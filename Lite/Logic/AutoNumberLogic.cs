using System;
using System.Collections.Generic;
using System.Linq;
using System.ServiceModel;
using System.Text;
using System.Threading.Tasks;

using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Metadata;
using Microsoft.Crm.Sdk.Messages;

namespace BriteGlobal.Xrm.Plugins.AutoNumber.Logic
{
    /// <summary>
    /// This is the Lite Version of the AutoNumberLogic class
    /// </summary>
    public class AutoNumberLogic : IDisposable
    {
        enum AttributeTypeCode
        {
            SingleLineOfText = 1,
            WholeNumber = 2
        }

        enum RelationshipTypeCode
        {
            PrimaryEntity = 1,
            RelatedEntity = 2,
            RelatedEntityMulti = 3,
            OptionSet = 4,
            OptionSetMulti = 5,
            RelatedEntityCounter = 6
        }

        #region Constructors and Global Variables

        private IOrganizationService service;
        private ITracingService tracing;
        public AutoNumberLogic(IOrganizationService _service)
        {
            service = _service;
        }

        public AutoNumberLogic(IOrganizationService _service, ITracingService _tracing)
        {
            service = _service;
            tracing = _tracing;
        }

        #endregion

        #region Logic Class Entry Point Functions

        public void CreateEntityLogic(string entityLogicalName, Guid entityId, LicenseMode license)
        {
            tracing.Trace("Entering CreateEntityLogic function");

            EntityCollection autonumbers = RetrieveAutoNumberSettings(entityLogicalName);
            if (autonumbers.Entities.Count > 0)
            {
                foreach (Entity autonumber in autonumbers.Entities)
                {
                    Guid autoNumberId = autonumber.Id;

                    string entityName = autonumber.Attributes["xrm_entityname"].ToString();
                    string attributeName = autonumber.Attributes["xrm_fieldname"].ToString();

                    UpdateAutoNumber(autonumber, entityLogicalName, attributeName, entityId, entityName);
                }
            }
        }

        public void AutoNumberPublishLogic(string entityLogicalName, Guid entityId)
        {
            Entity autonumber = RetrieveAutoNumberSettings(entityLogicalName, entityId);
            string entityName = autonumber.Attributes["xrm_entityname"].ToString();

            UserContext callingContext = autonumber.Contains("xrm_executioncontextcode") ? (UserContext)((OptionSetValue)(autonumber.Attributes["xrm_executioncontextcode"])).Value : UserContext.Creator;

            Guid impersonatingUserId = Guid.Empty;
            if (callingContext == UserContext.Creator)
            {
                impersonatingUserId = ((EntityReference)(autonumber.Attributes["createdby"])).Id;
            }

            // Publish SDK Message
            PublishSDKMessageProcessingStep(entityName, impersonatingUserId);
        }


        /// <summary>
        /// 
        /// </summary>
        /// <param name="entityLogicalName">xrm_autonumber</param>
        /// <param name="entityId">The Id of the AUtoNumber Record</param>
        public void FillAutoNumberLogic(string entityLogicalName, Guid autoNumberId)
        {
            ColumnSet columns = new ColumnSet("xrm_entityname", "xrm_fieldname", "xrm_relationshiptypecode");
            Entity autonumber = service.Retrieve(entityLogicalName, autoNumberId, columns);

            string entityName = autonumber.Attributes["xrm_entityname"].ToString();
            string attributeName = autonumber.Attributes["xrm_fieldname"].ToString();

            int relationshipTypeCode = ((OptionSetValue)(autonumber.Attributes["xrm_relationshiptypecode"])).Value;

            EntityCollection results = RetrieveEntities(entityName, attributeName, true);
            if (results.Entities.Count > 0)
            {
                foreach (Entity result in results.Entities)
                {
                    Guid entityId = result.Id;
                    UpdateAutoNumber(autonumber, entityLogicalName, attributeName, entityId, entityName);
                }
            }

        }

        #endregion

        private void UpdateAutoNumber(Entity autonumber, string entityLogicalName, string attributeName, Guid entityId, string entityName)
        {
            if (entityName.Contains(':'))
            {
                // Dual Entities
                bool entityExists = false;
                string[] entityNames = entityName.Split(':');
                int entityPosition = 0;
                foreach (string current in entityNames)
                {
                    entityPosition++;
                    if (current == entityLogicalName)
                    {
                        entityExists = true;
                        break;
                    }
                }

                if (entityExists)
                {
                    string[] attributeNames = attributeName.Split(':');
                    int attributePosition = 0;
                    foreach (string current in attributeNames)
                    {
                        attributePosition++;
                        if (attributePosition == entityPosition)
                        {
                            if (current.Contains(';'))
                            {
                                // Update Multiple Attributes
                                string[] attributes = current.Split(';');
                                UpdateAutoNumber(autonumber, entityId, entityLogicalName, attributes);
                            }
                            else
                            {
                                // Update Single Attribute
                                string[] attributes = new string[] { current };
                                UpdateAutoNumber(autonumber, entityId, entityLogicalName, attributes);
                            }
                            break;
                        }
                    }
                }
            }
            else
            {
                // Single Entity Name
                // Since using contains we have to compare the entity name to the logic name
                // Example would be entity called client and clientInformation would both be retrieved if searching for client
                if (entityName == entityLogicalName)
                {
                    string[] attributes = new string[] { attributeName };
                    UpdateAutoNumber(autonumber, entityId, entityLogicalName, attributes);
                }
            }

        }

        private void UpdateAutoNumber(Entity autonumber, Guid entityId, string entityName, string[] attributeNames)
        {
            tracing.Trace("Entering UpdateAutoNumber function");

            Entity update = new Entity(entityName);
            update.Id = entityId;

            int initialValue = Convert.ToInt32(autonumber.Attributes["xrm_initialvalue"]);
            int currentValue = autonumber.Attributes.ContainsKey("xrm_currentvalue") ? Convert.ToInt32(autonumber.Attributes["xrm_currentvalue"]) : 0;
            int nextValue = autonumber.Attributes.ContainsKey("xrm_nextvalue") ? Convert.ToInt32(autonumber.Attributes["xrm_nextvalue"]) : initialValue;
            int attributeType = ((OptionSetValue)(autonumber.Attributes["xrm_attributetype"])).Value;

            bool overwriteExistingValues = autonumber.Contains("xrm_overwriteexistingvalues") ? (bool)autonumber.Attributes["xrm_overwriteexistingvalues"] : false; 
            string currentAutoNumberValue = RetrieveRelatedEntityFieldValue(entityName, entityId, attributeNames[0]);

            tracing.Trace("Overwrite Existing Values: {0}", overwriteExistingValues.ToString());
            tracing.Trace("Current AutoNumber Value: {0}", currentAutoNumberValue);

            if (attributeType == AttributeTypeCode.SingleLineOfText.ToInt()) // Single Line of Text
            {
                string prefix = autonumber.Attributes.ContainsKey("xrm_prefix") ? autonumber.Attributes["xrm_prefix"].ToString() : string.Empty;
                string suffix = autonumber.Attributes.ContainsKey("xrm_suffix") ? autonumber.Attributes["xrm_suffix"].ToString() : string.Empty;
                int length = autonumber.Attributes.ContainsKey("xrm_length") ? ((OptionSetValue)(autonumber.Attributes["xrm_length"])).Value : 0;
                string nextStringValue = Helper.GenerateNextValue(prefix, suffix, nextValue, length);

                foreach (string attributeName in attributeNames)
                {
                    update[attributeName] = nextStringValue;
                }
                autonumber.Attributes["xrm_preview"] = nextStringValue;
            }
            else if (attributeType == AttributeTypeCode.WholeNumber.ToInt()) // Whole Number
            {
                foreach (string attributeName in attributeNames)
                {
                    update[attributeName] = nextValue;
                }
                autonumber.Attributes["xrm_preview"] = nextValue.ToString();
            }

            autonumber.Attributes["xrm_currentvalue"] = nextValue;
            autonumber.Attributes["xrm_nextvalue"] = nextValue + 1;

            try
            {
                if (overwriteExistingValues)
                {
                    service.Update(update);
                    service.Update(autonumber); // Update AutoNumber Entity
                }
                else
                {
                    if (string.IsNullOrEmpty(currentAutoNumberValue))
                    {
                        service.Update(update);
                        service.Update(autonumber); // Update AutoNumber Entity
                    }
                }
            }
            catch (FaultException<OrganizationServiceFault> ex)
            {
                throw new InvalidPluginExecutionException(
                    String.Format("An error occurred in the {0} function of the {1} plug-in.",
                       "UpdateAutoNumber", this.GetType().ToString()), ex);
            }
        }
                
        private EntityCollection RetrieveAutoNumberSettings(string entityName)
        {
            QueryExpression query = new QueryExpression("xrm_autonumber")
            {
                ColumnSet = new ColumnSet("xrm_entityname", "xrm_fieldname"),
                Criteria =
                {
                    FilterOperator = LogicalOperator.And,
                    Conditions =
                    {
                        new ConditionExpression("xrm_entityname", ConditionOperator.Equal, entityName),
                        new ConditionExpression("statuscode", ConditionOperator.Equal, 1),
                        new ConditionExpression("xrm_status", ConditionOperator.Equal, Convert.ToBoolean(1))
                    }
                }
            };

            RetrieveMultipleRequest request = new RetrieveMultipleRequest();
            request.Query = query;

            try
            {
                RetrieveMultipleResponse response = (RetrieveMultipleResponse)service.Execute(request);
                int count = response.EntityCollection.Entities.Count;
                return response.EntityCollection;
            }
            catch (FaultException<OrganizationServiceFault> ex)
            {
                throw new InvalidPluginExecutionException(
                    String.Format("An error occurred in the {0} function of the {1} plug-in. The encountered error was: {2}",
                       "RetrieveAutoNumberSettings", this.GetType().ToString(), ex.Message), ex);
            }
        }

        private Entity RetrieveAutoNumberSettings(string entityName, Guid entityId)
        {
            ColumnSet columns = new ColumnSet("xrm_entityname", "xrm_executioncontextcode", "createdby");
            Entity result = service.Retrieve(entityName, entityId, columns);
            return result;
        }

        private EntityCollection RetrieveEntities(string entityName, string nameField, bool nullOnly = false)
        {
            string idField = entityName + "id";
            EntityCollection results = RetrieveEntities(entityName, idField, nameField, nullOnly);
            return results;
        }

        private EntityCollection RetrieveEntities(string entityName, string idField, string nameField, bool nullOnly = false)
        {
            EntityCollection results = new EntityCollection();
            try
            {
                ColumnSet columns = new ColumnSet(idField, nameField);
                QueryExpression query = new QueryExpression();
                query.EntityName = entityName;
                query.ColumnSet = columns;

                if (nullOnly)
                {
                    query.Criteria = new FilterExpression();
                    ConditionExpression ce = new ConditionExpression(nameField, ConditionOperator.Null);
                    query.Criteria.Conditions.Add(ce);
                }

                results = service.RetrieveMultiple(query);
            }
            catch (FaultException<OrganizationServiceFault> ex)
            {
                throw new InvalidPluginExecutionException(
                    String.Format("An error occurred in the {0} function of the {1} plug-in.",
                       "RetrieveEntities", this.GetType().ToString()), ex);
            }
            return results;

        }

        private List<KeyValuePair<string, string>> RetrieveEntityKeyValues(string entityName, string idField, string nameField)
        {
            EntityCollection results = new EntityCollection();
            try
            {
                ColumnSet columns = new ColumnSet(idField, nameField);
                QueryExpression query = new QueryExpression();
                query.EntityName = entityName;
                query.ColumnSet = columns;

                results = service.RetrieveMultiple(query);

                List<KeyValuePair<string, string>> list = new List<KeyValuePair<string, string>>();
                if (results.Entities.Count > 0)
                {
                    foreach (Entity result in results.Entities)
                    {
                        string idFieldValue = result.Id.ToString();
                        string nameFieldValue = result.Contains(nameField) ? result[nameField].ToString() : "";
                        list.Add(new KeyValuePair<string, string>(idFieldValue, nameFieldValue));
                    }

                }
                return list;
            }
            catch (FaultException<OrganizationServiceFault> ex)
            {
                throw new InvalidPluginExecutionException(
                    String.Format("An error occurred in the {0} function of the {1} plug-in.",
                       "RetrieveEntities", this.GetType().ToString()), ex);
            }

        }


        private Guid RetrieveLookupId(string entityName, Guid entityId, string lookupFieldName)
        {
            Entity result = service.Retrieve(entityName, entityId, new ColumnSet(lookupFieldName));
            Guid rc = result.Contains(lookupFieldName) ? ((EntityReference)(result[lookupFieldName])).Id : Guid.Empty;
            return rc;
        }

        private int? RetrieveOptionSetValue(string entityName, Guid entityId, string lookupFieldName)
        {
            Entity result = service.Retrieve(entityName, entityId, new ColumnSet(lookupFieldName));
            int? rc = null;

            if (result.Contains(lookupFieldName))
                rc = ((OptionSetValue)(result[lookupFieldName])).Value;
            return rc;
        }

        private string RetrieveRelatedEntityFieldValue(string entityName, Guid entityId, string dataFieldName)
        {
            Entity result = service.Retrieve(entityName, entityId, new ColumnSet(dataFieldName));
            string rc = result.Contains(dataFieldName) ? result[dataFieldName].ToString() : string.Empty;
            return rc;
        }

        private string RetrieveRelatedEntityFieldValue(string entityName, Guid entityId, string dataFieldName, string counterFieldName, out int counterFieldValue)
        {
            Entity result = service.Retrieve(entityName, entityId, new ColumnSet(dataFieldName, counterFieldName));
            string rc = result.Contains(dataFieldName) ? result[dataFieldName].ToString() : string.Empty;

            counterFieldValue = 0;
            if (result.Contains(counterFieldName))
            {
                Int32.TryParse(result[counterFieldName].ToString(), out counterFieldValue);
            }

            return rc;
        }

        private void UpdateRelatedEntityCounter(string entityName, Guid relatedEntityId, string relatedCounterAttributeName, int nextValue)
        {
            Entity update = new Entity(entityName);
            update.Id = relatedEntityId;
            update[relatedCounterAttributeName] = nextValue;

            try
            {
                service.Update(update);
            }
            catch (FaultException<OrganizationServiceFault> ex)
            {
                throw new InvalidPluginExecutionException(
                    String.Format("An error occurred in the {0} function of the {1} plug-in.",
                       "UpdateRelatedEntityCounter", this.GetType().ToString()), ex);
            }
        }

        #region Metadata Functions

        private List<KeyValuePair<string, string>> RetrieveAllEntities()
        {
            List<KeyValuePair<string, string>> entities = new List<KeyValuePair<string, string>>();

            RetrieveAllEntitiesRequest request = new RetrieveAllEntitiesRequest()
            {
                EntityFilters = Microsoft.Xrm.Sdk.Metadata.EntityFilters.Entity,
                RetrieveAsIfPublished = true,
            };

            try
            {
                RetrieveAllEntitiesResponse response = (RetrieveAllEntitiesResponse)service.Execute(request);
                if (response.EntityMetadata.Length > 0)
                {
                    foreach (EntityMetadata entity in response.EntityMetadata)
                    {
                        string entityName = entity.LogicalName;
                        string displayName = entity.DisplayName.UserLocalizedLabel.Label;
                        string primaryAttribute = entity.PrimaryNameAttribute;
                        entities.Add(new KeyValuePair<string, string>(entityName, displayName));
                    }
                }
            }
            catch (FaultException<OrganizationServiceFault> ex)
            {

            }
            return entities;

        }

        private List<KeyValuePair<string, string>> RetrieveAttributeMetadata(string entityName, string attributeName)
        {
            RetrieveAttributeRequest request = new RetrieveAttributeRequest()
            {
                EntityLogicalName = entityName,
                LogicalName = attributeName,
                RetrieveAsIfPublished = true
            };

            RetrieveAttributeResponse response = (RetrieveAttributeResponse)service.Execute(request);
            PicklistAttributeMetadata picklistMetadata = (PicklistAttributeMetadata)response.AttributeMetadata;
            OptionSetMetadata optionsetMetadata = (OptionSetMetadata)picklistMetadata.OptionSet;

            List<KeyValuePair<string, string>> list = new List<KeyValuePair<string, string>>();
            foreach (OptionMetadata option in optionsetMetadata.Options)
            {
                int? optionKey = option.Value;
                string optionValue = option.Label.UserLocalizedLabel.Label;
                if (optionKey.HasValue)
                {
                    int optionKeyValue = optionKey.Value;
                    list.Add(new KeyValuePair<string, string>(optionKeyValue.ToString(), optionValue));
                }
            }

            return list;
        }

        private KeyValuePair<string, string> RetrieveEntityMetadata(string entityName)
        {
            RetrieveEntityRequest request = new RetrieveEntityRequest()
            {
                EntityFilters = EntityFilters.Entity,
                LogicalName = entityName
            };

            RetrieveEntityResponse response = (RetrieveEntityResponse)service.Execute(request);
            KeyValuePair<string, string> result = new KeyValuePair<string, string>(response.EntityMetadata.PrimaryIdAttribute, response.EntityMetadata.PrimaryNameAttribute);
            return result;
        }

        private int? RetrieveEntityMetadataObjectTypeCode(string entityName)
        {
            RetrieveEntityRequest request = new RetrieveEntityRequest()
            {
                EntityFilters = EntityFilters.Entity,
                LogicalName = entityName
            };

            RetrieveEntityResponse response = (RetrieveEntityResponse)service.Execute(request);
            int? result = response.EntityMetadata.ObjectTypeCode;
            return result;
        }

        #endregion

        #region SDK Message Processing Step Publishing

        const string ASSEMBLY_NAME = "BriteGlobal.Xrm.Plugins.AutoNumber";
        const string PLUGIN_TYPE_NAME = "BriteGlobal.Xrm.Plugins.AutoNumber.PluginEntryPoint";

        enum StepMode
        {
            Sync = 0,
            Async = 1
        }

        enum PluginStage
        {
            PreValidation = 10,
            PreOperation = 20,
            PostOperation = 40
        }
        
        enum PluginStepDeployment
        {
            ServerOnly = 0,
            OfflineOnly = 1,
            Both = 2
        }

        enum StepInvocationSource
        {
            Parent = 0,
            Child = 1
        }

        enum SdkMessageName
        {
            Create,
            Delete,
            Retrieve,
            Update
        }

        enum UserContext
        {
            Creator = 1,
            CallingUser = 2
        }

        public void PublishSDKMessageProcessingStep(string entityName, Guid impersonatingUserId)
        {
            Guid sdkMessageId = GetSdkMessageId("Create");
            int? objectTypeCode = RetrieveEntityMetadataObjectTypeCode(entityName.ToLower());
            if (objectTypeCode.HasValue)
            {
                bool stepExists = RetrieveSdkMessageProcessingStep(sdkMessageId, objectTypeCode.Value);

                if (!stepExists)
                {

                    CreateSdkMessageProcessingStep(string.Format("{0}: Create of {1}", PLUGIN_TYPE_NAME, entityName), entityName.ToLower(), "", StepMode.Sync, 10, PluginStage.PostOperation, StepInvocationSource.Parent, impersonatingUserId);
                }
            }
        }


        /// <summary>
        /// 
        /// </summary>
        /// <param name="name">Name of the Sdk Message Processing Step to Create</param>
        /// <param name="configuration">Leave Empty</param>
        /// <param name="mode">Sync or Async (enumeration)</param>
        /// <param name="rank">Default to 10</param>
        /// <param name="stage">Prevalidation, Preoperation or Postoperation</param>
        /// <param name="invocationSource">Parent or Child</param>
        /// <param name="pluginTypeId"></param>
        /// <param name="messageId"></param>
        /// <param name="messageFilterId"></param>
        /// <returns></returns>
        private Guid CreateSdkMessageProcessingStep(string name, string entityName, string configuration, StepMode mode, int rank, PluginStage stage, StepInvocationSource invocationSource, Guid impersonatingUserId)
        {
            
            Entity step = new Entity("sdkmessageprocessingstep");
            step["name"] = name;
            step["description"] = name;
            step["configuration"] = configuration;
            step["mode"] = new OptionSetValue(mode.ToInt());
            step["rank"] = rank;
            step["stage"] = new OptionSetValue(stage.ToInt());
            step["supporteddeployment"] = new OptionSetValue(0); // Server Only
            step["invocationsource"] = new OptionSetValue(invocationSource.ToInt());

            Guid sdkMessageId = GetSdkMessageId("Create");
            Guid sdkMessageFilterId = GetSdkMessageFilterId(entityName, sdkMessageId);

            Guid assemblyId = GetPluginAssemblyId(ASSEMBLY_NAME);
            Guid pluginTypeId = GetPluginTypeId(assemblyId, PLUGIN_TYPE_NAME);

            step["plugintypeid"] = new EntityReference("plugintype", pluginTypeId);
            step["sdkmessageid"] = new EntityReference("sdkmessage", sdkMessageId);
            step["sdkmessagefilterid"] = new EntityReference("sdkmessagefilter", sdkMessageFilterId);

            if (impersonatingUserId != Guid.Empty)
                step["impersonatinguserid"] = new EntityReference("systemuser", impersonatingUserId);

            try
            {
                Guid stepId = service.Create(step);
                return stepId;
            }
            catch (InvalidPluginExecutionException invalidPluginExecutionException)
            {
                throw invalidPluginExecutionException;
            }
            catch (Exception exception)
            {
                throw exception;
            }
        }

        private bool RetrieveSdkMessageProcessingStep(Guid sdkMessageId, int primaryObjectTypeCode)
        {
            bool rc = false;
            QueryExpression query = new QueryExpression()
            {
                EntityName = "sdkmessageprocessingstep",
                ColumnSet = new ColumnSet(new string[] { "name", "description", "plugintypeid", "sdkmessageid", "eventhandler" }),
                Criteria =
                {
                    Conditions =
                    {
                        new ConditionExpression("sdkmessageid", ConditionOperator.Equal, sdkMessageId)
                    }
                },
                LinkEntities =
                {
                   new LinkEntity()
                   {
                       Columns = new ColumnSet("primaryobjecttypecode", "sdkmessageid"),
                       LinkToEntityName = "sdkmessagefilter",
                       LinkToAttributeName = "sdkmessagefilterid",
                       LinkFromEntityName = "sdkmessageprocessingstep",
                       LinkFromAttributeName = "sdkmessagefilterid",
                       LinkCriteria =
                       {
                            Conditions =
                            {
                                new ConditionExpression("primaryobjecttypecode", ConditionOperator.Equal, primaryObjectTypeCode)
                            }
                       }
                   },
                   new LinkEntity()
                   {
                       LinkToEntityName = "plugintype",
                       LinkToAttributeName = "plugintypeid",
                       LinkFromEntityName = "sdkmessageprocessingstep",
                       LinkFromAttributeName = "plugintypeid",
                       LinkCriteria =
                       {
                           Conditions =
                           {
                               new ConditionExpression("assemblyname", ConditionOperator.Equal, ASSEMBLY_NAME)
                           }
                       }
                   }
                }
            };

            EntityCollection sdkMessageProcessingSteps = service.RetrieveMultiple(query);

            if (sdkMessageProcessingSteps.Entities.Count > 0)
            {
                rc = true;
            }

            return rc;
        }

        private Guid GetSdkMessageFilterId(string EntityLogicalName, Guid sdkMessageId)
        {
            try
            {
                //GET SDK MESSAGE FILTER QUERY
                QueryExpression sdkMessageFilterQueryExpression = new QueryExpression("sdkmessagefilter");
                sdkMessageFilterQueryExpression.ColumnSet = new ColumnSet("sdkmessagefilterid");
                sdkMessageFilterQueryExpression.Criteria = new FilterExpression
                {
                    Conditions =
                        {
                            new ConditionExpression
                            {
                                AttributeName = "primaryobjecttypecode",
                                Operator = ConditionOperator.Equal,
                                Values = {EntityLogicalName}
                            },
                            new ConditionExpression
                            {
                                AttributeName = "sdkmessageid",
                                Operator = ConditionOperator.Equal,
                                Values = {sdkMessageId}
                            },
                        }
                };

                //RETRIEVE SDK MESSAGE FILTER
                EntityCollection sdkMessageFilters = service.RetrieveMultiple(sdkMessageFilterQueryExpression);

                if (sdkMessageFilters.Entities.Count != 0)
                {
                    return sdkMessageFilters.Entities.First().Id;
                }
                throw new Exception(String.Format("SDK Message Filter for {0} was not found.", EntityLogicalName));
            }
            catch (InvalidPluginExecutionException invalidPluginExecutionException)
            {
                throw invalidPluginExecutionException;
            }
            catch (Exception exception)
            {
                throw exception;
            }
        }

        private Guid GetSdkMessageId(string SdkMessageName)
        {
            try
            {
                //GET SDK MESSAGE QUERY
                QueryExpression sdkMessageQueryExpression = new QueryExpression("sdkmessage");
                sdkMessageQueryExpression.ColumnSet = new ColumnSet("sdkmessageid");
                sdkMessageQueryExpression.Criteria = new FilterExpression
                {
                    Conditions =
                        {
                            new ConditionExpression
                            {
                                AttributeName = "name",
                                Operator = ConditionOperator.Equal,
                                Values = {SdkMessageName}
                            },
                        }
                };

                //RETRIEVE SDK MESSAGE
                EntityCollection sdkMessages = service.RetrieveMultiple(sdkMessageQueryExpression);
                if (sdkMessages.Entities.Count != 0)
                {
                    return sdkMessages.Entities.First().Id;
                }
                throw new Exception(String.Format("SDK MessageName {0} was not found.", SdkMessageName));
            }
            catch (InvalidPluginExecutionException invalidPluginExecutionException)
            {
                throw invalidPluginExecutionException;
            }
            catch (Exception exception)
            {
                throw exception;
            }
        }

        private Guid GetPluginAssemblyId(string assemblyName)
        {
            try
            {
                //GET ASSEMBLY QUERY
                QueryExpression pluginAssemblyQueryExpression = new QueryExpression("pluginassembly");
                pluginAssemblyQueryExpression.ColumnSet = new ColumnSet("pluginassemblyid");
                pluginAssemblyQueryExpression.Criteria = new FilterExpression
                {
                    Conditions =
                        {
                            new ConditionExpression
                            {
                                AttributeName = "name",
                                Operator = ConditionOperator.Equal,
                                Values = {assemblyName}
                            },
                        }
                };

                //RETRIEVE ASSEMBLY
                EntityCollection pluginAssemblies = service.RetrieveMultiple(pluginAssemblyQueryExpression);
                Guid assemblyId = Guid.Empty;
                if (pluginAssemblies.Entities.Count != 0)
                {
                    //ASSIGN ASSEMBLY ID TO VARIABLE
                    assemblyId = pluginAssemblies.Entities.First().Id;
                }
                return assemblyId;
            }
            catch (InvalidPluginExecutionException invalidPluginExecutionException)
            {
                throw invalidPluginExecutionException;
            }
            catch (Exception exception)
            {
                throw exception;
            }
        }

        private Guid GetPluginTypeId(Guid assemblyId, string PluginTypeName)
        {
            try
            {
                    //GET PLUGIN TYPES WITHIN ASSEMBLY
                    QueryExpression pluginTypeQueryExpression = new QueryExpression("plugintype");
                    pluginTypeQueryExpression.ColumnSet = new ColumnSet("plugintypeid");
                    pluginTypeQueryExpression.Criteria = new FilterExpression
                    {
                        Conditions =
                        {
                            new ConditionExpression
                            {
                                AttributeName = "pluginassemblyid",
                                Operator = ConditionOperator.Equal,
                                Values = {assemblyId}
                            },
                            new ConditionExpression
                            {
                                AttributeName = "typename",
                                Operator = ConditionOperator.Equal,
                                Values = {PluginTypeName}
                            },
                        }
                    };

                    //RETRIEVE PLUGIN TYPES IN ASSEMBLY
                    EntityCollection pluginTypes = service.RetrieveMultiple(pluginTypeQueryExpression);

                    //RETURN PLUGIN TYPE ID
                    if (pluginTypes.Entities.Count != 0)
                    {
                        return pluginTypes.Entities.First().Id;
                    }
                    throw new Exception(String.Format("Plugin Type {0} was not found in Assembly", PluginTypeName));
            }
            catch (InvalidPluginExecutionException invalidPluginExecutionException)
            {
                throw invalidPluginExecutionException;
            }
            catch (Exception exception)
            {
                throw exception;
            }
        }

        #endregion

        #region IDisposable Support
        private bool disposedValue = false; // To detect redundant calls

        protected virtual void Dispose(bool disposing)
        {
            if (!disposedValue)
            {
                if (disposing)
                {
                    // TODO: dispose managed state (managed objects).
                }

                // TODO: free unmanaged resources (unmanaged objects) and override a finalizer below.
                // TODO: set large fields to null.

                disposedValue = true;
            }
        }

        // TODO: override a finalizer only if Dispose(bool disposing) above has code to free unmanaged resources.
        // ~ExpenditureLogic() {
        //   // Do not change this code. Put cleanup code in Dispose(bool disposing) above.
        //   Dispose(false);
        // }

        // This code added to correctly implement the disposable pattern.
        public void Dispose()
        {
            // Do not change this code. Put cleanup code in Dispose(bool disposing) above.
            Dispose(true);
            // TODO: uncomment the following line if the finalizer is overridden above.
            // GC.SuppressFinalize(this);
        }
        #endregion
    }
}
