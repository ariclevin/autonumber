using System;
using System.Collections.Generic;
using System.Linq;
using System.ServiceModel;
using System.Text;
using System.Threading.Tasks;

using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk.Messages;

namespace BriteGlobal.Xrm.Plugins.Auto.Logic
{
    public class AutoNumberLogic
    {
        enum AttributeTypeCode
        {
            SingleLineOfText = 1,
            WholeNumber = 2
        }

        enum RelationshipTypeCode
        {
            PrimaryEntity = 1,
            RelatedEntity = 2
        }

        private IOrganizationService service;
        public AutoNumberLogic(IOrganizationService _service)
        {
            service = _service;
        }


        public void CreateAutoNumber(string entityLogicalName, Guid entityId)
        {
            EntityCollection autonumbers = RetrieveAutoNumberSettings(entityLogicalName);
            if (autonumbers.Entities.Count > 0)
            {
                foreach (Entity autonumber in autonumbers.Entities)
                {
                    Guid autoNumberId = autonumber.Id;

                    string entityName = autonumber.Attributes["xrm_entityname"].ToString();
                    string attributeName = autonumber.Attributes["xrm_fieldname"].ToString();

                    int relationshipTypeCode = ((OptionSetValue)(autonumber.Attributes["xrm_relationshipTypeCode"])).Value;
                    if (relationshipTypeCode == RelationshipTypeCode.RelatedEntity.ToInt())
                    {
                        UpdateLookupAutoNumber(autonumber, entityId, entityLogicalName, attributeName);
                    }
                    else
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
                }
            }
        }

        private void UpdateAutoNumber(Entity autonumber, Guid entityId, string entityName, string[] attributeNames)
        {
            Entity update = new Entity(entityName);
            update.Id = entityId;

            int initialValue = Convert.ToInt32(autonumber.Attributes["xrm_initialvalue"]);
            int currentValue = autonumber.Attributes.ContainsKey("xrm_currentvalue") ? Convert.ToInt32(autonumber.Attributes["xrm_currentvalue"]) : 0;
            int nextValue = autonumber.Attributes.ContainsKey("xrm_nextvalue") ? Convert.ToInt32(autonumber.Attributes["xrm_nextvalue"]) : initialValue;
            int attributeType = ((OptionSetValue)(autonumber.Attributes["xrm_attributetype"])).Value;

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
                service.Update(update);
                service.Update(autonumber); // Update AutoNumber Entity
            }
            catch (FaultException<OrganizationServiceFault> ex)
            {
                throw new InvalidPluginExecutionException(
                    String.Format("An error occurred in the {0} function of the {1} plug-in.",
                       "CreateAutoNumber", this.GetType().ToString()), ex);
            }
        }

        private void UpdateLookupAutoNumber(Entity autonumber, Guid entityId, string entityName, string attributeName)
        {
            Entity update = new Entity(entityName);
            update.Id = entityId;

            int initialValue = Convert.ToInt32(autonumber.Attributes["xrm_initialvalue"]);
            int currentValue = autonumber.Attributes.ContainsKey("xrm_currentvalue") ? Convert.ToInt32(autonumber.Attributes["xrm_currentvalue"]) : 0;
            int nextValue = autonumber.Attributes.ContainsKey("xrm_nextvalue") ? Convert.ToInt32(autonumber.Attributes["xrm_nextvalue"]) : initialValue;
            int attributeType = ((OptionSetValue)(autonumber.Attributes["xrm_attributetype"])).Value;

            if (attributeType == AttributeTypeCode.SingleLineOfText.ToInt()) // Single Line of Text
            {
                string lookupFieldName = autonumber.Attributes.ContainsKey("xrm_lookupfieldname") ? autonumber.Attributes["xrm_lookupfieldname"].ToString() : string.Empty;
                Guid relatedEntityId = RetrieveLookupId(entityName, entityId, lookupFieldName);

                if (relatedEntityId != Guid.Empty)
                {
                    string relatedEntityName = autonumber.Attributes.ContainsKey("xrm_relatedentityname") ? autonumber.Attributes["xrm_relatedentityname"].ToString() : string.Empty;
                    string relatedAttributeName = autonumber.Attributes.ContainsKey("xrm_relatedentityfieldname") ? autonumber.Attributes["xrm_relatedentityfieldname"].ToString() : string.Empty;
                    string relatedEntityAttributeValue = RetrieveRelatedEntityFieldValue(relatedEntityName, relatedEntityId, relatedAttributeName);
                    string separator = autonumber.Attributes.ContainsKey("xrm_separatorcharacter") ? autonumber.Attributes["xrm_separatorcharacter"].ToString() : string.Empty;

                    string prefix = autonumber.Attributes.ContainsKey("xrm_prefix") ? autonumber.Attributes["xrm_prefix"].ToString() : string.Empty;
                    string suffix = autonumber.Attributes.ContainsKey("xrm_suffix") ? autonumber.Attributes["xrm_suffix"].ToString() : string.Empty;
                    int length = autonumber.Attributes.ContainsKey("xrm_length") ? ((OptionSetValue)(autonumber.Attributes["xrm_length"])).Value : 0;
                    string nextStringValue = Helper.GenerateNextValue(prefix, suffix, relatedEntityAttributeValue, separator, nextValue, length);

                    update[attributeName] = nextStringValue;
                    autonumber.Attributes["xrm_preview"] = nextStringValue;
                    autonumber.Attributes["xrm_currentvalue"] = nextValue;
                    autonumber.Attributes["xrm_nextvalue"] = nextValue + 1;
                }
            }

            try
            {
                service.Update(update);
                service.Update(autonumber); // Update AutoNumber Entity
            }
            catch (FaultException<OrganizationServiceFault> ex)
            {
                throw new InvalidPluginExecutionException(
                    String.Format("An error occurred in the {0} function of the {1} plug-in.",
                       "CreateAutoNumber", this.GetType().ToString()), ex);
            }
        }


        public void FillAutoNumber(string entityLogicalName, Guid entityId)
        {
            Entity entity = service.Retrieve(entityLogicalName, entityId, new ColumnSet(true));

            string entityName = entity.Attributes["xrm_entityname"].ToString();
            string fieldName = entity.Attributes["xrm_fieldname"].ToString();
            int initialValue = Convert.ToInt32(entity.Attributes["xrm_initialvalue"]);
            int currentValue = entity.Attributes.ContainsKey("xrm_currentvalue") ? Convert.ToInt32(entity.Attributes["xrm_currentvalue"]) : 0;
            int nextValue = entity.Attributes.ContainsKey("xrm_nextvalue") ? Convert.ToInt32(entity.Attributes["xrm_currentvalue"]) : initialValue;


            int attributeType = ((OptionSetValue)(entity.Attributes["xrm_attributetype"])).Value;

            EntityCollection results = RetrieveEntities(entityName, fieldName);
            if (results.Entities.Count > 0)
            {
                string prefix = "", suffix = ""; int length = 0;
                if (attributeType == 1) // Single Line of Text
                {
                    prefix = entity.Attributes.ContainsKey("xrm_prefix") ? entity.Attributes["xrm_prefix"].ToString() : string.Empty;
                    suffix = entity.Attributes.ContainsKey("xrm_suffix") ? entity.Attributes["xrm_suffix"].ToString() : string.Empty;
                    length = entity.Attributes.ContainsKey("xrm_length") ? ((OptionSetValue)(entity.Attributes["xrm_length"])).Value : 0;
                }

                foreach (Entity current in results.Entities)
                {
                    Guid currentId = current.Id;
                    if (attributeType == AttributeTypeCode.SingleLineOfText.ToInt()) // Single Line of Text
                    {
                        string nextStringValue = Helper.GenerateNextValue(prefix, suffix, nextValue, length);
                        current[fieldName] = nextStringValue;
                    }
                    else if (attributeType == AttributeTypeCode.WholeNumber.ToInt()) // Whole Number
                    {
                        current[fieldName] = nextValue;
                    }

                    try
                    {
                        service.Update(current); // Update Entity
                    }
                    catch (FaultException<OrganizationServiceFault> ex)
                    {
                        throw new InvalidPluginExecutionException(
                            String.Format("An error occurred in the {0} function of the {1} plug-in.",
                                "CreateAutoNumber", this.GetType().ToString()), ex);
                    }
                    nextValue++;
                }

                // Update AutoNumber Entity
                entity.Attributes["xrm_currentvalue"] = nextValue - 1;
                entity.Attributes["xrm_nextvalue"] = nextValue;
                entity.Attributes["xrm_autofilled"] = false;
                if (attributeType == 1) // Single Line of Text
                    entity.Attributes["xrm_preview"] = Helper.GenerateNextValue(prefix, suffix, nextValue, length);
                else if (attributeType == 2) // Whole Number
                    entity.Attributes["xrm_preview"] = nextValue.ToString();

                try
                {
                    service.Update(entity); // Update AutoNumber Entity
                }
                catch (FaultException<OrganizationServiceFault> ex)
                {
                    throw new InvalidPluginExecutionException(
                        String.Format("An error occurred in the {0} function of the {1} plug-in.",
                            "CreateAutoNumber", this.GetType().ToString()), ex);
                }
            }

        }

        private EntityCollection RetrieveAutoNumberSettings(string entityName)
        {
            QueryExpression query = new QueryExpression("xrm_autonumber")
            {
                ColumnSet = new ColumnSet(true),
                Criteria =
                {
                    FilterOperator = LogicalOperator.And,
                    Conditions =
                    {
                        new ConditionExpression("xrm_entityname", ConditionOperator.Like, "%" + entityName + "%"),
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

        private EntityCollection RetrieveEntities(string entityname, string fieldname)
        {
            EntityCollection results = new EntityCollection();
            try
            {
                string[] columns = new string[] { entityname + "id", fieldname };

                QueryExpression query = new QueryExpression();
                query.EntityName = entityname;
                query.ColumnSet = new ColumnSet(columns);

                query.Criteria = new FilterExpression();
                ConditionExpression ce = new ConditionExpression(fieldname, ConditionOperator.Null);
                query.Criteria.Conditions.Add(ce);

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

        private Guid RetrieveLookupId(string entityName, Guid entityId, string lookupFieldName)
        {
            Entity result = service.Retrieve(entityName, entityId, new ColumnSet(lookupFieldName));
            Guid rc = result.Contains(lookupFieldName) ? ((EntityReference)(result[lookupFieldName])).Id : Guid.Empty;
            return rc;
        }

        private string RetrieveRelatedEntityFieldValue(string entityName, Guid entityId, string dataFieldName)
        {
            Entity result = service.Retrieve(entityName, entityId, new ColumnSet(dataFieldName));
            string rc = result.Contains(dataFieldName) ? result[dataFieldName].ToString() : string.Empty;
            return rc;
        }


    }
}
