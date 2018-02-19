namespace BriteGlobal.Xrm
{
    using System;
    using System.Collections.Generic;
    using System.Collections.ObjectModel;
    using System.Globalization;
    using System.IO;
    using System.Text;
    using System.Text.RegularExpressions;
    using System.Linq;
    using System.Net.Mail;
    using System.Security.Cryptography;
    using System.ServiceModel;
    using Microsoft.Xrm.Sdk;
    using Microsoft.Xrm.Sdk.Messages;
    using Microsoft.Xrm.Sdk.Query;
    using Microsoft.Crm.Sdk.Messages;

    using BriteGlobal.Xrm.Plugins.AutoNumber;

    public static class Helper
    {
        public static string SecurityCode { get; set; }

        public static string Decrypt(string cryptedString)
        {
            byte[] bytes = System.Text.Encoding.ASCII.GetBytes(SecurityCode.ToString());
            if (String.IsNullOrEmpty(cryptedString))
            {
                throw new ArgumentNullException
                   ("The string which needs to be decrypted can not be null.");
            }
            DESCryptoServiceProvider cryptoProvider = new DESCryptoServiceProvider();
            MemoryStream memoryStream = new MemoryStream
                    (Convert.FromBase64String(cryptedString));
            CryptoStream cryptoStream = new CryptoStream(memoryStream,
                cryptoProvider.CreateDecryptor(bytes, bytes), CryptoStreamMode.Read);
            StreamReader reader = new StreamReader(cryptoStream);
            return reader.ReadToEnd();
        }

        public static string Encrypt(string originalString)
        {
            byte[] bytes = System.Text.Encoding.ASCII.GetBytes(SecurityCode.ToString());
            if (String.IsNullOrEmpty(originalString))
            {
                throw new ArgumentNullException
                       ("The string which needs to be encrypted can not be null.");
            }
            DESCryptoServiceProvider cryptoProvider = new DESCryptoServiceProvider();
            MemoryStream memoryStream = new MemoryStream();
            CryptoStream cryptoStream = new CryptoStream(memoryStream,
                cryptoProvider.CreateEncryptor(bytes, bytes), CryptoStreamMode.Write);
            StreamWriter writer = new StreamWriter(cryptoStream);
            writer.Write(originalString);
            writer.Flush();
            cryptoStream.FlushFinalBlock();
            writer.Flush();
            return Convert.ToBase64String(memoryStream.GetBuffer(), 0, (int)memoryStream.Length);
        }

        public static string StripPunctuation(this string s)
        {
            var sb = new StringBuilder();
            foreach (char c in s)
            {
                if (char.IsLetterOrDigit(c))
                    sb.Append(c);
            }
            return sb.ToString();
        }

        public static int ToInt(this Enum i)
        {
            return Convert.ToInt32(i);
        }

        public static bool IsEmailRegex(string emailAddress)
        {
            string patternStrict = @"^(([^<>()[\]\\.,;:\s@\""]+" + @"(\.[^<>()[\]\\.,;:\s@\""]+)*)|(\"".+\""))@"
                  + @"((\[[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}" + @"\.[0-9]{1,3}\])|(([a-zA-Z\-0-9]+\.)+" + @"[a-zA-Z]{2,}))$";
            Regex reStrict = new Regex(patternStrict);

            bool isStrictMatch = reStrict.IsMatch(emailAddress);
            return isStrictMatch;
        }

        public static string GenerateNextValue(string prefix, string suffix, int nextValue, int length)
        {
            string rc = "";

            if (nextValue.ToString().Length < length)
                rc = nextValue.ToString().PadLeft(length, '0');
            else
                rc = nextValue.ToString();

            rc = prefix + rc;
            rc = rc + suffix;

            return rc;
        }

        public static string GenerateNextValue(string prefix, string suffix, string dataFieldValue, string separator, int nextValue, int length)
        {
            string rc = "";

            if (nextValue.ToString().Length < length)
                rc = nextValue.ToString().PadLeft(length, '0');
            else
                rc = nextValue.ToString();

            rc = dataFieldValue + separator + rc;

            rc = prefix + rc;
            rc = rc + suffix;

            return rc;
        }


        public static LicenseMode ValidateLicense(out string error)
        {
            bool rc = false;
            LicenseMode license = LicenseMode.Unavailable;

            error = "";
            bool isLicenseValid = ValidateKey(License.LicenseKey);
            if (!isLicenseValid)
            {
                DateTime today = DateTime.UtcNow;
                long todayTicks = today.Ticks;

                if (!string.IsNullOrEmpty(License.ExpirationDateTick))
                {
                    long expDateTicks = long.MinValue;
                    bool isInt64 = long.TryParse(License.ExpirationDateTick, out expDateTicks);
                    if (isInt64)
                    {
                        if (expDateTicks > todayTicks)
                        {
                            rc = true;
                        }
                    }
                    else
                    {

                    }
                }
            }
            else
                rc = true;

            if (!isLicenseValid)
            {
                license = LicenseMode.Unavailable;
            }
            else
            {
                license = LicenseMode.Enterprise;
            }
            return license; // rc;
        }

        private static bool ValidateKey(string key)
        {
            const string KEY_FORMAT = "%#^^?-^^%%^-^^?^^-#%#%#-^^%#?";
            List<Int32> list = new List<Int32>();

            if (!string.IsNullOrEmpty(key))
            {
                if (KEY_FORMAT.Length == key.Length)
                {
                    for (int i = 0; i < key.Length; i++)
                    {
                        string k = key.Substring(i, 1);
                        string f = KEY_FORMAT.Substring(i, 1);

                        char current = Convert.ToChar(k);
                        switch (f)
                        {
                            case "^": // Upper or Lowecase Letter
                                if (!char.IsLetter(current))
                                    return false;
                                break;
                            case "#":
                                if (!char.IsDigit(current))
                                    return false;
                                break;
                            case "?":
                                if (!char.IsLetterOrDigit(current))
                                    return false;
                                break;
                            case "%":
                                list.Add(Convert.ToInt32(k));
                                break;
                            default:
                                if (!(k.ToLower() == f.ToLower()))
                                    return false;
                                break;
                        }
                    }

                    int sum = 0;
                    foreach (int val in list)
                    {
                        sum += val;
                    }

                    if ((sum % 7) != 0)
                        return false;

                    return true;
                }
                else
                {
                    return false;
                }
            }
            else
            {
                return false;
            }
        }

    }

    public static class CRMHelper
    {
        public static IOrganizationService OrganizationService { private get; set; }
        public static ITracingService TracingService { private get; set; }

        public static string ExceptionMethodName { get; set; }
        public static string CallingMethodName { get; set; }
        public static string CallingClassName { get; set; }

        public enum OwnerTypeCode : int
        {
            SystemUser = 8,
            Team = 9
        }

        public enum StateCode : int
        {
            Active = 0,
            Inactive = 1
        }

        /// <summary>
        /// 
        /// </summary>
        /// <remarks>Must set the Organization Service before calling this function</remarks>
        /// <param name="key"></param>
        /// <returns></returns>
        public static string RetrieveApplicationSettingValue(string key, bool encrypted = false)
        {
            string rc = "";
            QueryByAttribute query = new QueryByAttribute();
            query.ColumnSet = new ColumnSet("xrm_key", "xrm_value", "xrm_securevalue");
            query.EntityName = "xrm_applicationsetting";
            query.Attributes.Add("xrm_key");
            query.Values.Add(key);

            RetrieveMultipleRequest request = new RetrieveMultipleRequest();
            request.Query = query;

            try
            {
                RetrieveMultipleResponse response = (RetrieveMultipleResponse)OrganizationService.Execute(request);
                EntityCollection results = response.EntityCollection;
                if (results.Entities.Count > 0)
                {
                    if (encrypted)
                        rc = results.Entities[0].Attributes["xrm_securevalue"].ToString();
                    else
                        rc = results.Entities[0].Attributes["xrm_value"].ToString();
                }
                return rc;
            }
            catch (FaultException<OrganizationServiceFault> ex)
            {
                throw new InvalidPluginExecutionException(
                    String.Format("An error occurred in the {0} function of the {1} plug-in.",
                       "RetrieveApplicationSettingValue", "CRMHelper"), ex);
            }
        }

        /// <summary>
        /// Retrieve Guid of CRM System User based on their AD login name
        /// </summary>
        public static Guid RetrieveUserIdByDomainName(string domainname)
        {
            QueryByAttribute query = new QueryByAttribute();
            query.ColumnSet = new ColumnSet(new string[] { "systemuserid", "fullname" });
            query.EntityName = "systemuser";
            query.Attributes.Add("domainname"); // TODO: Should be changed
            query.Values.Add(domainname);

            RetrieveMultipleRequest request = new RetrieveMultipleRequest();
            request.Query = query;

            RetrieveMultipleResponse response = (RetrieveMultipleResponse)OrganizationService.Execute(request);
            Entity results = (Entity)response.EntityCollection.Entities[0];

            try
            {
                Guid lookupId = new Guid();
                lookupId = results.Id;
                return lookupId;
            }

            catch (FaultException<OrganizationServiceFault> ex)
            {
                SetExceptionDetails(System.Reflection.MethodBase.GetCurrentMethod().Name, 2);
                string message = String.Format("An error occurred in the {0} function. This function was called from {1} of the {2} plugin. The error that was returned was {3}", ExceptionMethodName, CallingMethodName, CallingClassName, ex.Message);
                throw new InvalidPluginExecutionException(message);
            }
        }

        public static Guid RetrieveUserIdByEmailAddress(string emailAddress)
        {
            QueryByAttribute query = new QueryByAttribute();
            query.ColumnSet = new ColumnSet(new string[] { "systemuserid", "fullname" });
            query.EntityName = "systemuser";
            query.Attributes.Add("internalemailaddress");
            query.Values.Add(emailAddress);

            RetrieveMultipleRequest request = new RetrieveMultipleRequest();
            request.Query = query;

            RetrieveMultipleResponse response = (RetrieveMultipleResponse)OrganizationService.Execute(request);
            Entity results = (Entity)response.EntityCollection.Entities[0];

            try
            {
                Guid lookupId = new Guid();
                lookupId = results.Id;
                return lookupId;
            }

            catch (FaultException<OrganizationServiceFault> ex)
            {
                SetExceptionDetails(System.Reflection.MethodBase.GetCurrentMethod().Name, 2);
                string message = String.Format("An error occurred in the {0} function. This function was called from {1} of the {2} plugin. The error that was returned was {3}", ExceptionMethodName, CallingMethodName, CallingClassName, ex.Message);
                throw new InvalidPluginExecutionException(message);
            }
        }


        /// <summary>
        /// Retrieve Guid of CRM Team based on their abbreviated name
        /// </summary>
        public static Guid RetrieveTeamIdByName(string name)
        {
            QueryByAttribute query = new QueryByAttribute();
            query.ColumnSet = new ColumnSet(new string[] { "teamid", "name" });
            query.EntityName = "team";
            query.AddAttributeValue("name", name);

            RetrieveMultipleRequest request = new RetrieveMultipleRequest();
            request.Query = query;

            RetrieveMultipleResponse response = (RetrieveMultipleResponse)OrganizationService.Execute(request);
            Entity results = (Entity)response.EntityCollection.Entities[0];

            try
            {
                Guid lookupId = new Guid();
                lookupId = results.Id;
                return lookupId;
            }

            catch (FaultException<OrganizationServiceFault> ex)
            {
                throw new InvalidPluginExecutionException(
                    String.Format("An error occurred in the RetrieveUserIdByDomainName function of the SetStateCommitteeMembership plug-in. The error received was: {0}", ex.Message));
            }
        }

        public static Guid RetrieveOwnerId(Guid entityid, string entityname, ref string entitytype)
        {
            ColumnSet columns = new ColumnSet(new string[] { "ownerid" });
            Entity dynamic = OrganizationService.Retrieve(entityname, entityid, columns);
            entitytype = ((EntityReference)(dynamic.Attributes["ownerid"])).LogicalName;
            Guid ownerid = ((EntityReference)(dynamic.Attributes["ownerid"])).Id;
            return ownerid;
        }


        public static string RetrieveSystemUserInfo(Guid systemUserId, bool excludeEmail = false)
        {
            string firstName = "", lastName = "", emailAddress = "";
            RetrieveSystemUserInfo(systemUserId, ref firstName, ref lastName, ref emailAddress);
            if (excludeEmail)
                return firstName + " " + lastName;
            else
                return firstName + " " + lastName + "[" + emailAddress + "]";
        }

        /// <summary>
        /// Retrieve the first name, last name and email address of CRM System User
        /// </summary>
        public static void RetrieveSystemUserInfo(Guid systemUserId, ref string firstName, ref string lastName, ref string emailAddress)
        {
            ColumnSet columns = new ColumnSet(new string[] { "firstname", "lastname", "internalemailaddress" });
            Entity systemuser = OrganizationService.Retrieve("systemuser", systemUserId, columns);
            firstName = systemuser.Attributes["firstname"].ToString();
            lastName = systemuser.Attributes["lastname"].ToString();
            emailAddress = systemuser.Attributes["internalemailaddress"].ToString();
        }

        public static string RetrieveTeamInfo(Guid teamId)
        {
            string teamName = "", teamEmailAddress = "";
            RetrieveTeamInfo(teamId, ref teamName, ref teamEmailAddress);
            return teamName + "[" + teamEmailAddress + "]";
        }

        /// <summary>
        /// Retrieves the name and email address of a team record
        /// </summary>
        public static void RetrieveTeamInfo(Guid teamid, ref string teamname, ref string emailaddress)
        {
            ColumnSet columns = new ColumnSet(new string[] { "name", "emailaddress" });
            Entity systemuser = OrganizationService.Retrieve("team", teamid, columns);
            teamname = systemuser.Attributes["name"].ToString();
            emailaddress = systemuser.Attributes["emailaddress"].ToString();
        }

        public static string RetrieveQueueInfo(Guid queueId)
        {
            string queueName = "", queueEmailAddress = "";
            RetrieveQueueInfo(queueId, ref queueName, ref queueEmailAddress);
            return queueName + "[" + queueEmailAddress + "]";
        }

        /// <summary>
        /// Retrieves the name and email address of a queue record
        /// </summary>
        public static void RetrieveQueueInfo(Guid queueId, ref string queueName, ref string emailAddress)
        {
            ColumnSet columns = new ColumnSet(new string[] { "name", "emailaddress" });
            Entity systemuser = OrganizationService.Retrieve("queue", queueId, columns);
            queueName = systemuser.Attributes["name"].ToString();
            emailAddress = systemuser.Attributes["emailaddress"].ToString();
        }

        /// <summary>
        /// Assigns ownership of a record to a new user or team
        /// </summary>
        /// <param name="entityname">The Name of the Entity</param>
        /// <param name="id">The unique Id of the record that is to be assigned</param>
        /// <param name="ownerid">The new owner of the record</param>
        /// <param name="ownertype">The type of owner of the record</param>
        public static void AssignOwnership(string entityname, Guid id, Guid ownerid, OwnerTypeCode ownertype)
        {
            AssignRequest request = new AssignRequest();
            request.Target = new EntityReference(entityname, id);
            request.Assignee = new EntityReference(ownertype.ToString().ToLower(), ownerid);
            AssignResponse response = (AssignResponse)OrganizationService.Execute(request);
        }

        public static void SetState(string entityname, Guid entityid, StateCode state)
        {
            EntityReference moniker = new EntityReference();
            moniker.LogicalName = entityname;
            moniker.Id = entityid;

            SetStateRequest request = new SetStateRequest();
            request.EntityMoniker = moniker;

            if (state == StateCode.Active)
            {
                request.State = new OptionSetValue((int)state); // 0
                request.Status = new OptionSetValue((int)state + 1); // 1

            }
            else if (state == StateCode.Inactive)
            {
                request.State = new OptionSetValue((int)state); // 1
                request.Status = new OptionSetValue((int)state + 1); // 2
            }

            try
            {
                SetStateResponse response = (SetStateResponse)OrganizationService.Execute(request);
            }
            catch (FaultException<OrganizationServiceFault> ex)
            {
                string errMessage = "";
                if (ex.Message.StartsWith("SecLib::AccessCheckEx"))
                    errMessage = "Your current security settings do not allow you to deactivate this record.";
                else
                    errMessage = ex.Message;

                SetExceptionDetails(System.Reflection.MethodBase.GetCurrentMethod().Name, 2);
                string message = String.Format("An error occurred in the {0} function. This function was called from {1} of the {2} plugin. The error that was returned was {3}", ExceptionMethodName, CallingMethodName, CallingClassName, errMessage);
                throw new InvalidPluginExecutionException(message);
            }
        }

        public static StateCode RetrieveEntityState(string entityName, Guid entityId)
        {
            Entity entity = OrganizationService.Retrieve(entityName, entityId, new ColumnSet("statecode"));
            StateCode stateCode = (StateCode)((OptionSetValue)(entity.Attributes["statecode"])).Value;
            return stateCode;
        }

        public static Entity RetrieveSystemMessageFields(string name)
        {
            string rc = "";
            QueryByAttribute query = new QueryByAttribute();
            query.ColumnSet = new ColumnSet("xrm_name", "xrm_description");
            query.EntityName = "xrm_systemmessage";
            query.Attributes.Add("xrm_name");
            query.Values.Add(name);

            RetrieveMultipleRequest request = new RetrieveMultipleRequest();
            request.Query = query;

            try
            {
                RetrieveMultipleResponse response = (RetrieveMultipleResponse)OrganizationService.Execute(request);
                EntityCollection results = response.EntityCollection;
                return results.Entities[0];
            }
            catch (FaultException<OrganizationServiceFault> ex)
            {
                throw new InvalidPluginExecutionException(
                    String.Format("An error occurred in the {0} function of the {1} plug-in.",
                       "RetrieveSystemMessageFields", "CRMHelper"), ex);
            }
        }

        public static string RetrieveWebResourceByName(string name)
        {
            string rc = "";
            QueryByAttribute query = new QueryByAttribute();
            query.ColumnSet = new ColumnSet("content");
            query.EntityName = "webresource";
            query.Attributes.Add("name");
            query.Values.Add(name);

            RetrieveMultipleRequest request = new RetrieveMultipleRequest();
            request.Query = query;

            try
            {
                RetrieveMultipleResponse response = (RetrieveMultipleResponse)OrganizationService.Execute(request);
                EntityCollection results = response.EntityCollection;

                Entity firstResult = results.Entities[0];
                if (firstResult.Contains("content"))
                {
                    byte[] encodedContent = Convert.FromBase64String(firstResult.Attributes["content"].ToString());
                    rc = UnicodeEncoding.UTF8.GetString(encodedContent);
                }
                    

                TracingService.Trace("Content: {0}", rc);
                return rc;
            }
            catch (FaultException<OrganizationServiceFault> ex)
            {
                throw new InvalidPluginExecutionException(
                    String.Format("An error occurred in the {0} function of the {1} plug-in.",
                       "RetrieveWebResourceByName", "CRMHelper"), ex);
            }
        }


        public static int GetRecordCount(string entityName, string attributeName, ConditionExpression expression)
        {
            int rc = 0;
            StringBuilder sb = new StringBuilder();
            sb.AppendLine("<fetch distinct='false' mapping='logical' aggregate='true'> ");
            sb.AppendLine(" <entity name='" + entityName + "'> ");
            sb.AppendLine("  <attribute name='" + attributeName + "' alias='" + attributeName + "_count' aggregate='count'/> ");
            sb.AppendLine("  <filter type='and'>");
            if (expression.Values.Count > 1)
            {
                sb.AppendLine("   <condition attribute='" + expression.AttributeName + "' operator='in'>");
                for (int i = 0; i < expression.Values.Count; i++)
                {
                    sb.AppendLine("     <value>" + expression.Values[i].ToString() + "</value>");
                }
                sb.AppendLine("   </condition>");
            }
            else
            {
                sb.AppendLine("   <condition attribute='" + expression.AttributeName + "' operator='" + expression.Operator.ToFetchXmlOperator() + "' value='" + expression.Values[0].ToString() + "' /> ");
            }
            sb.AppendLine("   <condition attribute='statecode' operator='eq' value='0' /> ");
            sb.AppendLine("  </filter> ");
            sb.AppendLine(" </entity> ");
            sb.AppendLine("</fetch>");

            EntityCollection totalCount = OrganizationService.RetrieveMultiple(new FetchExpression(sb.ToString()));
            if (totalCount.Entities.Count > 0)
            {
                foreach (var t in totalCount.Entities)
                {
                    rc = (int)((AliasedValue)t[attributeName + "_count"]).Value;
                    break;
                }
            }
            return rc;
        }

        public static void SetExceptionDetails(string currentMethod, int previousLevelCount)
        {
            System.Diagnostics.StackTrace trace = new System.Diagnostics.StackTrace();
            System.Diagnostics.StackFrame frame = trace.GetFrame(previousLevelCount);
            CallingMethodName = frame.GetMethod().Name;
            CallingClassName = frame.GetType().Name;
            ExceptionMethodName = currentMethod;
        }

        /// <summary>
        /// Converts Query Expressions Condition Operator to Fetch Xml Operator
        /// </summary>
        /// <param name="s"></param>
        /// <returns></returns>
        public static string ToFetchXmlOperator(this ConditionOperator s)
        {
            string rc = "";
            switch (s)
            {
                case ConditionOperator.Equal:
                    rc = "eq"; break;
                case ConditionOperator.NotEqual:
                    rc = "ne"; break;
                case ConditionOperator.Contains:
                    rc = "like"; break;
                case ConditionOperator.Like:
                    rc = "like"; break;
                case ConditionOperator.Null:
                    rc = "null"; break;
                case ConditionOperator.EqualUserId:
                    rc = "eq-userid"; break;
                case ConditionOperator.NotLike:
                    rc = "not-like"; break;
                case ConditionOperator.NotNull:
                    rc = "not-null"; break;
                case ConditionOperator.NotEqualUserId:
                    rc = "ne-userid"; break;
                case ConditionOperator.GreaterEqual:
                    rc = "ge"; break;
                case ConditionOperator.GreaterThan:
                    rc = "gt"; break;
                case ConditionOperator.LessEqual:
                    rc = "le"; break;
                case ConditionOperator.LessThan:
                    rc = "lt"; break;
                case ConditionOperator.In:
                    rc = "in"; break;
                case ConditionOperator.NotIn:
                    rc = "not-in"; break;
                case ConditionOperator.Between:
                    rc = "between"; break;
                case ConditionOperator.EqualBusinessId:
                    rc = "eq-businessid"; break;
                case ConditionOperator.NotEqualBusinessId:
                    rc = "ne-businessid"; break;
            }
            return rc;
        }

        public static string GenerateEntityUrl(string organizationName, string entityName, Guid entityId)
        {
            StringBuilder sb = new StringBuilder();
            sb.Append(organizationName);
            if (!organizationName.EndsWith(@"/"))
                sb.Append(@"/");

            sb.Append("main.aspx?pagetype=entityrecord");
            sb.Append("&etn=" + entityName);
            sb.Append("&id=" + entityId.ToString());

            return sb.ToString();
        }
    }

    public static class License
    {
        public static string LicenseKey
        {
            get
            {
                return "%%watermark0%%";
            }
        }

        public static string CompanyName
        {
            get
            {
                return "%%watermark1%%";
            }
        }

        public static string CompanyWebSite
        {
            get
            {
                return "%%watermark2%%";
            }
        }

        public static string ExpirationDateTick
        {
            get
            {
                return "%%watermark3%%";
            }
        }

    }

    public class MailService
    {

        public string Domain { get; set; }
        public string UserName { get; set; }
        public string Password { get; set; }
        public string MailServer { get; set; }
        public int PortNumber { get; set; }

        public int SendEmail(string From, string To, string Cc, string Bcc, string Subject, string Body)
        {
            MailMessage message = new MailMessage();
            int BracketPos = 0; string eMailAddress = "", DisplayName = "";

            BracketPos = From.IndexOf("[");
            if (BracketPos > 0)
            {
                eMailAddress = From.Substring(BracketPos + 1, From.Length - BracketPos - 2);
                DisplayName = From.Substring(0, BracketPos);
                message.From = new MailAddress(eMailAddress, DisplayName);
            }
            else
                message.From = new MailAddress(From, From);

            BracketPos = To.IndexOf("[");
            if (BracketPos > 0)
            {
                eMailAddress = To.Substring(BracketPos + 1, To.Length - BracketPos - 2);
                DisplayName = To.Substring(0, BracketPos);
                message.To.Add(new MailAddress(eMailAddress, DisplayName));
            }
            else
                message.To.Add(new MailAddress(To, To));

            if (!string.IsNullOrEmpty(Cc))
            {
                BracketPos = Cc.IndexOf("[");
                if (BracketPos > 0)
                {
                    eMailAddress = Cc.Substring(BracketPos + 1, Cc.Length - BracketPos - 2);
                    DisplayName = Cc.Substring(0, BracketPos);
                    message.CC.Add(new MailAddress(eMailAddress, DisplayName));
                }
                else
                    message.CC.Add(new MailAddress(Cc, Cc));
            }
            if (!string.IsNullOrEmpty(Bcc))
            {
                BracketPos = Bcc.IndexOf("[");
                if (BracketPos > 0)
                {
                    eMailAddress = Bcc.Substring(BracketPos + 1, Bcc.Length - BracketPos - 2);
                    DisplayName = Bcc.Substring(0, BracketPos);
                    message.Bcc.Add(new MailAddress(eMailAddress, DisplayName));
                }
                else
                    message.Bcc.Add(new MailAddress(Bcc, Bcc));
            }

            message.Subject = Subject;
            message.Body = Body;
            message.IsBodyHtml = true;

            SmtpClient client = new SmtpClient(MailServer, 25);
            client.DeliveryMethod = SmtpDeliveryMethod.Network;
            string user = Domain + @"\" + UserName;
            client.Credentials = new System.Net.NetworkCredential(user, Password);

            int retVal = 0;
            try
            {
                client.Send(message);
            }
            catch (System.Exception ex)
            {

            }

            return retVal;
        }
    }

}