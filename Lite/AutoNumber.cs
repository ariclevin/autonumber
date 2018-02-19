namespace BriteGlobal.Xrm.Plugins.AutoNumber
{
    using System;
    using System.Collections;
    using System.Collections.Generic;
    using System.Linq;
    using System.ServiceModel;
    using System.Text;
    using System.Threading.Tasks;
    using System.Xml;
    using Microsoft.Xrm.Sdk;
    using Microsoft.Xrm.Sdk.Query;
    
    using BriteGlobal.Xrm.Plugins.AutoNumber.Logic;

    public class PluginEntryPoint : PluginBase
    {
        LicenseMode licenseMode = LicenseMode.Unavailable;

        /// <summary>
        /// Initializes a new instance of the <see cref="PluginEntryPoint"/> class.
        /// </summary>
        /// <param name="unsecure">Contains public (unsecured) configuration information.</param>
        /// <param name="secure">Contains non-public (secured) configuration information. 
        /// When using Microsoft Dynamics CRM for Outlook with Offline Access, 
        /// the secure string is not passed to a plug-in that executes while the client is offline.</param>
        public PluginEntryPoint(string unsecure, string secure)
            : base(typeof(PluginEntryPoint))
        {

            // TODO: Implement your custom configuration handling.
        }


        /// <summary>
        /// Main entry point for he business logic that the plug-in is to execute.
        /// </summary>
        /// <param name="localContext">The <see cref="LocalPluginContext"/> which contains the
        /// <see cref="IPluginExecutionContext"/>,
        /// <see cref="IOrganizationService"/>
        /// and <see cref="ITracingService"/>
        /// </param>
        /// <remarks>
        /// For improved performance, Microsoft Dynamics CRM caches plug-in instances.
        /// The plug-in's Execute method should be written to be stateless as the constructor
        /// is not called for every invocation of the plug-in. Also, multiple system threads
        /// could execute the plug-in at the same time. All per invocation state information
        /// is stored in the context. This means that you should not use global variables in plug-ins.
        /// </remarks>
        protected override void ExecuteCrmPlugin(LocalPluginContext localContext)
        {
            if (localContext == null)
            {
                throw new ArgumentNullException("localContext");
            }

            // TODO: Implement your custom plug-in business logic.

            if (localContext.PluginExecutionContext.PrimaryEntityName.ToLower() == "xrm_autonumber")
            {
                //CRMHelper.OrganizationService = localContext.OrganizationService;
                //CRMHelper.TracingService = localContext.TracingService;
                //string licenseXml = CRMHelper.RetrieveWebResourceByName("xrm_license.xml");

                //InitializeLicense(localContext.TracingService);
                //localContext.TracingService.Trace("Called InitializeLicense");
                //licenseMode = GetLicenseFromXml(licenseXml, localContext.TracingService);

                GetLicenseMode();
                localContext.TracingService.Trace(licenseMode.ToString());

                //localContext.TracingService.Trace(localContext.PluginExecutionContext.MessageName);
                switch (localContext.PluginExecutionContext.MessageName)
                {
                    case "Create":
                        ExecutePostAutoNumberCreate(localContext);
                        break;
                    case "xrm_PublishAutoNumberEntity":
                        ExecutePostAutoNumberPublish(localContext);
                        break;
                    case "xrm_UnpublishAutoNumberEntity":
                        break;
                    case "xrm_FillEntityAutoNumber":
                        ExecutePostAutoNumberFill(localContext);
                        break;
                }
            }
            else
            {
                if (localContext.PluginExecutionContext.MessageName == "Create")
                {
                    ExecutePostEntityCreate(localContext);
                }
            }
        }


        private void ExecutePostAutoNumberCreate(LocalPluginContext localContext)
        {
            if (licenseMode != LicenseMode.Unavailable)
            {
                // using (AutoNumberLogic logic = new AutoNumberLogic(localContext.OrganizationService, localContext.TracingService))
                //{
                //    logic.CreateAutoNumberLogic(localContext.PluginExecutionContext.PrimaryEntityName, localContext.PluginExecutionContext.PrimaryEntityId);
                //}
            }
        }

        private void ExecutePostEntityCreate(LocalPluginContext localContext)
        {
            using (AutoNumberLogic logic = new AutoNumberLogic(localContext.OrganizationService, localContext.TracingService))
            {
                logic.CreateEntityLogic(localContext.PluginExecutionContext.PrimaryEntityName, localContext.PluginExecutionContext.PrimaryEntityId, licenseMode);
            }
        }

        private void ExecutePostAutoNumberFill(LocalPluginContext localContext)
        {
            if (licenseMode != LicenseMode.Unavailable)
            {
                using (AutoNumberLogic logic = new AutoNumberLogic(localContext.OrganizationService, localContext.TracingService))
                    logic.FillAutoNumberLogic(localContext.PluginExecutionContext.PrimaryEntityName, localContext.PluginExecutionContext.PrimaryEntityId);
            }
        }

        private void ExecutePostAutoNumberPublish(LocalPluginContext localContext)
        {
            if (licenseMode != LicenseMode.Unavailable)
            {
                using (AutoNumberLogic logic = new AutoNumberLogic(localContext.OrganizationService, localContext.TracingService))
                    logic.AutoNumberPublishLogic(localContext.PluginExecutionContext.PrimaryEntityName, localContext.PluginExecutionContext.PrimaryEntityId);
            }
        }

        private void ExecutePostAutoNumberUnPublish(LocalPluginContext localContext)
        {
            // Unpublish should not remove from SDK Message Processing Step

        }

        #region Licensing Module

        /*
        CryptoLicense license;

        private void InitializeLicense(ITracingService tracingService)
        {
            tracingService.Trace("Before license = new CryptoLicense");
            license = new CryptoLicense();
            // Get by Pressing Ctrl+K in Crypto License Generator
            tracingService.Trace("Before license.ValidationKey");
            try
            {
                license.ValidationKey = "AMAAMADYLN6qRVjvUEWSzEw0UL1lN8bnWzRVTeB0YCgvKtr8mYotNNvZhpqRF0Ij/cKUN1MDAAEAAQ==";
            }
            catch (LicenseServiceException ex)
            {
                throw new InvalidPluginExecutionException(ex.Message);
            }
            catch (System.Exception ex)
            {
                throw new InvalidPluginExecutionException(ex.Message);
            }
            
            
            tracingService.Trace("After license.ValidationKey");
        }

        private LicenseMode GetLicenseFromXml(string xmlLicense, ITracingService tracingService)
        {
            XmlDocument doc = new XmlDocument();
            doc.Load(xmlLicense);

            tracingService.Trace("License Loaded to XmlDocument");

            XmlNode node = doc.DocumentElement.FirstChild;
            string licenseValue = node.InnerText;
            tracingService.Trace("License Value: {0}", licenseValue);

            license.LicenseCode = licenseValue;
            
            LicenseMode licenseMode = LicenseMode.Unavailable;
            if (license.Status == LicenseStatus.Valid)
            {
                string licenseType = GetDataField("LicenseType");
                switch (licenseType.ToLower())
                {
                    case "lite":
                        licenseMode = LicenseMode.Lite;
                        break;
                    case "enterprise":
                        licenseMode = LicenseMode.Enterprise;
                        break;
                    default:
                        break;
                }
            }

            return licenseMode;
        }

                    private string GetDataField(string fieldName)
        {
            Hashtable data = license.ParseUserData("#");
            string rc = data[fieldName].ToString();

            return rc;
        }

        */

        private void GetLicenseMode()
        {
            string licenseType = "lite"; // "%%watermark0%%";
            switch (licenseType.ToLower())
            {
                case "lite":
                    licenseMode = LicenseMode.Lite;
                    break;
                case "enterprise":
                    licenseMode = LicenseMode.Enterprise;
                    break;
                default:
                    break;
            }
        }




        #endregion




    }

}
