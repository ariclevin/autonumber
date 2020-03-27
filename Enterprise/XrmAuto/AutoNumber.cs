namespace BriteGlobal.Xrm.Plugins.Auto
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.ServiceModel;
    using System.Text;
    using System.Threading.Tasks;
    using Microsoft.Xrm.Sdk;
    using Microsoft.Xrm.Sdk.Query;


    using BriteGlobal.Xrm.Plugins.Auto.Logic;

    
    public class AutoNumber : Plugin
    {
        public AutoNumber()
            : base(typeof(AutoNumber))
        {
            this.RegisteredEvents.Add(new Tuple<int, string, string, Action<LocalPluginContext>>(40, "Create", "xrm_autonumber", new Action<LocalPluginContext>(ExecutePostAutoNumberCreate)));
            this.RegisteredEvents.Add(new Tuple<int, string, string, Action<LocalPluginContext>>(40, "xrm_AutoFill", "xrm_autonumber", new Action<LocalPluginContext>(ExecutePostAutoNumberFill)));
        }


        protected void ExecutePostAutoNumberCreate(LocalPluginContext localContext)
        {
            if (localContext == null)
            {
                throw new ArgumentNullException("localContext");
            }

            string errMessage = "";
            bool isLicensed = Helper.ValidateLicense(out errMessage);

            if (isLicensed)
            {
                AutoNumberLogic logic = new AutoNumberLogic(localContext.OrganizationService);
                logic.CreateAutoNumber(localContext.PluginExecutionContext.PrimaryEntityName, localContext.PluginExecutionContext.PrimaryEntityId);
            }
        }

        protected void ExecutePostAutoNumberFill(LocalPluginContext localContext)
        {
            if (localContext == null)
            {
                throw new ArgumentNullException("localContext");
            }

            string errMessage = "";
            bool isLicensed = Helper.ValidateLicense(out errMessage);

            if (isLicensed)
            {
                AutoNumberLogic logic = new AutoNumberLogic(localContext.OrganizationService);
                logic.FillAutoNumber(localContext.PluginExecutionContext.PrimaryEntityName, localContext.PluginExecutionContext.PrimaryEntityId);
            }

        }

    }
}
