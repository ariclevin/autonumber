using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using LogicNP.CryptoLicensing;

namespace BriteGlobal.Xrm.Plugins.AutoNumber.Common
{
    public class Licensing
    {
        CryptoLicense license;
        public string LicenseFile { get; private set; }

        public Licensing()
        {
            license = new CryptoLicense();
            // Get by Pressing Ctrl+K in Crypto License Generator
            license.ValidationKey = "AMAAMADYLN6qRVjvUEWSzEw0UL1lN8bnWzRVTeB0YCgvKtr8mYotNNvZhpqRF0Ij/cKUN1MDAAEAAQ=="; 
        }

        public LicenseMode GetLicenseFromXml(string xmlLicense)
        {
            XmlDocument doc = new XmlDocument();
            doc.Load(xmlLicense);

            XmlNode node = doc.DocumentElement.FirstChild;
            string licenseValue = node.InnerText;

            LicenseFile = licenseValue;
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
    }
}
