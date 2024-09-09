using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ImportAttachmentData
{
    public class AppConfig
    {
        public string SAPBobCompany { get; set; }
        public string SAPBobServer { get; set; }
        public string SAPBobLicenceServer { get; set; }
        public string SAPBobUsername { get; set; }
        public string SAPBobPassword { get; set; }
        public string SAPBobDbUser { get; set; }
        public string SAPBobDbPassword { get; set; }

        public bool SAPBobTrusted { get; set; }

        // sap_bob yapıcı metotunda config değerleri alınır
        public void getSettings()
        {
            // App.config ayarları
            System.Configuration.AppSettingsReader appReader = new System.Configuration.AppSettingsReader();

            // SAP BOB oku
            SAPBobCompany = appReader.GetValue("SAPBobCompany", typeof(string)).ToString();
            SAPBobServer = appReader.GetValue("SAPBobServer", typeof(string)).ToString();
            SAPBobLicenceServer = appReader.GetValue("SAPBobLicenceServer", typeof(string)).ToString();
            SAPBobUsername = appReader.GetValue("SAPBobUsername", typeof(string)).ToString();
            SAPBobPassword = appReader.GetValue("SAPBobPassword", typeof(string)).ToString();
            SAPBobDbUser = appReader.GetValue("SAPBobDbUser", typeof(string)).ToString();
            SAPBobDbPassword = appReader.GetValue("SAPBobDbPassword", typeof(string)).ToString();

            SAPBobTrusted = (bool)appReader.GetValue("SAPBobTrusted", typeof(bool));
        }
    }
}
