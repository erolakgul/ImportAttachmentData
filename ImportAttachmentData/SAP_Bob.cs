using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ImportAttachmentData
{
    public class SAP_Bob
    {
        public SAPbobsCOM.Company oCompany = new SAPbobsCOM.Company();

        private readonly string _SAPBobCompany;
        private readonly string _SAPBobServer;
        private readonly string _SAPBobLicenceServer;

        private readonly string _SAPBobUsername;
        private readonly string _SAPBobPassword;

        private readonly string _SAPBobDbUser;
        private readonly string _SAPBobDbPassword;

        private readonly bool _SAPBobTrusted;
        

        public SAP_Bob()
        {
            AppConfig Config = new AppConfig();
            Config.getSettings();

            _SAPBobCompany = Config.SAPBobCompany;
            _SAPBobServer = Config.SAPBobServer;
            _SAPBobUsername = Config.SAPBobUsername;
            _SAPBobPassword = Config.SAPBobPassword;
            _SAPBobTrusted = Config.SAPBobTrusted;
            _SAPBobLicenceServer = Config.SAPBobLicenceServer;
            _SAPBobDbUser = Config.SAPBobDbUser;
            _SAPBobDbPassword = Config.SAPBobDbPassword;
        }

        public SAPbobsCOM.Company _getCompanyConnection()
        {
            oCompany = new SAPbobsCOM.Company();

            oCompany.Server = _SAPBobServer;
            oCompany.LicenseServer = _SAPBobLicenceServer;
            //oCompany.SLDServer = "srvhausb1app:40000";

            oCompany.CompanyDB = _SAPBobCompany;
            oCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_MSSQL2016; //dst_MSSQL2016
            //oCompany.language = SAPbobsCOM.BoSuppLangs.ln_Turkish_Tr;

            oCompany.UserName = _SAPBobUsername;
            oCompany.Password = _SAPBobPassword;

            //oCompany.DbUserName = _SAPBobDbUser;
            //oCompany.DbPassword = _SAPBobDbPassword;

            //oCompany.UseTrusted = _SAPBobTrusted;
            
            oCompany.Connect();

            return oCompany;
        }
    }
}
