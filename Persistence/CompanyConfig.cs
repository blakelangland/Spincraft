using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;


namespace HorizonScientific
{
    public class CompanyConfiguration
    {
        public string CompanyId;

        public static CompanyConfiguration DefaultCompanyConfiguration
        {
            get { return g_oDefaultCompanyConfiguration; }
            set { g_oDefaultCompanyConfiguration = value; }
        }

        private static CompanyConfiguration g_oDefaultCompanyConfiguration;

        public static string SPINCRAFT_UK_COMPANY_ID = "51515";
        public static string SPINCRAFT_MA_COMPANY_ID = "51504";
        public static string SPINCRAFT_WI_COMPANY_ID = "51503";
        public static string SPINCRAFT_CA_COMPANY_ID = "MCS";
    }
}