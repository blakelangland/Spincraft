using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using System.IO;

using Ice.Core;
using Erp.BO;
using Erp.Proxy.BO;
using Ice.Lib.Framework;
using Ice.Adapters;
using Ice.BO;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Drawing.Spreadsheet;
using SpreadsheetLight;
using HSPersistence;
using SpreadsheetLight.Charts;
using Ice.Lib.Searches;
using System.Net;

namespace HorizonScientific
{
    public class HSCheckHead
    {
        public HSCheckHead(DataRow oDataRow)
        {
            if (oDataRow["CheckHed_BankAcctID"] != DBNull.Value)
            {
                m_sBankAccountId = (string)oDataRow["CheckHed_BankAcctID"];
            }
            if (oDataRow["BankAcct_CheckingAccount"] != DBNull.Value)
            {
                m_sCheckingAccount = (string)oDataRow["BankAcct_CheckingAccount"];
            }
            if (oDataRow["CheckHed_CheckNum"] != DBNull.Value)
            {
                m_iCheckNum = (int)oDataRow["CheckHed_CheckNum"];
            }
            if (oDataRow["CheckHed_CheckDate"] != DBNull.Value)
            {
                m_dtCheckDate = (DateTime)oDataRow["CheckHed_CheckDate"];
            }
            if (oDataRow["CheckHed_CheckAmt"] != DBNull.Value)
            {
                m_dPaymentAmount = (decimal)oDataRow["CheckHed_CheckAmt"];
            }
            if (oDataRow["Vendor_Name"] != DBNull.Value)
            {
                m_sVendorName = (string)oDataRow["Vendor_Name"];
            }
        }

        public string BankAccountId
        {
            get { return m_sBankAccountId; }
            set { m_sBankAccountId = value; }
        }

        public string CheckingAccount
        {
            get { return m_sCheckingAccount; }
            set { m_sCheckingAccount = value; }
        }

        public int CheckNum
        {
            get { return m_iCheckNum; }
            set { m_iCheckNum = value; }
        }

        public DateTime CheckDate
        {
            get { return m_dtCheckDate; }
            set { m_dtCheckDate = value; }
        }

        public decimal PaymentAmount
        {
            get { return m_dPaymentAmount; }
            set { m_dPaymentAmount = value; }
        }

        public string VendorName
        {
            get { return m_sVendorName; }
            set { m_sVendorName = value; }
        }

        private string m_sBankAccountId;
        private string m_sCheckingAccount;
        private int m_iCheckNum;
        private DateTime m_dtCheckDate;
        private decimal m_dPaymentAmount;
        private string m_sVendorName;
    }

    public class HSPositivePay
    {
        static public bool Initialize(Session oSession, int iStartCheckNum, int iEndCheckNum, DateTime dtStartCheckDate, DateTime dtEndCheckDate)
        {
            bool bSuccess = true;

            // now we will get the details for any checks within the stated parameters
            g_oCheckHeadRecords.Clear();

            // then we need to query each sales order to determine its status
            Ice.Proxy.BO.DynamicQueryImpl oDynamicQuery = WCFServiceSupport.CreateImpl<Ice.Proxy.BO.DynamicQueryImpl>(oSession, Ice.Proxy.BO.DynamicQueryImpl.UriPath);
            QueryExecutionDataSet oQueryExecutionDataSet = oDynamicQuery.GetQueryExecutionParametersByID(BAQConstants.QUERY_POSITIVE_PAY);
            oQueryExecutionDataSet.Clear();
            if (iStartCheckNum != 0)
            {
                oQueryExecutionDataSet.ExecutionParameter.AddExecutionParameterRow("StartCheckNum", iStartCheckNum.ToString(), "int", false, Guid.NewGuid(), "A");
            }
            if (iEndCheckNum != 0)
            {
                oQueryExecutionDataSet.ExecutionParameter.AddExecutionParameterRow("EndCheckNum", iEndCheckNum.ToString(), "int", false, Guid.NewGuid(), "A");
            }
            if (dtStartCheckDate != DateTime.MinValue)
            {
                oQueryExecutionDataSet.ExecutionParameter.AddExecutionParameterRow("StartCheckDate", dtStartCheckDate.ToShortDateString(), "date", false, Guid.NewGuid(), "A");
            }
            if (dtEndCheckDate != DateTime.MinValue)
            {
                oQueryExecutionDataSet.ExecutionParameter.AddExecutionParameterRow("EndCheckDate", dtEndCheckDate.ToShortDateString(), "date", false, Guid.NewGuid(), "A");
            }
            oQueryExecutionDataSet.AcceptChanges();
            DataSet oDataSet = oDynamicQuery.ExecuteByID(BAQConstants.QUERY_POSITIVE_PAY, oQueryExecutionDataSet);
            foreach (DataRow oRow in oDataSet.Tables[0].Rows)
            {
                g_oCheckHeadRecords.Add(new HSCheckHead(oRow));
            }

            return bSuccess;
        }

        static public void GenerateFile(string sArchiveFileDirectory, string sTmpFileDirectory, string sCheckNumberPrefix, HSUser oRequestingUser)
        {
            #region General Info
            bool bGeneratedData = false;

            DateTime dtToday = DateTime.Now;
            string sDestinationFileName = sTmpFileDirectory + "ETG_PositivePay-" + dtToday.Month.ToString() + "-" + dtToday.Day.ToString() + "-" + dtToday.Year.ToString() + ".txt";
            string sArchiveFileName = sArchiveFileDirectory + "ETG_PositivePay-" + dtToday.Month.ToString() + "-" + dtToday.Day.ToString() + "-" + dtToday.Year.ToString() + ".txt";

            // if the file is already there then get rid of it as we are reprocessing for some reason
            if (File.Exists(sDestinationFileName) == true)
            {
                try
                {
                    File.Delete(sDestinationFileName);
                }
                catch (Exception e)
                {
                    // we will ingore this if we cannot delete the file
                }
            }

            HSUser oServiceAccount = HSUser.GetUserById(HSUser.SPINCRAFT_SERVICE_ACCOUNT_ID);
            List<string> oToAddresses = new List<string>();
            if (oRequestingUser == null)
            {
                HSUser.GetEmailsForUsersInGroup(oToAddresses, HSUser.REPORT_ON_POSITIVE_PAY);
            }
            else
            {
                // this report was requested by someone directly from Epicor
                oToAddresses.Add(oRequestingUser.Email);
            }
            oToAddresses.Add(HSUser.SPINCRAFT_ROOT_USER);

            #endregion

            #region Data Output
            if (g_oCheckHeadRecords.Count > 0)
            {
                string sDelimeter = string.Empty;
                // could be comma?
                int iLengthOfAccount = 17;
                int iLengthOfCheckNum = 10;
                int iLengthOfPaymentAmount = 10;
                // date = yymmdd
                int iLengthOfVendorName = 80;

                // RECORD TYPE (1 character -- hard coded to "6")
                string sFixedPrefixForCheck = "6";
                // STATUS CODE (1 character, 2 = Add Issue, 4 = Void Issue)
                sFixedPrefixForCheck += "2";
                // ACCOUNT NUMBER LENGTH (2 characters -- hard coded to "17")
                sFixedPrefixForCheck += iLengthOfAccount.ToString();
                // FILLER (4 spaces)
                sFixedPrefixForCheck += "    ";
                // BANK NUMBER LENGTH (4 characters - hard coded to "3037")
                sFixedPrefixForCheck += "3037";
                // FILLER (10 spaces)
                sFixedPrefixForCheck += "          ";

                // build up posfix string after check info
                string sPayerIdAndRegionDeptNumber = "";
                // PAYER ID NUMBER (15 characters, not used)
                // REGION DEPT NUMBER (40 characters, not used)
                int iNumberOfSpacesUntilVendorName = 55; // we need 55 blank spaces until the vendor name is inserted
                sPayerIdAndRegionDeptNumber = sPayerIdAndRegionDeptNumber.PadLeft(iNumberOfSpacesUntilVendorName, ' ');

                StringBuilder oOutputString = new StringBuilder();
                using (StreamWriter oOutput = new StreamWriter(sDestinationFileName))
                {
                    // loop through all checks and write the checks according to the specification
                    foreach (HSCheckHead oCheckHead in g_oCheckHeadRecords)
                    {
                        // write out the fixed prefix for each check line
                        oOutputString.Append(sFixedPrefixForCheck);

                        // build the account number up
                        // ACCOUNT NUMBER (however many spaces were defined above, currently 17)
                        string sAccountNumber = "";
                        // temp fix to force use of Santanders account number
                        int iNumberOfPaddedZeros = iLengthOfAccount - oCheckHead.CheckingAccount.Length;
                        if (iNumberOfPaddedZeros < 0)
                        {
                            iNumberOfPaddedZeros = 0;
                        }
                        for (int iCounter = 0; iCounter < iNumberOfPaddedZeros; iCounter++)
                        {
                            sAccountNumber = sAccountNumber + "0";
                        }
                        sAccountNumber += oCheckHead.CheckingAccount;
                        oOutputString.Append(sAccountNumber);

                        // build up the check number
                        // SERIAL NUMBER (10 characters -- check number)
                        string sCheckNum = oCheckHead.CheckNum.ToString();
                        // hardcoding the first 5 digits of the checknumber for both MA and WI
                        sCheckNum = sCheckNumberPrefix + sCheckNum;
                        if (iLengthOfCheckNum > sCheckNum.Length)
                        {
                            sCheckNum = sCheckNum.PadLeft(iLengthOfCheckNum, '0');
                        }
                        oOutputString.Append(sCheckNum);

                        // build up the amount
                        // ISSUE AMOUNT (10 characters - no periods)
                        int iCheckAmount = (int)(oCheckHead.PaymentAmount * 100);
                        string sCheckAmount = iCheckAmount.ToString();
                        if (iLengthOfPaymentAmount > sCheckAmount.Length)
                        {
                            sCheckAmount = sCheckAmount.PadLeft(iLengthOfPaymentAmount, '0');
                        }
                        oOutputString.Append(sCheckAmount);

                        // build up the date --yymmdd
                        // ISSUE DATE (6 characters, YYMMDD)
                        string sYear = oCheckHead.CheckDate.ToString("yy");
                        string sMonth = oCheckHead.CheckDate.ToString("MM");
                        string sDay = oCheckHead.CheckDate.ToString("dd");
                        oOutputString.Append(sYear + sMonth + sDay);

                        // add the info for payer Id and region dept number 
                        // -- this is blank currently and 55 characters so 
                        // -- we create this in advance above
                        oOutputString.Append(sPayerIdAndRegionDeptNumber);

                        // build up the vendor name
                        // PAYEE NAME (80 characters)
                        string sVendorName = "";
                        if (oCheckHead.VendorName != null)
                        {
                            sVendorName = oCheckHead.VendorName.ToUpper();
                        }
                        // remove any commas in vendor name
                        sVendorName = sVendorName.Replace(',', ' ');
                        if (iLengthOfVendorName > sVendorName.Length)
                        {
                            sVendorName = sVendorName.PadRight(iLengthOfVendorName, ' ');
                        }
                        oOutputString.Append(sVendorName);

                        // terminate the line for this check
                        //oOutputString.Append(",");
                        oOutputString.Append("\r\n");
                    }

                    // write out the footer
                    //oOutputString.Append(sFixedFooterForFile);

                    // finally write out the file
                    oOutput.Write(oOutputString.ToString());
                    bGeneratedData = true;
                }
            }

            #endregion

            // send the email
            if (bGeneratedData == true)
            {
                // we will place a copy of this in the archive folder if one does not already exist there
                // this file will be used to help measure KPIs at the end of the day
                if (File.Exists(sArchiveFileName) == false)
                {
                    // cant copy in this environment
                    //File.Copy(sDestinationFileName, sArchiveFileName);
                }

                List<string> oAttachments = new List<string>();
                oAttachments.Add(sDestinationFileName);
                string sBody = "Positive Pay File";
                HSEmailHelper.SendEmail(oToAddresses, "Positive Pay", sBody, oAttachments);
            }
        }

        private static List<HSCheckHead> g_oCheckHeadRecords = new List<HSCheckHead>();
    }
}

