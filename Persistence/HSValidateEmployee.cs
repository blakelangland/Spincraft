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

using SpreadsheetLight;


namespace HorizonScientific
{
    public class HSEmployeeClockedIn
    {
        #region Constructors

        public HSEmployeeClockedIn(DataRow oDataRow)
        {
            if ((oDataRow["EmpBasic_FirstName"] != DBNull.Value) && (string.IsNullOrEmpty((string)oDataRow["EmpBasic_FirstName"]) == false))
            {
                m_sFirstName = (string)oDataRow["EmpBasic_FirstName"];
            }
            if ((oDataRow["EmpBasic_LastName"] != DBNull.Value) && (string.IsNullOrEmpty((string)oDataRow["EmpBasic_LastName"]) == false))
            {
                m_sLastName = (string)oDataRow["EmpBasic_LastName"];
            }
            if (oDataRow["Calculated_CurrentPayrollDate"] != DBNull.Value)
            {
                m_dtPayrollDate = (DateTime)oDataRow["Calculated_CurrentPayrollDate"];
            }
            if ((oDataRow["EmpBasic_EMailAddress"] != DBNull.Value) && (string.IsNullOrEmpty((string)oDataRow["EmpBasic_EMailAddress"]) == false))
            {
                m_sEmail = (string)oDataRow["EmpBasic_EMailAddress"];
            }
            if ((oDataRow["EmpBasic_SupervisorID"] != DBNull.Value) && (string.IsNullOrEmpty((string)oDataRow["EmpBasic_SupervisorID"]) == false))
            {
                m_sSupervisorId = (string)oDataRow["EmpBasic_SupervisorID"];
            }
        }

        #endregion

        #region Properties

        public string FirstName
        {
            get { return m_sFirstName; }
        }
        public string LastName
        {
            get { return m_sLastName; }
        }
        public DateTime PayrollDate
        {
            get { return m_dtPayrollDate; }
        }
        public string Email
        {
            get { return m_sEmail; }
        }
        public string SupervisorId
        {
            get { return m_sSupervisorId; }
        }

        #endregion

        #region Data Members

        private string m_sFirstName;
        private string m_sLastName;
        private DateTime m_dtPayrollDate;
        private string m_sEmail;
        private string m_sSupervisorId;

        #endregion
    }

    public class HSValidateEmployee
    {
        #region Constructors

        public HSValidateEmployee(DataRow oDataRow)
        {
        }

        #endregion

        #region Methods

        static public bool Initialize(Session oSession)
        {
            bool bSuccess = true;

            //****Set a parameter Value***** 
            Ice.Proxy.BO.DynamicQueryImpl oDynamicQuery = WCFServiceSupport.CreateImpl<Ice.Proxy.BO.DynamicQueryImpl>(oSession, Ice.Proxy.BO.DynamicQueryImpl.UriPath);
            QueryExecutionDataSet oQueryExecutionDataSet = oDynamicQuery.GetQueryExecutionParametersByID(QUERY_EMPLOYEES_CLOCKED_IN);
            oQueryExecutionDataSet.ExecutionParameter.Clear();
            DataSet oDataSet = oDynamicQuery.ExecuteByID(QUERY_EMPLOYEES_CLOCKED_IN, oQueryExecutionDataSet);

            //foreach (DataColumn x in oDataSet.Tables[0].Columns)
            //{
            //    Console.WriteLine("ColumnName = :" + x.ColumnName);
            //}

            g_oEmployeesClockedIn.Clear();
            foreach (DataRow oRow in oDataSet.Tables[0].Rows)
            {
                g_oEmployeesClockedIn.Add(new HSEmployeeClockedIn(oRow));
            }

            return bSuccess;
        }

        static public void PerformValidation(string sTmpFileDirectory)
        {
            DateTime dtToday = DateTime.Now;
            string sDestinationFileName = sTmpFileDirectory + "EmployeeValidation-" + dtToday.Month.ToString() + "-" + dtToday.Day.ToString() + "-" + dtToday.Year.ToString() + ".xlsx";

            // if the file is already there then get rid of it as we are reprocessing for some reason
            if (File.Exists(sDestinationFileName) == true)
            {
                File.Delete(sDestinationFileName);
            }

            HSUser oServiceAccount = HSUser.GetUserById(HSUser.SPINCRAFT_SERVICE_ACCOUNT_ID);

            //
            // Production Manager issues
            //
            SLDocument oSLProductionDocument = new SLDocument();
            SLStyle oClockedInTooLong = oSLProductionDocument.CreateStyle();
            oClockedInTooLong.ApplyNamedCellStyle(SLNamedCellStyleValues.Bad);

            List<string> oToAddresses = HSUser.GetEmailsForUsersInGroup(HSUser.REPORT_ON_EMPLOYEES_CLOCKED_IN);
            oToAddresses.Add(HSUser.SPINCRAFT_ROOT_USER);
            bool bFirstWorksheet = true;
            if (HSValidateEmployee.EmployeesClockedIn.Count != 0)
            {
                if (bFirstWorksheet == true)
                {
                    oSLProductionDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "Clocked In");
                    bFirstWorksheet = false;
                }
                else
                {
                    oSLProductionDocument.AddWorksheet("Clocked In");
                }

                //set up column headers
                int iNumOfColumns = 0;
                iNumOfColumns++;
                oSLProductionDocument.SetCellValue(1, iNumOfColumns, "First Name");
                oSLProductionDocument.SetColumnWidth(iNumOfColumns, 15);
                iNumOfColumns++;
                oSLProductionDocument.SetCellValue(1, iNumOfColumns, "Last Name");
                oSLProductionDocument.SetColumnWidth(iNumOfColumns, 15);
                iNumOfColumns++;
                oSLProductionDocument.SetCellValue(1, iNumOfColumns, "Clock In Date");
                oSLProductionDocument.SetColumnWidth(iNumOfColumns, 15);
                iNumOfColumns++;
                oSLProductionDocument.SetCellValue(1, iNumOfColumns, "Employees Still Clocked In");
                oSLProductionDocument.SetColumnWidth(iNumOfColumns, 30);

                // check for excess production
                int iNumOfRows = 1;
                foreach (HSEmployeeClockedIn oEmployeeClockedIn in HSValidateEmployee.EmployeesClockedIn)
                {
                    iNumOfRows++;
                    oSLProductionDocument.SetCellValue(iNumOfRows, 1, oEmployeeClockedIn.FirstName);
                    oSLProductionDocument.SetCellValue(iNumOfRows, 2, oEmployeeClockedIn.LastName);
                    oSLProductionDocument.SetCellValue(iNumOfRows, 3, oEmployeeClockedIn.PayrollDate.ToShortDateString());
                    if (oEmployeeClockedIn.PayrollDate < DateTime.Now.AddDays(-3))
                    {
                        oSLProductionDocument.SetCellStyle(iNumOfRows, 3, oClockedInTooLong);
                    }
                }
            }
            // send the email to the production manager
            if (bFirstWorksheet == false)
            {
                // we create a worksheet so we have some data in the spreadsheet so we email the spreadsheet
                oSLProductionDocument.SaveAs(sDestinationFileName);
                List<string> oAttachments = new List<string>();
                oAttachments.Add(sDestinationFileName);
                HSEmailHelper.SendEmail(oToAddresses, "Employees Clocked In", "Employees Clocked In", oAttachments);
            }
        }

        #endregion

        #region Properties

        //
        // Employees Clocked In
        //
        public static List<HSEmployeeClockedIn> EmployeesClockedIn
        {
            get
            {
                return g_oEmployeesClockedIn;
            }
        }

        #endregion

        #region Data Members

        // Employee Issues
        private static List<HSEmployeeClockedIn> g_oEmployeesClockedIn = new List<HSEmployeeClockedIn>();

        private static string QUERY_EMPLOYEES_CLOCKED_IN = "EmployeesClockedIn";

        #endregion
    }
}
