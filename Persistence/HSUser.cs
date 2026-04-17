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
using Erp.Adapters;
using Ice.Tablesets;
using System.DirectoryServices;
using System.DirectoryServices.AccountManagement;

using SpreadsheetLight;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.ListView;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Drawing.Charts;
using Erp.UI;
using System.Runtime.CompilerServices;


namespace HorizonScientific
{
    public class ValidateEmployee
    {
        #region Constructors

        public ValidateEmployee(DataRow oDataRow)
        {
            if ((oDataRow["EmpBasic_EmpID"] != DBNull.Value) && (string.IsNullOrEmpty((string)oDataRow["EmpBasic_EmpID"]) == false))
            {
                m_sEmployeeId = (string)oDataRow["EmpBasic_EmpID"];
            }
            if ((oDataRow["EmpBasic_EmpStatus"] != DBNull.Value) && (string.IsNullOrEmpty((string)oDataRow["EmpBasic_EmpStatus"]) == false))
            {
                m_sStatus = (string)oDataRow["EmpBasic_EmpStatus"];
            }
            if ((oDataRow["EmpBasic_FirstName"] != DBNull.Value) && (string.IsNullOrEmpty((string)oDataRow["EmpBasic_FirstName"]) == false))
            {
                m_sFirstName = (string)oDataRow["EmpBasic_FirstName"];
            }
            if ((oDataRow["EmpBasic_LastName"] != DBNull.Value) && (string.IsNullOrEmpty((string)oDataRow["EmpBasic_LastName"]) == false))
            {
                m_sLastName = (string)oDataRow["EmpBasic_LastName"];
            }
            if ((oDataRow["EmpBasic_EMailAddress"] != DBNull.Value) && (string.IsNullOrEmpty((string)oDataRow["EmpBasic_EMailAddress"]) == false))
            {
                m_sEmail = (string)oDataRow["EmpBasic_EMailAddress"];
            }
            if ((oDataRow["EmpBasic_SupervisorID"] != DBNull.Value) && (string.IsNullOrEmpty((string)oDataRow["EmpBasic_SupervisorID"]) == false))
            {
                m_sSupervisorEmployeeId = (string)oDataRow["EmpBasic_SupervisorID"];
            }
            if (oDataRow["EmpBasic_Shift"] != DBNull.Value)
            {
                m_iShift = (int)oDataRow["EmpBasic_Shift"];
            }
            if ((oDataRow["EmpBasic_ExpenseCode"] != DBNull.Value) && (string.IsNullOrEmpty((string)oDataRow["EmpBasic_ExpenseCode"]) == false))
            {
                m_sExpenseCode = (string)oDataRow["EmpBasic_ExpenseCode"];
            }
        }

        #endregion

        #region Methods

        public static bool Initialize(Session oSession)
        {
            bool bSuccess = true;

            //****Set a parameter Value***** 
            Ice.Proxy.BO.DynamicQueryImpl oDynamicQuery = WCFServiceSupport.CreateImpl<Ice.Proxy.BO.DynamicQueryImpl>(oSession, Ice.Proxy.BO.DynamicQueryImpl.UriPath);
            QueryExecutionDataSet oQueryExecutionDataSet = oDynamicQuery.GetQueryExecutionParametersByID(QUERY_VALIDATE_EMPLOYEE);
            oQueryExecutionDataSet.ExecutionParameter.Clear();
            DataSet oDataSet = oDynamicQuery.ExecuteByID(QUERY_VALIDATE_EMPLOYEE, oQueryExecutionDataSet);

            g_oNoEmpSupervisor.Clear();
            g_oNoEmpDept.Clear();
            g_oBadEmpLaborRate.Clear();
            g_oNoEmpEmail.Clear();

            foreach (DataRow oDataRow in oDataSet.Tables[0].Rows)
            {
                ValidateEmployee oTmpUser = new ValidateEmployee(oDataRow);
                g_oAllEmployees.Add(oTmpUser);

                if ((oDataRow["Calculated_NoEmpSupervisor"] != DBNull.Value) && (string.IsNullOrEmpty((string)oDataRow["Calculated_NoEmpSupervisor"]) == false))
                {
                    g_oNoEmpSupervisor.Add(oTmpUser);
                }
                if ((oDataRow["Calculated_NoEmpDept"] != DBNull.Value) && (string.IsNullOrEmpty((string)oDataRow["Calculated_NoEmpDept"]) == false))
                {
                    g_oNoEmpDept.Add(oTmpUser);
                }
                if ((oDataRow["Calculated_EmpBadLaborRate"] != DBNull.Value) && (string.IsNullOrEmpty((string)oDataRow["Calculated_EmpBadLaborRate"]) == false))
                {
                    g_oBadEmpLaborRate.Add(oTmpUser);
                }
                if ((oDataRow["Calculated_NoEmpEmail"] != DBNull.Value) && (string.IsNullOrEmpty((string)oDataRow["Calculated_NoEmpEmail"]) == false))
                {
                    g_oNoEmpEmail.Add(oTmpUser);
                }
            }

            return bSuccess;
        }

        static public void PerformValidation(string sTmpFileDirectory)
        {
            // get the file name
            DateTime dtToday = DateTime.Now;
            string sDate = dtToday.Month.ToString() + "-" + dtToday.Day.ToString() + "-" + dtToday.Year.ToString();
            string sDestinationFileName = sTmpFileDirectory + "\\-EmployeeSetupReport-" + sDate + ".xlsx";

            // if the file is already there then get rid of it as we are reprocessing for some reason
            if (File.Exists(sDestinationFileName) == true)
            {
                File.Delete(sDestinationFileName);
            }

            HSUser oServiceAccount = HSUser.GetUserById(HSUser.SPINCRAFT_SERVICE_ACCOUNT_ID);

            List<string> oToAddresses = HSUser.GetEmailsForUsersInGroup(HSUser.REPORT_ON_EMPLOYEE_SETUP_ISSUES);
            oToAddresses.Add(HSUser.SPINCRAFT_ROOT_USER);

            SLDocument oSLDocument = new SLDocument();
            bool bFirstWorksheet = true;
            #region IT Issues
            //
            // IT Management issues
            //
            if (ValidateEmployee.NoEmpSupervisor.Count != 0)
            {
                if (bFirstWorksheet == true)
                {
                    oSLDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "No Employee Supervisor");
                    bFirstWorksheet = false;
                }
                else
                {
                    oSLDocument.AddWorksheet("No Employee Supervisor");
                }

                int iNumberOfRows = 1;
                int iNumberOfColumns = 1;
                //set column header
                oSLDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "User Id");
                oSLDocument.SetColumnWidth(iNumberOfRows++, iNumberOfColumns++, 30);

                foreach (ValidateEmployee oEmployee in ValidateEmployee.NoEmpSupervisor)
                {
                    oSLDocument.SetCellValue(iNumberOfRows++, 1, oEmployee.EmployeeId);
                }
            }
            if (ValidateEmployee.NoEmpDept.Count != 0)
            {
                if (bFirstWorksheet == true)
                {
                    oSLDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "No Employee Department");
                    bFirstWorksheet = false;
                }
                else
                {
                    oSLDocument.AddWorksheet("No Employee Department");
                }

                int iNumberOfRows = 1;
                int iNumberOfColumns = 1;
                //set column header
                oSLDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "User Id");
                oSLDocument.SetColumnWidth(iNumberOfRows++, iNumberOfColumns++, 30);

                foreach (ValidateEmployee oEmployee in ValidateEmployee.NoEmpDept)
                {
                    oSLDocument.SetCellValue(iNumberOfRows++, 1, oEmployee.EmployeeId);
                }
            }
            if (ValidateEmployee.BadEmpLaborRate.Count != 0)
            {
                if (bFirstWorksheet == true)
                {
                    oSLDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "No Labor Rate");
                    bFirstWorksheet = false;
                }
                else
                {
                    oSLDocument.AddWorksheet("No Labor Rate");
                }

                int iNumberOfRows = 1;
                int iNumberOfColumns = 1;
                //set column header
                oSLDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "User Id");
                oSLDocument.SetColumnWidth(iNumberOfRows++, iNumberOfColumns++, 30);

                foreach (ValidateEmployee oEmployee in ValidateEmployee.BadEmpLaborRate)
                {
                    oSLDocument.SetCellValue(iNumberOfRows++, 1, oEmployee.EmployeeId);
                }
            }
            if (ValidateEmployee.NoEmpEmail.Count != 0)
            {
                if (bFirstWorksheet == true)
                {
                    oSLDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "No Employee Email");
                    bFirstWorksheet = false;
                }
                else
                {
                    oSLDocument.AddWorksheet("No Employee Email");
                }

                int iNumberOfRows = 1;
                int iNumberOfColumns = 1;
                //set column header
                oSLDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "User Id");
                oSLDocument.SetColumnWidth(iNumberOfRows++, iNumberOfColumns++, 30);

                foreach (ValidateEmployee oEmployee in ValidateEmployee.NoEmpEmail)
                {
                    oSLDocument.SetCellValue(iNumberOfRows++, 1, oEmployee.EmployeeId);
                }
            }
            if (ValidateEmployee.BadEmpLaborRate.Count != 0)
            {
                if (bFirstWorksheet == true)
                {
                    oSLDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "No Hourly Rate");
                    bFirstWorksheet = false;
                }
                else
                {
                    oSLDocument.AddWorksheet("No Hourly Rate");
                }

                int iNumberOfRows = 1;
                int iNumberOfColumns = 1;
                //set column header
                oSLDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "User Id");
                oSLDocument.SetColumnWidth(iNumberOfRows++, iNumberOfColumns++, 30);

                foreach (ValidateEmployee oEmployee in ValidateEmployee.BadEmpLaborRate)
                {
                    oSLDocument.SetCellValue(iNumberOfRows++, 1, oEmployee.EmployeeId);
                }
            }
            // send the email
            if (bFirstWorksheet == false)
            {
                // we create a worksheet so we have some data in the spreadsheet so we email the spreadsheet
                oSLDocument.SaveAs(sDestinationFileName);
                List<string> oAttachments = new List<string>();
                oAttachments.Add(sDestinationFileName);
                HSEmailHelper.SendEmail(oToAddresses, "Employee Setup", "Employee Setup", oAttachments);
            }
            #endregion
        }
        #endregion

        #region Properties

        public static List<ValidateEmployee> NoEmpSupervisor
        {
            get { return g_oNoEmpSupervisor; }
        }
        public static List<ValidateEmployee> NoEmpDept
        {
            get { return g_oNoEmpDept; }
        }
        public static List<ValidateEmployee> BadEmpLaborRate
        {
            get { return g_oBadEmpLaborRate; }
        }
        public static List<ValidateEmployee> NoEmpEmail
        {
            get { return g_oNoEmpEmail; }
        }
        public static List<ValidateEmployee> AllEmployees
        {
            get { return g_oAllEmployees; }
        }

        public string EmployeeId
        {
            get { return m_sEmployeeId; }
        }
        public string Status
        {
            get { return m_sStatus; }
        }
        public string FirstName
        {
            get { return m_sFirstName; }
        }
        public string LastName
        {
            get { return m_sLastName; }
        }
        public string Email
        {
            get { return m_sEmail; }
        }
        public string SupervisorEmployeeId
        {
            get { return m_sSupervisorEmployeeId; }
        }
        public int Shift
        {
            get { return m_iShift; }
        }
        public string ExpenseCode
        {
            get { return m_sExpenseCode; }
        }
        #endregion

        #region Data Members

        private static List<ValidateEmployee> g_oNoEmpSupervisor = new List<ValidateEmployee>();
        private static List<ValidateEmployee> g_oNoEmpDept = new List<ValidateEmployee>();
        private static List<ValidateEmployee> g_oBadEmpLaborRate = new List<ValidateEmployee>();
        private static List<ValidateEmployee> g_oNoEmpEmail = new List<ValidateEmployee>();
        private static List<ValidateEmployee> g_oAllEmployees = new List<ValidateEmployee>();

        private string m_sEmployeeId;
        private string m_sStatus;
        private string m_sFirstName;
        private string m_sLastName;
        private string m_sEmail;
        private string m_sSupervisorEmployeeId;
        private int m_iShift;
        private string m_sExpenseCode;

        private static string QUERY_VALIDATE_EMPLOYEE = "ValidateEmployee";

        #endregion
    }

    public class HSUser
    {
        #region Constructors

        public HSUser(DataRow oDataRow)
        {
            if ((oDataRow["UserComp_Company"] != DBNull.Value) && (string.IsNullOrEmpty((string)oDataRow["UserComp_Company"]) == false))
            {
                m_sCompany = (string)oDataRow["UserComp_Company"];
                m_sCompany = m_sCompany.Trim();
            }
            m_sUserId = (string)oDataRow["UserFile_DcdUserID"];
            if ((oDataRow["UserFile_Name"] != DBNull.Value) && (string.IsNullOrEmpty((string)oDataRow["UserFile_Name"]) == false))
            {
                m_sName = (string)oDataRow["UserFile_Name"];
                m_sName = m_sName.Trim();
            }
            if ((oDataRow["UserFile_EMailAddress"] != DBNull.Value) && (string.IsNullOrEmpty((string)oDataRow["UserFile_EMailAddress"]) == false))
            {
                m_sEmail = (string)oDataRow["UserFile_EMailAddress"];
                m_sEmail = m_sEmail.Trim();
            }
            if (oDataRow["UserFile_UserDisabled"] != DBNull.Value)
            {
                m_bDisabled = (bool)oDataRow["UserFile_UserDisabled"];
            }
            if (oDataRow["UserFile_SecurityMgr"] != DBNull.Value)
            {
                m_bSecurityManager = (bool)oDataRow["UserFile_SecurityMgr"];
            }
            if ((oDataRow["EmpBasic_EmpID"] != DBNull.Value) && (string.IsNullOrEmpty((string)oDataRow["EmpBasic_EmpID"]) == false))
            {
                m_sEmployeeId = (string)oDataRow["EmpBasic_EmpID"];
                m_sEmployeeId = m_sEmployeeId.Trim();
            }
            if ((oDataRow["EmpBasic_SupervisorID"] != DBNull.Value) && (string.IsNullOrEmpty((string)oDataRow["EmpBasic_SupervisorID"]) == false))
            {
                m_sSupervisorEmployeeId = (string)oDataRow["EmpBasic_SupervisorID"];
                m_sSupervisorEmployeeId = m_sSupervisorEmployeeId.Trim();
            }
            if ((oDataRow["UserFile_OSUserID"] != DBNull.Value) && (string.IsNullOrEmpty((string)oDataRow["UserFile_OSUserID"]) == false))
            {
                m_sOSUserId = (string)oDataRow["UserFile_OSUserID"];
                m_sOSUserId = m_sOSUserId.Trim();
            }
            if ((oDataRow["UserFile_DomainName"] != DBNull.Value) && (string.IsNullOrEmpty((string)oDataRow["UserFile_DomainName"]) == false))
            {
                m_sDomainName = (string)oDataRow["UserFile_DomainName"];
                m_sDomainName = m_sDomainName.Trim();
            }
            if (oDataRow["SysUserFile_HasADAccount_c"] != DBNull.Value)
            {
                m_bHasADAccount = (bool)oDataRow["SysUserFile_HasADAccount_c"];
            }

            if ((oDataRow["UserFile_GroupList"] != DBNull.Value) && (string.IsNullOrEmpty((string)oDataRow["UserFile_GroupList"]) == false))
            {
                string sTmp = (string)oDataRow["UserFile_GroupList"];
                m_oGroups.AddRange(sTmp.Split('~'));
            }
        }

        #endregion

        #region Methods

        public static bool Initialize(Session oSession)
        {
            bool bSuccess = true;

            //****Set a parameter Value***** 
            Ice.Proxy.BO.DynamicQueryImpl oDynamicQuery = WCFServiceSupport.CreateImpl<Ice.Proxy.BO.DynamicQueryImpl>(oSession, Ice.Proxy.BO.DynamicQueryImpl.UriPath);
            QueryExecutionDataSet oQueryExecutionDataSet = oDynamicQuery.GetQueryExecutionParametersByID(BAQConstants.QUERY_LIST_ALL_USERS);
            oQueryExecutionDataSet.ExecutionParameter.Clear();
            DataSet oDataSet = oDynamicQuery.ExecuteByID(BAQConstants.QUERY_LIST_ALL_USERS, oQueryExecutionDataSet);

            g_oIncorrectPasswordExpiration.Clear();
            g_oNoUserEmail.Clear();
            g_oNoUserCompany.Clear();
            g_oNoGroups.Clear();
            g_oNoEmployee.Clear();
            g_oNoDomainName.Clear();
            g_oNoOSUserId.Clear();

            // clear out anything in our list
            g_oHSUsers.Clear();
            foreach (DataRow oDataRow in oDataSet.Tables[0].Rows)
            {
                HSUser oTmpUser = new HSUser(oDataRow);
                g_oHSUsers.Add(oTmpUser);
           
                if ((oDataRow["Calculated_BadPasswordExp"] != DBNull.Value) && (string.IsNullOrEmpty((string)oDataRow["Calculated_BadPasswordExp"]) == false))
                {
                    g_oIncorrectPasswordExpiration.Add(oTmpUser);
                }
                if ((oDataRow["Calculated_NoUserEmail"] != DBNull.Value) && (string.IsNullOrEmpty((string)oDataRow["Calculated_NoUserEmail"]) == false))
                {
                    g_oNoUserEmail.Add(oTmpUser);
                }
                if ((oDataRow["Calculated_NoUserCompany"] != DBNull.Value) && (string.IsNullOrEmpty((string)oDataRow["Calculated_NoUserCompany"]) == false))
                {
                    g_oNoUserCompany.Add(oTmpUser);
                }
                if ((oDataRow["Calculated_NoUserGroupList"] != DBNull.Value) && (string.IsNullOrEmpty((string)oDataRow["Calculated_NoUserGroupList"]) == false))
                {
                    g_oNoGroups.Add(oTmpUser);
                }
                if ((oDataRow["Calculated_NoUserEmployee"] != DBNull.Value) && (string.IsNullOrEmpty((string)oDataRow["Calculated_NoUserEmployee"]) == false))
                {
                    g_oNoEmployee.Add(oTmpUser);
                }
                if ((oDataRow["Calculated_NoDomainName"] != DBNull.Value) && (string.IsNullOrEmpty((string)oDataRow["Calculated_NoDomainName"]) == false))
                {
                    g_oNoDomainName.Add(oTmpUser);
                }
                if ((oDataRow["Calculated_NoOSUserId"] != DBNull.Value) && (string.IsNullOrEmpty((string)oDataRow["Calculated_NoOSUserId"]) == false))
                {
                    g_oNoOSUserId.Add(oTmpUser);
                }
            }

            // there is a dependency on HSValidateEmployee -- need to be able to get the supervisor
            ValidateEmployee.Initialize(oSession);

            // now build up the list of active directory users
            // COMMENTING THIS OUT FOR NOW AS IT WAS NOT AVAILABLE ON 12/31/2024
            HSActiveDirectoryUser.Initialize();

            // build up the list of active payroll employees for this company
            HSPayrollUser.Initialize();
            return bSuccess;
        }

        static public void PerformValidation(string sTmpFileDirectory)
        {
            // get the file name
            DateTime dtToday = DateTime.Now;
            string sDate = dtToday.Month.ToString() + "-" + dtToday.Day.ToString() + "-" + dtToday.Year.ToString();
            string sDestinationFileName = sTmpFileDirectory + "\\-UserSetupReport-" + sDate + ".xlsx";

            // if the file is already there then get rid of it as we are reprocessing for some reason
            if (File.Exists(sDestinationFileName) == true)
            {
                File.Delete(sDestinationFileName);
            }

            HSUser oServiceAccount = HSUser.GetUserById(HSUser.SPINCRAFT_SERVICE_ACCOUNT_ID);

            List<string> oToAddresses = HSUser.GetEmailsForUsersInGroup(HSUser.REPORT_ON_EMPLOYEE_SETUP_ISSUES);
            oToAddresses.Add(HSUser.SPINCRAFT_ROOT_USER);

            System.Drawing.Color oHighlightYellow = System.Drawing.Color.FromArgb(255, 255, 0);

            SLStyle oHighlightHeaderStyle = new SLStyle();
            oHighlightHeaderStyle.SetFontBold(true);
            oHighlightHeaderStyle.SetFont(FontSchemeValues.Major, 12);
            oHighlightHeaderStyle.SetHorizontalAlignment(HorizontalAlignmentValues.Center);
            oHighlightHeaderStyle.SetPatternFill(PatternValues.Solid, oHighlightYellow, oHighlightYellow);
            oHighlightHeaderStyle.SetFontColor(SLThemeColorIndexValues.Dark2Color);

            SLDocument oSLDocument = new SLDocument();
            bool bFirstWorksheet = true;

            #region IT Issues
            //
            // IT Management issues
            //
            if (HSUser.IncorrectPasswordExpiration.Count != 0)
            {
                if (bFirstWorksheet == true)
                {
                    oSLDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "Bad Password Expiration");
                    bFirstWorksheet = false;
                }
                else
                {
                    oSLDocument.AddWorksheet("Bad Password Expiration");
                }

                int iNumberOfRows = 1;
                int iNumberOfColumns = 1;
                //set column header
                oSLDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "User Id");
                oSLDocument.SetColumnWidth(iNumberOfRows++, iNumberOfColumns++, 30);

                foreach (HSUser oUser in HSUser.IncorrectPasswordExpiration)
                {
                    oSLDocument.SetCellValue(iNumberOfRows++, 1, oUser.UserId);
                }
            }
            if (HSUser.NoUserEmail.Count != 0)
            {
                if (bFirstWorksheet == true)
                {
                    oSLDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "No Email");
                    bFirstWorksheet = false;
                }
                else
                {
                    oSLDocument.AddWorksheet("No Email");
                }

                int iNumberOfRows = 1;
                int iNumberOfColumns = 1;
                //set column header
                oSLDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "User Id");
                oSLDocument.SetColumnWidth(iNumberOfRows++, iNumberOfColumns++, 30);

                foreach (HSUser oUser in HSUser.NoUserEmail)
                {
                    oSLDocument.SetCellValue(iNumberOfRows++, 1, oUser.UserId);
                }
            }
            if (HSUser.NoUserCompany.Count != 0)
            {
                if (bFirstWorksheet == true)
                {
                    oSLDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "No Company");
                    bFirstWorksheet = false;
                }
                else
                {
                    oSLDocument.AddWorksheet("No Company");
                }

                int iNumberOfRows = 1;
                int iNumberOfColumns = 1;
                //set column header
                oSLDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "User Id");
                oSLDocument.SetColumnWidth(iNumberOfRows++, iNumberOfColumns++, 30);

                foreach (HSUser oUser in HSUser.NoUserCompany)
                {
                    oSLDocument.SetCellValue(iNumberOfRows++, 1, oUser.UserId);
                }
            }
            if (HSUser.NoGroups.Count != 0)
            {
                if (bFirstWorksheet == true)
                {
                    oSLDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "No Groups");
                    bFirstWorksheet = false;
                }
                else
                {
                    oSLDocument.AddWorksheet("No Groups");
                }

                int iNumberOfRows = 1;
                int iNumberOfColumns = 1;
                //set column header
                oSLDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "User Id");
                oSLDocument.SetColumnWidth(iNumberOfRows++, iNumberOfColumns++, 30);

                foreach (HSUser oUser in HSUser.NoGroups)
                {
                    oSLDocument.SetCellValue(iNumberOfRows++, 1, oUser.UserId);
                }
            }
            if (HSUser.NoEmployee.Count != 0)
            {
                if (bFirstWorksheet == true)
                {
                    oSLDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "No Employee Record");
                    bFirstWorksheet = false;
                }
                else
                {
                    oSLDocument.AddWorksheet("No Employee Record");
                }

                int iNumberOfRows = 1;
                int iNumberOfColumns = 1;
                //set column header
                oSLDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "User Id");
                oSLDocument.SetColumnWidth(iNumberOfRows++, iNumberOfColumns++, 30);

                foreach (HSUser oUser in HSUser.NoEmployee)
                {
                    oSLDocument.SetCellValue(iNumberOfRows++, 1, oUser.UserId);
                }
            }
            if (HSUser.NoDomainName.Count != 0)
            {
                if (bFirstWorksheet == true)
                {
                    oSLDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "No Domain");
                    bFirstWorksheet = false;
                }
                else
                {
                    oSLDocument.AddWorksheet("No Domain");
                }

                int iNumberOfRows = 1;
                int iNumberOfColumns = 1;
                //set column header
                oSLDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "User Id");
                oSLDocument.SetColumnWidth(iNumberOfRows++, iNumberOfColumns++, 30);

                foreach (HSUser oUser in HSUser.NoDomainName)
                {
                    oSLDocument.SetCellValue(iNumberOfRows++, 1, oUser.UserId);
                }
            }
            if (HSUser.NoOSUserId.Count != 0)
            {
                if (bFirstWorksheet == true)
                {
                    oSLDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "No OS User Id");
                    bFirstWorksheet = false;
                }
                else
                {
                    oSLDocument.AddWorksheet("No OS User Id");
                }

                int iNumberOfRows = 1;
                int iNumberOfColumns = 1;
                //set column header
                oSLDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "User Id");
                oSLDocument.SetColumnWidth(iNumberOfRows++, iNumberOfColumns++, 30);

                foreach (HSUser oUser in HSUser.NoOSUserId)
                {
                    oSLDocument.SetCellValue(iNumberOfRows++, 1, oUser.UserId);
                }
            }

            // send the email to the IT Managers
            if (bFirstWorksheet == false)
            {
                // we create a worksheet so we have some data in the spreadsheet so we email the spreadsheet
                oSLDocument.SaveAs(sDestinationFileName);
                List<string> oAttachments = new List<string>();
                oAttachments.Add(sDestinationFileName);
                HSEmailHelper.SendEmail(oToAddresses, "User Setup", "User Setup", oAttachments);
            }
            #endregion

            #region Account Audit Issues

            oToAddresses.Add(HSUser.SPINCRAFT_ROOT_USER);
            sDestinationFileName = sTmpFileDirectory + "\\-EmployeeAuditReport-" + sDate + ".xlsx";
            // if the file is already there then get rid of it as we are reprocessing for some reason
            if (File.Exists(sDestinationFileName) == true)
            {
                try
                {
                    File.Delete(sDestinationFileName);
                }
                catch (Exception)
                {
                    // we will ingore this if we cannot delete the file
                }
            }
            SLDocument oSLAuditDocument = new SLDocument();
            bFirstWorksheet = true;


            g_oActiveDirectoryAccountDisabledEpicorEnabled.Clear();
            // need to check if the AD account is disabled but the Epicor account is enabled
            foreach (HSActiveDirectoryUser oADUser in HSActiveDirectoryUser.AllADUsers)
            {
                if (oADUser.AccountDisabled == true)
                {
                    // AD has the user disabled so see if the user is in the Epicor list
                    // find the HSUser by account name and email address
                    HSUser oTmpUser = g_oHSUsers.FirstOrDefault(x => (string.Compare(x.OSUserId, oADUser.AccountName, true) == 0) && (string.Compare(x.Email, oADUser.Email, true) == 0));
                    if (oTmpUser != null)
                    {
                        if (oTmpUser.Disabled == false)
                        {
                            g_oActiveDirectoryAccountDisabledEpicorEnabled.Add(oTmpUser);
                        }
                    }
                }
            }

            /*
            List<string> oFindUsers = new List<string>();
            oFindUsers.Add("pthompson");
            foreach (string sUserName in oFindUsers)
            {
                HSActiveDirectoryUser oAD1 = HSActiveDirectoryUser.AllADUsers.FirstOrDefault(x => (string.Compare(x.AccountName, sUserName, true) == 0));
                if (oAD1 != null)
                {
                    Console.Write("FOUND " + sUserName);
                }
                else
                {
                    Console.WriteLine("Could not find " + sUserName);
                }
            }
            */

            g_oEpicorAccountDisabledActiveDirectoryEnabled.Clear();
            List<string> oEpicorDisabledAccountsToIgnore = new List<string>();
            oEpicorDisabledAccountsToIgnore.Add("scalvin");
            // need to check if the Epicor account is disabled but the AD account is enabled
            foreach (HSUser oTmpUser in g_oHSUsers)
            {
                if (oTmpUser.Disabled == true)
                {
                    if (oEpicorDisabledAccountsToIgnore.Contains(oTmpUser.UserId) == false)
                    {
                        // find the HSUser by account name and email address
                        HSActiveDirectoryUser oTmpADUser = HSActiveDirectoryUser.AllADUsers.FirstOrDefault(x => (string.Compare(x.AccountName, oTmpUser.OSUserId, true) == 0) && (string.Compare(x.Email, oTmpUser.Email, true) == 0));
                        if (oTmpADUser != null)
                        {
                            if (oTmpADUser.AccountDisabled == false)
                            {
                                g_oEpicorAccountDisabledActiveDirectoryEnabled.Add(oTmpUser);
                            }
                        }
                    }
                }
            }

            g_oEpicorAccountNoActiveDirectoryAccount.Clear();
            List<string> oActiveEpicorWithoutMatchingADAccounts = new List<string>();
            oActiveEpicorWithoutMatchingADAccounts.Add("blanglandadmin");
            oActiveEpicorWithoutMatchingADAccounts.Add("jmorrilladmin");
            oActiveEpicorWithoutMatchingADAccounts.Add("mnorthfieldadmin");
            oActiveEpicorWithoutMatchingADAccounts.Add("cottleadmin");
            oActiveEpicorWithoutMatchingADAccounts.Add("scalvinadmin");
            oActiveEpicorWithoutMatchingADAccounts.Add("Manager");
            oActiveEpicorWithoutMatchingADAccounts.Add("print");
            oActiveEpicorWithoutMatchingADAccounts.Add("SpincraftService");
            // check to see if all active Epicor users have an AD Account
            foreach (HSUser oTmpUser in g_oHSUsers)
            {
                if (oTmpUser.Disabled == false)
                {
                    // ensure this user id is not one for which we have an exception
                    if (oActiveEpicorWithoutMatchingADAccounts.Contains(oTmpUser.UserId) == false)
                    {
                        // check to make sure this user is supposed to have an AD account
                        if (oTmpUser.HasADAccount == true)
                        {
                            // find the HSUser by account name and email address
                            HSActiveDirectoryUser oTmpADUser = HSActiveDirectoryUser.AllADUsers.FirstOrDefault(x => (string.Compare(x.AccountName, oTmpUser.OSUserId, true) == 0) && (string.Compare(x.Email, oTmpUser.Email, true) == 0));
                            if (oTmpADUser == null)
                            {
                                // could not find an AD account for this Epicor user
                                g_oEpicorAccountNoActiveDirectoryAccount.Add(oTmpUser);
                            }
                        }
                    }
                }
            }

            if (HSUser.ActiveDirectoryAccountDisabledEpicorEnabled.Count != 0)
            {
                if (bFirstWorksheet == true)
                {
                    oSLAuditDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "AD Disabled");
                    bFirstWorksheet = false;
                }
                else
                {
                    oSLAuditDocument.AddWorksheet("AD Disabled");
                }

                int iNumberOfRows = 1;
                int iNumberOfColumns = 1;
                //set column header
                oSLAuditDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Epicor User Id");
                oSLAuditDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 30);
                oSLAuditDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Epicor Email");
                oSLAuditDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 30);
                oSLAuditDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Epicor Name");
                oSLAuditDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 30);
                oSLAuditDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "AD User Id");
                oSLAuditDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 30);
                oSLAuditDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "AD Email");
                oSLAuditDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 30);
                oSLAuditDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "AD Disabled But Epicor Is Enabled");
                oSLAuditDocument.SetCellStyle(iNumberOfRows++, iNumberOfColumns, oHighlightHeaderStyle);
                oSLAuditDocument.SetColumnWidth(iNumberOfColumns++, 90);

                foreach (HSUser oUser in HSUser.ActiveDirectoryAccountDisabledEpicorEnabled)
                {
                    oSLAuditDocument.SetCellValue(iNumberOfRows, 1, oUser.UserId);
                    oSLAuditDocument.SetCellValue(iNumberOfRows, 2, oUser.Email);
                    oSLAuditDocument.SetCellValue(iNumberOfRows, 3, oUser.Name);
                    HSActiveDirectoryUser oTmpADUser = HSActiveDirectoryUser.AllADUsers.FirstOrDefault(x => string.Compare(x.AccountName, oUser.OSUserId, true) == 0);
                    if (oTmpADUser != null)
                    {
                        oSLAuditDocument.SetCellValue(iNumberOfRows, 4, oTmpADUser.AccountName);
                        oSLAuditDocument.SetCellValue(iNumberOfRows, 5, oTmpADUser.Email);
                    }
                    iNumberOfRows++;
                }
            }
            if (HSUser.EpicorAccountDisabledActiveDirectoryEnabled.Count != 0)
            {
                if (bFirstWorksheet == true)
                {
                    oSLAuditDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "Epicor Disabled");
                    bFirstWorksheet = false;
                }
                else
                {
                    oSLAuditDocument.AddWorksheet("Epicor Disabled");
                }

                int iNumberOfRows = 1;
                int iNumberOfColumns = 1;
                //set column header
                oSLAuditDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Epicor User Id");
                oSLAuditDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 30);
                oSLAuditDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Epicor Email");
                oSLAuditDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 30);
                oSLAuditDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Epicor Name");
                oSLAuditDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 30);
                oSLAuditDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "AD User Id");
                oSLAuditDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 30);
                oSLAuditDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "AD Email");
                oSLAuditDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 30);
                oSLAuditDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Epicor Disabled But AD Is Enabled");
                oSLAuditDocument.SetCellStyle(iNumberOfRows++, iNumberOfColumns, oHighlightHeaderStyle);
                oSLAuditDocument.SetColumnWidth(iNumberOfColumns++, 90);

                // we dont actually have a user becuase the Epicor account is disabled or missing so we just show the AD name and email account
                foreach (HSUser oUser in HSUser.EpicorAccountDisabledActiveDirectoryEnabled)
                {
                    oSLAuditDocument.SetCellValue(iNumberOfRows, 1, oUser.UserId);
                    oSLAuditDocument.SetCellValue(iNumberOfRows, 2, oUser.Email);
                    oSLAuditDocument.SetCellValue(iNumberOfRows, 3, oUser.Name);
                    HSActiveDirectoryUser oTmpADUser = HSActiveDirectoryUser.AllADUsers.FirstOrDefault(x => string.Compare(x.AccountName, oUser.OSUserId, true) == 0);
                    if (oTmpADUser != null)
                    {
                        oSLAuditDocument.SetCellValue(iNumberOfRows, 4, oTmpADUser.AccountName);
                        oSLAuditDocument.SetCellValue(iNumberOfRows, 5, oTmpADUser.Email);
                    }
                    iNumberOfRows++;
                }
            }
            if (HSUser.g_oEpicorAccountNoActiveDirectoryAccount.Count != 0)
            {
                if (bFirstWorksheet == true)
                {
                    oSLAuditDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "No AD Account");
                    bFirstWorksheet = false;
                }
                else
                {
                    oSLAuditDocument.AddWorksheet("No AD Account");
                }

                int iNumberOfRows = 1;
                int iNumberOfColumns = 1;
                //set column header
                oSLAuditDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Epicor User Id");
                oSLAuditDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 30);
                oSLAuditDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Epicor Email");
                oSLAuditDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 30);
                oSLAuditDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Epicor Name");
                oSLAuditDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 30);

                oSLAuditDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Epicor User Has No AD Account");
                oSLAuditDocument.SetCellStyle(iNumberOfRows++, iNumberOfColumns, oHighlightHeaderStyle);
                oSLAuditDocument.SetColumnWidth(iNumberOfColumns++, 90);

                // we dont actually have a user becuase the Epicor account is disabled or missing so we just show the AD name and email account
                foreach (HSUser oUser in HSUser.EpicorAccountNoActiveDirectoryAccount)
                {
                    oSLAuditDocument.SetCellValue(iNumberOfRows, 1, oUser.UserId);
                    oSLAuditDocument.SetCellValue(iNumberOfRows, 2, oUser.Email);
                    oSLAuditDocument.SetCellValue(iNumberOfRows, 3, oUser.Name);

                    iNumberOfRows++;
                }
            }

            // send the email to IT
            if (bFirstWorksheet == false)
            {
                // we create a worksheet so we have some data in the spreadsheet so we email the spreadsheet
                oSLAuditDocument.SaveAs(sDestinationFileName);
                List<string> oAttachments = new List<string>();
                oAttachments.Add(sDestinationFileName);
                HSEmailHelper.SendEmail(oToAddresses, "Employee Audit", "Employee Audit", oAttachments);
            }
            #endregion
        }

        public static List<HSUser> GetUsersInGroup(string sGroup)
        {
            List<HSUser> oUsersInGroup = new List<HSUser>();
            oUsersInGroup = g_oHSUsers.Where(oItem => oItem.m_oGroups.Contains(sGroup) && oItem.m_bDisabled == false).ToList();
            return oUsersInGroup;
        }

        public static List<string> GetEmailsForUsersInGroup(string sGroup)
        {
            List<string> oEmailAddresses = new List<string>();
            List<HSUser> oUsersInGroup = GetUsersInGroup(sGroup);
            oEmailAddresses = oUsersInGroup.Where(oItem => string.IsNullOrEmpty(oItem.Email) == false).Select(oItem => oItem.Email).ToList();
            return oEmailAddresses;
        }

        public static List<string> GetEmailsForUsersInGroup(List<string> oCurrentEmails, string sGroup)
        {
            List<HSUser> oUsersInGroup = GetUsersInGroup(sGroup);
            List<string> oTmpEmails = oUsersInGroup.Select(oItem => oItem.Email).ToList();
            foreach (string sTmpEmail in oTmpEmails)
            {
                if (string.IsNullOrEmpty(sTmpEmail) == false)
                {
                    if (oCurrentEmails.Contains(sTmpEmail) == false)
                    {
                        oCurrentEmails.Add(sTmpEmail);
                    }
                }
            }
            return oCurrentEmails;
        }

        public static string GetEmailForUser(string sUserId)
        {
            string sEmailAddress = string.Empty;
            HSUser oUser = GetUserById(sUserId);
            if (oUser != null)
            {
                sEmailAddress = oUser.Email;
            }
            return sEmailAddress;
        }

        public static List<string> GetEmailForUser(List<string> oCurrentEmails, string sUserId)
        {
            HSUser oUser = GetUserById(sUserId);
            if ((oUser != null) && (string.IsNullOrEmpty(oUser.Email) == false))
            {
                if (oCurrentEmails.Contains(oUser.Email) == false)
                {
                    oCurrentEmails.Add(oUser.Email);
                }
            }
            return oCurrentEmails;
        }

        public static HSUser GetUserById(string sUserId)
        {
            // sometimes we pass in the employee Id or supervisor Id and these are not exact matches
            // due to the 8 character limitation so we see if it mostly matches
            HSUser oUser = null;
            if (string.IsNullOrEmpty(sUserId) == false)
            {
                oUser = g_oHSUsers.FirstOrDefault(oItem => oItem.m_sUserId.IndexOf(sUserId) != -1);
            }
            return oUser;
        }

        public static HSUser GetSupervisor(string sUserId)
        {
            HSUser oSupervisor = null;
            HSUser oCurrentUser = GetUserById(sUserId);
            if (oCurrentUser != null)
            {
                oSupervisor = g_oHSUsers.FirstOrDefault(oItem => string.Compare(oItem.EmployeeId, oCurrentUser.SupervisorId, true) == 0);
            }
            return oSupervisor;
        }

        #endregion

        #region Properties

        public static List<HSUser> HSUsers
        {
            get { return g_oHSUsers; }
        }
        public static List<HSUser> IncorrectPasswordExpiration
        {
            get { return g_oIncorrectPasswordExpiration; }
        }
        public static List<HSUser> NoUserEmail
        {
            get { return g_oNoUserEmail; }
        }
        public static List<HSUser> NoUserCompany
        {
            get { return g_oNoUserCompany; }
        }
        public static List<HSUser> NoGroups
        {
            get { return g_oNoGroups; }
        }
        public static List<HSUser> NoEmployee
        {
            get { return g_oNoEmployee; }
        }
        public static List<HSUser> NoDomainName
        {
            get {  return g_oNoDomainName; }
        }
        public static List<HSUser> NoOSUserId
        {
            get { return g_oNoOSUserId; }
        }

        public string OSUserId
        {
            get { return m_sOSUserId; }
        }

        public string DomainName
        {
            get { return m_sDomainName; }
        }
        public bool HasADAccount
        {
            get { return m_bHasADAccount; }
        }
        public string Company
        {
            get { return m_sCompany; }
        }
        public string UserId
        {
            get { return m_sUserId; }
        }
        public string Name
        {
            get { return m_sName; }
        }
        public string Email
        {
            get { return m_sEmail; }
        }
        public bool Disabled
        {
            get { return m_bDisabled; }
        }
        public bool SecurityManager
        {
            get { return m_bSecurityManager; }
        }
        public string EmployeeId
        {
            get { return m_sEmployeeId; }
        }
        public string SupervisorId
        {
            get { return m_sSupervisorEmployeeId; }
        }

        public static List<HSUser> ActiveDirectoryAccountDisabledEpicorEnabled
        {
            get { return g_oActiveDirectoryAccountDisabledEpicorEnabled; }
        }

        public static List<HSUser> EpicorAccountDisabledActiveDirectoryEnabled
        {
            get { return g_oEpicorAccountDisabledActiveDirectoryEnabled; }
        }

        public static List<HSUser> EpicorAccountNoActiveDirectoryAccount
        {
            get { return g_oEpicorAccountNoActiveDirectoryAccount; }
        }
        #endregion

        #region Data Members

        private static List<HSUser> g_oHSUsers = new List<HSUser>();

        private static List<HSUser> g_oIncorrectPasswordExpiration = new List<HSUser>();
        private static List<HSUser> g_oNoUserEmail = new List<HSUser>();
        private static List<HSUser> g_oNoUserCompany = new List<HSUser>();
        private static List<HSUser> g_oNoGroups = new List<HSUser>();
        private static List<HSUser> g_oNoEmployee = new List<HSUser>();
        private static List<HSUser> g_oNoOSUserId = new List<HSUser>();
        private static List<HSUser> g_oNoDomainName = new List<HSUser>();

        private static List<HSUser> g_oActiveDirectoryAccountDisabledEpicorEnabled = new List<HSUser>();
        private static List<HSUser> g_oEpicorAccountNoActiveDirectoryAccount = new List<HSUser>();
        private static List<HSUser> g_oEpicorAccountDisabledActiveDirectoryEnabled = new List<HSUser>();

        private string m_sCompany;
        private string m_sUserId;
        private string m_sEmployeeId;
        private string m_sName;
        private string m_sEmail;
        private string m_sShortEmail;
        private bool m_bDisabled;
        private List<string> m_oGroups = new List<string>();
        private bool m_bSecurityManager;
        private string m_sSupervisorEmployeeId;
        private int m_iADPId;
        private string m_sResourceGroup;
        private string m_sOSUserId;
        private string m_sDomainName;
        private bool m_bHasADAccount;
        //private string m_sSupervisorId;


        // User Roles In Epicor
        public static string ACCOUNTING_AP = "ACCT AP";
        public static string ACCOUNTING_AR = "ACCT AR";
        public static string ACCOUNTING_MANAGER = "ACCT-MGR";
        public static string ACCOUNTING_REPORTS = "ACCT RP";
        public static string AUDIT = "AUDIT";
        public static string BAQ_USERS = "BAQ_USER";
        public static string BOM_MANAGER = "BOM-MGR";
        public static string CHIEF_FINANCIAL_OFFICE = "CFO";
        public static string RECEIVING = "RECV";
        public static string RECEIVING_MANAGER = "RCV-MGR";
        // begin user roles specifc to reporting groups
        public static string REPORT_ON_EMPLOYEE_SETUP_ISSUES = "EMP_SETUP"; // reports on problems with employee setup
        public static string REPORT_ON_EMPLOYEES_CLOCKED_IN = "EMP_CLOCK"; // reports on problems with employee setup
        public static string REPORT_ON_POSITIVE_PAY = "POS_PAY"; // reports on positive pay

        public static string REPORT_ON_JOB_ESTIMATES = "RPT_JOB_EST"; // reports on job estimates
        public static string REPORT_ON_PROCUREMENT_ISSUES = "RPT_PO"; // reports on any procurement issues
        public static string REPORT_ON_SO_ISSUES = "RPT_SO"; // reports on any sales order issues
        public static string REPORT_ON_PURCHASED_PART_ISSUES = "RPT_PART"; // reports on problems with purchased parts
        public static string REPORT_ON_MANUFACTURED_PART_ISSUES = "RPT_BOM"; // reports on problems with manufactured parts
        public static string REPORT_ON_JOBS = "RPT_JOBS"; // reports on job validation
        public static string REPORT_ON_SHIPPING = "RPT_SHIP"; // reports on shipping validation
        public static string REPORT_ON_RECEIVING = "RPT_RCPT"; // reports on receipt validation
        public static string REPORT_ON_QUOTES = "RPT_QUOTE"; // reports on quote validation
        public static string REPORTS_PARTS_IN_INSPECTION = "PART_INSP"; // reports on parts received still in inspection
        public static string REPORTS_PO_RECEIVE_ISSUES = "PO_RCV_ISSUE"; // reports on parts received issues

        // enduser roles specifc to reporting groups
        public static string REPORTS = "RPTS";
        public static string TRACKERS = "TRCK";

        public static string SPINCRAFT_SERVICE_ACCOUNT_ID => HSCredentialStore.Get(HSCredentialStore.Keys.SpincraftServiceAccountId);
        public static string SPINCRAFT_SERVICE_PASSWORD => HSCredentialStore.Get(HSCredentialStore.Keys.SpincraftServicePassword);

        // this person gets emailed for all emails sent from Epicor
        public static string SPINCRAFT_ROOT_USER = "blangland@horizonscientific.com";

        #endregion
    }

    public class HSActiveDirectoryUser
    {
        #region Constructors
        public HSActiveDirectoryUser()
        {

        }

        public HSActiveDirectoryUser(string sOrganizationalUnit, string sAccountName, string sDomainName, string sDisplayName, string sEmail, bool bNormalAccount, bool bAccountDisabled, bool bAccountLockedOut)
        {
            OrganizationalUnit = sOrganizationalUnit.Trim();
            AccountName = sAccountName.Trim();
            DomainName = sDomainName.Trim();
            DisplayName = sDisplayName.Trim();
            Email = sEmail.Trim();
            NormalAccount = bNormalAccount;
            AccountDisabled = bAccountDisabled;
            Lockout = bAccountLockedOut;
        }
        #endregion

        #region Methods
        public static void Initialize()
        {
            g_oAllActiveDirectoryUsers.Clear();
            g_oAllActiveDirectoryUsersForThisCompany.Clear();

            // Set the context to your domain (null = current domain)
            using (PrincipalContext context = new PrincipalContext(ContextType.Domain))
            {
                // Create a UserPrincipal filter
                using (UserPrincipal userFilter = new UserPrincipal(context))
                {
                    // Create a PrincipalSearcher using the filter
                    using (PrincipalSearcher searcher = new PrincipalSearcher(userFilter))
                    {
                        foreach (var result in searcher.FindAll())
                        {
                            string sOrganizationalUnit = string.Empty;
                            bool bAccountDisabled = false;
                            bool bAccountLockedOut = false;
                            string sAccountName = string.Empty;
                            string sDomainName = string.Empty;
                            string sDisplayName = string.Empty;
                            string sEmail = string.Empty;
                            bool bNormalAccount = false;
                            UserPrincipal oUserPrincipal = result as UserPrincipal;
                            if (oUserPrincipal != null)
                            {
                                sAccountName = oUserPrincipal.SamAccountName;
                                if (string.IsNullOrEmpty(oUserPrincipal.DisplayName) == false)
                                {
                                    sDisplayName = oUserPrincipal.DisplayName;
                                }
                                if (string.IsNullOrEmpty(oUserPrincipal.Context.Name) == false)
                                {
                                    sDomainName = oUserPrincipal.Context.Name;
                                }
                                if (string.IsNullOrEmpty(oUserPrincipal.EmailAddress) == false)
                                {
                                    sEmail = oUserPrincipal.EmailAddress;
                                }
                                if (oUserPrincipal.Enabled.HasValue)
                                {
                                    if (oUserPrincipal.Enabled.Value == true)
                                    {
                                        bAccountDisabled = false;
                                    }
                                    else
                                    {
                                        bAccountDisabled = true;
                                    }
                                }
                                // Check if the account is locked out
                                bAccountLockedOut = oUserPrincipal.IsAccountLockedOut();

                                // now get data from directory entry object
                                DirectoryEntry oDirectoryEntry = (DirectoryEntry)result.GetUnderlyingObject();
                                if (oDirectoryEntry != null)
                                {
                                    int iUserAccountControl = oDirectoryEntry.Properties["userAccountControl"].Value != null ? Convert.ToInt32(oDirectoryEntry.Properties["userAccountControl"].Value) : 0;
                                    // Bit 0x0200 = NORMAL_ACCOUNT
                                    bNormalAccount = (iUserAccountControl & NORMAL_ACCOUNT) != 0;

                                    // Extract OU from Distinguished Name
                                    string sDistinguishedName = oDirectoryEntry.Properties["distinguishedName"].Value as string;
                                    sOrganizationalUnit = GetOrganizationalUnit(sDistinguishedName);
                                }
                            }

                            HSActiveDirectoryUser oADUser = new HSActiveDirectoryUser(sOrganizationalUnit, sAccountName, sDomainName, sDisplayName, sEmail, bNormalAccount, bAccountDisabled, bAccountLockedOut);
                            g_oAllActiveDirectoryUsers.Add(oADUser);
                        }
                    }
                }
            }

            List<string> oOrgUnitsToInclude = new List<string>();

            oOrgUnitsToInclude.Add("Users/Los Angeles/ETG");
            oOrgUnitsToInclude.Add("Users/New Berlin/ETG");
            oOrgUnitsToInclude.Add("Users/North Billerica/ETG");
            oOrgUnitsToInclude.Add("Shared Accounts/BeyondTrust");

            //oOrgUnitsToInclude.Add("_Scientific");
            //oOrgUnitsToInclude.Add("_Corporate");
            //oOrgUnitsToInclude.Add("_Federal");
            //oOrgUnitsToInclude.Add("_Hydraulics");
            //oOrgUnitsToInclude.Add("_Innovent");
            //oOrgUnitsToInclude.Add("_ETG");
            //oOrgUnitsToInclude.Add("Disabled Users");
            g_oAllActiveDirectoryUsersForThisCompany = g_oAllActiveDirectoryUsers.Where(oItem => (!string.IsNullOrEmpty(oItem.OrganizationalUnit)) && (oItem.OrganizationalUnit.ContainsAny(oOrgUnitsToInclude, StringComparison.CurrentCultureIgnoreCase) == true)).ToList();

            
            // Get distinct Organizational Units (ignoring null or empty entries)
            //var uniqueOUs = AllADUsers
            //    .Where(u => !string.IsNullOrWhiteSpace(u.OrganizationalUnit))
            //    .Select(u => u.OrganizationalUnit)
            //    .Distinct()
            //    .OrderBy(ou => ou)
            //    .ToList();

            // Output them
            //Console.WriteLine("Unique Organizational Units:");
            //foreach (var ou in uniqueOUs)
            //{
            //    Console.WriteLine(ou);
            //}
            
        }

        static string GetOrganizationalUnit(string distinguishedName)
        {
            if (string.IsNullOrEmpty(distinguishedName))
                return "";

            string[] parts = distinguishedName.Split(',');
            string ouPath = "";
            foreach (string part in parts)
            {
                if (part.StartsWith("OU=", StringComparison.OrdinalIgnoreCase))
                {
                    ouPath += part.Substring(3) + "/";
                }
            }
            return ouPath.TrimEnd('/');
        }

        #endregion

        #region Properties
        public string OrganizationalUnit 
        { 
            get { return m_sOrganizationalUnit; }
            set { m_sOrganizationalUnit = value.Trim();  }
        }
        public string AccountName
        {
            get { return m_sAccountName; }
            set { m_sAccountName = value.Trim(); }
        }

        public string DomainName
        {
            get { return m_sDomainName; }
            set { m_sDomainName = value.Trim(); }
        }

        public string DisplayName
        {
            get { return m_sDisplayName; }
            set { m_sDisplayName = value.Trim(); }

        }

        public string Email
        {
            get
            {
                return m_sEmail;
            }
            set
            {
                m_sEmail = value;
                m_sShortEmail = string.Empty;
                if (string.IsNullOrEmpty(m_sEmail) == false)
                {
                    string[] sTmp = m_sEmail.Split('@');
                    if (sTmp.Length > 0)
                    {
                        m_sShortEmail = sTmp[0];
                    }
                }
            }
        }
        public string ShortEmail
        {
            get { return m_sShortEmail; }
        }
        public bool NormalAccount
        {
            get { return m_bNormalAccount; }
            set { m_bNormalAccount = value; }
        }
        public bool AccountDisabled
        {
            get { return m_bAccountDisabled; }
            set { m_bAccountDisabled = value; }
        }

        public bool Lockout
        {
            get { return m_bLockout; }
            set { m_bLockout = value; }
        }

        public static List<HSActiveDirectoryUser> AllADUsers
        {
            get { return g_oAllActiveDirectoryUsers; }
        }

        public static List<HSActiveDirectoryUser> AllADUsersForThisCompany
        {
            get { return g_oAllActiveDirectoryUsersForThisCompany; }
        }
        #endregion

        #region Data Members

        public static List<HSActiveDirectoryUser> g_oAllActiveDirectoryUsers = new List<HSActiveDirectoryUser>();
        public static List<HSActiveDirectoryUser> g_oAllActiveDirectoryUsersForThisCompany = new List<HSActiveDirectoryUser>();

        public static long SCRIPT = 0x0001;
        public static long ACCOUNTDISABLE = 0x0002;
        public static long HOMEDIR_REQUIRED = 0x0008;
        public static long LOCKOUT = 0x0010;
        public static long PASSWD_NOTREQD = 0x0020;
        public static long PASSWD_CANT_CHANGE = 0x0040;
        public static long ENCRYPTED_TEXT_PWD_ALLOWED = 0x0080;
        public static long TEMP_DUPLICATE_ACCOUNT = 0x0100;
        public static long NORMAL_ACCOUNT = 0x0200;
        public static long INTERDOMAIN_TRUST_ACCOUNT = 0x0800;
        public static long WORKSTATION_TRUST_ACCOUNT = 0x1000;
        public static long SERVER_TRUST_ACCOUNT = 0x2000;
        public static long DONT_EXPIRE_PASSWORD = 0x10000;
        public static long MNS_LOGON_ACCOUNT = 0x20000;
        public static long SMARTCARD_REQUIRED = 0x40000;
        public static long TRUSTED_FOR_DELEGATION = 0x80000;
        public static long NOT_DELEGATED = 0x100000;
        public static long USE_DES_KEY_ONLY = 0x200000;
        public static long DONT_REQ_PREAUTH = 0x400000;
        public static long PASSWORD_EXPIRED = 0x800000;
        public static long TRUSTED_TO_AUTH_FOR_DELEGATION = 0x1000000;
        public static long PARTIAL_SECRETS_ACCOUNT = 0x04000000;

        private string m_sOrganizationalUnit;
        private string m_sAccountName;
        private string m_sDomainName;
        private string m_sDisplayName;
        private string m_sEmail;
        private string m_sShortEmail;
        private bool m_bNormalAccount;
        private bool m_bAccountDisabled;
        private bool m_bLockout;
        #endregion
    }

    public class HSPayrollUser
    {
        #region Constructors
        public HSPayrollUser()
        {

        }

        public HSPayrollUser(string sOrganizationalUnit, string sFirstName, string sLastName, bool bActive)
        {
            OrganizationalUnit = sOrganizationalUnit.Trim();
            FirstName = sFirstName.Trim();
            LastName = sLastName.Trim();
            Active = bActive;
        }
        #endregion

        #region Methods
        public static void Initialize()
        {
            g_oAllPayrollUsers.Clear();
            // get the active payroll employess from Day Force and populate the list
        }


        #endregion

        #region Properties
        public string OrganizationalUnit
        {
            get { return m_sOrganizationalUnit; }
            set { m_sOrganizationalUnit = value.Trim(); }
        }
        public string FirstName
        {
            get { return m_sFirstName; }
            set { m_sFirstName = value.Trim(); }
        }
        public string LastName
        {
            get { return m_sLastName; }
            set { m_sLastName = value.Trim(); }
        }

        public bool Active
        {
            get { return m_bActive; }
            set { m_bActive = value; }
        }


        #endregion

            #region Data Members

        private static List<HSPayrollUser> g_oAllPayrollUsers = new List<HSPayrollUser>();

        private string m_sOrganizationalUnit;
        private string m_sFirstName;
        private string m_sLastName;
        private bool m_bActive;
        #endregion
    }
}
