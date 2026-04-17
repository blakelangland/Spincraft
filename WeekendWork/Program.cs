using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using System.Globalization;

using Ice.Core;
using System.IO;
using HorizonScientific;

// MUST TARGET .NET 4.8
// We should reference the adapter and contract classes to both read and write to Epicor
// We start by connecting to Epicor through a session object
// The Epicor business objects (CreateImpl) are used to invoke the methods

namespace WeekendWork
{
    internal class Program
    {
        static void Main(string[] args)
        {
            try
            {
                bool bHSValidateEmployeeValidation = false;
                bool bValidateParts = false;
                bool bValidateBOMs = false;
                bool bValidateOperations = false;
                bool bValidateReourceGroups = false;
                bool bValidateResources = false;

                // we will see if this application has any arguments being passed in
                HSUser oRequestingUser = null;
                string sRequestingUserId = "";
                bool bGetUserId = false;

                bool bGetCompany = false;
                string sCompany = CompanyConfiguration.SPINCRAFT_MA_COMPANY_ID;

                string sProductionCalendarId = ProductionCalendarCollection.SPINCRAFT_MA_PRODUCTION_CALENDAR;

                HSValidateParts oValidateParts = null;
                BOMSupport oBOMSupport = null;

                bool bExecutedFromTaskScheduler = false;

                if (args.Count() > 0)
                {
                    foreach (string sArg in args)
                    {
                        if (bGetUserId == true)
                        {
                            bGetUserId = false;
                            sRequestingUserId = sArg;
                            continue;
                        }
                        if (bGetCompany == true)
                        {
                            bGetCompany = false;
                            sCompany = sArg;
                            continue;
                        }

                        else if (string.Compare(sArg, "USER_ID", true) == 0)
                        {
                            // the next argument will be the user id
                            bGetUserId = true;
                        }
                        else if (string.Compare(sArg, "COMPANY", true) == 0)
                        {
                            // the next argument will be the company id
                            bGetCompany = true;
                        }
                        else if (string.Compare(sArg, "TASK_SCHEDULER", true) == 0)
                        {
                            // this process was kicked off from the task scheduler
                            bExecutedFromTaskScheduler = true;
                        }
                    }
                }


                if (bExecutedFromTaskScheduler == true)
                {
                    // tasks we perform when kicked off from the task scheduler
                    bHSValidateEmployeeValidation = true;
                    bValidateParts = true;
                    bValidateBOMs = true;
                    bValidateOperations = true;
                    bValidateReourceGroups = true;
                    bValidateResources = true;
                }

                // get rid of any existing files in this temp directory to clean things up
                DirectoryInfo oDirectoryInfo = new DirectoryInfo(TEMP_FILE_DIRECTORY);
                foreach (FileInfo oTmpFile in oDirectoryInfo.EnumerateFiles())
                {
                    try
                    {
                        File.Delete(oTmpFile.FullName);
                    }
                    catch (Exception)
                    {
                        // we ignore any exeptions with file deletion as files
                        // could have been dropped off from other remote servers to process
                    }
                }

                if (string.Compare(sCompany, CompanyConfiguration.SPINCRAFT_MA_COMPANY_ID, true) == 0)
                {
                    g_sConfigLocation = g_sSpincraftMAConfiguration;
                    sCompany = CompanyConfiguration.SPINCRAFT_MA_COMPANY_ID;
                    sProductionCalendarId = ProductionCalendarCollection.SPINCRAFT_MA_PRODUCTION_CALENDAR;
                }
                else
                {
                    // default is WI if not stated
                    g_sConfigLocation = g_sSpincraftWIConfiguration;
                    sCompany = CompanyConfiguration.SPINCRAFT_WI_COMPANY_ID;
                    sProductionCalendarId = ProductionCalendarCollection.SPINCRAFT_WI_PRODUCTION_CALENDAR;
                }

                g_oSession = new Session(HSUser.SPINCRAFT_SERVICE_ACCOUNT_ID, HSUser.SPINCRAFT_SERVICE_PASSWORD, Session.LicenseType.Default, g_sConfigLocation);
                if (g_oSession != null)
                {
                    #region Metadata Validation
                    try
                    {
                        if (HSUser.Initialize(AppSession) == false)
                        {
                            Console.WriteLine("Failed to load all users!");
                            return;
                        }
                        if (string.IsNullOrEmpty(sRequestingUserId) == false)
                        {
                            oRequestingUser = HSUser.GetUserById(sRequestingUserId);
                        }
                        // set up the production calendar
                        DateTime dtFirstDayOfLastYear = new DateTime(DateTime.Now.Year - 1, 1, 1);
                        DateTime dtLastDayOfNextYear = new DateTime(DateTime.Now.Year + 1, 12, 31);
                        if (ProductionCalendarCollection.Initialize(AppSession, dtFirstDayOfLastYear, dtLastDayOfNextYear) == false)
                        {
                            Console.WriteLine("Fail to set up production calendar");
                            return;
                        }

                        HSUser.PerformValidation(TEMP_FILE_DIRECTORY);
                    }
                    catch (Exception ex)
                    {
                        ReportException(ex, "User");
                    }
                    #endregion

                    // need all parts in the database for several validations so we will load it only once
                    #region Part Validation
                    if (bValidateParts == true)
                    {
                        try
                        {
                            oValidateParts = new HSValidateParts();
                            if (oValidateParts.Initialize(AppSession) == false)
                            {
                                Console.WriteLine("Failed to load the validate parts!");
                            }
                            oValidateParts.PerformPurchasePartValidation(sCompany, TEMP_FILE_DIRECTORY);
                        }
                        catch (Exception ex)
                        {
                            ReportException(ex, "Validate Parts");
                        }
                    }
                    #endregion

                    // we create the BOM Support object as we will use it in several places
                    oBOMSupport = new BOMSupport(sCompany);

                    #region Production Issues

                    #endregion

                    #region Op And Resource Validation
                    if (bValidateOperations == true)
                    {
                        try
                        {
                            if (oBOMSupport.InitializeOperations(AppSession) == false)
                            {
                                Console.WriteLine("Failed to load the BOM Support object!");
                            }
                            oBOMSupport.PerformOperationValidation(sCompany, TEMP_FILE_DIRECTORY);
                        }
                        catch (Exception ex)
                        {
                            ReportException(ex, "Validate Operations");
                        }
                    }

                    if (bValidateReourceGroups == true)
                    {
                        try
                        {
                            if (oBOMSupport.InitializeResourceGroups(AppSession) == false)
                            {
                                Console.WriteLine("Failed to load the BOM Support object!");
                            }
                            oBOMSupport.PerformResourceGroupValidation(sCompany, TEMP_FILE_DIRECTORY);
                        }
                        catch (Exception ex)
                        {
                            ReportException(ex, "Validate Resource Groups");
                        }
                    }

                    if (bValidateResources == true)
                    {
                        try
                        {
                            if (oBOMSupport.InitializeResources(AppSession) == false)
                            {
                                Console.WriteLine("Failed to load the BOM Support object!");
                            }
                            oBOMSupport.PerformResourceValidation(sCompany, TEMP_FILE_DIRECTORY);
                        }
                        catch (Exception ex)
                        {
                            ReportException(ex, "Validate Resources");
                        }
                    }
                    #endregion

                    #region BOM Validation
                    if (bValidateBOMs == true)
                    {
                        try
                        {
                            if (oBOMSupport.Initialize(AppSession, oValidateParts) == false)
                            {
                                Console.WriteLine("Failed to load the BOM Support object!");
                            }
                            oBOMSupport.PerformValidation(sCompany, TEMP_FILE_DIRECTORY);
                        }
                        catch (Exception ex)
                        {
                            ReportException(ex, "Validate BOMs");
                        }
                    }
                    #endregion
                }
            }
            catch (Exception error)
            {
                // send an email out with error message
                HSUser oServiceAccount = HSUser.GetUserById(HSUser.SPINCRAFT_SERVICE_ACCOUNT_ID);
                List<string> oToAddresses = new List<string>();
                oToAddresses.Add(HSUser.SPINCRAFT_ROOT_USER);

                StringBuilder oStringBuilder = new StringBuilder();
                oStringBuilder.Clear();
                oStringBuilder.Append("An error occured while executing Weekend Work:\n");

                oStringBuilder.Append("Exception Message:\n");
                oStringBuilder.Append(error.Message);
                oStringBuilder.Append("\n\n");
                oStringBuilder.Append("Exception Inner Message:\n");
                oStringBuilder.Append(error.InnerException);
                oStringBuilder.Append("\n\n");
                HSEmailHelper.SendEmail(oToAddresses, "Error Executing Weekend Work", oStringBuilder.ToString());
            }
            finally
            {
                HSUser oServiceAccount = HSUser.GetUserById(HSUser.SPINCRAFT_SERVICE_ACCOUNT_ID);
                List<string> oToAddresses = new List<string>();
                oToAddresses.Add(HSUser.SPINCRAFT_ROOT_USER);

                StringBuilder oStringBuilder = new StringBuilder();
                oStringBuilder.Clear();
                oStringBuilder.Append("The Process WeekendWork was executed for Spincraft.");

                HSEmailHelper.SendEmail(oToAddresses, "The Weekend Work Process Ran", oStringBuilder.ToString());

                if (g_oSession != null)
                {
                    g_oSession.Dispose();
                    g_oSession = null;
                }
            }

        }

        public static void ReportException(Exception ex, string sCaller)
        {
            // send an email out with error message
            HSUser oServiceAccount = HSUser.GetUserById(HSUser.SPINCRAFT_SERVICE_ACCOUNT_ID);
            List<string> oToAddresses = new List<string>();
            oToAddresses.Add(HSUser.SPINCRAFT_ROOT_USER);

            StringBuilder oStringBuilder = new StringBuilder();
            oStringBuilder.Clear();
            oStringBuilder.Append("An error occured while executing WeekdayWork for Spincraft -- Subprocess " + sCaller + ":\n");

            oStringBuilder.Append("Exception Message:\n");
            oStringBuilder.Append(ex.Message);
            oStringBuilder.Append("\n\n");
            oStringBuilder.Append("Exception Inner Message:\n");
            oStringBuilder.Append(ex.InnerException);
            oStringBuilder.Append("\n\n");
            HSEmailHelper.SendEmail(oToAddresses, "Error Executing Spincraft Weekday Work", oStringBuilder.ToString());
        }


        public static Session AppSession
        {
            get
            {
                return g_oSession;
            }
        }

        public static Session g_oSession = null;

        public static string g_sConfigLocation = "";

        //public static string g_sSpincraftMAConfiguration = @"E:\Epicor\ERP11\LocalClients\SpincraftMALive\config\SpincraftMAPilotSSO.sysconfig";
        //public static string g_sSpincraftWIConfiguration = @"E:\Epicor\ERP11\LocalClients\SpincraftWILive\config\SpincraftWIPilotSSO.sysconfig";
        //public static string g_sSpincraftMAConfiguration = @"E:\Epicor\ERP11\LocalClients\SpincraftCALive\config\SpincraftMATestSSO.sysconfig";

        // ON RDS SERVER
        //public static string g_sSpincraftMAConfiguration = @"C:\Epicor\ERP11.2.400Client\Client\config\SpincraftMALiveSSO.sysconfig";
        //public static string g_sSpincraftWIConfiguration = @"C:\Epicor\ERP11.2.400Client\Client\config\SpincraftWILiveSSO.sysconfig";
        //public static string g_sSpincraftCAConfiguration = @"C:\Epicor\ERP11.2.400Client\Client\config\SpincraftCALiveSSO.sysconfig";
        //private static string TEMP_FILE_DIRECTORY = @"C:\Epicor\Spincraft\TempFiles\";
        //private static string TEMPLATES_FILE_DIRECTORY = @"C:\Epicor\Spincraft\Templates\";

        // ON APP SERVER
        public static string g_sSpincraftMAConfiguration = @"E:\Epicor\ERP11\LocalClients\SpincraftMALive\config\SpincraftMALiveSSO.sysconfig";
        public static string g_sSpincraftWIConfiguration = @"E:\Epicor\ERP11\LocalClients\SpincraftWILive\config\SpincraftWILiveSSO.sysconfig";
        public static string g_sSpincraftCAConfiguration = @"E:\Epicor\ERP11\LocalClients\SpincraftCALive\config\SpincraftCALiveSSO.sysconfig";
        private static string TEMP_FILE_DIRECTORY = @"E:\Epicor\ERP11\Spincraft\TempFiles\";
        private static string TEMPLATES_FILE_DIRECTORY = @"E:\Epicor\ERP11\Spincraft\Templates\";

        private static string UPLOAD_FILE_DIRECTORY = @"\\ETUS25AN-AP0001\UploadedFiles\";
        private static string ARCHIVE_FILE_DIRECTORY = @"\\ETUS25AN-AP0001\UploadedFiles\Archived\";
    }
}
