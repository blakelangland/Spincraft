using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using System.Globalization;

using Ice.Core;
using System.IO;
using static System.Collections.Specialized.BitVector32;


// MUST TARGET .NET 4.8
// We should reference the adapter and contract classes to both read and write to Epicor
// We start by connecting to Epicor through a session object
// The Epicor business objects (CreateImpl) are used to invoke the methods

namespace HorizonScientific
{
    class Program
    {
        static void Main(string[] args)
        {
            try
            {

                // we will see if this application has any arguments being passed in
                HSUser oRequestingUser = null;
                string sRequestingUserId = "";
                bool bGetUserId = false;

                bool bGetPartNum = false;
                string sPartNum = "";
                DateTime dtEndDate = DateTime.MinValue;
                bool bGetEndDate = false;
                bool bGetStartDate = false;
                DateTime dtStartDate = DateTime.MinValue;

                int iQuoteNum = 0;
                bool bProcessSF1411 = false;
                bool bGetQuoteNum = false;
                bool bGenerateSF1411WhatIf = false;
                bool bQuoteCostBreakdown = false;

                bool bRetrieveJobSchedule = false;
                bool bUpdateJobSchedule = false;
                bool bGenerateMasterSchedule = false;

                bool bImportQuote = false;

                bool bGetCompany = false;
                string sCompany = "";

                string sProductionCalendarId = ProductionCalendarCollection.SPINCRAFT_MA_PRODUCTION_CALENDAR;

                // positive pay parameters
                bool bGeneratePositivePay = false;
                bool bGetStartCheckNum = false;
                int iStartCheckNum = 0;
                bool bGetEndCheckNum = false;
                int iEndCheckNum = 0;
                bool bGetStartCheckDate = false;
                DateTime dtStartCheckDate = DateTime.MinValue;
                bool bGetEndCheckDate = false;
                DateTime dtEndCheckDate = DateTime.MinValue;

                string sCheckNumberPrefix = "";

                bool bGenerateAuditSODReport = false;
                bool bGenerateGLInformation = false;

                bool bGetOrderNum = false;
                int iOrderNum = 0;
                bool bGetOrderLine = false;
                int iOrderLine = 0;
                bool bGetOrderRelNum = false;
                int iOrderRelNum = 0;

                bool bJobEstimateValidation = false;
                bool bJobEstimateValidationByJob = false;
                bool bJobSummary = false;
                bool bForceJobAnalysis = false;
                bool bAcceptActualsForMissingEstimates = false;
                bool bGetJobNum = false;
                string sJobNum = "";
                bool bGetAssemblySequence = false;
                int iAssemblySequence = 0;
                bool bGetOperationSequence = false;
                int iOperationSequence = 0;
                bool bGetEmployeeId = false;
                string sEmployeeId = "";
                bool bCompletePriorOperations = false;

                bool bLoadAssociatedJobs = false;

                bool bCheckIfOperationBeStarted = false;
                bool bAnalyzePartShortages = false;
                bool bAnalyzePartShortagesForOpenOrders = false;
                bool bOnlyPartShortages = false;
                bool bGetCutOffDate = false;
                DateTime dtCutOffDate = DateTime.MinValue;
                decimal dJobEstimatePercentError = 0M;
                bool bGetJobEstimatePercentError = false;
                decimal dJobEstimateAbsoluteError = 0M;
                bool bGetJobEstimateAbsoluteError = false;
                bool bGetJobMarginThreshold = false;
                decimal dJobMarginThreshold = 0M;
                bool bJustMissingCosts = false;

                bool bGenerateOverallEmployeeEffectiveness = false;

                bool bValidateSalesOrders = false;
                bool bValidateJobs = false;
                bool bValidatePurchaseOrders = false;
                bool bValidatePOReceipts = false;

                bool bCalcuateInventoryAge = false;

                // list of all parts in the system -- will get initialized if someone needs it
                HSValidateParts oValidateParts = null;

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
                        if (bGetPartNum == true)
                        {
                            bGetPartNum = false;
                            sPartNum = sArg;
                            continue;
                        }
                        if (bGetQuoteNum == true)
                        {
                            bGetQuoteNum = false;
                            if (int.TryParse(sArg, out iQuoteNum) == false)
                            {
                                iQuoteNum = 0;
                            }
                        }
                        if (bGetOrderNum == true)
                        {
                            bGetOrderNum = false;
                            if (int.TryParse(sArg, out iOrderNum) == false)
                            {
                                iOrderNum = 0;
                            }
                        }
                        if (bGetOrderLine == true)
                        {
                            bGetOrderLine = false;
                            if (int.TryParse(sArg, out iOrderLine) == false)
                            {
                                iOrderLine = 0;
                            }
                        }
                        if (bGetOrderRelNum == true)
                        {
                            bGetOrderRelNum = false;
                            if (int.TryParse(sArg, out iOrderRelNum) == false)
                            {
                                iOrderRelNum = 0;
                            }
                        }
                        if (bGetCompany == true)
                        {
                            bGetCompany = false;
                            sCompany = sArg;
                            continue;
                        }
                        if (bGetJobNum == true)
                        {
                            bGetJobNum = false;
                            sJobNum = sArg;
                            continue;
                        }
                        if (bGetAssemblySequence == true)
                        {
                            bGetAssemblySequence = false;
                            if (int.TryParse(sArg, out iAssemblySequence) == false)
                            {
                                iAssemblySequence = 0;
                            }
                            continue;
                        }
                        if (bGetOperationSequence == true)
                        {
                            bGetOperationSequence = false;
                            if (int.TryParse(sArg, out iOperationSequence) == false)
                            {
                                iOperationSequence = 0;
                            }
                            continue;
                        }
                        if (bGetEmployeeId == true)
                        {
                            bGetEmployeeId = false;
                            sEmployeeId = sArg;
                            continue;
                        }
                        if (bGetJobEstimatePercentError == true)
                        {
                            bGetJobEstimatePercentError = false;
                            if (Decimal.TryParse(sArg, out dJobEstimatePercentError) == false)
                            {
                                dJobEstimatePercentError = 0M;
                            }
                            // need to divide by 100
                            //dJobEstimatePercentError = dJobEstimatePercentError / 100.0M;
                            continue;
                        }
                        if (bGetJobEstimateAbsoluteError == true)
                        {
                            bGetJobEstimateAbsoluteError = false;
                            if (Decimal.TryParse(sArg, out dJobEstimateAbsoluteError) == false)
                            {
                                dJobEstimateAbsoluteError = 0M;
                            }
                            continue;
                        }
                        if (bGetJobMarginThreshold == true)
                        {
                            bGetJobMarginThreshold = false;
                            if (Decimal.TryParse(sArg, out dJobMarginThreshold) == false)
                            {
                                dJobMarginThreshold = 0M;
                            }
                            continue;
                        }
                        if (bGetStartCheckNum == true)
                        {
                            bGetStartCheckNum = false;
                            if (int.TryParse(sArg, out iStartCheckNum) == false)
                            {
                                iStartCheckNum = 0;
                            }
                            continue;
                        }
                        if (bGetEndCheckNum == true)
                        {
                            bGetEndCheckNum = false;
                            if (int.TryParse(sArg, out iEndCheckNum) == false)
                            {
                                iEndCheckNum = 0;
                            }
                            continue;
                        }
                        if (bGetStartCheckDate == true)
                        {
                            bGetStartCheckDate = false;
                            try
                            {
                                dtStartCheckDate = DateTime.Parse(sArg, CultureInfo.InvariantCulture, DateTimeStyles.RoundtripKind);
                            }
                            catch (Exception)
                            {
                                dtStartCheckDate = DateTime.MinValue;
                            }

                            //////////if (DateTime.TryParse(sArg, out dtStartCheckDate) == false)
                            //////////{
                            //////////    dtStartCheckDate = DateTime.MinValue;
                            //////////}
                            continue;
                        }
                        if (bGetEndCheckDate == true)
                        {
                            bGetEndCheckDate = false;
                            try
                            {
                                dtEndCheckDate = DateTime.Parse(sArg, CultureInfo.InvariantCulture, DateTimeStyles.RoundtripKind);
                            }
                            catch (Exception)
                            {
                                dtEndCheckDate = DateTime.MinValue;
                            }

                            //////////if (DateTime.TryParse(sArg, out dtEndCheckDate) == false)
                            //////////{
                            //////////    dtEndCheckDate = DateTime.MinValue;
                            //////////}
                            continue;
                        }
                        if (bGetStartDate == true)
                        {
                            bGetStartDate = false;
                            try
                            {
                                dtStartDate = DateTime.Parse(sArg, CultureInfo.InvariantCulture, DateTimeStyles.RoundtripKind);
                            }
                            catch (Exception)
                            {
                                dtStartDate = DateTime.MinValue;
                            }

                            //////////if (DateTime.TryParse(sArg, out dtEndCheckDate) == false)
                            //////////{
                            //////////    dtStartDate = DateTime.MinValue;
                            //////////}
                            continue;
                        }
                        if (bGetEndDate == true)
                        {
                            bGetEndDate = false;
                            try
                            {
                                dtEndDate = DateTime.Parse(sArg, CultureInfo.InvariantCulture, DateTimeStyles.RoundtripKind);
                            }
                            catch (Exception)
                            {
                                dtEndDate = DateTime.MinValue;
                            }

                            //////////if (DateTime.TryParse(sArg, out dtEndCheckDate) == false)
                            //////////{
                            //////////    dtEndDate = DateTime.MinValue;
                            //////////}
                            continue;
                        }

                        else if (string.Compare(sArg, "USER_ID", true) == 0)
                        {
                            // the next argument will be the user id
                            bGetUserId = true;
                        }
                        else if (string.Compare(sArg, "PROCESS_SF1411", true) == 0)
                        {
                            bProcessSF1411 = true;
                        }
                        else if (string.Compare(sArg, "SF1411_WHATIF", true) == 0)
                        {
                            bGenerateSF1411WhatIf = true;
                        }
                        else if (string.Compare(sArg, "QUOTE_COST_BREAKDOWN", true) == 0)
                        {
                            bQuoteCostBreakdown = true;
                        }
                        else if (string.Compare(sArg, "IMPORT_QUOTE", true) == 0)
                        {
                            bImportQuote = true;
                        }
                        else if (string.Compare(sArg, "QUOTE_NUM", true) == 0)
                        {
                            // the next argument will be the quote num
                            bGetQuoteNum = true;
                        }
                        else if (string.Compare(sArg, "ORDER_NUM", true) == 0)
                        {
                            // next argument will be the order num
                            bGetOrderNum = true;
                        }
                        else if (string.Compare(sArg, "ORDER_LINE", true) == 0)
                        {
                            // next argument will be the order line
                            bGetOrderLine = true;
                        }
                        else if (string.Compare(sArg, "ORDER_REL_NUM", true) == 0)
                        {
                            // next argument will be the order rel num
                            bGetOrderRelNum = true;
                        }
                        else if (string.Compare(sArg, "PART_NUM", true) == 0)
                        {
                            // the next argument will be the part num
                            bGetPartNum = true;
                        }
                        else if (string.Compare(sArg, "COMPANY", true) == 0)
                        {
                            // the next argument will be the company id
                            bGetCompany = true;
                        }
                        else if (string.Compare(sArg, "JOB_NUM", true) == 0)
                        {
                            // the next argument will be the job num
                            bGetJobNum = true;
                        }
                        else if (string.Compare(sArg, "ASM_SEQ", true) == 0)
                        {
                            // the next argument will be the assembly seq
                            bGetAssemblySequence = true;
                        }
                        else if (string.Compare(sArg, "OPR_SEQ", true) == 0)
                        {
                            // the next argument will be the operation seq
                            bGetOperationSequence = true;
                        }
                        else if (string.Compare(sArg, "EMPLOYEE_ID", true) == 0)
                        {
                            // the next argument will be the employee id
                            bGetEmployeeId = true;
                        }
                        else if (string.Compare(sArg, "TASK_SCHEDULER", true) == 0)
                        {
                            // this process was kicked off from the task scheduler
                            bExecutedFromTaskScheduler = true;
                        }
                        else if (string.Compare(sArg, "CUT_OFF_DATE", true) == 0)
                        {
                            // the next argument will be the cut off date
                            bGetCutOffDate = true;
                        }
                        else if (string.Compare(sArg, "JOB_ONLY_MISSING_COSTS", true) == 0)
                        {
                            bJustMissingCosts = true;
                        }
                        else if (string.Compare(sArg, "JOB_EST_PERCENT_ERROR", true) == 0)
                        {
                            bGetJobEstimatePercentError = true;
                        }
                        else if (string.Compare(sArg, "JOB_EST_ABSOLUTE_ERROR", true) == 0)
                        {
                            bGetJobEstimateAbsoluteError = true;
                        }
                        else if (string.Compare(sArg, "JOB_MARGIN_THRESHOLD", true) == 0)
                        {
                            bGetJobMarginThreshold = true;
                        }
                        else if (string.Compare(sArg, "RETRIEVE_JOB_SCHEDULE", true) == 0)
                        {
                            bRetrieveJobSchedule = true;
                        }
                        else if (string.Compare(sArg, "UPDATE_JOB_SCHEDULE", true) == 0)
                        {
                            bUpdateJobSchedule = true;
                        }
                        else if (string.Compare(sArg, "ONLY_PART_SHORTAGES", true) == 0)
                        {
                            bOnlyPartShortages = true;
                        }
                        else if (string.Compare(sArg, "ANALYZE_PART_SHORTAGES", true) == 0)
                        {
                            bAnalyzePartShortages = true;
                        }
                        else if (string.Compare(sArg, "ANALYZE_PART_SHORTAGS_FOR_ORDERS", true) == 0)
                        {
                            bAnalyzePartShortagesForOpenOrders = true;
                        }
                        else if (string.Compare(sArg, "GENERATE_MASTER_SCHEDULE", true) == 0)
                        {
                            bGenerateMasterSchedule = true;
                        }
                        else if (string.Compare(sArg, "GENERATE_POSITIVE_PAY", true) == 0)
                        {
                            bGeneratePositivePay = true;
                        }
                        else if (string.Compare(sArg, "START_CHECK_NUM", true) == 0)
                        {
                            bGetStartCheckNum = true;
                        }
                        else if (string.Compare(sArg, "END_CHECK_NUM", true) == 0)
                        {
                            bGetEndCheckNum = true;
                        }
                        else if (string.Compare(sArg, "START_CHECK_DATE", true) == 0)
                        {
                            bGetStartCheckDate = true;
                        }
                        else if (string.Compare(sArg, "END_DATE", true) == 0)
                        {
                            bGetEndDate = true;
                        }
                        else if (string.Compare(sArg, "START_DATE", true) == 0)
                        {
                            bGetStartDate = true;
                        }
                        else if (string.Compare(sArg, "END_CHECK_DATE", true) == 0)
                        {
                            bGetEndCheckDate = true;
                        }
                        else if (string.Compare(sArg, "GENERATE_AUDIT_SOD", true) == 0)
                        {
                            bGenerateAuditSODReport = true;
                        }
                        else if (string.Compare(sArg, "JOB_ESTIMATE_VALIDATION", true) == 0)
                        {
                            bJobEstimateValidation = true;
                        }
                        else if (string.Compare(sArg, "JOB_ESTIMATE_VALIDATION_BY_JOB", true) == 0)
                        {
                            bJobEstimateValidationByJob = true;
                        }
                        else if (string.Compare(sArg, "JOB_SUMMARY", true) == 0)
                        {
                            bJobSummary = true;
                        }
                        else if (string.Compare(sArg, "FORCE_JOB_ANALYSIS", true) == 0)
                        {
                            bForceJobAnalysis = true;
                        }
                        else if (string.Compare(sArg, "CHECK_IF_OP_CAN_START", true) == 0)
                        {
                            bCheckIfOperationBeStarted = true;
                        }
                        else if (string.Compare(sArg, "COMPLETE_PRIOR_OPERATIONS", true) == 0)
                        {
                            bCompletePriorOperations = true;
                        }
                        else if (string.Compare(sArg, "LOAD_ASSOCIATED_JOBS", true) == 0)
                        {
                            bLoadAssociatedJobs = true;
                        }
                        else if (string.Compare(sArg, "USE_ACT_FOR_MISSING_EST", true) == 0)
                        {
                            bAcceptActualsForMissingEstimates = true;
                        }
                        else if (string.Compare(sArg, "GENERATE_OVERALL_EMPLOYEE_EFFECTIVENESS", true) == 0)
                        {
                            bGenerateOverallEmployeeEffectiveness = true;
                        }
                        else if (string.Compare(sArg, "GENERATE_GL_INFORMATION", true) == 0)
                        {
                            bGenerateGLInformation = true;
                        }
                        else if (string.Compare(sArg, "CALCULATE_INVENTORY_AGE", true) == 0)
                        {
                            bCalcuateInventoryAge = true;
                        }
                    }
                }

                if (bExecutedFromTaskScheduler == true)
                {
                    // tasks we perform when kicked off from the task scheduler
                    bValidateSalesOrders = true;
                    bValidateJobs = true;
                    bCheckIfOperationBeStarted = false;
                    bLoadAssociatedJobs = false;
                    bCompletePriorOperations = false;
                    bValidatePurchaseOrders = true;
                    bValidatePOReceipts = true;
                }

                // get rid of any existing files in this temp directory to clean things up
                DirectoryInfo oDirectoryInfo = new DirectoryInfo(TEMP_FILE_DIRECTORY);
                foreach (FileInfo oTmpFile in oDirectoryInfo.EnumerateFiles())
                {
                    File.Delete(oTmpFile.FullName);
                }

                if (string.Compare(sCompany, CompanyConfiguration.SPINCRAFT_MA_COMPANY_ID, true) == 0)
                {
                    g_sConfigLocation = g_sSpincraftMAConfiguration;
                    sCompany = CompanyConfiguration.SPINCRAFT_MA_COMPANY_ID;
                    sCheckNumberPrefix = "06600";
                    sProductionCalendarId = ProductionCalendarCollection.SPINCRAFT_MA_PRODUCTION_CALENDAR;
                }
                else
                {
                    // default is WI if not stated
                    g_sConfigLocation = g_sSpincraftWIConfiguration;
                    sCompany = CompanyConfiguration.SPINCRAFT_WI_COMPANY_ID;
                    sCheckNumberPrefix = "06500";
                    sProductionCalendarId = ProductionCalendarCollection.SPINCRAFT_WI_PRODUCTION_CALENDAR;
                }
  

                g_oSession = new Session(HSUser.SPINCRAFT_SERVICE_ACCOUNT_ID, HSUser.SPINCRAFT_SERVICE_PASSWORD, Session.LicenseType.Default, g_sConfigLocation);
                if (g_oSession != null)
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

                    // TESTING EMAIL SETUP
                    //HSUser.PerformValidation(TEMP_FILE_DIRECTORY);

                    oValidateParts = null;

                    // THIS WAS JUST TO TEST CODE
                    //sPartNum = "636-1360-1-BLK";
                    //HSPartData.Initialize(AppSession, sPartNum);

                    //
                    // daily work
                    //
                    if (bCheckIfOperationBeStarted == true)
                    {
                        //sJobNum = "31239-39";
                        //iAssemblySequence = 1;
                        //iOperationSequence = 130;
                        string sResult = "";
                        // we need the job num, assembly sequence, and operation sequence
                        AnalyzeOperationsToComplete oAnalyzeOperationsToComplete = new AnalyzeOperationsToComplete();
                        oAnalyzeOperationsToComplete.LoadDataForJob(AppSession, sJobNum, false);
                        bool bCanOperationBeStarted = oAnalyzeOperationsToComplete.CanOperationBeStarted(iAssemblySequence, iOperationSequence, out sResult);
                        if (bCanOperationBeStarted == false)
                        {
                            Console.WriteLine("Prior Operations Must Be Completed First:\n");
                            Console.WriteLine(sResult);
                        }
                        else
                        {
                            Console.WriteLine("Job: " + sJobNum + " Asm: " + iAssemblySequence.ToString() + " Opr: " + iOperationSequence.ToString() + " can be started.");
                        }
                    }
                    if (bCompletePriorOperations == true)
                    {
                        //sJobNum = "31306-02";
                        //iOrderNum = 24056;
                        //iOrderLine = 1;
                        //iOrderRelNum = 1;
                        //sPartNum = "2088FA";
                        //iAssemblySequence = 2;
                        //iOperationSequence = 10;
                        //sEmployeeId = "168";
                        //bLoadAssociatedJobs = true;

                        string sOperationIssues = "";
                        // we need the job num, assembly sequence, and operation sequence
                        List<AnalyzeOperationsToComplete> oAllAnalysis = new List<AnalyzeOperationsToComplete>();
                        if (string.IsNullOrEmpty(sJobNum) == false)
                        {
                            AnalyzeOperationsToComplete oAnalyzeOperationsToComplete = new AnalyzeOperationsToComplete();
                            oAnalyzeOperationsToComplete.LoadDataForJob(AppSession, sJobNum, bLoadAssociatedJobs);
                            oAllAnalysis.Add(oAnalyzeOperationsToComplete);
                        }
                        else if (string.IsNullOrEmpty(sPartNum) == false)
                        {
                            // we need to read in a list of jobs
                            List<string> oJobNums = AnalyzeOperationsToComplete.GetJobListByPartNum(AppSession, sPartNum);
                            // we now need to set up a list of these analyze job operation objects
                            foreach (string sTmpJonNum in oJobNums)
                            {
                                AnalyzeOperationsToComplete oAnalyzeOperationsToComplete = new AnalyzeOperationsToComplete();
                                oAnalyzeOperationsToComplete.LoadDataForJob(AppSession, sTmpJonNum, bLoadAssociatedJobs);
                                oAllAnalysis.Add(oAnalyzeOperationsToComplete);
                            }
                        }    
                        else  if ( (iOrderNum != 0) && (iOrderLine != 0) && (iOrderRelNum != 0) )
                        {
                            AnalyzeOperationsToComplete oAnalyzeOperationsToComplete = new AnalyzeOperationsToComplete();
                            oAnalyzeOperationsToComplete.LoadDataForJob(AppSession, iOrderNum, iOrderLine, iOrderRelNum, bLoadAssociatedJobs);
                            oAllAnalysis.Add(oAnalyzeOperationsToComplete);
                        }

                        // now we need to walk through the list of jobs and their operations
                        foreach (AnalyzeOperationsToComplete oAnalyze in oAllAnalysis)
                        {
                            List<JobOperationHierarchy> oOperationsToComplete = oAnalyze.EvaluateOperationsToComplete(iAssemblySequence, iOperationSequence);
                            sOperationIssues = oAnalyze.CompleteOperations(AppSession, oOperationsToComplete, sEmployeeId);

                            //oAnalyze.TotalNumberOfOperations.ToString();
                            //oAnalyze.NumberOfOperationsCompleted.ToString();
                            //oAnalyze.PercentageCompleteByTimeEstiamtes.ToString();
                            //oAnalyze.TotalEstimatedTime.ToString();
                            //oAnalyze.EstimatedTimeCompleted.ToString();
                            //oAnalyze oNextOp = oAnalyzeOperationsToComplete.NextOperationsToComplete();

                            if (string.IsNullOrEmpty(sOperationIssues) == false)
                            {
                                Exception ex = new Exception(sOperationIssues);
                                ReportException(ex, "OperationsToComplete");
                            }
                        }
                    }

                    if (bValidateJobs == true)
                    {
                        try
                        {
                            if (oValidateParts == null)
                            {
                                oValidateParts = new HSValidateParts();
                                oValidateParts.Initialize(AppSession);
                            }

                            HSJobValidation oJobValidation = new HSJobValidation(sCompany);
                            if (oJobValidation.Initialize(AppSession, oValidateParts) == false)
                            {
                                Console.WriteLine("Failed to load all sales orders!");
                                return;
                            }
                            oJobValidation.PerformValidation(sCompany, TEMP_FILE_DIRECTORY, oRequestingUser);
                        }
                        catch (Exception ex)
                        {
                            ReportException(ex, "Validate Jobs");
                        }
                    }

                    if (bValidateSalesOrders == true)
                    {
                        try
                        {
                            if (oValidateParts == null)
                            {
                                oValidateParts = new HSValidateParts();
                                oValidateParts.Initialize(AppSession);
                            }

                            ValidateSalesOrders oValidateSalesOrders = new ValidateSalesOrders();
                            if (oValidateSalesOrders.Initialize(AppSession, sCompany, oValidateParts) == false)
                            {
                                Console.WriteLine("Failed to load all sales orders!");
                                return;
                            }
                            oValidateSalesOrders.PerformValidation(sCompany, TEMP_FILE_DIRECTORY);
                        }
                        catch (Exception ex)
                        {
                            ReportException(ex, "Validate Sales Order");
                        }
                    }

                    if (bValidatePurchaseOrders == true)
                    {
                        try
                        {
                            if (oValidateParts == null)
                            {
                                oValidateParts = new HSValidateParts();
                                oValidateParts.Initialize(AppSession);
                            }

                            HSPOValidation oValidatePurchaseOrders = new HSPOValidation();
                            if (oValidatePurchaseOrders.Initialize(AppSession, oValidateParts) == false)
                            {
                                Console.WriteLine("Failed to load all purchase orders!");
                                return;
                            }
                            oValidatePurchaseOrders.PerformValidation(sCompany, TEMP_FILE_DIRECTORY);
                        }
                        catch (Exception ex)
                        {
                            ReportException(ex, "Validate Purchase Orders");
                        }
                    }

                    if (bValidatePOReceipts == true)
                    {
                        try
                        {
                            HSPOReceiptValidation oReceiptValidation = new HSPOReceiptValidation();
                            if (oReceiptValidation.Initialize(AppSession) == false)
                            {
                                Console.WriteLine("Failed to load all purchase orders receipts!");
                                return;
                            }
                            oReceiptValidation.PerformValidation(sCompany, TEMP_FILE_DIRECTORY);
                        }
                        catch (Exception ex)
                        {
                            ReportException(ex, "Validate Purchase Orders");
                        }
                    }

                    if (bProcessSF1411 == true)
                    {
                        try
                        {
                            //iQuoteNum = 5763;
                            //iQuoteNum = 5704;
                            List<HSQuote> oQuotes = HSQuote.Initialize(AppSession, iQuoteNum);
                            if ((oQuotes != null) && (oQuotes.Count > 0))
                            {
                                HSQuote.CreateAndSendReport(TEMP_FILE_DIRECTORY, oRequestingUser, oQuotes);
                            }
                            else
                            {
                                Console.WriteLine("Fail to load SF1411 quote data");
                                return;
                            }
                        }
                        catch (Exception ex)
                        {
                            ReportException(ex, "QuoteSF1411");
                        }
                    }

                    if (bGenerateSF1411WhatIf == true)
                    {
                        try
                        {
                            //iQuoteNum = 5618;
                            //iQuoteNum = 5704;
                            //iQuoteNum = 5745;
                            List<HSQuote> oQuotes = HSQuote.Initialize(AppSession, iQuoteNum);
                            if ((oQuotes != null) && (oQuotes.Count > 0))
                            {
                                HSQuote.CreateAndSendWhatIfReport(TEMP_FILE_DIRECTORY, oRequestingUser, oQuotes);
                            }
                            else
                            {
                                Console.WriteLine("Fail to load SF1411 quote data");
                                return;
                            }
                        }
                        catch (Exception ex)
                        {
                            ReportException(ex, "QuoteSF1411");
                        }
                    }

                    if (bQuoteCostBreakdown == true)
                    {
                        try
                        {
                            //iQuoteNum = 5618;
                            //iQuoteNum = 5585;
                            List<HSQuote> oQuotes = HSQuote.Initialize(AppSession, iQuoteNum);
                            if ((oQuotes != null) && (oQuotes.Count > 0))
                            {
                                HSQuote.CreateQuoteCostBreakdown(TEMPLATES_FILE_DIRECTORY + "\\Quote Cost Breakdown.xlsx", TEMP_FILE_DIRECTORY, oRequestingUser, oQuotes);
                            }
                            else
                            {
                                Console.WriteLine("Fail to load quote data for cost breakdown!");
                                return;
                            }
                        }
                        catch (Exception ex)
                        {
                            ReportException(ex, "Quote Cost Breakdown");
                        }
                    }

                    if (bImportQuote == true)
                    {
                        try
                        {
                            if (HSQuote.ImportQuotes(AppSession, UPLOAD_FILE_DIRECTORY + "Quote.xlsx", oRequestingUser, TEMP_FILE_DIRECTORY) == false)
                            {
                                Console.WriteLine("Failed to import quote data");
                            }
                        }
                        catch (Exception ex)
                        {
                            ReportException(ex, "Import Quote");
                        }
                    }

                    if (bRetrieveJobSchedule == true)
                    {
                        try
                        {
                            List<HSOperationDetail> oOperationDetails = HSJobSchedule.Initialize(AppSession, sPartNum);
                            if ((oOperationDetails != null) && (oOperationDetails.Count > 0))
                            {
                                HSJobSchedule.CreateAndSendReport(TEMP_FILE_DIRECTORY, oRequestingUser, oOperationDetails);
                            }
                            else
                            {
                                Console.WriteLine("Fail to load operation details data");
                                return;
                            }
                        }
                        catch (Exception ex)
                        {
                            ReportException(ex, "Process Job Schedule");
                        }
                    }

                    if (bUpdateJobSchedule == true)
                    {
                        try
                        {
                            // get the op details as they currently are in the system
                            List<HSOperationDetail> oOriginalOperationDetails = HSJobSchedule.Initialize(AppSession, sPartNum);
                            HSJobSchedule.UpdateJobOperations(AppSession, UPLOAD_FILE_DIRECTORY + "JobScheduleUpdates.xlsx", oRequestingUser, TEMP_FILE_DIRECTORY, oOriginalOperationDetails);
                        }
                        catch (Exception ex)
                        {
                            ReportException(ex, "Process Job Schedule");
                        }
                    }

                    if (bGenerateMasterSchedule == true)
                    {
                        try
                        {
                            // get the op details as they currently are in the system
                            //sPartNum = "901-062-004-405";
                            List<HSOperationDetail> oOriginalOperationDetails = HSJobSchedule.Initialize(AppSession, sPartNum);

                            HSJobSchedule.UpdateJobOperations(AppSession, UPLOAD_FILE_DIRECTORY + "JobScheduleUpdates.xlsx", oRequestingUser, TEMP_FILE_DIRECTORY, oOriginalOperationDetails);
                        }
                        catch (Exception ex)
                        {
                            ReportException(ex, "Process Job Schedule");
                        }
                    }

                    if ((bJobEstimateValidation == true) || (bJobEstimateValidationByJob == true) || (bJobSummary == true))
                    {
                        try
                        {
                            //sJobNum = "31118-TEST";
                            //sJobNum = "31304-10";
                            //sJobNum = "013160-7-3";
                            //dJobEstimateAbsoluteError = 1000.0M;
                            //dJobEstimatePercentError = 10.0M;
                            //dJobMarginThreshold = 25.0M;
                            //sPartNum = "2085FASP";
                            //bAcceptActualsForMissingEstimates = true;

                            if (oValidateParts == null)
                            {
                                oValidateParts = new HSValidateParts();
                                oValidateParts.Initialize(AppSession);
                            }

                            // load the materials and operations for all open jobs
                            HSJobEstimateValidation oJobEstimateValidator = new HSJobEstimateValidation(sCompany, dJobEstimateAbsoluteError, dJobEstimatePercentError, dJobMarginThreshold, bForceJobAnalysis, bAcceptActualsForMissingEstimates, bJustMissingCosts);

                            oJobEstimateValidator.Initialize(AppSession, oValidateParts, sJobNum, sPartNum);
                            if (bJobEstimateValidation == true)
                            {
                                oJobEstimateValidator.PerformJobValidation(sCompany, TEMP_FILE_DIRECTORY, oRequestingUser);
                            }
                            if (bJobEstimateValidationByJob == true)
                            {
                                oJobEstimateValidator.PerformJobValidationByJob(sCompany, TEMP_FILE_DIRECTORY, oRequestingUser);
                            }
                            if (bJobSummary == true)
                            {
                                oJobEstimateValidator.PerformJobSummary(sCompany, TEMP_FILE_DIRECTORY, oRequestingUser);
                            }
                        }
                        catch (Exception ex)
                        {
                            ReportException(ex, "Job Estimate Validation");
                        }
                    }

                    if (bAnalyzePartShortages == true)
                    {
                        HSAnalyzePartShortages oAnalyzePartShortages = new HSAnalyzePartShortages();
                        if (oAnalyzePartShortages.Initialize(AppSession, sCompany, dtCutOffDate) == false)
                        {
                            Console.WriteLine("Failed to load all part shortage information!");
                            return;
                        }
                        //sJobNum = "350710";
                        //bOnlyPartShortages = true;
                        //dtCutOffDate = DateTime.Now.AddMonths(1);
                        oAnalyzePartShortages.CreateReport(ARCHIVE_FILE_DIRECTORY, TEMP_FILE_DIRECTORY, sJobNum, bOnlyPartShortages, dtCutOffDate, oRequestingUser);
                    }

                    //
                    // review what we can build that has been orderd and we have parts for
                    //
                    if (bAnalyzePartShortagesForOpenOrders == true)
                    {
                        HSAnalyzePartShortages oAnalyzePartShortages = new HSAnalyzePartShortages();
                        if (oAnalyzePartShortages.Initialize2(AppSession, sCompany) == false)
                        {
                            Console.WriteLine("Failed to load all part shortage information!");
                            return;
                        }

                        //
                        // WHAT ABOUT PO THAT IS BUY DIRECT TO A JOB - THIS SHOULD NOT BE COUNTED AS PARTS THAT ARE NOT GENERALLY AVAILABLE AS IT CAN ONLY SUPPLY THAT JOB
                        //

                        //
                        // WHAT ABOUT JOBS THAT DEPEND ON OTHER JOBS THAT HAVE PART SHORTAGES???
                        // WE NEED TO BUILD UP A DEPENDENCY TREE TO SEE IF WE HAVE THE PARTS 
                        // NEEDED FOR ALL CHILD JOBS
                        //
                    }

                    if (bGenerateOverallEmployeeEffectiveness == true)
                    {
                    }

                    if (bCalcuateInventoryAge == true)
                    {
                        try
                        {
                            // need to compute the age of inventory on hand
                            HSCalculateInventoryAge oCalculateInventoryAge = new HSCalculateInventoryAge();
                            oCalculateInventoryAge.Initialize(AppSession);
                            oCalculateInventoryAge.CreateReport(ARCHIVE_FILE_DIRECTORY, TEMP_FILE_DIRECTORY, sCompany, oRequestingUser);
                        }
                        catch (Exception ex)
                        {
                            ReportException(ex, "CalculateInventoryAge");
                        }
                    }

                    #region Positive Pay
                    if (bGeneratePositivePay == true)
                    {
                        try
                        {
                            //iStartCheckNum = 41191;
                            //iEndCheckNum = 41197;
                            //dtStartCheckDate = new DateTime(2026, 1, 1);
                            //dtEndCheckDate = new DateTime(2026, 1, 15);
                            if (HSPositivePay.Initialize(AppSession, iStartCheckNum, iEndCheckNum, dtStartCheckDate, dtEndCheckDate) == false)
                            {
                                Console.WriteLine("Failed to load all Positive Pay Documents!");
                                return;
                            }
                            HSPositivePay.GenerateFile(ARCHIVE_FILE_DIRECTORY, TEMP_FILE_DIRECTORY, sCheckNumberPrefix, oRequestingUser);
                        }
                        catch (Exception ex)
                        {
                            ReportException(ex, "PositivePay");
                        }
                    }
                    #endregion

                    #region Audit SOD
                    if (bGenerateAuditSODReport == true)
                    {
                        try
                        {
                            if (HSAudit.Initialize(AppSession) == false)
                            {
                                Console.WriteLine("Failed to load all Audit SOD Documents!");
                                return;
                            }
                            HSAudit.GenerateReport(TEMP_FILE_DIRECTORY, oRequestingUser);
                        }
                        catch (Exception ex)
                        {
                            ReportException(ex, "Generate SOD Report");
                        }
                    }
                    #endregion

                    #region GL Information
                    if (bGenerateGLInformation == true)
                    {
                        HSGLAccountCategories oAccountCategories = new HSGLAccountCategories();
                        oAccountCategories.Initialize(AppSession);
                        oAccountCategories.GenerateReport(sCompany, TEMP_FILE_DIRECTORY, oRequestingUser);
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
                oStringBuilder.Append("An error occured while executing Weekday Work:\n");

                oStringBuilder.Append("Exception Message:\n");
                oStringBuilder.Append(error.Message);
                oStringBuilder.Append("\n\n");
                oStringBuilder.Append("Exception Inner Message:\n");
                oStringBuilder.Append(error.InnerException);
                oStringBuilder.Append("\n\n");
                HSEmailHelper.SendEmail(oToAddresses, "Error Executing Weekday Work", oStringBuilder.ToString());
            }
            finally
            {
                HSUser oServiceAccount = HSUser.GetUserById(HSUser.SPINCRAFT_SERVICE_ACCOUNT_ID);
                List<string> oToAddresses = new List<string>();
                oToAddresses.Add(HSUser.SPINCRAFT_ROOT_USER);

                StringBuilder oStringBuilder = new StringBuilder();
                oStringBuilder.Clear();
                oStringBuilder.Append("The Process WeekdayWork was executed for Spincraft.");

                HSEmailHelper.SendEmail(oToAddresses, "The Weekday Work Process Ran", oStringBuilder.ToString());

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
