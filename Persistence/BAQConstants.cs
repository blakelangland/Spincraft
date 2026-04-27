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
    public static class StringExt
    {
        public static string Truncate(this string variable, int Length)
        {
            if (string.IsNullOrEmpty(variable))
                return variable;
            return variable.Length <= Length ? variable : variable.Substring(0, Length);
        }

        public static string CleanString(string sTmp)
        {
            StringBuilder oReplacement = new StringBuilder();

            // fix description
            foreach (char s in sTmp)
            {
                if (char.IsLetterOrDigit(s) == true)
                {
                    oReplacement.Append(s);
                }
                else if ((s == ' ') || (s == '-') || (s == '\t') || (s == '\r') || (s == '\n'))
                {
                    oReplacement.Append(s);
                }
                else
                {
                    oReplacement.Append(' ');
                }
            }
            return oReplacement.ToString();
        }
        public static bool ContainsAny(this string input, IEnumerable<string> containsKeywords, StringComparison comparisonType = StringComparison.InvariantCulture)
        {
            return containsKeywords.Any(keyword => input.IndexOf(keyword, comparisonType) >= 0);
        }

        public static void FindDifferenceInStrings(string s1, string s2)
        {
            int maxLength = Math.Max(s1.Length, s2.Length);
            bool differencesFound = false;

            for (int i = 0; i < maxLength; i++)
            {
                char c1 = i < s1.Length ? s1[i] : '\0'; // '\0' = end of string
                char c2 = i < s2.Length ? s2[i] : '\0';

                if (c1 != c2)
                {
                    differencesFound = true;
                    Console.WriteLine($"Difference at position {i}: '{c1}' vs '{c2}'");
                }
            }

            if (!differencesFound)
            {
                Console.WriteLine("The strings match exactly.");
            }
        }
    }

    public static class BAQConstants
    {
        public static string QUERY_GET_JOBS_TIED_TO_PART_NUM = @"GetJobsTiedToPartNum";
        public static string QUERY_GET_JOB_TIED_TO_ORDER_LINE_RELEASE = @"GetJobTiedToOrderLineRelease";
        public static string QUERY_GET_JOBS_TIED_TO_JOB = @"GetJobsTiedToJob";

        public static string QUERY_GET_JOB_OPERATION_HIERARCHY = @"GetJobOperationHierarchy";
        public static string QUERY_GET_JOB_MATERIAL_HIERARCHY = @"GetJobMaterialHierarchy";
        public static string QUERY_GET_JOB_ASSEMBLY_HIERARCHY = @"GetJobAssemblyHierarchy";
        public static string QUERY_GET_EMPLOYEE_INFO = @"GetEmployeeInfo";


        public static string QUERY_QUOTE_COSTS = @"QuoteCosts";
        public static string QUERY_QUOTE_OPERATION_DETAILS = @"QuoteOperationDetails";
        public static string QUERY_QUOTE_MATERIAL_DETAILS = @"QuoteMaterialDetails";

        public static string QUERY_QUOTE_OPERATIONS = @"QuoteOperations";
        public static string QUERY_JOB_OPERATION_DETAILS_BY_PART = @"JobOpDetailsByPart";

        public static string QUERY_POSITIVE_PAY = @"PositivePay";

        public static string QUERY_LIST_ALL_USERS = "ListAllUsers";
        public static string QUERY_ALL_USERS = @"AllUsers";
        public static string QUERY_ACTIVE_GROUPS = @"ActiveGroups";
        public static string QUERY_AUDIT_MENU_SECURITY = @"AuditMenuSecurity";

        public static string QUERY_PRODUCTION_CALENDAR = @"ProductionCalendar";

        public static string QUERY_LIST_ALL_PARTS = @"ListAllParts";

        public static string QUERY_LIST_ALL_MFG_PARTS_FOR_BOM_COMPARISON = @"ListAllMfgPartsForBOMComp";
        public static string LIST_ALL_PART_MTL = @"ListAllPartMtl";
        public static string QUERY_LIST_ALL_PART_OPERATIONS = @"ListAllPartOperations";

        public static string QUERY_LIST_ALL_OPERATIONS = @"ListAllOperations";
        public static string QUERY_LIST_ALL_RESOURCE_GROUPS = @"ListAllResourceGroups";
        public static string QUERY_LIST_ALL_RESOURCES = @"ListAllResources";

        public static string QUERY_LIST_ALL_UNFIRM_JOBS = @"ListAllUnfirmJobs";

        public static string QUERY_LIST_ALL_JOB_MATERIALS = @"ListAllJobMaterials";
        public static string QUERY_LIST_ALL_JOB_OPERATIONS = @"ListAllJobOperations";
        public static string QUERY_JOB_OPS_EST_VS_ACTUAL_COSTS = @"JobOpsEstVsActualCosts";
        public static string QUERY_JOB_EST_VS_ACTUAL_COSTS_QTY = @"JobEstVsActualCostsQty";

        // used for validating open sales orders
        public static string QUERY_SALES_ORDER_BACKLOG = @"SalesOrderBacklog";

        // used for validating purchase orders
        public static string QUERY_LIST_VALIDATE_PURCHASE_ORDERS = @"valPurchaseOrder";

        // used for validating PO receipts
        public static string QUERY_VALIDATE_PO_RECEIPT_FOR_INSPECTION = @"ValidatePORecieptForInsp";
        public static string QUERY_RECEIVE_DETAIL_OPEN = @"RcvDtlOpen";

        // used for validating shipments
        public static string QUERY_PACKS_WITHOUT_LINES = @"PacksWithoutLines";
        public static string QUERY_PACKS_STILL_OPEN = @"PacksStillOpen";
        public static string QUERY_PACKS_CLOSED_NOT_SHIPPED = @"PacksClosedNotShipped";

        // PULL IN NEW QUERIES
        public static string QUERY_OPEN_JOBS = @"OpenJobs";
        public static string QUERY_SOURCE_FOR_PARTS_FROM_PO = @"SourceForPartsFromPOs";
        public static string QUERY_DEMAND_FOR_PARTS_IN_TIME = @"DemandForPartsInTime";
        public static string QUERY_PART_DEMAND_FROM_REQUISITIONS = @"PartDemandFromReqs";
        public static string QUERY_DEMAND_FOR_PARTS_FROM_ORDERS = @"DemandForPartsFromOrders";
        public static string QUERY_DEMAND_FOR_PARTS_FROM_JOBS = @"DemandForPartsFromJobs";
        public static string QUERY_PARTS_ON_HAND = @"PartsOnHand";
        public static string QUERY_SO_BACKLOG_FOR_PART_ANALYSIS = @"SOBacklogForPartAnalysis";
        public static string QUERY_GET_PART = @"GetPart";
        public static string QUERY_PART_COST_HISTORY_IN_DATE_RANGE = @"PartCostHistoryInDateRange";
        public static string QUERY_PART_COST_HISTORY_IN_DATE_RANGE_BY_PART2 = @"PartCostHistoryDtRngByPart";
        public static string QUERY_INVENTORY_ADJUSTMENTS_IN_DATE_RANGE = @"InventoryAdjustmentsInDtRng";
        public static string QUERY_INVENTORY_ADJUSTMENTS_IN_DATE_RANGE_BY_PART = @"InvAdjustmentsInDtRngByPart";
        public static string QUERY_PO_DETAIL_HISTORY = @"PODetailHistory";
        public static string QUERY_PO_DETAIL_HISTORY_BY_PART = @"PODetailHistoryByPart";
        public static string QUERY_PURCHASE_PARTS_CONSUMED = @"PurPartsConsumed";
        public static string QUERY_PURCHASED_PARTS_CONSUMED_BY_PART = @"PurPartsConsumedByPart";

        public static string QUERY_GL_ACCOUNT_CATEGORIES = @"GLAccountCategories";
        public static string QUERY_GL_SEGMENTS = @"GLSegments";
        public static string QUERY_GL_ACCOUNTS = @"GLAccounts";
        public static string QUERY_GL_CONTROLS = @"GLControls";

        public static string QUERY_PART_INVENTORY_AGING = @"PartInventoryAging";
        public static string QUERY_PART_RECEIPTS_IN_DATE_RANGE = @"PartReceiptsInDateRange";
        public static string QUERY_SIMPLE_WHERE_USED = @"SimpleWhereUsed";

        public static string QUERY_RECEIPTS_IN_TIME = @"ReceiptsInTime";
        //public static string QUERY_ALL_SERIAL_NUMBERS_IN_INVENTORY = @"AllSerialNumbersInInventory";
        //public static string QUERY_GET_MFG_STK_TRANS_IN_DATE_RANGE = @"GetMfgStkTransactionsInDR";
    }
}
