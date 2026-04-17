using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Text;
using System.Linq;

using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Spreadsheet;
using Ice.BO;
using Ice.Core;
using Ice.Lib.Framework;
using Ice.Lib.PerformanceCanvasXmla;
using DocumentFormat.OpenXml.Drawing.Charts;
using SpreadsheetLight;
using SpreadsheetLight.Charts;
using System.Xml;
using System.Xml.Schema;
using System.Runtime.CompilerServices;

using HSPersistence;
using System.Windows.Forms;
using Ice.Tablesets;
using Epicor.Utilities;
using System.Runtime.InteropServices.WindowsRuntime;
using System.Text.RegularExpressions;

namespace HorizonScientific
{
    public class HSAnalyzePartShortages
    {
        #region Constructors

        public HSAnalyzePartShortages()
        {
        }

        #endregion

        #region Methods
        public bool Initialize(Session oSession, string sCompany, DateTime dtCutOffDate)
        {
            bool bSuccess = true;
            m_sCompany = sCompany;
            //////////////////////////////////////////////////
            /// WE NEED TO ANALYZE ALL OUTSTANDING DEMAND FOR RAW PARTS
            /// AND DETERMINE WHICH JOBS WITHIN THE SPECIFIED TIME
            /// WINDOW WILL NOT BE ABLE TO BE STARTED DUE TO A LACK OF PARTS
            /// 
            /// 
            /// Sales orders that are not yet confirmed would be placed at
            /// least one year into the future until the scheduler can 
            /// determine when the job will have parts and can be worked.
            /// 
            /// The scheduler is moving jobs around using the scheudling tool and 
            /// trying to determine if the jobs could be started and completed by
            /// considering the availability of parts and the work load. In this case
            /// the scheduler will find a spot where the job can be fit based on
            /// labor availabiltiy and then will use this tool to answer the question
            /// as to whether there are parts available to complete the job.
            /// 
            /// This method is invoked by the main program in order to load the 
            /// data needed to perform that analysis.
            /// 
            /// This program will consider demand from all jobs(firm or unfirm) within the specified date range.
            /// This program will consider demand from all sales orders within the specified date range.
            /// This program will assume that all outstanding POs will be delivered per the scheduled promise date.
            /// The program allocates parts to jobs and sales orders based on the earliest start date and works its way forward in time.
            /// The output of the program displays all jobs and sales orders that have part shortages based on this analysis.
            /// It will list the job number, the part number and quantity that it is short by.
            /// It will also list the sales order, part number and quantity that it is short by.
            //////////////////////////////////////////////////

            //****Set a parameter Value***** 
            Ice.Proxy.BO.DynamicQueryImpl oDynamicQuery = WCFServiceSupport.CreateImpl<Ice.Proxy.BO.DynamicQueryImpl>(oSession, Ice.Proxy.BO.DynamicQueryImpl.UriPath);
            #region Production Calendar
            QueryExecutionDataSet oQueryExecutionDataSet = oDynamicQuery.GetQueryExecutionParametersByID(BAQConstants.QUERY_PRODUCTION_CALENDAR);
            DateTime dtToday = DateTime.Today;
            DateTime dtFirstProductionDay = new DateTime(DateTime.Now.Year - 1, 1, 1);
            DateTime dtLastProductionDay = new DateTime(DateTime.Now.Year + 1, 12, 31);
            string sCalendarID = "D5H8";
            foreach (DataRow oParameter in oQueryExecutionDataSet.ExecutionParameter)
            {
                if (string.Compare(oParameter["ParameterID"].ToString(), "StartDate") == 0)
                {
                    oParameter["ParameterValue"] = dtFirstProductionDay;
                }
                if (string.Compare(oParameter["ParameterID"].ToString(), "EndDate") == 0)
                {
                    oParameter["ParameterValue"] = dtLastProductionDay;
                }
                if (string.Compare(oParameter["ParameterID"].ToString(), "CalendarID") == 0)
                {
                    oParameter["ParameterValue"] = sCalendarID;
                }
            }
            oQueryExecutionDataSet.AcceptChanges();
            DataSet oDataSet = oDynamicQuery.ExecuteByID(BAQConstants.QUERY_PRODUCTION_CALENDAR, oQueryExecutionDataSet);
            m_oProductionCalendarDays.Clear();
            foreach (DataRow oDataRow in oDataSet.Tables[0].Rows)
            {
                m_oProductionCalendarDays.Add(new ProductionCalendar(oDataRow));
            }
            #endregion

            #region Open Jobs
            oQueryExecutionDataSet = oDynamicQuery.GetQueryExecutionParametersByID(BAQConstants.QUERY_OPEN_JOBS);
            oQueryExecutionDataSet.ExecutionParameter.Clear();
            oDataSet = oDynamicQuery.ExecuteByID(BAQConstants.QUERY_OPEN_JOBS, oQueryExecutionDataSet);
            m_oAllOpenJobs.Clear();
            m_oSupplyJobs.Clear();
            // see what filters we have set for the job supply
            m_bJobSupplyEngineered = false;
            m_bJobSupplyFirmed = false;
            m_bJobSupplyReleased = false;
            foreach (DataRow oDataRow in oDataSet.Tables[0].Rows)
            {
                HSOpenJob oJob = new HSOpenJob(oDataRow);
                // this job is always included in the list of open jobs
                m_oAllOpenJobs.Add(oJob);

                bool bIncludeJob = true;

                // Do we need to filter out supply jobs that are not engineered, firmed, or released?
                if (m_bJobSupplyEngineered == true)
                {
                    if (oJob.Engineered == false)
                    {
                        bIncludeJob = false;
                    }
                }
                if (m_bJobSupplyFirmed == true)
                {
                    if ((oJob.Engineered == false) || (oJob.Firm == false))
                    {
                        bIncludeJob = false;
                    }
                }
                if (m_bJobSupplyReleased == true)
                {
                    if ((oJob.Released == false) || (oJob.Engineered == false) || (oJob.Firm == false))
                    {
                        bIncludeJob = false;
                    }
                }
                if (dtCutOffDate != DateTime.MinValue)
                {
                    // see if job is outside the window
                    if (oJob.DueDate > dtCutOffDate)
                    {
                        bIncludeJob = false;
                    }
                }
                if (bIncludeJob == true)
                {
                    m_oSupplyJobs.Add(oJob);
                }
            }
            // sort these by when the job will be completed
            m_oAllOpenJobs = m_oAllOpenJobs.OrderBy(oItem => oItem.DueDate).ToList();
            #endregion

            #region Part Info
            oQueryExecutionDataSet = oDynamicQuery.GetQueryExecutionParametersByID(BAQConstants.QUERY_LIST_ALL_PARTS);
            oQueryExecutionDataSet.ExecutionParameter.Clear();
            oDataSet = oDynamicQuery.ExecuteByID(BAQConstants.QUERY_LIST_ALL_PARTS, oQueryExecutionDataSet);
            m_oAllPartData.Clear();
            foreach (DataRow oDataRow in oDataSet.Tables[0].Rows)
            {
                m_oAllPartData.Add(new HSPartData(oDataRow));
            }
            #endregion

            #region Purchase Orders
            oQueryExecutionDataSet = oDynamicQuery.GetQueryExecutionParametersByID(BAQConstants.QUERY_SOURCE_FOR_PARTS_FROM_PO);
            oQueryExecutionDataSet.ExecutionParameter.Clear();
            oDataSet = oDynamicQuery.ExecuteByID(BAQConstants.QUERY_SOURCE_FOR_PARTS_FROM_PO, oQueryExecutionDataSet);
            m_oPartsSuppliedByPOs.Clear();
            // see what filters we have set for the po supply
            m_bPOSupplyConfirmed = false;
            m_bPOSupplyApproved = false;
            foreach (DataRow oDataRow in oDataSet.Tables[0].Rows)
            {
                HSSourceForPartFromPO oPO = new HSSourceForPartFromPO(oDataRow);
                bool bIncludePO = true;
                // do we need to filter out ones that are not approved and confirmed?
                if (m_bPOSupplyConfirmed == true)
                {
                    if (oPO.Confirmed == false)
                    {
                        bIncludePO = false;
                    }
                }
                if (m_bPOSupplyApproved == true)
                {
                    if ((oPO.Confirmed == false) || (oPO.Approved == false))
                    {
                        bIncludePO = false;
                    }
                }
                if (bIncludePO == true)
                {
                    m_oPartsSuppliedByPOs.Add(oPO);
                }
            }
            // sort by PO Expected Arrival Date
            m_oPartsSuppliedByPOs = m_oPartsSuppliedByPOs.OrderBy(oItem => oItem.DueDate).ToList();
            #endregion

            #region Requisitions
            oQueryExecutionDataSet = oDynamicQuery.GetQueryExecutionParametersByID(BAQConstants.QUERY_PART_DEMAND_FROM_REQUISITIONS);
            oQueryExecutionDataSet.ExecutionParameter.Clear();
            oDataSet = oDynamicQuery.ExecuteByID(BAQConstants.QUERY_PART_DEMAND_FROM_REQUISITIONS, oQueryExecutionDataSet);
            m_oPartDemandFromRequisitions.Clear();
            foreach (DataRow oRow in oDataSet.Tables[0].Rows)
            {
                HSDemandForPartsFromRequisition oReq = new HSDemandForPartsFromRequisition(oRow);
                bool bIncludeReq = true;
                // check to see if req is outside of window
                if (dtCutOffDate != DateTime.MinValue)
                {
                    if (oReq.DueDate > dtCutOffDate)
                    {
                        bIncludeReq = false;
                    }
                }
                if (bIncludeReq == true)
                {
                    m_oPartDemandFromRequisitions.Add(new HSDemandForPartsFromRequisition(oRow));
                }
            }
            // sort by requisition due date
            m_oPartDemandFromRequisitions = m_oPartDemandFromRequisitions.OrderBy(oItem => oItem.DueDate).ToList();
            #endregion

            #region Demand From Orders
            oQueryExecutionDataSet = oDynamicQuery.GetQueryExecutionParametersByID(BAQConstants.QUERY_DEMAND_FOR_PARTS_FROM_ORDERS);
            oQueryExecutionDataSet.ExecutionParameter.Clear();
            oDataSet = oDynamicQuery.ExecuteByID(BAQConstants.QUERY_DEMAND_FOR_PARTS_FROM_ORDERS, oQueryExecutionDataSet);
            m_oPartDemandFromOrders.Clear();
            // see what filters we have set for the sales order demand
            List<HSDemandForPartFromOrder> oSalesOrdersExcludedFromDemand = new List<HSDemandForPartFromOrder>();
            m_bSalesOrderDemandFirmed = false;
            foreach (DataRow oRow in oDataSet.Tables[0].Rows)
            {
                HSDemandForPartFromOrder oSalesOrder = new HSDemandForPartFromOrder(oRow);
                bool bIncludeSalesOrder = true;
                if (m_bSalesOrderDemandFirmed == true)
                {
                    if (oSalesOrder.Firm == false)
                    {
                        bIncludeSalesOrder = false;
                    }
                }
                // check to see if this is outside of the window
                if (dtCutOffDate != DateTime.MinValue)
                {
                    if (oSalesOrder.ReleaseRequiredDate > dtCutOffDate)
                    {
                        bIncludeSalesOrder = false;
                    }
                }
                if (bIncludeSalesOrder == true)
                {
                    m_oPartDemandFromOrders.Add(oSalesOrder);
                }
                else
                {
                    // we need to track these sales order so we can pull the demand out 
                    // from the part demand records we already read in from the database
                    oSalesOrdersExcludedFromDemand.Add(oSalesOrder);
                }
            }
            // sort by release ship by date
            m_oPartDemandFromOrders = m_oPartDemandFromOrders.OrderBy(oItem => oItem.ReleaseRequiredDate).ToList();
            #endregion

            #region Demand From Jobs
            oQueryExecutionDataSet = oDynamicQuery.GetQueryExecutionParametersByID(BAQConstants.QUERY_DEMAND_FOR_PARTS_FROM_JOBS);
            oQueryExecutionDataSet.ExecutionParameter.Clear();
            oDataSet = oDynamicQuery.ExecuteByID(BAQConstants.QUERY_DEMAND_FOR_PARTS_FROM_JOBS, oQueryExecutionDataSet);
            m_oPartDemandFromJobs.Clear();
            // see what filters we have set for the job demand
            List<HSDemandForPartFromJob> oJobsExcludedFromDemand = new List<HSDemandForPartFromJob>();
            m_bJobDemandEngineered = false;
            m_bJobDemandFirmed = false;
            m_bJobDemandReleased = false;
            foreach (DataRow oRow in oDataSet.Tables[0].Rows)
            {
                HSDemandForPartFromJob oJobDemand = new HSDemandForPartFromJob(oRow);
                bool bIncludeJobDemand = true;
                if (m_bJobDemandEngineered == true)
                {
                    if (oJobDemand.Engineered == false)
                    {
                        bIncludeJobDemand = false;
                    }
                }
                if (m_bJobDemandFirmed == true)
                {
                    if ((oJobDemand.Engineered == false) || (oJobDemand.Firm == false))
                    {
                        bIncludeJobDemand = false;
                    }
                }
                if (m_bJobDemandReleased == true)
                {
                    if ((oJobDemand.Released == false) || (oJobDemand.Engineered == false) || (oJobDemand.Firm == false))
                    {
                        bIncludeJobDemand = false;
                    }
                }
                // see if we are outside of the window
                if (dtCutOffDate != DateTime.MinValue)
                {
                    if (oJobDemand.JobRequiredDate > dtCutOffDate)
                    {
                        bIncludeJobDemand = false;
                    }
                }
                if (bIncludeJobDemand == true)
                {
                    m_oPartDemandFromJobs.Add(oJobDemand);
                }
                else
                {
                    // we need to track these jobs so we can pull the demand out 
                    // from the part demand records we already read in from the database
                    oJobsExcludedFromDemand.Add(oJobDemand);
                }
            }
            // sort by job material required by date
            m_oPartDemandFromJobs = m_oPartDemandFromJobs.OrderBy(oItem => oItem.JobRequiredDate).ToList();
            #endregion

            #region Part Demand In Time
            oQueryExecutionDataSet = oDynamicQuery.GetQueryExecutionParametersByID(BAQConstants.QUERY_DEMAND_FOR_PARTS_IN_TIME);
            oQueryExecutionDataSet.ExecutionParameter.Clear();
            oDataSet = oDynamicQuery.ExecuteByID(BAQConstants.QUERY_DEMAND_FOR_PARTS_IN_TIME, oQueryExecutionDataSet);
            m_oPartDemandInTime.Clear();
            foreach (DataRow oRow in oDataSet.Tables[0].Rows)
            {
                HSDemandForPartsInTime oPartDemand = new HSDemandForPartsInTime(oRow);
                bool bIncludePartDemand = true;
                // need to check to make sure the demand is not comming from a job that is to be excluded
                HSDemandForPartFromJob oExcludedJob = oJobsExcludedFromDemand.FirstOrDefault(oItem => string.Compare(oItem.JobNumber, oPartDemand.JobNum, true) == 0);
                if (oExcludedJob != null)
                {
                    // this part demand record is tied to a job that we do not want to consider so we do not add it to the demand list
                    bIncludePartDemand = false;
                }
                // need to check to make sure the demand is not comming from a sales order that is to be excluded
                HSDemandForPartFromOrder oExcludedSalesOrder = oSalesOrdersExcludedFromDemand.FirstOrDefault(oItem => (oItem.OrderNumber == oPartDemand.OrderNum) && (oItem.LineNumber == oPartDemand.OrderLine) &&
                    (oItem.ReleaseNumber == oPartDemand.OrderRelease));
                if (oExcludedSalesOrder != null)
                {
                    // this part demand record is tied to a sales order that we do not want to consider so we do not add it to the demand list
                    bIncludePartDemand = false;
                }
                // check to see if we are outside of window
                if (dtCutOffDate != DateTime.MinValue)
                {
                    if (oPartDemand.DueDate > dtCutOffDate)
                    {
                        bIncludePartDemand = false;
                    }
                }
                if (bIncludePartDemand == true)
                {
                    m_oPartDemandInTime.Add(oPartDemand);
                }
            }
            // sort by due date -- this is the time that the change will go into effect (either issued to job, po received, etc)
            m_oPartDemandInTime = m_oPartDemandInTime.OrderBy(oItem => oItem.DueDate).ToList();
            #endregion

            #region Parts On Hand
            // get the parts we have on hand
            oQueryExecutionDataSet = oDynamicQuery.GetQueryExecutionParametersByID(BAQConstants.QUERY_PARTS_ON_HAND);
            oQueryExecutionDataSet.ExecutionParameter.Clear();
            oDataSet = oDynamicQuery.ExecuteByID(BAQConstants.QUERY_PARTS_ON_HAND, oQueryExecutionDataSet);
            m_oPartsOnHand.Clear();
            foreach (DataRow oRow in oDataSet.Tables[0].Rows)
            {
                m_oPartsOnHand.Add(new HSPartsOnHand(oRow));
            }
            #endregion

            #region Perform Analysis
            FillInTimeDomain();

            EvaluatePartShortagesForJobs();

            EvaluatePartShortagesForSalesOrders();
            #endregion

            return bSuccess;
        }

        public bool Initialize2(Session oSession, string sCompany)
        {
            m_sCompany = sCompany;

            //****Set a parameter Value***** 
            Ice.Proxy.BO.DynamicQueryImpl oDynamicQuery = WCFServiceSupport.CreateImpl<Ice.Proxy.BO.DynamicQueryImpl>(oSession, Ice.Proxy.BO.DynamicQueryImpl.UriPath);

            #region Backlog Details
            // get the sales order backlog detail
            QueryExecutionDataSet oQueryExecutionDataSet = oDynamicQuery.GetQueryExecutionParametersByID(BAQConstants.QUERY_SO_BACKLOG_FOR_PART_ANALYSIS);
            oQueryExecutionDataSet.ExecutionParameter.Clear();
            DataSet oDataSet = oDynamicQuery.ExecuteByID(BAQConstants.QUERY_SO_BACKLOG_FOR_PART_ANALYSIS, oQueryExecutionDataSet);
            // initialize the backlog details list
            m_oAllBacklogDetails.Clear();
            foreach (DataRow oDataRow in oDataSet.Tables[0].Rows)
            {
                m_oAllBacklogDetails.Add(new SOBacklogForPartAnalysis(oDataRow));
            }
            #endregion

            #region Group Orders As Required
            // now lets compute the effective ship date, and whether this is a customer service order for all of these sales orders
            foreach (SOBacklogForPartAnalysis oSalesOrder in m_oAllBacklogDetails)
            {
                DateTime dtEffectiveShipDate = oSalesOrder.RequiredByDate;
               
                // finally set the effective ship date
                oSalesOrder.EffectiveShipDate = dtEffectiveShipDate;
            }
            // by default we order these by the effective ship date
            m_oAllBacklogDetails = m_oAllBacklogDetails.OrderBy(oItem => oItem.EffectiveShipDate).ToList();

            // let group these orders together based whether the order must ship complete or if the line must ship complete
            m_oOrdersToAnalyze.Clear();
            foreach (SOBacklogForPartAnalysis oTmp in m_oAllBacklogDetails)
            {
                List<SOBacklogForPartAnalysis> oAllItemsInGroup = new List<SOBacklogForPartAnalysis>();
                if (oTmp.Processed == false)
                {
                    if (oTmp.ShipOrderComplete == true)
                    {
                        // the order is set to ship complete so we put all orders with this order num in this group
                        oAllItemsInGroup = m_oAllBacklogDetails.Where(oItem => oItem.OrderNum == oTmp.OrderNum).ToList();
                    }
                    else if (oTmp.TypeCode == "K")
                    {
                        // if this is a sales kit then we should pull in all lines tied to this parent line
                        oAllItemsInGroup = m_oAllBacklogDetails.Where(oItem => (oItem.OrderNum == oTmp.OrderNum) && (oItem.KitParentLine == oTmp.OrderLine)).ToList();
                    }
                    else if (oTmp.ShipLineComplete == true)
                    {
                        // the order line is set to ship complete so we put all orders with the order num and line num in the group
                        oAllItemsInGroup = m_oAllBacklogDetails.Where(oItem => (oItem.OrderNum == oTmp.OrderNum) && (oItem.OrderLine == oTmp.OrderLine)).ToList();
                    }
                    else
                    {
                        // we just have this one order in the group
                        oAllItemsInGroup.Add(oTmp);
                    }

                    // now we set all orders in the group to processed so we dont reprocesses something
                    foreach (SOBacklogForPartAnalysis oSO in oAllItemsInGroup)
                    {
                        oSO.Processed = true;
                    }
                    m_oOrdersToAnalyze.Add(new SalesOrderGroup(oAllItemsInGroup));
                }
            }

            // now that we have the sales order in the proper groupings we will reset the processed flag
            // as this will be needed to determine if there are enough parts to satisfy the order
            foreach (SOBacklogForPartAnalysis oTmp in m_oAllBacklogDetails)
            {
                oTmp.Processed = false;
            }

            // prioritize the groups of orders based on their total value
            m_oOrdersToAnalyze = m_oOrdersToAnalyze.OrderBy(oItem => oItem.TotalValue).ToList();
            #endregion

            #region Production Calendar
            oQueryExecutionDataSet = oDynamicQuery.GetQueryExecutionParametersByID(BAQConstants.QUERY_PRODUCTION_CALENDAR);
            DateTime dtToday = DateTime.Today;
            DateTime dtFirstProductionDay = new DateTime(DateTime.Now.Year - 1, 1, 1);
            DateTime dtLastProductionDay = new DateTime(DateTime.Now.Year + 1, 12, 31);
            string sCalendarID = "D5H8";
            foreach (DataRow oParameter in oQueryExecutionDataSet.ExecutionParameter)
            {
                if (string.Compare(oParameter["ParameterID"].ToString(), "StartDate") == 0)
                {
                    oParameter["ParameterValue"] = dtFirstProductionDay;
                }
                if (string.Compare(oParameter["ParameterID"].ToString(), "EndDate") == 0)
                {
                    oParameter["ParameterValue"] = dtLastProductionDay;
                }
                if (string.Compare(oParameter["ParameterID"].ToString(), "CalendarID") == 0)
                {
                    oParameter["ParameterValue"] = sCalendarID;
                }
            }
            oQueryExecutionDataSet.AcceptChanges();
            oDataSet = oDynamicQuery.ExecuteByID(BAQConstants.QUERY_PRODUCTION_CALENDAR, oQueryExecutionDataSet);
            m_oProductionCalendarDays.Clear();
            foreach (DataRow oDataRow in oDataSet.Tables[0].Rows)
            {
                m_oProductionCalendarDays.Add(new ProductionCalendar(oDataRow));
            }
            #endregion

            #region Open Jobs
            oQueryExecutionDataSet = oDynamicQuery.GetQueryExecutionParametersByID(BAQConstants.QUERY_OPEN_JOBS);
            oQueryExecutionDataSet.ExecutionParameter.Clear();
            oDataSet = oDynamicQuery.ExecuteByID(BAQConstants.QUERY_OPEN_JOBS, oQueryExecutionDataSet);
            m_oAllOpenJobs.Clear();
            m_oSupplyJobs.Clear();
            // see what filters we have set for the job supply
            m_bJobSupplyEngineered = false;
            m_bJobSupplyFirmed = false;
            m_bJobSupplyReleased = false;
            foreach (DataRow oDataRow in oDataSet.Tables[0].Rows)
            {
                HSOpenJob oJob = new HSOpenJob(oDataRow);
                // this job is always included in the list of open jobs
                m_oAllOpenJobs.Add(oJob);

                bool bIncludeJob = false;

                // See if this Job is related to an order that needs to be analyzed in the list
                if (string.Compare(oJob.CustomerName, STOCKING_JOB, true) == 0)
                {
                    // this is a stocking job so we will include it
                    bIncludeJob = true;
                }
                else
                {
                    // this is a make direct job so we need to find the related sales order 
                    foreach (SalesOrderGroup oGroup in m_oOrdersToAnalyze)
                    {
                        foreach (SOBacklogForPartAnalysis oSOBacklog in oGroup.OrdersThatMustShipTogether)
                        {
                            if (string.Compare(oSOBacklog.JobNum, oJob.JobNum, true) == 0)
                            {
                                bIncludeJob = true;
                                break;
                            }
                        }

                        if (bIncludeJob == true)
                        {
                            break;
                        }
                    }
                }

                if (bIncludeJob == true)
                {
                    // Do we need to filter out supply jobs that are not engineered, firmed, or released?
                    if (m_bJobSupplyEngineered == true)
                    {
                        if (oJob.Engineered == false)
                        {
                            bIncludeJob = false;
                        }
                    }
                    if (m_bJobSupplyFirmed == true)
                    {
                        if ((oJob.Engineered == false) || (oJob.Firm == false))
                        {
                            bIncludeJob = false;
                        }
                    }
                    if (m_bJobSupplyReleased == true)
                    {
                        if ((oJob.Released == false) || (oJob.Engineered == false) || (oJob.Firm == false))
                        {
                            bIncludeJob = false;
                        }
                    }
                }

                // finally see if we should still include this job supply
                if (bIncludeJob == true)
                {
                    m_oSupplyJobs.Add(oJob);
                }
            }
            // sort these by when the job will be completed
            m_oAllOpenJobs = m_oAllOpenJobs.OrderBy(oItem => oItem.DueDate).ToList();
            #endregion

            #region Part Info
            oQueryExecutionDataSet = oDynamicQuery.GetQueryExecutionParametersByID(BAQConstants.QUERY_LIST_ALL_PARTS);
            oQueryExecutionDataSet.ExecutionParameter.Clear();
            oDataSet = oDynamicQuery.ExecuteByID(BAQConstants.QUERY_LIST_ALL_PARTS, oQueryExecutionDataSet);
            m_oAllPartData.Clear();
            foreach (DataRow oDataRow in oDataSet.Tables[0].Rows)
            {
                m_oAllPartData.Add(new HSPartData(oDataRow));
            }
            #endregion

            #region Purchase Orders
            oQueryExecutionDataSet = oDynamicQuery.GetQueryExecutionParametersByID(BAQConstants.QUERY_SOURCE_FOR_PARTS_FROM_PO);
            oQueryExecutionDataSet.ExecutionParameter.Clear();
            oDataSet = oDynamicQuery.ExecuteByID(BAQConstants.QUERY_SOURCE_FOR_PARTS_FROM_PO, oQueryExecutionDataSet);
            m_oPartsSuppliedByPOs.Clear();
            // see what filters we have set for the po supply
            m_bPOSupplyConfirmed = false;
            m_bPOSupplyApproved = false;
            foreach (DataRow oDataRow in oDataSet.Tables[0].Rows)
            {
                HSSourceForPartFromPO oPO = new HSSourceForPartFromPO(oDataRow);
                bool bIncludePO = true;
                // do we need to filter out ones that are not approved and confirmed?
                if (m_bPOSupplyConfirmed == true)
                {
                    if (oPO.Confirmed == false)
                    {
                        bIncludePO = false;
                    }
                }
                if (m_bPOSupplyApproved == true)
                {
                    if ((oPO.Confirmed == false) || (oPO.Approved == false))
                    {
                        bIncludePO = false;
                    }
                }
                if (bIncludePO == true)
                {
                    m_oPartsSuppliedByPOs.Add(oPO);
                }
            }
            // sort by PO Expected Arrival Date
            m_oPartsSuppliedByPOs = m_oPartsSuppliedByPOs.OrderBy(oItem => oItem.DueDate).ToList();
            #endregion

            #region Requisitions
            oQueryExecutionDataSet = oDynamicQuery.GetQueryExecutionParametersByID(BAQConstants.QUERY_PART_DEMAND_FROM_REQUISITIONS);
            oQueryExecutionDataSet.ExecutionParameter.Clear();
            oDataSet = oDynamicQuery.ExecuteByID(BAQConstants.QUERY_PART_DEMAND_FROM_REQUISITIONS, oQueryExecutionDataSet);
            m_oPartDemandFromRequisitions.Clear();
            foreach (DataRow oRow in oDataSet.Tables[0].Rows)
            {
                HSDemandForPartsFromRequisition oReq = new HSDemandForPartsFromRequisition(oRow);
                m_oPartDemandFromRequisitions.Add(new HSDemandForPartsFromRequisition(oRow));
            }
            // sort by requisition due date
            m_oPartDemandFromRequisitions = m_oPartDemandFromRequisitions.OrderBy(oItem => oItem.DueDate).ToList();
            #endregion

            #region Demand From Orders
            oQueryExecutionDataSet = oDynamicQuery.GetQueryExecutionParametersByID(BAQConstants.QUERY_DEMAND_FOR_PARTS_FROM_ORDERS);
            oQueryExecutionDataSet.ExecutionParameter.Clear();
            oDataSet = oDynamicQuery.ExecuteByID(BAQConstants.QUERY_DEMAND_FOR_PARTS_FROM_ORDERS, oQueryExecutionDataSet);
            m_oPartDemandFromOrders.Clear();
            // see what filters we have set for the sales order demand
            List<HSDemandForPartFromOrder> oSalesOrdersExcludedFromDemand = new List<HSDemandForPartFromOrder>();
            m_bSalesOrderDemandFirmed = false;
            foreach (DataRow oRow in oDataSet.Tables[0].Rows)
            {
                HSDemandForPartFromOrder oSalesOrder = new HSDemandForPartFromOrder(oRow);

                bool bIncludeSalesOrder = false;
                // check to see if we are eliminating this sales order from the analysis
                foreach (SalesOrderGroup oGroup in m_oOrdersToAnalyze)
                {
                    foreach (SOBacklogForPartAnalysis oSOBacklog in oGroup.OrdersThatMustShipTogether)
                    {
                        if ((oSOBacklog.OrderNum == oSalesOrder.OrderNumber) && (oSOBacklog.OrderLine == oSalesOrder.LineNumber) && (oSOBacklog.OrderRelease == oSalesOrder.ReleaseNumber))
                        {
                            bIncludeSalesOrder = true;
                            break;
                        }
                    }

                    if (bIncludeSalesOrder == true)
                    {
                        break;
                    }
                }

                if (bIncludeSalesOrder == true)
                {
                    if (m_bSalesOrderDemandFirmed == true)
                    {
                        if (oSalesOrder.Firm == false)
                        {
                            bIncludeSalesOrder = false;
                        }
                    }
                }

                if (bIncludeSalesOrder == true)
                {
                    m_oPartDemandFromOrders.Add(oSalesOrder);
                }
                else
                {
                    // we need to track these sales order so we can pull the demand out 
                    // from the part demand records we already read in from the database
                    oSalesOrdersExcludedFromDemand.Add(oSalesOrder);
                }
            }
            // sort by release ship by date
            m_oPartDemandFromOrders = m_oPartDemandFromOrders.OrderBy(oItem => oItem.ReleaseRequiredDate).ToList();
            #endregion

            #region Demand From Jobs
            oQueryExecutionDataSet = oDynamicQuery.GetQueryExecutionParametersByID(BAQConstants.QUERY_DEMAND_FOR_PARTS_FROM_JOBS);
            oQueryExecutionDataSet.ExecutionParameter.Clear();
            oDataSet = oDynamicQuery.ExecuteByID(BAQConstants.QUERY_DEMAND_FOR_PARTS_FROM_JOBS, oQueryExecutionDataSet);
            m_oPartDemandFromJobs.Clear();
            // see what filters we have set for the job demand
            List<HSDemandForPartFromJob> oJobsExcludedFromDemand = new List<HSDemandForPartFromJob>();
            m_bJobDemandEngineered = false;
            m_bJobDemandFirmed = false;
            m_bJobDemandReleased = false;
            foreach (DataRow oRow in oDataSet.Tables[0].Rows)
            {
                HSDemandForPartFromJob oJobDemand = new HSDemandForPartFromJob(oRow);
                bool bIncludeJobDemand = false;

                // See if this Job is related to an order that needs to be analyzed in the list
                if ((oJobDemand.OrderNum == 0) && (oJobDemand.OrderLine == 0) && (oJobDemand.OrderRelNum == 0))
                {
                    // this is a stocking job so we will include it
                    bIncludeJobDemand = true;
                }
                else
                {
                    // this is a make direct job so we need to find the related sales order 
                    foreach (SalesOrderGroup oGroup in m_oOrdersToAnalyze)
                    {
                        foreach (SOBacklogForPartAnalysis oSOBacklog in oGroup.OrdersThatMustShipTogether)
                        {
                            if ((oSOBacklog.OrderNum == oJobDemand.OrderNum) && (oSOBacklog.OrderLine == oJobDemand.OrderLine) && (oSOBacklog.OrderRelease == oJobDemand.OrderRelNum))
                            {
                                bIncludeJobDemand = true;
                                break;
                            }
                        }

                        if (bIncludeJobDemand == true)
                        {
                            break;
                        }
                    }
                }

                if (bIncludeJobDemand == true)
                {
                    if (m_bJobDemandEngineered == true)
                    {
                        if (oJobDemand.Engineered == false)
                        {
                            bIncludeJobDemand = false;
                        }
                    }
                    if (m_bJobDemandFirmed == true)
                    {
                        if ((oJobDemand.Engineered == false) || (oJobDemand.Firm == false))
                        {
                            bIncludeJobDemand = false;
                        }
                    }
                    if (m_bJobDemandReleased == true)
                    {
                        if ((oJobDemand.Released == false) || (oJobDemand.Engineered == false) || (oJobDemand.Firm == false))
                        {
                            bIncludeJobDemand = false;
                        }
                    }
                }

                if (bIncludeJobDemand == true)
                {
                    m_oPartDemandFromJobs.Add(oJobDemand);
                }
                else
                {
                    // we need to track these jobs so we can pull the demand out 
                    // from the part demand records we already read in from the database
                    oJobsExcludedFromDemand.Add(oJobDemand);
                }
            }
            // sort by job material required by date
            m_oPartDemandFromJobs = m_oPartDemandFromJobs.OrderBy(oItem => oItem.JobRequiredDate).ToList();
            #endregion

            #region Part Demand In Time
            oQueryExecutionDataSet = oDynamicQuery.GetQueryExecutionParametersByID(BAQConstants.QUERY_DEMAND_FOR_PARTS_IN_TIME);
            oQueryExecutionDataSet.ExecutionParameter.Clear();
            oDataSet = oDynamicQuery.ExecuteByID(BAQConstants.QUERY_DEMAND_FOR_PARTS_IN_TIME, oQueryExecutionDataSet);
            m_oPartDemandInTime.Clear();
            foreach (DataRow oRow in oDataSet.Tables[0].Rows)
            {
                HSDemandForPartsInTime oPartDemand = new HSDemandForPartsInTime(oRow);

                bool bIncludePartDemand = true;
                // need to check to make sure the demand is not comming from a job that is to be excluded
                HSDemandForPartFromJob oExcludedJob = oJobsExcludedFromDemand.FirstOrDefault(oItem => string.Compare(oItem.JobNumber, oPartDemand.JobNum, true) == 0);
                if (oExcludedJob != null)
                {
                    // this part demand record is tied to a job that we do not want to consider so we do not add it to the demand list
                    bIncludePartDemand = false;
                }
                // need to check to make sure the demand is not comming from a sales order that is to be excluded
                HSDemandForPartFromOrder oExcludedSalesOrder = oSalesOrdersExcludedFromDemand.FirstOrDefault(oItem => (oItem.OrderNumber == oPartDemand.OrderNum) && (oItem.LineNumber == oPartDemand.OrderLine) &&
                    (oItem.ReleaseNumber == oPartDemand.OrderRelease));
                if (oExcludedSalesOrder != null)
                {
                    // this part demand record is tied to a sales order that we do not want to consider so we do not add it to the demand list
                    bIncludePartDemand = false;
                }
                if (bIncludePartDemand == true)
                {
                    m_oPartDemandInTime.Add(oPartDemand);
                }
            }
            // sort by due date -- this is the time that the change will go into effect (either issued to job, po received, etc)
            m_oPartDemandInTime = m_oPartDemandInTime.OrderBy(oItem => oItem.DueDate).ToList();
            #endregion

            #region Parts On Hand
            // get the parts we have on hand
            oQueryExecutionDataSet = oDynamicQuery.GetQueryExecutionParametersByID(BAQConstants.QUERY_PARTS_ON_HAND);
            oQueryExecutionDataSet.ExecutionParameter.Clear();
            oDataSet = oDynamicQuery.ExecuteByID(BAQConstants.QUERY_PARTS_ON_HAND, oQueryExecutionDataSet);
            m_oPartsOnHand.Clear();
            foreach (DataRow oRow in oDataSet.Tables[0].Rows)
            {
                m_oPartsOnHand.Add(new HSPartsOnHand(oRow));
            }
            #endregion

            #region Perform Analysis
            FillInTimeDomain();

            EvaluatePartShortagesForJobs();

            EvaluatePartShortagesForSalesOrders();

            EvaluateJobDependencies();

            EvaluateSalesOrderDependencyOnJobs();

            #endregion

            m_oOrdersToAnalyze = m_oOrdersToAnalyze.OrderBy(oItem => oItem.EffectiveShipDate).ToList();

            // when processing these orders we need to indicate where the part shortage exists
            // they cannot be fulfilled as well
            // 1) Insufficient Inventory (partnum qty short)
            // 2) order must ship complete and one line cant ship (which line)
            // 3) line must ship complete and one release cant ship (which release)
            // 4) customer indicates dont ship until

            // for orders that dont have a part shortage we need to highlight any other issues
            // 1) customer on credit hold
            // 2) order not reviewed
            // 3) order placed on hold


            return true;
        }

        public void FillInTimeDomain()
        {
            #region Fill In Time Domain

            // get all part numbers for all open jobs for which we need to provide an analysis
            m_oAllPartNumbers = m_oPartDemandFromJobs.Select(oItem => oItem.PartNum).ToList();
            // get all part numbers for all open orders for which we need to provide an analysis
            m_oAllPartNumbers.AddRange(m_oPartDemandFromOrders.Select(oItem => oItem.PartNum).ToList());
            //
            // get all part numbers for all open reqs for which we need to provide an analysis
            m_oAllPartNumbers.AddRange(m_oPartDemandFromRequisitions.Select(oItem => oItem.PartNum).ToList());
            // now get the distinct list of part numbers for which we have to do analysis
            m_oAllPartNumbers = m_oAllPartNumbers.Distinct().ToList();
            m_oAllPartTimeDomainInfo = new List<PartTimeDomainInfo>();

            // walk through each part and determine what the demand looks like through time (not including PO suggestions)
            foreach (string sPartNum in m_oAllPartNumbers)
            {
                HSPartData oPartInfo = m_oAllPartData.FirstOrDefault(oItem => string.Compare(sPartNum, oItem.PartNum, true) == 0);
                if (oPartInfo != null)
                {
                    PartTimeDomainInfo oPartTimeDomainInfo = new PartTimeDomainInfo(oPartInfo, m_sCompany);
                    m_oAllPartTimeDomainInfo.Add(oPartTimeDomainInfo);

                    // get all information for this part (jobs for part, orders for part, POs for part, etc)
                    oPartTimeDomainInfo.Jobs = m_oPartDemandFromJobs.Where(oItem => oItem.PartNum == sPartNum).ToList();
                    oPartTimeDomainInfo.Orders = m_oPartDemandFromOrders.Where(oItem => oItem.PartNum == sPartNum).ToList();
                    oPartTimeDomainInfo.Requisitions = m_oPartDemandFromRequisitions.Where(oItem => oItem.PartNum == sPartNum).ToList();
                    oPartTimeDomainInfo.POs = m_oPartsSuppliedByPOs.Where(oItem => oItem.PartNum == sPartNum).ToList();
                    oPartTimeDomainInfo.JobsSupplyingParts = m_oSupplyJobs.Where(oItem => oItem.PartNum == sPartNum).ToList();
                    oPartTimeDomainInfo.PartsOnHand = m_oPartsOnHand.Where(oItem => oItem.PartNum == sPartNum).ToList();

                    // we will now combine all of the above data into a list of events for a single part
                    oPartTimeDomainInfo.PartInfoTimeLineData = new List<HSPartInfoTimeLine>();

                    // get the start and end dates for this timeline
                    List<HSDemandForPartsInTime> oPartAndTimeInfo = m_oPartDemandInTime.Where(oItem => string.Compare(oItem.PartNum, sPartNum, true) == 0).ToList();
                    DateTime dtStartDate = DateTime.MinValue;
                    DateTime dtEndDate = DateTime.MaxValue;
                    if (oPartAndTimeInfo.Count > 0)
                    {
                        dtStartDate = oPartAndTimeInfo.Min(x => x.DueDate);
                        dtEndDate = oPartAndTimeInfo.Max(x => x.DueDate);
                    }
                    // we subtract a day just in case there are no parts on hand in the demand in time list
                    if (dtStartDate != DateTime.MinValue)
                    {
                        dtStartDate = dtStartDate.AddDays(-1);
                    }
                    oPartTimeDomainInfo.StartDate = dtStartDate;
                    oPartTimeDomainInfo.EndDate = dtEndDate;
                    HSPartInfoTimeLine oTmp;

                    // add in each outstanding PO
                    foreach (HSSourceForPartFromPO oOutstandingPO in oPartTimeDomainInfo.POs)
                    {
                        oTmp = new HSPartInfoTimeLine();
                        oTmp.DateOfChange = oOutstandingPO.DueDate;
                        // ensure we have the min start date and max end dates
                        if (oTmp.DateOfChange < oPartTimeDomainInfo.StartDate)
                        {
                            oPartTimeDomainInfo.StartDate = oTmp.DateOfChange;
                        }
                        if (oTmp.DateOfChange > oPartTimeDomainInfo.EndDate)
                        {
                            oPartTimeDomainInfo.EndDate = oTmp.DateOfChange;
                        }
                        oTmp.NetChange = oOutstandingPO.NetDemandQuantity;
                        if (oTmp.NetChange < 0)
                        {
                            oTmp.NetChange = 0;
                        }
                        oTmp.ActionType = "PO";
                        oTmp.ReasonForChange = "PO: " + oOutstandingPO.PONumber.ToString() + "-" + oOutstandingPO.LineNumber.ToString() + "-" + oOutstandingPO.ReleaseNumber.ToString();
                        oTmp.PONum = oOutstandingPO.PONumber;
                        oTmp.POLine = oOutstandingPO.LineNumber;
                        oTmp.PORelease = oOutstandingPO.ReleaseNumber;
                        oTmp.PromiseDate = oOutstandingPO.PromiseDate;
                        oPartTimeDomainInfo.PartInfoTimeLineData.Add(oTmp);
                    }

                    // add in each open job
                    foreach (HSOpenJob oOpenJob in oPartTimeDomainInfo.JobsSupplyingParts)
                    {
                        oTmp = new HSPartInfoTimeLine();
                        oTmp.DateOfChange = oOpenJob.DueDate;
                        // ensure we have the min start date and max end dates
                        if (oTmp.DateOfChange < oPartTimeDomainInfo.StartDate)
                        {
                            oPartTimeDomainInfo.StartDate = oTmp.DateOfChange;
                        }
                        if (oTmp.DateOfChange > oPartTimeDomainInfo.EndDate)
                        {
                            oPartTimeDomainInfo.EndDate = oTmp.DateOfChange;
                        }
                        oTmp.NetChange = oOpenJob.ProductionQty;
                        if (oTmp.NetChange < 0)
                        {
                            oTmp.NetChange = 0;
                        }
                        oTmp.ActionType = "JOB SUPPLY";
                        oTmp.ReasonForChange = "JOB: " + oOpenJob.JobNum.ToString();
                        oTmp.JobNum = oOpenJob.JobNum;
                        oPartTimeDomainInfo.PartInfoTimeLineData.Add(oTmp);
                    }

                    // add in each requisition entry
                    foreach (HSDemandForPartsFromRequisition oRequisition in oPartTimeDomainInfo.Requisitions)
                    {
                        oTmp = new HSPartInfoTimeLine();
                        oTmp.DateOfChange = oRequisition.DueDate;
                        // ensure we have the min start date and max end dates
                        if (oTmp.DateOfChange < oPartTimeDomainInfo.StartDate)
                        {
                            oPartTimeDomainInfo.StartDate = oTmp.DateOfChange;
                        }
                        if (oTmp.DateOfChange > oPartTimeDomainInfo.EndDate)
                        {
                            oPartTimeDomainInfo.EndDate = oTmp.DateOfChange;
                        }
                        oTmp.NetChange = oRequisition.RequiredQuantity;
                        // if required quantity is less than minimum quantity then increase the requested amount
                        if (oTmp.NetChange < oPartTimeDomainInfo.MinimumOrderQuantity)
                        {
                            oTmp.NetChange = oPartTimeDomainInfo.MinimumOrderQuantity;
                        }
                        oTmp.ActionType = "REQUISITION";
                        oTmp.ReasonForChange = "Requisition: " + oRequisition.ReqNumber.ToString();
                        oTmp.RequisitionNumber = oRequisition.ReqNumber;
                        oTmp.RequestedBy = oRequisition.RequestorId;
                        oPartTimeDomainInfo.PartInfoTimeLineData.Add(oTmp);
                    }

                    // add in each job demand
                    foreach (HSDemandForPartFromJob oJobDemand in oPartTimeDomainInfo.Jobs)
                    {
                        oTmp = new HSPartInfoTimeLine();
                        oTmp.DateOfChange = oJobDemand.JobRequiredDate;
                        // ensure we have the min start date and max end dates
                        if (oTmp.DateOfChange < oPartTimeDomainInfo.StartDate)
                        {
                            oPartTimeDomainInfo.StartDate = oTmp.DateOfChange;
                        }
                        if (oTmp.DateOfChange > oPartTimeDomainInfo.EndDate)
                        {
                            oPartTimeDomainInfo.EndDate = oTmp.DateOfChange;
                        }
                        oTmp.NetChange = oJobDemand.NetDemandQuantity;
                        if (oTmp.NetChange > 0)
                        {
                            oTmp.NetChange = 0;
                        }
                        oTmp.ActionType = "JOB";
                        oTmp.ReasonForChange = "Job: " + oJobDemand.JobNumber;
                        oTmp.JobNum = oJobDemand.JobNumber;
                        oTmp.AssemblySequence = oJobDemand.AssemblySequence;
                        oTmp.MaterialSequence = oJobDemand.MaterialSequence;
                        oTmp.RelatedOperation = oJobDemand.RelatedOperation;
                        oPartTimeDomainInfo.PartInfoTimeLineData.Add(oTmp);
                    }

                    // add in each order demand
                    foreach (HSDemandForPartFromOrder oOrderDemand in oPartTimeDomainInfo.Orders)
                    {
                        oTmp = new HSPartInfoTimeLine();
                        oTmp.DateOfChange = oOrderDemand.ReleaseRequiredDate;
                        // ensure we have the min start date and max end dates
                        if (oTmp.DateOfChange < oPartTimeDomainInfo.StartDate)
                        {
                            oPartTimeDomainInfo.StartDate = oTmp.DateOfChange;
                        }
                        if (oTmp.DateOfChange > oPartTimeDomainInfo.EndDate)
                        {
                            oPartTimeDomainInfo.EndDate = oTmp.DateOfChange;
                        }
                        oTmp.NetChange = oOrderDemand.NetDemandQuantity;
                        if (oTmp.NetChange > 0)
                        {
                            oTmp.NetChange = 0;
                        }
                        oTmp.ActionType = "ORDER";
                        oTmp.ReasonForChange = "Order Shipped: " + oOrderDemand.OrderNumber.ToString() + "-" + oOrderDemand.LineNumber.ToString() + "-" + oOrderDemand.ReleaseNumber.ToString();
                        oTmp.OrderNum = oOrderDemand.OrderNumber;
                        oTmp.OrderLine = oOrderDemand.LineNumber;
                        oTmp.OrderRelNum = oOrderDemand.ReleaseNumber;
                        oPartTimeDomainInfo.PartInfoTimeLineData.Add(oTmp);
                    }

                    // add in the existing inventory on hand
                    oTmp = new HSPartInfoTimeLine();
                    oTmp.DateOfChange = oPartTimeDomainInfo.StartDate;
                    oPartTimeDomainInfo.StartingInventoryOnHand = oPartTimeDomainInfo.PartsOnHand.Sum(oItem => oItem.TotalOnHand);
                    oTmp.NetChange = oPartTimeDomainInfo.StartingInventoryOnHand;
                    oTmp.ActionType = "INVENTORY ON HAND";
                    oTmp.ReasonForChange = "INVENTORY ON HAND";
                    oPartTimeDomainInfo.PartInfoTimeLineData.Add(oTmp);

                    // place all events in chronological order
                    oPartTimeDomainInfo.PartInfoTimeLineData = oPartTimeDomainInfo.PartInfoTimeLineData.OrderBy(oItem => oItem.DateOfChange).ToList();

                    // get the maximum and minimum quantity for this period
                    decimal dMaximumQuantity = 0M;
                    decimal dMinimumQuantity = 0M;
                    decimal dCurrentInventoryLevel = 0M;

                    DateTime dtLastDate = oPartTimeDomainInfo.StartDate;
                    // check to make sure this date is not a min value
                    if (dtLastDate != DateTime.MinValue)
                    {
                        dtLastDate = dtLastDate.AddDays(-1);
                    }
                    DateTime dtDateInventoryWentNegative = DateTime.MinValue;
                    foreach (HSPartInfoTimeLine oTmpPartInfo in oPartTimeDomainInfo.PartInfoTimeLineData)
                    {
                        if (oPartInfo.PartQtyBearing == true)
                        {
                            dCurrentInventoryLevel = dCurrentInventoryLevel + oTmpPartInfo.NetChange;
                        }
                        // establish the current inventory level at this point in time
                        oTmpPartInfo.CurrentInventoryLevel = dCurrentInventoryLevel;
                        if (dCurrentInventoryLevel > dMaximumQuantity)
                        {
                            // we have a new maximum
                            dMaximumQuantity = dCurrentInventoryLevel;
                        }
                        if (dCurrentInventoryLevel < dMinimumQuantity)
                        {
                            // we have a new minimum
                            dMinimumQuantity = dCurrentInventoryLevel;
                        }

                        // if two events happen on the exact same date then we will artificially add 1 sec to separate the two events
                        if (oTmpPartInfo.DateOfChange == dtLastDate)
                        {
                            oTmpPartInfo.DateOfChange = oTmpPartInfo.DateOfChange.AddSeconds(1);
                        }

                        dtLastDate = oTmpPartInfo.DateOfChange;

                        // see if we are currently in negative territory - only look forward from todays date
                        if ((dCurrentInventoryLevel < 0) && (oTmpPartInfo.DateOfChange >= DateTime.Now))
                        {
                            if (dtDateInventoryWentNegative == DateTime.MinValue)
                            {
                                dtDateInventoryWentNegative = oTmpPartInfo.DateOfChange;
                            }
                            else if ((oTmpPartInfo.DateOfChange - dtDateInventoryWentNegative).TotalDays > 30)
                            {
                                // this part has been in negative territory for 1 month
                                oPartTimeDomainInfo.NegativeForMoreThanOneMonth = true;
                            }
                        }
                        else
                        {
                            // reset this as the inventory is now positive
                            dtDateInventoryWentNegative = DateTime.MinValue;
                        }
                    }

                    oPartTimeDomainInfo.MinQuantity = dMinimumQuantity;
                    oPartTimeDomainInfo.MaxQuantity = dMaximumQuantity;
                    oPartTimeDomainInfo.EndingQuantity = dCurrentInventoryLevel;

                    // now have all child objects point to the parent object
                    foreach (HSPartInfoTimeLine oPartInfoTimeLine in oPartTimeDomainInfo.PartInfoTimeLineData)
                    {
                        oPartInfoTimeLine.Parent = oPartTimeDomainInfo;
                    }
                }
            }

            #endregion
        }

        public void EvaluatePartShortagesForJobs()
        {
            #region Evaluate Job Needs

            // we will analyze the part needs by jobs -- looking for any jobs where the part inventory level goes negative
            List<string> oAllJobs = new List<string>();
            foreach (PartTimeDomainInfo oTimeDomain in m_oAllPartTimeDomainInfo)
            {
                List<string> oAllJobsForPart = oTimeDomain.PartInfoTimeLineData.Select(oItem => oItem.JobNum).ToList();
                oAllJobs.AddRange(oAllJobsForPart);
            }
            // get rid of all duplicate job numbers
            oAllJobs = oAllJobs.Distinct().ToList();

            // evaluate each job
            m_oPartDemandsForJobs.Clear();
            foreach (string sJobNum in oAllJobs)
            {
                if (string.IsNullOrEmpty(sJobNum) == true)
                {
                    // we do not process items when there is no job number specified
                    continue;
                }
                HSOpenJob oJobToReview = m_oAllOpenJobs.FirstOrDefault(oItem => string.Compare(oItem.JobNum, sJobNum, true) == 0);
                if (oJobToReview != null)
                {
                    foreach (PartTimeDomainInfo oTimeDomain in m_oAllPartTimeDomainInfo)
                    {
                        // get the list of part shortages for this job
                        List<HSPartInfoTimeLine> oTimeLinesForPartShortages = oTimeDomain.PartInfoTimeLineData.Where(oItem => (string.Compare(oItem.JobNum, sJobNum, true) == 0) && (oItem.NetChange < 0) && (oItem.CurrentInventoryLevel < 0)).ToList();
                        // get the list of all part demands for this job
                        List<HSPartInfoTimeLine> oTimeLinesForAllParts = oTimeDomain.PartInfoTimeLineData.Where(oItem => (string.Compare(oItem.JobNum, sJobNum, true) == 0)).ToList();

                        // if there are part shortages then add the list of short parts to the job info
                        if (oTimeLinesForPartShortages.Count > 0)
                        {
                            // indicate that this job has part shortages
                            oJobToReview.PartShortage = true;
                            // indicate that these parts go negative for this job
                            foreach (HSPartInfoTimeLine oPartShort in oTimeLinesForPartShortages)
                            {
                                oPartShort.PartShortage = true;
                            }
                        }

                        if (oTimeLinesForAllParts.Count > 0)
                        {
                            // we always create a list of part demand for open jobs
                            m_oPartDemandsForJobs.AddRange(oTimeLinesForAllParts);
                        }
                    }
                }
            }

            #endregion
        }

        public void EvaluatePartShortagesForSalesOrders()
        {
            #region Evaluate Sales Order Needs

            //// we will analyze the part needs by sales order -- looking for any sales orders where the part inventory level goes negative
            //List<UniqueOrder> oAllSalesOrders = new List<UniqueOrder>();
            //foreach (PartTimeDomainInfo oTimeDomain in m_oAllPartTimeDomainInfo)
            //{
            //    // we will remove all mfg parts and only look at purchased parts for these orders
            //    HSPartData oPart = m_oAllPartData.FirstOrDefault(oItem => string.Compare(oItem.PartNum, oTimeDomain.PartNum, true) == 0);
            //    if (string.Compare(oPart.TypeCode, "P", true) == 0)
            //    {
            //        foreach (HSPartInfoTimeLine oTimeLine in oTimeDomain.PartInfoTimeLineData)
            //        {
            //            UniqueOrder oOrder = new UniqueOrder();
            //            oOrder.m_iOrderNum = oTimeLine.OrderNum;
            //            oOrder.m_iOrderLine = oTimeLine.OrderLine;
            //            oOrder.m_iOrderRel = oTimeLine.OrderRelNum;
            //            bool bAdd = true;
            //            foreach (UniqueOrder o in oAllSalesOrders)
            //            {
            //                if ((oOrder.m_iOrderNum == o.m_iOrderNum) && (oOrder.m_iOrderLine == o.m_iOrderLine) && (oOrder.m_iOrderRel == o.m_iOrderRel))
            //                {
            //                    bAdd = false;
            //                    break;
            //                }
            //            }
            //            if (bAdd == true)
            //            {
            //                oAllSalesOrders.Add(oOrder);
            //            }
            //        }
            //    }
            //}

            // evaluate each sales order
            m_oPartDemandForOrders.Clear();

            // examine each order group to see if there are part shortages
            foreach (SalesOrderGroup oOrderGroup in m_oOrdersToAnalyze)
            {
                // look at each sales order in this group to see if there are part shortages
                foreach (SOBacklogForPartAnalysis oOrder in oOrderGroup.OrdersThatMustShipTogether)
                {
                    foreach (PartTimeDomainInfo oTimeDomain in m_oAllPartTimeDomainInfo)
                    {
                        // get the list of part shortages for this order
                        List<HSPartInfoTimeLine> oTimeLinesForPartShortages = oTimeDomain.PartInfoTimeLineData.Where(oItem => (oItem.OrderNum == oOrder.OrderNum) && (oItem.OrderLine == oOrder.OrderLine) &&
                        (oItem.OrderRelNum == oOrder.OrderRelease) && (oItem.CurrentInventoryLevel < 0)).ToList();

                        HSPartInfoTimeLine oTimeLine = oTimeDomain.PartInfoTimeLineData.FirstOrDefault(oItem => (oItem.OrderNum == oOrder.OrderNum) &&
                            (oItem.OrderLine == oOrder.OrderLine) && (oItem.OrderRelNum == oOrder.OrderRelease));

                        // if there are part shortages then add the list of short parts to the order
                        if (oTimeLinesForPartShortages.Count > 0)
                        {
                            // indicate that this order has part shortages
                            oOrder.PartShortage = true;
                            // indicate that these parts go negative for this job
                            foreach (HSPartInfoTimeLine oPartShort in oTimeLinesForPartShortages)
                            {
                                oPartShort.PartShortage = true;
                            }
                        }

                        if (oTimeLine != null)
                        {
                            m_oPartDemandForOrders.Add(oTimeLine);
                        }
                    }
                }
            }


            #endregion
        }

        public void EvaluateJobDependencies()
        {
            // walk through all open jobs and if they are tied to another job put them in the dependency list
            foreach (HSOpenJob oJob in m_oAllOpenJobs)
            {
                // see if this job is tied to another job
                if (string.IsNullOrEmpty(oJob.TargetJobNum) == false)
                {
                    // find the parent job
                    string sParentJobNum = oJob.TargetJobNum;
                    HSOpenJob oParentJob = m_oAllOpenJobs.FirstOrDefault(oItem => string.Compare(oItem.JobNum, sParentJobNum, true) == 0);
                    if (oParentJob != null)
                    {
                        oParentJob.AddRelatedJob(oJob);
                    }
                }
            }
        }

        public void EvaluateSalesOrderDependencyOnJobs()
        {
            // walk through all open jobs and see if the job is tied to an order and if so put the job in the sales order's dependency list
            foreach (HSOpenJob oJob in m_oAllOpenJobs)
            {
                // see if this job is tied to an order
                if ((oJob.OrderNum != 0) && (oJob.OrderLine != 0) && (oJob.OrderRelNum != 0))
                {
                    // find the parent sales order
                    SOBacklogForPartAnalysis oParentOrder = m_oAllBacklogDetails.FirstOrDefault(oItem => (oItem.OrderNum == oJob.OrderNum) && (oItem.OrderLine == oJob.OrderLine) && (oItem.OrderRelease == oJob.OrderRelNum)); ;
                    if (oParentOrder != null)
                    {
                        oParentOrder.AddRelatedJob(oJob);
                    }
                }
            }
        }

        public void CreateReport(string sArchiveFileDirectory, string sTmpFileDirectory, string sJobNum, bool bOnlyPartShortages, DateTime dtCutOffDate, HSUser oRequestingUser)
        {
            HSOpenJob oJob = m_oAllOpenJobs.FirstOrDefault(oItem => string.Compare(oItem.JobNum, sJobNum, true) == 0);
            if (oJob != null)
            {
                // now get the list of parts to print out for this job
                List<HSPartInfoTimeLine> oFinalPartList = new List<HSPartInfoTimeLine>();
                List<HSPartInfoTimeLine> oPartsForJob = m_oPartDemandsForJobs.Where(oItem => string.Compare(oItem.JobNum, sJobNum, true) == 0).ToList();
                if (oPartsForJob.Count > 0)
                {
                    // put these in order of assembly then mtl seq
                    oPartsForJob = oPartsForJob.OrderBy(oItem => oItem.AssemblySequence).ThenBy(x => x.MaterialSequence).ToList();
                    foreach (HSPartInfoTimeLine oPartInfo in oPartsForJob)
                    {
                        bool bPrintPartInfo = true;
                        if (bOnlyPartShortages == true)
                        {
                            if (oPartInfo.CurrentInventoryLevel >= 0)
                            {
                                bPrintPartInfo = false;
                            }
                        }
                        if (bPrintPartInfo == true)
                        {
                            oFinalPartList.Add(oPartInfo);
                        }
                    }

                    GenerateReport(sArchiveFileDirectory, sTmpFileDirectory, oJob, oFinalPartList, dtCutOffDate, oRequestingUser);
                }
            }

        }

        public void GenerateReport(string sArchiveFileDirectory, string sTmpFileDirectory, HSOpenJob oJob, List<HSPartInfoTimeLine> oFinalPartList, DateTime dtCutOffDate, HSUser oRequestingUser)
        {
            #region General Setup
            DateTime dtToday = DateTime.Now;
            string sDestinationFileName = sTmpFileDirectory + "Federal Part Analysis For Job-" + dtToday.Month.ToString() + "-" + dtToday.Day.ToString() + "-" + dtToday.Year.ToString() + ".xlsx";
            string sArchiveFileName = sArchiveFileDirectory + "Federal Part Analysis For Job-" + dtToday.Month.ToString() + "-" + dtToday.Day.ToString() + "-" + dtToday.Year.ToString() + ".xlsx";

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
            if (oRequestingUser != null)
            {
                // this report was requested by someone directly from Epicor
                oToAddresses.Add(oRequestingUser.Email);
            }
            oToAddresses.Add(HSUser.SPINCRAFT_ROOT_USER);

            bool bFirstWorksheet = true;
            SLDocument oPartAnalysisForJobDocument = new SLDocument();

            // set up the style of cells
            SLStyle oMoney = oPartAnalysisForJobDocument.CreateStyle();
            oMoney.FormatCode = "$#,##0.00";
            SLStyle oPercentage = oPartAnalysisForJobDocument.CreateStyle();
            oPercentage.FormatCode = "0.00%";
            SLStyle oGood = oPartAnalysisForJobDocument.CreateStyle();
            oGood.ApplyNamedCellStyle(SLNamedCellStyleValues.Good);
            SLStyle oNeutral = oPartAnalysisForJobDocument.CreateStyle();
            oNeutral.ApplyNamedCellStyle(SLNamedCellStyleValues.Neutral);
            SLStyle oBad = oPartAnalysisForJobDocument.CreateStyle();
            oBad.ApplyNamedCellStyle(SLNamedCellStyleValues.Bad);

            System.Drawing.Color oRed = System.Drawing.Color.FromArgb(255, 32, 32);
            System.Drawing.Color oGreen = System.Drawing.Color.FromArgb(0, 100, 5);
            System.Drawing.Color oDRed = System.Drawing.Color.FromArgb(255, 100, 100);
            System.Drawing.Color oDGreen = System.Drawing.Color.FromArgb(50, 100, 32);
            System.Drawing.Color oDYellow = System.Drawing.Color.FromArgb(200, 200, 50);
            System.Drawing.Color oDBlue = System.Drawing.Color.FromArgb(0, 100, 200);
            System.Drawing.Color oDOrange = System.Drawing.Color.FromArgb(255, 128, 2);

            SLStyle oBelowSafetyStyle = new SLStyle();
            oBelowSafetyStyle.SetFontBold(true);
            oBelowSafetyStyle.SetFontColor(oDOrange);

            SLStyle oBelowMinimumStyle = new SLStyle();
            oBelowMinimumStyle.SetFontBold(true);
            oBelowMinimumStyle.SetFontColor(oDYellow);

            SLStyle oBelowZeroStyle = new SLStyle();
            oBelowZeroStyle.SetFontBold(true);
            oBelowZeroStyle.SetFontColor(oRed);

            SLStyle oLatePOStyle = new SLStyle();
            oLatePOStyle.SetFontBold(true);
            oLatePOStyle.SetFontColor(oRed);

            SLFill oSLFill = new SLFill();
            oSLFill.SetPatternBackgroundColor(SLThemeColorIndexValues.Dark2Color);

            SLAlignment oLeftAlignment = new SLAlignment();
            oLeftAlignment.Horizontal = HorizontalAlignmentValues.Left;

            SLStyle oLeftAlignmentStyle = new SLStyle();
            oLeftAlignmentStyle.Alignment = oLeftAlignment;

            SLAlignment oRightAlignment = new SLAlignment();
            oRightAlignment.Horizontal = HorizontalAlignmentValues.Right;

            SLStyle oRightAlignmentStyle = new SLStyle();
            oRightAlignmentStyle.Alignment = oRightAlignment;

            SLStyle oBoldStyle = new SLStyle();
            oBoldStyle.SetFontBold(true);
            oBoldStyle.SetTopBorder(BorderStyleValues.Thin, SLThemeColorIndexValues.Dark2Color);
            oBoldStyle.SetBottomBorder(BorderStyleValues.Thin, SLThemeColorIndexValues.Dark2Color);
            oBoldStyle.SetLeftBorder(BorderStyleValues.Thin, SLThemeColorIndexValues.Dark2Color);
            oBoldStyle.SetRightBorder(BorderStyleValues.Thin, SLThemeColorIndexValues.Dark2Color);
            oBoldStyle.SetFontColor(SLThemeColorIndexValues.Dark2Color);

            m_oGood = oPartAnalysisForJobDocument.CreateStyle();
            m_oGood.ApplyNamedCellStyle(SLNamedCellStyleValues.Good);

            m_oNeutral = oPartAnalysisForJobDocument.CreateStyle();
            m_oNeutral.ApplyNamedCellStyle(SLNamedCellStyleValues.Neutral);

            m_oBad = oPartAnalysisForJobDocument.CreateStyle();
            m_oBad.ApplyNamedCellStyle(SLNamedCellStyleValues.Bad);

            m_oHeader = oPartAnalysisForJobDocument.CreateStyle();
            m_oHeader.SetFontBold(true);
            m_oHeader.SetFont("Calibri", 12);
            m_oHeader.Fill.SetPattern(PatternValues.Solid, SLThemeColorIndexValues.Accent5Color, SLThemeColorIndexValues.Dark1Color);

            m_oSubHeader = oPartAnalysisForJobDocument.CreateStyle();
            m_oSubHeader.SetFontBold(true);
            m_oSubHeader.SetFont("Calibri", 12);
            m_oSubHeader.Fill.SetPattern(PatternValues.Solid, SLThemeColorIndexValues.Accent4Color, SLThemeColorIndexValues.Dark1Color);

            m_oMoney1 = oPartAnalysisForJobDocument.CreateStyle();
            m_oMoney1.Fill.SetPattern(PatternValues.Solid, SLThemeColorIndexValues.Light1Color, SLThemeColorIndexValues.Accent1Color);
            m_oMoney1.FormatCode = "$#,##0.00";

            m_oMoney2 = oPartAnalysisForJobDocument.CreateStyle();
            m_oMoney2.Fill.SetPattern(PatternValues.Solid, SLThemeColorIndexValues.Light2Color, SLThemeColorIndexValues.Accent2Color);
            m_oMoney2.FormatCode = "$#,##0.00";

            m_oBoldMoney = oPartAnalysisForJobDocument.CreateStyle();
            m_oBoldMoney.SetFontBold(true);
            m_oBoldMoney.FormatCode = "$#,##0.00";

            m_oOrder1 = oPartAnalysisForJobDocument.CreateStyle();
            m_oOrder1.Fill.SetPattern(PatternValues.Solid, SLThemeColorIndexValues.Light1Color, SLThemeColorIndexValues.Accent1Color);

            m_oOrder2 = oPartAnalysisForJobDocument.CreateStyle();
            m_oOrder2.Fill.SetPattern(PatternValues.Solid, SLThemeColorIndexValues.Light2Color, SLThemeColorIndexValues.Accent2Color);

            #endregion

            //set up column headers
            int iNumOfColumns = 0;
            int iNumOfRows = 0;

            if (bFirstWorksheet == true)
            {
                oPartAnalysisForJobDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "Part Analysis");
                bFirstWorksheet = false;
            }
            else
            {
                oPartAnalysisForJobDocument.AddWorksheet("Part Analysis");
            }


            #region General Job Info

            //set up column headers
            iNumOfColumns = 0;
            iNumOfRows = 1;
            iNumOfColumns++;
            oPartAnalysisForJobDocument.SetCellValue(iNumOfRows, iNumOfColumns, "Job Num");
            oPartAnalysisForJobDocument.SetColumnWidth(iNumOfColumns, 15);
            oPartAnalysisForJobDocument.SetCellStyle(iNumOfRows, iNumOfColumns, m_oHeader);
            iNumOfColumns++;
            oPartAnalysisForJobDocument.SetCellValue(iNumOfRows, iNumOfColumns, "Mfg Part Num");
            oPartAnalysisForJobDocument.SetColumnWidth(iNumOfColumns, 15);
            oPartAnalysisForJobDocument.SetCellStyle(iNumOfRows, iNumOfColumns, m_oHeader);
            iNumOfColumns++;
            oPartAnalysisForJobDocument.SetCellValue(iNumOfRows, iNumOfColumns, "Start Date");
            oPartAnalysisForJobDocument.SetColumnWidth(iNumOfColumns, 15);
            oPartAnalysisForJobDocument.SetCellStyle(iNumOfRows, iNumOfColumns, m_oHeader);
            iNumOfColumns++;
            oPartAnalysisForJobDocument.SetCellValue(iNumOfRows, iNumOfColumns, "Req Date");
            oPartAnalysisForJobDocument.SetColumnWidth(iNumOfColumns, 15);
            oPartAnalysisForJobDocument.SetCellStyle(iNumOfRows, iNumOfColumns, m_oHeader);
            iNumOfColumns++;
            oPartAnalysisForJobDocument.SetCellValue(iNumOfRows, iNumOfColumns, "Engineered");
            oPartAnalysisForJobDocument.SetColumnWidth(iNumOfColumns, 15);
            oPartAnalysisForJobDocument.SetCellStyle(iNumOfRows, iNumOfColumns, m_oHeader);
            iNumOfColumns++;
            oPartAnalysisForJobDocument.SetCellValue(iNumOfRows, iNumOfColumns, "Firm");
            oPartAnalysisForJobDocument.SetColumnWidth(iNumOfColumns, 15);
            oPartAnalysisForJobDocument.SetCellStyle(iNumOfRows, iNumOfColumns, m_oHeader);
            iNumOfColumns++;
            oPartAnalysisForJobDocument.SetCellValue(iNumOfRows, iNumOfColumns, "Released");
            oPartAnalysisForJobDocument.SetColumnWidth(iNumOfColumns, 15);
            oPartAnalysisForJobDocument.SetCellStyle(iNumOfRows, iNumOfColumns, m_oHeader);
            iNumOfColumns++;
            oPartAnalysisForJobDocument.SetCellValue(iNumOfRows, iNumOfColumns, "Production Qty");
            oPartAnalysisForJobDocument.SetColumnWidth(iNumOfColumns, 15);
            oPartAnalysisForJobDocument.SetCellStyle(iNumOfRows, iNumOfColumns, m_oHeader);
            iNumOfColumns++;
            oPartAnalysisForJobDocument.SetCellValue(iNumOfRows, iNumOfColumns, "Cust ID");
            oPartAnalysisForJobDocument.SetColumnWidth(iNumOfColumns, 15);
            oPartAnalysisForJobDocument.SetCellStyle(iNumOfRows, iNumOfColumns, m_oHeader);
            iNumOfColumns++;
            oPartAnalysisForJobDocument.SetCellValue(iNumOfRows, iNumOfColumns, "Customer Name");
            oPartAnalysisForJobDocument.SetColumnWidth(iNumOfColumns, 30);
            oPartAnalysisForJobDocument.SetCellStyle(iNumOfRows, iNumOfColumns, m_oHeader);
            iNumOfColumns++;
            oPartAnalysisForJobDocument.SetCellValue(iNumOfRows, iNumOfColumns, "CUT OFF DATE");
            oPartAnalysisForJobDocument.SetColumnWidth(iNumOfColumns, 15);
            oPartAnalysisForJobDocument.SetCellStyle(iNumOfRows, iNumOfColumns, m_oHeader);

            // now fill in job info
            iNumOfRows++;
            iNumOfColumns = 1;
            oPartAnalysisForJobDocument.SetCellValue(iNumOfRows, iNumOfColumns++, oJob.JobNum);
            oPartAnalysisForJobDocument.SetCellValue(iNumOfRows, iNumOfColumns++, oJob.PartNum);
            oPartAnalysisForJobDocument.SetCellValue(iNumOfRows, iNumOfColumns++, oJob.StartDate.ToShortDateString());
            oPartAnalysisForJobDocument.SetCellValue(iNumOfRows, iNumOfColumns, oJob.RequiredDate.ToShortDateString());
            if (oJob.RequiredDate <= DateTime.Now)
            {
                oPartAnalysisForJobDocument.SetCellStyle(iNumOfRows, iNumOfColumns, m_oBad);
            }
            iNumOfColumns++;
            oPartAnalysisForJobDocument.SetCellValue(iNumOfRows, iNumOfColumns++, oJob.Engineered);
            oPartAnalysisForJobDocument.SetCellValue(iNumOfRows, iNumOfColumns++, oJob.Firm);
            oPartAnalysisForJobDocument.SetCellValue(iNumOfRows, iNumOfColumns++, oJob.Released);
            oPartAnalysisForJobDocument.SetCellValue(iNumOfRows, iNumOfColumns++, oJob.ProductionQty.ToString());
            oPartAnalysisForJobDocument.SetCellValue(iNumOfRows, iNumOfColumns++, oJob.CustomerID);
            oPartAnalysisForJobDocument.SetCellValue(iNumOfRows, iNumOfColumns++, oJob.CustomerName);
            if (dtCutOffDate != DateTime.MinValue)
            {
                oPartAnalysisForJobDocument.SetCellValue(iNumOfRows, iNumOfColumns++, dtCutOffDate.ToShortDateString());
            }
            #endregion


            #region Part Info

            // we will skip a few rows and one column start a new header
            iNumOfRows++;
            iNumOfRows++;

            iNumOfColumns = 2;
            oPartAnalysisForJobDocument.SetCellValue(iNumOfRows, iNumOfColumns, "Assy Seq");
            oPartAnalysisForJobDocument.SetColumnWidth(iNumOfColumns, 15);
            oPartAnalysisForJobDocument.SetCellStyle(iNumOfRows, iNumOfColumns, m_oHeader);
            iNumOfColumns++;
            oPartAnalysisForJobDocument.SetCellValue(iNumOfRows, iNumOfColumns, "Mat Seq");
            oPartAnalysisForJobDocument.SetColumnWidth(iNumOfColumns, 15);
            oPartAnalysisForJobDocument.SetCellStyle(iNumOfRows, iNumOfColumns, m_oHeader);
            iNumOfColumns++;
            oPartAnalysisForJobDocument.SetCellValue(iNumOfRows, iNumOfColumns, "Opr");
            oPartAnalysisForJobDocument.SetColumnWidth(iNumOfColumns, 15);
            oPartAnalysisForJobDocument.SetCellStyle(iNumOfRows, iNumOfColumns, m_oHeader);
            iNumOfColumns++;
            oPartAnalysisForJobDocument.SetCellValue(iNumOfRows, iNumOfColumns, "Mtl Part Num");
            oPartAnalysisForJobDocument.SetColumnWidth(iNumOfColumns, 15);
            oPartAnalysisForJobDocument.SetCellStyle(iNumOfRows, iNumOfColumns, m_oHeader);
            iNumOfColumns++;
            oPartAnalysisForJobDocument.SetCellValue(iNumOfRows, iNumOfColumns, "Req Qty");
            oPartAnalysisForJobDocument.SetColumnWidth(iNumOfColumns, 15);
            oPartAnalysisForJobDocument.SetCellStyle(iNumOfRows, iNumOfColumns, m_oHeader);
            iNumOfColumns++;
            oPartAnalysisForJobDocument.SetCellValue(iNumOfRows, iNumOfColumns, "Req Date");
            oPartAnalysisForJobDocument.SetColumnWidth(iNumOfColumns, 15);
            oPartAnalysisForJobDocument.SetCellStyle(iNumOfRows, iNumOfColumns, m_oHeader);
            iNumOfColumns++;
            oPartAnalysisForJobDocument.SetCellValue(iNumOfRows, iNumOfColumns, "Balance");
            oPartAnalysisForJobDocument.SetColumnWidth(iNumOfColumns, 15);
            oPartAnalysisForJobDocument.SetCellStyle(iNumOfRows, iNumOfColumns, m_oHeader);


            if (oFinalPartList.Count > 0)
            {
                // put these in order of assembly then mtl seq
                oFinalPartList = oFinalPartList.OrderBy(oItem => oItem.AssemblySequence).ThenBy(x => x.MaterialSequence).ToList();
                foreach (HSPartInfoTimeLine oPartInfo in oFinalPartList)
                {
                    iNumOfRows++;
                    iNumOfColumns = 2;

                    oPartAnalysisForJobDocument.SetCellValue(iNumOfRows, iNumOfColumns++, oPartInfo.AssemblySequence.ToString());
                    oPartAnalysisForJobDocument.SetCellValue(iNumOfRows, iNumOfColumns++, oPartInfo.MaterialSequence.ToString());
                    oPartAnalysisForJobDocument.SetCellValue(iNumOfRows, iNumOfColumns++, oPartInfo.RelatedOperation.ToString());
                    oPartAnalysisForJobDocument.SetCellValue(iNumOfRows, iNumOfColumns++, oPartInfo.Parent.PartNum);
                    oPartAnalysisForJobDocument.SetCellValue(iNumOfRows, iNumOfColumns++, Math.Abs(oPartInfo.NetChange).ToString());
                    oPartAnalysisForJobDocument.SetCellValue(iNumOfRows, iNumOfColumns++, oPartInfo.DateOfChange.ToShortDateString());
                    oPartAnalysisForJobDocument.SetCellValue(iNumOfRows, iNumOfColumns, oPartInfo.CurrentInventoryLevel.ToString());
                    if (oPartInfo.CurrentInventoryLevel < 0)
                    {
                        oPartAnalysisForJobDocument.SetCellStyle(iNumOfRows, iNumOfColumns, m_oBad);
                    }
                    iNumOfColumns++;
                }
            }
            #endregion

            // send the email to the requestor
            if (bFirstWorksheet == false)
            {
                // we create a worksheet so we have some data in the spreadsheet so we email the spreadsheet
                oPartAnalysisForJobDocument.SaveAs(sDestinationFileName);
                List<string> oAttachments = new List<string>();
                oAttachments.Add(sDestinationFileName);
                HSEmailHelper.SendEmail(oToAddresses, "Federal - Part Analysis For Job", "Federal - Part Analysis For Job", oAttachments);
            }
        }

        #endregion

        #region Data Members

        public static string STOCKING_JOB = "STOCK";

        private string m_sCompany;

        // available parts
        private List<HSPartsOnHand> m_oPartsOnHand = new List<HSPartsOnHand>();
        // production calendar
        private List<ProductionCalendar> m_oProductionCalendarDays = new List<ProductionCalendar>();

        // all sales order backlog
        private List<SOBacklogForPartAnalysis> m_oAllBacklogDetails = new List<SOBacklogForPartAnalysis>();
        // list of orders we will analyze placed in groups per shipping constraints (order complete, line complete)
        private List<SalesOrderGroup> m_oOrdersToAnalyze = new List<SalesOrderGroup>();

        // required parts from jobs
        private List<HSDemandForPartFromJob> m_oPartDemandFromJobs = new List<HSDemandForPartFromJob>();
        private bool m_bJobDemandEngineered = false;
        private bool m_bJobDemandFirmed = false;
        private bool m_bJobDemandReleased = false;
        // required parts from sales orders
        private List<HSDemandForPartFromOrder> m_oPartDemandFromOrders = new List<HSDemandForPartFromOrder>();
        private bool m_bSalesOrderDemandFirmed;
        // required parts from requisitions
        private List<HSDemandForPartsFromRequisition> m_oPartDemandFromRequisitions = new List<HSDemandForPartsFromRequisition>();
        // supplied parts from POs
        private List<HSSourceForPartFromPO> m_oPartsSuppliedByPOs = new List<HSSourceForPartFromPO>();
        private bool m_bPOSupplyConfirmed = false;
        private bool m_bPOSupplyApproved = false;
        // list of all open jobs
        private List<HSOpenJob> m_oAllOpenJobs = new List<HSOpenJob>();
        // list of supply jobs (jobs that are making parts for other jobs)
        private List<HSOpenJob> m_oSupplyJobs = new List<HSOpenJob>();
        private bool m_bJobSupplyFirmed = false;
        private bool m_bJobSupplyEngineered = false;
        private bool m_bJobSupplyReleased = false;

        // part demand through time
        private List<HSDemandForPartsInTime> m_oPartDemandInTime = new List<HSDemandForPartsInTime>();

        // part information
        private List<string> m_oAllPartNumbers = new List<string>();
        private List<HSPartData> m_oAllPartData = new List<HSPartData>();
        private List<PartTimeDomainInfo> m_oAllPartTimeDomainInfo = new List<PartTimeDomainInfo>();

        // all open jobs with related part info
        List<HSPartInfoTimeLine> m_oPartDemandsForJobs = new List<HSPartInfoTimeLine>();

        // all open sales orders with related part info
        List<HSPartInfoTimeLine> m_oPartDemandForOrders = new List<HSPartInfoTimeLine>();

        private SLStyle m_oGood;
        private SLStyle m_oNeutral;
        private SLStyle m_oBad;
        private SLStyle m_oHeader;
        private SLStyle m_oSubHeader;
        private SLStyle m_oMoney1;
        private SLStyle m_oMoney2;
        private SLStyle m_oBoldMoney;
        private SLStyle m_oOrder1;
        private SLStyle m_oOrder2;
        private bool m_bOrderStyle1 = true;

        #endregion
    }

}
