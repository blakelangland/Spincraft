using Epicor.Utilities;
using Erp.Contracts;
using Erp.Proxy.BO;
using Erp.Tablesets;
using HorizonScientific;
using Ice.BO;
using Ice.Core;
using Ice.Lib.Framework;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace HorizonScientific
{
    public class HSOpenJob
    {
        #region Constructors

        public HSOpenJob(DataRow oRow)
        {
            if ((oRow["JobHead_Company"] != DBNull.Value) && (string.IsNullOrEmpty((string)oRow["JobHead_Company"]) == false))
            {
                m_sCompany = (string)oRow["JobHead_Company"];
            }
            if ((oRow["JobHead_JobNum"] != DBNull.Value) && (string.IsNullOrEmpty((string)oRow["JobHead_JobNum"]) == false))
            {
                m_sJobNum = (string)oRow["JobHead_JobNum"];
            }
            if ((oRow["JobHead_PartNum"] != DBNull.Value) && (string.IsNullOrEmpty((string)oRow["JobHead_PartNum"]) == false))
            {
                m_sPartNum = (string)oRow["JobHead_PartNum"];
            }
            if ((oRow["JobHead_RevisionNum"] != DBNull.Value) && (string.IsNullOrEmpty((string)oRow["JobHead_RevisionNum"]) == false))
            {
                m_sRevNum = (string)oRow["JobHead_RevisionNum"];
            }
            if ((oRow["JobHead_PartDescription"] != DBNull.Value) && (string.IsNullOrEmpty((string)oRow["JobHead_PartDescription"]) == false))
            {
                m_sDescription = (string)oRow["JobHead_PartDescription"];
                m_sDescription = FormatString(m_sDescription);
            }
            if (oRow["JobHead_JobComplete"] != DBNull.Value)
            {
                m_bComplete = (bool)oRow["JobHead_JobComplete"];
            }
            if (oRow["JobHead_JobEngineered"] != DBNull.Value)
            {
                m_bEngineered = (bool)oRow["JobHead_JobEngineered"];
            }
            if (oRow["JobHead_JobReleased"] != DBNull.Value)
            {
                m_bReleased = (bool)oRow["JobHead_JobReleased"];
            }
            if (oRow["JobHead_JobFirm"] != DBNull.Value)
            {
                m_bFirm = (bool)oRow["JobHead_JobFirm"];
            }
            if (oRow["JobHead_ProdQty"] != DBNull.Value)
            {
                m_dProductionQty = (decimal)oRow["JobHead_ProdQty"];
            }
            if (oRow["JobHead_StartDate"] != DBNull.Value)
            {
                m_dtStartDate = (DateTime)oRow["JobHead_StartDate"];
                m_bStartDateNotSet = false;
            }
            else
            {
                m_dtStartDate = DateTime.Now;
                m_bStartDateNotSet = true;
            }
            if (oRow["JobHead_DueDate"] != DBNull.Value)
            {
                m_dtDueDate = (DateTime)oRow["JobHead_DueDate"];
                m_bDueDateNotSet = false;
            }
            else
            {
                m_dtDueDate = DateTime.Now;
                m_bDueDateNotSet = true;
            }
            if (oRow["JobHead_ReqDueDate"] != DBNull.Value)
            {
                m_dtRequiredDate = (DateTime)oRow["JobHead_ReqDueDate"];
                m_bRequiredDateNotSet = false;
            }
            else
            {
                m_dtRequiredDate = DateTime.Now;
                m_bRequiredDateNotSet = true;
            }
            if ((oRow["Customer_CustID"] != DBNull.Value) && (string.IsNullOrEmpty((string)oRow["Customer_CustID"]) == false))
            {
                m_sCustID = (string)oRow["Customer_CustID"];
            }
            if ((oRow["Customer_Name"] != DBNull.Value) && (string.IsNullOrEmpty((string)oRow["Customer_Name"]) == false))
            {
                m_sCustomerName = (string)oRow["Customer_Name"];
            }

            // check for stocking orders
            if (string.IsNullOrEmpty(m_sCustID) == true)
            {
                m_sCustID = HSAnalyzePartShortages.STOCKING_JOB;
            }
            if (string.IsNullOrEmpty(m_sCustomerName) == true)
            {
                m_sCustomerName = HSAnalyzePartShortages.STOCKING_JOB;
            }

            if ((oRow["JobProd_TargetJobNum"] != DBNull.Value) && (string.IsNullOrEmpty((string)oRow["JobProd_TargetJobNum"]) == false))
            {
                m_sTargetJobNum = (string)oRow["JobProd_TargetJobNum"];
            }
            if (oRow["JobProd_TargetAssemblySeq"] != DBNull.Value)
            {
                m_iTargetAssemblyNum = (int)oRow["JobProd_TargetAssemblySeq"];
            }
            if (oRow["JobProd_TargetMtlSeq"] != DBNull.Value)
            {
                m_iTargetMaterialSeq = (int)oRow["JobProd_TargetMtlSeq"];
            }
            if (oRow["JobProd_OrderNum"] != DBNull.Value)
            {
                m_iOrderNum = (int)oRow["JobProd_OrderNum"];
            }
            if (oRow["JobProd_OrderLine"] != DBNull.Value)
            {
                m_iOrderLine = (int)oRow["JobProd_OrderLine"];
            }
            if (oRow["JobProd_OrderRelNum"] != DBNull.Value)
            {
                m_iOrderRelNum = (int)oRow["JobProd_OrderRelNum"];
            }
            if (oRow["JobProd_ShippedQty"] != DBNull.Value)
            {
                m_dShippedQuantity = (decimal)oRow["JobProd_ShippedQty"];
            }
            if (oRow["JobProd_ReceivedQty"] != DBNull.Value)
            {
                m_dReceivedQuantity = (decimal)oRow["JobProd_ReceivedQty"];
            }
            if (oRow["Calculated_LastClockInDate"] != DBNull.Value)
            {
                m_dtLastClockInDate = (DateTime)oRow["Calculated_LastClockInDate"];
            }
        }
        #endregion

        #region Methods
        public void AddRelatedJob(HSOpenJob oJob)
        {
            m_oRelatedJobs.Add(oJob);
        }
        private string FormatString(string sTextToFormat)
        {
            string sResult = sTextToFormat;
            if (string.IsNullOrEmpty(sTextToFormat) == false)
            {
                int iMax = sTextToFormat.Length;
                if (iMax > 50)
                {
                    iMax = 50;
                }
                sResult = sTextToFormat.SubString(0, iMax - 1);
            }
            return sResult;
        }
        #endregion

        #region Properties

        public string Company
        {
            get { return m_sCompany; }
        }
        public string JobNum
        {
            get { return m_sJobNum; }
        }
        public string PartNum
        {
            get { return m_sPartNum; }
        }
        public string RevNum
        {
            get { return m_sRevNum; }
        }
        public string Description
        {
            get { return m_sDescription; }
        }
        public bool Complete
        {
            get { return m_bComplete; }
        }
        public bool Engineered
        {
            get { return m_bEngineered; }
        }
        public bool Released
        {
            get { return m_bReleased; }
        }
        public bool Firm
        {
            get { return m_bFirm; }
        }
        public decimal ProductionQty
        {
            get { return m_dProductionQty; }
        }
        public DateTime StartDate
        {
            get { return m_dtStartDate; }
        }
        public DateTime DueDate
        {
            get { return m_dtDueDate; }
        }
        public DateTime RequiredDate
        {
            get { return m_dtRequiredDate; }
        }
        public string CustomerID
        {
            get { return m_sCustID; }
        }
        public string CustomerName
        {
            get { return m_sCustomerName; }
        }
        public string TargetJobNum
        {
            get { return m_sTargetJobNum; }
        }
        public int TargetAssemblyNum
        {
            get { return m_iTargetAssemblyNum; }
        }
        public int TargetMaterialSeq
        {
            get { return m_iTargetMaterialSeq; }
        }
        public int OrderNum
        {
            get { return m_iOrderNum; }
        }
        public int OrderLine
        {
            get { return m_iOrderLine; }
        }
        public int OrderRelNum
        {
            get { return m_iOrderRelNum; }
        }
        public decimal ShippedQuantity
        {
            get { return m_dShippedQuantity; }
        }
        public decimal ReceivedQuantity
        {
            get { return m_dReceivedQuantity; }
        }
        public DateTime LastClockInDate
        {
            get { return m_dtLastClockInDate; }
        }

        public decimal RemainingQty
        {
            get
            {
                decimal dRemainingQty = m_dProductionQty - m_dShippedQuantity - m_dReceivedQuantity;
                if (dRemainingQty < 0)
                {
                    dRemainingQty = 0;
                }
                return dRemainingQty;
            }
        }
        public bool StartDateNotSet
        {
            get { return m_bStartDateNotSet; }
        }
        public bool DueDateNotSet
        {
            get { return m_bDueDateNotSet; }
        }
        public bool RequiredDateNotSet
        {
            get { return m_bRequiredDateNotSet; }
        }

        public bool PartShortage
        {
            get
            {
                bool bShortage = m_bPartShortage;
                if (bShortage == false)
                {
                    // if there is not a part shortage on this job then
                    // we need to walk through the list of dependent jobs
                    // and see if there is a part shortage on any of those
                    foreach (HSOpenJob oJob in m_oRelatedJobs)
                    {
                        if (oJob.PartShortage == true)
                        {
                            bShortage = true;
                            break;
                        }
                    }
                }
                return bShortage;
            }
            set { m_bPartShortage = value; }
        }

        public List<HSOpenJob> RelatedJobs
        {
            get { return m_oRelatedJobs; }
        }
        #endregion

        #region Data Members

        private string m_sCompany;
        private string m_sJobNum;
        private string m_sPartNum;
        private string m_sRevNum;
        private string m_sDescription;
        private bool m_bComplete;
        private bool m_bEngineered;
        private bool m_bReleased;
        private bool m_bFirm;
        private decimal m_dProductionQty;
        private DateTime m_dtStartDate;
        private DateTime m_dtDueDate;
        private DateTime m_dtRequiredDate;
        private string m_sCustID;
        private string m_sCustomerName;
        private string m_sTargetJobNum;
        private int m_iTargetAssemblyNum;
        private int m_iTargetMaterialSeq;
        private int m_iOrderNum;
        private int m_iOrderLine;
        private int m_iOrderRelNum;
        private decimal m_dShippedQuantity;
        private decimal m_dReceivedQuantity;
        private DateTime m_dtLastClockInDate;

        // derived attribute
        private bool m_bStartDateNotSet;
        private bool m_bDueDateNotSet;
        private bool m_bRequiredDateNotSet;
        private bool m_bPartShortage;
        private List<HSOpenJob> m_oRelatedJobs = new List<HSOpenJob>();

        #endregion
    }

    public class HSUnfirmJob
    {
        #region Constructors

        public HSUnfirmJob(DataRow oRow)
        {
            if ((oRow["JobHead_Company"] != DBNull.Value) && (string.IsNullOrEmpty((string)oRow["JobHead_Company"]) == false))
            {
                m_sCompany = (string)oRow["JobHead_Company"];
            }
            if ((oRow["JobHead_JobNum"] != DBNull.Value) && (string.IsNullOrEmpty((string)oRow["JobHead_JobNum"]) == false))
            {
                m_sJobNum = (string)oRow["JobHead_JobNum"];
            }
        }
        #endregion

        #region Methods
        #endregion

        #region Properties

        public string Company
        {
            get { return m_sCompany; }
        }
        public string JobNum
        {
            get { return m_sJobNum; }
        }
        #endregion

        #region Data Members
        private string m_sCompany;
        private string m_sJobNum;
        #endregion
    }

    public class HSFixUnfirmJobs
    {
        #region Constructors 
        public HSFixUnfirmJobs(Session oSession)
        {
            // get the list of unfirm jobs
            // get a list of all materials for open jobs
            m_oUnfirmJobs = new List<HSUnfirmJob>();
            Ice.Proxy.BO.DynamicQueryImpl oDynamicQuery = WCFServiceSupport.CreateImpl<Ice.Proxy.BO.DynamicQueryImpl>(oSession, Ice.Proxy.BO.DynamicQueryImpl.UriPath);
            QueryExecutionDataSet oQueryExecutionDataSet = oDynamicQuery.GetQueryExecutionParametersByID(BAQConstants.QUERY_LIST_ALL_UNFIRM_JOBS);
            oQueryExecutionDataSet.Clear();
            oQueryExecutionDataSet.AcceptChanges();
            DataSet oDataSet = oDynamicQuery.ExecuteByID(BAQConstants.QUERY_LIST_ALL_UNFIRM_JOBS, oQueryExecutionDataSet);
            foreach (DataRow oRow in oDataSet.Tables[0].Rows)
            {
                HSUnfirmJob oUnfirmJob = new HSUnfirmJob(oRow);
                m_oUnfirmJobs.Add(oUnfirmJob);
            }
        }
        #endregion

        #region Methods
        public void FirmJobs(Session oSession)
        {
            foreach (HSUnfirmJob oUnfirmJob in m_oUnfirmJobs)
            {
                JobEntryImpl oJobEntryImpl = WCFServiceSupport.CreateImpl<JobEntryImpl>(oSession, Erp.Proxy.BO.JobEntryImpl.UriPath);

                // Get dataset
                Erp.BO.JobEntryDataSet oJobEntryDataSet = oJobEntryImpl.GetByID(oUnfirmJob.JobNum);
                if (oJobEntryDataSet != null)
                {
                    if (oJobEntryDataSet.JobHead.Count == 1)
                    {
                        Erp.BO.JobEntryDataSet.JobHeadRow oJobHeadRow = oJobEntryDataSet.JobHead[0];

                        oJobHeadRow.BeginEdit();
                        oJobHeadRow.JobFirm = true;
                        string sJobComment = oJobHeadRow.CommentText;
                        sJobComment += "\nAuto Firmed";
                        oJobHeadRow.CommentText = sJobComment;
                        oJobHeadRow.RowMod = "U";
                        oJobHeadRow.EndEdit();
                        oJobEntryImpl.Update(oJobEntryDataSet);
                    }
                }
            }

        }

        #endregion

        #region Data Members

        private List<HSUnfirmJob> m_oUnfirmJobs;

        #endregion

    }
}
