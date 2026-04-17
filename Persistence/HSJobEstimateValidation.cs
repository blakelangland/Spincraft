using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Ice.Core;
using Erp.Adapters;
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
using SpreadsheetLight.Charts;
using Ice.Lib.Searches;
using System.IO;


namespace HorizonScientific
{
    public class HSAssembly
    {
        #region Constructors
        public HSAssembly(HSJob oJob, HSAssembly oParentAssembly, int iAssemblySeq, List<JobOperation> oAllJobOperations, List<JobMaterial> oAllJobMaterials, List<JobOpsEstVsActualCosts> oAllOperationCosts)
        {
            m_oJob = oJob;
            m_oParentAssembly = oParentAssembly;
            m_iAssemblySeq = iAssemblySeq;

            // find all materials for this job and assembly seq
            m_oMaterials = oAllJobMaterials.Where(oItem => (oItem.AssemblySeq == m_iAssemblySeq)).ToList();
            // order materials by related operation sequence as this will be when they are consumed and then by material sequence
            m_oMaterials = m_oMaterials.OrderBy(oItem => oItem.RelatedOperation).ThenBy(x => x.MtlSeq).ToList();

            // find all operations for this job and assembly seq
            m_oOperations = oAllJobOperations.Where(oItem => (oItem.AssemblySeq == m_iAssemblySeq)).ToList();

            // find all operation costs for this job and assembly seq
            m_oOperationCosts = oAllOperationCosts.Where(oItem => (oItem.AssemblySeq == m_iAssemblySeq)).ToList();
 
            // now compute all costs for this assembly
            if (m_oOperationCosts != null)
            {
                m_dEstTotalBurdenCosts = m_oOperationCosts.Sum(oItem => oItem.EstBurCost);
                m_dActTotalBurdenCosts = m_oOperationCosts.Sum(oItem => oItem.ActBurdenCost);
                m_dEstTotalLaborCosts = m_oOperationCosts.Sum(oItem => oItem.EstLaborCost);
                m_dActTotalLaborCosts = m_oOperationCosts.Sum(oItem => oItem.ActLaborCost);
                m_dEstTotalSubcontractCosts = m_oOperationCosts.Sum(oItem => oItem.EstSubCost);
                m_dActTotalSubcontractCosts = m_oOperationCosts.Sum(oItem => oItem.ActSubCost);
            }
            if (m_oMaterials != null)
            {
                m_dEstTotalMaterialCosts = m_oMaterials.Sum(oItem => oItem.EstTotalCost);
                m_dActTotalMaterialCosts = m_oMaterials.Sum(oItem => oItem.ActTotalCost);
            }
            m_dEstTotalCost = m_dEstTotalBurdenCosts + m_dEstTotalLaborCosts + m_dEstTotalSubcontractCosts + m_dEstTotalMaterialCosts;
            m_dActTotalCost = m_dActTotalBurdenCosts + m_dActTotalLaborCosts + m_dActTotalSubcontractCosts + m_dActTotalMaterialCosts;

            // now we see if we have any child assemblies
            List<int> oChildAssemblies = new List<int>();
            List<JobMaterial> oChildJobMaterials = new List<JobMaterial>();
            List<JobOpsEstVsActualCosts> oChildJobOps = new List<JobOpsEstVsActualCosts>();
            // if we are the root assembly (seq 0) then we need to eliminate elements with the assembly seq of 0
            if (m_iAssemblySeq == 0)
            {
                oChildJobMaterials = oAllJobMaterials.Where(oItem => (oItem.ParentAssemblySeq == m_iAssemblySeq) && (oItem.AssemblySeq != 0)).ToList();
                oChildJobOps = oAllOperationCosts.Where(oItem => (oItem.ParentAssemblySeq == m_iAssemblySeq) && (oItem.AssemblySeq != 0)).ToList();
            }
            else
            {
                oChildJobMaterials = oAllJobMaterials.Where(oItem => (oItem.ParentAssemblySeq == m_iAssemblySeq)).ToList();
                oChildJobOps = oAllOperationCosts.Where(oItem => (oItem.ParentAssemblySeq == m_iAssemblySeq)).ToList();
            }
            oChildAssemblies.AddRange(oChildJobMaterials.Select(oItem => oItem.AssemblySeq));
            oChildAssemblies.AddRange(oChildJobOps.Select(oItem => oItem.AssemblySeq));
            oChildAssemblies = oChildAssemblies.Distinct().ToList();
            // now we order these by assembly seq
            oChildAssemblies.Sort();
            foreach (int iChildAssemblySeq in oChildAssemblies)
            {
                HSAssembly oChildAssembly = new HSAssembly(m_oJob, this, iChildAssemblySeq, oAllJobOperations, oAllJobMaterials, oAllOperationCosts);
                m_oChildAssemblies.Add(oChildAssembly);
            }

            // no we will set some basic properties on the assembly
            if (m_oMaterials.Count > 0)
            {
                m_sCompany = m_oMaterials[0].Company;
                m_sJobNum = m_oJob.JobNum;
                m_sPartNum = m_oMaterials[0].ParentPartNum;
                m_sPartRevNum = m_oMaterials[0].ParentRevNum;
            }
            else if (m_oOperations.Count > 0)
            {
                m_sCompany = m_oOperations[0].Company;
                m_sJobNum = m_oJob.JobNum;
                m_sPartNum = m_oOperations[0].ParentPartNum;
                m_sPartRevNum = m_oOperations[0].ParentRevNum;
            }
            // check to see if we need a NO-OP for this assembly -- basically we have material not tied to an operation
            List<JobMaterial> oMaterialsWithoutRelatedOperation = m_oMaterials.Where(oItem => oItem.RelatedOperation == 0).ToList();
            JobOperation oNoOp = null;
            JobOpsEstVsActualCosts oNoOpCost = null;
            if (oMaterialsWithoutRelatedOperation.Count > 0)
            {
                oNoOp = new JobOperation(m_oJob, this);
                oNoOpCost = new JobOpsEstVsActualCosts(m_oJob, this);
            }
            if (oNoOp != null)
            {
                // we need to add in a no op
                m_oOperations.Add(oNoOp);
            }
            // order operations by sequence number
            m_oOperations = m_oOperations.OrderBy(oItem => oItem.OperationSeq).ToList();

            if (oNoOpCost != null)
            {
                // we need to add in a no op
                m_oOperationCosts.Add(oNoOpCost);
            }
            // order operation costs by sequence number
            m_oOperationCosts = m_oOperationCosts.OrderBy(oItem => oItem.OprSeq).ToList();
        }
        #endregion

        #region Methods
        public void GetAssembliesInOrder(List<HSAssembly> oOrderedAssemblies)
        {
            // walk through the list of assemblies and put the root assembly last
            foreach (HSAssembly oAssembly in m_oChildAssemblies)
            {
                oAssembly.GetAssembliesInOrder(oOrderedAssemblies);
            }
            // after we walk through all the children assemblies we then add ourselves
            oOrderedAssemblies.Add(this);
        }
        #endregion

        #region Properties
        public HSJob    MyJob
        {
            get { return m_oJob; }
        }
        public int      AssemblySeq
        {
            get { return m_iAssemblySeq; }
        }

        // these are set for convenience from the child materials or ops
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
        public string PartRevNum
        {
            get { return m_sPartRevNum; }
        }
        //

        public List<JobOperation> Operations
        {
            get { return m_oOperations; }
        }
        public List<JobOpsEstVsActualCosts> OperationCosts
        {
            get { return m_oOperationCosts; }
        }
        public List<JobMaterial> Materials
        {
            get { return m_oMaterials; }
        }
        public List<HSAssembly> ChildAssemblies
        {
            get { return m_oChildAssemblies; }
        }
        public HSAssembly   ParentAssembly
        {
            get { return m_oParentAssembly; }
        }

        public decimal EstMaterialCost
        {
            get { return m_dEstTotalMaterialCosts; }
        }
        public decimal ActMaterialCost
        {
            get { return m_dActTotalMaterialCosts; }
        }
        public decimal EstLaborCost
        {
            get { return m_dEstTotalLaborCosts; }
        }
        public decimal ActLaborCost
        {
            get { return m_dActTotalLaborCosts; }
        }
        public decimal EstBurdenCost
        {
            get { return m_dEstTotalBurdenCosts; }
        }
        public decimal ActBurdenCost
        {
            get { return m_dActTotalBurdenCosts; }
        }
        public decimal EstSubcontractCost
        {
            get { return m_dEstTotalSubcontractCosts; }
        }
        public decimal ActSubcontractCost
        {
            get { return m_dActTotalSubcontractCosts; }
        }
        public decimal EstTotaCost
        {
            get { return m_dEstTotalCost; }
        }
        public decimal ActTotaCost
        {
            get { return m_dActTotalCost; }
        }
        #endregion

        #region Data Members
        private List<JobOperation> m_oOperations = new List<JobOperation>();
        private List<JobMaterial> m_oMaterials = new List<JobMaterial>();
        private List<JobOpsEstVsActualCosts> m_oOperationCosts = new List<JobOpsEstVsActualCosts>();

        private HSJob m_oJob;
        int m_iAssemblySeq;
        private HSAssembly m_oParentAssembly;

        private string m_sCompany;
        private string m_sJobNum;
        private string m_sPartNum;
        private string m_sPartRevNum;

        private decimal m_dEstTotalCost;
        private decimal m_dActTotalCost;
        private decimal m_dEstTotalSubcontractCosts;
        private decimal m_dActTotalSubcontractCosts;
        private decimal m_dEstTotalBurdenCosts;
        private decimal m_dActTotalBurdenCosts;
        private decimal m_dEstTotalLaborCosts;
        private decimal m_dActTotalLaborCosts;
        private decimal m_dEstTotalMaterialCosts;
        private decimal m_dActTotalMaterialCosts;

        private List<HSAssembly> m_oChildAssemblies = new List<HSAssembly>();
        #endregion
    }

    public class HSJob
    {
        #region Constructors
        public HSJob(string sJobNum, HSValidateParts oValidateParts, List<JobMaterial> oAllJobMaterials, List<JobOperation> oAllJobOperations, List<JobOpsEstVsActualCosts> oAllJobOperationCosts, JobEstVsActualCostsQty oJobEstVsActualCosts, bool bAcceptActualsForMissingEstimates)
        {
            m_sJobNum = sJobNum;
            if (oAllJobOperationCosts.Count > 0)
            {
                m_iOrderNum = oAllJobOperationCosts[0].OrderNum;
                m_iOrderLine = oAllJobOperationCosts[0].OrderLine;
                m_iOrderRel = oAllJobOperationCosts[0].OrderRel;
                m_sWarehouseCode = oAllJobOperationCosts[0].WarehouseCode;
            }
            else if (oJobEstVsActualCosts != null)
            {
                m_iOrderNum = oJobEstVsActualCosts.OrderNum;
                m_iOrderLine = oJobEstVsActualCosts.OrderLine;
                m_iOrderRel = oJobEstVsActualCosts.OrderRelNum;
                m_sWarehouseCode = oJobEstVsActualCosts.WarehouseCode;
            }

            // get all of our materials, ops, etc.
            m_oJobMaterials = oAllJobMaterials;
            m_oJobOperations = oAllJobOperations;
            m_oJobOperationCosts = oAllJobOperationCosts;
            m_oJobEstVsActualCosts = oJobEstVsActualCosts;

            // now we need to provide a structure as to the ordering of these operations and materials
            // we will order items by assembly and then by operation
            // assembly 0 is always the root 
            m_oRootAssembly = new HSAssembly(this, null, 0, oAllJobOperations, oAllJobMaterials, oAllJobOperationCosts);

            // now compute all costs
            if (m_oJobOperationCosts != null)
            {
                m_dEstTotalBurdenCosts = m_oJobOperationCosts.Sum(oItem => oItem.EstBurCost);
                m_dActTotalBurdenCosts = m_oJobOperationCosts.Sum(oItem => oItem.ActBurdenCost);
                m_dEstTotalLaborCosts = m_oJobOperationCosts.Sum(oItem => oItem.EstLaborCost);
                m_dActTotalLaborCosts = m_oJobOperationCosts.Sum(oItem => oItem.ActLaborCost);
                m_dEstTotalSubcontractCosts = m_oJobOperationCosts.Sum(oItem => oItem.EstSubCost);
                m_dActTotalSubcontractCosts = m_oJobOperationCosts.Sum(oItem => oItem.ActSubCost);
            }
            if (m_oJobMaterials != null)
            {
                m_dEstTotalMaterialCosts = m_oJobMaterials.Sum(oItem => oItem.EstTotalCost);
                m_dActTotalMaterialCosts = m_oJobMaterials.Sum(oItem => oItem.ActTotalCost);
            }
            m_dEstTotalCost = m_dEstTotalBurdenCosts + m_dEstTotalLaborCosts + m_dEstTotalSubcontractCosts + m_dEstTotalMaterialCosts;
            m_dActTotalCost = m_dActTotalBurdenCosts + m_dActTotalLaborCosts + m_dActTotalSubcontractCosts + m_dActTotalMaterialCosts;

            // the following attributes will not be modified regardless of whether we accept
            // actuals for missing estimates or not so we go ahead and establish these values
            if (m_oJobEstVsActualCosts != null)
            {
                m_dUnitPrice = oJobEstVsActualCosts.UnitPrice;
                m_dQty = oJobEstVsActualCosts.OurReqQty;
                m_sProductPortfolioCode = oJobEstVsActualCosts.ProductPortfolio;
                m_dtStartDate = oJobEstVsActualCosts.StartDate;
                m_dtRequiredDate = oJobEstVsActualCosts.RequiredDate;
                m_dtLastLoginDate = oJobEstVsActualCosts.LastClockInDate;
                m_dtDueDate = oJobEstVsActualCosts.DueDate;
                m_sCustomerCode = oJobEstVsActualCosts.CustID;
                m_sCustomerName = oJobEstVsActualCosts.CustomerName;
                m_sMarketSegment = oJobEstVsActualCosts.MarketSegment;
                m_dProductionQty = oJobEstVsActualCosts.ProdQty;
                m_dShippedQty = oJobEstVsActualCosts.ShippedQty;
                m_dReceivedQty = oJobEstVsActualCosts.ReceivedQty;
                m_bOpenRelease = oJobEstVsActualCosts.OpenRelease;
                m_dExtPrice = UnitPrice * Qty;

                m_sPartNum = oJobEstVsActualCosts.PartNum;
                m_sPartRev = oJobEstVsActualCosts.PartRevNum;
                m_oPartMaster = oValidateParts.GetPart(m_sPartNum);
            }

            // now compute the estimated margin and profit from all operations and materials not yet completed
            m_dRemainingCostsEstimated = 0M;

            // this will hold what will be the total costs of the job according to the estimates for ops and parts
            decimal dPercentEstTotalCost = 0M;

            // establish the part and rev numbers -- also set remaining est cost for operations
            if ((m_oJobOperationCosts != null) && (m_oJobOperationCosts.Count > 0))
            {
                foreach (JobOpsEstVsActualCosts oOperationCost in m_oJobOperationCosts)
                {
                    m_dRemainingCostsEstimated += oOperationCost.EstimatedRemainingCost;
                    dPercentEstTotalCost += oOperationCost.EstCost * oOperationCost.PercentComplete;
                }
            }
            // establish the part and rev numbers -- also set remaining est cost for materials
            if ((m_oJobMaterials != null) && (m_oJobMaterials.Count > 0))
            {
                foreach (JobMaterial oMaterial in m_oJobMaterials)
                {
                    m_dRemainingCostsEstimated += oMaterial.EstRemainingCost;
                    dPercentEstTotalCost += oMaterial.EstTotalCost * oMaterial.PercentComplete;
                }
            }

            m_dProfitEstimated = ExtPrice - m_oJobEstVsActualCosts.ActTotalCost - m_dRemainingCostsEstimated;
            if (ExtPrice != 0)
            {
                m_dMarginEstimated = m_dProfitEstimated / ExtPrice;
            }
            else
            {
                m_dMarginEstimated = 0M;
            }

            // determine if we have any transactions against this job yet
            if (m_dActTotalCost > 0)
            {
                m_bHasTransactions = true;
            }

            if (m_oJobEstVsActualCosts != null)
            {
                if (bAcceptActualsForMissingEstimates == true)
                {
                    // we need to recompute the percentage complete
                    m_dPercentComplete = 0M;
                    if (m_dEstTotalCost != 0)
                    {
                        m_dPercentComplete = dPercentEstTotalCost / m_dEstTotalCost;
                    }
                    if (m_dPercentComplete > 1)
                    {
                        m_dPercentComplete = 1.0M;
                    }
                    m_oJobEstVsActualCosts.PercentageComplete = m_dPercentComplete;

                    // we need to recompute the projected total cost -- as this did not have any 
                    if (m_dPercentComplete > 0)
                    {
                        m_oJobEstVsActualCosts.ProjectedTotalCost = ActTotalCost / m_dPercentComplete;
                    }
                    else
                    {
                        // we use estimated total cost
                        m_oJobEstVsActualCosts.ProjectedTotalCost = m_dEstTotalCost;
                    }

                    // reset the projected costs based on the recalculation of the projected total cost
                    m_dRemainingCostsProjected = m_oJobEstVsActualCosts.ProjectedTotalCost - m_oJobEstVsActualCosts.ActTotalCost;

                    // we need to recompute margin
                    if (ExtPrice * PercentComplete > 0)
                    {
                        m_dMarginToDate = ((ExtPrice * PercentComplete) - ActTotalCost) / (ExtPrice * PercentComplete);
                    }
                    else
                    {
                        m_dMarginToDate = 0M;
                    }
                    m_oJobEstVsActualCosts.MarginToDate = m_dMarginToDate;

                    // we need to recompute the projected profit
                    if (PercentComplete > 0)
                    {
                        m_oJobEstVsActualCosts.ProjectedTotalCost = ActTotalCost / PercentComplete;
                    }
                    else
                    {
                        m_oJobEstVsActualCosts.ProjectedTotalCost = EstTotalCost;
                    }
                    m_dProfitProjected = ExtPrice - m_oJobEstVsActualCosts.ProjectedTotalCost;

                    // we need to recompute margin
                    m_dMarginProjected = m_dMarginToDate;
                }
                else
                {
                    // no modification necessary from original query
                    m_dMarginToDate = m_oJobEstVsActualCosts.MarginToDate;
                    m_dPercentComplete = m_oJobEstVsActualCosts.PercentageComplete;
                    m_dRemainingCostsProjected = m_oJobEstVsActualCosts.ProjectedTotalCost - m_oJobEstVsActualCosts.ActTotalCost;
                    m_dProfitProjected = ExtPrice - m_oJobEstVsActualCosts.ProjectedTotalCost;
                    m_dMarginProjected = m_dMarginToDate;
                }
            }
        }

        #endregion

        #region Methods
        public List<HSAssembly> GetAssembliesInOrder()
        {
            List<HSAssembly> oOrderedAssemblies = new List<HSAssembly>();
            m_oRootAssembly.GetAssembliesInOrder(oOrderedAssemblies);
            return oOrderedAssemblies;
        }
        #endregion

        #region Properties

        public string JobNum
        {
            get { return m_sJobNum; }
            set { m_sJobNum = value; }
        }
        public int OrderNum
        {
            get { return m_iOrderNum; }
            set { m_iOrderNum = value; }
        }

        public int OrderLine
        {
            get { return m_iOrderLine; }
            set { m_iOrderLine = value; }
        }
        public int OrderRel
        {
            get { return m_iOrderRel; }
            set { m_iOrderRel = value; }
        }
        public string CustomerCode
        {
            get { return m_sCustomerCode; }
            set { m_sCustomerCode = value; }
        }
        public string CustomerName
        {
            get { return m_sCustomerName; }
            set { m_sCustomerName = value; }
        }
        public string MarketSegment
        {
            get { return m_sMarketSegment; }
            set { m_sMarketSegment = value; }
        }
        public decimal UnitPrice
        {
            get { return m_dUnitPrice; }
            set { m_dUnitPrice = value; }
        }
        public decimal Qty
        {
            get { return m_dQty; }
            set { m_dQty = value; }
        }
        public decimal ExtPrice
        {
            get { return m_dExtPrice; }
            set { m_dExtPrice = value; }
        }
        public string ProductPortfolioCode
        {
            get { return m_sProductPortfolioCode; }
            set { m_sProductPortfolioCode = value; }
        }
        public DateTime StartDate
        {
            get { return m_dtStartDate; }
            set { m_dtStartDate = value; }
        }
        public DateTime LastLoginDate
        {
            get { return m_dtLastLoginDate; }
            set { m_dtLastLoginDate = value; }
        }
        public DateTime DueDate
        {
            get { return m_dtDueDate; }
            set { m_dtDueDate = value; }
        }
        public DateTime RequiredDate
        {
            get { return m_dtRequiredDate; }
            set { m_dtRequiredDate = value; }
        }
        public string WarehouseCode
        {
            get { return m_sWarehouseCode; }
            set { m_sWarehouseCode = value; }
        }
        public string PartNum
        {
            get { return m_sPartNum; }
            set { m_sPartNum = value; }
        }
        public string PartRevNum
        {
            get { return m_sPartRev; }
            set { m_sPartRev = value; }
        }
        public HSPartData PartMaster
        {
            get { return m_oPartMaster; }
        }
        public bool HasTransactions
        {
            get { return m_bHasTransactions; }
        }

        public decimal ProductionQty
        {
            get { return m_dProductionQty; }
            set { m_dProductionQty = value; }
        }

        public decimal ShippedQty
        {
            get { return m_dShippedQty; }
            set { m_dShippedQty = value; }
        }

        public decimal ReceivedQty
        {
            get { return m_dReceivedQty; }
            set { m_dReceivedQty = value; }
        }
        public bool OpenRelease
        {
            get { return m_bOpenRelease; }
            set { m_bOpenRelease = value; }
        }
        public HSAssembly RootAssembly
        {
            get { return m_oRootAssembly; }
            set { m_oRootAssembly = value; }
        }

        public List<JobMaterial> AllMaterials
        {
            get { return m_oJobMaterials; }
        }
        public List<JobOperation> AllOperations
        {
            get { return m_oJobOperations; }
        }
        public List<JobOpsEstVsActualCosts> AllOperationCosts
        {
            get { return m_oJobOperationCosts; }
        }

        public decimal EstMaterialCost
        {
            get { return m_dEstTotalMaterialCosts; }
        }
        public decimal ActMaterialCost
        {
            get { return m_dActTotalMaterialCosts; }
        }
        public decimal EstLaborCost
        {
            get { return m_dEstTotalLaborCosts; }
        }
        public decimal ActLaborCost
        {
            get { return m_dActTotalLaborCosts; }
        }
        public decimal EstBurdenCost
        {
            get { return m_dEstTotalBurdenCosts; }
        }
        public decimal ActBurdenCost
        {
            get { return m_dActTotalBurdenCosts; }
        }
        public decimal EstSubcontractCost
        {
            get { return m_dEstTotalSubcontractCosts; }
        }
        public decimal ActSubcontractCost
        {
            get { return m_dActTotalSubcontractCosts; }
        }
        public decimal EstTotalCost
        {
            get { return m_dEstTotalCost; }
        }
        public decimal ActTotalCost
        {
            get { return m_dActTotalCost; }
        }

        public decimal MarginToDate
        {
            get { return m_dMarginToDate; }
            set { m_dMarginToDate = value; }
        }
        public decimal PercentComplete
        {
            get { return m_dPercentComplete; }
            set { m_dPercentComplete = value; }
        }
        public decimal RemainingCostsEstimated
        {
            get { return m_dRemainingCostsEstimated; }
            set { m_dRemainingCostsEstimated = value; }
        }
        public decimal RemainingCostsProjected
        {
            get { return m_dRemainingCostsProjected; }
            set { m_dRemainingCostsProjected = value; }
        }
        public decimal ProfitEstimated
        {
            get { return m_dProfitEstimated; }
            set { m_dProfitEstimated = value; }
        }
        public decimal ProfitProjected
        {
            get { return m_dProfitProjected; }
            set { m_dProfitProjected = value; }
        }
        public decimal MarginEstimated
        {
            get { return m_dMarginEstimated; }
            set { m_dMarginEstimated = value; }
        }
        public decimal MarginProjected
        {
            get { return m_dMarginProjected; }
            set { m_dMarginProjected = value; }
        }
        #endregion

        #region Data Members
        private string m_sJobNum;
        private int m_iOrderNum;
        private int m_iOrderLine;
        private int m_iOrderRel;
        private string m_sCustomerCode;
        private string m_sCustomerName;
        private string m_sMarketSegment;
        private decimal m_dUnitPrice;
        private decimal m_dQty;
        private decimal m_dExtPrice;
        private string m_sProductPortfolioCode;
        private DateTime m_dtStartDate;
        private DateTime m_dtLastLoginDate;
        private DateTime m_dtDueDate;
        private DateTime m_dtRequiredDate;
        private string m_sWarehouseCode;
        private string m_sPartNum;
        private string m_sPartRev;
        private HSPartData m_oPartMaster;
        private bool m_bHasTransactions;
        private decimal m_dProductionQty;
        private decimal m_dShippedQty;
        private decimal m_dReceivedQty;
        private bool m_bOpenRelease;

        private HSAssembly m_oRootAssembly;

        private List<JobMaterial> m_oJobMaterials = new List<JobMaterial>();
        private List<JobOperation> m_oJobOperations = new List<JobOperation>();
        private List<JobOpsEstVsActualCosts> m_oJobOperationCosts = new List<JobOpsEstVsActualCosts>();
        private JobEstVsActualCostsQty m_oJobEstVsActualCosts = new JobEstVsActualCostsQty();

        private decimal m_dEstTotalCost;
        private decimal m_dActTotalCost;
        private decimal m_dEstTotalSubcontractCosts;
        private decimal m_dActTotalSubcontractCosts;
        private decimal m_dEstTotalBurdenCosts;
        private decimal m_dActTotalBurdenCosts;
        private decimal m_dEstTotalLaborCosts;
        private decimal m_dActTotalLaborCosts;
        private decimal m_dEstTotalMaterialCosts;
        private decimal m_dActTotalMaterialCosts;

        private decimal m_dMarginToDate;
        private decimal m_dPercentComplete;
        private decimal m_dRemainingCostsEstimated;
        private decimal m_dRemainingCostsProjected;
        private decimal m_dProfitEstimated;
        private decimal m_dProfitProjected;
        private decimal m_dMarginEstimated;
        private decimal m_dMarginProjected;

        #endregion
    }

    public class HSJobEstimateValidation
    {
        #region constructors
        public HSJobEstimateValidation(string sCompany, decimal dJobEstimateAbsoluteError, decimal dJobEstimatePercentError, decimal dJobMarginThreshold, bool bForceJobAnalysis, bool bAcceptActualsForMissingEstimates, bool bJustMissingCosts)
        {
            m_sCompany = sCompany;
            m_dJobEstimateAbsoluteError = dJobEstimateAbsoluteError;
            if (m_dJobEstimateAbsoluteError == 0)
            {
                m_dJobEstimateAbsoluteError = MIN_ABSOLUTE_ERROR;
            }
            m_dJobEstimatePercentError = dJobEstimatePercentError;
            if (m_dJobEstimatePercentError == 0)
            {
                m_dJobEstimatePercentError = MIN_ABSOLUTE_PERCENTAGE_ERROR;
            }
            m_dJobMarginThreshold = dJobMarginThreshold;
            if (m_dJobMarginThreshold == 0)
            {
                m_dJobMarginThreshold = MIN_MARGIN_THRESHOLD;
            }
            m_bForceJobAnalysis = bForceJobAnalysis;
            m_bAcceptActualsForMissingEstimates = bAcceptActualsForMissingEstimates;
            m_bJustMissingCosts = bJustMissingCosts;

            if (string.Compare(m_sCompany, CompanyConfiguration.SPINCRAFT_WI_COMPANY_ID, true) == 0)
            {
                // Wisconsin ignores parts in the following classes
                m_oPartClassesToIgnore.Add("CATL");
                m_oPartClassesToIgnore.Add("COTL");
                m_oPartClassesToIgnore.Add("ENGD");
                m_oPartClassesToIgnore.Add("FA");
                m_oPartClassesToIgnore.Add("GOVT");
                m_oPartClassesToIgnore.Add("INSP");
                m_oPartClassesToIgnore.Add("LTAT");
                m_oPartClassesToIgnore.Add("MFG");
                m_oPartClassesToIgnore.Add("PAIN");
                m_oPartClassesToIgnore.Add("PUR");
                m_oPartClassesToIgnore.Add("SA");
                m_oPartClassesToIgnore.Add("SHIP");
                m_oPartClassesToIgnore.Add("SPNS");
                m_oPartClassesToIgnore.Add("SPTL");
                m_oPartClassesToIgnore.Add("SUPL");
                m_oPartClassesToIgnore.Add("WELD");
            }
        }

        #endregion

        #region Methods
        public void Initialize(Session oSession, HSValidateParts oValidateParts, string sJobNum, string sPartNum)
        {
            // loading all parts from the part master
            if (oValidateParts == null)
            {
                if (m_oValidateParts.Initialize(oSession) == false)
                {
                    Console.WriteLine("Failed to load the validate parts!");
                }
                oValidateParts = m_oValidateParts;
            }
            else
            {
                m_oValidateParts = oValidateParts;
            }

            // get all resources
            m_oBOMSupport = new BOMSupport(m_sCompany);
            if (m_oBOMSupport.InitializeResourceGroups(oSession) == false)
            {
                Console.WriteLine("Failed to load the BOM Support object!");
            }

            // get all resouce groups
            if (m_oBOMSupport.InitializeResources(oSession) == false)
            {
                Console.WriteLine("Failed to load the BOM Support object!");
            }

            // get a list of all materials for open jobs
            m_oJobMaterials.Clear();
            Ice.Proxy.BO.DynamicQueryImpl oDynamicQuery = WCFServiceSupport.CreateImpl<Ice.Proxy.BO.DynamicQueryImpl>(oSession, Ice.Proxy.BO.DynamicQueryImpl.UriPath);
            QueryExecutionDataSet oQueryExecutionDataSet = oDynamicQuery.GetQueryExecutionParametersByID(BAQConstants.QUERY_LIST_ALL_JOB_MATERIALS);
            oQueryExecutionDataSet.Clear();
            if (string.IsNullOrEmpty(sJobNum) == false)
            {
                oQueryExecutionDataSet.ExecutionParameter.AddExecutionParameterRow("JobNum", sJobNum, "nvarchar", false, Guid.NewGuid(), "A");
            }
            if (string.IsNullOrEmpty(sPartNum) == false)
            {
                oQueryExecutionDataSet.ExecutionParameter.AddExecutionParameterRow("PartNum", sPartNum, "nvarchar", false, Guid.NewGuid(), "A");
            }
            oQueryExecutionDataSet.AcceptChanges();
            DataSet oDataSet = oDynamicQuery.ExecuteByID(BAQConstants.QUERY_LIST_ALL_JOB_MATERIALS, oQueryExecutionDataSet);
            foreach (DataRow oRow in oDataSet.Tables[0].Rows)
            {
                JobMaterial oJobMaterial = new JobMaterial(oRow, m_oValidateParts, m_bAcceptActualsForMissingEstimates);
                m_oJobMaterials.Add(oJobMaterial);
                string sTmpJobNum = oJobMaterial.JobNum;
                List<JobMaterial> oTmpJobMaterials = null;
                if (m_oFastJobMaterials.ContainsKey(sTmpJobNum) == true)
                {
                    oTmpJobMaterials = m_oFastJobMaterials[sTmpJobNum];
                }
                else
                {
                    oTmpJobMaterials = new List<JobMaterial>();
                    m_oFastJobMaterials[sTmpJobNum] = oTmpJobMaterials;
                }
                oTmpJobMaterials.Add(oJobMaterial);
            }

            m_oJobOperations.Clear();
            oDynamicQuery = WCFServiceSupport.CreateImpl<Ice.Proxy.BO.DynamicQueryImpl>(oSession, Ice.Proxy.BO.DynamicQueryImpl.UriPath);
            oQueryExecutionDataSet = oDynamicQuery.GetQueryExecutionParametersByID(BAQConstants.QUERY_LIST_ALL_JOB_OPERATIONS);
            oQueryExecutionDataSet.Clear();
            if (string.IsNullOrEmpty(sJobNum) == false)
            {
                oQueryExecutionDataSet.ExecutionParameter.AddExecutionParameterRow("JobNum", sJobNum, "nvarchar", false, Guid.NewGuid(), "A");
            }
            if (string.IsNullOrEmpty(sPartNum) == false)
            {
                oQueryExecutionDataSet.ExecutionParameter.AddExecutionParameterRow("PartNum", sPartNum, "nvarchar", false, Guid.NewGuid(), "A");
            }
            oQueryExecutionDataSet.AcceptChanges();
            oDataSet = oDynamicQuery.ExecuteByID(BAQConstants.QUERY_LIST_ALL_JOB_OPERATIONS, oQueryExecutionDataSet);
            foreach (DataRow oRow in oDataSet.Tables[0].Rows)
            {
                JobOperation oJobOperation = new JobOperation(oRow);
                m_oJobOperations.Add(oJobOperation);

                string sTmpJobNum = oJobOperation.JobNum;
                List<JobOperation> oTmpJobOperations = null;
                if (m_oFastJobOperations.ContainsKey(sTmpJobNum) == true)
                {
                    oTmpJobOperations = m_oFastJobOperations[sTmpJobNum];
                }
                else
                {
                    oTmpJobOperations = new List<JobOperation>();
                    m_oFastJobOperations[sTmpJobNum] = oTmpJobOperations;
                }
                oTmpJobOperations.Add(oJobOperation);
            }

            // pull in POC estimate vs actual costs for all open jobs
            m_oJobOpsEstVsActualCosts.Clear();
            oDynamicQuery = WCFServiceSupport.CreateImpl<Ice.Proxy.BO.DynamicQueryImpl>(oSession, Ice.Proxy.BO.DynamicQueryImpl.UriPath);
            oQueryExecutionDataSet = oDynamicQuery.GetQueryExecutionParametersByID(BAQConstants.QUERY_JOB_OPS_EST_VS_ACTUAL_COSTS);
            oQueryExecutionDataSet.Clear();
            if (string.IsNullOrEmpty(sJobNum) == false)
            {
                oQueryExecutionDataSet.ExecutionParameter.AddExecutionParameterRow("JobNum", sJobNum, "nvarchar", false, Guid.NewGuid(), "A");
            }
            if (string.IsNullOrEmpty(sPartNum) == false)
            {
                oQueryExecutionDataSet.ExecutionParameter.AddExecutionParameterRow("PartNum", sPartNum, "nvarchar", false, Guid.NewGuid(), "A");
            }
            oQueryExecutionDataSet.AcceptChanges();
            oDataSet = oDynamicQuery.ExecuteByID(BAQConstants.QUERY_JOB_OPS_EST_VS_ACTUAL_COSTS, oQueryExecutionDataSet);
            foreach (DataRow oRow in oDataSet.Tables[0].Rows)
            {
                JobOpsEstVsActualCosts oJobEstVsActualCosts = new JobOpsEstVsActualCosts(oRow, m_bAcceptActualsForMissingEstimates);
                m_oJobOpsEstVsActualCosts.Add(oJobEstVsActualCosts);

                string sTmpJobNum = oJobEstVsActualCosts.JobNum;
                List<JobOpsEstVsActualCosts> oTmpJobOperationCosts = null;
                if (m_oFastJobOperationCosts.ContainsKey(sTmpJobNum) == true)
                {
                    oTmpJobOperationCosts = m_oFastJobOperationCosts[sTmpJobNum];
                }
                else
                {
                    oTmpJobOperationCosts = new List<JobOpsEstVsActualCosts>();
                    m_oFastJobOperationCosts[sTmpJobNum] = oTmpJobOperationCosts;
                }
                oTmpJobOperationCosts.Add(oJobEstVsActualCosts);
            }

            m_oJobEstVsActualCostsQtys.Clear();
            oDynamicQuery = WCFServiceSupport.CreateImpl<Ice.Proxy.BO.DynamicQueryImpl>(oSession, Ice.Proxy.BO.DynamicQueryImpl.UriPath);
            oQueryExecutionDataSet = oDynamicQuery.GetQueryExecutionParametersByID(BAQConstants.QUERY_JOB_EST_VS_ACTUAL_COSTS_QTY);
            oQueryExecutionDataSet.Clear();
            if (string.IsNullOrEmpty(sJobNum) == false)
            {
                oQueryExecutionDataSet.ExecutionParameter.AddExecutionParameterRow("JobNum", sJobNum, "nvarchar", false, Guid.NewGuid(), "A");
            }
            if (string.IsNullOrEmpty(sPartNum) == false)
            {
                oQueryExecutionDataSet.ExecutionParameter.AddExecutionParameterRow("PartNum", sPartNum, "nvarchar", false, Guid.NewGuid(), "A");
            }
            oQueryExecutionDataSet.AcceptChanges();
            oDataSet = oDynamicQuery.ExecuteByID(BAQConstants.QUERY_JOB_EST_VS_ACTUAL_COSTS_QTY, oQueryExecutionDataSet);
            foreach (DataRow oRow in oDataSet.Tables[0].Rows)
            {
                JobEstVsActualCostsQty oJobEstVsActualCostsQty = new JobEstVsActualCostsQty(oRow);
                m_oJobEstVsActualCostsQtys.Add(oJobEstVsActualCostsQty);

                string sTmpJobNum = oJobEstVsActualCostsQty.JobNum;
                if (m_oFastJobEstVsActualCosts.ContainsKey(sTmpJobNum) == false)
                {
                    m_oFastJobEstVsActualCosts[sTmpJobNum] = oJobEstVsActualCostsQty;
                }
            }


            m_sPartNum = sPartNum;

            // we need to get the unique list of job numbers
            List<string> oAllJobNums = new List<string>();
            List<string> oAllJobsForOperations = m_oJobOperations.Select(oItem => oItem.JobNum).ToList();
            oAllJobNums.AddRange(oAllJobsForOperations);
            List<string> oAllJobsForMaterials = m_oJobMaterials.Select(oItem => oItem.JobNum).ToList();
            oAllJobNums.AddRange(oAllJobsForMaterials);
            List<string> oAllJobsForOpCosts = m_oJobOpsEstVsActualCosts.Select(oItem => oItem.JobNum).ToList();
            oAllJobNums.AddRange(oAllJobsForOpCosts);
            List<string> oAllJobEstVsActualCosts = m_oJobEstVsActualCostsQtys.Select(oItem => oItem.JobNum).ToList();
            oAllJobNums.AddRange(oAllJobEstVsActualCosts);

            // get rid of all duplicate job numbers
            m_oJobNums = oAllJobNums.Distinct().ToList();

            // now we create a list of jobs
            foreach (string sTmpJobNum in m_oJobNums)
            {
                // get the list of materials and ops specific to this job
                List<JobMaterial> oTmpJobMaterials = new List<JobMaterial>();
                if (m_oFastJobMaterials.ContainsKey(sTmpJobNum) == true)
                {
                    oTmpJobMaterials = m_oFastJobMaterials[sTmpJobNum];
                }

                List<JobOperation> oTmpJobOperations = new List<JobOperation>();
                if (m_oFastJobOperations.ContainsKey(sTmpJobNum) == true)
                {
                    oTmpJobOperations = m_oFastJobOperations[sTmpJobNum];
                }

                List<JobOpsEstVsActualCosts> oTmpJobOperationCosts = new List<JobOpsEstVsActualCosts>();
                if (m_oFastJobOperationCosts.ContainsKey(sTmpJobNum) == true)
                {
                    oTmpJobOperationCosts = m_oFastJobOperationCosts[sTmpJobNum];
                }

                JobEstVsActualCostsQty oTmpJobEstVsActualCosts = new JobEstVsActualCostsQty();
                if (m_oFastJobEstVsActualCosts.ContainsKey(sTmpJobNum) == true)
                {
                    oTmpJobEstVsActualCosts = m_oFastJobEstVsActualCosts[sTmpJobNum];
                }

                HSJob oJob = new HSJob(sTmpJobNum, m_oValidateParts, oTmpJobMaterials, oTmpJobOperations, oTmpJobOperationCosts, oTmpJobEstVsActualCosts, m_bAcceptActualsForMissingEstimates);
                m_oAllJobs.Add(oJob);
            }
        }

        public void PerformJobValidation(string sCompany, string sTmpFileDirectory, HSUser oRequestingUser)
        {
            #region Setup

            // get the file name
            DateTime dtToday = DateTime.Now;
            string sDate = dtToday.Month.ToString() + "-" + dtToday.Day.ToString() + "-" + dtToday.Year.ToString();
            string sDestinationFileName = sTmpFileDirectory + "\\" + sCompany + "-JobValidationReport-" + sDate + ".xlsx";
            int iNumberOfRows = 1;
            int iNumberOfColumns = 1;

            HSUser oServiceAccount = HSUser.GetUserById(HSUser.SPINCRAFT_SERVICE_ACCOUNT_ID);
            List<string> oToAddresses = new List<string>();
            if (oRequestingUser == null)
            {
                // get users in the engineering, production, and quoting groups
                HSUser.GetEmailsForUsersInGroup(oToAddresses, HSUser.REPORT_ON_JOB_ESTIMATES);
            }
            else
            {
                // this report was requested by someone directly from Epicor
                oToAddresses.Add(oRequestingUser.Email);
            }
            oToAddresses.Add(HSUser.SPINCRAFT_ROOT_USER);

            SLFill oSLFill = new SLFill();
            oSLFill.SetPatternBackgroundColor(SLThemeColorIndexValues.Accent1Color);
            oSLFill.SetPattern(PatternValues.Solid, SLThemeColorIndexValues.Light2Color, SLThemeColorIndexValues.Light2Color);

            SLAlignment oCenterAlignment = new SLAlignment();
            oCenterAlignment.Horizontal = HorizontalAlignmentValues.Center;

            SLAlignment oRightAlignment = new SLAlignment();
            oRightAlignment.Horizontal = HorizontalAlignmentValues.Right;

            SLAlignment oLeftAlignment = new SLAlignment();
            oLeftAlignment.Horizontal = HorizontalAlignmentValues.Left;

            SLStyle oBoldStyle = new SLStyle();
            oBoldStyle.SetFontBold(true);
            oBoldStyle.SetTopBorder(BorderStyleValues.Thin, SLThemeColorIndexValues.Dark2Color);
            oBoldStyle.SetBottomBorder(BorderStyleValues.Thin, SLThemeColorIndexValues.Dark2Color);
            oBoldStyle.SetLeftBorder(BorderStyleValues.Thin, SLThemeColorIndexValues.Dark2Color);
            oBoldStyle.SetRightBorder(BorderStyleValues.Thin, SLThemeColorIndexValues.Dark2Color);
            oBoldStyle.SetFontColor(SLThemeColorIndexValues.Dark2Color);

            SLStyle oBoldStyle2 = new SLStyle();
            oBoldStyle2.SetFontBold(true);
            oBoldStyle2.SetTopBorder(BorderStyleValues.Thick, SLThemeColorIndexValues.Dark2Color);
            oBoldStyle2.SetBottomBorder(BorderStyleValues.Thick, SLThemeColorIndexValues.Dark2Color);
            oBoldStyle2.SetLeftBorder(BorderStyleValues.Thick, SLThemeColorIndexValues.Dark2Color);
            oBoldStyle.SetRightBorder(BorderStyleValues.Thick, SLThemeColorIndexValues.Dark2Color);
            oBoldStyle2.SetFontColor(System.Drawing.Color.IndianRed);

            SLStyle oCurrencyStyle = new SLStyle();
            oCurrencyStyle.ApplyNamedCellStyle(SLNamedCellStyleValues.Currency);
            oCurrencyStyle.Alignment = oLeftAlignment;
            oCurrencyStyle.FormatCode = "$#,##0.00";

            SLStyle oDecimalStyle = new SLStyle();
            oDecimalStyle.Alignment = oLeftAlignment;
            oDecimalStyle.FormatCode = "###.00";

            System.Drawing.Color oHighlightYellow = System.Drawing.Color.FromArgb(255, 255, 0);

            SLStyle oHighlightHeaderStyle = new SLStyle();
            oHighlightHeaderStyle.SetFontBold(true);
            oHighlightHeaderStyle.SetFont(FontSchemeValues.Major, 12);
            oHighlightHeaderStyle.SetHorizontalAlignment(HorizontalAlignmentValues.Center);
            oHighlightHeaderStyle.SetPatternFill(PatternValues.Solid, oHighlightYellow, oHighlightYellow);
            oHighlightHeaderStyle.SetFontColor(SLThemeColorIndexValues.Dark2Color);

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
            SLDocument oSLBOMDocument = new SLDocument();

            bool bDataInReport = false;
            bool bFirstWorksheet = true;

            int iStandardColumnWidth = 20;
            #endregion

            #region Job Issues
            // jobs with zero qty
            List<HSJob> oJobsWithNoProductionQty = m_oAllJobs.Where(oItem => oItem.ProductionQty == 0).ToList();
            if (oJobsWithNoProductionQty.Count > 0)
            {
                // sort these by job number
                oJobsWithNoProductionQty = oJobsWithNoProductionQty.OrderBy(oItem => oItem.JobNum).ToList();
                iNumberOfRows = 1;
                iNumberOfColumns = 1;
                if (bFirstWorksheet == true)
                {
                    oSLBOMDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "Jobs Without Production Qty");
                    bFirstWorksheet = false;
                }
                else
                {
                    oSLBOMDocument.AddWorksheet("Jobs Without Production Qty");
                }
                //set column header
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "JobNum");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Part Num");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Part Rev Num");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Jobs Should Have A Production Qty -- SOP REQUIRES JOBS TO HAVE PRODUCTION QTY");
                oSLBOMDocument.SetCellStyle(iNumberOfRows++, iNumberOfColumns, oHighlightHeaderStyle);
                oSLBOMDocument.SetColumnWidth(iNumberOfColumns++, 90);

                foreach (HSJob oJob in oJobsWithNoProductionQty)
                {
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 1, oJob.JobNum);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 2, oJob.PartNum);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 3, oJob.PartRevNum);

                    iNumberOfRows++;
                    bDataInReport = true;
                }
            }

            // jobs missing due date
            List<HSJob> oJobsMissingDueDate = m_oAllJobs.Where(oItem => oItem.DueDate == DateTime.MinValue).ToList();
            if (oJobsMissingDueDate.Count > 0)
            {
                // sort these by job number
                oJobsMissingDueDate = oJobsMissingDueDate.OrderBy(oItem => oItem.JobNum).ToList();
                iNumberOfRows = 1;
                iNumberOfColumns = 1;
                if (bFirstWorksheet == true)
                {
                    oSLBOMDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "Jobs Without Due Date");
                    bFirstWorksheet = false;
                }
                else
                {
                    oSLBOMDocument.AddWorksheet("Jobs Without Due Date");
                }
                //set column header
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "JobNum");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Part Num");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Part Rev Num");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Jobs Should Have Due Date Set For Planning and Purchasing Purposes -- SOP REQUIRES DUE DATE TO BE SET");
                oSLBOMDocument.SetCellStyle(iNumberOfRows++, iNumberOfColumns, oHighlightHeaderStyle);
                oSLBOMDocument.SetColumnWidth(iNumberOfColumns++, 90);

                foreach (HSJob oJob in oJobsMissingDueDate)
                {
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 1, oJob.JobNum);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 2, oJob.PartNum);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 3, oJob.PartRevNum);

                    iNumberOfRows++;
                    bDataInReport = true;
                }
            }

            // jobs missing required by date
            List<HSJob> oJobsMissingRequiredByDate = m_oAllJobs.Where(oItem => oItem.RequiredDate == DateTime.MinValue).ToList();
            if (oJobsMissingRequiredByDate.Count > 0)
            {
                // sort these by job number
                oJobsMissingRequiredByDate = oJobsMissingRequiredByDate.OrderBy(oItem => oItem.JobNum).ToList();
                iNumberOfRows = 1;
                iNumberOfColumns = 1;
                if (bFirstWorksheet == true)
                {
                    oSLBOMDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "Jobs Without Required Date");
                    bFirstWorksheet = false;
                }
                else
                {
                    oSLBOMDocument.AddWorksheet("Jobs Without Required Date");
                }
                //set column header
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "JobNum");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Part Num");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Part Rev Num");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Jobs Should Have Required By Date Set For Planning and Purchasing Purposes -- SOP STATES REQUIRED DATE SHOULD BE SET");
                oSLBOMDocument.SetCellStyle(iNumberOfRows++, iNumberOfColumns, oHighlightHeaderStyle);
                oSLBOMDocument.SetColumnWidth(iNumberOfColumns++, 90);

                foreach (HSJob oJob in oJobsMissingRequiredByDate)
                {
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 1, oJob.JobNum);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 2, oJob.PartNum);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 3, oJob.PartRevNum);

                    iNumberOfRows++;
                    bDataInReport = true;
                }
            }

            // jobs missing start date
            List<HSJob> oJobsMissingStartDate = m_oAllJobs.Where(oItem => oItem.StartDate == DateTime.MinValue).ToList();
            if (oJobsMissingRequiredByDate.Count > 0)
            {
                // sort these by job number
                oJobsMissingStartDate = oJobsMissingStartDate.OrderBy(oItem => oItem.JobNum).ToList();
                iNumberOfRows = 1;
                iNumberOfColumns = 1;
                if (bFirstWorksheet == true)
                {
                    oSLBOMDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "Jobs Without Start Date");
                    bFirstWorksheet = false;
                }
                else
                {
                    oSLBOMDocument.AddWorksheet("Jobs Without Start Date");
                }
                //set column header
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "JobNum");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Part Num");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Part Rev Num");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Jobs Should Have Start Date Set For Planning and Purchasing Purposes -- SOP REQUIRES START DATE TO BE SET");
                oSLBOMDocument.SetCellStyle(iNumberOfRows++, iNumberOfColumns, oHighlightHeaderStyle);
                oSLBOMDocument.SetColumnWidth(iNumberOfColumns++, 90);

                foreach (HSJob oJob in oJobsMissingStartDate)
                {
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 1, oJob.JobNum);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 2, oJob.PartNum);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 3, oJob.PartRevNum);

                    iNumberOfRows++;
                    bDataInReport = true;
                }
            }

            // jobs that have parts on the fly
            List<HSJob> oJobsWithoutPartMaster = m_oAllJobs.Where(oItem => (oItem.PartMaster == null)).ToList();
            if (oJobsWithoutPartMaster.Count > 0)
            {
                // sort these by job number
                oJobsWithoutPartMaster = oJobsWithoutPartMaster.OrderBy(oItem => oItem.JobNum).ToList();
                iNumberOfRows = 1;
                iNumberOfColumns = 1;
                if (bFirstWorksheet == true)
                {
                    oSLBOMDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "Jobs Without Part Master");
                    bFirstWorksheet = false;
                }
                else
                {
                    oSLBOMDocument.AddWorksheet("Jobs Without Part Master");
                }
                //set column header
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "JobNum");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Part Num");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Part Rev Num");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "The Part Being Built For This Jobs Is Not In The Part Master -- SOP IS TO PULL BOM FROM PART MASTER");
                oSLBOMDocument.SetCellStyle(iNumberOfRows++, iNumberOfColumns, oHighlightHeaderStyle);
                oSLBOMDocument.SetColumnWidth(iNumberOfColumns++, 90);

                foreach (HSJob oJob in oJobsWithoutPartMaster)
                {
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 1, oJob.JobNum);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 2, oJob.PartNum);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 3, oJob.PartRevNum);

                    iNumberOfRows++;
                    bDataInReport = true;
                }
            }

            // jobs that have make direct finished goods but are not tied to SO -- issue trying to compute margin for POC
            List<HSJob> oMakeDirectJobsNotTiedToSalesOrder = m_oAllJobs.Where(oItem => (oItem.OrderNum == 0) && (oItem.PartMaster != null) && (oItem.PartMaster.PartNonStock == true)).ToList();
            if (oMakeDirectJobsNotTiedToSalesOrder.Count > 0)
            {
                // sort these by job number
                oMakeDirectJobsNotTiedToSalesOrder = oMakeDirectJobsNotTiedToSalesOrder.OrderBy(oItem => oItem.JobNum).ToList();
                iNumberOfRows = 1;
                iNumberOfColumns = 1;
                if (bFirstWorksheet == true)
                {
                    oSLBOMDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "Make Direct No SO");
                    bFirstWorksheet = false;
                }
                else
                {
                    oSLBOMDocument.AddWorksheet("Make Direct No SO");
                }
                //set column header
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "JobNum");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Part Num");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Part Rev Num");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "The Part Being Built Is Set To Be Make Direct But The Job Is Not Tied To A Sales Order -- POC ISSUE");
                oSLBOMDocument.SetCellStyle(iNumberOfRows++, iNumberOfColumns, oHighlightHeaderStyle);
                oSLBOMDocument.SetColumnWidth(iNumberOfColumns++, 90);

                foreach (HSJob oJob in oMakeDirectJobsNotTiedToSalesOrder)
                {
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 1, oJob.JobNum);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 2, oJob.PartNum);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 3, oJob.PartRevNum);

                    iNumberOfRows++;
                    bDataInReport = true;
                }
            }

            // jobs that are stocking but tied to sales order
            List<HSJob> oStockJobsTiedToSalesOrder = m_oAllJobs.Where(oItem => (oItem.OrderNum != 0) && (oItem.PartMaster != null) && (oItem.PartMaster.PartNonStock == false)).ToList();
            if (oStockJobsTiedToSalesOrder.Count > 0)
            {
                // sort these by job number
                oStockJobsTiedToSalesOrder = oStockJobsTiedToSalesOrder.OrderBy(oItem => oItem.JobNum).ToList();
                iNumberOfRows = 1;
                iNumberOfColumns = 1;
                if (bFirstWorksheet == true)
                {
                    oSLBOMDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "Stock MFG Tied To SO");
                    bFirstWorksheet = false;
                }
                else
                {
                    oSLBOMDocument.AddWorksheet("Stock MFG Tied To SO");
                }
                //set column header
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "JobNum");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Part Num");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Part Rev Num");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "The Part Being Built Is Set To Stocking But The Job Is Tied Directly To A Sales Order -- POC ISSUE");
                oSLBOMDocument.SetCellStyle(iNumberOfRows++, iNumberOfColumns, oHighlightHeaderStyle);
                oSLBOMDocument.SetColumnWidth(iNumberOfColumns++, 90);

                foreach (HSJob oJob in oStockJobsTiedToSalesOrder)
                {
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 1, oJob.JobNum);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 2, oJob.PartNum);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 3, oJob.PartRevNum);

                    iNumberOfRows++;
                    bDataInReport = true;
                }
            }

            if (m_bJustMissingCosts == false)
            {
                // jobs that are below the margin threshold
                List<HSJob> oJobsBelowMargin = m_oAllJobs.Where(oItem => (oItem.HasTransactions == true) && (oItem.MarginToDate * 100.0M < m_dJobMarginThreshold)).ToList();
                if (oJobsBelowMargin.Count > 0)
                {
                    // sort these by job number
                    oJobsBelowMargin = oJobsBelowMargin.OrderBy(oItem => oItem.JobNum).ToList();
                    iNumberOfRows = 1;
                    iNumberOfColumns = 1;
                    if (bFirstWorksheet == true)
                    {
                        oSLBOMDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "Jobs Below Margin");
                        bFirstWorksheet = false;
                    }
                    else
                    {
                        oSLBOMDocument.AddWorksheet("Jobs Below Margin");
                    }
                    //set column header
                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "JobNum");
                    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Part Num");
                    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Part Rev Num");
                    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Current Margin");
                    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Jobs Is Below Margin Threshold -- POC ISSUE PLEASE REVIEW");
                    oSLBOMDocument.SetCellStyle(iNumberOfRows++, iNumberOfColumns, oHighlightHeaderStyle);
                    oSLBOMDocument.SetColumnWidth(iNumberOfColumns++, 90);

                    foreach (HSJob oJob in oJobsBelowMargin)
                    {
                        oSLBOMDocument.SetCellValue(iNumberOfRows, 1, oJob.JobNum);
                        oSLBOMDocument.SetCellValue(iNumberOfRows, 2, oJob.PartNum);
                        oSLBOMDocument.SetCellValue(iNumberOfRows, 3, oJob.PartRevNum);
                        oSLBOMDocument.SetCellValue(iNumberOfRows, 4, oJob.MarginToDate);
                        oSLBOMDocument.SetCellStyle(iNumberOfRows, 4, oDecimalStyle);

                        iNumberOfRows++;
                        bDataInReport = true;
                    }
                }
            }
            #endregion

            #region Material Attributes

            List<JobMaterial> oMaterialsWithZeroQty = m_oJobMaterials.Where(oItem => (oItem.RequiredQty == 0) && (oItem.QtyBearing == true) && ((oItem.PartMaster == null) || (m_oPartClassesToIgnore.Contains(oItem.PartMaster.ClassID) == false))).ToList();
            if (oMaterialsWithZeroQty.Count > 0)
            {
                // sort these by job number
                oMaterialsWithZeroQty = oMaterialsWithZeroQty.OrderBy(oItem => oItem.JobNum).ToList();
                iNumberOfRows = 1;
                iNumberOfColumns = 1;
                if (bFirstWorksheet == true)
                {
                    oSLBOMDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "Mtl Qty Zero");
                    bFirstWorksheet = false;
                }
                else
                {
                    oSLBOMDocument.AddWorksheet("Mtl Qty Zero");
                }
                //set column header
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "JobNum");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Parent Part Num");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Parent Part Rev Num");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Assembly Seq");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Mtl Seq");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Mtl Part Num");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Mtl Part Rev Num");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Class");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Material On Job Has Qty Set To Zero -- POC ISSUE PLEASE SET QTY REQUIRED");
                oSLBOMDocument.SetCellStyle(iNumberOfRows++, iNumberOfColumns, oHighlightHeaderStyle);
                oSLBOMDocument.SetColumnWidth(iNumberOfColumns++, 90);

                foreach (JobMaterial oJobMateiral in oMaterialsWithZeroQty)
                {
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 1, oJobMateiral.JobNum);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 2, oJobMateiral.ParentPartNum);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 3, oJobMateiral.ParentRevNum);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 4, oJobMateiral.AssemblySeq.ToString());
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 5, oJobMateiral.MtlSeq.ToString());
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 6, oJobMateiral.MaterialPartNum);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 7, oJobMateiral.MaterialPartRevNum);
                    if (oJobMateiral.PartMaster != null)
                    {
                        oSLBOMDocument.SetCellValue(iNumberOfRows, 8, oJobMateiral.PartMaster.ClassID);
                    }

                    iNumberOfRows++;
                    bDataInReport = true;
                }
            }

            #region IGNORING THESE ISSUES
            // materials not tied to operatiopn
            //////////List<JobMaterial> oMaterialsNotAssociatedToOp = m_oJobMaterials.Where(oItem => (oItem.RelatedOperation == 0) && (oItem.PartMaster != null) && (m_oPartClassesToIgnore.Contains(oItem.PartMaster.ClassID) == false)).ToList();
            //////////if (oMaterialsNotAssociatedToOp.Count > 0)
            //////////{
            //////////    // sort these by job number
            //////////    oMaterialsNotAssociatedToOp = oMaterialsNotAssociatedToOp.OrderBy(oItem => oItem.JobNum).ToList();
            //////////    iNumberOfRows = 1;
            //////////    iNumberOfColumns = 1;
            //////////    if (bFirstWorksheet == true)
            //////////    {
            //////////        oSLBOMDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "Mtl No Op");
            //////////        bFirstWorksheet = false;
            //////////    }
            //////////    else
            //////////    {
            //////////        oSLBOMDocument.AddWorksheet("Mtl No Op");
            //////////    }
            //////////    //set column header
            //////////    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "JobNum");
            //////////    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
            //////////    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Parent Part Num");
            //////////    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
            //////////    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Parent Part Rev Num");
            //////////    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
            //////////    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Assembly Seq");
            //////////    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
            //////////    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Mtl Seq");
            //////////    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
            //////////    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Mtl Part Num");
            //////////    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
            //////////    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Mtl Part Rev Num");
            //////////    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
            //////////    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Class");
            //////////    oSLBOMDocument.SetColumnWidth(iNumberOfRows++, iNumberOfColumns++, iStandardColumnWidth);

            //////////    foreach (JobMaterial oJobMateiral in oMaterialsNotAssociatedToOp)
            //////////    {
            //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, 1, oJobMateiral.JobNum);
            //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, 2, oJobMateiral.ParentPartNum);
            //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, 3, oJobMateiral.ParentRevNum);
            //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, 4, oJobMateiral.AssemblySeq.ToString());
            //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, 5, oJobMateiral.MtlSeq.ToString());
            //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, 6, oJobMateiral.MaterialPartNum);
            //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, 7, oJobMateiral.MaterialPartRevNum);
            //////////        if (oJobMateiral.PartMaster != null)
            //////////        {
            //////////            oSLBOMDocument.SetCellValue(iNumberOfRows, 8, oJobMateiral.PartMaster.ClassID);
            //////////        }

            //////////        iNumberOfRows++;
            //////////        bDataInReport = true;
            //////////    }
            //////////}

            // materials no lead time
            //////////List<JobMaterial> oMaterialsNoLeadTime = m_oJobMaterials.Where(oItem => (oItem.LeadTime == 0) && (oItem.PartMaster != null) && (m_oPartClassesToIgnore.Contains(oItem.PartMaster.ClassID) == false)).ToList();
            //////////if (oMaterialsNoLeadTime.Count > 0)
            //////////{
            //////////    // sort these by job number
            //////////    oMaterialsNoLeadTime = oMaterialsNoLeadTime.OrderBy(oItem => oItem.JobNum).ToList();
            //////////    iNumberOfRows = 1;
            //////////    iNumberOfColumns = 1;
            //////////    if (bFirstWorksheet == true)
            //////////    {
            //////////        oSLBOMDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "Mtl No Lead Time");
            //////////        bFirstWorksheet = false;
            //////////    }
            //////////    else
            //////////    {
            //////////        oSLBOMDocument.AddWorksheet("Mtl No Lead Time");
            //////////    }
            //////////    //set column header
            //////////    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "JobNum");
            //////////    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
            //////////    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Parent Part Num");
            //////////    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
            //////////    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Parent Part Rev Num");
            //////////    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
            //////////    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Assembly Seq");
            //////////    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
            //////////    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Mtl Seq");
            //////////    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
            //////////    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Mtl Part Num");
            //////////    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
            //////////    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Mtl Part Rev Num");
            //////////    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
            //////////    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Class");
            //////////    oSLBOMDocument.SetColumnWidth(iNumberOfRows++, iNumberOfColumns++, iStandardColumnWidth);

            //////////    foreach (JobMaterial oJobMateiral in oMaterialsNoLeadTime)
            //////////    {
            //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, 1, oJobMateiral.JobNum);
            //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, 2, oJobMateiral.ParentPartNum);
            //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, 3, oJobMateiral.ParentRevNum);
            //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, 4, oJobMateiral.AssemblySeq.ToString());
            //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, 5, oJobMateiral.MtlSeq.ToString());
            //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, 6, oJobMateiral.MaterialPartNum);
            //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, 7, oJobMateiral.MaterialPartRevNum);
            //////////        if (oJobMateiral.PartMaster != null)
            //////////        {
            //////////            oSLBOMDocument.SetCellValue(iNumberOfRows, 8, oJobMateiral.PartMaster.ClassID);
            //////////        }

            //////////        iNumberOfRows++;
            //////////        bDataInReport = true;
            //////////    }
            //////////}

            // materials no required date
            //////////List<JobMaterial> oMaterialsNoRequiredDate = m_oJobMaterials.Where(oItem => (oItem.RequiredDate == null) || (oItem.RequiredDate == DateTime.MinValue) && (oItem.PartMaster != null) && (m_oPartClassesToIgnore.Contains(oItem.PartMaster.ClassID) == false)).ToList();
            //////////if (oMaterialsNoRequiredDate.Count > 0)
            //////////{
            //////////    // sort these by job number
            //////////    oMaterialsNoRequiredDate = oMaterialsNoRequiredDate.OrderBy(oItem => oItem.JobNum).ToList();
            //////////    iNumberOfRows = 1;
            //////////    iNumberOfColumns = 1;
            //////////    if (bFirstWorksheet == true)
            //////////    {
            //////////        oSLBOMDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "Mtl No Required Date");
            //////////        bFirstWorksheet = false;
            //////////    }
            //////////    else
            //////////    {
            //////////        oSLBOMDocument.AddWorksheet("Mtl No Required Date");
            //////////    }
            //////////    //set column header
            //////////    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "JobNum");
            //////////    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
            //////////    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Parent Part Num");
            //////////    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
            //////////    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Parent Part Rev Num");
            //////////    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
            //////////    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Assembly Seq");
            //////////    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
            //////////    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Mtl Seq");
            //////////    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
            //////////    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Mtl Part Num");
            //////////    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
            //////////    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Mtl Part Rev Num");
            //////////    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
            //////////    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Class");
            //////////    oSLBOMDocument.SetColumnWidth(iNumberOfRows++, iNumberOfColumns++, iStandardColumnWidth);

            //////////    foreach (JobMaterial oJobMateiral in oMaterialsNoRequiredDate)
            //////////    {
            //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, 1, oJobMateiral.JobNum);
            //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, 2, oJobMateiral.ParentPartNum);
            //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, 3, oJobMateiral.ParentRevNum);
            //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, 4, oJobMateiral.AssemblySeq.ToString());
            //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, 5, oJobMateiral.MtlSeq.ToString());
            //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, 6, oJobMateiral.MaterialPartNum);
            //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, 7, oJobMateiral.MaterialPartRevNum);
            //////////        if (oJobMateiral.PartMaster != null)
            //////////        {
            //////////            oSLBOMDocument.SetCellValue(iNumberOfRows, 8, oJobMateiral.PartMaster.ClassID);
            //////////        }

            //////////        iNumberOfRows++;
            //////////        bDataInReport = true;
            //////////    }
            //////////}

            // materials set to backflush
            //////////List<JobMaterial> oMaterialsSetToBackflush = m_oJobMaterials.Where(oItem => oItem.Backflush == true).ToList();
            //////////if (oMaterialsSetToBackflush.Count > 0)
            //////////{
            //////////    // sort these by job number
            //////////    oMaterialsSetToBackflush = oMaterialsSetToBackflush.OrderBy(oItem => oItem.JobNum).ToList();
            //////////    iNumberOfRows = 1;
            //////////    iNumberOfColumns = 1;
            //////////    if (bFirstWorksheet == true)
            //////////    {
            //////////        oSLBOMDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "Mtl Set To Backflush");
            //////////        bFirstWorksheet = false;
            //////////    }
            //////////    else
            //////////    {
            //////////        oSLBOMDocument.AddWorksheet("Mtl Set To Backflush");
            //////////    }
            //////////    //set column header
            //////////    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "JobNum");
            //////////    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
            //////////    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Parent Part Num");
            //////////    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
            //////////    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Parent Part Rev Num");
            //////////    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
            //////////    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Assembly Seq");
            //////////    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
            //////////    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Mtl Seq");
            //////////    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
            //////////    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Mtl Part Num");
            //////////    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
            //////////    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Mtl Part Rev Num");
            //////////    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
            //////////    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Class");
            //////////    oSLBOMDocument.SetColumnWidth(iNumberOfRows++, iNumberOfColumns++, iStandardColumnWidth);

            //////////    foreach (JobMaterial oJobMateiral in oMaterialsSetToBackflush)
            //////////    {
            //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, 1, oJobMateiral.JobNum);
            //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, 2, oJobMateiral.ParentPartNum);
            //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, 3, oJobMateiral.ParentRevNum);
            //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, 4, oJobMateiral.AssemblySeq.ToString());
            //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, 5, oJobMateiral.MtlSeq.ToString());
            //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, 6, oJobMateiral.MaterialPartNum);
            //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, 7, oJobMateiral.MaterialPartRevNum);
            //////////        if (oJobMateiral.PartMaster != null)
            //////////        {
            //////////            oSLBOMDocument.SetCellValue(iNumberOfRows, 8, oJobMateiral.PartMaster.ClassID);
            //////////        }

            //////////        iNumberOfRows++;
            //////////        bDataInReport = true;
            //////////    }
            //////////}

            // materials on hold

            // materials inactive

            // materials bad rev

            // materials MRP turned off

            // group code not set

            // class code not set
            #endregion

            // material issues
            List<JobMaterial> oExcessMaterialIssued = m_oJobMaterials.Where(oItem => (oItem.IssuedQty > oItem.RequiredQty) || ((oItem.IssuedQty < oItem.RequiredQty) && (oItem.QtyBearing == true) && (oItem.OpComplete == true) && ((oItem.PartMaster == null) || (m_oPartClassesToIgnore.Contains(oItem.PartMaster.ClassID) == false)))).ToList();
            if (oExcessMaterialIssued.Count > 0)
            {
                // sort these by job number
                oExcessMaterialIssued = oExcessMaterialIssued.OrderBy(oItem => oItem.JobNum).ToList();
                iNumberOfRows = 1;
                iNumberOfColumns = 1;
                if (bFirstWorksheet == true)
                {
                    oSLBOMDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "Mtl Issued");
                    bFirstWorksheet = false;
                }
                else
                {
                    oSLBOMDocument.AddWorksheet("Mtl Issued");
                }
                //set column header
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "JobNum");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Parent Part Num");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Parent Part Rev Num");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Assembly Seq");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Mtl Seq");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Mtl Part Num");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Mtl Part Rev Num");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Class");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Mtl Req");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Mtl Issued");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Variance Qty");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Variance %");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Material Issued To Job Differs From Estimate -- POC ISSUE PLEASE REVIEW");
                oSLBOMDocument.SetCellStyle(iNumberOfRows++, iNumberOfColumns, oHighlightHeaderStyle);
                oSLBOMDocument.SetColumnWidth(iNumberOfColumns++, 90);

                foreach (JobMaterial oJobMateiral in oExcessMaterialIssued)
                {
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 1, oJobMateiral.JobNum);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 2, oJobMateiral.ParentPartNum);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 3, oJobMateiral.ParentRevNum);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 4, oJobMateiral.AssemblySeq.ToString());
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 5, oJobMateiral.MtlSeq.ToString());
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 6, oJobMateiral.MaterialPartNum);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 7, oJobMateiral.MaterialPartRevNum);
                    if (oJobMateiral.PartMaster != null)
                    {
                        oSLBOMDocument.SetCellValue(iNumberOfRows, 8, oJobMateiral.PartMaster.ClassID);
                    }
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 9, oJobMateiral.RequiredQty);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 10, oJobMateiral.IssuedQty);
                    decimal dDelta = oJobMateiral.IssuedQty - oJobMateiral.RequiredQty;
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 11, dDelta);
                    oSLBOMDocument.SetCellStyle(iNumberOfRows, 11, oDecimalStyle);
                    decimal dPercentageError = 0M;
                    if (oJobMateiral.RequiredQty != 0)
                    {
                        dPercentageError = dDelta / oJobMateiral.RequiredQty * 100.0M;
                    }
                    else
                    {
                        dPercentageError = 100.0M;
                    }
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 12, dPercentageError);
                    oSLBOMDocument.SetCellStyle(iNumberOfRows, 12, oDecimalStyle);

                    iNumberOfRows++;
                    bDataInReport = true;
                }
            }

            #endregion

            #region Material Costs
            // materials without cost
            List<JobMaterial> oMaterialsWithoutCost = m_oJobMaterials.Where(oItem => (oItem.EstUnitCost == 0) && (oItem.QtyBearing == true) && ((oItem.PartMaster == null) || ((oItem.PartMaster != null) && (m_oPartClassesToIgnore.Contains(oItem.PartMaster.ClassID) == false)))).ToList();
            if (oMaterialsWithoutCost.Count > 0)
            {
                // sort these by job number
                oMaterialsWithoutCost = oMaterialsWithoutCost.OrderBy(oItem => oItem.JobNum).ToList();
                iNumberOfRows = 1;
                iNumberOfColumns = 1;
                if (bFirstWorksheet == true)
                {
                    oSLBOMDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "Mtl Cost Is Zero");
                    bFirstWorksheet = false;
                }
                else
                {
                    oSLBOMDocument.AddWorksheet("Mtl Cost Is Zero");
                }
                //set column header
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "JobNum");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Parent Part Num");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Parent Part Rev Num");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Assembly Seq");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Mtl Seq");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Mtl Part Num");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Mtl Part Rev Num");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Class");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Material Cost Is Not Set -- POC ISSUE PLEASE SET COST FOR MATERIAL");
                oSLBOMDocument.SetCellStyle(iNumberOfRows++, iNumberOfColumns, oHighlightHeaderStyle);
                oSLBOMDocument.SetColumnWidth(iNumberOfColumns++, 90);

                foreach (JobMaterial oJobMateiral in oMaterialsWithoutCost)
                {
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 1, oJobMateiral.JobNum);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 2, oJobMateiral.ParentPartNum);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 3, oJobMateiral.ParentRevNum);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 4, oJobMateiral.AssemblySeq.ToString());
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 5, oJobMateiral.MtlSeq.ToString());
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 6, oJobMateiral.MaterialPartNum);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 7, oJobMateiral.MaterialPartRevNum);
                    if (oJobMateiral.PartMaster != null)
                    {
                        oSLBOMDocument.SetCellValue(iNumberOfRows, 8, oJobMateiral.PartMaster.ClassID);
                    }

                    iNumberOfRows++;
                    bDataInReport = true;
                }
            }

            #region IGNORING THESE ISSUES
            //////////if (m_bJustMissingCosts == false)
            //////////{
            //////////    // materials with unexpected material cost
            //////////    List<JobMaterial> oMtlWithUnexpectedMtlCost = m_oJobMaterials.Where(oItem => (oItem.MaterialCost > (oItem.EstMtlUnitCost * oItem.RequiredQty)) || ((oItem.MaterialCost < (oItem.EstMtlUnitCost * oItem.RequiredQty)) && (oItem.OpComplete == true))).ToList();
            //////////    if (oMtlWithUnexpectedMtlCost.Count > 0)
            //////////    {
            //////////        // sort these by job number
            //////////        oMtlWithUnexpectedMtlCost = oMtlWithUnexpectedMtlCost.OrderBy(oItem => oItem.JobNum).ToList();
            //////////        iNumberOfRows = 1;
            //////////        iNumberOfColumns = 1;
            //////////        if (bFirstWorksheet == true)
            //////////        {
            //////////            oSLBOMDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "Mtl Unexpected Mtl Cost");
            //////////            bFirstWorksheet = false;
            //////////        }
            //////////        else
            //////////        {
            //////////            oSLBOMDocument.AddWorksheet("Mtl Unexpected Mtl Cost");
            //////////        }
            //////////        //set column header
            //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "JobNum");
            //////////        oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
            //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Parent Part Num");
            //////////        oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
            //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Parent Part Rev Num");
            //////////        oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
            //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Assembly Seq");
            //////////        oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
            //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Mtl Seq");
            //////////        oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
            //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Mtl Part Num");
            //////////        oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
            //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Mtl Part Rev Num");
            //////////        oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
            //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Class");
            //////////        oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
            //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Est Mtl Cost");
            //////////        oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
            //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Act Mtl Cost");
            //////////        oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
            //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Variance $");
            //////////        oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
            //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Variance %");
            //////////        oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
            //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Material Actual Cost Differs From Estimate -- POC ISSUE PLEASE REVIEW");
            //////////        oSLBOMDocument.SetCellStyle(iNumberOfRows++, iNumberOfColumns, oHighlightHeaderStyle);
            //////////        oSLBOMDocument.SetColumnWidth(iNumberOfColumns++, 90);

            //////////        foreach (JobMaterial oJobMaterial in oMtlWithUnexpectedMtlCost)
            //////////        {
            //////////            decimal dEstTotal = oJobMaterial.EstMtlUnitCost * oJobMaterial.RequiredQty;
            //////////            decimal dDelta = oJobMaterial.MaterialCost - dEstTotal;
            //////////            decimal dPercentageError = 0M;
            //////////            decimal dMarginToDate = 0M;
            //////////            if (dEstTotal != 0)
            //////////            {
            //////////                dPercentageError = dDelta / dEstTotal * 100.0M;
            //////////            }
            //////////            else
            //////////            {
            //////////                dPercentageError = 100.0M;
            //////////            }
            //////////            HSJob oTmpJob = m_oAllJobs.FirstOrDefault(oItem => string.Compare(oJobMaterial.JobNum, oItem.JobNum, true) == 0);
            //////////            if (oTmpJob != null)
            //////////            {
            //////////                dMarginToDate = oTmpJob.MarginToDate * 100.0M;
            //////////            }

            //////////            bool bExceedsThreholds = false;
            //////////            if ((Math.Abs(dDelta) > m_dJobEstimateAbsoluteError) && (Math.Abs(dPercentageError) > m_dJobEstimatePercentError) && (dMarginToDate < m_dJobMarginThreshold))
            //////////            {
            //////////                // check to see if it exceeds both thresholds
            //////////                bExceedsThreholds = true;
            //////////            }
            //////////            if (m_bForceJobAnalysis == true)
            //////////            {
            //////////                // force it to report
            //////////                bExceedsThreholds = true;
            //////////            }

            //////////            if (bExceedsThreholds == true)
            //////////            {

            //////////                oSLBOMDocument.SetCellValue(iNumberOfRows, 1, oJobMaterial.JobNum);
            //////////                oSLBOMDocument.SetCellValue(iNumberOfRows, 2, oJobMaterial.ParentPartNum);
            //////////                oSLBOMDocument.SetCellValue(iNumberOfRows, 3, oJobMaterial.ParentRevNum);
            //////////                oSLBOMDocument.SetCellValue(iNumberOfRows, 4, oJobMaterial.AssemblySeq.ToString());
            //////////                oSLBOMDocument.SetCellValue(iNumberOfRows, 5, oJobMaterial.MtlSeq.ToString());
            //////////                oSLBOMDocument.SetCellValue(iNumberOfRows, 6, oJobMaterial.MaterialPartNum);
            //////////                oSLBOMDocument.SetCellValue(iNumberOfRows, 7, oJobMaterial.MaterialPartRevNum);
            //////////                if (oJobMaterial.PartMaster != null)
            //////////                {
            //////////                    oSLBOMDocument.SetCellValue(iNumberOfRows, 8, oJobMaterial.PartMaster.ClassID);
            //////////                }
            //////////                oSLBOMDocument.SetCellValue(iNumberOfRows, 9, oJobMaterial.EstMtlUnitCost * oJobMaterial.RequiredQty);
            //////////                oSLBOMDocument.SetCellStyle(iNumberOfRows, 9, oCurrencyStyle);
            //////////                oSLBOMDocument.SetCellValue(iNumberOfRows, 10, oJobMaterial.MaterialCost);
            //////////                oSLBOMDocument.SetCellStyle(iNumberOfRows, 10, oCurrencyStyle);

            //////////                oSLBOMDocument.SetCellValue(iNumberOfRows, 11, dDelta);
            //////////                oSLBOMDocument.SetCellStyle(iNumberOfRows, 11, oCurrencyStyle);

            //////////                oSLBOMDocument.SetCellValue(iNumberOfRows, 12, dPercentageError);
            //////////                oSLBOMDocument.SetCellStyle(iNumberOfRows, 12, oDecimalStyle);

            //////////                iNumberOfRows++;
            //////////                bDataInReport = true;
            //////////            }
            //////////        }
            //////////    }

            //////////    // materials with unexpected burden
            //////////    List<JobMaterial> oMtlWithUnexpectedBurden = m_oJobMaterials.Where(oItem => (oItem.BurdenCost > (oItem.EstBurdenUnitCost * oItem.RequiredQty)) || ((oItem.BurdenCost < (oItem.EstBurdenUnitCost * oItem.RequiredQty)) && (oItem.OpComplete == true))).ToList();
            //////////    if (oMtlWithUnexpectedBurden.Count > 0)
            //////////    {
            //////////        // sort these by job number
            //////////        oMtlWithUnexpectedBurden = oMtlWithUnexpectedBurden.OrderBy(oItem => oItem.JobNum).ToList();
            //////////        iNumberOfRows = 1;
            //////////        iNumberOfColumns = 1;
            //////////        if (bFirstWorksheet == true)
            //////////        {
            //////////            oSLBOMDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "Mtl Unexpected Burden");
            //////////            bFirstWorksheet = false;
            //////////        }
            //////////        else
            //////////        {
            //////////            oSLBOMDocument.AddWorksheet("Mtl Unexpected Burden");
            //////////        }
            //////////        //set column header
            //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "JobNum");
            //////////        oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
            //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Parent Part Num");
            //////////        oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
            //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Parent Part Rev Num");
            //////////        oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
            //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Assembly Seq");
            //////////        oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
            //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Mtl Seq");
            //////////        oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
            //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Mtl Part Num");
            //////////        oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
            //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Mtl Part Rev Num");
            //////////        oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
            //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Class");
            //////////        oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
            //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Est Burden");
            //////////        oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
            //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Act Burden");
            //////////        oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
            //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Variance $");
            //////////        oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
            //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Variance %");
            //////////        oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
            //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Material Burden Cost Job Differs From Estimate -- POC ISSUE PLEASE REVIEW");
            //////////        oSLBOMDocument.SetCellStyle(iNumberOfRows++, iNumberOfColumns, oHighlightHeaderStyle);
            //////////        oSLBOMDocument.SetColumnWidth(iNumberOfColumns++, 90);

            //////////        foreach (JobMaterial oJobMaterial in oMtlWithUnexpectedBurden)
            //////////        {
            //////////            decimal dEstTotal = oJobMaterial.EstBurdenUnitCost * oJobMaterial.RequiredQty;
            //////////            decimal dDelta = oJobMaterial.BurdenCost - dEstTotal;
            //////////            decimal dPercentageError = 0M;
            //////////            decimal dMarginToDate = 0M;
            //////////            if (dEstTotal != 0)
            //////////            {
            //////////                dPercentageError = dDelta / dEstTotal * 100.0M;
            //////////            }
            //////////            else
            //////////            {
            //////////                dPercentageError = 100.0M;
            //////////            }
            //////////            HSJob oTmpJob = m_oAllJobs.FirstOrDefault(oItem => string.Compare(oJobMaterial.JobNum, oItem.JobNum, true) == 0);
            //////////            if (oTmpJob != null)
            //////////            {
            //////////                dMarginToDate = oTmpJob.MarginToDate * 100.0M;
            //////////            }

            //////////            bool bExceedsThreholds = false;
            //////////            if ((Math.Abs(dDelta) > m_dJobEstimateAbsoluteError) && (Math.Abs(dPercentageError) > m_dJobEstimatePercentError) && (dMarginToDate < m_dJobMarginThreshold))
            //////////            {
            //////////                // check to see if it exceeds both thresholds
            //////////                bExceedsThreholds = true;
            //////////            }
            //////////            if (m_bForceJobAnalysis == true)
            //////////            {
            //////////                // force it to report
            //////////                bExceedsThreholds = true;
            //////////            }

            //////////            if (bExceedsThreholds == true)
            //////////            {
            //////////                oSLBOMDocument.SetCellValue(iNumberOfRows, 1, oJobMaterial.JobNum);
            //////////                oSLBOMDocument.SetCellValue(iNumberOfRows, 2, oJobMaterial.ParentPartNum);
            //////////                oSLBOMDocument.SetCellValue(iNumberOfRows, 3, oJobMaterial.ParentRevNum);
            //////////                oSLBOMDocument.SetCellValue(iNumberOfRows, 4, oJobMaterial.AssemblySeq.ToString());
            //////////                oSLBOMDocument.SetCellValue(iNumberOfRows, 5, oJobMaterial.MtlSeq.ToString());
            //////////                oSLBOMDocument.SetCellValue(iNumberOfRows, 6, oJobMaterial.MaterialPartNum);
            //////////                oSLBOMDocument.SetCellValue(iNumberOfRows, 7, oJobMaterial.MaterialPartRevNum);
            //////////                if (oJobMaterial.PartMaster != null)
            //////////                {
            //////////                    oSLBOMDocument.SetCellValue(iNumberOfRows, 8, oJobMaterial.PartMaster.ClassID);
            //////////                }
            //////////                oSLBOMDocument.SetCellValue(iNumberOfRows, 9, oJobMaterial.EstBurdenUnitCost * oJobMaterial.RequiredQty);
            //////////                oSLBOMDocument.SetCellStyle(iNumberOfRows, 9, oCurrencyStyle);
            //////////                oSLBOMDocument.SetCellValue(iNumberOfRows, 10, oJobMaterial.BurdenCost);
            //////////                oSLBOMDocument.SetCellStyle(iNumberOfRows, 10, oCurrencyStyle);

            //////////                oSLBOMDocument.SetCellValue(iNumberOfRows, 11, dDelta);
            //////////                oSLBOMDocument.SetCellStyle(iNumberOfRows, 11, oCurrencyStyle);

            //////////                oSLBOMDocument.SetCellValue(iNumberOfRows, 12, dPercentageError);
            //////////                oSLBOMDocument.SetCellStyle(iNumberOfRows, 12, oDecimalStyle);

            //////////                iNumberOfRows++;
            //////////                bDataInReport = true;
            //////////            }
            //////////        }
            //////////    }

            //////////    // materials with unexpected labor
            //////////    List<JobMaterial> oMtlWithUnexpectedLabor = m_oJobMaterials.Where(oItem => (oItem.LaborCost > (oItem.EstLaborUnitCost * oItem.RequiredQty)) || ((oItem.LaborCost < (oItem.EstLaborUnitCost * oItem.RequiredQty)) && (oItem.OpComplete == true))).ToList();
            //////////    if (oMtlWithUnexpectedLabor.Count > 0)
            //////////    {
            //////////        // sort these by job number
            //////////        oMtlWithUnexpectedLabor = oMtlWithUnexpectedLabor.OrderBy(oItem => oItem.JobNum).ToList();
            //////////        iNumberOfRows = 1;
            //////////        iNumberOfColumns = 1;
            //////////        if (bFirstWorksheet == true)
            //////////        {
            //////////            oSLBOMDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "Mtl Unexpected Labor");
            //////////            bFirstWorksheet = false;
            //////////        }
            //////////        else
            //////////        {
            //////////            oSLBOMDocument.AddWorksheet("Mtl Unexpected Labor");
            //////////        }
            //////////        //set column header
            //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "JobNum");
            //////////        oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
            //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Parent Part Num");
            //////////        oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
            //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Parent Part Rev Num");
            //////////        oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
            //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Assembly Seq");
            //////////        oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
            //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Mtl Seq");
            //////////        oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
            //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Mtl Part Num");
            //////////        oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
            //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Mtl Part Rev Num");
            //////////        oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
            //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Class");
            //////////        oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
            //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Est Labor");
            //////////        oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
            //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Act Labor");
            //////////        oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
            //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Variance $");
            //////////        oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
            //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Variance %");
            //////////        oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
            //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Material Labor Cost Differs From Estimate -- POC ISSUE PLEASE REVIEW");
            //////////        oSLBOMDocument.SetCellStyle(iNumberOfRows++, iNumberOfColumns, oHighlightHeaderStyle);
            //////////        oSLBOMDocument.SetColumnWidth(iNumberOfColumns++, 90);

            //////////        foreach (JobMaterial oJobMaterial in oMtlWithUnexpectedLabor)
            //////////        {
            //////////            decimal dEstTotal = oJobMaterial.EstLaborUnitCost * oJobMaterial.RequiredQty;
            //////////            decimal dDelta = oJobMaterial.LaborCost - dEstTotal;
            //////////            decimal dPercentageError = 0M;
            //////////            decimal dMarginToDate = 0M;
            //////////            if (dEstTotal != 0)
            //////////            {
            //////////                dPercentageError = dDelta / dEstTotal * 100.0M;
            //////////            }
            //////////            else
            //////////            {
            //////////                dPercentageError = 100.0M;
            //////////            }
            //////////            HSJob oTmpJob = m_oAllJobs.FirstOrDefault(oItem => string.Compare(oJobMaterial.JobNum, oItem.JobNum, true) == 0);
            //////////            if (oTmpJob != null)
            //////////            {
            //////////                dMarginToDate = oTmpJob.MarginToDate * 100.0M;
            //////////            }

            //////////            bool bExceedsThreholds = false;
            //////////            if ((Math.Abs(dDelta) > m_dJobEstimateAbsoluteError) && (Math.Abs(dPercentageError) > m_dJobEstimatePercentError) && (dMarginToDate < m_dJobMarginThreshold))
            //////////            {
            //////////                // check to see if it exceeds both thresholds
            //////////                bExceedsThreholds = true;
            //////////            }
            //////////            if (m_bForceJobAnalysis == true)
            //////////            {
            //////////                // force it to report
            //////////                bExceedsThreholds = true;
            //////////            }

            //////////            if (bExceedsThreholds == true)
            //////////            {
            //////////                oSLBOMDocument.SetCellValue(iNumberOfRows, 1, oJobMaterial.JobNum);
            //////////                oSLBOMDocument.SetCellValue(iNumberOfRows, 2, oJobMaterial.ParentPartNum);
            //////////                oSLBOMDocument.SetCellValue(iNumberOfRows, 3, oJobMaterial.ParentRevNum);
            //////////                oSLBOMDocument.SetCellValue(iNumberOfRows, 4, oJobMaterial.AssemblySeq.ToString());
            //////////                oSLBOMDocument.SetCellValue(iNumberOfRows, 5, oJobMaterial.MtlSeq.ToString());
            //////////                oSLBOMDocument.SetCellValue(iNumberOfRows, 6, oJobMaterial.MaterialPartNum);
            //////////                oSLBOMDocument.SetCellValue(iNumberOfRows, 7, oJobMaterial.MaterialPartRevNum);
            //////////                if (oJobMaterial.PartMaster != null)
            //////////                {
            //////////                    oSLBOMDocument.SetCellValue(iNumberOfRows, 8, oJobMaterial.PartMaster.ClassID);
            //////////                }
            //////////                oSLBOMDocument.SetCellValue(iNumberOfRows, 9, oJobMaterial.EstLaborUnitCost * oJobMaterial.RequiredQty);
            //////////                oSLBOMDocument.SetCellStyle(iNumberOfRows, 9, oCurrencyStyle);
            //////////                oSLBOMDocument.SetCellValue(iNumberOfRows, 10, oJobMaterial.LaborCost);
            //////////                oSLBOMDocument.SetCellStyle(iNumberOfRows, 10, oCurrencyStyle);

            //////////                oSLBOMDocument.SetCellValue(iNumberOfRows, 11, dDelta);
            //////////                oSLBOMDocument.SetCellStyle(iNumberOfRows, 11, oCurrencyStyle);

            //////////                oSLBOMDocument.SetCellValue(iNumberOfRows, 12, dPercentageError);
            //////////                oSLBOMDocument.SetCellStyle(iNumberOfRows, 12, oDecimalStyle);

            //////////                iNumberOfRows++;
            //////////                bDataInReport = true;
            //////////            }
            //////////        }
            //////////    }

            //////////    // materials wtih unepected subcontract costs
            //////////    List<JobMaterial> oMtlWithUnexpectedSubcontractCosts = m_oJobMaterials.Where(oItem => (oItem.SubcontractCost > (oItem.EstSubcontractUnitCost * oItem.RequiredQty)) || ((oItem.SubcontractCost < (oItem.EstSubcontractUnitCost * oItem.RequiredQty)) && (oItem.OpComplete == true))).ToList();
            //////////    if (oMtlWithUnexpectedSubcontractCosts.Count > 0)
            //////////    {
            //////////        // sort these by job number
            //////////        oMtlWithUnexpectedSubcontractCosts = oMtlWithUnexpectedSubcontractCosts.OrderBy(oItem => oItem.JobNum).ToList();
            //////////        iNumberOfRows = 1;
            //////////        iNumberOfColumns = 1;
            //////////        if (bFirstWorksheet == true)
            //////////        {
            //////////            oSLBOMDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "Mtl Unexpected Subcontract");
            //////////            bFirstWorksheet = false;
            //////////        }
            //////////        else
            //////////        {
            //////////            oSLBOMDocument.AddWorksheet("Mtl Unexpected Subcontract");
            //////////        }
            //////////        //set column header
            //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "JobNum");
            //////////        oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
            //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Parent Part Num");
            //////////        oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
            //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Parent Part Rev Num");
            //////////        oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
            //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Assembly Seq");
            //////////        oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
            //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Mtl Seq");
            //////////        oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
            //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Mtl Part Num");
            //////////        oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
            //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Mtl Part Rev Num");
            //////////        oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
            //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Class");
            //////////        oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
            //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Est Subcontract");
            //////////        oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
            //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Act Subcontract");
            //////////        oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
            //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Variance $");
            //////////        oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
            //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Variance %");
            //////////        oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
            //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Material Subcontract Cost Differs From Estimate -- POC ISSUE PLEASE REVIEW");
            //////////        oSLBOMDocument.SetCellStyle(iNumberOfRows++, iNumberOfColumns, oHighlightHeaderStyle);
            //////////        oSLBOMDocument.SetColumnWidth(iNumberOfColumns++, 90);

            //////////        foreach (JobMaterial oJobMaterial in oMtlWithUnexpectedLabor)
            //////////        {
            //////////            decimal dEstTotal = oJobMaterial.EstSubcontractUnitCost * oJobMaterial.RequiredQty;
            //////////            decimal dDelta = oJobMaterial.SubcontractCost - dEstTotal;
            //////////            decimal dPercentageError = 0M;
            //////////            decimal dMarginToDate = 0M;
            //////////            if (dEstTotal != 0)
            //////////            {
            //////////                dPercentageError = dDelta / dEstTotal * 100.0M;
            //////////            }
            //////////            else
            //////////            {
            //////////                dPercentageError = 100.0M;
            //////////            }
            //////////            HSJob oTmpJob = m_oAllJobs.FirstOrDefault(oItem => string.Compare(oJobMaterial.JobNum, oItem.JobNum, true) == 0);
            //////////            if (oTmpJob != null)
            //////////            {
            //////////                dMarginToDate = oTmpJob.MarginToDate * 100.0M;
            //////////            }

            //////////            bool bExceedsThreholds = false;
            //////////            if ((Math.Abs(dDelta) > m_dJobEstimateAbsoluteError) && (Math.Abs(dPercentageError) > m_dJobEstimatePercentError) && (dMarginToDate < m_dJobMarginThreshold))
            //////////            {
            //////////                // check to see if it exceeds both thresholds
            //////////                bExceedsThreholds = true;
            //////////            }
            //////////            if (m_bForceJobAnalysis == true)
            //////////            {
            //////////                // force it to report
            //////////                bExceedsThreholds = true;
            //////////            }

            //////////            if (bExceedsThreholds == true)
            //////////            {
            //////////                oSLBOMDocument.SetCellValue(iNumberOfRows, 1, oJobMaterial.JobNum);
            //////////                oSLBOMDocument.SetCellValue(iNumberOfRows, 2, oJobMaterial.ParentPartNum);
            //////////                oSLBOMDocument.SetCellValue(iNumberOfRows, 3, oJobMaterial.ParentRevNum);
            //////////                oSLBOMDocument.SetCellValue(iNumberOfRows, 4, oJobMaterial.AssemblySeq.ToString());
            //////////                oSLBOMDocument.SetCellValue(iNumberOfRows, 5, oJobMaterial.MtlSeq.ToString());
            //////////                oSLBOMDocument.SetCellValue(iNumberOfRows, 6, oJobMaterial.MaterialPartNum);
            //////////                oSLBOMDocument.SetCellValue(iNumberOfRows, 7, oJobMaterial.MaterialPartRevNum);
            //////////                if (oJobMaterial.PartMaster != null)
            //////////                {
            //////////                    oSLBOMDocument.SetCellValue(iNumberOfRows, 8, oJobMaterial.PartMaster.ClassID);
            //////////                }
            //////////                oSLBOMDocument.SetCellValue(iNumberOfRows, 9, oJobMaterial.EstSubcontractUnitCost * oJobMaterial.RequiredQty);
            //////////                oSLBOMDocument.SetCellStyle(iNumberOfRows, 9, oCurrencyStyle);
            //////////                oSLBOMDocument.SetCellValue(iNumberOfRows, 10, oJobMaterial.SubcontractCost);
            //////////                oSLBOMDocument.SetCellStyle(iNumberOfRows, 10, oCurrencyStyle);

            //////////                oSLBOMDocument.SetCellValue(iNumberOfRows, 11, dDelta);
            //////////                oSLBOMDocument.SetCellStyle(iNumberOfRows, 11, oCurrencyStyle);

            //////////                oSLBOMDocument.SetCellValue(iNumberOfRows, 12, dPercentageError);
            //////////                oSLBOMDocument.SetCellStyle(iNumberOfRows, 12, oDecimalStyle);

            //////////                iNumberOfRows++;
            //////////                bDataInReport = true;
            //////////            }
            //////////        }
            //////////    }
            //////////}
            #endregion

            #endregion

            #region Operation Attributes
            // resource id set on operation
            // we should not specify a resource id on the operation -- too specific
            if (m_bJustMissingCosts == false)
            {
                List<JobOperation> oOperationsWithResourceIdSet = m_oJobOperations.Where(oItem => string.IsNullOrEmpty(oItem.ResourceId) == false).ToList();
                if (oOperationsWithResourceIdSet.Count > 0)
                {
                    // sort these by job number
                    oOperationsWithResourceIdSet = oOperationsWithResourceIdSet.OrderBy(oItem => oItem.JobNum).ToList();
                    iNumberOfRows = 1;
                    iNumberOfColumns = 1;
                    if (bFirstWorksheet == true)
                    {
                        oSLBOMDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "Resource Id Set On Op");
                        bFirstWorksheet = false;
                    }
                    else
                    {
                        oSLBOMDocument.AddWorksheet("Resource Id Set On Op");
                    }
                    //set column header
                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "JobNum");
                    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Parent Part Num");
                    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Parent Part Rev Num");
                    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Assembly Seq");
                    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Opr Seq");
                    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Op Code");
                    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Resource");
                    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Resource Is Set On Operation -- SOP IS TO ONLY SET THE RESOURCE GROUP");
                    oSLBOMDocument.SetCellStyle(iNumberOfRows++, iNumberOfColumns, oHighlightHeaderStyle);
                    oSLBOMDocument.SetColumnWidth(iNumberOfColumns++, 90);

                    foreach (JobOperation oJobOperation in oOperationsWithResourceIdSet)
                    {
                        oSLBOMDocument.SetCellValue(iNumberOfRows, 1, oJobOperation.JobNum);
                        oSLBOMDocument.SetCellValue(iNumberOfRows, 2, oJobOperation.ParentPartNum);
                        oSLBOMDocument.SetCellValue(iNumberOfRows, 3, oJobOperation.ParentRevNum);
                        oSLBOMDocument.SetCellValue(iNumberOfRows, 4, oJobOperation.AssemblySeq.ToString());
                        oSLBOMDocument.SetCellValue(iNumberOfRows, 5, oJobOperation.OperationSeq.ToString());
                        oSLBOMDocument.SetCellValue(iNumberOfRows, 6, oJobOperation.OpCode);
                        oSLBOMDocument.SetCellValue(iNumberOfRows++, 7, oJobOperation.ResourceId);

                        bDataInReport = true;
                    }
                }

                List<JobOperation> oOperationsWithoutDueDate = m_oJobOperations.Where(oItem => oItem.DueDate == DateTime.MinValue).ToList();
                if (oOperationsWithoutDueDate.Count > 0)
                {
                    // sort these by job number
                    oOperationsWithoutDueDate = oOperationsWithoutDueDate.OrderBy(oItem => oItem.JobNum).ToList();
                    iNumberOfRows = 1;
                    iNumberOfColumns = 1;
                    if (bFirstWorksheet == true)
                    {
                        oSLBOMDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "Op No Due Date");
                        bFirstWorksheet = false;
                    }
                    else
                    {
                        oSLBOMDocument.AddWorksheet("Op No Due Date");
                    }
                    //set column header
                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "JobNum");
                    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Parent Part Num");
                    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Parent Part Rev Num");
                    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Assembly Seq");
                    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Opr Seq");
                    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Op Code");
                    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Operation Due Date Not Set On Operation -- SOP IS TO SET OP DUE DATE FOR SCHEUDLING AND PLANNING");
                    oSLBOMDocument.SetCellStyle(iNumberOfRows++, iNumberOfColumns, oHighlightHeaderStyle);
                    oSLBOMDocument.SetColumnWidth(iNumberOfColumns++, 90);

                    foreach (JobOperation oJobOperation in oOperationsWithoutDueDate)
                    {
                        oSLBOMDocument.SetCellValue(iNumberOfRows, 1, oJobOperation.JobNum);
                        oSLBOMDocument.SetCellValue(iNumberOfRows, 2, oJobOperation.ParentPartNum);
                        oSLBOMDocument.SetCellValue(iNumberOfRows, 3, oJobOperation.ParentRevNum);
                        oSLBOMDocument.SetCellValue(iNumberOfRows, 4, oJobOperation.AssemblySeq.ToString());
                        oSLBOMDocument.SetCellValue(iNumberOfRows, 5, oJobOperation.OperationSeq.ToString());
                        oSLBOMDocument.SetCellValue(iNumberOfRows++, 6, oJobOperation.OpCode);

                        bDataInReport = true;
                    }
                }

                List<JobOperation> oOperationsWithoutStartDate = m_oJobOperations.Where(oItem => oItem.StartDate == DateTime.MinValue).ToList();
                if (oOperationsWithoutStartDate.Count > 0)
                {
                    // sort these by job number
                    oOperationsWithoutStartDate = oOperationsWithoutStartDate.OrderBy(oItem => oItem.JobNum).ToList();
                    iNumberOfRows = 1;
                    iNumberOfColumns = 1;
                    if (bFirstWorksheet == true)
                    {
                        oSLBOMDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "Op No Start Date");
                        bFirstWorksheet = false;
                    }
                    else
                    {
                        oSLBOMDocument.AddWorksheet("Op No Start Date");
                    }
                    //set column header
                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "JobNum");
                    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Parent Part Num");
                    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Parent Part Rev Num");
                    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Assembly Seq");
                    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Opr Seq");
                    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Op Code");
                    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Operation Start Date Not Set On Operation -- SOP IS TO SET OP START DATE FOR SCHEUDLING AND PLANNING");
                    oSLBOMDocument.SetCellStyle(iNumberOfRows++, iNumberOfColumns, oHighlightHeaderStyle);
                    oSLBOMDocument.SetColumnWidth(iNumberOfColumns++, 90);

                    foreach (JobOperation oJobOperation in oOperationsWithoutStartDate)
                    {
                        oSLBOMDocument.SetCellValue(iNumberOfRows, 1, oJobOperation.JobNum);
                        oSLBOMDocument.SetCellValue(iNumberOfRows, 2, oJobOperation.ParentPartNum);
                        oSLBOMDocument.SetCellValue(iNumberOfRows, 3, oJobOperation.ParentRevNum);
                        oSLBOMDocument.SetCellValue(iNumberOfRows, 4, oJobOperation.AssemblySeq.ToString());
                        oSLBOMDocument.SetCellValue(iNumberOfRows, 5, oJobOperation.OperationSeq.ToString());
                        oSLBOMDocument.SetCellValue(iNumberOfRows++, 6, oJobOperation.OpCode);

                        bDataInReport = true;
                    }
                }
            }

            // check to make sure the resource group is active
            List<JobOperation> oOperationsWithResourceGroup = m_oJobOperations.Where(oItem => (string.IsNullOrEmpty(oItem.ResourceGroupId) == false)).ToList();
            List<JobOperation> oOpsWithInactiveResourceGroups = new List<JobOperation>();
            foreach (JobOperation oJobOp in oOperationsWithResourceGroup)
            {
                if (m_oBOMSupport.IsResourceGroupActive(oJobOp.ResourceGroupId) == false)
                {
                    oOpsWithInactiveResourceGroups.Add(oJobOp);
                }
            }
            if (oOpsWithInactiveResourceGroups.Count > 0)
            {
                // sort these by job number
                oOpsWithInactiveResourceGroups = oOpsWithInactiveResourceGroups.OrderBy(oItem => oItem.JobNum).ToList();
                iNumberOfRows = 1;
                iNumberOfColumns = 1;
                if (bFirstWorksheet == true)
                {
                    oSLBOMDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "Op With Inactive Res Grp");
                    bFirstWorksheet = false;
                }
                else
                {
                    oSLBOMDocument.AddWorksheet("Op With Inactive Res Grp");
                }
                //set column header
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "JobNum");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Parent Part Num");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Parent Part Rev Num");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Assembly Seq");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Opr Seq");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Op Code");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Resource Group");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Operation Uses Inactive Resource Group -- POC ISSUE THIS JOB CANNOT BE SCHEUDLED WITH INACTIVE RESOURCE GROUP");
                oSLBOMDocument.SetCellStyle(iNumberOfRows++, iNumberOfColumns, oHighlightHeaderStyle);
                oSLBOMDocument.SetColumnWidth(iNumberOfColumns++, 90);

                foreach (JobOperation oJobOperation in oOpsWithInactiveResourceGroups)
                {
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 1, oJobOperation.JobNum);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 2, oJobOperation.ParentPartNum);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 3, oJobOperation.ParentRevNum);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 4, oJobOperation.AssemblySeq.ToString());
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 5, oJobOperation.OperationSeq.ToString());
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 6, oJobOperation.OpCode);
                    oSLBOMDocument.SetCellValue(iNumberOfRows++, 7, oJobOperation.ResourceGroupId);

                    bDataInReport = true;
                }
            }

            // check to make sure the resource is active
            List<JobOperation> oOperationsWithResources = m_oJobOperations.Where(oItem => (string.IsNullOrEmpty(oItem.ResourceId) == false)).ToList();
            List<JobOperation> oOpsWithInactiveResources = new List<JobOperation>();
            foreach (JobOperation oJobOp in oOperationsWithResources)
            {
                if (m_oBOMSupport.IsResourceActive(oJobOp.ResourceId) == false)
                {
                    oOpsWithInactiveResources.Add(oJobOp);
                }
            }
            if (oOpsWithInactiveResources.Count > 0)
            {
                // sort these by job number
                oOpsWithInactiveResources = oOpsWithInactiveResources.OrderBy(oItem => oItem.JobNum).ToList();
                iNumberOfRows = 1;
                iNumberOfColumns = 1;
                if (bFirstWorksheet == true)
                {
                    oSLBOMDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "Op With Inactive Res");
                    bFirstWorksheet = false;
                }
                else
                {
                    oSLBOMDocument.AddWorksheet("Op With Inactive Res");
                }
                //set column header
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "JobNum");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Parent Part Num");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Parent Part Rev Num");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Assembly Seq");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Opr Seq");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Op Code");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Resource Grouo");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Resource");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);

                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Operation Uses Inactive Resource -- POC ISSUE THIS JOB CANNOT BE SCHEUDLED WITH INACTIVE RESOURCES");
                oSLBOMDocument.SetCellStyle(iNumberOfRows++, iNumberOfColumns, oHighlightHeaderStyle);
                oSLBOMDocument.SetColumnWidth(iNumberOfColumns++, 90);

                foreach (JobOperation oJobOperation in oOpsWithInactiveResources)
                {
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 1, oJobOperation.JobNum);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 2, oJobOperation.ParentPartNum);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 3, oJobOperation.ParentRevNum);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 4, oJobOperation.AssemblySeq.ToString());
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 5, oJobOperation.OperationSeq.ToString());
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 6, oJobOperation.OpCode);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 7, oJobOperation.ResourceGroupId);
                    oSLBOMDocument.SetCellValue(iNumberOfRows++, 8, oJobOperation.ResourceId);

                    bDataInReport = true;
                }
            }

            // prod standard is zero
            List<JobOperation> oOperationsWithZeroEstimatedTime = m_oJobOperations.Where(oItem => (oItem.ProdStandard == 0) && (oItem.Subcontract == false)).ToList();
            if (oOperationsWithZeroEstimatedTime.Count > 0)
            {
                // sort these by job number
                oOperationsWithZeroEstimatedTime = oOperationsWithZeroEstimatedTime.OrderBy(oItem => oItem.JobNum).ToList();
                iNumberOfRows = 1;
                iNumberOfColumns = 1;
                if (bFirstWorksheet == true)
                {
                    oSLBOMDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "Op No Est Time");
                    bFirstWorksheet = false;
                }
                else
                {
                    oSLBOMDocument.AddWorksheet("Op No Est Time");
                }
                //set column header
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "JobNum");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Parent Part Num");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Parent Part Rev Num");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Assembly Seq");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Opr Seq");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Op Code");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Operation Has No Time Set -- POC ISSUE PLEASE SET HOURS FOR OPERATION");
                oSLBOMDocument.SetCellStyle(iNumberOfRows++, iNumberOfColumns, oHighlightHeaderStyle);
                oSLBOMDocument.SetColumnWidth(iNumberOfColumns++, 90);

                foreach (JobOperation oJobOperation in oOperationsWithZeroEstimatedTime)
                {
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 1, oJobOperation.JobNum);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 2, oJobOperation.ParentPartNum);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 3, oJobOperation.ParentRevNum);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 4, oJobOperation.AssemblySeq.ToString());
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 5, oJobOperation.OperationSeq.ToString());
                    oSLBOMDocument.SetCellValue(iNumberOfRows++, 6, oJobOperation.OpCode);

                    bDataInReport = true;
                }
            }

            // test if labor rate is set to zero
            // first get list of operations with time set and a run qty
            List<JobOperation> oOperationsWithHours = m_oJobOperations.Where(oItem => (oItem.ProdStandard != 0) && (oItem.RunQty != 0) && (oItem.Subcontract == false)).ToList();
            // get a list of operations where labor cost is zero
            List<JobOpsEstVsActualCosts> oOperationsWithNoLaborCost = m_oJobOpsEstVsActualCosts.Where(oItem => oItem.EstLaborCost == 0).ToList();
            // order by job num
            oOperationsWithNoLaborCost = oOperationsWithNoLaborCost.OrderBy(oItem => oItem.JobNum).ToList();
            // walk through this list of ops with no labor cost and if they have operational time then this means the labor rate is zero
            bool bSetHeader = false;
            foreach (JobOpsEstVsActualCosts oOpCost in oOperationsWithNoLaborCost)
            {
                JobOperation oJobOperation = oOperationsWithHours.FirstOrDefault(oItem => (string.Compare(oItem.JobNum, oOpCost.JobNum, true) == 0) && (oItem.AssemblySeq == oOpCost.AssemblySeq) && (oItem.OperationSeq == oOpCost.OprSeq) && (string.Compare(oItem.OpCode, oOpCost.OpCode, true) == 0) );
                if (oJobOperation != null)
                {
                    // this operation has hours but the labor comes out to $0 so that means we must have a labor rate of $0
                    if (bSetHeader == false)
                    {
                        bSetHeader = true;
                        iNumberOfRows = 1;
                        iNumberOfColumns = 1;
                        if (bFirstWorksheet == true)
                        {
                            oSLBOMDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "Ops With No Labor Rate");
                            bFirstWorksheet = false;
                        }
                        else
                        {
                            oSLBOMDocument.AddWorksheet("Ops With No Labor Rate");
                        }
                        //set column header
                        oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "JobNum");
                        oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                        oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Parent Part Num");
                        oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                        oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Parent Part Rev Num");
                        oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                        oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Assembly Seq");
                        oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                        oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Opr Seq");
                        oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                        oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Op Code");
                        oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                        oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Labor Rate Is $0 For Operation -- POC ISSUE PLEASE SET LABOR RATE ON OPERATION");
                        oSLBOMDocument.SetCellStyle(iNumberOfRows++, iNumberOfColumns, oHighlightHeaderStyle);
                        oSLBOMDocument.SetColumnWidth(iNumberOfColumns++, 90);
                    }

                    // indicate which operation has a zero dollar labor rate
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 1, oJobOperation.JobNum);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 2, oJobOperation.ParentPartNum);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 3, oJobOperation.ParentRevNum);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 4, oJobOperation.AssemblySeq.ToString());
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 5, oJobOperation.OperationSeq.ToString());
                    oSLBOMDocument.SetCellValue(iNumberOfRows++, 6, oJobOperation.OpCode);

                    bDataInReport = true;
                }
            }

            // test if burden rate is set to zero
            // we will use the list of operations with time set and a run qty from above
            // get a list of operations where burden cost is zero
            List<JobOpsEstVsActualCosts> oOperationsWithNoBurdenCost = m_oJobOpsEstVsActualCosts.Where(oItem => oItem.EstBurCost == 0).ToList();
            // walk through this list of ops with no burden cost and if they have operational time then this means the burden rate is zero
            bSetHeader = false;
            foreach (JobOpsEstVsActualCosts oOpCost in oOperationsWithNoBurdenCost)
            {
                JobOperation oJobOperation = oOperationsWithHours.FirstOrDefault(oItem => (string.Compare(oItem.JobNum, oOpCost.JobNum, true) == 0) && (oItem.AssemblySeq == oOpCost.AssemblySeq) && (oItem.OperationSeq == oOpCost.OprSeq) && (string.Compare(oItem.OpCode, oOpCost.OpCode, true) == 0));
                if (oJobOperation != null)
                {
                    // this operation has hours but the burden comes out to $0 so that means we must have a burden rate of $0
                    if (bSetHeader == false)
                    {
                        bSetHeader = true;
                        iNumberOfRows = 1;
                        iNumberOfColumns = 1;
                        if (bFirstWorksheet == true)
                        {
                            oSLBOMDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "Ops With No Burden Rate");
                            bFirstWorksheet = false;
                        }
                        else
                        {
                            oSLBOMDocument.AddWorksheet("Ops With No Burden Rate");
                        }
                        //set column header
                        oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "JobNum");
                        oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                        oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Parent Part Num");
                        oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                        oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Parent Part Rev Num");
                        oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                        oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Assembly Seq");
                        oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                        oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Opr Seq");
                        oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                        oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Op Code");
                        oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                        oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Burden Rate Is $0 For Operation -- POC ISSUE PLEASE SET BURDEN RATE ON OPERATION");
                        oSLBOMDocument.SetCellStyle(iNumberOfRows++, iNumberOfColumns, oHighlightHeaderStyle);
                        oSLBOMDocument.SetColumnWidth(iNumberOfColumns++, 90);
                    }

                    // indicate which operation has a zero dollar burden rate
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 1, oJobOperation.JobNum);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 2, oJobOperation.ParentPartNum);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 3, oJobOperation.ParentRevNum);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 4, oJobOperation.AssemblySeq.ToString());
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 5, oJobOperation.OperationSeq.ToString());
                    oSLBOMDocument.SetCellValue(iNumberOfRows++, 6, oJobOperation.OpCode);

                    bDataInReport = true;
                }
            }

            #region IGNORING THESE ISSUES
            // std format is wrong
            // standard format should be "HP" hours / piece
            //////////List<JobOperation> oOperationsWithBadStandardFormat = m_oJobOperations.Where(oItem => (string.Compare(oItem.StdFormat, "HP", true) != 0)).ToList();
            //////////if (oOperationsWithBadStandardFormat.Count > 0)
            //////////{
            //////////    // sort these by job number
            //////////    oOperationsWithBadStandardFormat = oOperationsWithBadStandardFormat.OrderBy(oItem => oItem.JobNum).ToList();
            //////////    iNumberOfRows = 1;
            //////////    iNumberOfColumns = 1;
            //////////    if (bFirstWorksheet == true)
            //////////    {
            //////////        oSLBOMDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "Op Bad Std Format");
            //////////        bFirstWorksheet = false;
            //////////    }
            //////////    else
            //////////    {
            //////////        oSLBOMDocument.AddWorksheet("Op Bad Std Format");
            //////////    }
            //////////    //set column header
            //////////    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "JobNum");
            //////////    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
            //////////    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Parent Part Num");
            //////////    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
            //////////    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Parent Part Rev Num");
            //////////    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
            //////////    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Assembly Seq");
            //////////    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
            //////////    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Opr Seq");
            //////////    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
            //////////    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Op Code");
            //////////    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
            //////////    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Std Format");
            //////////    oSLBOMDocument.SetColumnWidth(iNumberOfRows++, iNumberOfColumns++, iStandardColumnWidth);

            //////////    foreach (JobOperation oJobOperation in oOperationsWithBadStandardFormat)
            //////////    {
            //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, 1, oJobOperation.JobNum);
            //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, 2, oJobOperation.ParentPartNum);
            //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, 3, oJobOperation.ParentRevNum);
            //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, 4, oJobOperation.AssemblySeq.ToString());
            //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, 5, oJobOperation.OperationSeq.ToString());
            //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, 6, oJobOperation.OpCode);
            //////////        oSLBOMDocument.SetCellValue(iNumberOfRows++, 7, oJobOperation.StdFormat);

            //////////        bDataInReport = true;
            //////////    }
            //////////}
            #endregion

            if (m_bJustMissingCosts == false)
            {
                // if std format is "OM" or "OH" then the OpsPerPart field must be > 0
                List<JobOperation> oOperationsPerPartSetToZero = m_oJobOperations.Where(oItem => ((string.Compare(oItem.StdFormat, "OM", true) == 0) || (string.Compare(oItem.StdFormat, "OH", true) == 0)) && (oItem.OperationsPerPart == 0)).ToList();
                if (oOperationsPerPartSetToZero.Count > 0)
                {
                    // sort these by job number
                    oOperationsPerPartSetToZero = oOperationsPerPartSetToZero.OrderBy(oItem => oItem.JobNum).ToList();
                    iNumberOfRows = 1;
                    iNumberOfColumns = 1;
                    if (bFirstWorksheet == true)
                    {
                        oSLBOMDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "Ops Per Part Is Zero");
                        bFirstWorksheet = false;
                    }
                    else
                    {
                        oSLBOMDocument.AddWorksheet("Ops Per Part Is Zero");
                    }
                    //set column header
                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "JobNum");
                    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Parent Part Num");
                    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Parent Part Rev Num");
                    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Assembly Seq");
                    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Opr Seq");
                    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Op Code");
                    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Std Format");
                    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Operation Is Set To Operations Per Part But Is Missing Operation Qty -- POC ISSUE PLEASE SET OPERATION QTY");
                    oSLBOMDocument.SetCellStyle(iNumberOfRows++, iNumberOfColumns, oHighlightHeaderStyle);
                    oSLBOMDocument.SetColumnWidth(iNumberOfColumns++, 90);

                    foreach (JobOperation oJobOperation in oOperationsPerPartSetToZero)
                    {
                        oSLBOMDocument.SetCellValue(iNumberOfRows, 1, oJobOperation.JobNum);
                        oSLBOMDocument.SetCellValue(iNumberOfRows, 2, oJobOperation.ParentPartNum);
                        oSLBOMDocument.SetCellValue(iNumberOfRows, 3, oJobOperation.ParentRevNum);
                        oSLBOMDocument.SetCellValue(iNumberOfRows, 4, oJobOperation.AssemblySeq.ToString());
                        oSLBOMDocument.SetCellValue(iNumberOfRows, 5, oJobOperation.OperationSeq.ToString());
                        oSLBOMDocument.SetCellValue(iNumberOfRows, 6, oJobOperation.OpCode);
                        oSLBOMDocument.SetCellValue(iNumberOfRows++, 7, oJobOperation.StdFormat);

                        bDataInReport = true;
                    }
                }

                // if std format is anything other than "OM" or "OH" then the OpsPerPart field should be zero
                List<JobOperation> oOperationsPerPartNotSetToZero = m_oJobOperations.Where(oItem => (string.Compare(oItem.StdFormat, "OM", true) != 0) && (string.Compare(oItem.StdFormat, "OH", true) != 0) && (oItem.OperationsPerPart != 0)).ToList();
                if (oOperationsPerPartNotSetToZero.Count > 0)
                {
                    // sort these by job number
                    oOperationsPerPartNotSetToZero = oOperationsPerPartNotSetToZero.OrderBy(oItem => oItem.JobNum).ToList();
                    iNumberOfRows = 1;
                    iNumberOfColumns = 1;
                    if (bFirstWorksheet == true)
                    {
                        oSLBOMDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "Ops Per Part Not Zero");
                        bFirstWorksheet = false;
                    }
                    else
                    {
                        oSLBOMDocument.AddWorksheet("Ops Per Part Not Zero");
                    }
                    //set column header
                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "JobNum");
                    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Parent Part Num");
                    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Parent Part Rev Num");
                    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Assembly Seq");
                    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Opr Seq");
                    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Op Code");
                    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Std Format");
                    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Ops Per Part");
                    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Operations Per Part Is Not Set To Zero -- SOP IS TO SET OPERATIONS PER PART TO ZERO");
                    oSLBOMDocument.SetCellStyle(iNumberOfRows++, iNumberOfColumns, oHighlightHeaderStyle);
                    oSLBOMDocument.SetColumnWidth(iNumberOfColumns++, 90);

                    foreach (JobOperation oJobOperation in oOperationsPerPartNotSetToZero)
                    {
                        oSLBOMDocument.SetCellValue(iNumberOfRows, 1, oJobOperation.JobNum);
                        oSLBOMDocument.SetCellValue(iNumberOfRows, 2, oJobOperation.ParentPartNum);
                        oSLBOMDocument.SetCellValue(iNumberOfRows, 3, oJobOperation.ParentRevNum);
                        oSLBOMDocument.SetCellValue(iNumberOfRows, 4, oJobOperation.AssemblySeq.ToString());
                        oSLBOMDocument.SetCellValue(iNumberOfRows, 5, oJobOperation.OperationSeq.ToString());
                        oSLBOMDocument.SetCellValue(iNumberOfRows, 6, oJobOperation.OpCode);
                        oSLBOMDocument.SetCellValue(iNumberOfRows, 7, oJobOperation.StdFormat);
                        oSLBOMDocument.SetCellValue(iNumberOfRows++, 8, oJobOperation.OperationsPerPart);

                        bDataInReport = true;
                    }
                }

                // std basis is wrong
                List<JobOperation> oOperationsStdBasisWrong = m_oJobOperations.Where(oItem => (string.Compare(oItem.StdBasis, "E", true) != 0) && (oItem.Subcontract == false)).ToList();
                if (oOperationsStdBasisWrong.Count > 0)
                {
                    // sort these by job number
                    oOperationsStdBasisWrong = oOperationsStdBasisWrong.OrderBy(oItem => oItem.JobNum).ToList();
                    iNumberOfRows = 1;
                    iNumberOfColumns = 1;
                    if (bFirstWorksheet == true)
                    {
                        oSLBOMDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "Op Bad Std Basis");
                        bFirstWorksheet = false;
                    }
                    else
                    {
                        oSLBOMDocument.AddWorksheet("Op Bad Std Basis");
                    }
                    //set column header
                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "JobNum");
                    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Parent Part Num");
                    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Parent Part Rev Num");
                    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Assembly Seq");
                    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Opr Seq");
                    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Op Code");
                    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Std Basis");
                    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Operation Standard Basis Is Not Set To Each -- SOP IS TO SET THIS TO EACH NOT PER HUNDRED OR PER THOUSAND");
                    oSLBOMDocument.SetCellStyle(iNumberOfRows++, iNumberOfColumns, oHighlightHeaderStyle);
                    oSLBOMDocument.SetColumnWidth(iNumberOfColumns++, 90);

                    foreach (JobOperation oJobOperation in oOperationsStdBasisWrong)
                    {
                        oSLBOMDocument.SetCellValue(iNumberOfRows, 1, oJobOperation.JobNum);
                        oSLBOMDocument.SetCellValue(iNumberOfRows, 2, oJobOperation.ParentPartNum);
                        oSLBOMDocument.SetCellValue(iNumberOfRows, 3, oJobOperation.ParentRevNum);
                        oSLBOMDocument.SetCellValue(iNumberOfRows, 4, oJobOperation.AssemblySeq.ToString());
                        oSLBOMDocument.SetCellValue(iNumberOfRows, 5, oJobOperation.OperationSeq.ToString());
                        oSLBOMDocument.SetCellValue(iNumberOfRows, 6, oJobOperation.OpCode);
                        oSLBOMDocument.SetCellValue(iNumberOfRows++, 7, oJobOperation.StdBasis);

                        bDataInReport = true;
                    }
                }
            }

            // labor entry method not correct
            List<JobOperation> oOperationsBadLaborEntryMethod = m_oJobOperations.Where(oItem => (string.Compare(oItem.LaborEntryMethod, "T", true) != 0)).ToList();
            if (oOperationsBadLaborEntryMethod.Count > 0)
            {
                // sort these by job number
                oOperationsBadLaborEntryMethod = oOperationsBadLaborEntryMethod.OrderBy(oItem => oItem.JobNum).ToList();
                iNumberOfRows = 1;
                iNumberOfColumns = 1;
                if (bFirstWorksheet == true)
                {
                    oSLBOMDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "Op Bad Labor Entry Method");
                    bFirstWorksheet = false;
                }
                else
                {
                    oSLBOMDocument.AddWorksheet("Op Bad Labor Entry Method");
                }
                //set column header
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "JobNum");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Parent Part Num");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Parent Part Rev Num");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Assembly Seq");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Opr Seq");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Op Code");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Labor Entry");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Labor Entry Method Not Set To Time And Quantity -- SOP IS TO ALWAYS USE TIME AND QUANTITY WHEN REPORTING ON JOB");
                oSLBOMDocument.SetCellStyle(iNumberOfRows++, iNumberOfColumns, oHighlightHeaderStyle);
                oSLBOMDocument.SetColumnWidth(iNumberOfColumns++, 90);

                foreach (JobOperation oJobOperation in oOperationsBadLaborEntryMethod)
                {
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 1, oJobOperation.JobNum);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 2, oJobOperation.ParentPartNum);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 3, oJobOperation.ParentRevNum);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 4, oJobOperation.AssemblySeq.ToString());
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 5, oJobOperation.OperationSeq.ToString());
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 6, oJobOperation.OpCode);
                    oSLBOMDocument.SetCellValue(iNumberOfRows++, 7, oJobOperation.LaborEntryMethod);

                    bDataInReport = true;
                }
            }

            #region IGNORING THESE ISSUES
            // if this is a subcontract we should include the days out
            //////////List<JobOperation> oSubcontractOperationNoDaysOutSet = m_oJobOperations.Where(oItem => (oItem.Subcontract == true) && (oItem.DaysOut == 0)).ToList();
            //////////if (oSubcontractOperationNoDaysOutSet.Count > 0)
            //////////{
            //////////    // sort these by job number
            //////////    oSubcontractOperationNoDaysOutSet = oSubcontractOperationNoDaysOutSet.OrderBy(oItem => oItem.JobNum).ToList();
            //////////    iNumberOfRows = 1;
            //////////    iNumberOfColumns = 1;
            //////////    if (bFirstWorksheet == true)
            //////////    {
            //////////        oSLBOMDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "Sub Op No Days Out Set");
            //////////        bFirstWorksheet = false;
            //////////    }
            //////////    else
            //////////    {
            //////////        oSLBOMDocument.AddWorksheet("Sub Op No Days Out Set");
            //////////    }
            //////////    //set column header
            //////////    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "JobNum");
            //////////    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
            //////////    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Parent Part Num");
            //////////    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
            //////////    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Parent Part Rev Num");
            //////////    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
            //////////    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Assembly Seq");
            //////////    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
            //////////    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Opr Seq");
            //////////    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
            //////////    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Op Code");
            //////////    oSLBOMDocument.SetColumnWidth(iNumberOfRows++, iNumberOfColumns++, iStandardColumnWidth);

            //////////    foreach (JobOperation oJobOperation in oOperationsBadLaborEntryMethod)
            //////////    {
            //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, 1, oJobOperation.JobNum);
            //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, 2, oJobOperation.ParentPartNum);
            //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, 3, oJobOperation.ParentRevNum);
            //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, 4, oJobOperation.AssemblySeq.ToString());
            //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, 5, oJobOperation.OperationSeq.ToString());
            //////////        oSLBOMDocument.SetCellValue(iNumberOfRows++, 6, oJobOperation.OpCode);

            //////////        bDataInReport = true;
            //////////    }
            //////////}

            // sub op with no vendor set
            //////////List<JobOperation> oSubcontractOperationNoVendor = m_oJobOperations.Where(oItem => (oItem.Subcontract == true) && (oItem.VendorNum == 0)).ToList();
            //////////if (oSubcontractOperationNoVendor.Count > 0)
            //////////{
            //////////    // sort these by job number
            //////////    oSubcontractOperationNoVendor = oSubcontractOperationNoVendor.OrderBy(oItem => oItem.JobNum).ToList();
            //////////    iNumberOfRows = 1;
            //////////    iNumberOfColumns = 1;
            //////////    if (bFirstWorksheet == true)
            //////////    {
            //////////        oSLBOMDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "Sub Op No Vendor");
            //////////        bFirstWorksheet = false;
            //////////    }
            //////////    else
            //////////    {
            //////////        oSLBOMDocument.AddWorksheet("Sub Op No Vendor");
            //////////    }
            //////////    //set column header
            //////////    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "JobNum");
            //////////    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
            //////////    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Parent Part Num");
            //////////    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
            //////////    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Parent Part Rev Num");
            //////////    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
            //////////    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Assembly Seq");
            //////////    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
            //////////    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Opr Seq");
            //////////    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
            //////////    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Op Code");
            //////////    oSLBOMDocument.SetColumnWidth(iNumberOfRows++, iNumberOfColumns++, iStandardColumnWidth);

            //////////    foreach (JobOperation oJobOperation in oSubcontractOperationNoVendor)
            //////////    {
            //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, 1, oJobOperation.JobNum);
            //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, 2, oJobOperation.ParentPartNum);
            //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, 3, oJobOperation.ParentRevNum);
            //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, 4, oJobOperation.AssemblySeq.ToString());
            //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, 5, oJobOperation.OperationSeq.ToString());
            //////////        oSLBOMDocument.SetCellValue(iNumberOfRows++, 6, oJobOperation.OpCode);

            //////////        bDataInReport = true;
            //////////    }
            //////////}
            #endregion

            if (m_bJustMissingCosts == false)
            {
                // qty per should be set for subcontract
                List<JobOperation> oSubcontractOperationZeroQtyPer = m_oJobOperations.Where(oItem => (oItem.Subcontract == true) && (oItem.QtyPer == 0)).ToList();
                if (oSubcontractOperationZeroQtyPer.Count > 0)
                {
                    // sort these by job number
                    oSubcontractOperationZeroQtyPer = oSubcontractOperationZeroQtyPer.OrderBy(oItem => oItem.JobNum).ToList();
                    iNumberOfRows = 1;
                    iNumberOfColumns = 1;
                    if (bFirstWorksheet == true)
                    {
                        oSLBOMDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "Sub Op Zero Qty Per");
                        bFirstWorksheet = false;
                    }
                    else
                    {
                        oSLBOMDocument.AddWorksheet("Sub Op Zero Qty Per");
                    }
                    //set column header
                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "JobNum");
                    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Parent Part Num");
                    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Parent Part Rev Num");
                    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Assembly Seq");
                    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Opr Seq");
                    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Op Code");
                    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Subcontract Has Quantity Set To Zero -- POC ISSUE PLEASE SET SUBCONTRACT QUANTITY");
                    oSLBOMDocument.SetCellStyle(iNumberOfRows++, iNumberOfColumns, oHighlightHeaderStyle);
                    oSLBOMDocument.SetColumnWidth(iNumberOfColumns++, 90);

                    foreach (JobOperation oJobOperation in oSubcontractOperationZeroQtyPer)
                    {
                        oSLBOMDocument.SetCellValue(iNumberOfRows, 1, oJobOperation.JobNum);
                        oSLBOMDocument.SetCellValue(iNumberOfRows, 2, oJobOperation.ParentPartNum);
                        oSLBOMDocument.SetCellValue(iNumberOfRows, 3, oJobOperation.ParentRevNum);
                        oSLBOMDocument.SetCellValue(iNumberOfRows, 4, oJobOperation.AssemblySeq.ToString());
                        oSLBOMDocument.SetCellValue(iNumberOfRows, 5, oJobOperation.OperationSeq.ToString());
                        oSLBOMDocument.SetCellValue(iNumberOfRows++, 6, oJobOperation.OpCode);

                        bDataInReport = true;
                    }
                }
            }

            #region IGNORING THESE ISSUES
            // burden does not equals labor
            //////////List<JobOperation> oOperationWithBurdenNotEqualToLabor = m_oJobOperations.Where(oItem => (oItem.BurdenEqualsLabor == false)).ToList();
            //////////if (oOperationWithBurdenNotEqualToLabor.Count > 0)
            //////////{
            //////////    // sort these by job number
            //////////    oOperationWithBurdenNotEqualToLabor = oOperationWithBurdenNotEqualToLabor.OrderBy(oItem => oItem.JobNum).ToList();
            //////////    iNumberOfRows = 1;
            //////////    iNumberOfColumns = 1;
            //////////    if (bFirstWorksheet == true)
            //////////    {
            //////////        oSLBOMDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "Burden Not Eq Labor");
            //////////        bFirstWorksheet = false;
            //////////    }
            //////////    else
            //////////    {
            //////////        oSLBOMDocument.AddWorksheet("Burden Not Eq Labor");
            //////////    }
            //////////    //set column header
            //////////    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "JobNum");
            //////////    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
            //////////    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Parent Part Num");
            //////////    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
            //////////    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Parent Part Rev Num");
            //////////    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
            //////////    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Assembly Seq");
            //////////    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
            //////////    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Opr Seq");
            //////////    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
            //////////    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Op Code");
            //////////    oSLBOMDocument.SetColumnWidth(iNumberOfRows++, iNumberOfColumns++, iStandardColumnWidth);

            //////////    foreach (JobOperation oJobOperation in oOperationWithBurdenNotEqualToLabor)
            //////////    {
            //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, 1, oJobOperation.JobNum);
            //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, 2, oJobOperation.ParentPartNum);
            //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, 3, oJobOperation.ParentRevNum);
            //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, 4, oJobOperation.AssemblySeq.ToString());
            //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, 5, oJobOperation.OperationSeq.ToString());
            //////////        oSLBOMDocument.SetCellValue(iNumberOfRows++, 6, oJobOperation.OpCode);

            //////////        bDataInReport = true;
            //////////    }
            //////////}
            #endregion

            if (m_bJustMissingCosts == false)
            {
                // use estimates
                List<JobOperation> oOperationUseEstimatesSet = m_oJobOperations.Where(oItem => (oItem.UseEstimates == true)).ToList();
                if (oOperationUseEstimatesSet.Count > 0)
                {
                    // sort these by job number
                    oOperationUseEstimatesSet = oOperationUseEstimatesSet.OrderBy(oItem => oItem.JobNum).ToList();
                    iNumberOfRows = 1;
                    iNumberOfColumns = 1;
                    if (bFirstWorksheet == true)
                    {
                        oSLBOMDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "Use Estimates Set");
                        bFirstWorksheet = false;
                    }
                    else
                    {
                        oSLBOMDocument.AddWorksheet("Use Estimates Set");
                    }
                    //set column header
                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "JobNum");
                    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Parent Part Num");
                    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Parent Part Rev Num");
                    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Assembly Seq");
                    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Opr Seq");
                    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Op Code");
                    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Operation Has Use Estimates Set -- SOP IS TO NOT USE THIS FEATURE");
                    oSLBOMDocument.SetCellStyle(iNumberOfRows++, iNumberOfColumns, oHighlightHeaderStyle);
                    oSLBOMDocument.SetColumnWidth(iNumberOfColumns++, 90);

                    foreach (JobOperation oJobOperation in oOperationUseEstimatesSet)
                    {
                        oSLBOMDocument.SetCellValue(iNumberOfRows, 1, oJobOperation.JobNum);
                        oSLBOMDocument.SetCellValue(iNumberOfRows, 2, oJobOperation.ParentPartNum);
                        oSLBOMDocument.SetCellValue(iNumberOfRows, 3, oJobOperation.ParentRevNum);
                        oSLBOMDocument.SetCellValue(iNumberOfRows, 4, oJobOperation.AssemblySeq.ToString());
                        oSLBOMDocument.SetCellValue(iNumberOfRows, 5, oJobOperation.OperationSeq.ToString());
                        oSLBOMDocument.SetCellValue(iNumberOfRows++, 6, oJobOperation.OpCode);

                        bDataInReport = true;
                    }
                }

                // split operations
                List<JobOperation> oSplitOperationsSet = m_oJobOperations.Where(oItem => (oItem.SplitOperations == true)).ToList();
                if (oSplitOperationsSet.Count > 0)
                {
                    // sort these by job number
                    oSplitOperationsSet = oSplitOperationsSet.OrderBy(oItem => oItem.JobNum).ToList();
                    iNumberOfRows = 1;
                    iNumberOfColumns = 1;
                    if (bFirstWorksheet == true)
                    {
                        oSLBOMDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "Split Operations Set");
                        bFirstWorksheet = false;
                    }
                    else
                    {
                        oSLBOMDocument.AddWorksheet("Split Operations Set");
                    }
                    //set column header
                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "JobNum");
                    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Parent Part Num");
                    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Parent Part Rev Num");
                    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Assembly Seq");
                    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Opr Seq");
                    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Op Code");
                    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Operation Allows Work To Be Split -- SOP DOES NOT PERMIT OPERATIONS TO BE SET TO BE SPLIT");
                    oSLBOMDocument.SetCellStyle(iNumberOfRows++, iNumberOfColumns, oHighlightHeaderStyle);
                    oSLBOMDocument.SetColumnWidth(iNumberOfColumns++, 90);

                    foreach (JobOperation oJobOperation in oSplitOperationsSet)
                    {
                        oSLBOMDocument.SetCellValue(iNumberOfRows, 1, oJobOperation.JobNum);
                        oSLBOMDocument.SetCellValue(iNumberOfRows, 2, oJobOperation.ParentPartNum);
                        oSLBOMDocument.SetCellValue(iNumberOfRows, 3, oJobOperation.ParentRevNum);
                        oSLBOMDocument.SetCellValue(iNumberOfRows, 4, oJobOperation.AssemblySeq.ToString());
                        oSLBOMDocument.SetCellValue(iNumberOfRows, 5, oJobOperation.OperationSeq.ToString());
                        oSLBOMDocument.SetCellValue(iNumberOfRows++, 6, oJobOperation.OpCode);

                        bDataInReport = true;
                    }
                }

                // split burden
                List<JobOperation> oSplitBurdenSet = m_oJobOperations.Where(oItem => (oItem.SplitBurden == true)).ToList();
                if (oSplitOperationsSet.Count > 0)
                {
                    // sort these by job number
                    oSplitBurdenSet = oSplitBurdenSet.OrderBy(oItem => oItem.JobNum).ToList();
                    iNumberOfRows = 1;
                    iNumberOfColumns = 1;
                    if (bFirstWorksheet == true)
                    {
                        oSLBOMDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "Split Burden Set");
                        bFirstWorksheet = false;
                    }
                    else
                    {
                        oSLBOMDocument.AddWorksheet("Split Burden Set");
                    }
                    //set column header
                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "JobNum");
                    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Parent Part Num");
                    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Parent Part Rev Num");
                    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Assembly Seq");
                    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Opr Seq");
                    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Op Code");
                    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Operation Is Set To Split Burden Across Resources -- SOP IS TO NOT ALLOW SPLIT BURDEN");
                    oSLBOMDocument.SetCellStyle(iNumberOfRows++, iNumberOfColumns, oHighlightHeaderStyle);
                    oSLBOMDocument.SetColumnWidth(iNumberOfColumns++, 90);

                    foreach (JobOperation oJobOperation in oSplitBurdenSet)
                    {
                        oSLBOMDocument.SetCellValue(iNumberOfRows, 1, oJobOperation.JobNum);
                        oSLBOMDocument.SetCellValue(iNumberOfRows, 2, oJobOperation.ParentPartNum);
                        oSLBOMDocument.SetCellValue(iNumberOfRows, 3, oJobOperation.ParentRevNum);
                        oSLBOMDocument.SetCellValue(iNumberOfRows, 4, oJobOperation.AssemblySeq.ToString());
                        oSLBOMDocument.SetCellValue(iNumberOfRows, 5, oJobOperation.OperationSeq.ToString());
                        oSLBOMDocument.SetCellValue(iNumberOfRows++, 6, oJobOperation.OpCode);

                        bDataInReport = true;
                    }
                }
            }

            // subcontract estimate is zero
            List<JobOperation> oSubOperationZeroEst = m_oJobOperations.Where(oItem => (oItem.Subcontract == true) && (oItem.EstUnitCost == 0)).ToList();
            if (oSubOperationZeroEst.Count > 0)
            {
                // sort these by job number
                oSubOperationZeroEst = oSubOperationZeroEst.OrderBy(oItem => oItem.JobNum).ToList();
                iNumberOfRows = 1;
                iNumberOfColumns = 1;
                if (bFirstWorksheet == true)
                {
                    oSLBOMDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "Sub Op Zero Est");
                    bFirstWorksheet = false;
                }
                else
                {
                    oSLBOMDocument.AddWorksheet("Sub Op Zero Est");
                }
                //set column header
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "JobNum");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Parent Part Num");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Parent Part Rev Num");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Assembly Seq");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Opr Seq");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Op Code");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Subcontract Estimate is Zero -- POC ISSUE PLEASE SET THE SUBCONTRACT ESTIMATE");
                oSLBOMDocument.SetCellStyle(iNumberOfRows++, iNumberOfColumns, oHighlightHeaderStyle);
                oSLBOMDocument.SetColumnWidth(iNumberOfColumns++, 90);

                foreach (JobOperation oJobOperation in oSubOperationZeroEst)
                {
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 1, oJobOperation.JobNum);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 2, oJobOperation.ParentPartNum);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 3, oJobOperation.ParentRevNum);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 4, oJobOperation.AssemblySeq.ToString());
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 5, oJobOperation.OperationSeq.ToString());
                    oSLBOMDocument.SetCellValue(iNumberOfRows++, 6, oJobOperation.OpCode);

                    bDataInReport = true;
                }
            }
            #endregion

            #region Operation Costs

            #region IGNORING THESE ISSUES
            //////////if (m_bJustMissingCosts == false)
            //////////{
            //////////    // unexpected hours -- if we are over the estimated amount or if we are less then the estimated amount and the op is completed
            //////////    List<JobOpsEstVsActualCosts> oJobOperationHoursDiffers = m_oJobOpsEstVsActualCosts.Where(oItem => (oItem.ActOprHours > oItem.EstOprHours) || ((oItem.ActOprHours < oItem.EstOprHours) && (oItem.OpComplete == true))).ToList();
            //////////    if (oJobOperationHoursDiffers.Count > 0)
            //////////    {
            //////////        // sort these by job number
            //////////        oJobOperationHoursDiffers = oJobOperationHoursDiffers.OrderBy(oItem => oItem.JobNum).ToList();
            //////////        iNumberOfRows = 1;
            //////////        iNumberOfColumns = 1;
            //////////        if (bFirstWorksheet == true)
            //////////        {
            //////////            oSLBOMDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "Op Hours");
            //////////            bFirstWorksheet = false;
            //////////        }
            //////////        else
            //////////        {
            //////////            oSLBOMDocument.AddWorksheet("Op Hours");
            //////////        }
            //////////        //set column header
            //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "JobNum");
            //////////        oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
            //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Parent Part Num");
            //////////        oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
            //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Parent Part Rev Num");
            //////////        oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
            //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Assembly Seq");
            //////////        oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
            //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Opr Seq");
            //////////        oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
            //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Opr Code");
            //////////        oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
            //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Est Hours");
            //////////        oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
            //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Act Hours");
            //////////        oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
            //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Variance Qty");
            //////////        oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
            //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Variance %");
            //////////        oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
            //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Operation Hours Differs From Estimate -- POC ISSUE PLEASE REVIEW");
            //////////        oSLBOMDocument.SetCellStyle(iNumberOfRows++, iNumberOfColumns, oHighlightHeaderStyle);
            //////////        oSLBOMDocument.SetColumnWidth(iNumberOfColumns++, 90);

            //////////        foreach (JobOpsEstVsActualCosts oJobOperation in oJobOperationHoursDiffers)
            //////////        {
            //////////            decimal dDelta = oJobOperation.ActOprHours - oJobOperation.EstOprHours;
            //////////            decimal dPercentageError = 0M;
            //////////            decimal dMarginToDate = 0M;

            //////////            if (oJobOperation.EstOprHours != 0)
            //////////            {
            //////////                dPercentageError = dDelta / oJobOperation.EstOprHours * 100.0M;
            //////////            }
            //////////            else
            //////////            {
            //////////                dPercentageError = 100.0M;
            //////////            }
            //////////            HSJob oTmpJob = m_oAllJobs.FirstOrDefault(oItem => string.Compare(oJobOperation.JobNum, oItem.JobNum, true) == 0);
            //////////            if (oTmpJob != null)
            //////////            {
            //////////                dMarginToDate = oTmpJob.MarginToDate * 100.0M;
            //////////            }

            //////////            bool bExceedsThreholds = false;

            //////////            // HERE WE JUST CHECK % ERROR AS WE HAVE NO $$$$
            //////////            if ((Math.Abs(dPercentageError) > m_dJobEstimatePercentError) && (dMarginToDate < m_dJobMarginThreshold))
            //////////            {
            //////////                // check to see if it exceeds both thresholds
            //////////                bExceedsThreholds = true;
            //////////            }

            //////////            if (bExceedsThreholds == true)
            //////////            {
            //////////                oSLBOMDocument.SetCellValue(iNumberOfRows, 1, oJobOperation.JobNum);
            //////////                oSLBOMDocument.SetCellValue(iNumberOfRows, 2, oJobOperation.PartNum);
            //////////                oSLBOMDocument.SetCellValue(iNumberOfRows, 3, oJobOperation.PartRevNum);
            //////////                oSLBOMDocument.SetCellValue(iNumberOfRows, 4, oJobOperation.AssemblySeq.ToString());
            //////////                oSLBOMDocument.SetCellValue(iNumberOfRows, 5, oJobOperation.OprSeq.ToString());
            //////////                oSLBOMDocument.SetCellValue(iNumberOfRows, 6, oJobOperation.OpCode);
            //////////                oSLBOMDocument.SetCellValue(iNumberOfRows, 7, oJobOperation.EstOprHours);
            //////////                oSLBOMDocument.SetCellStyle(iNumberOfRows, 7, oDecimalStyle);
            //////////                oSLBOMDocument.SetCellValue(iNumberOfRows, 8, oJobOperation.ActOprHours);
            //////////                oSLBOMDocument.SetCellStyle(iNumberOfRows, 8, oDecimalStyle);

            //////////                oSLBOMDocument.SetCellValue(iNumberOfRows, 9, dDelta);
            //////////                oSLBOMDocument.SetCellStyle(iNumberOfRows, 9, oDecimalStyle);

            //////////                oSLBOMDocument.SetCellValue(iNumberOfRows, 10, dPercentageError);
            //////////                oSLBOMDocument.SetCellStyle(iNumberOfRows++, 10, oDecimalStyle);
            //////////                bDataInReport = true;
            //////////            }
            //////////        }
            //////////    }

            //////////    // est labor cost differs from actual
            //////////    List<JobOpsEstVsActualCosts> oJobOperationLaborDiffers = m_oJobOpsEstVsActualCosts.Where(oItem => (oItem.ActLaborCost > oItem.EstLaborCost) || ((oItem.ActLaborCost < oItem.EstLaborCost) && (oItem.OpComplete == true))).ToList();
            //////////    if (oJobOperationLaborDiffers.Count > 0)
            //////////    {
            //////////        // sort these by job number
            //////////        oJobOperationLaborDiffers = oJobOperationLaborDiffers.OrderBy(oItem => oItem.JobNum).ToList();
            //////////        iNumberOfRows = 1;
            //////////        iNumberOfColumns = 1;
            //////////        if (bFirstWorksheet == true)
            //////////        {
            //////////            oSLBOMDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "Op Labor");
            //////////            bFirstWorksheet = false;
            //////////        }
            //////////        else
            //////////        {
            //////////            oSLBOMDocument.AddWorksheet("Op Labor");
            //////////        }
            //////////        //set column header
            //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "JobNum");
            //////////        oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
            //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Parent Part Num");
            //////////        oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
            //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Parent Part Rev Num");
            //////////        oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
            //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Assembly Seq");
            //////////        oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
            //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Opr Seq");
            //////////        oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
            //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Opr Code");
            //////////        oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
            //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Est Labor");
            //////////        oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
            //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Act Labor");
            //////////        oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
            //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Variance $");
            //////////        oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
            //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Variance %");
            //////////        oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
            //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Operation Labor Cost Differs From Estimate -- POC ISSUE PLEASE REVIEW");
            //////////        oSLBOMDocument.SetCellStyle(iNumberOfRows++, iNumberOfColumns, oHighlightHeaderStyle);
            //////////        oSLBOMDocument.SetColumnWidth(iNumberOfColumns++, 90);

            //////////        foreach (JobOpsEstVsActualCosts oJobOperation in oJobOperationLaborDiffers)
            //////////        {
            //////////            decimal dDelta = oJobOperation.ActLaborCost - oJobOperation.EstLaborCost;
            //////////            decimal dPercentageError = 0M;
            //////////            decimal dMarginToDate = 0M;
            //////////            if (oJobOperation.EstLaborCost != 0)
            //////////            {
            //////////                dPercentageError = dDelta / oJobOperation.EstLaborCost * 100.0M;
            //////////            }
            //////////            else
            //////////            {
            //////////                dPercentageError = 100.0M;
            //////////            }
            //////////            HSJob oTmpJob = m_oAllJobs.FirstOrDefault(oItem => string.Compare(oJobOperation.JobNum, oItem.JobNum, true) == 0);
            //////////            if (oTmpJob != null)
            //////////            {
            //////////                dMarginToDate = oTmpJob.MarginToDate * 100.0M;
            //////////            }

            //////////            bool bExceedsThreholds = false;

            //////////            if ((Math.Abs(dDelta) > m_dJobEstimateAbsoluteError) && (Math.Abs(dPercentageError) > m_dJobEstimatePercentError) && (dMarginToDate < m_dJobMarginThreshold))
            //////////            {
            //////////                // check to see if it exceeds both thresholds
            //////////                bExceedsThreholds = true;
            //////////            }
            //////////            if (m_bForceJobAnalysis == true)
            //////////            {
            //////////                // force it to report
            //////////                bExceedsThreholds = true;
            //////////            }

            //////////            if (bExceedsThreholds == true)
            //////////            {
            //////////                oSLBOMDocument.SetCellValue(iNumberOfRows, 1, oJobOperation.JobNum);
            //////////                oSLBOMDocument.SetCellValue(iNumberOfRows, 2, oJobOperation.PartNum);
            //////////                oSLBOMDocument.SetCellValue(iNumberOfRows, 3, oJobOperation.PartRevNum);
            //////////                oSLBOMDocument.SetCellValue(iNumberOfRows, 4, oJobOperation.AssemblySeq.ToString());
            //////////                oSLBOMDocument.SetCellValue(iNumberOfRows, 5, oJobOperation.OprSeq.ToString());
            //////////                oSLBOMDocument.SetCellValue(iNumberOfRows, 6, oJobOperation.OpCode);
            //////////                oSLBOMDocument.SetCellValue(iNumberOfRows, 7, oJobOperation.EstLaborCost);
            //////////                oSLBOMDocument.SetCellStyle(iNumberOfRows, 7, oCurrencyStyle);
            //////////                oSLBOMDocument.SetCellValue(iNumberOfRows, 8, oJobOperation.ActLaborCost);
            //////////                oSLBOMDocument.SetCellStyle(iNumberOfRows, 8, oCurrencyStyle);

            //////////                oSLBOMDocument.SetCellValue(iNumberOfRows, 9, dDelta);
            //////////                oSLBOMDocument.SetCellStyle(iNumberOfRows, 9, oCurrencyStyle);

            //////////                oSLBOMDocument.SetCellValue(iNumberOfRows, 10, dPercentageError);
            //////////                oSLBOMDocument.SetCellStyle(iNumberOfRows++, 10, oDecimalStyle);
            //////////                bDataInReport = true;
            //////////            }
            //////////        }
            //////////    }

            //////////    // est burden cost differs from actual
            //////////    List<JobOpsEstVsActualCosts> oJobOperationBurdenDiffers = m_oJobOpsEstVsActualCosts.Where(oItem => (oItem.ActBurdenCost > oItem.EstBurCost) || ((oItem.ActBurdenCost < oItem.EstBurCost) && (oItem.OpComplete == true))).ToList();
            //////////    if (oJobOperationBurdenDiffers.Count > 0)
            //////////    {
            //////////        // sort these by job number
            //////////        oJobOperationBurdenDiffers = oJobOperationBurdenDiffers.OrderBy(oItem => oItem.JobNum).ToList();
            //////////        iNumberOfRows = 1;
            //////////        iNumberOfColumns = 1;
            //////////        if (bFirstWorksheet == true)
            //////////        {
            //////////            oSLBOMDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "Op Burden");
            //////////            bFirstWorksheet = false;
            //////////        }
            //////////        else
            //////////        {
            //////////            oSLBOMDocument.AddWorksheet("Op Burden");
            //////////        }
            //////////        //set column header
            //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "JobNum");
            //////////        oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
            //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Parent Part Num");
            //////////        oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
            //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Parent Part Rev Num");
            //////////        oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
            //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Assembly Seq");
            //////////        oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
            //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Opr Seq");
            //////////        oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
            //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Opr Code");
            //////////        oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
            //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Est Burden");
            //////////        oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
            //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Act Burden");
            //////////        oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
            //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Variance $");
            //////////        oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
            //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Variance %");
            //////////        oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
            //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Operation Burden Cost Differs From Estimate -- POC ISSUE PLEASE REVIEW");
            //////////        oSLBOMDocument.SetCellStyle(iNumberOfRows++, iNumberOfColumns, oHighlightHeaderStyle);
            //////////        oSLBOMDocument.SetColumnWidth(iNumberOfColumns++, 90);

            //////////        foreach (JobOpsEstVsActualCosts oJobOperation in oJobOperationBurdenDiffers)
            //////////        {
            //////////            decimal dDelta = oJobOperation.ActBurdenCost - oJobOperation.EstBurCost;
            //////////            decimal dPercentageError = 0M;
            //////////            decimal dMarginToDate = 0M;
            //////////            if (oJobOperation.EstBurCost != 0)
            //////////            {
            //////////                dPercentageError = dDelta / oJobOperation.EstBurCost * 100.0M;
            //////////            }
            //////////            else
            //////////            {
            //////////                dPercentageError = 100.0M;
            //////////            }
            //////////            HSJob oTmpJob = m_oAllJobs.FirstOrDefault(oItem => string.Compare(oJobOperation.JobNum, oItem.JobNum, true) == 0);
            //////////            if (oTmpJob != null)
            //////////            {
            //////////                dMarginToDate = oTmpJob.MarginToDate * 100.0M;
            //////////            }

            //////////            bool bExceedsThreholds = false;
            //////////            if ((Math.Abs(dDelta) > m_dJobEstimateAbsoluteError) && (Math.Abs(dPercentageError) > m_dJobEstimatePercentError) && (dMarginToDate < m_dJobMarginThreshold))
            //////////            {
            //////////                // check to see if it exceeds both thresholds
            //////////                bExceedsThreholds = true;
            //////////            }
            //////////            if (m_bForceJobAnalysis == true)
            //////////            {
            //////////                // force it to report
            //////////                bExceedsThreholds = true;
            //////////            }

            //////////            if (bExceedsThreholds == true)
            //////////            {
            //////////                oSLBOMDocument.SetCellValue(iNumberOfRows, 1, oJobOperation.JobNum);
            //////////                oSLBOMDocument.SetCellValue(iNumberOfRows, 2, oJobOperation.PartNum);
            //////////                oSLBOMDocument.SetCellValue(iNumberOfRows, 3, oJobOperation.PartRevNum);
            //////////                oSLBOMDocument.SetCellValue(iNumberOfRows, 4, oJobOperation.AssemblySeq.ToString());
            //////////                oSLBOMDocument.SetCellValue(iNumberOfRows, 5, oJobOperation.OprSeq.ToString());
            //////////                oSLBOMDocument.SetCellValue(iNumberOfRows, 6, oJobOperation.OpCode);
            //////////                oSLBOMDocument.SetCellValue(iNumberOfRows, 7, oJobOperation.EstBurCost);
            //////////                oSLBOMDocument.SetCellStyle(iNumberOfRows, 7, oCurrencyStyle);
            //////////                oSLBOMDocument.SetCellValue(iNumberOfRows, 8, oJobOperation.ActBurdenCost);
            //////////                oSLBOMDocument.SetCellStyle(iNumberOfRows, 8, oCurrencyStyle);

            //////////                oSLBOMDocument.SetCellValue(iNumberOfRows, 9, dDelta);
            //////////                oSLBOMDocument.SetCellStyle(iNumberOfRows, 9, oCurrencyStyle);

            //////////                oSLBOMDocument.SetCellValue(iNumberOfRows, 10, dPercentageError);
            //////////                oSLBOMDocument.SetCellStyle(iNumberOfRows++, 10, oDecimalStyle);
            //////////                bDataInReport = true;
            //////////            }
            //////////        }
            //////////    }

            //////////    // est sub cost differs from actual
            //////////    List<JobOpsEstVsActualCosts> oJobOperationSubDiffers = m_oJobOpsEstVsActualCosts.Where(oItem => (oItem.ActSubCost > oItem.EstSubCost) || ((oItem.ActSubCost < oItem.EstSubCost) && (oItem.OpComplete == true))).ToList();
            //////////    if (oJobOperationSubDiffers.Count > 0)
            //////////    {
            //////////        // sort these by job number
            //////////        oJobOperationSubDiffers = oJobOperationSubDiffers.OrderBy(oItem => oItem.JobNum).ToList();
            //////////        iNumberOfRows = 1;
            //////////        iNumberOfColumns = 1;
            //////////        if (bFirstWorksheet == true)
            //////////        {
            //////////            oSLBOMDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "Subcontract");
            //////////            bFirstWorksheet = false;
            //////////        }
            //////////        else
            //////////        {
            //////////            oSLBOMDocument.AddWorksheet("Subcontract");
            //////////        }
            //////////        //set column header
            //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "JobNum");
            //////////        oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
            //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Parent Part Num");
            //////////        oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
            //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Parent Part Rev Num");
            //////////        oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
            //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Assembly Seq");
            //////////        oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
            //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Opr Seq");
            //////////        oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
            //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Opr Code");
            //////////        oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
            //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Est Sub");
            //////////        oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
            //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Act Sub");
            //////////        oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
            //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Variance $");
            //////////        oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
            //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Variance %");
            //////////        oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
            //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Subcontract Actual Cost Differs From Estimate -- POC ISSUE PLEASE REVIEW");
            //////////        oSLBOMDocument.SetCellStyle(iNumberOfRows++, iNumberOfColumns, oHighlightHeaderStyle);
            //////////        oSLBOMDocument.SetColumnWidth(iNumberOfColumns++, 90);

            //////////        foreach (JobOpsEstVsActualCosts oJobOperation in oJobOperationSubDiffers)
            //////////        {
            //////////            decimal dDelta = oJobOperation.ActSubCost - oJobOperation.EstSubCost;
            //////////            decimal dPercentageError = 0M;
            //////////            decimal dMarginToDate = 0M;
            //////////            if (oJobOperation.EstSubCost != 0)
            //////////            {
            //////////                dPercentageError = dDelta / oJobOperation.EstSubCost * 100.0M;
            //////////            }
            //////////            else
            //////////            {
            //////////                dPercentageError = 100.0M;
            //////////            }
            //////////            HSJob oTmpJob = m_oAllJobs.FirstOrDefault(oItem => string.Compare(oJobOperation.JobNum, oItem.JobNum, true) == 0);
            //////////            if (oTmpJob != null)
            //////////            {
            //////////                dMarginToDate = oTmpJob.MarginToDate * 100.0M;
            //////////            }

            //////////            bool bExceedsThreholds = false;
            //////////            if ((Math.Abs(dDelta) > m_dJobEstimateAbsoluteError) && (Math.Abs(dPercentageError) > m_dJobEstimatePercentError) && (dMarginToDate < m_dJobMarginThreshold))
            //////////            {
            //////////                // check to see if it exceeds both thresholds
            //////////                bExceedsThreholds = true;
            //////////            }
            //////////            if (m_bForceJobAnalysis == true)
            //////////            {
            //////////                // force it to report
            //////////                bExceedsThreholds = true;
            //////////            }

            //////////            if (bExceedsThreholds == true)
            //////////            {
            //////////                oSLBOMDocument.SetCellValue(iNumberOfRows, 1, oJobOperation.JobNum);
            //////////                oSLBOMDocument.SetCellValue(iNumberOfRows, 2, oJobOperation.PartNum);
            //////////                oSLBOMDocument.SetCellValue(iNumberOfRows, 3, oJobOperation.PartRevNum);
            //////////                oSLBOMDocument.SetCellValue(iNumberOfRows, 4, oJobOperation.AssemblySeq.ToString());
            //////////                oSLBOMDocument.SetCellValue(iNumberOfRows, 5, oJobOperation.OprSeq.ToString());
            //////////                oSLBOMDocument.SetCellValue(iNumberOfRows, 6, oJobOperation.OpCode);
            //////////                oSLBOMDocument.SetCellValue(iNumberOfRows, 7, oJobOperation.EstSubCost);
            //////////                oSLBOMDocument.SetCellStyle(iNumberOfRows, 7, oCurrencyStyle);
            //////////                oSLBOMDocument.SetCellValue(iNumberOfRows, 8, oJobOperation.ActSubCost);
            //////////                oSLBOMDocument.SetCellStyle(iNumberOfRows, 8, oCurrencyStyle);

            //////////                oSLBOMDocument.SetCellValue(iNumberOfRows, 9, dDelta);
            //////////                oSLBOMDocument.SetCellStyle(iNumberOfRows, 9, oCurrencyStyle);

            //////////                oSLBOMDocument.SetCellValue(iNumberOfRows, 10, dPercentageError);
            //////////                oSLBOMDocument.SetCellStyle(iNumberOfRows++, 10, oDecimalStyle);
            //////////                bDataInReport = true;
            //////////            }
            //////////        }
            //////////    }

            //////////    // total op cost differs from actual
            //////////    List<JobOpsEstVsActualCosts> oJobOperationTotalDiffers = m_oJobOpsEstVsActualCosts.Where(oItem => (oItem.ActCost > oItem.EstCost) || ((oItem.ActCost < oItem.EstCost) && (oItem.OpComplete == true))).ToList();
            //////////    if (oJobOperationTotalDiffers.Count > 0)
            //////////    {
            //////////        // sort these by job number
            //////////        oJobOperationTotalDiffers = oJobOperationTotalDiffers.OrderBy(oItem => oItem.JobNum).ToList();
            //////////        iNumberOfRows = 1;
            //////////        iNumberOfColumns = 1;
            //////////        if (bFirstWorksheet == true)
            //////////        {
            //////////            oSLBOMDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "Op Total");
            //////////            bFirstWorksheet = false;
            //////////        }
            //////////        else
            //////////        {
            //////////            oSLBOMDocument.AddWorksheet("Op Total");
            //////////        }
            //////////        //set column header
            //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "JobNum");
            //////////        oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
            //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Parent Part Num");
            //////////        oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
            //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Parent Part Rev Num");
            //////////        oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
            //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Assembly Seq");
            //////////        oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
            //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Opr Seq");
            //////////        oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
            //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Opr Code");
            //////////        oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
            //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Est Total");
            //////////        oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
            //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Act Total");
            //////////        oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
            //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Variance $");
            //////////        oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
            //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Variance %");
            //////////        oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
            //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Total Operation Cost Differs From Estimate -- POC ISSUE PLEASE REVIEW");
            //////////        oSLBOMDocument.SetCellStyle(iNumberOfRows++, iNumberOfColumns, oHighlightHeaderStyle);
            //////////        oSLBOMDocument.SetColumnWidth(iNumberOfColumns++, 90);

            //////////        foreach (JobOpsEstVsActualCosts oJobOperation in oJobOperationTotalDiffers)
            //////////        {
            //////////            decimal dDelta = oJobOperation.ActCost - oJobOperation.EstCost;
            //////////            decimal dPercentageError = 0M;
            //////////            decimal dMarginToDate = 0M;
            //////////            if (oJobOperation.EstCost != 0)
            //////////            {
            //////////                dPercentageError = dDelta / oJobOperation.EstCost * 100.0M;
            //////////            }
            //////////            else
            //////////            {
            //////////                dPercentageError = 100.0M;
            //////////            }
            //////////            HSJob oTmpJob = m_oAllJobs.FirstOrDefault(oItem => string.Compare(oJobOperation.JobNum, oItem.JobNum, true) == 0);
            //////////            if (oTmpJob != null)
            //////////            {
            //////////                dMarginToDate = oTmpJob.MarginToDate * 100.0M;
            //////////            }

            //////////            bool bExceedsThreholds = false;
            //////////            if ((Math.Abs(dDelta) > m_dJobEstimateAbsoluteError) && (Math.Abs(dPercentageError) > m_dJobEstimatePercentError) && (dMarginToDate < m_dJobMarginThreshold))
            //////////            {
            //////////                // check to see if it exceeds both thresholds
            //////////                bExceedsThreholds = true;
            //////////            }
            //////////            if (m_bForceJobAnalysis == true)
            //////////            {
            //////////                // force it to report
            //////////                bExceedsThreholds = true;
            //////////            }

            //////////            if (bExceedsThreholds == true)
            //////////            {
            //////////                oSLBOMDocument.SetCellValue(iNumberOfRows, 1, oJobOperation.JobNum);
            //////////                oSLBOMDocument.SetCellValue(iNumberOfRows, 2, oJobOperation.PartNum);
            //////////                oSLBOMDocument.SetCellValue(iNumberOfRows, 3, oJobOperation.PartRevNum);
            //////////                oSLBOMDocument.SetCellValue(iNumberOfRows, 4, oJobOperation.AssemblySeq.ToString());
            //////////                oSLBOMDocument.SetCellValue(iNumberOfRows, 5, oJobOperation.OprSeq.ToString());
            //////////                oSLBOMDocument.SetCellValue(iNumberOfRows, 6, oJobOperation.OpCode);
            //////////                oSLBOMDocument.SetCellValue(iNumberOfRows, 7, oJobOperation.EstCost);
            //////////                oSLBOMDocument.SetCellStyle(iNumberOfRows, 7, oCurrencyStyle);
            //////////                oSLBOMDocument.SetCellValue(iNumberOfRows, 8, oJobOperation.ActCost);
            //////////                oSLBOMDocument.SetCellStyle(iNumberOfRows, 8, oCurrencyStyle);

            //////////                oSLBOMDocument.SetCellValue(iNumberOfRows, 9, dDelta);
            //////////                oSLBOMDocument.SetCellStyle(iNumberOfRows, 9, oCurrencyStyle);

            //////////                oSLBOMDocument.SetCellValue(iNumberOfRows, 10, dPercentageError);
            //////////                oSLBOMDocument.SetCellStyle(iNumberOfRows++, 10, oDecimalStyle);
            //////////                bDataInReport = true;
            //////////            }
            //////////        }
            //////////    }
            //////////}
            #endregion

            #endregion

            if (bDataInReport == true)
            {
                oSLBOMDocument.SaveAs(sDestinationFileName);
                // Check to see if we created a file and if so email it
                if (File.Exists(sDestinationFileName) == true)
                {
                    List<string> oAttachments = new List<string>();
                    oAttachments.Add(sDestinationFileName);
                    if (oRequestingUser != null)
                    {
                        oToAddresses.Add(oRequestingUser.Email);
                    }
                    HSEmailHelper.SendEmail(oToAddresses, sCompany + " Job Validation Report", sCompany + " Job Validation Report for " + sDate, oAttachments);
                }
            }
        }

        public void PerformJobValidationByJob(string sCompany, string sTmpFileDirectory, HSUser oRequestingUser)
        {
            // we will evaluate this one job at a time
            foreach (HSJob oJob in m_oAllJobs)
            {
                #region Setup

                // get the file name
                DateTime dtToday = DateTime.Now;
                string sDate = dtToday.Month.ToString() + "-" + dtToday.Day.ToString() + "-" + dtToday.Year.ToString();
                string sDestinationFileName = sTmpFileDirectory + "\\" + sCompany + "-Job-" + oJob.JobNum + "-ValidationReport-" + sDate + ".xlsx";
                int iNumberOfRows = 1;
                int iNumberOfColumns = 1;

                HSUser oServiceAccount = HSUser.GetUserById(HSUser.SPINCRAFT_SERVICE_ACCOUNT_ID);
                List<string> oToAddresses = new List<string>();
                if (oRequestingUser == null)
                {
                    // get users in the engineering, production, and quoting groups
                    HSUser.GetEmailsForUsersInGroup(oToAddresses, HSUser.REPORT_ON_JOB_ESTIMATES);
                }
                else
                {
                    // this report was requested by someone directly from Epicor
                    oToAddresses.Add(oRequestingUser.Email);
                }
                oToAddresses.Add(HSUser.SPINCRAFT_ROOT_USER);

                SLFill oSLFill = new SLFill();
                oSLFill.SetPatternBackgroundColor(SLThemeColorIndexValues.Accent1Color);
                oSLFill.SetPattern(PatternValues.Solid, SLThemeColorIndexValues.Light2Color, SLThemeColorIndexValues.Light2Color);

                SLAlignment oCenterAlignment = new SLAlignment();
                oCenterAlignment.Horizontal = HorizontalAlignmentValues.Center;

                SLAlignment oRightAlignment = new SLAlignment();
                oRightAlignment.Horizontal = HorizontalAlignmentValues.Right;

                SLAlignment oLeftAlignment = new SLAlignment();
                oLeftAlignment.Horizontal = HorizontalAlignmentValues.Left;

                SLStyle oBoldStyle = new SLStyle();
                oBoldStyle.SetFontBold(true);
                oBoldStyle.SetTopBorder(BorderStyleValues.Thin, SLThemeColorIndexValues.Dark2Color);
                oBoldStyle.SetBottomBorder(BorderStyleValues.Thin, SLThemeColorIndexValues.Dark2Color);
                oBoldStyle.SetLeftBorder(BorderStyleValues.Thin, SLThemeColorIndexValues.Dark2Color);
                oBoldStyle.SetRightBorder(BorderStyleValues.Thin, SLThemeColorIndexValues.Dark2Color);
                oBoldStyle.SetFontColor(SLThemeColorIndexValues.Dark2Color);

                SLStyle oBoldStyle2 = new SLStyle();
                oBoldStyle2.SetFontBold(true);
                oBoldStyle2.SetTopBorder(BorderStyleValues.Thick, SLThemeColorIndexValues.Dark2Color);
                oBoldStyle2.SetBottomBorder(BorderStyleValues.Thick, SLThemeColorIndexValues.Dark2Color);
                oBoldStyle2.SetLeftBorder(BorderStyleValues.Thick, SLThemeColorIndexValues.Dark2Color);
                oBoldStyle.SetRightBorder(BorderStyleValues.Thick, SLThemeColorIndexValues.Dark2Color);
                oBoldStyle2.SetFontColor(System.Drawing.Color.IndianRed);

                SLStyle oCurrencyStyle = new SLStyle();
                oCurrencyStyle.ApplyNamedCellStyle(SLNamedCellStyleValues.Currency);
                oCurrencyStyle.Alignment = oLeftAlignment;
                oCurrencyStyle.FormatCode = "$#,##0.00";

                SLStyle oDecimalStyle = new SLStyle();
                oDecimalStyle.Alignment = oLeftAlignment;
                oDecimalStyle.FormatCode = "###.00";

                System.Drawing.Color oHighlightYellow = System.Drawing.Color.FromArgb(255, 255, 0);

                SLStyle oHighlightHeaderStyle = new SLStyle();
                oHighlightHeaderStyle.SetFontBold(true);
                oHighlightHeaderStyle.SetFont(FontSchemeValues.Major, 12);
                oHighlightHeaderStyle.SetHorizontalAlignment(HorizontalAlignmentValues.Center);
                oHighlightHeaderStyle.SetPatternFill(PatternValues.Solid, oHighlightYellow, oHighlightYellow);
                oHighlightHeaderStyle.SetFontColor(SLThemeColorIndexValues.Dark2Color);

                // colors for plotting
                System.Drawing.Color oLightRed = System.Drawing.Color.FromArgb(240, 160, 140);
                System.Drawing.Color oLightGreen = System.Drawing.Color.FromArgb(160, 225, 85);
                System.Drawing.Color oDRed = System.Drawing.Color.FromArgb(240, 80, 90);
                System.Drawing.Color oDGreen = System.Drawing.Color.FromArgb(160, 180, 125);
                System.Drawing.Color oBrightPurple = System.Drawing.Color.FromArgb(150, 47, 214);
                System.Drawing.Color oBrightBlue = System.Drawing.Color.FromArgb(10, 18, 248);
                System.Drawing.Color oBrightOrange = System.Drawing.Color.FromArgb(253, 150, 5);
                System.Drawing.Color oBrightYellow = System.Drawing.Color.FromArgb(249, 253, 5);
                System.Drawing.Color oSkyBlue = System.Drawing.Color.FromArgb(5, 232, 253);
                System.Drawing.Color oBrightPink = System.Drawing.Color.FromArgb(253, 5, 183);
                System.Drawing.Color oBrightRed = System.Drawing.Color.FromArgb(253, 5, 5);

                SLStyle oOperationCompleteStyle = new SLStyle();
                oOperationCompleteStyle.Fill.SetPattern(PatternValues.Solid, oLightGreen, oDGreen);

                SLStyle oCurrencyStyleUnderEstimate = new SLStyle();
                oCurrencyStyleUnderEstimate.ApplyNamedCellStyle(SLNamedCellStyleValues.Currency);
                oCurrencyStyleUnderEstimate.Alignment = oLeftAlignment;
                oCurrencyStyleUnderEstimate.FormatCode = "$#,##0.00";
                oCurrencyStyleUnderEstimate.SetFontColor(oDGreen);

                SLStyle oCurrencyStyleOverEstimate = new SLStyle();
                oCurrencyStyleOverEstimate.ApplyNamedCellStyle(SLNamedCellStyleValues.Currency);
                oCurrencyStyleOverEstimate.Alignment = oLeftAlignment;
                oCurrencyStyleOverEstimate.FormatCode = "$#,##0.00";
                oCurrencyStyleOverEstimate.SetFontColor(oDRed);

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
                SLDocument oSLBOMDocument = new SLDocument();

                bool bDataInReport = false;
                bool bFirstWorksheet = true;

                int iStandardColumnWidth = 20;
                #endregion

                #region Job Issues
                // jobs with zero qty
                if (oJob.ProductionQty == 0)
                {
                    iNumberOfRows = 1;
                    iNumberOfColumns = 1;
                    if (bFirstWorksheet == true)
                    {
                        oSLBOMDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "Jobs Without Production Qty");
                        bFirstWorksheet = false;
                    }
                    else
                    {
                        oSLBOMDocument.AddWorksheet("Jobs Without Production Qty");
                    }
                    //set column header
                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "JobNum");
                    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Part Num");
                    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Part Rev Num");
                    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Jobs Should Have A Production Qty -- SOP REQUIRES JOBS TO HAVE PRODUCTION QTY");
                    oSLBOMDocument.SetCellStyle(iNumberOfRows++, iNumberOfColumns, oHighlightHeaderStyle);
                    oSLBOMDocument.SetColumnWidth(iNumberOfColumns++, 90);

                    oSLBOMDocument.SetCellValue(iNumberOfRows, 1, oJob.JobNum);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 2, oJob.PartNum);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 3, oJob.PartRevNum);

                    iNumberOfRows++;
                    bDataInReport = true;
                    
                }

                // jobs missing due date
                if (oJob.DueDate == DateTime.MinValue)
                {
                    iNumberOfRows = 1;
                    iNumberOfColumns = 1;
                    if (bFirstWorksheet == true)
                    {
                        oSLBOMDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "Jobs Without Due Date");
                        bFirstWorksheet = false;
                    }
                    else
                    {
                        oSLBOMDocument.AddWorksheet("Jobs Without Due Date");
                    }
                    //set column header
                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "JobNum");
                    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Part Num");
                    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Part Rev Num");
                    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Jobs Should Have Due Date Set For Planning and Purchasing Purposes -- SOP REQUIRES DUE DATE TO BE SET");
                    oSLBOMDocument.SetCellStyle(iNumberOfRows++, iNumberOfColumns, oHighlightHeaderStyle);
                    oSLBOMDocument.SetColumnWidth(iNumberOfColumns++, 90);

                    oSLBOMDocument.SetCellValue(iNumberOfRows, 1, oJob.JobNum);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 2, oJob.PartNum);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 3, oJob.PartRevNum);

                    iNumberOfRows++;
                    bDataInReport = true;
                }

                // jobs missing required by date
                if (oJob.RequiredDate == DateTime.MinValue)
                {
                    iNumberOfRows = 1;
                    iNumberOfColumns = 1;
                    if (bFirstWorksheet == true)
                    {
                        oSLBOMDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "Jobs Without Required Date");
                        bFirstWorksheet = false;
                    }
                    else
                    {
                        oSLBOMDocument.AddWorksheet("Jobs Without Required Date");
                    }
                    //set column header
                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "JobNum");
                    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Part Num");
                    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Part Rev Num");
                    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Jobs Should Have Required By Date Set For Planning and Purchasing Purposes -- SOP STATES REQUIRED DATE SHOULD BE SET");
                    oSLBOMDocument.SetCellStyle(iNumberOfRows++, iNumberOfColumns, oHighlightHeaderStyle);
                    oSLBOMDocument.SetColumnWidth(iNumberOfColumns++, 90);

                    oSLBOMDocument.SetCellValue(iNumberOfRows, 1, oJob.JobNum);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 2, oJob.PartNum);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 3, oJob.PartRevNum);

                    iNumberOfRows++;
                    bDataInReport = true;
                }

                // jobs missing start date
                if (oJob.StartDate == DateTime.MinValue)
                {
                    iNumberOfRows = 1;
                    iNumberOfColumns = 1;
                    if (bFirstWorksheet == true)
                    {
                        oSLBOMDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "Jobs Without Start Date");
                        bFirstWorksheet = false;
                    }
                    else
                    {
                        oSLBOMDocument.AddWorksheet("Jobs Without Start Date");
                    }
                    //set column header
                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "JobNum");
                    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Part Num");
                    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Part Rev Num");
                    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Jobs Should Have Start Date Set For Planning and Purchasing Purposes -- SOP REQUIRES START DATE TO BE SET");
                    oSLBOMDocument.SetCellStyle(iNumberOfRows++, iNumberOfColumns, oHighlightHeaderStyle);
                    oSLBOMDocument.SetColumnWidth(iNumberOfColumns++, 90);

                    oSLBOMDocument.SetCellValue(iNumberOfRows, 1, oJob.JobNum);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 2, oJob.PartNum);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 3, oJob.PartRevNum);

                    iNumberOfRows++;
                    bDataInReport = true;
                }

                // jobs that have parts on the fly
                if (oJob.PartMaster == null)
                {
                    iNumberOfRows = 1;
                    iNumberOfColumns = 1;
                    if (bFirstWorksheet == true)
                    {
                        oSLBOMDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "Jobs Without Part Master");
                        bFirstWorksheet = false;
                    }
                    else
                    {
                        oSLBOMDocument.AddWorksheet("Jobs Without Part Master");
                    }
                    //set column header
                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "JobNum");
                    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Part Num");
                    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Part Rev Num");
                    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "The Part Being Built Is Not Formally In The Part Master");
                    oSLBOMDocument.SetCellStyle(iNumberOfRows++, iNumberOfColumns, oHighlightHeaderStyle);
                    oSLBOMDocument.SetColumnWidth(iNumberOfColumns++, 90);

                    oSLBOMDocument.SetCellValue(iNumberOfRows, 1, oJob.JobNum);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 2, oJob.PartNum);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 3, oJob.PartRevNum);

                    iNumberOfRows++;
                    bDataInReport = true;
                }

                // jobs that have make direct finished goods but are not tied to SO -- issue trying to compute margin for POC
                if ((oJob.OrderNum == 0) && (oJob.PartMaster != null) && (oJob.PartMaster.PartNonStock == true))
                {
                    iNumberOfRows = 1;
                    iNumberOfColumns = 1;
                    if (bFirstWorksheet == true)
                    {
                        oSLBOMDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "Make Direct No SO");
                        bFirstWorksheet = false;
                    }
                    else
                    {
                        oSLBOMDocument.AddWorksheet("Make Direct No SO");
                    }
                    //set column header
                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "JobNum");
                    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Part Num");
                    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Part Rev Num");
                    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "The Part Is Set To Make Direct But The Job Is Not Tied To A Sales Order - POC ISSUE");
                    oSLBOMDocument.SetCellStyle(iNumberOfRows++, iNumberOfColumns, oHighlightHeaderStyle);
                    oSLBOMDocument.SetColumnWidth(iNumberOfColumns++, 90);

                    oSLBOMDocument.SetCellValue(iNumberOfRows, 1, oJob.JobNum);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 2, oJob.PartNum);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 3, oJob.PartRevNum);

                    iNumberOfRows++;
                    bDataInReport = true;
                }

                // jobs that are stocking but tied to sales order
                if ((oJob.OrderNum != 0) && (oJob.PartMaster != null) && (oJob.PartMaster.PartNonStock == false))
                {
                    iNumberOfRows = 1;
                    iNumberOfColumns = 1;
                    if (bFirstWorksheet == true)
                    {
                        oSLBOMDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "Stock MFG Tied To SO");
                        bFirstWorksheet = false;
                    }
                    else
                    {
                        oSLBOMDocument.AddWorksheet("Stock MFG Tied To SO");
                    }
                    //set column header
                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "JobNum");
                    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Part Num");
                    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Part Rev Num");
                    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "The Part Is Set To Stocking But The Job Is Directly Tied To A Sales Order - POC ISSUE");
                    oSLBOMDocument.SetCellStyle(iNumberOfRows++, iNumberOfColumns, oHighlightHeaderStyle);
                    oSLBOMDocument.SetColumnWidth(iNumberOfColumns++, 90);

                    oSLBOMDocument.SetCellValue(iNumberOfRows, 1, oJob.JobNum);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 2, oJob.PartNum);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 3, oJob.PartRevNum);

                    iNumberOfRows++;
                    bDataInReport = true;
                }

                // jobs that are below the margin threshold
                if ((oJob.HasTransactions == true) && (oJob.MarginToDate * 100.0M < m_dJobMarginThreshold))
                {
                    iNumberOfRows = 1;
                    iNumberOfColumns = 1;
                    if (bFirstWorksheet == true)
                    {
                        oSLBOMDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "Jobs Below Margin");
                        bFirstWorksheet = false;
                    }
                    else
                    {
                        oSLBOMDocument.AddWorksheet("Jobs Below Margin");
                    }
                    //set column header
                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "JobNum");
                    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Part Num");
                    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Part Rev Num");
                    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Current Margin");
                    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "The Job Is Below The Margin Threshold -- REVIEW");
                    oSLBOMDocument.SetCellStyle(iNumberOfRows++, iNumberOfColumns, oHighlightHeaderStyle);
                    oSLBOMDocument.SetColumnWidth(iNumberOfColumns++, 90);

                    oSLBOMDocument.SetCellValue(iNumberOfRows, 1, oJob.JobNum);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 2, oJob.PartNum);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 3, oJob.PartRevNum);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 4, oJob.MarginToDate);
                    oSLBOMDocument.SetCellStyle(iNumberOfRows, 4, oDecimalStyle);

                    iNumberOfRows++;
                    bDataInReport = true;
                }
                #endregion

                #region Material Attributes

                List<JobMaterial> oMaterialsWithZeroQty = oJob.AllMaterials.Where(oItem => (oItem.RequiredQty == 0) && (oItem.PartMaster != null) && ((oItem.QtyBearing == true) || (m_oPartClassesToIgnore.Contains(oItem.PartMaster.ClassID) == false))).ToList();
                if (oMaterialsWithZeroQty.Count > 0)
                {
                    // sort these by job number
                    oMaterialsWithZeroQty = oMaterialsWithZeroQty.OrderBy(oItem => oItem.JobNum).ToList();
                    iNumberOfRows = 1;
                    iNumberOfColumns = 1;
                    if (bFirstWorksheet == true)
                    {
                        oSLBOMDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "Mtl Qty Zero");
                        bFirstWorksheet = false;
                    }
                    else
                    {
                        oSLBOMDocument.AddWorksheet("Mtl Qty Zero");
                    }
                    //set column header
                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "JobNum");
                    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Parent Part Num");
                    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Parent Part Rev Num");
                    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Assembly Seq");
                    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Mtl Seq");
                    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Mtl Part Num");
                    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Mtl Part Rev Num");
                    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Class");
                    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "The Job Material Qty Is Zero -- POC ISSUE PLEASE SET QTY");
                    oSLBOMDocument.SetCellStyle(iNumberOfRows++, iNumberOfColumns, oHighlightHeaderStyle);
                    oSLBOMDocument.SetColumnWidth(iNumberOfColumns++, 90);

                    foreach (JobMaterial oJobMateiral in oMaterialsWithZeroQty)
                    {
                        oSLBOMDocument.SetCellValue(iNumberOfRows, 1, oJobMateiral.JobNum);
                        oSLBOMDocument.SetCellValue(iNumberOfRows, 2, oJobMateiral.ParentPartNum);
                        oSLBOMDocument.SetCellValue(iNumberOfRows, 3, oJobMateiral.ParentRevNum);
                        oSLBOMDocument.SetCellValue(iNumberOfRows, 4, oJobMateiral.AssemblySeq.ToString());
                        oSLBOMDocument.SetCellValue(iNumberOfRows, 5, oJobMateiral.MtlSeq.ToString());
                        oSLBOMDocument.SetCellValue(iNumberOfRows, 6, oJobMateiral.MaterialPartNum);
                        oSLBOMDocument.SetCellValue(iNumberOfRows, 7, oJobMateiral.MaterialPartRevNum);
                        if (oJobMateiral.PartMaster != null)
                        {
                            oSLBOMDocument.SetCellValue(iNumberOfRows, 8, oJobMateiral.PartMaster.ClassID);
                        }

                        iNumberOfRows++;
                        bDataInReport = true;
                    }
                }

                #region IGNORING THESE ISSUES
                // materials not tied to operatiopn
                //////////List<JobMaterial> oMaterialsNotAssociatedToOp = m_oJobMaterials.Where(oItem => (oItem.RelatedOperation == 0) && (oItem.PartMaster != null) && (m_oPartClassesToIgnore.Contains(oItem.PartMaster.ClassID) == false)).ToList();
                //////////if (oMaterialsNotAssociatedToOp.Count > 0)
                //////////{
                //////////    // sort these by job number
                //////////    oMaterialsNotAssociatedToOp = oMaterialsNotAssociatedToOp.OrderBy(oItem => oItem.JobNum).ToList();
                //////////    iNumberOfRows = 1;
                //////////    iNumberOfColumns = 1;
                //////////    if (bFirstWorksheet == true)
                //////////    {
                //////////        oSLBOMDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "Mtl No Op");
                //////////        bFirstWorksheet = false;
                //////////    }
                //////////    else
                //////////    {
                //////////        oSLBOMDocument.AddWorksheet("Mtl No Op");
                //////////    }
                //////////    //set column header
                //////////    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "JobNum");
                //////////    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                //////////    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Parent Part Num");
                //////////    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                //////////    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Parent Part Rev Num");
                //////////    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                //////////    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Assembly Seq");
                //////////    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                //////////    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Mtl Seq");
                //////////    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                //////////    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Mtl Part Num");
                //////////    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                //////////    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Mtl Part Rev Num");
                //////////    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                //////////    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Class");
                //////////    oSLBOMDocument.SetColumnWidth(iNumberOfRows++, iNumberOfColumns++, iStandardColumnWidth);

                //////////    foreach (JobMaterial oJobMateiral in oMaterialsNotAssociatedToOp)
                //////////    {
                //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, 1, oJobMateiral.JobNum);
                //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, 2, oJobMateiral.ParentPartNum);
                //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, 3, oJobMateiral.ParentRevNum);
                //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, 4, oJobMateiral.AssemblySeq.ToString());
                //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, 5, oJobMateiral.MtlSeq.ToString());
                //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, 6, oJobMateiral.MaterialPartNum);
                //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, 7, oJobMateiral.MaterialPartRevNum);
                //////////        if (oJobMateiral.PartMaster != null)
                //////////        {
                //////////            oSLBOMDocument.SetCellValue(iNumberOfRows, 8, oJobMateiral.PartMaster.ClassID);
                //////////        }

                //////////        iNumberOfRows++;
                //////////        bDataInReport = true;
                //////////    }
                //////////}

                // materials no lead time
                //////////List<JobMaterial> oMaterialsNoLeadTime = m_oJobMaterials.Where(oItem => (oItem.LeadTime == 0) && (oItem.PartMaster != null) && (m_oPartClassesToIgnore.Contains(oItem.PartMaster.ClassID) == false)).ToList();
                //////////if (oMaterialsNoLeadTime.Count > 0)
                //////////{
                //////////    // sort these by job number
                //////////    oMaterialsNoLeadTime = oMaterialsNoLeadTime.OrderBy(oItem => oItem.JobNum).ToList();
                //////////    iNumberOfRows = 1;
                //////////    iNumberOfColumns = 1;
                //////////    if (bFirstWorksheet == true)
                //////////    {
                //////////        oSLBOMDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "Mtl No Lead Time");
                //////////        bFirstWorksheet = false;
                //////////    }
                //////////    else
                //////////    {
                //////////        oSLBOMDocument.AddWorksheet("Mtl No Lead Time");
                //////////    }
                //////////    //set column header
                //////////    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "JobNum");
                //////////    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                //////////    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Parent Part Num");
                //////////    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                //////////    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Parent Part Rev Num");
                //////////    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                //////////    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Assembly Seq");
                //////////    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                //////////    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Mtl Seq");
                //////////    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                //////////    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Mtl Part Num");
                //////////    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                //////////    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Mtl Part Rev Num");
                //////////    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                //////////    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Class");
                //////////    oSLBOMDocument.SetColumnWidth(iNumberOfRows++, iNumberOfColumns++, iStandardColumnWidth);

                //////////    foreach (JobMaterial oJobMateiral in oMaterialsNoLeadTime)
                //////////    {
                //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, 1, oJobMateiral.JobNum);
                //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, 2, oJobMateiral.ParentPartNum);
                //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, 3, oJobMateiral.ParentRevNum);
                //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, 4, oJobMateiral.AssemblySeq.ToString());
                //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, 5, oJobMateiral.MtlSeq.ToString());
                //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, 6, oJobMateiral.MaterialPartNum);
                //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, 7, oJobMateiral.MaterialPartRevNum);
                //////////        if (oJobMateiral.PartMaster != null)
                //////////        {
                //////////            oSLBOMDocument.SetCellValue(iNumberOfRows, 8, oJobMateiral.PartMaster.ClassID);
                //////////        }

                //////////        iNumberOfRows++;
                //////////        bDataInReport = true;
                //////////    }
                //////////}

                // materials no required date
                //////////List<JobMaterial> oMaterialsNoRequiredDate = m_oJobMaterials.Where(oItem => (oItem.RequiredDate == null) || (oItem.RequiredDate == DateTime.MinValue) && (oItem.PartMaster != null) && (m_oPartClassesToIgnore.Contains(oItem.PartMaster.ClassID) == false)).ToList();
                //////////if (oMaterialsNoRequiredDate.Count > 0)
                //////////{
                //////////    // sort these by job number
                //////////    oMaterialsNoRequiredDate = oMaterialsNoRequiredDate.OrderBy(oItem => oItem.JobNum).ToList();
                //////////    iNumberOfRows = 1;
                //////////    iNumberOfColumns = 1;
                //////////    if (bFirstWorksheet == true)
                //////////    {
                //////////        oSLBOMDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "Mtl No Required Date");
                //////////        bFirstWorksheet = false;
                //////////    }
                //////////    else
                //////////    {
                //////////        oSLBOMDocument.AddWorksheet("Mtl No Required Date");
                //////////    }
                //////////    //set column header
                //////////    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "JobNum");
                //////////    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                //////////    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Parent Part Num");
                //////////    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                //////////    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Parent Part Rev Num");
                //////////    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                //////////    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Assembly Seq");
                //////////    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                //////////    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Mtl Seq");
                //////////    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                //////////    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Mtl Part Num");
                //////////    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                //////////    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Mtl Part Rev Num");
                //////////    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                //////////    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Class");
                //////////    oSLBOMDocument.SetColumnWidth(iNumberOfRows++, iNumberOfColumns++, iStandardColumnWidth);

                //////////    foreach (JobMaterial oJobMateiral in oMaterialsNoRequiredDate)
                //////////    {
                //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, 1, oJobMateiral.JobNum);
                //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, 2, oJobMateiral.ParentPartNum);
                //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, 3, oJobMateiral.ParentRevNum);
                //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, 4, oJobMateiral.AssemblySeq.ToString());
                //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, 5, oJobMateiral.MtlSeq.ToString());
                //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, 6, oJobMateiral.MaterialPartNum);
                //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, 7, oJobMateiral.MaterialPartRevNum);
                //////////        if (oJobMateiral.PartMaster != null)
                //////////        {
                //////////            oSLBOMDocument.SetCellValue(iNumberOfRows, 8, oJobMateiral.PartMaster.ClassID);
                //////////        }

                //////////        iNumberOfRows++;
                //////////        bDataInReport = true;
                //////////    }
                //////////}

                // materials set to backflush
                //////////List<JobMaterial> oMaterialsSetToBackflush = m_oJobMaterials.Where(oItem => oItem.Backflush == true).ToList();
                //////////if (oMaterialsSetToBackflush.Count > 0)
                //////////{
                //////////    // sort these by job number
                //////////    oMaterialsSetToBackflush = oMaterialsSetToBackflush.OrderBy(oItem => oItem.JobNum).ToList();
                //////////    iNumberOfRows = 1;
                //////////    iNumberOfColumns = 1;
                //////////    if (bFirstWorksheet == true)
                //////////    {
                //////////        oSLBOMDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "Mtl Set To Backflush");
                //////////        bFirstWorksheet = false;
                //////////    }
                //////////    else
                //////////    {
                //////////        oSLBOMDocument.AddWorksheet("Mtl Set To Backflush");
                //////////    }
                //////////    //set column header
                //////////    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "JobNum");
                //////////    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                //////////    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Parent Part Num");
                //////////    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                //////////    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Parent Part Rev Num");
                //////////    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                //////////    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Assembly Seq");
                //////////    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                //////////    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Mtl Seq");
                //////////    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                //////////    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Mtl Part Num");
                //////////    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                //////////    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Mtl Part Rev Num");
                //////////    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                //////////    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Class");
                //////////    oSLBOMDocument.SetColumnWidth(iNumberOfRows++, iNumberOfColumns++, iStandardColumnWidth);

                //////////    foreach (JobMaterial oJobMateiral in oMaterialsSetToBackflush)
                //////////    {
                //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, 1, oJobMateiral.JobNum);
                //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, 2, oJobMateiral.ParentPartNum);
                //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, 3, oJobMateiral.ParentRevNum);
                //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, 4, oJobMateiral.AssemblySeq.ToString());
                //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, 5, oJobMateiral.MtlSeq.ToString());
                //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, 6, oJobMateiral.MaterialPartNum);
                //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, 7, oJobMateiral.MaterialPartRevNum);
                //////////        if (oJobMateiral.PartMaster != null)
                //////////        {
                //////////            oSLBOMDocument.SetCellValue(iNumberOfRows, 8, oJobMateiral.PartMaster.ClassID);
                //////////        }

                //////////        iNumberOfRows++;
                //////////        bDataInReport = true;
                //////////    }
                //////////}

                // materials on hold

                // materials inactive

                // materials bad rev

                // materials MRP turned off

                // group code not set

                // class code not set
                #endregion

                // material issues
                List<JobMaterial> oExcessMaterialIssued = oJob.AllMaterials.Where(oItem => (oItem.IssuedQty > oItem.RequiredQty) || ((oItem.IssuedQty < oItem.RequiredQty) && (oItem.QtyBearing == true) && (oItem.OpComplete == true) && ((oItem.PartMaster == null) || (m_oPartClassesToIgnore.Contains(oItem.PartMaster.ClassID) == false)))).ToList();
                if (oExcessMaterialIssued.Count > 0)
                {
                    // sort these by job number
                    oExcessMaterialIssued = oExcessMaterialIssued.OrderBy(oItem => oItem.JobNum).ToList();
                    iNumberOfRows = 1;
                    iNumberOfColumns = 1;
                    if (bFirstWorksheet == true)
                    {
                        oSLBOMDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "Mtl Issued");
                        bFirstWorksheet = false;
                    }
                    else
                    {
                        oSLBOMDocument.AddWorksheet("Mtl Issued");
                    }
                    //set column header
                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "JobNum");
                    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Parent Part Num");
                    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Parent Part Rev Num");
                    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Assembly Seq");
                    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Mtl Seq");
                    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Mtl Part Num");
                    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Mtl Part Rev Num");
                    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Class");
                    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Mtl Req");
                    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Mtl Issued");
                    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Variance Qty");
                    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Variance %");
                    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "The Job Material Issued Differs From Estimate -- POC ISSUE PLEASE REVIEW QTY");
                    oSLBOMDocument.SetCellStyle(iNumberOfRows++, iNumberOfColumns, oHighlightHeaderStyle);
                    oSLBOMDocument.SetColumnWidth(iNumberOfColumns++, 90);

                    foreach (JobMaterial oJobMateiral in oExcessMaterialIssued)
                    {
                        oSLBOMDocument.SetCellValue(iNumberOfRows, 1, oJobMateiral.JobNum);
                        oSLBOMDocument.SetCellValue(iNumberOfRows, 2, oJobMateiral.ParentPartNum);
                        oSLBOMDocument.SetCellValue(iNumberOfRows, 3, oJobMateiral.ParentRevNum);
                        oSLBOMDocument.SetCellValue(iNumberOfRows, 4, oJobMateiral.AssemblySeq.ToString());
                        oSLBOMDocument.SetCellValue(iNumberOfRows, 5, oJobMateiral.MtlSeq.ToString());
                        oSLBOMDocument.SetCellValue(iNumberOfRows, 6, oJobMateiral.MaterialPartNum);
                        oSLBOMDocument.SetCellValue(iNumberOfRows, 7, oJobMateiral.MaterialPartRevNum);
                        if (oJobMateiral.PartMaster != null)
                        {
                            oSLBOMDocument.SetCellValue(iNumberOfRows, 8, oJobMateiral.PartMaster.ClassID);
                        }
                        oSLBOMDocument.SetCellValue(iNumberOfRows, 9, oJobMateiral.RequiredQty);
                        oSLBOMDocument.SetCellValue(iNumberOfRows, 10, oJobMateiral.IssuedQty);
                        decimal dDelta = oJobMateiral.IssuedQty - oJobMateiral.RequiredQty;
                        oSLBOMDocument.SetCellValue(iNumberOfRows, 11, dDelta);
                        oSLBOMDocument.SetCellStyle(iNumberOfRows, 11, oDecimalStyle);
                        decimal dPercentageError = 0M;
                        if (oJobMateiral.RequiredQty != 0)
                        {
                            dPercentageError = dDelta / oJobMateiral.RequiredQty * 100.0M;
                        }
                        else
                        {
                            dPercentageError = 100.0M;
                        }
                        oSLBOMDocument.SetCellValue(iNumberOfRows, 12, dPercentageError);
                        oSLBOMDocument.SetCellStyle(iNumberOfRows, 12, oDecimalStyle);

                        iNumberOfRows++;
                        bDataInReport = true;
                    }
                }

                #endregion

                #region Material Costs
                // materials without cost
                List<JobMaterial> oMaterialsWithoutCost = oJob.AllMaterials.Where(oItem => (oItem.EstUnitCost == 0) && (oItem.QtyBearing == true) && ((oItem.PartMaster == null) || ((oItem.PartMaster != null) && (m_oPartClassesToIgnore.Contains(oItem.PartMaster.ClassID) == false)))).ToList();
                if (oMaterialsWithoutCost.Count > 0)
                {
                    // sort these by job number
                    oMaterialsWithoutCost = oMaterialsWithoutCost.OrderBy(oItem => oItem.JobNum).ToList();
                    iNumberOfRows = 1;
                    iNumberOfColumns = 1;
                    if (bFirstWorksheet == true)
                    {
                        oSLBOMDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "Mtl Cost Is Zero");
                        bFirstWorksheet = false;
                    }
                    else
                    {
                        oSLBOMDocument.AddWorksheet("Mtl Cost Is Zero");
                    }
                    //set column header
                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "JobNum");
                    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Parent Part Num");
                    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Parent Part Rev Num");
                    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Assembly Seq");
                    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Mtl Seq");
                    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Mtl Part Num");
                    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Mtl Part Rev Num");
                    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Class");
                    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "The Job Material Cost Is Zero -- POC ISSUE PLEASE SET COST");
                    oSLBOMDocument.SetCellStyle(iNumberOfRows++, iNumberOfColumns, oHighlightHeaderStyle);
                    oSLBOMDocument.SetColumnWidth(iNumberOfColumns++, 90);

                    foreach (JobMaterial oJobMateiral in oMaterialsWithoutCost)
                    {
                        oSLBOMDocument.SetCellValue(iNumberOfRows, 1, oJobMateiral.JobNum);
                        oSLBOMDocument.SetCellValue(iNumberOfRows, 2, oJobMateiral.ParentPartNum);
                        oSLBOMDocument.SetCellValue(iNumberOfRows, 3, oJobMateiral.ParentRevNum);
                        oSLBOMDocument.SetCellValue(iNumberOfRows, 4, oJobMateiral.AssemblySeq.ToString());
                        oSLBOMDocument.SetCellValue(iNumberOfRows, 5, oJobMateiral.MtlSeq.ToString());
                        oSLBOMDocument.SetCellValue(iNumberOfRows, 6, oJobMateiral.MaterialPartNum);
                        oSLBOMDocument.SetCellValue(iNumberOfRows, 7, oJobMateiral.MaterialPartRevNum);
                        if (oJobMateiral.PartMaster != null)
                        {
                            oSLBOMDocument.SetCellValue(iNumberOfRows, 8, oJobMateiral.PartMaster.ClassID);
                        }

                        iNumberOfRows++;
                        bDataInReport = true;
                    }
                }

                #region IGNORING
                //////////// materials with unexpected material cost
                //////////List<JobMaterial> oMtlWithUnexpectedMtlCost = oJob.AllMaterials.Where(oItem => (oItem.MaterialCost > (oItem.EstMtlUnitCost * oItem.RequiredQty)) || ((oItem.MaterialCost < (oItem.EstMtlUnitCost * oItem.RequiredQty)) && (oItem.OpComplete == true))).ToList();
                //////////if (oMtlWithUnexpectedMtlCost.Count > 0)
                //////////{
                //////////    // sort these by job number
                //////////    oMtlWithUnexpectedMtlCost = oMtlWithUnexpectedMtlCost.OrderBy(oItem => oItem.JobNum).ToList();
                //////////    iNumberOfRows = 1;
                //////////    iNumberOfColumns = 1;
                //////////    if (bFirstWorksheet == true)
                //////////    {
                //////////        oSLBOMDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "Mtl Unexpected Mtl Cost");
                //////////        bFirstWorksheet = false;
                //////////    }
                //////////    else
                //////////    {
                //////////        oSLBOMDocument.AddWorksheet("Mtl Unexpected Mtl Cost");
                //////////    }
                //////////    //set column header
                //////////    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "JobNum");
                //////////    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                //////////    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Parent Part Num");
                //////////    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                //////////    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Parent Part Rev Num");
                //////////    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                //////////    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Assembly Seq");
                //////////    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                //////////    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Mtl Seq");
                //////////    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                //////////    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Mtl Part Num");
                //////////    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                //////////    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Mtl Part Rev Num");
                //////////    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                //////////    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Class");
                //////////    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                //////////    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Est Mtl Cost");
                //////////    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                //////////    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Act Mtl Cost");
                //////////    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                //////////    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Variance $");
                //////////    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                //////////    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Variance %");
                //////////    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                //////////    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "The Job Material Cost Differs From Estimate -- POC ISSUE PLEASE REVIEW COSTS");
                //////////    oSLBOMDocument.SetCellStyle(iNumberOfRows++, iNumberOfColumns, oHighlightHeaderStyle);
                //////////    oSLBOMDocument.SetColumnWidth(iNumberOfColumns++, 90);

                //////////    foreach (JobMaterial oJobMaterial in oMtlWithUnexpectedMtlCost)
                //////////    {
                //////////        decimal dEstTotal = oJobMaterial.EstMtlUnitCost * oJobMaterial.RequiredQty;
                //////////        decimal dDelta = oJobMaterial.MaterialCost - dEstTotal;
                //////////        decimal dPercentageError = 0M;
                //////////        decimal dMarginToDate = 0M;
                //////////        if (dEstTotal != 0)
                //////////        {
                //////////            dPercentageError = dDelta / dEstTotal * 100.0M;
                //////////        }
                //////////        else
                //////////        {
                //////////            dPercentageError = 100.0M;
                //////////        }
                //////////        HSJob oTmpJob = m_oAllJobs.FirstOrDefault(oItem => string.Compare(oJobMaterial.JobNum, oItem.JobNum, true) == 0);
                //////////        if (oTmpJob != null)
                //////////        {
                //////////            dMarginToDate = oTmpJob.MarginToDate * 100.0M;
                //////////        }

                //////////        bool bExceedsThreholds = false;
                //////////        if ((Math.Abs(dDelta) > m_dJobEstimateAbsoluteError) && (Math.Abs(dPercentageError) > m_dJobEstimatePercentError) && (dMarginToDate < m_dJobMarginThreshold))
                //////////        {
                //////////            // check to see if it exceeds both thresholds
                //////////            bExceedsThreholds = true;
                //////////        }
                //////////        if (m_bForceJobAnalysis == true)
                //////////        {
                //////////            // force it to report
                //////////            bExceedsThreholds = true;
                //////////        }

                //////////        if (bExceedsThreholds == true)
                //////////        {

                //////////            oSLBOMDocument.SetCellValue(iNumberOfRows, 1, oJobMaterial.JobNum);
                //////////            oSLBOMDocument.SetCellValue(iNumberOfRows, 2, oJobMaterial.ParentPartNum);
                //////////            oSLBOMDocument.SetCellValue(iNumberOfRows, 3, oJobMaterial.ParentRevNum);
                //////////            oSLBOMDocument.SetCellValue(iNumberOfRows, 4, oJobMaterial.AssemblySeq.ToString());
                //////////            oSLBOMDocument.SetCellValue(iNumberOfRows, 5, oJobMaterial.MtlSeq.ToString());
                //////////            oSLBOMDocument.SetCellValue(iNumberOfRows, 6, oJobMaterial.MaterialPartNum);
                //////////            oSLBOMDocument.SetCellValue(iNumberOfRows, 7, oJobMaterial.MaterialPartRevNum);
                //////////            if (oJobMaterial.PartMaster != null)
                //////////            {
                //////////                oSLBOMDocument.SetCellValue(iNumberOfRows, 8, oJobMaterial.PartMaster.ClassID);
                //////////            }
                //////////            oSLBOMDocument.SetCellValue(iNumberOfRows, 9, oJobMaterial.EstMtlUnitCost * oJobMaterial.RequiredQty);
                //////////            oSLBOMDocument.SetCellStyle(iNumberOfRows, 9, oCurrencyStyle);
                //////////            oSLBOMDocument.SetCellValue(iNumberOfRows, 10, oJobMaterial.MaterialCost);
                //////////            oSLBOMDocument.SetCellStyle(iNumberOfRows, 10, oCurrencyStyle);

                //////////            oSLBOMDocument.SetCellValue(iNumberOfRows, 11, dDelta);
                //////////            oSLBOMDocument.SetCellStyle(iNumberOfRows, 11, oCurrencyStyle);

                //////////            oSLBOMDocument.SetCellValue(iNumberOfRows, 12, dPercentageError);
                //////////            oSLBOMDocument.SetCellStyle(iNumberOfRows, 12, oDecimalStyle);

                //////////            iNumberOfRows++;
                //////////            bDataInReport = true;
                //////////        }
                //////////    }
                //////////}

                //////////// materials with unexpected burden
                //////////List<JobMaterial> oMtlWithUnexpectedBurden = oJob.AllMaterials.Where(oItem => (oItem.BurdenCost > (oItem.EstBurdenUnitCost * oItem.RequiredQty)) || ((oItem.BurdenCost < (oItem.EstBurdenUnitCost * oItem.RequiredQty)) && (oItem.OpComplete == true))).ToList();
                //////////if (oMtlWithUnexpectedBurden.Count > 0)
                //////////{
                //////////    // sort these by job number
                //////////    oMtlWithUnexpectedBurden = oMtlWithUnexpectedBurden.OrderBy(oItem => oItem.JobNum).ToList();
                //////////    iNumberOfRows = 1;
                //////////    iNumberOfColumns = 1;
                //////////    if (bFirstWorksheet == true)
                //////////    {
                //////////        oSLBOMDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "Mtl Unexpected Burden");
                //////////        bFirstWorksheet = false;
                //////////    }
                //////////    else
                //////////    {
                //////////        oSLBOMDocument.AddWorksheet("Mtl Unexpected Burden");
                //////////    }
                //////////    //set column header
                //////////    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "JobNum");
                //////////    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                //////////    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Parent Part Num");
                //////////    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                //////////    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Parent Part Rev Num");
                //////////    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                //////////    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Assembly Seq");
                //////////    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                //////////    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Mtl Seq");
                //////////    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                //////////    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Mtl Part Num");
                //////////    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                //////////    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Mtl Part Rev Num");
                //////////    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                //////////    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Class");
                //////////    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                //////////    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Est Burden");
                //////////    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                //////////    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Act Burden");
                //////////    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                //////////    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Variance $");
                //////////    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                //////////    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Variance %");
                //////////    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                //////////    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "The Job Material Burden Cost Differs From Estimate -- POC ISSUE PLEASE REVIEW");
                //////////    oSLBOMDocument.SetCellStyle(iNumberOfRows++, iNumberOfColumns, oHighlightHeaderStyle);
                //////////    oSLBOMDocument.SetColumnWidth(iNumberOfColumns++, 90);

                //////////    foreach (JobMaterial oJobMaterial in oMtlWithUnexpectedBurden)
                //////////    {
                //////////        decimal dEstTotal = oJobMaterial.EstBurdenUnitCost * oJobMaterial.RequiredQty;
                //////////        decimal dDelta = oJobMaterial.BurdenCost - dEstTotal;
                //////////        decimal dPercentageError = 0M;
                //////////        decimal dMarginToDate = 0M;
                //////////        if (dEstTotal != 0)
                //////////        {
                //////////            dPercentageError = dDelta / dEstTotal * 100.0M;
                //////////        }
                //////////        else
                //////////        {
                //////////            dPercentageError = 100.0M;
                //////////        }
                //////////        HSJob oTmpJob = m_oAllJobs.FirstOrDefault(oItem => string.Compare(oJobMaterial.JobNum, oItem.JobNum, true) == 0);
                //////////        if (oTmpJob != null)
                //////////        {
                //////////            dMarginToDate = oTmpJob.MarginToDate * 100.0M;
                //////////        }

                //////////        bool bExceedsThreholds = false;
                //////////        if ((Math.Abs(dDelta) > m_dJobEstimateAbsoluteError) && (Math.Abs(dPercentageError) > m_dJobEstimatePercentError) && (dMarginToDate < m_dJobMarginThreshold))
                //////////        {
                //////////            // check to see if it exceeds both thresholds
                //////////            bExceedsThreholds = true;
                //////////        }
                //////////        if (m_bForceJobAnalysis == true)
                //////////        {
                //////////            // force it to report
                //////////            bExceedsThreholds = true;
                //////////        }

                //////////        if (bExceedsThreholds == true)
                //////////        {
                //////////            oSLBOMDocument.SetCellValue(iNumberOfRows, 1, oJobMaterial.JobNum);
                //////////            oSLBOMDocument.SetCellValue(iNumberOfRows, 2, oJobMaterial.ParentPartNum);
                //////////            oSLBOMDocument.SetCellValue(iNumberOfRows, 3, oJobMaterial.ParentRevNum);
                //////////            oSLBOMDocument.SetCellValue(iNumberOfRows, 4, oJobMaterial.AssemblySeq.ToString());
                //////////            oSLBOMDocument.SetCellValue(iNumberOfRows, 5, oJobMaterial.MtlSeq.ToString());
                //////////            oSLBOMDocument.SetCellValue(iNumberOfRows, 6, oJobMaterial.MaterialPartNum);
                //////////            oSLBOMDocument.SetCellValue(iNumberOfRows, 7, oJobMaterial.MaterialPartRevNum);
                //////////            if (oJobMaterial.PartMaster != null)
                //////////            {
                //////////                oSLBOMDocument.SetCellValue(iNumberOfRows, 8, oJobMaterial.PartMaster.ClassID);
                //////////            }
                //////////            oSLBOMDocument.SetCellValue(iNumberOfRows, 9, oJobMaterial.EstBurdenUnitCost * oJobMaterial.RequiredQty);
                //////////            oSLBOMDocument.SetCellStyle(iNumberOfRows, 9, oCurrencyStyle);
                //////////            oSLBOMDocument.SetCellValue(iNumberOfRows, 10, oJobMaterial.BurdenCost);
                //////////            oSLBOMDocument.SetCellStyle(iNumberOfRows, 10, oCurrencyStyle);

                //////////            oSLBOMDocument.SetCellValue(iNumberOfRows, 11, dDelta);
                //////////            oSLBOMDocument.SetCellStyle(iNumberOfRows, 11, oCurrencyStyle);

                //////////            oSLBOMDocument.SetCellValue(iNumberOfRows, 12, dPercentageError);
                //////////            oSLBOMDocument.SetCellStyle(iNumberOfRows, 12, oDecimalStyle);

                //////////            iNumberOfRows++;
                //////////            bDataInReport = true;
                //////////        }
                //////////    }
                //////////}

                //////////// materials with unexpected labor
                //////////List<JobMaterial> oMtlWithUnexpectedLabor = oJob.AllMaterials.Where(oItem => (oItem.LaborCost > (oItem.EstLaborUnitCost * oItem.RequiredQty)) || ((oItem.LaborCost < (oItem.EstLaborUnitCost * oItem.RequiredQty)) && (oItem.OpComplete == true))).ToList();
                //////////if (oMtlWithUnexpectedLabor.Count > 0)
                //////////{
                //////////    // sort these by job number
                //////////    oMtlWithUnexpectedLabor = oMtlWithUnexpectedLabor.OrderBy(oItem => oItem.JobNum).ToList();
                //////////    iNumberOfRows = 1;
                //////////    iNumberOfColumns = 1;
                //////////    if (bFirstWorksheet == true)
                //////////    {
                //////////        oSLBOMDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "Mtl Unexpected Labor");
                //////////        bFirstWorksheet = false;
                //////////    }
                //////////    else
                //////////    {
                //////////        oSLBOMDocument.AddWorksheet("Mtl Unexpected Labor");
                //////////    }
                //////////    //set column header
                //////////    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "JobNum");
                //////////    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                //////////    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Parent Part Num");
                //////////    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                //////////    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Parent Part Rev Num");
                //////////    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                //////////    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Assembly Seq");
                //////////    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                //////////    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Mtl Seq");
                //////////    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                //////////    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Mtl Part Num");
                //////////    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                //////////    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Mtl Part Rev Num");
                //////////    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                //////////    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Class");
                //////////    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                //////////    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Est Labor");
                //////////    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                //////////    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Act Labor");
                //////////    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                //////////    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Variance $");
                //////////    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                //////////    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Variance %");
                //////////    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                //////////    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "The Job Material Labor Cost Differs From Estimate -- POC ISSUE PLEASE REVIEW");
                //////////    oSLBOMDocument.SetCellStyle(iNumberOfRows++, iNumberOfColumns, oHighlightHeaderStyle);
                //////////    oSLBOMDocument.SetColumnWidth(iNumberOfColumns++, 90);

                //////////    foreach (JobMaterial oJobMaterial in oMtlWithUnexpectedLabor)
                //////////    {
                //////////        decimal dEstTotal = oJobMaterial.EstLaborUnitCost * oJobMaterial.RequiredQty;
                //////////        decimal dDelta = oJobMaterial.LaborCost - dEstTotal;
                //////////        decimal dPercentageError = 0M;
                //////////        decimal dMarginToDate = 0M;
                //////////        if (dEstTotal != 0)
                //////////        {
                //////////            dPercentageError = dDelta / dEstTotal * 100.0M;
                //////////        }
                //////////        else
                //////////        {
                //////////            dPercentageError = 100.0M;
                //////////        }
                //////////        HSJob oTmpJob = m_oAllJobs.FirstOrDefault(oItem => string.Compare(oJobMaterial.JobNum, oItem.JobNum, true) == 0);
                //////////        if (oTmpJob != null)
                //////////        {
                //////////            dMarginToDate = oTmpJob.MarginToDate * 100.0M;
                //////////        }

                //////////        bool bExceedsThreholds = false;
                //////////        if ((Math.Abs(dDelta) > m_dJobEstimateAbsoluteError) && (Math.Abs(dPercentageError) > m_dJobEstimatePercentError) && (dMarginToDate < m_dJobMarginThreshold))
                //////////        {
                //////////            // check to see if it exceeds both thresholds
                //////////            bExceedsThreholds = true;
                //////////        }
                //////////        if (m_bForceJobAnalysis == true)
                //////////        {
                //////////            // force it to report
                //////////            bExceedsThreholds = true;
                //////////        }

                //////////        if (bExceedsThreholds == true)
                //////////        {
                //////////            oSLBOMDocument.SetCellValue(iNumberOfRows, 1, oJobMaterial.JobNum);
                //////////            oSLBOMDocument.SetCellValue(iNumberOfRows, 2, oJobMaterial.ParentPartNum);
                //////////            oSLBOMDocument.SetCellValue(iNumberOfRows, 3, oJobMaterial.ParentRevNum);
                //////////            oSLBOMDocument.SetCellValue(iNumberOfRows, 4, oJobMaterial.AssemblySeq.ToString());
                //////////            oSLBOMDocument.SetCellValue(iNumberOfRows, 5, oJobMaterial.MtlSeq.ToString());
                //////////            oSLBOMDocument.SetCellValue(iNumberOfRows, 6, oJobMaterial.MaterialPartNum);
                //////////            oSLBOMDocument.SetCellValue(iNumberOfRows, 7, oJobMaterial.MaterialPartRevNum);
                //////////            if (oJobMaterial.PartMaster != null)
                //////////            {
                //////////                oSLBOMDocument.SetCellValue(iNumberOfRows, 8, oJobMaterial.PartMaster.ClassID);
                //////////            }
                //////////            oSLBOMDocument.SetCellValue(iNumberOfRows, 9, oJobMaterial.EstLaborUnitCost * oJobMaterial.RequiredQty);
                //////////            oSLBOMDocument.SetCellStyle(iNumberOfRows, 9, oCurrencyStyle);
                //////////            oSLBOMDocument.SetCellValue(iNumberOfRows, 10, oJobMaterial.LaborCost);
                //////////            oSLBOMDocument.SetCellStyle(iNumberOfRows, 10, oCurrencyStyle);

                //////////            oSLBOMDocument.SetCellValue(iNumberOfRows, 11, dDelta);
                //////////            oSLBOMDocument.SetCellStyle(iNumberOfRows, 11, oCurrencyStyle);

                //////////            oSLBOMDocument.SetCellValue(iNumberOfRows, 12, dPercentageError);
                //////////            oSLBOMDocument.SetCellStyle(iNumberOfRows, 12, oDecimalStyle);

                //////////            iNumberOfRows++;
                //////////            bDataInReport = true;
                //////////        }
                //////////    }
                //////////}

                //////////// materials wtih unepected subcontract costs
                //////////List<JobMaterial> oMtlWithUnexpectedSubcontractCosts = oJob.AllMaterials.Where(oItem => (oItem.SubcontractCost > (oItem.EstSubcontractUnitCost * oItem.RequiredQty)) || ((oItem.SubcontractCost < (oItem.EstSubcontractUnitCost * oItem.RequiredQty)) && (oItem.OpComplete == true))).ToList();
                //////////if (oMtlWithUnexpectedSubcontractCosts.Count > 0)
                //////////{
                //////////    // sort these by job number
                //////////    oMtlWithUnexpectedSubcontractCosts = oMtlWithUnexpectedSubcontractCosts.OrderBy(oItem => oItem.JobNum).ToList();
                //////////    iNumberOfRows = 1;
                //////////    iNumberOfColumns = 1;
                //////////    if (bFirstWorksheet == true)
                //////////    {
                //////////        oSLBOMDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "Mtl Unexpected Subcontract");
                //////////        bFirstWorksheet = false;
                //////////    }
                //////////    else
                //////////    {
                //////////        oSLBOMDocument.AddWorksheet("Mtl Unexpected Subcontract");
                //////////    }
                //////////    //set column header
                //////////    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "JobNum");
                //////////    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                //////////    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Parent Part Num");
                //////////    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                //////////    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Parent Part Rev Num");
                //////////    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                //////////    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Assembly Seq");
                //////////    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                //////////    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Mtl Seq");
                //////////    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                //////////    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Mtl Part Num");
                //////////    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                //////////    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Mtl Part Rev Num");
                //////////    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                //////////    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Class");
                //////////    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                //////////    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Est Subcontract");
                //////////    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                //////////    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Act Subcontract");
                //////////    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                //////////    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Variance $");
                //////////    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                //////////    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Variance %");
                //////////    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                //////////    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "The Job Material Subcontract Cost Diffres From Estimate -- POC ISSUE PLEASE REVIEW");
                //////////    oSLBOMDocument.SetCellStyle(iNumberOfRows++, iNumberOfColumns, oHighlightHeaderStyle);
                //////////    oSLBOMDocument.SetColumnWidth(iNumberOfColumns++, 90);

                //////////    foreach (JobMaterial oJobMaterial in oMtlWithUnexpectedLabor)
                //////////    {
                //////////        decimal dEstTotal = oJobMaterial.EstSubcontractUnitCost * oJobMaterial.RequiredQty;
                //////////        decimal dDelta = oJobMaterial.SubcontractCost - dEstTotal;
                //////////        decimal dPercentageError = 0M;
                //////////        decimal dMarginToDate = 0M;
                //////////        if (dEstTotal != 0)
                //////////        {
                //////////            dPercentageError = dDelta / dEstTotal * 100.0M;
                //////////        }
                //////////        else
                //////////        {
                //////////            dPercentageError = 100.0M;
                //////////        }
                //////////        HSJob oTmpJob = m_oAllJobs.FirstOrDefault(oItem => string.Compare(oJobMaterial.JobNum, oItem.JobNum, true) == 0);
                //////////        if (oTmpJob != null)
                //////////        {
                //////////            dMarginToDate = oTmpJob.MarginToDate * 100.0M;
                //////////        }

                //////////        bool bExceedsThreholds = false;
                //////////        if ((Math.Abs(dDelta) > m_dJobEstimateAbsoluteError) && (Math.Abs(dPercentageError) > m_dJobEstimatePercentError) && (dMarginToDate < m_dJobMarginThreshold))
                //////////        {
                //////////            // check to see if it exceeds both thresholds
                //////////            bExceedsThreholds = true;
                //////////        }
                //////////        if (m_bForceJobAnalysis == true)
                //////////        {
                //////////            // force it to report
                //////////            bExceedsThreholds = true;
                //////////        }

                //////////        if (bExceedsThreholds == true)
                //////////        {
                //////////            oSLBOMDocument.SetCellValue(iNumberOfRows, 1, oJobMaterial.JobNum);
                //////////            oSLBOMDocument.SetCellValue(iNumberOfRows, 2, oJobMaterial.ParentPartNum);
                //////////            oSLBOMDocument.SetCellValue(iNumberOfRows, 3, oJobMaterial.ParentRevNum);
                //////////            oSLBOMDocument.SetCellValue(iNumberOfRows, 4, oJobMaterial.AssemblySeq.ToString());
                //////////            oSLBOMDocument.SetCellValue(iNumberOfRows, 5, oJobMaterial.MtlSeq.ToString());
                //////////            oSLBOMDocument.SetCellValue(iNumberOfRows, 6, oJobMaterial.MaterialPartNum);
                //////////            oSLBOMDocument.SetCellValue(iNumberOfRows, 7, oJobMaterial.MaterialPartRevNum);
                //////////            if (oJobMaterial.PartMaster != null)
                //////////            {
                //////////                oSLBOMDocument.SetCellValue(iNumberOfRows, 8, oJobMaterial.PartMaster.ClassID);
                //////////            }
                //////////            oSLBOMDocument.SetCellValue(iNumberOfRows, 9, oJobMaterial.EstSubcontractUnitCost * oJobMaterial.RequiredQty);
                //////////            oSLBOMDocument.SetCellStyle(iNumberOfRows, 9, oCurrencyStyle);
                //////////            oSLBOMDocument.SetCellValue(iNumberOfRows, 10, oJobMaterial.SubcontractCost);
                //////////            oSLBOMDocument.SetCellStyle(iNumberOfRows, 10, oCurrencyStyle);

                //////////            oSLBOMDocument.SetCellValue(iNumberOfRows, 11, dDelta);
                //////////            oSLBOMDocument.SetCellStyle(iNumberOfRows, 11, oCurrencyStyle);

                //////////            oSLBOMDocument.SetCellValue(iNumberOfRows, 12, dPercentageError);
                //////////            oSLBOMDocument.SetCellStyle(iNumberOfRows, 12, oDecimalStyle);

                //////////            iNumberOfRows++;
                //////////            bDataInReport = true;
                //////////        }
                //////////    }
                //////////}
                #endregion
                
                #endregion

                #region Operation Attributes
                // resource id set on operation
                // we should not specify a resource id on the operation -- too specific
                List<JobOperation> oOperationsWithResourceIdSet = oJob.AllOperations.Where(oItem => string.IsNullOrEmpty(oItem.ResourceId) == false).ToList();
                if (oOperationsWithResourceIdSet.Count > 0)
                {
                    // sort these by job number
                    oOperationsWithResourceIdSet = oOperationsWithResourceIdSet.OrderBy(oItem => oItem.JobNum).ToList();
                    iNumberOfRows = 1;
                    iNumberOfColumns = 1;
                    if (bFirstWorksheet == true)
                    {
                        oSLBOMDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "Resource Id Set On Op");
                        bFirstWorksheet = false;
                    }
                    else
                    {
                        oSLBOMDocument.AddWorksheet("Resource Id Set On Op");
                    }
                    //set column header
                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "JobNum");
                    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Parent Part Num");
                    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Parent Part Rev Num");
                    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Assembly Seq");
                    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Opr Seq");
                    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Op Code");
                    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Resource");
                    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "The Resource Id Is Set On Operation -- SOP IS TO NOT SET RESOURCE ID ONLY RESOURCE GROUP");
                    oSLBOMDocument.SetCellStyle(iNumberOfRows++, iNumberOfColumns, oHighlightHeaderStyle);
                    oSLBOMDocument.SetColumnWidth(iNumberOfColumns++, 90);

                    foreach (JobOperation oJobOperation in oOperationsWithResourceIdSet)
                    {
                        oSLBOMDocument.SetCellValue(iNumberOfRows, 1, oJobOperation.JobNum);
                        oSLBOMDocument.SetCellValue(iNumberOfRows, 2, oJobOperation.ParentPartNum);
                        oSLBOMDocument.SetCellValue(iNumberOfRows, 3, oJobOperation.ParentRevNum);
                        oSLBOMDocument.SetCellValue(iNumberOfRows, 4, oJobOperation.AssemblySeq.ToString());
                        oSLBOMDocument.SetCellValue(iNumberOfRows, 5, oJobOperation.OperationSeq.ToString());
                        oSLBOMDocument.SetCellValue(iNumberOfRows, 6, oJobOperation.OpCode);
                        oSLBOMDocument.SetCellValue(iNumberOfRows++, 7, oJobOperation.ResourceId);

                        bDataInReport = true;
                    }
                }

                // ops must have due date
                List<JobOperation> oOperationsWithoutDueDate = oJob.AllOperations.Where(oItem => oItem.DueDate == DateTime.MinValue).ToList();
                if (oOperationsWithoutDueDate.Count > 0)
                {
                    // sort these by job number
                    oOperationsWithoutDueDate = oOperationsWithoutDueDate.OrderBy(oItem => oItem.JobNum).ToList();
                    iNumberOfRows = 1;
                    iNumberOfColumns = 1;
                    if (bFirstWorksheet == true)
                    {
                        oSLBOMDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "Op No Due Date");
                        bFirstWorksheet = false;
                    }
                    else
                    {
                        oSLBOMDocument.AddWorksheet("Op No Due Date");
                    }
                    //set column header
                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "JobNum");
                    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Parent Part Num");
                    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Parent Part Rev Num");
                    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Assembly Seq");
                    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Opr Seq");
                    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Op Code");
                    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Operation Due Date Not Set On Operation -- SOP IS TO SET OP DUE DATE FOR SCHEUDLING AND PLANNING");
                    oSLBOMDocument.SetCellStyle(iNumberOfRows++, iNumberOfColumns, oHighlightHeaderStyle);
                    oSLBOMDocument.SetColumnWidth(iNumberOfColumns++, 90);

                    foreach (JobOperation oJobOperation in oOperationsWithoutDueDate)
                    {
                        oSLBOMDocument.SetCellValue(iNumberOfRows, 1, oJobOperation.JobNum);
                        oSLBOMDocument.SetCellValue(iNumberOfRows, 2, oJobOperation.ParentPartNum);
                        oSLBOMDocument.SetCellValue(iNumberOfRows, 3, oJobOperation.ParentRevNum);
                        oSLBOMDocument.SetCellValue(iNumberOfRows, 4, oJobOperation.AssemblySeq.ToString());
                        oSLBOMDocument.SetCellValue(iNumberOfRows, 5, oJobOperation.OperationSeq.ToString());
                        oSLBOMDocument.SetCellValue(iNumberOfRows++, 6, oJobOperation.OpCode);

                        bDataInReport = true;
                    }
                }

                // ops must have start date
                List<JobOperation> oOperationsWithoutStartDate = oJob.AllOperations.Where(oItem => oItem.StartDate == DateTime.MinValue).ToList();
                if (oOperationsWithoutStartDate.Count > 0)
                {
                    // sort these by job number
                    oOperationsWithoutStartDate = oOperationsWithoutStartDate.OrderBy(oItem => oItem.JobNum).ToList();
                    iNumberOfRows = 1;
                    iNumberOfColumns = 1;
                    if (bFirstWorksheet == true)
                    {
                        oSLBOMDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "Op No Start Date");
                        bFirstWorksheet = false;
                    }
                    else
                    {
                        oSLBOMDocument.AddWorksheet("Op No Start Date");
                    }
                    //set column header
                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "JobNum");
                    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Parent Part Num");
                    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Parent Part Rev Num");
                    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Assembly Seq");
                    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Opr Seq");
                    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Op Code");
                    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Operation Start Date Not Set On Operation -- SOP IS TO SET OP START DATE FOR SCHEUDLING AND PLANNING");
                    oSLBOMDocument.SetCellStyle(iNumberOfRows++, iNumberOfColumns, oHighlightHeaderStyle);
                    oSLBOMDocument.SetColumnWidth(iNumberOfColumns++, 90);

                    foreach (JobOperation oJobOperation in oOperationsWithoutStartDate)
                    {
                        oSLBOMDocument.SetCellValue(iNumberOfRows, 1, oJobOperation.JobNum);
                        oSLBOMDocument.SetCellValue(iNumberOfRows, 2, oJobOperation.ParentPartNum);
                        oSLBOMDocument.SetCellValue(iNumberOfRows, 3, oJobOperation.ParentRevNum);
                        oSLBOMDocument.SetCellValue(iNumberOfRows, 4, oJobOperation.AssemblySeq.ToString());
                        oSLBOMDocument.SetCellValue(iNumberOfRows, 5, oJobOperation.OperationSeq.ToString());
                        oSLBOMDocument.SetCellValue(iNumberOfRows++, 6, oJobOperation.OpCode);

                        bDataInReport = true;
                    }
                }

                // check to make sure the resource group is active
                List<JobOperation> oOperationsWithResourceGroup = oJob.AllOperations.Where(oItem => (string.IsNullOrEmpty(oItem.ResourceGroupId) == false)).ToList();
                List<JobOperation> oOpsWithInactiveResourceGroups = new List<JobOperation>();
                foreach (JobOperation oJobOp in oOperationsWithResourceGroup)
                {
                    if (m_oBOMSupport.IsResourceGroupActive(oJobOp.ResourceGroupId) == false)
                    {
                        oOpsWithInactiveResourceGroups.Add(oJobOp);
                    }
                }
                if (oOpsWithInactiveResourceGroups.Count > 0)
                {
                    // sort these by job number
                    oOpsWithInactiveResourceGroups = oOpsWithInactiveResourceGroups.OrderBy(oItem => oItem.JobNum).ToList();
                    iNumberOfRows = 1;
                    iNumberOfColumns = 1;
                    if (bFirstWorksheet == true)
                    {
                        oSLBOMDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "Op With Inactive Res Grp");
                        bFirstWorksheet = false;
                    }
                    else
                    {
                        oSLBOMDocument.AddWorksheet("Op With Inactive Res Grp");
                    }
                    //set column header
                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "JobNum");
                    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Parent Part Num");
                    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Parent Part Rev Num");
                    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Assembly Seq");
                    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Opr Seq");
                    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Op Code");
                    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Resource Group");
                    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Operation Uses Inactive Resource Group -- POC ISSUE THIS JOB CANNOT BE SCHEUDLED WITH INACTIVE RESOURCE GROUP");
                    oSLBOMDocument.SetCellStyle(iNumberOfRows++, iNumberOfColumns, oHighlightHeaderStyle);
                    oSLBOMDocument.SetColumnWidth(iNumberOfColumns++, 90);

                    foreach (JobOperation oJobOperation in oOpsWithInactiveResourceGroups)
                    {
                        oSLBOMDocument.SetCellValue(iNumberOfRows, 1, oJobOperation.JobNum);
                        oSLBOMDocument.SetCellValue(iNumberOfRows, 2, oJobOperation.ParentPartNum);
                        oSLBOMDocument.SetCellValue(iNumberOfRows, 3, oJobOperation.ParentRevNum);
                        oSLBOMDocument.SetCellValue(iNumberOfRows, 4, oJobOperation.AssemblySeq.ToString());
                        oSLBOMDocument.SetCellValue(iNumberOfRows, 5, oJobOperation.OperationSeq.ToString());
                        oSLBOMDocument.SetCellValue(iNumberOfRows, 6, oJobOperation.OpCode);
                        oSLBOMDocument.SetCellValue(iNumberOfRows++, 7, oJobOperation.ResourceGroupId);

                        bDataInReport = true;
                    }
                }

                // check to make sure the resource is active
                List<JobOperation> oOperationsWithResources = oJob.AllOperations.Where(oItem => (string.IsNullOrEmpty(oItem.ResourceId) == false)).ToList();
                List<JobOperation> oOpsWithInactiveResources = new List<JobOperation>();
                foreach (JobOperation oJobOp in oOperationsWithResources)
                {
                    if (m_oBOMSupport.IsResourceActive(oJobOp.ResourceId) == false)
                    {
                        oOpsWithInactiveResources.Add(oJobOp);
                    }
                }
                if (oOpsWithInactiveResources.Count > 0)
                {
                    // sort these by job number
                    oOpsWithInactiveResources = oOpsWithInactiveResources.OrderBy(oItem => oItem.JobNum).ToList();
                    iNumberOfRows = 1;
                    iNumberOfColumns = 1;
                    if (bFirstWorksheet == true)
                    {
                        oSLBOMDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "Op With Inactive Res");
                        bFirstWorksheet = false;
                    }
                    else
                    {
                        oSLBOMDocument.AddWorksheet("Op With Inactive Res");
                    }
                    //set column header
                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "JobNum");
                    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Parent Part Num");
                    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Parent Part Rev Num");
                    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Assembly Seq");
                    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Opr Seq");
                    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Op Code");
                    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Resource Grouo");
                    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Resource");
                    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);

                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Operation Uses Inactive Resource -- POC ISSUE THIS JOB CANNOT BE SCHEUDLED WITH INACTIVE RESOURCES");
                    oSLBOMDocument.SetCellStyle(iNumberOfRows++, iNumberOfColumns, oHighlightHeaderStyle);
                    oSLBOMDocument.SetColumnWidth(iNumberOfColumns++, 90);

                    foreach (JobOperation oJobOperation in oOpsWithInactiveResources)
                    {
                        oSLBOMDocument.SetCellValue(iNumberOfRows, 1, oJobOperation.JobNum);
                        oSLBOMDocument.SetCellValue(iNumberOfRows, 2, oJobOperation.ParentPartNum);
                        oSLBOMDocument.SetCellValue(iNumberOfRows, 3, oJobOperation.ParentRevNum);
                        oSLBOMDocument.SetCellValue(iNumberOfRows, 4, oJobOperation.AssemblySeq.ToString());
                        oSLBOMDocument.SetCellValue(iNumberOfRows, 5, oJobOperation.OperationSeq.ToString());
                        oSLBOMDocument.SetCellValue(iNumberOfRows, 6, oJobOperation.OpCode);
                        oSLBOMDocument.SetCellValue(iNumberOfRows, 7, oJobOperation.ResourceGroupId);
                        oSLBOMDocument.SetCellValue(iNumberOfRows++, 8, oJobOperation.ResourceId);

                        bDataInReport = true;
                    }
                }

                // prod standard is zero
                List<JobOperation> oOperationsWithZeroEstimatedTime = oJob.AllOperations.Where(oItem => (oItem.ProdStandard == 0) && (oItem.Subcontract == false)).ToList();
                if (oOperationsWithZeroEstimatedTime.Count > 0)
                {
                    // sort these by job number
                    oOperationsWithZeroEstimatedTime = oOperationsWithZeroEstimatedTime.OrderBy(oItem => oItem.JobNum).ToList();
                    iNumberOfRows = 1;
                    iNumberOfColumns = 1;
                    if (bFirstWorksheet == true)
                    {
                        oSLBOMDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "Op No Est Time");
                        bFirstWorksheet = false;
                    }
                    else
                    {
                        oSLBOMDocument.AddWorksheet("Op No Est Time");
                    }
                    //set column header
                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "JobNum");
                    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Parent Part Num");
                    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Parent Part Rev Num");
                    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Assembly Seq");
                    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Opr Seq");
                    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Op Code");
                    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "The Operation Has No Time Estimate -- POC ISSUE PLEASE SET EXPECTED TIME");
                    oSLBOMDocument.SetCellStyle(iNumberOfRows++, iNumberOfColumns, oHighlightHeaderStyle);
                    oSLBOMDocument.SetColumnWidth(iNumberOfColumns++, 90);

                    foreach (JobOperation oJobOperation in oOperationsWithZeroEstimatedTime)
                    {
                        oSLBOMDocument.SetCellValue(iNumberOfRows, 1, oJobOperation.JobNum);
                        oSLBOMDocument.SetCellValue(iNumberOfRows, 2, oJobOperation.ParentPartNum);
                        oSLBOMDocument.SetCellValue(iNumberOfRows, 3, oJobOperation.ParentRevNum);
                        oSLBOMDocument.SetCellValue(iNumberOfRows, 4, oJobOperation.AssemblySeq.ToString());
                        oSLBOMDocument.SetCellValue(iNumberOfRows, 5, oJobOperation.OperationSeq.ToString());
                        oSLBOMDocument.SetCellValue(iNumberOfRows++, 6, oJobOperation.OpCode);

                        bDataInReport = true;
                    }
                }

                // test if labor rate is set to zero
                // first get list of operations with time set and a run qty
                List<JobOperation> oOperationsWithHours = oJob.AllOperations.Where(oItem => (oItem.ProdStandard != 0) && (oItem.RunQty != 0) && (oItem.Subcontract == false)).ToList();
                // get a list of operations where labor cost is zero
                List<JobOpsEstVsActualCosts> oOperationsWithNoLaborCost = oJob.AllOperationCosts.Where(oItem => oItem.EstLaborCost == 0).ToList();
                // order by job num
                oOperationsWithNoLaborCost = oOperationsWithNoLaborCost.OrderBy(oItem => oItem.JobNum).ToList();
                // walk through this list of ops with no labor cost and if they have operational time then this means the labor rate is zero
                bool bSetHeader = false;
                foreach (JobOpsEstVsActualCosts oOpCost in oOperationsWithNoLaborCost)
                {
                    JobOperation oJobOperation = oOperationsWithHours.FirstOrDefault(oItem => (string.Compare(oItem.JobNum, oOpCost.JobNum, true) == 0) && (oItem.AssemblySeq == oOpCost.AssemblySeq) && (oItem.OperationSeq == oOpCost.OprSeq) && (string.Compare(oItem.OpCode, oOpCost.OpCode, true) == 0));
                    if (oJobOperation != null)
                    {
                        // this operation has hours but the labor comes out to $0 so that means we must have a labor rate of $0
                        if (bSetHeader == false)
                        {
                            bSetHeader = true;
                            iNumberOfRows = 1;
                            iNumberOfColumns = 1;
                            if (bFirstWorksheet == true)
                            {
                                oSLBOMDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "Ops With No Labor Rate");
                                bFirstWorksheet = false;
                            }
                            else
                            {
                                oSLBOMDocument.AddWorksheet("Ops With No Labor Rate");
                            }
                            //set column header
                            oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "JobNum");
                            oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                            oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Parent Part Num");
                            oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                            oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Parent Part Rev Num");
                            oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                            oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Assembly Seq");
                            oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                            oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Opr Seq");
                            oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                            oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Op Code");
                            oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                            oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Labor Rate Is $0 For Operation -- POC ISSUE PLEASE SET LABOR RATE ON OPERATION");
                            oSLBOMDocument.SetCellStyle(iNumberOfRows++, iNumberOfColumns, oHighlightHeaderStyle);
                            oSLBOMDocument.SetColumnWidth(iNumberOfColumns++, 90);
                        }

                        // indicate which operation has a zero dollar labor rate
                        oSLBOMDocument.SetCellValue(iNumberOfRows, 1, oJobOperation.JobNum);
                        oSLBOMDocument.SetCellValue(iNumberOfRows, 2, oJobOperation.ParentPartNum);
                        oSLBOMDocument.SetCellValue(iNumberOfRows, 3, oJobOperation.ParentRevNum);
                        oSLBOMDocument.SetCellValue(iNumberOfRows, 4, oJobOperation.AssemblySeq.ToString());
                        oSLBOMDocument.SetCellValue(iNumberOfRows, 5, oJobOperation.OperationSeq.ToString());
                        oSLBOMDocument.SetCellValue(iNumberOfRows++, 6, oJobOperation.OpCode);

                        bDataInReport = true;
                    }
                }

                // test if burden rate is set to zero
                // we will use the list of operations with time set and a run qty from above
                // get a list of operations where burden cost is zero
                List<JobOpsEstVsActualCosts> oOperationsWithNoBurdenCost = oJob.AllOperationCosts.Where(oItem => oItem.EstBurCost == 0).ToList();
                // walk through this list of ops with no burden cost and if they have operational time then this means the burden rate is zero
                bSetHeader = false;
                foreach (JobOpsEstVsActualCosts oOpCost in oOperationsWithNoBurdenCost)
                {
                    JobOperation oJobOperation = oOperationsWithHours.FirstOrDefault(oItem => (string.Compare(oItem.JobNum, oOpCost.JobNum, true) == 0) && (oItem.AssemblySeq == oOpCost.AssemblySeq) && (oItem.OperationSeq == oOpCost.OprSeq) && (string.Compare(oItem.OpCode, oOpCost.OpCode, true) == 0));
                    if (oJobOperation != null)
                    {
                        // this operation has hours but the burden comes out to $0 so that means we must have a burden rate of $0
                        if (bSetHeader == false)
                        {
                            bSetHeader = true;
                            iNumberOfRows = 1;
                            iNumberOfColumns = 1;
                            if (bFirstWorksheet == true)
                            {
                                oSLBOMDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "Ops With No Burden Rate");
                                bFirstWorksheet = false;
                            }
                            else
                            {
                                oSLBOMDocument.AddWorksheet("Ops With No Burden Rate");
                            }
                            //set column header
                            oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "JobNum");
                            oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                            oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Parent Part Num");
                            oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                            oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Parent Part Rev Num");
                            oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                            oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Assembly Seq");
                            oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                            oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Opr Seq");
                            oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                            oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Op Code");
                            oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                            oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Burden Rate Is $0 For Operation -- POC ISSUE PLEASE SET BURDEN RATE ON OPERATION");
                            oSLBOMDocument.SetCellStyle(iNumberOfRows++, iNumberOfColumns, oHighlightHeaderStyle);
                            oSLBOMDocument.SetColumnWidth(iNumberOfColumns++, 90);
                        }

                        // indicate which operation has a zero dollar burden rate
                        oSLBOMDocument.SetCellValue(iNumberOfRows, 1, oJobOperation.JobNum);
                        oSLBOMDocument.SetCellValue(iNumberOfRows, 2, oJobOperation.ParentPartNum);
                        oSLBOMDocument.SetCellValue(iNumberOfRows, 3, oJobOperation.ParentRevNum);
                        oSLBOMDocument.SetCellValue(iNumberOfRows, 4, oJobOperation.AssemblySeq.ToString());
                        oSLBOMDocument.SetCellValue(iNumberOfRows, 5, oJobOperation.OperationSeq.ToString());
                        oSLBOMDocument.SetCellValue(iNumberOfRows++, 6, oJobOperation.OpCode);

                        bDataInReport = true;
                    }
                }

                #region IGNORING THESE ISSUES
                // std format is wrong
                // standard format should be "HP" hours / piece
                //////////List<JobOperation> oOperationsWithBadStandardFormat = m_oJobOperations.Where(oItem => (string.Compare(oItem.StdFormat, "HP", true) != 0)).ToList();
                //////////if (oOperationsWithBadStandardFormat.Count > 0)
                //////////{
                //////////    // sort these by job number
                //////////    oOperationsWithBadStandardFormat = oOperationsWithBadStandardFormat.OrderBy(oItem => oItem.JobNum).ToList();
                //////////    iNumberOfRows = 1;
                //////////    iNumberOfColumns = 1;
                //////////    if (bFirstWorksheet == true)
                //////////    {
                //////////        oSLBOMDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "Op Bad Std Format");
                //////////        bFirstWorksheet = false;
                //////////    }
                //////////    else
                //////////    {
                //////////        oSLBOMDocument.AddWorksheet("Op Bad Std Format");
                //////////    }
                //////////    //set column header
                //////////    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "JobNum");
                //////////    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                //////////    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Parent Part Num");
                //////////    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                //////////    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Parent Part Rev Num");
                //////////    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                //////////    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Assembly Seq");
                //////////    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                //////////    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Opr Seq");
                //////////    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                //////////    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Op Code");
                //////////    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                //////////    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Std Format");
                //////////    oSLBOMDocument.SetColumnWidth(iNumberOfRows++, iNumberOfColumns++, iStandardColumnWidth);

                //////////    foreach (JobOperation oJobOperation in oOperationsWithBadStandardFormat)
                //////////    {
                //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, 1, oJobOperation.JobNum);
                //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, 2, oJobOperation.ParentPartNum);
                //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, 3, oJobOperation.ParentRevNum);
                //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, 4, oJobOperation.AssemblySeq.ToString());
                //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, 5, oJobOperation.OperationSeq.ToString());
                //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, 6, oJobOperation.OpCode);
                //////////        oSLBOMDocument.SetCellValue(iNumberOfRows++, 7, oJobOperation.StdFormat);

                //////////        bDataInReport = true;
                //////////    }
                //////////}
                #endregion

                // if std format is "OM" or "OH" then the OpsPerPart field must be > 0
                List<JobOperation> oOperationsPerPartSetToZero = oJob.AllOperations.Where(oItem => ((string.Compare(oItem.StdFormat, "OM", true) == 0) || (string.Compare(oItem.StdFormat, "OH", true) == 0)) && (oItem.OperationsPerPart == 0)).ToList();
                if (oOperationsPerPartSetToZero.Count > 0)
                {
                    // sort these by job number
                    oOperationsPerPartSetToZero = oOperationsPerPartSetToZero.OrderBy(oItem => oItem.JobNum).ToList();
                    iNumberOfRows = 1;
                    iNumberOfColumns = 1;
                    if (bFirstWorksheet == true)
                    {
                        oSLBOMDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "Ops Per Part Is Zero");
                        bFirstWorksheet = false;
                    }
                    else
                    {
                        oSLBOMDocument.AddWorksheet("Ops Per Part Is Zero");
                    }
                    //set column header
                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "JobNum");
                    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Parent Part Num");
                    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Parent Part Rev Num");
                    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Assembly Seq");
                    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Opr Seq");
                    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Op Code");
                    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Std Format");
                    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "The Operation Is Set To Operations Per Part -- NEED TO SET QTY ON OPERATIONS PER PART");
                    oSLBOMDocument.SetCellStyle(iNumberOfRows++, iNumberOfColumns, oHighlightHeaderStyle);
                    oSLBOMDocument.SetColumnWidth(iNumberOfColumns++, 90);

                    foreach (JobOperation oJobOperation in oOperationsPerPartSetToZero)
                    {
                        oSLBOMDocument.SetCellValue(iNumberOfRows, 1, oJobOperation.JobNum);
                        oSLBOMDocument.SetCellValue(iNumberOfRows, 2, oJobOperation.ParentPartNum);
                        oSLBOMDocument.SetCellValue(iNumberOfRows, 3, oJobOperation.ParentRevNum);
                        oSLBOMDocument.SetCellValue(iNumberOfRows, 4, oJobOperation.AssemblySeq.ToString());
                        oSLBOMDocument.SetCellValue(iNumberOfRows, 5, oJobOperation.OperationSeq.ToString());
                        oSLBOMDocument.SetCellValue(iNumberOfRows, 6, oJobOperation.OpCode);
                        oSLBOMDocument.SetCellValue(iNumberOfRows++, 7, oJobOperation.StdFormat);

                        bDataInReport = true;
                    }
                }

                // if std format is anything other than "OM" or "OH" then the OpsPerPart field should be zero
                List<JobOperation> oOperationsPerPartNotSetToZero = oJob.AllOperations.Where(oItem => (string.Compare(oItem.StdFormat, "OM", true) != 0) && (string.Compare(oItem.StdFormat, "OH", true) != 0) && (oItem.OperationsPerPart != 0)).ToList();
                if (oOperationsPerPartNotSetToZero.Count > 0)
                {
                    // sort these by job number
                    oOperationsPerPartNotSetToZero = oOperationsPerPartNotSetToZero.OrderBy(oItem => oItem.JobNum).ToList();
                    iNumberOfRows = 1;
                    iNumberOfColumns = 1;
                    if (bFirstWorksheet == true)
                    {
                        oSLBOMDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "Ops Per Part Not Zero");
                        bFirstWorksheet = false;
                    }
                    else
                    {
                        oSLBOMDocument.AddWorksheet("Ops Per Part Not Zero");
                    }
                    //set column header
                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "JobNum");
                    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Parent Part Num");
                    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Parent Part Rev Num");
                    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Assembly Seq");
                    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Opr Seq");
                    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Op Code");
                    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Std Format");
                    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Ops Per Part");
                    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "The Operation Is NOT Set To Operations Per Part -- THE OPERATIONS PER PART FIELD SHOULD BE ZERO");
                    oSLBOMDocument.SetCellStyle(iNumberOfRows++, iNumberOfColumns, oHighlightHeaderStyle);
                    oSLBOMDocument.SetColumnWidth(iNumberOfColumns++, 90);

                    foreach (JobOperation oJobOperation in oOperationsPerPartNotSetToZero)
                    {
                        oSLBOMDocument.SetCellValue(iNumberOfRows, 1, oJobOperation.JobNum);
                        oSLBOMDocument.SetCellValue(iNumberOfRows, 2, oJobOperation.ParentPartNum);
                        oSLBOMDocument.SetCellValue(iNumberOfRows, 3, oJobOperation.ParentRevNum);
                        oSLBOMDocument.SetCellValue(iNumberOfRows, 4, oJobOperation.AssemblySeq.ToString());
                        oSLBOMDocument.SetCellValue(iNumberOfRows, 5, oJobOperation.OperationSeq.ToString());
                        oSLBOMDocument.SetCellValue(iNumberOfRows, 6, oJobOperation.OpCode);
                        oSLBOMDocument.SetCellValue(iNumberOfRows, 7, oJobOperation.StdFormat);
                        oSLBOMDocument.SetCellValue(iNumberOfRows++, 8, oJobOperation.OperationsPerPart);

                        bDataInReport = true;
                    }
                }

                // std basis is wrong
                List<JobOperation> oOperationsStdBasisWrong = oJob.AllOperations.Where(oItem => (string.Compare(oItem.StdBasis, "E", true) != 0) && (oItem.Subcontract == false)).ToList();
                if (oOperationsStdBasisWrong.Count > 0)
                {
                    // sort these by job number
                    oOperationsStdBasisWrong = oOperationsStdBasisWrong.OrderBy(oItem => oItem.JobNum).ToList();
                    iNumberOfRows = 1;
                    iNumberOfColumns = 1;
                    if (bFirstWorksheet == true)
                    {
                        oSLBOMDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "Op Bad Std Basis");
                        bFirstWorksheet = false;
                    }
                    else
                    {
                        oSLBOMDocument.AddWorksheet("Op Bad Std Basis");
                    }
                    //set column header
                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "JobNum");
                    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Parent Part Num");
                    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Parent Part Rev Num");
                    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Assembly Seq");
                    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Opr Seq");
                    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Op Code");
                    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Std Basis");
                    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "The Operation Standard Basis Is Not Set To EACH -- SOP IS TO SET THE BASIS TO EACH NOT PER HUNDRED OR PER THOUSAND");
                    oSLBOMDocument.SetCellStyle(iNumberOfRows++, iNumberOfColumns, oHighlightHeaderStyle);
                    oSLBOMDocument.SetColumnWidth(iNumberOfColumns++, 90);

                    foreach (JobOperation oJobOperation in oOperationsStdBasisWrong)
                    {
                        oSLBOMDocument.SetCellValue(iNumberOfRows, 1, oJobOperation.JobNum);
                        oSLBOMDocument.SetCellValue(iNumberOfRows, 2, oJobOperation.ParentPartNum);
                        oSLBOMDocument.SetCellValue(iNumberOfRows, 3, oJobOperation.ParentRevNum);
                        oSLBOMDocument.SetCellValue(iNumberOfRows, 4, oJobOperation.AssemblySeq.ToString());
                        oSLBOMDocument.SetCellValue(iNumberOfRows, 5, oJobOperation.OperationSeq.ToString());
                        oSLBOMDocument.SetCellValue(iNumberOfRows, 6, oJobOperation.OpCode);
                        oSLBOMDocument.SetCellValue(iNumberOfRows++, 7, oJobOperation.StdBasis);

                        bDataInReport = true;
                    }
                }

                // labor entry method not correct
                List<JobOperation> oOperationsBadLaborEntryMethod = oJob.AllOperations.Where(oItem => (string.Compare(oItem.LaborEntryMethod, "T", true) != 0)).ToList();
                if (oOperationsBadLaborEntryMethod.Count > 0)
                {
                    // sort these by job number
                    oOperationsBadLaborEntryMethod = oOperationsBadLaborEntryMethod.OrderBy(oItem => oItem.JobNum).ToList();
                    iNumberOfRows = 1;
                    iNumberOfColumns = 1;
                    if (bFirstWorksheet == true)
                    {
                        oSLBOMDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "Op Bad Labor Entry Method");
                        bFirstWorksheet = false;
                    }
                    else
                    {
                        oSLBOMDocument.AddWorksheet("Op Bad Labor Entry Method");
                    }
                    //set column header
                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "JobNum");
                    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Parent Part Num");
                    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Parent Part Rev Num");
                    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Assembly Seq");
                    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Opr Seq");
                    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Op Code");
                    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Labor Entry");
                    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "The Operation Labor Entry Method Not Set Correctly -- SOP IS TO USE TIME AND QUANTITY FOR REPORTING");
                    oSLBOMDocument.SetCellStyle(iNumberOfRows++, iNumberOfColumns, oHighlightHeaderStyle);
                    oSLBOMDocument.SetColumnWidth(iNumberOfColumns++, 90);

                    foreach (JobOperation oJobOperation in oOperationsBadLaborEntryMethod)
                    {
                        oSLBOMDocument.SetCellValue(iNumberOfRows, 1, oJobOperation.JobNum);
                        oSLBOMDocument.SetCellValue(iNumberOfRows, 2, oJobOperation.ParentPartNum);
                        oSLBOMDocument.SetCellValue(iNumberOfRows, 3, oJobOperation.ParentRevNum);
                        oSLBOMDocument.SetCellValue(iNumberOfRows, 4, oJobOperation.AssemblySeq.ToString());
                        oSLBOMDocument.SetCellValue(iNumberOfRows, 5, oJobOperation.OperationSeq.ToString());
                        oSLBOMDocument.SetCellValue(iNumberOfRows, 6, oJobOperation.OpCode);
                        oSLBOMDocument.SetCellValue(iNumberOfRows++, 7, oJobOperation.LaborEntryMethod);

                        bDataInReport = true;
                    }
                }

                #region IGNORING THESE ISSUES
                // if this is a subcontract we should include the days out
                //////////List<JobOperation> oSubcontractOperationNoDaysOutSet = oJob.AllOperations.Where(oItem => (oItem.Subcontract == true) && (oItem.DaysOut == 0)).ToList();
                //////////if (oSubcontractOperationNoDaysOutSet.Count > 0)
                //////////{
                //////////    // sort these by job number
                //////////    oSubcontractOperationNoDaysOutSet = oSubcontractOperationNoDaysOutSet.OrderBy(oItem => oItem.JobNum).ToList();
                //////////    iNumberOfRows = 1;
                //////////    iNumberOfColumns = 1;
                //////////    if (bFirstWorksheet == true)
                //////////    {
                //////////        oSLBOMDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "Sub Op No Days Out Set");
                //////////        bFirstWorksheet = false;
                //////////    }
                //////////    else
                //////////    {
                //////////        oSLBOMDocument.AddWorksheet("Sub Op No Days Out Set");
                //////////    }
                //////////    //set column header
                //////////    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "JobNum");
                //////////    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                //////////    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Parent Part Num");
                //////////    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                //////////    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Parent Part Rev Num");
                //////////    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                //////////    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Assembly Seq");
                //////////    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                //////////    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Opr Seq");
                //////////    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                //////////    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Op Code");
                //////////    oSLBOMDocument.SetColumnWidth(iNumberOfRows++, iNumberOfColumns++, iStandardColumnWidth);

                //////////    foreach (JobOperation oJobOperation in oOperationsBadLaborEntryMethod)
                //////////    {
                //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, 1, oJobOperation.JobNum);
                //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, 2, oJobOperation.ParentPartNum);
                //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, 3, oJobOperation.ParentRevNum);
                //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, 4, oJobOperation.AssemblySeq.ToString());
                //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, 5, oJobOperation.OperationSeq.ToString());
                //////////        oSLBOMDocument.SetCellValue(iNumberOfRows++, 6, oJobOperation.OpCode);

                //////////        bDataInReport = true;
                //////////    }
                //////////}

                // sub op with no vendor set
                //////////List<JobOperation> oSubcontractOperationNoVendor = m_oJobOperations.Where(oItem => (oItem.Subcontract == true) && (oItem.VendorNum == 0)).ToList();
                //////////if (oSubcontractOperationNoVendor.Count > 0)
                //////////{
                //////////    // sort these by job number
                //////////    oSubcontractOperationNoVendor = oSubcontractOperationNoVendor.OrderBy(oItem => oItem.JobNum).ToList();
                //////////    iNumberOfRows = 1;
                //////////    iNumberOfColumns = 1;
                //////////    if (bFirstWorksheet == true)
                //////////    {
                //////////        oSLBOMDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "Sub Op No Vendor");
                //////////        bFirstWorksheet = false;
                //////////    }
                //////////    else
                //////////    {
                //////////        oSLBOMDocument.AddWorksheet("Sub Op No Vendor");
                //////////    }
                //////////    //set column header
                //////////    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "JobNum");
                //////////    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                //////////    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Parent Part Num");
                //////////    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                //////////    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Parent Part Rev Num");
                //////////    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                //////////    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Assembly Seq");
                //////////    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                //////////    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Opr Seq");
                //////////    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                //////////    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Op Code");
                //////////    oSLBOMDocument.SetColumnWidth(iNumberOfRows++, iNumberOfColumns++, iStandardColumnWidth);

                //////////    foreach (JobOperation oJobOperation in oSubcontractOperationNoVendor)
                //////////    {
                //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, 1, oJobOperation.JobNum);
                //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, 2, oJobOperation.ParentPartNum);
                //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, 3, oJobOperation.ParentRevNum);
                //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, 4, oJobOperation.AssemblySeq.ToString());
                //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, 5, oJobOperation.OperationSeq.ToString());
                //////////        oSLBOMDocument.SetCellValue(iNumberOfRows++, 6, oJobOperation.OpCode);

                //////////        bDataInReport = true;
                //////////    }
                //////////}
                #endregion

                // qty per should be set for subcontract
                List<JobOperation> oSubcontractOperationZeroQtyPer = oJob.AllOperations.Where(oItem => (oItem.Subcontract == true) && (oItem.QtyPer == 0)).ToList();
                if (oSubcontractOperationZeroQtyPer.Count > 0)
                {
                    // sort these by job number
                    oSubcontractOperationZeroQtyPer = oSubcontractOperationZeroQtyPer.OrderBy(oItem => oItem.JobNum).ToList();
                    iNumberOfRows = 1;
                    iNumberOfColumns = 1;
                    if (bFirstWorksheet == true)
                    {
                        oSLBOMDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "Sub Op Zero Qty Per");
                        bFirstWorksheet = false;
                    }
                    else
                    {
                        oSLBOMDocument.AddWorksheet("Sub Op Zero Qty Per");
                    }
                    //set column header
                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "JobNum");
                    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Parent Part Num");
                    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Parent Part Rev Num");
                    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Assembly Seq");
                    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Opr Seq");
                    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Op Code");
                    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "The Subcontract Qty Is Zero -- POC ISSUE PLEASE SET QTY ON SUBCONTRACT OPERATIONS");
                    oSLBOMDocument.SetCellStyle(iNumberOfRows++, iNumberOfColumns, oHighlightHeaderStyle);
                    oSLBOMDocument.SetColumnWidth(iNumberOfColumns++, 90);

                    foreach (JobOperation oJobOperation in oSubcontractOperationZeroQtyPer)
                    {
                        oSLBOMDocument.SetCellValue(iNumberOfRows, 1, oJobOperation.JobNum);
                        oSLBOMDocument.SetCellValue(iNumberOfRows, 2, oJobOperation.ParentPartNum);
                        oSLBOMDocument.SetCellValue(iNumberOfRows, 3, oJobOperation.ParentRevNum);
                        oSLBOMDocument.SetCellValue(iNumberOfRows, 4, oJobOperation.AssemblySeq.ToString());
                        oSLBOMDocument.SetCellValue(iNumberOfRows, 5, oJobOperation.OperationSeq.ToString());
                        oSLBOMDocument.SetCellValue(iNumberOfRows++, 6, oJobOperation.OpCode);

                        bDataInReport = true;
                    }
                }

                #region IGNORING THESE ISSUES
                // burden does not equals labor
                //////////List<JobOperation> oOperationWithBurdenNotEqualToLabor = oJob.AllOperations.Where(oItem => (oItem.BurdenEqualsLabor == false)).ToList();
                //////////if (oOperationWithBurdenNotEqualToLabor.Count > 0)
                //////////{
                //////////    // sort these by job number
                //////////    oOperationWithBurdenNotEqualToLabor = oOperationWithBurdenNotEqualToLabor.OrderBy(oItem => oItem.JobNum).ToList();
                //////////    iNumberOfRows = 1;
                //////////    iNumberOfColumns = 1;
                //////////    if (bFirstWorksheet == true)
                //////////    {
                //////////        oSLBOMDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "Burden Not Eq Labor");
                //////////        bFirstWorksheet = false;
                //////////    }
                //////////    else
                //////////    {
                //////////        oSLBOMDocument.AddWorksheet("Burden Not Eq Labor");
                //////////    }
                //////////    //set column header
                //////////    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "JobNum");
                //////////    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                //////////    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Parent Part Num");
                //////////    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                //////////    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Parent Part Rev Num");
                //////////    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                //////////    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Assembly Seq");
                //////////    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                //////////    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Opr Seq");
                //////////    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                //////////    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Op Code");
                //////////    oSLBOMDocument.SetColumnWidth(iNumberOfRows++, iNumberOfColumns++, iStandardColumnWidth);

                //////////    foreach (JobOperation oJobOperation in oOperationWithBurdenNotEqualToLabor)
                //////////    {
                //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, 1, oJobOperation.JobNum);
                //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, 2, oJobOperation.ParentPartNum);
                //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, 3, oJobOperation.ParentRevNum);
                //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, 4, oJobOperation.AssemblySeq.ToString());
                //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, 5, oJobOperation.OperationSeq.ToString());
                //////////        oSLBOMDocument.SetCellValue(iNumberOfRows++, 6, oJobOperation.OpCode);

                //////////        bDataInReport = true;
                //////////    }
                //////////}
                #endregion

                // use estimates
                List<JobOperation> oOperationUseEstimatesSet = oJob.AllOperations.Where(oItem => (oItem.UseEstimates == true)).ToList();
                if (oOperationUseEstimatesSet.Count > 0)
                {
                    // sort these by job number
                    oOperationUseEstimatesSet = oOperationUseEstimatesSet.OrderBy(oItem => oItem.JobNum).ToList();
                    iNumberOfRows = 1;
                    iNumberOfColumns = 1;
                    if (bFirstWorksheet == true)
                    {
                        oSLBOMDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "Use Estimates Set");
                        bFirstWorksheet = false;
                    }
                    else
                    {
                        oSLBOMDocument.AddWorksheet("Use Estimates Set");
                    }
                    //set column header
                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "JobNum");
                    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Parent Part Num");
                    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Parent Part Rev Num");
                    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Assembly Seq");
                    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Opr Seq");
                    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Op Code");
                    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "The Operation Has Use Estimates Set -- SOP IS TO NOT USE THIS FEATURE");
                    oSLBOMDocument.SetCellStyle(iNumberOfRows++, iNumberOfColumns, oHighlightHeaderStyle);
                    oSLBOMDocument.SetColumnWidth(iNumberOfColumns++, 90);

                    foreach (JobOperation oJobOperation in oOperationUseEstimatesSet)
                    {
                        oSLBOMDocument.SetCellValue(iNumberOfRows, 1, oJobOperation.JobNum);
                        oSLBOMDocument.SetCellValue(iNumberOfRows, 2, oJobOperation.ParentPartNum);
                        oSLBOMDocument.SetCellValue(iNumberOfRows, 3, oJobOperation.ParentRevNum);
                        oSLBOMDocument.SetCellValue(iNumberOfRows, 4, oJobOperation.AssemblySeq.ToString());
                        oSLBOMDocument.SetCellValue(iNumberOfRows, 5, oJobOperation.OperationSeq.ToString());
                        oSLBOMDocument.SetCellValue(iNumberOfRows++, 6, oJobOperation.OpCode);

                        bDataInReport = true;
                    }
                }

                // split operations
                List<JobOperation> oSplitOperationsSet = oJob.AllOperations.Where(oItem => (oItem.SplitOperations == true)).ToList();
                if (oSplitOperationsSet.Count > 0)
                {
                    // sort these by job number
                    oSplitOperationsSet = oSplitOperationsSet.OrderBy(oItem => oItem.JobNum).ToList();
                    iNumberOfRows = 1;
                    iNumberOfColumns = 1;
                    if (bFirstWorksheet == true)
                    {
                        oSLBOMDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "Split Operations Set");
                        bFirstWorksheet = false;
                    }
                    else
                    {
                        oSLBOMDocument.AddWorksheet("Split Operations Set");
                    }
                    //set column header
                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "JobNum");
                    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Parent Part Num");
                    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Parent Part Rev Num");
                    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Assembly Seq");
                    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Opr Seq");
                    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Op Code");
                    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "The Operation Permits Spliting -- SOP IS TO NOT ALLOW THIS FEATURE");
                    oSLBOMDocument.SetCellStyle(iNumberOfRows++, iNumberOfColumns, oHighlightHeaderStyle);
                    oSLBOMDocument.SetColumnWidth(iNumberOfColumns++, 90);

                    foreach (JobOperation oJobOperation in oSplitOperationsSet)
                    {
                        oSLBOMDocument.SetCellValue(iNumberOfRows, 1, oJobOperation.JobNum);
                        oSLBOMDocument.SetCellValue(iNumberOfRows, 2, oJobOperation.ParentPartNum);
                        oSLBOMDocument.SetCellValue(iNumberOfRows, 3, oJobOperation.ParentRevNum);
                        oSLBOMDocument.SetCellValue(iNumberOfRows, 4, oJobOperation.AssemblySeq.ToString());
                        oSLBOMDocument.SetCellValue(iNumberOfRows, 5, oJobOperation.OperationSeq.ToString());
                        oSLBOMDocument.SetCellValue(iNumberOfRows++, 6, oJobOperation.OpCode);

                        bDataInReport = true;
                    }
                }

                // split burden
                List<JobOperation> oSplitBurdenSet = oJob.AllOperations.Where(oItem => (oItem.SplitBurden == true)).ToList();
                if (oSplitOperationsSet.Count > 0)
                {
                    // sort these by job number
                    oSplitBurdenSet = oSplitBurdenSet.OrderBy(oItem => oItem.JobNum).ToList();
                    iNumberOfRows = 1;
                    iNumberOfColumns = 1;
                    if (bFirstWorksheet == true)
                    {
                        oSLBOMDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "Split Burden Set");
                        bFirstWorksheet = false;
                    }
                    else
                    {
                        oSLBOMDocument.AddWorksheet("Split Burden Set");
                    }
                    //set column header
                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "JobNum");
                    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Parent Part Num");
                    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Parent Part Rev Num");
                    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Assembly Seq");
                    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Opr Seq");
                    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Op Code");
                    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "The Burden Rate Is Set To Split Across Resources -- SOP IS TO NOT ALLOW SPLITTING OF BURDEN");
                    oSLBOMDocument.SetCellStyle(iNumberOfRows++, iNumberOfColumns, oHighlightHeaderStyle);
                    oSLBOMDocument.SetColumnWidth(iNumberOfColumns++, 90);

                    foreach (JobOperation oJobOperation in oSplitBurdenSet)
                    {
                        oSLBOMDocument.SetCellValue(iNumberOfRows, 1, oJobOperation.JobNum);
                        oSLBOMDocument.SetCellValue(iNumberOfRows, 2, oJobOperation.ParentPartNum);
                        oSLBOMDocument.SetCellValue(iNumberOfRows, 3, oJobOperation.ParentRevNum);
                        oSLBOMDocument.SetCellValue(iNumberOfRows, 4, oJobOperation.AssemblySeq.ToString());
                        oSLBOMDocument.SetCellValue(iNumberOfRows, 5, oJobOperation.OperationSeq.ToString());
                        oSLBOMDocument.SetCellValue(iNumberOfRows++, 6, oJobOperation.OpCode);

                        bDataInReport = true;
                    }
                }

                // subcontract estimate is zero
                List<JobOperation> oSubOperationZeroEst = oJob.AllOperations.Where(oItem => (oItem.Subcontract == true) && (oItem.EstUnitCost == 0)).ToList();
                if (oSubOperationZeroEst.Count > 0)
                {
                    // sort these by job number
                    oSubOperationZeroEst = oSubOperationZeroEst.OrderBy(oItem => oItem.JobNum).ToList();
                    iNumberOfRows = 1;
                    iNumberOfColumns = 1;
                    if (bFirstWorksheet == true)
                    {
                        oSLBOMDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "Sub Op Zero Est");
                        bFirstWorksheet = false;
                    }
                    else
                    {
                        oSLBOMDocument.AddWorksheet("Sub Op Zero Est");
                    }
                    //set column header
                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "JobNum");
                    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Parent Part Num");
                    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Parent Part Rev Num");
                    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Assembly Seq");
                    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Opr Seq");
                    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Op Code");
                    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "The Subcontract Cost Is Zero -- POC ISSUE PLEASE SET COST");
                    oSLBOMDocument.SetCellStyle(iNumberOfRows++, iNumberOfColumns, oHighlightHeaderStyle);
                    oSLBOMDocument.SetColumnWidth(iNumberOfColumns++, 90);

                    foreach (JobOperation oJobOperation in oSubOperationZeroEst)
                    {
                        oSLBOMDocument.SetCellValue(iNumberOfRows, 1, oJobOperation.JobNum);
                        oSLBOMDocument.SetCellValue(iNumberOfRows, 2, oJobOperation.ParentPartNum);
                        oSLBOMDocument.SetCellValue(iNumberOfRows, 3, oJobOperation.ParentRevNum);
                        oSLBOMDocument.SetCellValue(iNumberOfRows, 4, oJobOperation.AssemblySeq.ToString());
                        oSLBOMDocument.SetCellValue(iNumberOfRows, 5, oJobOperation.OperationSeq.ToString());
                        oSLBOMDocument.SetCellValue(iNumberOfRows++, 6, oJobOperation.OpCode);

                        bDataInReport = true;
                    }
                }
                #endregion

                #region Operation Costs

                #region IGNORING THESE ISSUES
                //////////// unexpected hours -- if we are over the estimated amount or if we are less then the estimated amount and the op is completed
                //////////List<JobOpsEstVsActualCosts> oJobOperationHoursDiffers = oJob.AllOperationCosts.Where(oItem => (oItem.ActOprHours > oItem.EstOprHours) || ((oItem.ActOprHours < oItem.EstOprHours) && (oItem.OpComplete == true))).ToList();
                //////////if (oJobOperationHoursDiffers.Count > 0)
                //////////{
                //////////    // sort these by job number
                //////////    oJobOperationHoursDiffers = oJobOperationHoursDiffers.OrderBy(oItem => oItem.JobNum).ToList();
                //////////    iNumberOfRows = 1;
                //////////    iNumberOfColumns = 1;
                //////////    if (bFirstWorksheet == true)
                //////////    {
                //////////        oSLBOMDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "Op Hours");
                //////////        bFirstWorksheet = false;
                //////////    }
                //////////    else
                //////////    {
                //////////        oSLBOMDocument.AddWorksheet("Op Hours");
                //////////    }
                //////////    //set column header
                //////////    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "JobNum");
                //////////    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                //////////    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Parent Part Num");
                //////////    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                //////////    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Parent Part Rev Num");
                //////////    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                //////////    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Assembly Seq");
                //////////    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                //////////    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Opr Seq");
                //////////    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                //////////    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Opr Code");
                //////////    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                //////////    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Est Hours");
                //////////    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                //////////    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Act Hours");
                //////////    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                //////////    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Variance Qty");
                //////////    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                //////////    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Variance %");
                //////////    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                //////////    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "The Operation Hours Differs From Estimate -- POC ISSUE PLEASE REVIEW");
                //////////    oSLBOMDocument.SetCellStyle(iNumberOfRows++, iNumberOfColumns, oHighlightHeaderStyle);
                //////////    oSLBOMDocument.SetColumnWidth(iNumberOfColumns++, 90);

                //////////    foreach (JobOpsEstVsActualCosts oJobOperation in oJobOperationHoursDiffers)
                //////////    {
                //////////        decimal dDelta = oJobOperation.ActOprHours - oJobOperation.EstOprHours;
                //////////        decimal dPercentageError = 0M;
                //////////        decimal dMarginToDate = 0M;

                //////////        if (oJobOperation.EstOprHours != 0)
                //////////        {
                //////////            dPercentageError = dDelta / oJobOperation.EstOprHours * 100.0M;
                //////////        }
                //////////        else
                //////////        {
                //////////            dPercentageError = 100.0M;
                //////////        }
                //////////        HSJob oTmpJob = m_oAllJobs.FirstOrDefault(oItem => string.Compare(oJobOperation.JobNum, oItem.JobNum, true) == 0);
                //////////        if (oTmpJob != null)
                //////////        {
                //////////            dMarginToDate = oTmpJob.MarginToDate * 100.0M;
                //////////        }

                //////////        bool bExceedsThreholds = false;

                //////////        // HERE WE JUST CHECK % ERROR AS WE HAVE NO $$$$
                //////////        if ((Math.Abs(dPercentageError) > m_dJobEstimatePercentError) && (dMarginToDate < m_dJobMarginThreshold))
                //////////        {
                //////////            // check to see if it exceeds both thresholds
                //////////            bExceedsThreholds = true;
                //////////        }

                //////////        if (bExceedsThreholds == true)
                //////////        {
                //////////            oSLBOMDocument.SetCellValue(iNumberOfRows, 1, oJobOperation.JobNum);
                //////////            oSLBOMDocument.SetCellValue(iNumberOfRows, 2, oJobOperation.PartNum);
                //////////            oSLBOMDocument.SetCellValue(iNumberOfRows, 3, oJobOperation.PartRevNum);
                //////////            oSLBOMDocument.SetCellValue(iNumberOfRows, 4, oJobOperation.AssemblySeq.ToString());
                //////////            oSLBOMDocument.SetCellValue(iNumberOfRows, 5, oJobOperation.OprSeq.ToString());
                //////////            oSLBOMDocument.SetCellValue(iNumberOfRows, 6, oJobOperation.OpCode);
                //////////            oSLBOMDocument.SetCellValue(iNumberOfRows, 7, oJobOperation.EstOprHours);
                //////////            oSLBOMDocument.SetCellStyle(iNumberOfRows, 7, oDecimalStyle);
                //////////            oSLBOMDocument.SetCellValue(iNumberOfRows, 8, oJobOperation.ActOprHours);
                //////////            oSLBOMDocument.SetCellStyle(iNumberOfRows, 8, oDecimalStyle);

                //////////            oSLBOMDocument.SetCellValue(iNumberOfRows, 9, dDelta);
                //////////            oSLBOMDocument.SetCellStyle(iNumberOfRows, 9, oDecimalStyle);

                //////////            oSLBOMDocument.SetCellValue(iNumberOfRows, 10, dPercentageError);
                //////////            oSLBOMDocument.SetCellStyle(iNumberOfRows++, 10, oDecimalStyle);
                //////////            bDataInReport = true;
                //////////        }
                //////////    }
                //////////}

                //////////// est labor cost differs from actual
                //////////List<JobOpsEstVsActualCosts> oJobOperationLaborDiffers = oJob.AllOperationCosts.Where(oItem => (oItem.ActLaborCost > oItem.EstLaborCost) || ((oItem.ActLaborCost < oItem.EstLaborCost) && (oItem.OpComplete == true))).ToList();
                //////////if (oJobOperationLaborDiffers.Count > 0)
                //////////{
                //////////    // sort these by job number
                //////////    oJobOperationLaborDiffers = oJobOperationLaborDiffers.OrderBy(oItem => oItem.JobNum).ToList();
                //////////    iNumberOfRows = 1;
                //////////    iNumberOfColumns = 1;
                //////////    if (bFirstWorksheet == true)
                //////////    {
                //////////        oSLBOMDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "Op Labor");
                //////////        bFirstWorksheet = false;
                //////////    }
                //////////    else
                //////////    {
                //////////        oSLBOMDocument.AddWorksheet("Op Labor");
                //////////    }
                //////////    //set column header
                //////////    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "JobNum");
                //////////    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                //////////    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Parent Part Num");
                //////////    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                //////////    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Parent Part Rev Num");
                //////////    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                //////////    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Assembly Seq");
                //////////    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                //////////    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Opr Seq");
                //////////    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                //////////    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Opr Code");
                //////////    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                //////////    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Est Labor");
                //////////    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                //////////    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Act Labor");
                //////////    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                //////////    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Variance $");
                //////////    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                //////////    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Variance %");
                //////////    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                //////////    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "The Operation Cost Differs From Estimate -- POC ISSUE PLEASE REVIEW");
                //////////    oSLBOMDocument.SetCellStyle(iNumberOfRows++, iNumberOfColumns, oHighlightHeaderStyle);
                //////////    oSLBOMDocument.SetColumnWidth(iNumberOfColumns++, 90);

                //////////    foreach (JobOpsEstVsActualCosts oJobOperation in oJobOperationLaborDiffers)
                //////////    {
                //////////        decimal dDelta = oJobOperation.ActLaborCost - oJobOperation.EstLaborCost;
                //////////        decimal dPercentageError = 0M;
                //////////        decimal dMarginToDate = 0M;
                //////////        if (oJobOperation.EstLaborCost != 0)
                //////////        {
                //////////            dPercentageError = dDelta / oJobOperation.EstLaborCost * 100.0M;
                //////////        }
                //////////        else
                //////////        {
                //////////            dPercentageError = 100.0M;
                //////////        }
                //////////        HSJob oTmpJob = m_oAllJobs.FirstOrDefault(oItem => string.Compare(oJobOperation.JobNum, oItem.JobNum, true) == 0);
                //////////        if (oTmpJob != null)
                //////////        {
                //////////            dMarginToDate = oTmpJob.MarginToDate * 100.0M;
                //////////        }

                //////////        bool bExceedsThreholds = false;

                //////////        if ((Math.Abs(dDelta) > m_dJobEstimateAbsoluteError) && (Math.Abs(dPercentageError) > m_dJobEstimatePercentError) && (dMarginToDate < m_dJobMarginThreshold))
                //////////        {
                //////////            // check to see if it exceeds both thresholds
                //////////            bExceedsThreholds = true;
                //////////        }
                //////////        if (m_bForceJobAnalysis == true)
                //////////        {
                //////////            // force it to report
                //////////            bExceedsThreholds = true;
                //////////        }

                //////////        if (bExceedsThreholds == true)
                //////////        {
                //////////            oSLBOMDocument.SetCellValue(iNumberOfRows, 1, oJobOperation.JobNum);
                //////////            oSLBOMDocument.SetCellValue(iNumberOfRows, 2, oJobOperation.PartNum);
                //////////            oSLBOMDocument.SetCellValue(iNumberOfRows, 3, oJobOperation.PartRevNum);
                //////////            oSLBOMDocument.SetCellValue(iNumberOfRows, 4, oJobOperation.AssemblySeq.ToString());
                //////////            oSLBOMDocument.SetCellValue(iNumberOfRows, 5, oJobOperation.OprSeq.ToString());
                //////////            oSLBOMDocument.SetCellValue(iNumberOfRows, 6, oJobOperation.OpCode);
                //////////            oSLBOMDocument.SetCellValue(iNumberOfRows, 7, oJobOperation.EstLaborCost);
                //////////            oSLBOMDocument.SetCellStyle(iNumberOfRows, 7, oCurrencyStyle);
                //////////            oSLBOMDocument.SetCellValue(iNumberOfRows, 8, oJobOperation.ActLaborCost);
                //////////            oSLBOMDocument.SetCellStyle(iNumberOfRows, 8, oCurrencyStyle);

                //////////            oSLBOMDocument.SetCellValue(iNumberOfRows, 9, dDelta);
                //////////            oSLBOMDocument.SetCellStyle(iNumberOfRows, 9, oCurrencyStyle);

                //////////            oSLBOMDocument.SetCellValue(iNumberOfRows, 10, dPercentageError);
                //////////            oSLBOMDocument.SetCellStyle(iNumberOfRows++, 10, oDecimalStyle);
                //////////            bDataInReport = true;
                //////////        }
                //////////    }
                //////////}

                //////////// est burden cost differs from actual
                //////////List<JobOpsEstVsActualCosts> oJobOperationBurdenDiffers = oJob.AllOperationCosts.Where(oItem => (oItem.ActBurdenCost > oItem.EstBurCost) || ((oItem.ActBurdenCost < oItem.EstBurCost) && (oItem.OpComplete == true))).ToList();
                //////////if (oJobOperationBurdenDiffers.Count > 0)
                //////////{
                //////////    // sort these by job number
                //////////    oJobOperationBurdenDiffers = oJobOperationBurdenDiffers.OrderBy(oItem => oItem.JobNum).ToList();
                //////////    iNumberOfRows = 1;
                //////////    iNumberOfColumns = 1;
                //////////    if (bFirstWorksheet == true)
                //////////    {
                //////////        oSLBOMDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "Op Burden");
                //////////        bFirstWorksheet = false;
                //////////    }
                //////////    else
                //////////    {
                //////////        oSLBOMDocument.AddWorksheet("Op Burden");
                //////////    }
                //////////    //set column header
                //////////    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "JobNum");
                //////////    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                //////////    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Parent Part Num");
                //////////    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                //////////    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Parent Part Rev Num");
                //////////    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                //////////    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Assembly Seq");
                //////////    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                //////////    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Opr Seq");
                //////////    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                //////////    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Opr Code");
                //////////    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                //////////    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Est Burden");
                //////////    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                //////////    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Act Burden");
                //////////    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                //////////    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Variance $");
                //////////    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                //////////    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Variance %");
                //////////    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                //////////    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "The Operation Burden Differs From Estimate -- POC ISSUE PLEASE REVIEW");
                //////////    oSLBOMDocument.SetCellStyle(iNumberOfRows++, iNumberOfColumns, oHighlightHeaderStyle);
                //////////    oSLBOMDocument.SetColumnWidth(iNumberOfColumns++, 90);

                //////////    foreach (JobOpsEstVsActualCosts oJobOperation in oJobOperationBurdenDiffers)
                //////////    {
                //////////        decimal dDelta = oJobOperation.ActBurdenCost - oJobOperation.EstBurCost;
                //////////        decimal dPercentageError = 0M;
                //////////        decimal dMarginToDate = 0M;
                //////////        if (oJobOperation.EstBurCost != 0)
                //////////        {
                //////////            dPercentageError = dDelta / oJobOperation.EstBurCost * 100.0M;
                //////////        }
                //////////        else
                //////////        {
                //////////            dPercentageError = 100.0M;
                //////////        }
                //////////        HSJob oTmpJob = m_oAllJobs.FirstOrDefault(oItem => string.Compare(oJobOperation.JobNum, oItem.JobNum, true) == 0);
                //////////        if (oTmpJob != null)
                //////////        {
                //////////            dMarginToDate = oTmpJob.MarginToDate * 100.0M;
                //////////        }

                //////////        bool bExceedsThreholds = false;
                //////////        if ((Math.Abs(dDelta) > m_dJobEstimateAbsoluteError) && (Math.Abs(dPercentageError) > m_dJobEstimatePercentError) && (dMarginToDate < m_dJobMarginThreshold))
                //////////        {
                //////////            // check to see if it exceeds both thresholds
                //////////            bExceedsThreholds = true;
                //////////        }
                //////////        if (m_bForceJobAnalysis == true)
                //////////        {
                //////////            // force it to report
                //////////            bExceedsThreholds = true;
                //////////        }

                //////////        if (bExceedsThreholds == true)
                //////////        {
                //////////            oSLBOMDocument.SetCellValue(iNumberOfRows, 1, oJobOperation.JobNum);
                //////////            oSLBOMDocument.SetCellValue(iNumberOfRows, 2, oJobOperation.PartNum);
                //////////            oSLBOMDocument.SetCellValue(iNumberOfRows, 3, oJobOperation.PartRevNum);
                //////////            oSLBOMDocument.SetCellValue(iNumberOfRows, 4, oJobOperation.AssemblySeq.ToString());
                //////////            oSLBOMDocument.SetCellValue(iNumberOfRows, 5, oJobOperation.OprSeq.ToString());
                //////////            oSLBOMDocument.SetCellValue(iNumberOfRows, 6, oJobOperation.OpCode);
                //////////            oSLBOMDocument.SetCellValue(iNumberOfRows, 7, oJobOperation.EstBurCost);
                //////////            oSLBOMDocument.SetCellStyle(iNumberOfRows, 7, oCurrencyStyle);
                //////////            oSLBOMDocument.SetCellValue(iNumberOfRows, 8, oJobOperation.ActBurdenCost);
                //////////            oSLBOMDocument.SetCellStyle(iNumberOfRows, 8, oCurrencyStyle);

                //////////            oSLBOMDocument.SetCellValue(iNumberOfRows, 9, dDelta);
                //////////            oSLBOMDocument.SetCellStyle(iNumberOfRows, 9, oCurrencyStyle);

                //////////            oSLBOMDocument.SetCellValue(iNumberOfRows, 10, dPercentageError);
                //////////            oSLBOMDocument.SetCellStyle(iNumberOfRows++, 10, oDecimalStyle);
                //////////            bDataInReport = true;
                //////////        }
                //////////    }
                //////////}

                //////////// est sub cost differs from actual
                //////////List<JobOpsEstVsActualCosts> oJobOperationSubDiffers = oJob.AllOperationCosts.Where(oItem => (oItem.ActSubCost > oItem.EstSubCost) || ((oItem.ActSubCost < oItem.EstSubCost) && (oItem.OpComplete == true))).ToList();
                //////////if (oJobOperationSubDiffers.Count > 0)
                //////////{
                //////////    // sort these by job number
                //////////    oJobOperationSubDiffers = oJobOperationSubDiffers.OrderBy(oItem => oItem.JobNum).ToList();
                //////////    iNumberOfRows = 1;
                //////////    iNumberOfColumns = 1;
                //////////    if (bFirstWorksheet == true)
                //////////    {
                //////////        oSLBOMDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "Subcontract");
                //////////        bFirstWorksheet = false;
                //////////    }
                //////////    else
                //////////    {
                //////////        oSLBOMDocument.AddWorksheet("Subcontract");
                //////////    }
                //////////    //set column header
                //////////    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "JobNum");
                //////////    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                //////////    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Parent Part Num");
                //////////    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                //////////    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Parent Part Rev Num");
                //////////    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                //////////    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Assembly Seq");
                //////////    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                //////////    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Opr Seq");
                //////////    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                //////////    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Opr Code");
                //////////    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                //////////    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Est Sub");
                //////////    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                //////////    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Act Sub");
                //////////    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                //////////    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Variance $");
                //////////    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                //////////    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Variance %");
                //////////    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                //////////    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "The Actual Subcontract Cost Differs From Estimate -- POC ISSUE PLEASE REVIEW");
                //////////    oSLBOMDocument.SetCellStyle(iNumberOfRows++, iNumberOfColumns, oHighlightHeaderStyle);
                //////////    oSLBOMDocument.SetColumnWidth(iNumberOfColumns++, 90);

                //////////    foreach (JobOpsEstVsActualCosts oJobOperation in oJobOperationSubDiffers)
                //////////    {
                //////////        decimal dDelta = oJobOperation.ActSubCost - oJobOperation.EstSubCost;
                //////////        decimal dPercentageError = 0M;
                //////////        decimal dMarginToDate = 0M;
                //////////        if (oJobOperation.EstSubCost != 0)
                //////////        {
                //////////            dPercentageError = dDelta / oJobOperation.EstSubCost * 100.0M;
                //////////        }
                //////////        else
                //////////        {
                //////////            dPercentageError = 100.0M;
                //////////        }
                //////////        HSJob oTmpJob = m_oAllJobs.FirstOrDefault(oItem => string.Compare(oJobOperation.JobNum, oItem.JobNum, true) == 0);
                //////////        if (oTmpJob != null)
                //////////        {
                //////////            dMarginToDate = oTmpJob.MarginToDate * 100.0M;
                //////////        }

                //////////        bool bExceedsThreholds = false;
                //////////        if ((Math.Abs(dDelta) > m_dJobEstimateAbsoluteError) && (Math.Abs(dPercentageError) > m_dJobEstimatePercentError) && (dMarginToDate < m_dJobMarginThreshold))
                //////////        {
                //////////            // check to see if it exceeds both thresholds
                //////////            bExceedsThreholds = true;
                //////////        }
                //////////        if (m_bForceJobAnalysis == true)
                //////////        {
                //////////            // force it to report
                //////////            bExceedsThreholds = true;
                //////////        }

                //////////        if (bExceedsThreholds == true)
                //////////        {
                //////////            oSLBOMDocument.SetCellValue(iNumberOfRows, 1, oJobOperation.JobNum);
                //////////            oSLBOMDocument.SetCellValue(iNumberOfRows, 2, oJobOperation.PartNum);
                //////////            oSLBOMDocument.SetCellValue(iNumberOfRows, 3, oJobOperation.PartRevNum);
                //////////            oSLBOMDocument.SetCellValue(iNumberOfRows, 4, oJobOperation.AssemblySeq.ToString());
                //////////            oSLBOMDocument.SetCellValue(iNumberOfRows, 5, oJobOperation.OprSeq.ToString());
                //////////            oSLBOMDocument.SetCellValue(iNumberOfRows, 6, oJobOperation.OpCode);
                //////////            oSLBOMDocument.SetCellValue(iNumberOfRows, 7, oJobOperation.EstSubCost);
                //////////            oSLBOMDocument.SetCellStyle(iNumberOfRows, 7, oCurrencyStyle);
                //////////            oSLBOMDocument.SetCellValue(iNumberOfRows, 8, oJobOperation.ActSubCost);
                //////////            oSLBOMDocument.SetCellStyle(iNumberOfRows, 8, oCurrencyStyle);

                //////////            oSLBOMDocument.SetCellValue(iNumberOfRows, 9, dDelta);
                //////////            oSLBOMDocument.SetCellStyle(iNumberOfRows, 9, oCurrencyStyle);

                //////////            oSLBOMDocument.SetCellValue(iNumberOfRows, 10, dPercentageError);
                //////////            oSLBOMDocument.SetCellStyle(iNumberOfRows++, 10, oDecimalStyle);
                //////////            bDataInReport = true;
                //////////        }
                //////////    }
                //////////}

                //////////// total op cost differs from actual
                //////////List<JobOpsEstVsActualCosts> oJobOperationTotalDiffers = oJob.AllOperationCosts.Where(oItem => (oItem.ActCost > oItem.EstCost) || ((oItem.ActCost < oItem.EstCost) && (oItem.OpComplete == true))).ToList();
                //////////if (oJobOperationTotalDiffers.Count > 0)
                //////////{
                //////////    // sort these by job number
                //////////    oJobOperationTotalDiffers = oJobOperationTotalDiffers.OrderBy(oItem => oItem.JobNum).ToList();
                //////////    iNumberOfRows = 1;
                //////////    iNumberOfColumns = 1;
                //////////    if (bFirstWorksheet == true)
                //////////    {
                //////////        oSLBOMDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "Op Total");
                //////////        bFirstWorksheet = false;
                //////////    }
                //////////    else
                //////////    {
                //////////        oSLBOMDocument.AddWorksheet("Op Total");
                //////////    }
                //////////    //set column header
                //////////    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "JobNum");
                //////////    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                //////////    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Parent Part Num");
                //////////    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                //////////    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Parent Part Rev Num");
                //////////    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                //////////    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Assembly Seq");
                //////////    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                //////////    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Opr Seq");
                //////////    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                //////////    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Opr Code");
                //////////    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                //////////    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Est Total");
                //////////    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                //////////    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Act Total");
                //////////    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                //////////    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Variance $");
                //////////    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                //////////    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Variance %");
                //////////    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                //////////    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "The Total Operation Cost Differs From Estimate -- POC ISSUE PLEASE REVIEW");
                //////////    oSLBOMDocument.SetCellStyle(iNumberOfRows++, iNumberOfColumns, oHighlightHeaderStyle);
                //////////    oSLBOMDocument.SetColumnWidth(iNumberOfColumns++, 90);

                //////////    foreach (JobOpsEstVsActualCosts oJobOperation in oJobOperationTotalDiffers)
                //////////    {
                //////////        decimal dDelta = oJobOperation.ActCost - oJobOperation.EstCost;
                //////////        decimal dPercentageError = 0M;
                //////////        decimal dMarginToDate = 0M;
                //////////        if (oJobOperation.EstCost != 0)
                //////////        {
                //////////            dPercentageError = dDelta / oJobOperation.EstCost * 100.0M;
                //////////        }
                //////////        else
                //////////        {
                //////////            dPercentageError = 100.0M;
                //////////        }
                //////////        HSJob oTmpJob = m_oAllJobs.FirstOrDefault(oItem => string.Compare(oJobOperation.JobNum, oItem.JobNum, true) == 0);
                //////////        if (oTmpJob != null)
                //////////        {
                //////////            dMarginToDate = oTmpJob.MarginToDate * 100.0M;
                //////////        }

                //////////        bool bExceedsThreholds = false;
                //////////        if ((Math.Abs(dDelta) > m_dJobEstimateAbsoluteError) && (Math.Abs(dPercentageError) > m_dJobEstimatePercentError) && (dMarginToDate < m_dJobMarginThreshold))
                //////////        {
                //////////            // check to see if it exceeds both thresholds
                //////////            bExceedsThreholds = true;
                //////////        }
                //////////        if (m_bForceJobAnalysis == true)
                //////////        {
                //////////            // force it to report
                //////////            bExceedsThreholds = true;
                //////////        }

                //////////        if (bExceedsThreholds == true)
                //////////        {
                //////////            oSLBOMDocument.SetCellValue(iNumberOfRows, 1, oJobOperation.JobNum);
                //////////            oSLBOMDocument.SetCellValue(iNumberOfRows, 2, oJobOperation.PartNum);
                //////////            oSLBOMDocument.SetCellValue(iNumberOfRows, 3, oJobOperation.PartRevNum);
                //////////            oSLBOMDocument.SetCellValue(iNumberOfRows, 4, oJobOperation.AssemblySeq.ToString());
                //////////            oSLBOMDocument.SetCellValue(iNumberOfRows, 5, oJobOperation.OprSeq.ToString());
                //////////            oSLBOMDocument.SetCellValue(iNumberOfRows, 6, oJobOperation.OpCode);
                //////////            oSLBOMDocument.SetCellValue(iNumberOfRows, 7, oJobOperation.EstCost);
                //////////            oSLBOMDocument.SetCellStyle(iNumberOfRows, 7, oCurrencyStyle);
                //////////            oSLBOMDocument.SetCellValue(iNumberOfRows, 8, oJobOperation.ActCost);
                //////////            oSLBOMDocument.SetCellStyle(iNumberOfRows, 8, oCurrencyStyle);

                //////////            oSLBOMDocument.SetCellValue(iNumberOfRows, 9, dDelta);
                //////////            oSLBOMDocument.SetCellStyle(iNumberOfRows, 9, oCurrencyStyle);

                //////////            oSLBOMDocument.SetCellValue(iNumberOfRows, 10, dPercentageError);
                //////////            oSLBOMDocument.SetCellStyle(iNumberOfRows++, 10, oDecimalStyle);
                //////////            bDataInReport = true;
                //////////        }
                //////////    }
                //////////}
                #endregion

                #endregion

                #region Charts

                //////////SLChart chart;
                //////////double fChartHeight;
                //////////double fChartWidth;

                //////////// walk through the list of assemblies and operations in order
                //////////// when looking at the materials for an assembly, if the material
                //////////// is not tied to an operation it should appear on operation ZERO
                //////////List<HSAssembly> oJobAssemblies = oJob.GetAssembliesInOrder();

                //////////#region Costs By Area
                //////////decimal dEstMtlTotals = oJob.EstMaterialCost;
                //////////decimal dActMtlTotals = oJob.ActMaterialCost;
                //////////decimal dEstOprTotals = oJob.EstLaborCost + oJob.EstBurdenCost;
                //////////decimal dActOprTotals = oJob.ActLaborCost + oJob.ActBurdenCost;
                //////////decimal dEstSubTotals = oJob.EstSubcontractCost;
                //////////decimal dActSubTotals = oJob.ActSubcontractCost;

                //////////iNumberOfRows = 1;
                //////////iNumberOfColumns = 1;
                //////////if (bFirstWorksheet == true)
                //////////{
                //////////    oSLBOMDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "Costs By Area");
                //////////    bFirstWorksheet = false;
                //////////}
                //////////else
                //////////{
                //////////    oSLBOMDocument.AddWorksheet("Costs By Area");
                //////////}

                //////////// job info
                //////////oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "JobNum");
                //////////oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns, iStandardColumnWidth);
                //////////oSLBOMDocument.SetCellStyle(iNumberOfRows, iNumberOfColumns++, oBoldStyle);
                //////////oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Parent Part Num");
                //////////oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns, iStandardColumnWidth);
                //////////oSLBOMDocument.SetCellStyle(iNumberOfRows, iNumberOfColumns++, oBoldStyle);
                //////////oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Parent Part Rev Num");
                //////////oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns, iStandardColumnWidth);
                //////////oSLBOMDocument.SetCellStyle(iNumberOfRows, iNumberOfColumns++, oBoldStyle);

                //////////iNumberOfRows++;
                //////////iNumberOfColumns = 1;
                //////////oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, oJob.JobNum);
                //////////oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                //////////oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, oJob.PartNum);
                //////////oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                //////////oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, oJob.PartRevNum);
                //////////oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);

                //////////iNumberOfRows++; // blank row

                //////////// enter the data to chart

                //////////iNumberOfRows++;

                //////////int iStartDataRow = iNumberOfRows;

                //////////// add the header for estimate and actual rows
                //////////oSLBOMDocument.SetCellValue(iNumberOfRows + 1, 1, "Estimate");
                //////////oSLBOMDocument.SetCellValue(iNumberOfRows + 2, 1, "Actual");

                //////////// now we add the column header
                //////////iNumberOfColumns = 2;
                //////////oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Materials");
                //////////oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                //////////oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Operations");
                //////////oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                //////////oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Subcontracts");
                //////////oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);

                //////////// now we add the data
                //////////iNumberOfRows++;
                //////////iNumberOfColumns = 2;

                //////////oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, dEstMtlTotals);
                //////////oSLBOMDocument.SetCellStyle(iNumberOfRows, iNumberOfColumns, oCurrencyStyle);
                //////////oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);

                //////////oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, dEstOprTotals);
                //////////oSLBOMDocument.SetCellStyle(iNumberOfRows, iNumberOfColumns, oCurrencyStyle);
                //////////oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);

                //////////oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, dEstSubTotals);
                //////////oSLBOMDocument.SetCellStyle(iNumberOfRows, iNumberOfColumns, oCurrencyStyle);
                //////////oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);

                //////////// now we enter actuals on the next row
                //////////iNumberOfRows++;
                //////////iNumberOfColumns = 2;

                //////////oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, dActMtlTotals);
                //////////oSLBOMDocument.SetCellStyle(iNumberOfRows, iNumberOfColumns, oCurrencyStyle);
                //////////oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);

                //////////oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, dActOprTotals);
                //////////oSLBOMDocument.SetCellStyle(iNumberOfRows, iNumberOfColumns, oCurrencyStyle);
                //////////oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);

                //////////oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, dActSubTotals);
                //////////oSLBOMDocument.SetCellStyle(iNumberOfRows, iNumberOfColumns, oCurrencyStyle);
                //////////oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);

                //////////int iEndDataRow = iNumberOfRows;

                //////////// blank row
                //////////iNumberOfRows++;
                //////////iNumberOfColumns = 1;

                //////////// now we add the charts
                //////////iNumberOfRows++;
                //////////fChartHeight = 15;
                //////////fChartWidth = 4;

                //////////chart = oSLBOMDocument.CreateChart(iStartDataRow, 1, iEndDataRow, 4);
                //////////chart.SetChartType(SLColumnChartType.ClusteredColumn);
                //////////chart.SetChartPosition(iNumberOfRows, 0, iNumberOfRows + fChartHeight, fChartWidth);
                //////////oSLBOMDocument.InsertChart(chart);

                //////////#endregion

                //////////#region Cumulative Op Costs

                //////////iNumberOfRows = 1;
                //////////iNumberOfColumns = 1;
                //////////if (bFirstWorksheet == true)
                //////////{
                //////////    oSLBOMDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "Cumulative Op Costs");
                //////////    bFirstWorksheet = false;
                //////////}
                //////////else
                //////////{
                //////////    oSLBOMDocument.AddWorksheet("Cumulative Op Costs");
                //////////}

                //////////// job info
                //////////oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "JobNum");
                //////////oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns, iStandardColumnWidth);
                //////////oSLBOMDocument.SetCellStyle(iNumberOfRows, iNumberOfColumns++, oBoldStyle);
                //////////oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Parent Part Num");
                //////////oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns, iStandardColumnWidth);
                //////////oSLBOMDocument.SetCellStyle(iNumberOfRows, iNumberOfColumns++, oBoldStyle);
                //////////oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Parent Part Rev Num");
                //////////oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns, iStandardColumnWidth);
                //////////oSLBOMDocument.SetCellStyle(iNumberOfRows, iNumberOfColumns++, oBoldStyle);

                //////////iNumberOfRows++;
                //////////iNumberOfColumns = 1;
                //////////oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, oJob.JobNum);
                //////////oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                //////////oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, oJob.PartNum);
                //////////oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                //////////oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, oJob.PartRevNum);
                //////////oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);

                //////////// skip a row
                //////////iNumberOfRows++;

                //////////iNumberOfRows++;

                //////////int iStartOperationRow = iNumberOfRows;
                //////////iNumberOfColumns = 3;
                //////////int iTotalOperations = 0;
                //////////// establish the op name and cost
                //////////decimal dOperationsEstimateRunningTotal = 0M;
                //////////decimal dOperationsActualRunningTotal = 0M;
                //////////foreach (HSAssembly oAssembly in oJobAssemblies)
                //////////{
                //////////    foreach (JobOpsEstVsActualCosts oOperationCost in oAssembly.OperationCosts)
                //////////    {
                //////////        // we only want operations that are not supposed to be subcontract ops
                //////////        JobOperation oThisJobOp = oAssembly.Operations.FirstOrDefault(oItem => (oItem.AssemblySeq == oOperationCost.AssemblySeq) && (oItem.OpCode == oOperationCost.OpCode));
                //////////        if ((oThisJobOp != null) && (oThisJobOp.Subcontract == false))
                //////////        {
                //////////            iTotalOperations++;

                //////////            // column header
                //////////            string sAssemblyAndOp = oAssembly.AssemblySeq.ToString() + "-" + oOperationCost.OprSeq.ToString() + "-(" + oOperationCost.OpCode + ")";
                //////////            oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, sAssemblyAndOp);
                //////////            oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns, iStandardColumnWidth);
                //////////            if (oOperationCost.OpComplete == true)
                //////////            {
                //////////                oSLBOMDocument.SetCellStyle(iNumberOfRows, iNumberOfColumns, oOperationCompleteStyle);
                //////////            }

                //////////            // est operational costs
                //////////            dOperationsEstimateRunningTotal += oOperationCost.EstLaborCost + oOperationCost.EstBurCost;
                //////////            oSLBOMDocument.SetCellValue(iNumberOfRows + 1, iNumberOfColumns, dOperationsEstimateRunningTotal);
                //////////            oSLBOMDocument.SetCellStyle(iNumberOfRows + 1, iNumberOfColumns, oCurrencyStyle);

                //////////            // act operational costs
                //////////            dOperationsActualRunningTotal += oOperationCost.ActLaborCost + oOperationCost.ActBurdenCost;
                //////////            oSLBOMDocument.SetCellValue(iNumberOfRows + 2, iNumberOfColumns, dOperationsActualRunningTotal);
                //////////            if (dOperationsActualRunningTotal > dOperationsEstimateRunningTotal)
                //////////            {
                //////////                oSLBOMDocument.SetCellStyle(iNumberOfRows + 2, iNumberOfColumns, oCurrencyStyleOverEstimate);
                //////////            }
                //////////            else
                //////////            {
                //////////                oSLBOMDocument.SetCellStyle(iNumberOfRows + 2, iNumberOfColumns, oCurrencyStyleUnderEstimate);
                //////////            }

                //////////            iNumberOfColumns++;
                //////////        }
                //////////    }
                //////////}
                //////////// add the header for estimate and actual
                //////////oSLBOMDocument.SetCellValue(iStartOperationRow + 1, 2, "Estimate");
                //////////oSLBOMDocument.SetCellValue(iStartOperationRow + 2, 2, "Actual");

                //////////// only estimate and actual so height can be fixed
                //////////fChartHeight = 30.0;
                //////////// width needs to vary with the number of operations
                //////////// we need 1 units per operation
                //////////fChartWidth = iTotalOperations;

                //////////// THIS IS A LINE CHART
                ////////////chart = oSLBOMDocument.CreateChart(2, 2, 6, iNumberOfColumns - 1);
                ////////////chart.SetChartType(SLLineChartType.Line);
                ////////////chart.SetChartPosition(7, 1, 7 + fChartHeight, 1 + fChartWidth);
                ////////////oSLBOMDocument.InsertChart(chart);

                //////////// SIMILAR TO LINE CHART BUT COLOR CODING IS SHADES OF GREEN?
                ////////////chart = oSLBOMDocument.CreateChart(2, 2, 6, iNumberOfColumns - 1);
                ////////////chart.SetChartType(SLLineChartType.StackedLineWithMarkers);
                ////////////chart.SetChartStyle(SLChartStyle.Style5);
                ////////////chart.SetChartPosition(7, 1, 7 + fChartHeight, 1 + fChartWidth);
                ////////////oSLBOMDocument.InsertChart(chart);

                //////////// BAR CHART
                //////////chart = oSLBOMDocument.CreateChart(iStartOperationRow, 2, iStartOperationRow + 2, iNumberOfColumns - 1);
                //////////chart.SetChartType(SLColumnChartType.ClusteredColumn);
                //////////chart.SetChartPosition(iStartOperationRow + 4, 1, 7 + fChartHeight, 1 + fChartWidth);
                //////////oSLBOMDocument.InsertChart(chart);

                //////////#endregion

                //////////#region Cumulative Sub Costs

                //////////iNumberOfRows = 1;
                //////////iNumberOfColumns = 1;
                //////////int iTotalSubOperations = 0;
                //////////if (bFirstWorksheet == true)
                //////////{
                //////////    oSLBOMDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "Cumulative Sub Costs");
                //////////    bFirstWorksheet = false;
                //////////}
                //////////else
                //////////{
                //////////    oSLBOMDocument.AddWorksheet("Cumulative Sub Costs");
                //////////}

                //////////// job info
                //////////oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "JobNum");
                //////////oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns, iStandardColumnWidth);
                //////////oSLBOMDocument.SetCellStyle(iNumberOfRows, iNumberOfColumns++, oBoldStyle);
                //////////oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Parent Part Num");
                //////////oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns, iStandardColumnWidth);
                //////////oSLBOMDocument.SetCellStyle(iNumberOfRows, iNumberOfColumns++, oBoldStyle);
                //////////oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Parent Part Rev Num");
                //////////oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns, iStandardColumnWidth);
                //////////oSLBOMDocument.SetCellStyle(iNumberOfRows, iNumberOfColumns++, oBoldStyle);

                //////////iNumberOfRows++;
                //////////iNumberOfColumns = 1;
                //////////oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, oJob.JobNum);
                //////////oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                //////////oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, oJob.PartNum);
                //////////oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                //////////oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, oJob.PartRevNum);
                //////////oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);

                //////////// skip a row
                //////////iNumberOfRows++;

                //////////iNumberOfRows++;

                //////////int iStartSubcontractRow = iNumberOfRows;
                //////////iNumberOfColumns = 3;
                //////////// establish the sub name and cost
                //////////decimal dSubcontractEstimateRunningTotal = 0M;
                //////////decimal dSubcontractActualRunningTotal = 0M;
                //////////foreach (HSAssembly oAssembly in oJobAssemblies)
                //////////{
                //////////    foreach (JobOpsEstVsActualCosts oOperationCost in oAssembly.OperationCosts)
                //////////    {
                //////////        // we only want operations that are subcontract ops
                //////////        JobOperation oThisJobOp = oAssembly.Operations.FirstOrDefault(oItem => (oItem.AssemblySeq == oOperationCost.AssemblySeq) && (oItem.OpCode == oOperationCost.OpCode));
                //////////        if ((oThisJobOp != null) && (oThisJobOp.Subcontract == true))
                //////////        {
                //////////            iTotalSubOperations++;

                //////////            // column header
                //////////            string sAssemblyAndOp = oAssembly.AssemblySeq.ToString() + "-" + oOperationCost.OprSeq.ToString() + "-(" + oOperationCost.OpCode + ")";
                //////////            oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, sAssemblyAndOp);
                //////////            oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns, iStandardColumnWidth);
                //////////            if (oOperationCost.OpComplete == true)
                //////////            {
                //////////                oSLBOMDocument.SetCellStyle(iNumberOfRows, iNumberOfColumns, oOperationCompleteStyle);
                //////////            }

                //////////            // est operational costs
                //////////            dSubcontractEstimateRunningTotal += oOperationCost.EstSubCost;
                //////////            oSLBOMDocument.SetCellValue(iNumberOfRows + 1, iNumberOfColumns, dSubcontractEstimateRunningTotal);
                //////////            oSLBOMDocument.SetCellStyle(iNumberOfRows + 1, iNumberOfColumns, oCurrencyStyle);

                //////////            // act operational costs
                //////////            dSubcontractActualRunningTotal += oOperationCost.ActSubCost;
                //////////            oSLBOMDocument.SetCellValue(iNumberOfRows + 2, iNumberOfColumns, dSubcontractActualRunningTotal);
                //////////            if (dSubcontractActualRunningTotal > dSubcontractEstimateRunningTotal)
                //////////            {
                //////////                oSLBOMDocument.SetCellStyle(iNumberOfRows + 2, iNumberOfColumns, oCurrencyStyleOverEstimate);
                //////////            }
                //////////            else
                //////////            {
                //////////                oSLBOMDocument.SetCellStyle(iNumberOfRows + 2, iNumberOfColumns, oCurrencyStyleUnderEstimate);
                //////////            }

                //////////            iNumberOfColumns++;
                //////////        }
                //////////    }
                //////////}
                //////////// add the header for estimate and actual
                //////////oSLBOMDocument.SetCellValue(iStartOperationRow + 1, 2, "Estimate");
                //////////oSLBOMDocument.SetCellValue(iStartOperationRow + 2, 2, "Actual");

                //////////// only estimate and actual so height can be fixed
                //////////fChartHeight = 30.0;
                //////////// width needs to vary with the number of operations
                //////////// we need 1 units per operation
                //////////fChartWidth = iTotalSubOperations;

                //////////// BAR CHART
                //////////chart = oSLBOMDocument.CreateChart(iStartOperationRow, 2, iStartOperationRow + 2, iNumberOfColumns - 1);
                //////////chart.SetChartType(SLColumnChartType.ClusteredColumn);
                //////////chart.SetChartPosition(iStartOperationRow + 4, 1, 7 + fChartHeight, 1 + fChartWidth);
                //////////oSLBOMDocument.InsertChart(chart);

                //////////#endregion

                //////////#region Cumulative Mtl Costs

                //////////iNumberOfRows = 1;
                //////////iNumberOfColumns = 1;
                //////////int iTotalMaterials = 0;
                //////////if (bFirstWorksheet == true)
                //////////{
                //////////    oSLBOMDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "Cumulative Mtl Costs");
                //////////    bFirstWorksheet = false;
                //////////}
                //////////else
                //////////{
                //////////    oSLBOMDocument.AddWorksheet("Cumulative Mtl Costs");
                //////////}

                //////////// job info
                //////////oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "JobNum");
                //////////oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns, iStandardColumnWidth);
                //////////oSLBOMDocument.SetCellStyle(iNumberOfRows, iNumberOfColumns++, oBoldStyle);
                //////////oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Parent Part Num");
                //////////oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns, iStandardColumnWidth);
                //////////oSLBOMDocument.SetCellStyle(iNumberOfRows, iNumberOfColumns++, oBoldStyle);
                //////////oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Parent Part Rev Num");
                //////////oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns, iStandardColumnWidth);
                //////////oSLBOMDocument.SetCellStyle(iNumberOfRows, iNumberOfColumns++, oBoldStyle);

                //////////iNumberOfRows++;
                //////////iNumberOfColumns = 1;
                //////////oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, oJob.JobNum);
                //////////oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                //////////oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, oJob.PartNum);
                //////////oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                //////////oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, oJob.PartRevNum);
                //////////oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);

                //////////// skip a row
                //////////iNumberOfRows++;

                //////////iNumberOfRows++;
                //////////int iStartMaterialRow = iNumberOfRows;
                //////////iNumberOfColumns = 3;
                //////////// establish the mtl name and cost
                //////////decimal dMaterialsEstimateRunningTotal = 0M;
                //////////decimal dMaterialsActualRunningTotal = 0M;
                //////////// lets get all materials for this job that have an estimate or actual cost 
                ////////////List<JobMaterial> oAllMaterials = oJob.AllMaterials.Where(oItem => (oItem.EstUnitCost != 0) || (oItem.ActTotalCost != 0)).ToList();
                //////////List<JobMaterial> oAllMaterials = oJob.AllMaterials;
                //////////// order by assembly seq then mtl seq
                //////////oAllMaterials = oAllMaterials.OrderBy(oItem => oItem.AssemblySeq).ThenBy(x => x.MtlSeq).ToList();
                //////////foreach (JobMaterial oMaterial in oAllMaterials)
                //////////{
                //////////    iTotalMaterials++;

                //////////    // column header
                //////////    string sAssemblyAndMtlSeq = oMaterial.AssemblySeq.ToString() + "-" + oMaterial.MtlSeq.ToString() + "-(" + oMaterial.MaterialPartNum + ")";
                //////////    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, sAssemblyAndMtlSeq);
                //////////    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns, iStandardColumnWidth);
                //////////    if (oMaterial.IssuedComplete == true)
                //////////    {
                //////////        oSLBOMDocument.SetCellStyle(iNumberOfRows, iNumberOfColumns, oOperationCompleteStyle);
                //////////    }

                //////////    // est material costs
                //////////    dMaterialsEstimateRunningTotal += oMaterial.EstUnitCost;
                //////////    oSLBOMDocument.SetCellValue(iNumberOfRows + 1, iNumberOfColumns, dMaterialsEstimateRunningTotal);
                //////////    oSLBOMDocument.SetCellStyle(iNumberOfRows + 1, iNumberOfColumns, oCurrencyStyle);

                //////////    // act material costs
                //////////    dMaterialsActualRunningTotal += oMaterial.ActTotalCost;
                //////////    oSLBOMDocument.SetCellValue(iNumberOfRows + 2, iNumberOfColumns, dMaterialsActualRunningTotal);
                //////////    if (dMaterialsActualRunningTotal > dMaterialsEstimateRunningTotal)
                //////////    {
                //////////        oSLBOMDocument.SetCellStyle(iNumberOfRows + 2, iNumberOfColumns, oCurrencyStyleOverEstimate);
                //////////    }
                //////////    else
                //////////    {
                //////////        oSLBOMDocument.SetCellStyle(iNumberOfRows + 2, iNumberOfColumns, oCurrencyStyleUnderEstimate);
                //////////    }

                //////////    iNumberOfColumns++;
                //////////}
                //////////// add the header for estimate and actual
                //////////oSLBOMDocument.SetCellValue(iStartMaterialRow + 1, 2, "Estimate");
                //////////oSLBOMDocument.SetCellValue(iStartMaterialRow + 2, 2, "Actual");

                //////////// only estimate and actual so height can be fixed
                //////////fChartHeight = 30.0;
                //////////// width needs to vary with the number of operations
                //////////// we need 1 units per operation
                //////////fChartWidth = iTotalMaterials;

                //////////// BAR CHART
                //////////chart = oSLBOMDocument.CreateChart(iStartMaterialRow, 2, iStartMaterialRow + 2, iNumberOfColumns - 1);
                //////////chart.SetChartType(SLColumnChartType.ClusteredColumn);
                //////////chart.SetChartPosition(iStartMaterialRow + 4, 1, 7 + fChartHeight, 1 + fChartWidth);
                //////////oSLBOMDocument.InsertChart(chart);

                //////////#endregion

                //////////#region Total Costs
                //////////decimal dEstTotals = oJob.EstMaterialCost + oJob.EstLaborCost + oJob.EstBurdenCost + oJob.EstSubcontractCost;
                //////////decimal dActTotals = oJob.ActMaterialCost + oJob.ActLaborCost + oJob.ActBurdenCost + oJob.ActSubcontractCost;

                //////////iNumberOfRows = 1;
                //////////iNumberOfColumns = 1;
                //////////if (bFirstWorksheet == true)
                //////////{
                //////////    oSLBOMDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "Total Cost");
                //////////    bFirstWorksheet = false;
                //////////}
                //////////else
                //////////{
                //////////    oSLBOMDocument.AddWorksheet("Total Cost");
                //////////}

                //////////// job info
                //////////oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "JobNum");
                //////////oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns, iStandardColumnWidth);
                //////////oSLBOMDocument.SetCellStyle(iNumberOfRows, iNumberOfColumns++, oBoldStyle);
                //////////oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Parent Part Num");
                //////////oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns, iStandardColumnWidth);
                //////////oSLBOMDocument.SetCellStyle(iNumberOfRows, iNumberOfColumns++, oBoldStyle);
                //////////oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Parent Part Rev Num");
                //////////oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns, iStandardColumnWidth);
                //////////oSLBOMDocument.SetCellStyle(iNumberOfRows, iNumberOfColumns++, oBoldStyle);

                //////////iNumberOfRows++;
                //////////iNumberOfColumns = 1;
                //////////oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, oJob.JobNum);
                //////////oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                //////////oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, oJob.PartNum);
                //////////oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                //////////oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, oJob.PartRevNum);
                //////////oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);

                //////////iNumberOfRows++; // blank row

                //////////// enter the data to chart

                //////////iNumberOfRows++;

                //////////int iStartTotalCostDataRow = iNumberOfRows;

                //////////// add the header for estimate and actual rows
                //////////oSLBOMDocument.SetCellValue(iNumberOfRows + 1, 1, "Estimate");
                //////////oSLBOMDocument.SetCellValue(iNumberOfRows + 2, 1, "Actual");

                //////////// now we add the column header
                //////////iNumberOfColumns = 2;
                //////////oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Total");
                //////////oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);

                //////////// now we add the data
                //////////iNumberOfRows++;
                //////////iNumberOfColumns = 2;

                //////////oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, dEstTotals);
                //////////oSLBOMDocument.SetCellStyle(iNumberOfRows, iNumberOfColumns, oCurrencyStyle);
                //////////oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);

                //////////// now we enter actuals on the next row
                //////////iNumberOfRows++;
                //////////iNumberOfColumns = 2;

                //////////oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, dActTotals);
                //////////oSLBOMDocument.SetCellStyle(iNumberOfRows, iNumberOfColumns, oCurrencyStyle);
                //////////oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);

                //////////int iEndTotalCostDataRow = iNumberOfRows;

                //////////// blank row
                //////////iNumberOfRows++;
                //////////iNumberOfColumns = 1;

                //////////// now we add the charts
                //////////iNumberOfRows++;
                //////////fChartHeight = 15;
                //////////fChartWidth = 2;

                //////////chart = oSLBOMDocument.CreateChart(iStartTotalCostDataRow, 1, iEndTotalCostDataRow, 2);
                //////////chart.SetChartType(SLColumnChartType.ClusteredColumn);
                ////////////chart.SetChartStyle(SLChartStyle.Style41);
                //////////chart.SetChartPosition(iNumberOfRows, 0, iNumberOfRows + fChartHeight, fChartWidth);
                //////////oSLBOMDocument.InsertChart(chart);

                //////////#endregion

                //////////#region OTHER CHART OPTIONS NOT USED
                //////////// WOULD BE GOOD FOR JOB SCHEDULE
                //////////////////////oSLBOMDocument.SetColumnWidth(1, 25);

                //////////////////////oSLBOMDocument.SetCellValue(1, 2, "Start Date");
                //////////////////////oSLBOMDocument.SetCellValue(1, 3, "Days");

                //////////////////////oSLBOMDocument.SetCellValue(2, 1, "Do landscaping");
                //////////////////////oSLBOMDocument.SetCellValue(2, 2, new DateTime(2012, 11, 24));
                //////////////////////oSLBOMDocument.SetCellValue(2, 3, 2);

                //////////////////////oSLBOMDocument.SetCellValue(3, 1, "Change to wooden flooring");
                //////////////////////oSLBOMDocument.SetCellValue(3, 2, new DateTime(2012, 11, 26));
                //////////////////////oSLBOMDocument.SetCellValue(3, 3, 10);

                //////////////////////oSLBOMDocument.SetCellValue(4, 1, "Wax the floor");
                //////////////////////oSLBOMDocument.SetCellValue(4, 2, new DateTime(2012, 12, 10));
                //////////////////////oSLBOMDocument.SetCellValue(4, 3, 5);

                //////////////////////oSLBOMDocument.SetCellValue(5, 1, "Paint walls");
                //////////////////////oSLBOMDocument.SetCellValue(5, 2, new DateTime(2012, 12, 15));
                //////////////////////oSLBOMDocument.SetCellValue(5, 3, 4);

                //////////////////////oSLBOMDocument.SetCellValue(6, 1, "Move furniture in");
                //////////////////////oSLBOMDocument.SetCellValue(6, 2, new DateTime(2012, 12, 20));
                //////////////////////oSLBOMDocument.SetCellValue(6, 3, 3);

                //////////////////////oSLBOMDocument.SetCellValue(7, 1, "Rest");
                //////////////////////oSLBOMDocument.SetCellValue(7, 2, new DateTime(2012, 12, 23));
                //////////////////////oSLBOMDocument.SetCellValue(7, 3, 1);

                //////////////////////oSLBOMDocument.SetCellValue(7, 5, "Yeah we survived! Suck it Mayan long count calendar!");

                //////////////////////oSLBOMDocument.SetCellValue(8, 1, "Holiday party!");
                //////////////////////oSLBOMDocument.SetCellValue(8, 2, new DateTime(2012, 12, 25));
                //////////////////////oSLBOMDocument.SetCellValue(8, 3, 1);

                //////////////////////SLStyle style = new SLStyle();
                //////////////////////style.FormatCode = "d-mmm";
                //////////////////////oSLBOMDocument.SetCellStyle(2, 2, 8, 2, style);

                //////////////////////SLChart chart = oSLBOMDocument.CreateChart("A1", "C8");
                //////////////////////chart.SetChartType(SLBarChartType.StackedBar);
                //////////////////////chart.SetChartPosition(10, 1, 24, 11);

                //////////////////////chart.HideChartLegend();
                //////////////////////chart.PrimaryTextAxis.InReverseOrder = true;
                //////////////////////chart.PrimaryTextAxis.SetMaximumOtherAxisCrossing();
                //////////////////////// it's not exactly 1 Jan 1900 because there's the incorrect 29 Feb 1900
                //////////////////////// but it's not worth quibbling about...
                //////////////////////chart.PrimaryValueAxis.Minimum = (new DateTime(2012, 11, 24) - new DateTime(1900, 1, 1)).Days;
                //////////////////////// we add more days to the last day so the last day is also included
                //////////////////////// Go experiment with the values...
                //////////////////////chart.PrimaryValueAxis.Maximum = (new DateTime(2012, 12, 30) - new DateTime(1900, 1, 1)).Days;
                //////////////////////// 7 days. Set the interval as weekly
                //////////////////////chart.PrimaryValueAxis.MajorUnit = 7;

                //////////////////////SLDataSeriesOptions dso = chart.GetDataSeriesOptions(1);
                //////////////////////dso.Fill.SetNoFill();
                //////////////////////dso.Line.SetNoLine();
                //////////////////////chart.SetDataSeriesOptions(1, dso);

                //////////////////////oSLBOMDocument.InsertChart(chart);


                ////////////////////double fChartHeight = 15.0;
                ////////////////////double fChartWidth = 7.5;
                ////////////////////SLChart chart;
                ////////////////////chart = sl.CreateChart("N1", "R2");
                ////////////////////chart.SetChartType(SLPieChartType.Pie);
                ////////////////////chart.SetChartPosition(piechartRow, 0, piechartRow + fChartHeight, fChartWidth);
                ////////////////////chart.Title.Shadow.SetPreset(SpreadsheetLight.Drawing.SLShadowPresetValues.PerspectiveDiagonalUpperLeft);
                ////////////////////SLGroupDataLabelOptions grplabels = chart.CreateGroupDataLabelOptions();
                ////////////////////grplabels.ShowPercentage = true;
                ////////////////////grplabels.ShowValue = false;
                ////////////////////chart.SetGroupDataLabelOptions(grplabels);
                ////////////////////sl.InsertChart(chart);


                //////////// render the chart
                ////////////oSLBOMDocument.HideRow(3);
                ////////////oSLBOMDocument.HideRow(6);
                ////////////SLChart oEstCostChart = oSLBOMDocument.CreateChart(iStartDataRow, 1, iEndDataRow, 2);
                ////////////oEstCostChart.SetChartType(SLColumnChartType.P);
                ////////////oEstCostChart.PlotDataSeriesAsPrimaryLineChart(4, SLChartDataDisplayType.Normal, false);
                ////////////oEstCostChart.PlotDataSeriesAsSecondaryLineChart(2, SLChartDataDisplayType.Normal, false);
                ////////////oEstCostChart.HideSecondaryValueAxis();
                ////////////SLDataSeriesOptions op = oEstCostChart.GetDataSeriesOptions(1);
                ////////////op.Fill.SetSolidFill(oLightRed, 0);
                ////////////oEstCostChart.SetDataSeriesOptions(1, op);

                ////////////op = oEstCostChart.GetDataSeriesOptions(2);
                ////////////op.Line.SetSolidLine(oDRed, 0);
                ////////////oEstCostChart.SetDataSeriesOptions(2, op);

                ////////////op = oEstCostChart.GetDataSeriesOptions(3);
                ////////////op.Fill.SetSolidFill(oLightGreen, 0);
                ////////////oEstCostChart.SetDataSeriesOptions(3, op);

                ////////////op = oEstCostChart.GetDataSeriesOptions(4);
                ////////////op.Line.SetSolidLine(oDGreen, 0);
                ////////////oEstCostChart.SetDataSeriesOptions(4, op);
                ////////////oEstCostChart.SetChartPosition(dChartTopPosition, dChartLeftPosition, dChartTopPosition + dChartHeight, dChartLeftPosition + dChartWidth);
                ////////////oEstCostChart.Legend.LegendPosition = LegendPositionValues.TopRight;
                ////////////oEstCostChart.Title.SetTitle("Estimate Job Costs");
                ////////////oEstCostChart.ShowChartTitle(true);
                ////////////oEstCostChart.ShowEmptyCellsAs = 0;
                ////////////oSLBOMDocument.InsertChart(oEstCostChart);
                //////////#endregion

                #endregion

                if (bDataInReport == true)
                {
                    oSLBOMDocument.SaveAs(sDestinationFileName);
                    // Check to see if we created a file and if so email it
                    if (File.Exists(sDestinationFileName) == true)
                    {
                        List<string> oAttachments = new List<string>();
                        oAttachments.Add(sDestinationFileName);
                        if (oRequestingUser != null)
                        {
                            oToAddresses.Add(oRequestingUser.Email);
                        }
                        HSEmailHelper.SendEmail(oToAddresses, sCompany + " Job Validation Report", sCompany + " Job Validation Report for " + sDate, oAttachments);
                    }
                }
            }
        }

        public void PerformJobSummary(string sCompany, string sTmpFileDirectory, HSUser oRequestingUser)
        {
            #region Setup

            bool bDataInReport = false;

            // get the file name
            DateTime dtToday = DateTime.Now;
            string sDate = dtToday.Month.ToString() + "-" + dtToday.Day.ToString() + "-" + dtToday.Year.ToString();
            string sDestinationFileName = sTmpFileDirectory + "\\" + sCompany + "-Job Summary Report-" + sDate + ".xlsx";

            HSUser oServiceAccount = HSUser.GetUserById(HSUser.SPINCRAFT_SERVICE_ACCOUNT_ID);
            List<string> oToAddresses = new List<string>();
            if (oRequestingUser == null)
            {
                // get users in the engineering, production, and quoting groups
                HSUser.GetEmailsForUsersInGroup(oToAddresses, HSUser.REPORT_ON_JOB_ESTIMATES);
            }
            else
            {
                // this report was requested by someone directly from Epicor
                oToAddresses.Add(oRequestingUser.Email);
            }
            oToAddresses.Add(HSUser.SPINCRAFT_ROOT_USER);

            int iNumberOfRows = 1;
            int iNumberOfColumns = 1;

            // colors for job summary
            System.Drawing.Color oJobInformation = System.Drawing.Color.FromArgb(221, 217, 196);
            System.Drawing.Color oCustomerInformation = System.Drawing.Color.FromArgb(228, 223, 236);
            System.Drawing.Color oProductInformation = System.Drawing.Color.FromArgb(197, 217, 241);
            System.Drawing.Color oProfitMarginInformation = System.Drawing.Color.FromArgb(253, 233, 217);
            System.Drawing.Color oOperationAndMaterialInformation = System.Drawing.Color.FromArgb(255, 255, 204);

            System.Drawing.Color oErrorRed = System.Drawing.Color.FromArgb(240, 42, 42);
            System.Drawing.Color oGoodGreen = System.Drawing.Color.FromArgb(10, 160, 40);
            System.Drawing.Color oUsingActualsForEstimates = System.Drawing.Color.FromArgb(240, 175, 175);

            SLAlignment oCenterAlignment = new SLAlignment();
            oCenterAlignment.Horizontal = HorizontalAlignmentValues.Center;

            System.Drawing.Color oHighlightYellow = System.Drawing.Color.FromArgb(255, 255, 0);

            SLStyle oHighlightHeaderStyle = new SLStyle();
            oHighlightHeaderStyle.SetFontBold(true);
            oHighlightHeaderStyle.SetFont(FontSchemeValues.Major, 12);
            oHighlightHeaderStyle.SetHorizontalAlignment(HorizontalAlignmentValues.Center);
            oHighlightHeaderStyle.SetPatternFill(PatternValues.Solid, oHighlightYellow, oHighlightYellow);
            oHighlightHeaderStyle.SetFontColor(SLThemeColorIndexValues.Dark2Color);

            #region Job Information Styles
            SLStyle oJobHeaderStyle = new SLStyle();
            oJobHeaderStyle.SetFontBold(true);
            oJobHeaderStyle.SetFont(FontSchemeValues.Major, 12);
            oJobHeaderStyle.SetHorizontalAlignment(HorizontalAlignmentValues.Center);
            oJobHeaderStyle.SetPatternFill(PatternValues.Solid, oJobInformation, oJobInformation);
            oJobHeaderStyle.SetFontColor(SLThemeColorIndexValues.Dark2Color);

            SLStyle oJobStringStyle = new SLStyle();
            oJobStringStyle.Alignment = oCenterAlignment;
            oJobStringStyle.SetPatternFill(PatternValues.Solid, oJobInformation, oJobInformation);

            SLStyle oJobPercentStyle = new SLStyle();
            oJobPercentStyle.Alignment = oCenterAlignment;
            oJobPercentStyle.FormatCode = "###.00%";
            oJobPercentStyle.SetPatternFill(PatternValues.Solid, oJobInformation, oJobInformation);

            SLStyle oJobPercentStyleUnderEstimate = new SLStyle();
            oJobPercentStyleUnderEstimate.Alignment = oCenterAlignment;
            oJobPercentStyleUnderEstimate.FormatCode = "###.00%";
            oJobPercentStyleUnderEstimate.SetFontColor(oErrorRed);
            oJobPercentStyleUnderEstimate.SetPatternFill(PatternValues.Solid, oJobInformation, oJobInformation);

            SLStyle oJobPercentStyleOverEstimate = new SLStyle();
            oJobPercentStyleOverEstimate.Alignment = oCenterAlignment;
            oJobPercentStyleOverEstimate.FormatCode = "###.00%";
            oJobPercentStyleOverEstimate.SetFontColor(oGoodGreen);
            oJobPercentStyleOverEstimate.SetPatternFill(PatternValues.Solid, oJobInformation, oJobInformation);

            #endregion

            #region Customer Information Styles
            SLStyle oCustomerHeaderStyle = new SLStyle();
            oCustomerHeaderStyle.SetFontBold(true);
            oCustomerHeaderStyle.SetFont(FontSchemeValues.Major, 12);
            oCustomerHeaderStyle.SetHorizontalAlignment(HorizontalAlignmentValues.Center);
            oCustomerHeaderStyle.SetPatternFill(PatternValues.Solid, oCustomerInformation, oCustomerInformation);
            oCustomerHeaderStyle.SetFontColor(SLThemeColorIndexValues.Dark2Color);

            SLStyle oCustomerStringStyle = new SLStyle();
            oCustomerStringStyle.Alignment = oCenterAlignment;
            oCustomerStringStyle.SetPatternFill(PatternValues.Solid, oCustomerInformation, oCustomerInformation);

            SLStyle oCustomerIntStyle = new SLStyle();
            oCustomerIntStyle.Alignment = oCenterAlignment;
            oCustomerIntStyle.FormatCode = "#########";
            oCustomerIntStyle.SetPatternFill(PatternValues.Solid, oCustomerInformation, oCustomerInformation);
            #endregion

            #region Product Information Style
            SLStyle oProductHeaderStyle = new SLStyle();
            oProductHeaderStyle.SetFontBold(true);
            oProductHeaderStyle.SetFont(FontSchemeValues.Major, 12);
            oProductHeaderStyle.SetHorizontalAlignment(HorizontalAlignmentValues.Center);
            oProductHeaderStyle.SetPatternFill(PatternValues.Solid, oProductInformation, oProductInformation);
            oProductHeaderStyle.SetFontColor(SLThemeColorIndexValues.Dark2Color);

            SLStyle oProductStringStyle = new SLStyle();
            oProductStringStyle.Alignment = oCenterAlignment;
            oProductStringStyle.SetPatternFill(PatternValues.Solid, oProductInformation, oProductInformation);

            SLStyle oProductCurrencyStyle = new SLStyle();
            oProductCurrencyStyle.ApplyNamedCellStyle(SLNamedCellStyleValues.Currency);
            oProductCurrencyStyle.Alignment = oCenterAlignment;
            oProductCurrencyStyle.SetPatternFill(PatternValues.Solid, oProductInformation, oProductInformation);
            oProductCurrencyStyle.FormatCode = "$#,##0.00";

            SLStyle oProductCurrencyStyleBoxed = new SLStyle();
            oProductCurrencyStyleBoxed.ApplyNamedCellStyle(SLNamedCellStyleValues.Currency);
            oProductCurrencyStyleBoxed.SetTopBorder(BorderStyleValues.Thick, oErrorRed);
            oProductCurrencyStyleBoxed.SetBottomBorder(BorderStyleValues.Thick, oErrorRed);
            oProductCurrencyStyleBoxed.SetLeftBorder(BorderStyleValues.Thick, oErrorRed);
            oProductCurrencyStyleBoxed.SetRightBorder(BorderStyleValues.Thick, oErrorRed);
            oProductCurrencyStyleBoxed.Alignment = oCenterAlignment;
            oProductCurrencyStyleBoxed.SetPatternFill(PatternValues.Solid, oProductInformation, oProductInformation);
            oProductCurrencyStyleBoxed.FormatCode = "$#,##0.00";

            // only used for qty so only showing whole numbers
            SLStyle oProductDecimalStyle = new SLStyle();
            oProductDecimalStyle.Alignment = oCenterAlignment;
            oProductDecimalStyle.FormatCode = "###";
            oProductDecimalStyle.SetPatternFill(PatternValues.Solid, oProductInformation, oProductInformation);

            SLStyle oProductDateStyle = new SLStyle();
            oProductDateStyle.Alignment = oCenterAlignment;
            oProductDateStyle.SetPatternFill(PatternValues.Solid, oProductInformation, oProductInformation);

            SLStyle oProductDateStyleLate = new SLStyle();
            oProductDateStyleLate.Alignment = oCenterAlignment;
            oProductDateStyleLate.SetPatternFill(PatternValues.Solid, oProductInformation, oProductInformation);
            oProductDateStyleLate.SetFontColor(oErrorRed);
            #endregion

            #region Profit Margin Style
            SLStyle oProfitMarginHeaderStyle = new SLStyle();
            oProfitMarginHeaderStyle.SetFontBold(true);
            oProfitMarginHeaderStyle.SetFont(FontSchemeValues.Major, 12);
            oProfitMarginHeaderStyle.SetHorizontalAlignment(HorizontalAlignmentValues.Center);
            oProfitMarginHeaderStyle.SetPatternFill(PatternValues.Solid, oProfitMarginInformation, oProfitMarginInformation);
            oProfitMarginHeaderStyle.SetFontColor(SLThemeColorIndexValues.Dark2Color);

            SLStyle oProfitMarginCurrencyStyle = new SLStyle();
            oProfitMarginCurrencyStyle.ApplyNamedCellStyle(SLNamedCellStyleValues.Currency);
            oProfitMarginCurrencyStyle.Alignment = oCenterAlignment;
            oProfitMarginCurrencyStyle.SetPatternFill(PatternValues.Solid, oProfitMarginInformation, oProfitMarginInformation);
            oProfitMarginCurrencyStyle.FormatCode = "$#,##0.00";

            SLStyle oProfitMarginCurrencyStyleBoxed = new SLStyle();
            oProfitMarginCurrencyStyleBoxed.ApplyNamedCellStyle(SLNamedCellStyleValues.Currency);
            oProfitMarginCurrencyStyleBoxed.Alignment = oCenterAlignment;
            oProfitMarginCurrencyStyleBoxed.SetPatternFill(PatternValues.Solid, oProfitMarginInformation, oProfitMarginInformation);
            oProfitMarginCurrencyStyleBoxed.FormatCode = "$#,##0.00";
            oProfitMarginCurrencyStyleBoxed.SetTopBorder(BorderStyleValues.Thick, oErrorRed);
            oProfitMarginCurrencyStyleBoxed.SetBottomBorder(BorderStyleValues.Thick, oErrorRed);
            oProfitMarginCurrencyStyleBoxed.SetLeftBorder(BorderStyleValues.Thick, oErrorRed);
            oProfitMarginCurrencyStyleBoxed.SetRightBorder(BorderStyleValues.Thick, oErrorRed);

            SLStyle oProfitMarginCurrencyStyleGood = new SLStyle();
            oProfitMarginCurrencyStyleGood.ApplyNamedCellStyle(SLNamedCellStyleValues.Currency);
            oProfitMarginCurrencyStyleGood.Alignment = oCenterAlignment;
            oProfitMarginCurrencyStyleGood.SetPatternFill(PatternValues.Solid, oProfitMarginInformation, oProfitMarginInformation);
            oProfitMarginCurrencyStyleGood.FormatCode = "$#,##0.00";
            oProfitMarginCurrencyStyleGood.SetFontColor(oGoodGreen);
            oProfitMarginCurrencyStyleGood.SetTopBorder(BorderStyleValues.Thick, oErrorRed);
            oProfitMarginCurrencyStyleGood.SetBottomBorder(BorderStyleValues.Thick, oErrorRed);
            oProfitMarginCurrencyStyleGood.SetLeftBorder(BorderStyleValues.Thick, oErrorRed);
            oProfitMarginCurrencyStyleGood.SetRightBorder(BorderStyleValues.Thick, oErrorRed);

            SLStyle oProfitMarginCurrencyStyleBad = new SLStyle();
            oProfitMarginCurrencyStyleBad.ApplyNamedCellStyle(SLNamedCellStyleValues.Currency);
            oProfitMarginCurrencyStyleBad.Alignment = oCenterAlignment;
            oProfitMarginCurrencyStyleBad.SetPatternFill(PatternValues.Solid, oProfitMarginInformation, oProfitMarginInformation);
            oProfitMarginCurrencyStyleBad.FormatCode = "$#,##0.00";
            oProfitMarginCurrencyStyleBad.SetFontColor(oErrorRed);
            oProfitMarginCurrencyStyleBad.SetTopBorder(BorderStyleValues.Thick, oErrorRed);
            oProfitMarginCurrencyStyleBad.SetBottomBorder(BorderStyleValues.Thick, oErrorRed);
            oProfitMarginCurrencyStyleBad.SetLeftBorder(BorderStyleValues.Thick, oErrorRed);
            oProfitMarginCurrencyStyleBad.SetRightBorder(BorderStyleValues.Thick, oErrorRed);

            SLStyle oProfitMarginPercentStyleGood = new SLStyle();
            oProfitMarginPercentStyleGood.Alignment = oCenterAlignment;
            oProfitMarginPercentStyleGood.SetPatternFill(PatternValues.Solid, oProfitMarginInformation, oProfitMarginInformation);
            oProfitMarginPercentStyleGood.FormatCode = "###.00%";
            oProfitMarginPercentStyleGood.SetFontColor(oGoodGreen);
            oProfitMarginPercentStyleGood.SetTopBorder(BorderStyleValues.Thick, oErrorRed);
            oProfitMarginPercentStyleGood.SetBottomBorder(BorderStyleValues.Thick, oErrorRed);
            oProfitMarginPercentStyleGood.SetLeftBorder(BorderStyleValues.Thick, oErrorRed);
            oProfitMarginPercentStyleGood.SetRightBorder(BorderStyleValues.Thick, oErrorRed);

            SLStyle oProfitMarginPercentStyleBad = new SLStyle();
            oProfitMarginPercentStyleBad.Alignment = oCenterAlignment;
            oProfitMarginPercentStyleBad.SetPatternFill(PatternValues.Solid, oProfitMarginInformation, oProfitMarginInformation);
            oProfitMarginPercentStyleBad.FormatCode = "###.00%";
            oProfitMarginPercentStyleBad.SetFontColor(oErrorRed); oProfitMarginPercentStyleBad.SetFontColor(oErrorRed);
            oProfitMarginPercentStyleBad.SetTopBorder(BorderStyleValues.Thick, oErrorRed);
            oProfitMarginPercentStyleBad.SetBottomBorder(BorderStyleValues.Thick, oErrorRed);
            oProfitMarginPercentStyleBad.SetLeftBorder(BorderStyleValues.Thick, oErrorRed);
            oProfitMarginPercentStyleBad.SetRightBorder(BorderStyleValues.Thick, oErrorRed);

            SLStyle oProfitMarginStringStyle = new SLStyle();
            oProfitMarginStringStyle.Alignment = oCenterAlignment;
            oProfitMarginStringStyle.SetPatternFill(PatternValues.Solid, oProfitMarginInformation, oProfitMarginInformation);
            #endregion

            #region Operation And Material Style
            SLStyle oOpMtlHeaderStyle = new SLStyle();
            oOpMtlHeaderStyle.SetFontBold(true);
            oOpMtlHeaderStyle.SetFont(FontSchemeValues.Major, 12);
            oOpMtlHeaderStyle.SetHorizontalAlignment(HorizontalAlignmentValues.Center);
            oOpMtlHeaderStyle.SetPatternFill(PatternValues.Solid, oOperationAndMaterialInformation, oOperationAndMaterialInformation);
            oOpMtlHeaderStyle.SetFontColor(SLThemeColorIndexValues.Dark2Color);

            SLStyle oOpMtlStringStyle = new SLStyle();
            oOpMtlStringStyle.Alignment = oCenterAlignment;
            oOpMtlStringStyle.SetPatternFill(PatternValues.Solid, oOperationAndMaterialInformation, oOperationAndMaterialInformation);

            SLStyle oOpMtlIntStyle = new SLStyle();
            oOpMtlIntStyle.Alignment = oCenterAlignment;
            oOpMtlIntStyle.SetPatternFill(PatternValues.Solid, oOperationAndMaterialInformation, oOperationAndMaterialInformation);

            SLStyle oOpMtlCurrencyStyle = new SLStyle();
            oOpMtlCurrencyStyle.ApplyNamedCellStyle(SLNamedCellStyleValues.Currency);
            oOpMtlCurrencyStyle.Alignment = oCenterAlignment;
            oOpMtlCurrencyStyle.SetPatternFill(PatternValues.Solid, oOperationAndMaterialInformation, oOperationAndMaterialInformation);
            oOpMtlCurrencyStyle.FormatCode = "$#,##0.00";

            SLStyle oOpMtlUsingActualsCurrencyStyle = new SLStyle();
            oOpMtlUsingActualsCurrencyStyle.ApplyNamedCellStyle(SLNamedCellStyleValues.Currency);
            oOpMtlUsingActualsCurrencyStyle.Alignment = oCenterAlignment;
            oOpMtlUsingActualsCurrencyStyle.SetPatternFill(PatternValues.Solid, oUsingActualsForEstimates, oUsingActualsForEstimates);
            oOpMtlUsingActualsCurrencyStyle.FormatCode = "$#,##0.00";

            SLStyle oOpMtlCurrencyStyleGood = new SLStyle();
            oOpMtlCurrencyStyleGood.ApplyNamedCellStyle(SLNamedCellStyleValues.Currency);
            oOpMtlCurrencyStyleGood.Alignment = oCenterAlignment;
            oOpMtlCurrencyStyleGood.SetPatternFill(PatternValues.Solid, oOperationAndMaterialInformation, oOperationAndMaterialInformation);
            oOpMtlCurrencyStyleGood.FormatCode = "$#,##0.00";
            oOpMtlCurrencyStyleGood.SetFontColor(oGoodGreen);

            SLStyle oOpMtlCurrencyStyleBad = new SLStyle();
            oOpMtlCurrencyStyleBad.ApplyNamedCellStyle(SLNamedCellStyleValues.Currency);
            oOpMtlCurrencyStyleBad.Alignment = oCenterAlignment;
            oOpMtlCurrencyStyleBad.SetPatternFill(PatternValues.Solid, oOperationAndMaterialInformation, oOperationAndMaterialInformation);
            oOpMtlCurrencyStyleBad.FormatCode = "$#,##0.00";
            oOpMtlCurrencyStyleBad.SetFontColor(oErrorRed);

            SLStyle oOpMtlPercentStyle = new SLStyle();
            oOpMtlPercentStyle.Alignment = oCenterAlignment;
            oOpMtlPercentStyle.SetPatternFill(PatternValues.Solid, oOperationAndMaterialInformation, oOperationAndMaterialInformation);
            oOpMtlPercentStyle.FormatCode = "###.00%";

            SLStyle oOpMtlPercentStyleGood = new SLStyle();
            oOpMtlPercentStyleGood.Alignment = oCenterAlignment;
            oOpMtlPercentStyleGood.SetPatternFill(PatternValues.Solid, oOperationAndMaterialInformation, oOperationAndMaterialInformation);
            oOpMtlPercentStyleGood.FormatCode = "###.00%";
            oOpMtlPercentStyleGood.SetFontColor(oGoodGreen);

            SLStyle oOpMtlPercentStyleBad = new SLStyle();
            oOpMtlPercentStyleBad.Alignment = oCenterAlignment;
            oOpMtlPercentStyleBad.SetPatternFill(PatternValues.Solid, oOperationAndMaterialInformation, oOperationAndMaterialInformation);
            oOpMtlPercentStyleBad.FormatCode = "###.00%";
            oOpMtlPercentStyleBad.SetFontColor(oErrorRed); oProfitMarginPercentStyleBad.SetFontColor(oErrorRed);

            #endregion

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
            SLDocument oSLBOMDocument = new SLDocument();

            bool bFirstWorksheet = true;

            int iStandardColumnWidth = 25;

            #endregion

            // Each job will be a worksheet
            m_oAllJobs = m_oAllJobs.OrderBy(oItem => oItem.JobNum).ToList();

            foreach (HSJob oJob in m_oAllJobs)
            {
                if ((m_bForceJobAnalysis == true) || ((oJob.HasTransactions == true) && (oJob.MarginToDate * 100.0M < m_dJobMarginThreshold)))
                {
                    #region Job Information

                    iNumberOfRows = 1;
                    iNumberOfColumns = 1;
                    if (bFirstWorksheet == true)
                    {
                        oSLBOMDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, oJob.JobNum);
                        bFirstWorksheet = false;
                    }
                    else
                    {
                        oSLBOMDocument.AddWorksheet(oJob.JobNum);
                    }

                    //set column header
                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "JOB");
                    oSLBOMDocument.SetCellStyle(iNumberOfRows, iNumberOfColumns, oJobHeaderStyle);
                    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);

                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "PERCENT COMPLETE");
                    oSLBOMDocument.SetCellStyle(iNumberOfRows, iNumberOfColumns, oJobHeaderStyle);
                    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);

                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "CURRENT MARGIN");
                    oSLBOMDocument.SetCellStyle(iNumberOfRows, iNumberOfColumns, oJobHeaderStyle);
                    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);

                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "");
                    oSLBOMDocument.SetCellStyle(iNumberOfRows, iNumberOfColumns, oJobHeaderStyle);
                    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);

                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "");
                    oSLBOMDocument.SetCellStyle(iNumberOfRows, iNumberOfColumns, oJobHeaderStyle);
                    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);

                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "");
                    oSLBOMDocument.SetCellStyle(iNumberOfRows, iNumberOfColumns, oJobHeaderStyle);
                    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);

                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "");
                    oSLBOMDocument.SetCellStyle(iNumberOfRows, iNumberOfColumns, oJobHeaderStyle);
                    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);

                    iNumberOfRows++;
                    iNumberOfColumns = 1;

                    // JOB INFORMATION
                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, oJob.JobNum);
                    oSLBOMDocument.SetCellStyle(iNumberOfRows, iNumberOfColumns++, oJobStringStyle);

                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, oJob.PercentComplete);
                    oSLBOMDocument.SetCellStyle(iNumberOfRows, iNumberOfColumns++, oJobPercentStyle);

                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, oJob.MarginToDate);
                    if (oJob.MarginToDate < m_dJobMarginThreshold)
                    {
                        oSLBOMDocument.SetCellStyle(iNumberOfRows, iNumberOfColumns++, oJobPercentStyleUnderEstimate);
                    }
                    else
                    {
                        oSLBOMDocument.SetCellStyle(iNumberOfRows, iNumberOfColumns++, oJobPercentStyleOverEstimate);
                    }

                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "");
                    oSLBOMDocument.SetCellStyle(iNumberOfRows, iNumberOfColumns++, oJobStringStyle);

                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "");
                    oSLBOMDocument.SetCellStyle(iNumberOfRows, iNumberOfColumns++, oJobStringStyle);

                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "");
                    oSLBOMDocument.SetCellStyle(iNumberOfRows, iNumberOfColumns++, oJobStringStyle);

                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "");
                    oSLBOMDocument.SetCellStyle(iNumberOfRows, iNumberOfColumns++, oJobStringStyle);

                    iNumberOfRows++;
                    bDataInReport = true;

                    #endregion

                    // inject blank line
                    iNumberOfRows++;

                    #region Customer Information
                    iNumberOfColumns = 1;
                    //set column header
                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "CUSTOMER CODE");
                    oSLBOMDocument.SetCellStyle(iNumberOfRows, iNumberOfColumns, oCustomerHeaderStyle);
                    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);

                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "CUSTOMER NAME");
                    oSLBOMDocument.SetCellStyle(iNumberOfRows, iNumberOfColumns, oCustomerHeaderStyle);
                    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);

                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "MARKET SEGMENT");
                    oSLBOMDocument.SetCellStyle(iNumberOfRows, iNumberOfColumns, oCustomerHeaderStyle);
                    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);

                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "ORDER NUM");
                    oSLBOMDocument.SetCellStyle(iNumberOfRows, iNumberOfColumns, oCustomerHeaderStyle);
                    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);

                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "LINE");
                    oSLBOMDocument.SetCellStyle(iNumberOfRows, iNumberOfColumns, oCustomerHeaderStyle);
                    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);

                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "RELEASE");
                    oSLBOMDocument.SetCellStyle(iNumberOfRows, iNumberOfColumns, oCustomerHeaderStyle);
                    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);

                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "");
                    oSLBOMDocument.SetCellStyle(iNumberOfRows, iNumberOfColumns, oCustomerHeaderStyle);
                    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);

                    iNumberOfRows++;
                    iNumberOfColumns = 1;

                    // CUSTOMER INFORMATION
                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, oJob.CustomerCode);
                    oSLBOMDocument.SetCellStyle(iNumberOfRows, iNumberOfColumns++, oCustomerStringStyle);

                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, oJob.CustomerName);
                    oSLBOMDocument.SetCellStyle(iNumberOfRows, iNumberOfColumns++, oCustomerStringStyle);

                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, oJob.MarketSegment);
                    oSLBOMDocument.SetCellStyle(iNumberOfRows, iNumberOfColumns++, oCustomerStringStyle);

                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, oJob.OrderNum);
                    oSLBOMDocument.SetCellStyle(iNumberOfRows, iNumberOfColumns++, oCustomerIntStyle);

                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, oJob.OrderLine);
                    oSLBOMDocument.SetCellStyle(iNumberOfRows, iNumberOfColumns++, oCustomerIntStyle);

                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, oJob.OrderRel);
                    oSLBOMDocument.SetCellStyle(iNumberOfRows, iNumberOfColumns++, oCustomerIntStyle);

                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "");
                    oSLBOMDocument.SetCellStyle(iNumberOfRows, iNumberOfColumns++, oCustomerStringStyle);

                    iNumberOfRows++;
                    bDataInReport = true;
                    #endregion

                    // inject blank line
                    iNumberOfRows++;

                    #region Product Information
                    iNumberOfColumns = 1;
                    //set column header
                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "PRODUCT");
                    oSLBOMDocument.SetCellStyle(iNumberOfRows, iNumberOfColumns, oProductHeaderStyle);
                    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);

                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "UNIT PRICE");
                    oSLBOMDocument.SetCellStyle(iNumberOfRows, iNumberOfColumns, oProductHeaderStyle);
                    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);

                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "QTY");
                    oSLBOMDocument.SetCellStyle(iNumberOfRows, iNumberOfColumns, oProductHeaderStyle);
                    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);

                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "EXT PRICE");
                    oSLBOMDocument.SetCellStyle(iNumberOfRows, iNumberOfColumns, oProductHeaderStyle);
                    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);

                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "PRODUCT CODE");
                    oSLBOMDocument.SetCellStyle(iNumberOfRows, iNumberOfColumns, oProductHeaderStyle);
                    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);

                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "DUE DATE");
                    oSLBOMDocument.SetCellStyle(iNumberOfRows, iNumberOfColumns, oProductHeaderStyle);
                    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);

                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "");
                    oSLBOMDocument.SetCellStyle(iNumberOfRows, iNumberOfColumns, oProductHeaderStyle);
                    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);

                    iNumberOfRows++;
                    iNumberOfColumns = 1;

                    // PRODUCT INFORMATION
                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, oJob.PartNum);
                    oSLBOMDocument.SetCellStyle(iNumberOfRows, iNumberOfColumns++, oProductStringStyle);

                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, oJob.UnitPrice);
                    oSLBOMDocument.SetCellStyle(iNumberOfRows, iNumberOfColumns++, oProductCurrencyStyle);

                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, oJob.Qty);
                    oSLBOMDocument.SetCellStyle(iNumberOfRows, iNumberOfColumns++, oProductDecimalStyle);

                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, oJob.ExtPrice);
                    oSLBOMDocument.SetCellStyle(iNumberOfRows, iNumberOfColumns++, oProductCurrencyStyleBoxed);

                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, oJob.ProductPortfolioCode);
                    oSLBOMDocument.SetCellStyle(iNumberOfRows, iNumberOfColumns++, oProductStringStyle);

                    if (oJob.DueDate != DateTime.MinValue)
                    {
                        oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, oJob.DueDate.ToShortDateString());
                        if (oJob.DueDate < DateTime.Now)
                        {
                            oSLBOMDocument.SetCellStyle(iNumberOfRows, iNumberOfColumns++, oProductDateStyleLate);
                        }
                        else
                        {
                            oSLBOMDocument.SetCellStyle(iNumberOfRows, iNumberOfColumns++, oProductDateStyle);
                        }
                    }
                    else
                    {
                        oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "");
                        oSLBOMDocument.SetCellStyle(iNumberOfRows, iNumberOfColumns++, oProductStringStyle);
                    }

                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "");
                    oSLBOMDocument.SetCellStyle(iNumberOfRows, iNumberOfColumns++, oProductStringStyle);

                    iNumberOfRows++;
                    bDataInReport = true;
                    #endregion

                    // inject blank line
                    iNumberOfRows++;

                    #region Profit And Margin
                    iNumberOfColumns = 1;
                    //set column header
                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "");
                    oSLBOMDocument.SetCellStyle(iNumberOfRows, iNumberOfColumns, oProfitMarginHeaderStyle);
                    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);

                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "COSTS TO DATE");
                    oSLBOMDocument.SetCellStyle(iNumberOfRows, iNumberOfColumns, oProfitMarginHeaderStyle);
                    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);

                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "REMAINING COSTS");
                    oSLBOMDocument.SetCellStyle(iNumberOfRows, iNumberOfColumns, oProfitMarginHeaderStyle);
                    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);

                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "TOTAL COSTS");
                    oSLBOMDocument.SetCellStyle(iNumberOfRows, iNumberOfColumns, oProfitMarginHeaderStyle);
                    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);

                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "PROFIT");
                    oSLBOMDocument.SetCellStyle(iNumberOfRows, iNumberOfColumns, oProfitMarginHeaderStyle);
                    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);

                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "MARGIN");
                    oSLBOMDocument.SetCellStyle(iNumberOfRows, iNumberOfColumns, oProfitMarginHeaderStyle);
                    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);

                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "");
                    oSLBOMDocument.SetCellStyle(iNumberOfRows, iNumberOfColumns, oProfitMarginHeaderStyle);
                    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);

                    iNumberOfRows++;
                    iNumberOfColumns = 1;

                    // PROFIT AND MARGIN INFORMATION
                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "ESTIMATED");
                    oSLBOMDocument.SetCellStyle(iNumberOfRows, iNumberOfColumns++, oProfitMarginHeaderStyle);

                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, oJob.ActTotalCost);
                    oSLBOMDocument.SetCellStyle(iNumberOfRows, iNumberOfColumns++, oProfitMarginCurrencyStyle);

                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, oJob.RemainingCostsEstimated);
                    oSLBOMDocument.SetCellStyle(iNumberOfRows, iNumberOfColumns++, oProfitMarginCurrencyStyle);

                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, oJob.ActTotalCost + oJob.RemainingCostsEstimated);
                    oSLBOMDocument.SetCellStyle(iNumberOfRows, iNumberOfColumns++, oProfitMarginCurrencyStyleBoxed);

                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, oJob.ProfitEstimated);
                    if (oJob.ProfitEstimated > 0)
                    {
                        oSLBOMDocument.SetCellStyle(iNumberOfRows, iNumberOfColumns++, oProfitMarginCurrencyStyleGood);
                    }
                    else
                    {
                        oSLBOMDocument.SetCellStyle(iNumberOfRows, iNumberOfColumns++, oProfitMarginCurrencyStyleBad);
                    }

                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, oJob.MarginEstimated);
                    if (oJob.MarginEstimated * 100.0M > m_dJobMarginThreshold)
                    {
                        oSLBOMDocument.SetCellStyle(iNumberOfRows, iNumberOfColumns++, oProfitMarginPercentStyleGood);
                    }
                    else
                    {
                        oSLBOMDocument.SetCellStyle(iNumberOfRows, iNumberOfColumns++, oProfitMarginPercentStyleBad);
                    }

                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "");
                    oSLBOMDocument.SetCellStyle(iNumberOfRows, iNumberOfColumns++, oProfitMarginStringStyle);

                    iNumberOfRows++;
                    iNumberOfColumns = 1;

                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "PROJECTED");
                    oSLBOMDocument.SetCellStyle(iNumberOfRows, iNumberOfColumns++, oProfitMarginHeaderStyle);

                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, oJob.ActTotalCost);
                    oSLBOMDocument.SetCellStyle(iNumberOfRows, iNumberOfColumns++, oProfitMarginCurrencyStyle);

                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, oJob.RemainingCostsProjected);
                    oSLBOMDocument.SetCellStyle(iNumberOfRows, iNumberOfColumns++, oProfitMarginCurrencyStyle);

                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, oJob.ActTotalCost + oJob.RemainingCostsProjected);
                    oSLBOMDocument.SetCellStyle(iNumberOfRows, iNumberOfColumns++, oProfitMarginCurrencyStyleBoxed);

                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, oJob.ProfitProjected);
                    if (oJob.ProfitProjected > 0)
                    {
                        oSLBOMDocument.SetCellStyle(iNumberOfRows, iNumberOfColumns++, oProfitMarginCurrencyStyleGood);
                    }
                    else
                    {
                        oSLBOMDocument.SetCellStyle(iNumberOfRows, iNumberOfColumns++, oProfitMarginCurrencyStyleBad);
                    }

                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, oJob.MarginProjected);
                    if (oJob.MarginProjected * 100.0M > m_dJobMarginThreshold)
                    {
                        oSLBOMDocument.SetCellStyle(iNumberOfRows, iNumberOfColumns++, oProfitMarginPercentStyleGood);
                    }
                    else
                    {
                        oSLBOMDocument.SetCellStyle(iNumberOfRows, iNumberOfColumns++, oProfitMarginPercentStyleBad);
                    }

                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "");
                    oSLBOMDocument.SetCellStyle(iNumberOfRows, iNumberOfColumns++, oProfitMarginStringStyle);

                    bDataInReport = true;
                    #endregion

                    // inject blank line
                    iNumberOfRows++;

                    #region Op And Materials
                    iNumberOfRows++;
                    iNumberOfColumns = 1;
                    //set column header
                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "ASM");
                    oSLBOMDocument.SetCellStyle(iNumberOfRows, iNumberOfColumns, oOpMtlHeaderStyle);
                    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);

                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "OPR / MTL #");
                    oSLBOMDocument.SetCellStyle(iNumberOfRows, iNumberOfColumns, oOpMtlHeaderStyle);
                    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);

                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "TYPE");
                    oSLBOMDocument.SetCellStyle(iNumberOfRows, iNumberOfColumns, oOpMtlHeaderStyle);
                    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);

                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "PERCENT COMPLETE");
                    oSLBOMDocument.SetCellStyle(iNumberOfRows, iNumberOfColumns, oOpMtlHeaderStyle);
                    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);

                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "PERCENT OF JOB");
                    oSLBOMDocument.SetCellStyle(iNumberOfRows, iNumberOfColumns, oOpMtlHeaderStyle);
                    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);

                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "ESTIMATED COSTS");
                    oSLBOMDocument.SetCellStyle(iNumberOfRows, iNumberOfColumns, oOpMtlHeaderStyle);
                    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);

                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "ACTUAL COSTS");
                    oSLBOMDocument.SetCellStyle(iNumberOfRows, iNumberOfColumns, oOpMtlHeaderStyle);
                    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);

                    iNumberOfRows++;
                    iNumberOfColumns = 1;

                    // OP AND MTL INFORMATION

                    // walk through the list of assemblies and operations in order
                    // when looking at the materials for an assembly, if the material
                    // is not tied to an operation it should appear on operation ZERO
                    List<HSAssembly> oJobAssemblies = oJob.GetAssembliesInOrder();
                    foreach (HSAssembly oAssembly in oJobAssemblies)
                    {
                        foreach (JobOpsEstVsActualCosts oOperationCost in oAssembly.OperationCosts)
                        {
                            JobOperation oThisJobOp = oAssembly.Operations.FirstOrDefault(oItem => (oItem.AssemblySeq == oOperationCost.AssemblySeq) && (oItem.OpCode == oOperationCost.OpCode));
                            if (oThisJobOp != null)
                            {
                                iNumberOfColumns = 1;

                                int iAssemblySeq = oOperationCost.AssemblySeq;
                                int iOperationSeq = oOperationCost.OprSeq;
                                decimal dEstimateCosts = 0M;
                                decimal dActualCosts = 0M;
                                string sType = "";
                                if (oThisJobOp.Subcontract == true)
                                {
                                    sType = "Subcontract - " + oThisJobOp.OpCode;
                                    dEstimateCosts = oOperationCost.EstSubCost;
                                    dActualCosts = oOperationCost.ActSubCost;
                                }
                                else
                                {
                                    sType = "Operation - " + oThisJobOp.OpCode;
                                    dEstimateCosts = oOperationCost.EstBurCost + oOperationCost.EstLaborCost;
                                    dActualCosts = oOperationCost.ActBurdenCost + oOperationCost.ActLaborCost;
                                }

                                decimal dPercentComplete = oOperationCost.PercentComplete;
                                if (oOperationCost.OpComplete == true)
                                {
                                    dPercentComplete = 1.0M;
                                }

                                decimal dPercentOfJob = 0M;
                                if (oJob.EstTotalCost != 0)
                                {
                                    dPercentOfJob = (dEstimateCosts / oJob.EstTotalCost);
                                }

                                // decided to include all operations
                                //if ((dEstimateCosts != 0) || (dActualCosts != 0))
                                {
                                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, iAssemblySeq.ToString());
                                    oSLBOMDocument.SetCellStyle(iNumberOfRows, iNumberOfColumns++, oOpMtlIntStyle);

                                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, iOperationSeq);
                                    oSLBOMDocument.SetCellStyle(iNumberOfRows, iNumberOfColumns++, oOpMtlIntStyle);

                                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, sType);
                                    oSLBOMDocument.SetCellStyle(iNumberOfRows, iNumberOfColumns++, oOpMtlStringStyle);

                                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, dPercentComplete);
                                    if (oOperationCost.OpComplete == true)
                                    {
                                        oSLBOMDocument.SetCellStyle(iNumberOfRows, iNumberOfColumns++, oOpMtlPercentStyleGood);
                                    }
                                    else
                                    {
                                        oSLBOMDocument.SetCellStyle(iNumberOfRows, iNumberOfColumns++, oOpMtlPercentStyle);
                                    }

                                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, dPercentOfJob);
                                    oSLBOMDocument.SetCellStyle(iNumberOfRows, iNumberOfColumns++, oOpMtlPercentStyle);

                                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, dEstimateCosts);
                                    if (oOperationCost.UsedActualsForMissingEstimate == true)
                                    {
                                        oSLBOMDocument.SetCellStyle(iNumberOfRows, iNumberOfColumns++, oOpMtlUsingActualsCurrencyStyle);

                                    }
                                    else
                                    {
                                        oSLBOMDocument.SetCellStyle(iNumberOfRows, iNumberOfColumns++, oOpMtlCurrencyStyle);

                                    }

                                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, dActualCosts);
                                    if (dActualCosts > dEstimateCosts)
                                    {
                                        oSLBOMDocument.SetCellStyle(iNumberOfRows, iNumberOfColumns++, oOpMtlCurrencyStyleBad);
                                    }
                                    else
                                    {
                                        oSLBOMDocument.SetCellStyle(iNumberOfRows, iNumberOfColumns++, oOpMtlCurrencyStyleGood);
                                    }

                                    iNumberOfRows++;
                                }
                            }

                            // get all material tied to this operation
                            List<JobMaterial> oAllMaterialsTiedToOperation = oAssembly.Materials.Where(Oitem => Oitem.RelatedOperation == oOperationCost.OprSeq && Oitem.AssemblySeq == oOperationCost.AssemblySeq).ToList();
                            foreach (JobMaterial oMaterial in oAllMaterialsTiedToOperation)
                            {
                                iNumberOfColumns = 1;

                                int iAssemblySeq = oMaterial.AssemblySeq;
                                int iMtlSeq = oMaterial.MtlSeq;
                                decimal dEstimateCosts = 0M;
                                decimal dActualCosts = 0M;
                                string sType = "Material - " + oMaterial.MaterialPartNum;

                                dEstimateCosts = oMaterial.EstTotalCost;
                                dActualCosts = oMaterial.ActTotalCost;

                                decimal dPercentComplete = oMaterial.PercentComplete;
                                if (oMaterial.IssuedComplete == true)
                                {
                                    dPercentComplete = 1.0M;
                                }

                                decimal dPercentOfJob = 0M;
                                if (oJob.EstTotalCost != 0)
                                {
                                    dPercentOfJob = (dEstimateCosts / oJob.EstTotalCost);
                                }

                                if ((dEstimateCosts != 0) || (dActualCosts != 0))
                                {
                                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, iAssemblySeq.ToString());
                                    oSLBOMDocument.SetCellStyle(iNumberOfRows, iNumberOfColumns++, oOpMtlIntStyle);

                                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, iMtlSeq);
                                    oSLBOMDocument.SetCellStyle(iNumberOfRows, iNumberOfColumns++, oOpMtlIntStyle);

                                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, sType);
                                    oSLBOMDocument.SetCellStyle(iNumberOfRows, iNumberOfColumns++, oOpMtlStringStyle);

                                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, dPercentComplete);
                                    if (oMaterial.IssuedComplete == true)
                                    {
                                        oSLBOMDocument.SetCellStyle(iNumberOfRows, iNumberOfColumns++, oOpMtlPercentStyleGood);
                                    }
                                    else
                                    {
                                        oSLBOMDocument.SetCellStyle(iNumberOfRows, iNumberOfColumns++, oOpMtlPercentStyle);
                                    }

                                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, dPercentOfJob);
                                    oSLBOMDocument.SetCellStyle(iNumberOfRows, iNumberOfColumns++, oOpMtlPercentStyle);

                                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, dEstimateCosts);
                                    if (oMaterial.UsedActualsForMissingEstimate == true)
                                    {
                                        oSLBOMDocument.SetCellStyle(iNumberOfRows, iNumberOfColumns++, oOpMtlUsingActualsCurrencyStyle);

                                    }
                                    else
                                    {
                                        oSLBOMDocument.SetCellStyle(iNumberOfRows, iNumberOfColumns++, oOpMtlCurrencyStyle);

                                    }

                                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, dActualCosts);
                                    if (dActualCosts > dEstimateCosts)
                                    {
                                        oSLBOMDocument.SetCellStyle(iNumberOfRows, iNumberOfColumns++, oOpMtlCurrencyStyleBad);
                                    }
                                    else
                                    {
                                        oSLBOMDocument.SetCellStyle(iNumberOfRows, iNumberOfColumns++, oOpMtlCurrencyStyleGood);
                                    }

                                    iNumberOfRows++;
                                }
                            }
                        }
                    }

                    iNumberOfRows++;
                    bDataInReport = true;
                    #endregion
                }
            }

            if (bDataInReport == true)
            {
                oSLBOMDocument.SaveAs(sDestinationFileName);
                // Check to see if we created a file and if so email it
                if (File.Exists(sDestinationFileName) == true)
                {
                    List<string> oAttachments = new List<string>();
                    oAttachments.Add(sDestinationFileName);
                    if (oRequestingUser != null)
                    {
                        oToAddresses.Add(oRequestingUser.Email);
                    }
                    HSEmailHelper.SendEmail(oToAddresses, sCompany + " Job Summary Report", sCompany + " Job Job Summary Report for " + sDate, oAttachments);
                }
            }
        }
        #endregion

        #region Properties
        #endregion

        #region Data Members

        private HSValidateParts m_oValidateParts = new HSValidateParts();
        private BOMSupport m_oBOMSupport;

        private List<JobMaterial> m_oJobMaterials = new List<JobMaterial>();
        private List<JobOperation> m_oJobOperations = new List<JobOperation>();
        private List<JobOpsEstVsActualCosts> m_oJobOpsEstVsActualCosts = new List<JobOpsEstVsActualCosts>();
        private List<JobEstVsActualCostsQty> m_oJobEstVsActualCostsQtys = new List<JobEstVsActualCostsQty>();

        Dictionary<string, List<JobMaterial>> m_oFastJobMaterials = new Dictionary<string, List<JobMaterial>>();
        Dictionary<string, List<JobOperation>> m_oFastJobOperations = new Dictionary<string, List<JobOperation>>();
        Dictionary<string, List<JobOpsEstVsActualCosts>> m_oFastJobOperationCosts = new Dictionary<string, List<JobOpsEstVsActualCosts>>();
        Dictionary<string, JobEstVsActualCostsQty> m_oFastJobEstVsActualCosts = new Dictionary<string, JobEstVsActualCostsQty>();

        private string m_sCompany;
        private decimal m_dJobEstimateAbsoluteError;
        private decimal m_dJobEstimatePercentError;
        private decimal m_dJobMarginThreshold;
        private bool m_bForceJobAnalysis;
        private bool m_bAcceptActualsForMissingEstimates;
        private bool m_bJustMissingCosts;
        private string m_sPartNum;
        private List<string> m_oJobNums;
        private List<HSJob> m_oAllJobs = new List<HSJob>();

        private List<string> m_oPartClassesToIgnore = new List<string>();

        private const decimal MIN_ABSOLUTE_ERROR = 0.01M;
        private const decimal MIN_ABSOLUTE_PERCENTAGE_ERROR = 0.01M;
        private const decimal MIN_MARGIN_THRESHOLD = -.01M;
        #endregion

    }

    public class HSJobValidation
    {
        #region constructors
        public HSJobValidation(string sCompany)
        {
            m_sCompany = sCompany;

            if (string.Compare(m_sCompany, CompanyConfiguration.SPINCRAFT_WI_COMPANY_ID, true) == 0)
            {
                // Wisconsin ignores parts in the following classes
                m_oPartClassesToIgnore.Add("CATL");
                m_oPartClassesToIgnore.Add("COTL");
                m_oPartClassesToIgnore.Add("ENGD");
                m_oPartClassesToIgnore.Add("FA");
                m_oPartClassesToIgnore.Add("GOVT");
                m_oPartClassesToIgnore.Add("INSP");
                m_oPartClassesToIgnore.Add("LTAT");
                m_oPartClassesToIgnore.Add("MFG");
                m_oPartClassesToIgnore.Add("PAIN");
                m_oPartClassesToIgnore.Add("PUR");
                m_oPartClassesToIgnore.Add("SA");
                m_oPartClassesToIgnore.Add("SHIP");
                m_oPartClassesToIgnore.Add("SPNS");
                m_oPartClassesToIgnore.Add("SPTL");
                m_oPartClassesToIgnore.Add("SUPL");
                m_oPartClassesToIgnore.Add("WELD");
            }
        }
        #endregion

        #region Methods
        public bool Initialize(Session oSession, HSValidateParts oValidateParts)
        {
            bool bSuccess = true;

            // loading all parts from the part master
            if (oValidateParts == null)
            {
                if (m_oValidateParts.Initialize(oSession) == false)
                {
                    Console.WriteLine("Failed to load the validate parts!");
                }
                oValidateParts = m_oValidateParts;
            }
            else
            {
                m_oValidateParts = oValidateParts;
            }

            // get all resources
            m_oBOMSupport = new BOMSupport(m_sCompany);
            if (m_oBOMSupport.InitializeResourceGroups(oSession) == false)
            {
                Console.WriteLine("Failed to load the BOM Support object!");
            }

            // get all resouce groups
            if (m_oBOMSupport.InitializeResources(oSession) == false)
            {
                Console.WriteLine("Failed to load the BOM Support object!");
            }

            // get a list of all materials for open jobs
            m_oJobMaterials.Clear();
            Ice.Proxy.BO.DynamicQueryImpl oDynamicQuery = WCFServiceSupport.CreateImpl<Ice.Proxy.BO.DynamicQueryImpl>(oSession, Ice.Proxy.BO.DynamicQueryImpl.UriPath);
            QueryExecutionDataSet oQueryExecutionDataSet = oDynamicQuery.GetQueryExecutionParametersByID(BAQConstants.QUERY_LIST_ALL_JOB_MATERIALS);
            oQueryExecutionDataSet.Clear();
            oQueryExecutionDataSet.AcceptChanges();
            DataSet oDataSet = oDynamicQuery.ExecuteByID(BAQConstants.QUERY_LIST_ALL_JOB_MATERIALS, oQueryExecutionDataSet);
            foreach (DataRow oRow in oDataSet.Tables[0].Rows)
            {
                JobMaterial oJobMaterial = new JobMaterial(oRow, m_oValidateParts, false);
                m_oJobMaterials.Add(oJobMaterial);
                string sTmpJobNum = oJobMaterial.JobNum;
                List<JobMaterial> oTmpJobMaterials = null;
                if (m_oFastJobMaterials.ContainsKey(sTmpJobNum) == true)
                {
                    oTmpJobMaterials = m_oFastJobMaterials[sTmpJobNum];
                }
                else
                {
                    oTmpJobMaterials = new List<JobMaterial>();
                    m_oFastJobMaterials[sTmpJobNum] = oTmpJobMaterials;
                }
                oTmpJobMaterials.Add(oJobMaterial);
            }

            m_oJobOperations.Clear();
            oDynamicQuery = WCFServiceSupport.CreateImpl<Ice.Proxy.BO.DynamicQueryImpl>(oSession, Ice.Proxy.BO.DynamicQueryImpl.UriPath);
            oQueryExecutionDataSet = oDynamicQuery.GetQueryExecutionParametersByID(BAQConstants.QUERY_LIST_ALL_JOB_OPERATIONS);
            oQueryExecutionDataSet.Clear();
            oQueryExecutionDataSet.AcceptChanges();
            oDataSet = oDynamicQuery.ExecuteByID(BAQConstants.QUERY_LIST_ALL_JOB_OPERATIONS, oQueryExecutionDataSet);
            foreach (DataRow oRow in oDataSet.Tables[0].Rows)
            {
                JobOperation oJobOperation = new JobOperation(oRow);
                m_oJobOperations.Add(oJobOperation);

                string sTmpJobNum = oJobOperation.JobNum;
                List<JobOperation> oTmpJobOperations = null;
                if (m_oFastJobOperations.ContainsKey(sTmpJobNum) == true)
                {
                    oTmpJobOperations = m_oFastJobOperations[sTmpJobNum];
                }
                else
                {
                    oTmpJobOperations = new List<JobOperation>();
                    m_oFastJobOperations[sTmpJobNum] = oTmpJobOperations;
                }
                oTmpJobOperations.Add(oJobOperation);
            }

            // pull in POC estimate vs actual costs for all open jobs
            m_oJobOpsEstVsActualCosts.Clear();
            oDynamicQuery = WCFServiceSupport.CreateImpl<Ice.Proxy.BO.DynamicQueryImpl>(oSession, Ice.Proxy.BO.DynamicQueryImpl.UriPath);
            oQueryExecutionDataSet = oDynamicQuery.GetQueryExecutionParametersByID(BAQConstants.QUERY_JOB_OPS_EST_VS_ACTUAL_COSTS);
            oQueryExecutionDataSet.Clear();
            oQueryExecutionDataSet.AcceptChanges();
            oDataSet = oDynamicQuery.ExecuteByID(BAQConstants.QUERY_JOB_OPS_EST_VS_ACTUAL_COSTS, oQueryExecutionDataSet);
            foreach (DataRow oRow in oDataSet.Tables[0].Rows)
            {
                JobOpsEstVsActualCosts oJobEstVsActualCosts = new JobOpsEstVsActualCosts(oRow, false);
                m_oJobOpsEstVsActualCosts.Add(oJobEstVsActualCosts);

                string sTmpJobNum = oJobEstVsActualCosts.JobNum;
                List<JobOpsEstVsActualCosts> oTmpJobOperationCosts = null;
                if (m_oFastJobOperationCosts.ContainsKey(sTmpJobNum) == true)
                {
                    oTmpJobOperationCosts = m_oFastJobOperationCosts[sTmpJobNum];
                }
                else
                {
                    oTmpJobOperationCosts = new List<JobOpsEstVsActualCosts>();
                    m_oFastJobOperationCosts[sTmpJobNum] = oTmpJobOperationCosts;
                }
                oTmpJobOperationCosts.Add(oJobEstVsActualCosts);
            }

            m_oJobEstVsActualCostsQtys.Clear();
            oDynamicQuery = WCFServiceSupport.CreateImpl<Ice.Proxy.BO.DynamicQueryImpl>(oSession, Ice.Proxy.BO.DynamicQueryImpl.UriPath);
            oQueryExecutionDataSet = oDynamicQuery.GetQueryExecutionParametersByID(BAQConstants.QUERY_JOB_EST_VS_ACTUAL_COSTS_QTY);
            oQueryExecutionDataSet.Clear();
            oQueryExecutionDataSet.AcceptChanges();
            oDataSet = oDynamicQuery.ExecuteByID(BAQConstants.QUERY_JOB_EST_VS_ACTUAL_COSTS_QTY, oQueryExecutionDataSet);
            foreach (DataRow oRow in oDataSet.Tables[0].Rows)
            {
                JobEstVsActualCostsQty oJobEstVsActualCostsQty = new JobEstVsActualCostsQty(oRow);
                m_oJobEstVsActualCostsQtys.Add(oJobEstVsActualCostsQty);

                string sTmpJobNum = oJobEstVsActualCostsQty.JobNum;
                if (m_oFastJobEstVsActualCosts.ContainsKey(sTmpJobNum) == false)
                {
                    m_oFastJobEstVsActualCosts[sTmpJobNum] = oJobEstVsActualCostsQty;
                }
            }

            // we need to get the unique list of job numbers
            List<string> oAllJobNums = new List<string>();
            List<string> oAllJobsForOperations = m_oJobOperations.Select(oItem => oItem.JobNum).ToList();
            oAllJobNums.AddRange(oAllJobsForOperations);
            List<string> oAllJobsForMaterials = m_oJobMaterials.Select(oItem => oItem.JobNum).ToList();
            oAllJobNums.AddRange(oAllJobsForMaterials);
            List<string> oAllJobsForOpCosts = m_oJobOpsEstVsActualCosts.Select(oItem => oItem.JobNum).ToList();
            oAllJobNums.AddRange(oAllJobsForOpCosts);
            List<string> oAllJobEstVsActualCosts = m_oJobEstVsActualCostsQtys.Select(oItem => oItem.JobNum).ToList();
            oAllJobNums.AddRange(oAllJobEstVsActualCosts);

            // get rid of all duplicate job numbers
            m_oJobNums = oAllJobNums.Distinct().ToList();

            // now we create a list of jobs
            foreach (string sTmpJobNum in m_oJobNums)
            {
                // get the list of materials and ops specific to this job
                List<JobMaterial> oTmpJobMaterials = new List<JobMaterial>();
                if (m_oFastJobMaterials.ContainsKey(sTmpJobNum) == true)
                {
                    oTmpJobMaterials = m_oFastJobMaterials[sTmpJobNum];
                }

                List<JobOperation> oTmpJobOperations = new List<JobOperation>();
                if (m_oFastJobOperations.ContainsKey(sTmpJobNum) == true)
                {
                    oTmpJobOperations = m_oFastJobOperations[sTmpJobNum];
                }

                List<JobOpsEstVsActualCosts> oTmpJobOperationCosts = new List<JobOpsEstVsActualCosts>();
                if (m_oFastJobOperationCosts.ContainsKey(sTmpJobNum) == true)
                {
                    oTmpJobOperationCosts = m_oFastJobOperationCosts[sTmpJobNum];
                }

                JobEstVsActualCostsQty oTmpJobEstVsActualCosts = new JobEstVsActualCostsQty();
                if (m_oFastJobEstVsActualCosts.ContainsKey(sTmpJobNum) == true)
                {
                    oTmpJobEstVsActualCosts = m_oFastJobEstVsActualCosts[sTmpJobNum];
                }

                HSJob oJob = new HSJob(sTmpJobNum, m_oValidateParts, oTmpJobMaterials, oTmpJobOperations, oTmpJobOperationCosts, oTmpJobEstVsActualCosts, false);
                m_oAllJobs.Add(oJob);
            }

            return bSuccess;
        }

        public void PerformValidation(string sCompany, string sTmpFileDirectory, HSUser oRequestingUser)
        {
            #region Setup

            // get the file name
            DateTime dtToday = DateTime.Now;
            string sDate = dtToday.Month.ToString() + "-" + dtToday.Day.ToString() + "-" + dtToday.Year.ToString();
            string sDestinationFileName = sTmpFileDirectory + "\\" + sCompany + "-Job Validation Report-" + sDate + ".xlsx";
            int iNumberOfRows = 1;
            int iNumberOfColumns = 1;

            HSUser oServiceAccount = HSUser.GetUserById(HSUser.SPINCRAFT_SERVICE_ACCOUNT_ID);
            List<string> oToAddresses = new List<string>();
            if (oRequestingUser == null)
            {
                // get users in the engineering, production, and quoting groups
                HSUser.GetEmailsForUsersInGroup(oToAddresses, HSUser.REPORT_ON_JOBS);
            }
            else
            {
                // this report was requested by someone directly from Epicor
                oToAddresses.Add(oRequestingUser.Email);
            }
            oToAddresses.Add(HSUser.SPINCRAFT_ROOT_USER);

            SLFill oSLFill = new SLFill();
            oSLFill.SetPatternBackgroundColor(SLThemeColorIndexValues.Accent1Color);
            oSLFill.SetPattern(PatternValues.Solid, SLThemeColorIndexValues.Light2Color, SLThemeColorIndexValues.Light2Color);

            SLAlignment oCenterAlignment = new SLAlignment();
            oCenterAlignment.Horizontal = HorizontalAlignmentValues.Center;

            SLAlignment oRightAlignment = new SLAlignment();
            oRightAlignment.Horizontal = HorizontalAlignmentValues.Right;

            SLAlignment oLeftAlignment = new SLAlignment();
            oLeftAlignment.Horizontal = HorizontalAlignmentValues.Left;

            SLStyle oBoldStyle = new SLStyle();
            oBoldStyle.SetFontBold(true);
            oBoldStyle.SetTopBorder(BorderStyleValues.Thin, SLThemeColorIndexValues.Dark2Color);
            oBoldStyle.SetBottomBorder(BorderStyleValues.Thin, SLThemeColorIndexValues.Dark2Color);
            oBoldStyle.SetLeftBorder(BorderStyleValues.Thin, SLThemeColorIndexValues.Dark2Color);
            oBoldStyle.SetRightBorder(BorderStyleValues.Thin, SLThemeColorIndexValues.Dark2Color);
            oBoldStyle.SetFontColor(SLThemeColorIndexValues.Dark2Color);

            SLStyle oBoldStyle2 = new SLStyle();
            oBoldStyle2.SetFontBold(true);
            oBoldStyle2.SetTopBorder(BorderStyleValues.Thick, SLThemeColorIndexValues.Dark2Color);
            oBoldStyle2.SetBottomBorder(BorderStyleValues.Thick, SLThemeColorIndexValues.Dark2Color);
            oBoldStyle2.SetLeftBorder(BorderStyleValues.Thick, SLThemeColorIndexValues.Dark2Color);
            oBoldStyle.SetRightBorder(BorderStyleValues.Thick, SLThemeColorIndexValues.Dark2Color);
            oBoldStyle2.SetFontColor(System.Drawing.Color.IndianRed);

            SLStyle oCurrencyStyle = new SLStyle();
            oCurrencyStyle.ApplyNamedCellStyle(SLNamedCellStyleValues.Currency);
            oCurrencyStyle.Alignment = oLeftAlignment;
            oCurrencyStyle.FormatCode = "$#,##0.00";

            SLStyle oDecimalStyle = new SLStyle();
            oDecimalStyle.Alignment = oLeftAlignment;
            oDecimalStyle.FormatCode = "###.00";

            System.Drawing.Color oHighlightYellow = System.Drawing.Color.FromArgb(255, 255, 0);

            SLStyle oHighlightHeaderStyle = new SLStyle();
            oHighlightHeaderStyle.SetFontBold(true);
            oHighlightHeaderStyle.SetFont(FontSchemeValues.Major, 12);
            oHighlightHeaderStyle.SetHorizontalAlignment(HorizontalAlignmentValues.Center);
            oHighlightHeaderStyle.SetPatternFill(PatternValues.Solid, oHighlightYellow, oHighlightYellow);
            oHighlightHeaderStyle.SetFontColor(SLThemeColorIndexValues.Dark2Color);

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
            SLDocument oSLBOMDocument = new SLDocument();

            bool bDataInReport = false;
            bool bFirstWorksheet = true;

            int iStandardColumnWidth = 30;
            #endregion

            #region Job Issues
            // jobs with zero qty
            List<HSJob> oJobsWithNoProductionQty = m_oAllJobs.Where(oItem => oItem.ProductionQty == 0).ToList();
            if (oJobsWithNoProductionQty.Count > 0)
            {
                // sort these by job number
                oJobsWithNoProductionQty = oJobsWithNoProductionQty.OrderBy(oItem => oItem.JobNum).ToList();
                iNumberOfRows = 1;
                iNumberOfColumns = 1;
                if (bFirstWorksheet == true)
                {
                    oSLBOMDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "Jobs Without Production Qty");
                    bFirstWorksheet = false;
                }
                else
                {
                    oSLBOMDocument.AddWorksheet("Jobs Without Production Qty");
                }
                //set column header
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "JobNum");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Part Num");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Part Rev Num");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Jobs Should Have A Production Qty -- SOP REQUIRES JOBS TO HAVE PRODUCTION QTY");
                oSLBOMDocument.SetCellStyle(iNumberOfRows++, iNumberOfColumns, oHighlightHeaderStyle);
                oSLBOMDocument.SetColumnWidth(iNumberOfColumns++, 90);

                foreach (HSJob oJob in oJobsWithNoProductionQty)
                {
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 1, oJob.JobNum);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 2, oJob.PartNum);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 3, oJob.PartRevNum);

                    iNumberOfRows++;
                    bDataInReport = true;
                }
            }

            // jobs missing due date
            List<HSJob> oJobsMissingDueDate = m_oAllJobs.Where(oItem => oItem.DueDate == DateTime.MinValue).ToList();
            if (oJobsMissingDueDate.Count > 0)
            {
                // sort these by job number
                oJobsMissingDueDate = oJobsMissingDueDate.OrderBy(oItem => oItem.JobNum).ToList();
                iNumberOfRows = 1;
                iNumberOfColumns = 1;
                if (bFirstWorksheet == true)
                {
                    oSLBOMDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "Jobs Without Due Date");
                    bFirstWorksheet = false;
                }
                else
                {
                    oSLBOMDocument.AddWorksheet("Jobs Without Due Date");
                }
                //set column header
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "JobNum");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Part Num");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Part Rev Num");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Jobs Should Have Due Date Set For Planning and Purchasing Purposes -- SOP REQUIRES DUE DATE TO BE SET");
                oSLBOMDocument.SetCellStyle(iNumberOfRows++, iNumberOfColumns, oHighlightHeaderStyle);
                oSLBOMDocument.SetColumnWidth(iNumberOfColumns++, 90);

                foreach (HSJob oJob in oJobsMissingDueDate)
                {
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 1, oJob.JobNum);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 2, oJob.PartNum);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 3, oJob.PartRevNum);

                    iNumberOfRows++;
                    bDataInReport = true;
                }
            }

            // jobs missing required by date
            List<HSJob> oJobsMissingRequiredByDate = m_oAllJobs.Where(oItem => oItem.RequiredDate == DateTime.MinValue).ToList();
            if (oJobsMissingRequiredByDate.Count > 0)
            {
                // sort these by job number
                oJobsMissingRequiredByDate = oJobsMissingRequiredByDate.OrderBy(oItem => oItem.JobNum).ToList();
                iNumberOfRows = 1;
                iNumberOfColumns = 1;
                if (bFirstWorksheet == true)
                {
                    oSLBOMDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "Jobs Without Required Date");
                    bFirstWorksheet = false;
                }
                else
                {
                    oSLBOMDocument.AddWorksheet("Jobs Without Required Date");
                }
                //set column header
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "JobNum");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Part Num");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Part Rev Num");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Jobs Should Have Required By Date Set For Planning and Purchasing Purposes -- SOP STATES REQUIRED DATE SHOULD BE SET");
                oSLBOMDocument.SetCellStyle(iNumberOfRows++, iNumberOfColumns, oHighlightHeaderStyle);
                oSLBOMDocument.SetColumnWidth(iNumberOfColumns++, 90);

                foreach (HSJob oJob in oJobsMissingRequiredByDate)
                {
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 1, oJob.JobNum);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 2, oJob.PartNum);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 3, oJob.PartRevNum);

                    iNumberOfRows++;
                    bDataInReport = true;
                }
            }

            // jobs missing start date
            List<HSJob> oJobsMissingStartDate = m_oAllJobs.Where(oItem => oItem.StartDate == DateTime.MinValue).ToList();
            if (oJobsMissingRequiredByDate.Count > 0)
            {
                // sort these by job number
                oJobsMissingStartDate = oJobsMissingStartDate.OrderBy(oItem => oItem.JobNum).ToList();
                iNumberOfRows = 1;
                iNumberOfColumns = 1;
                if (bFirstWorksheet == true)
                {
                    oSLBOMDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "Jobs Without Start Date");
                    bFirstWorksheet = false;
                }
                else
                {
                    oSLBOMDocument.AddWorksheet("Jobs Without Start Date");
                }
                //set column header
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "JobNum");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Part Num");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Part Rev Num");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Jobs Should Have Start Date Set For Planning and Purchasing Purposes -- SOP REQUIRES START DATE TO BE SET");
                oSLBOMDocument.SetCellStyle(iNumberOfRows++, iNumberOfColumns, oHighlightHeaderStyle);
                oSLBOMDocument.SetColumnWidth(iNumberOfColumns++, 90);

                foreach (HSJob oJob in oJobsMissingStartDate)
                {
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 1, oJob.JobNum);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 2, oJob.PartNum);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 3, oJob.PartRevNum);

                    iNumberOfRows++;
                    bDataInReport = true;
                }
            }

            // jobs that have parts on the fly
            List<HSJob> oJobsWithoutPartMaster = m_oAllJobs.Where(oItem => (oItem.PartMaster == null)).ToList();
            if (oJobsWithoutPartMaster.Count > 0)
            {
                // sort these by job number
                oJobsWithoutPartMaster = oJobsWithoutPartMaster.OrderBy(oItem => oItem.JobNum).ToList();
                iNumberOfRows = 1;
                iNumberOfColumns = 1;
                if (bFirstWorksheet == true)
                {
                    oSLBOMDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "Jobs Without Part Master");
                    bFirstWorksheet = false;
                }
                else
                {
                    oSLBOMDocument.AddWorksheet("Jobs Without Part Master");
                }
                //set column header
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "JobNum");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Part Num");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Part Rev Num");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns, iStandardColumnWidth);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Part Does Not Exist In Part Master");
                oSLBOMDocument.SetCellStyle(iNumberOfRows++, iNumberOfColumns, oHighlightHeaderStyle);
                oSLBOMDocument.SetColumnWidth(iNumberOfColumns++, 60);

                foreach (HSJob oJob in oJobsWithoutPartMaster)
                {
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 1, oJob.JobNum);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 2, oJob.PartNum);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 3, oJob.PartRevNum);

                    iNumberOfRows++;
                    bDataInReport = true;
                }
            }

            // jobs that have make direct finished goods but are not tied to SO -- issue trying to compute margin for POC
            List<HSJob> oMakeDirectJobsNotTiedToSalesOrder = m_oAllJobs.Where(oItem => (oItem.OrderNum == 0) && (oItem.PartMaster != null) && (oItem.PartMaster.PartNonStock == true)).ToList();
            if (oMakeDirectJobsNotTiedToSalesOrder.Count > 0)
            {
                // sort these by job number
                oMakeDirectJobsNotTiedToSalesOrder = oMakeDirectJobsNotTiedToSalesOrder.OrderBy(oItem => oItem.JobNum).ToList();
                iNumberOfRows = 1;
                iNumberOfColumns = 1;
                if (bFirstWorksheet == true)
                {
                    oSLBOMDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "Make Direct No SO");
                    bFirstWorksheet = false;
                }
                else
                {
                    oSLBOMDocument.AddWorksheet("Make Direct No SO");
                }
                //set column header
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "JobNum");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Part Num");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Part Rev Num");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Job Is Make Direct But Not Tied To SO");
                oSLBOMDocument.SetCellStyle(iNumberOfRows++, iNumberOfColumns, oHighlightHeaderStyle);
                oSLBOMDocument.SetColumnWidth(iNumberOfColumns++, 60);

                foreach (HSJob oJob in oMakeDirectJobsNotTiedToSalesOrder)
                {
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 1, oJob.JobNum);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 2, oJob.PartNum);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 3, oJob.PartRevNum);

                    iNumberOfRows++;
                    bDataInReport = true;
                }
            }

            // make direct tied to closed order
            List<HSJob> oJobsTiedToClosedRelease = m_oAllJobs.Where(oItem => (oItem.OrderNum != 0) && (oItem.OpenRelease == false)).ToList();
            if (oJobsTiedToClosedRelease.Count > 0)
            {
                // sort these by job number
                oJobsTiedToClosedRelease = oJobsTiedToClosedRelease.OrderBy(oItem => oItem.JobNum).ToList();
                iNumberOfRows = 1;
                iNumberOfColumns = 1;
                if (bFirstWorksheet == true)
                {
                    oSLBOMDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "Job Closed Rel");
                    bFirstWorksheet = false;
                }
                else
                {
                    oSLBOMDocument.AddWorksheet("Job Closed Rel");
                }
                //set column header
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "JobNum");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Part Num");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Part Rev Num");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Order Num");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Order Line");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Order Rel");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Make Direct Job Tied To Closed Release");
                oSLBOMDocument.SetCellStyle(iNumberOfRows++, iNumberOfColumns, oHighlightHeaderStyle);
                oSLBOMDocument.SetColumnWidth(iNumberOfColumns++, 60);

                foreach (HSJob oJob in oJobsTiedToClosedRelease)
                {
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 1, oJob.JobNum);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 2, oJob.PartNum);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 3, oJob.PartRevNum);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 4, oJob.OrderNum);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 5, oJob.OrderLine);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 6, oJob.OrderRel);

                    iNumberOfRows++;
                    bDataInReport = true;
                }
            }

            // make direct job that was received to stock
            List<HSJob> oMakeDirectJobReceivedToStock = m_oAllJobs.Where(oItem => (oItem.OrderNum != 0) && (oItem.ReceivedQty > 0)).ToList();
            if (oMakeDirectJobReceivedToStock.Count > 0)
            {
                // sort these by job number
                oMakeDirectJobReceivedToStock = oMakeDirectJobReceivedToStock.OrderBy(oItem => oItem.JobNum).ToList();
                iNumberOfRows = 1;
                iNumberOfColumns = 1;
                if (bFirstWorksheet == true)
                {
                    oSLBOMDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "Job Rcvd Stock");
                    bFirstWorksheet = false;
                }
                else
                {
                    oSLBOMDocument.AddWorksheet("Job Rcvd Stock");
                }
                //set column header
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Job Num");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Part Num");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Part Rev Num");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Order Num");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Order Line");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Order Rel");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Received Qty");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Make Direct Job But We Received Qty Into Stock");
                oSLBOMDocument.SetCellStyle(iNumberOfRows++, iNumberOfColumns, oHighlightHeaderStyle);
                oSLBOMDocument.SetColumnWidth(iNumberOfColumns++, 60);

                foreach (HSJob oJob in oMakeDirectJobReceivedToStock)
                {
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 1, oJob.JobNum);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 2, oJob.PartNum);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 3, oJob.PartRevNum);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 4, oJob.OrderNum);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 5, oJob.OrderLine);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 6, oJob.OrderRel);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 7, oJob.ReceivedQty);

                    iNumberOfRows++;
                    bDataInReport = true;
                }
            }

            // make direct job prod qty does not match release qty
            List<HSJob> oJobQtyNotEqualToRelQty = m_oAllJobs.Where(oItem => (oItem.OrderNum != 0) && (oItem.ProductionQty != oItem.Qty)).ToList();
            if (oJobQtyNotEqualToRelQty.Count > 0)
            {
                // sort these by job number
                oJobQtyNotEqualToRelQty = oJobQtyNotEqualToRelQty.OrderBy(oItem => oItem.JobNum).ToList();
                iNumberOfRows = 1;
                iNumberOfColumns = 1;
                if (bFirstWorksheet == true)
                {
                    oSLBOMDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "Job Qty And Rel Qty");
                    bFirstWorksheet = false;
                }
                else
                {
                    oSLBOMDocument.AddWorksheet("Job Qty And Rel Qty");
                }
                //set column header
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "JobNum");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Part Num");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Part Rev Num");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Order Num");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Order Line");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Order Rel");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Release Qty");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Production Qty");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Job Qty Not The Same As Release Qty");
                oSLBOMDocument.SetCellStyle(iNumberOfRows++, iNumberOfColumns, oHighlightHeaderStyle);
                oSLBOMDocument.SetColumnWidth(iNumberOfColumns++, 60);

                foreach (HSJob oJob in oJobQtyNotEqualToRelQty)
                {
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 1, oJob.JobNum);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 2, oJob.PartNum);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 3, oJob.PartRevNum);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 4, oJob.OrderNum);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 5, oJob.OrderLine);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 6, oJob.OrderRel);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 7, oJob.Qty);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 8, oJob.ProductionQty);

                    iNumberOfRows++;
                    bDataInReport = true;
                }
            }

            // make direct jobs where the sales order shows more have been shipped for the rel than required -- could be job or sales order problem
            List<HSJob> oShippedMoreThanRelQty = m_oAllJobs.Where(oItem => (oItem.OrderNum != 0) && (oItem.ShippedQty > oItem.Qty)).ToList();
            if (oShippedMoreThanRelQty.Count > 0)
            {
                // sort these by job number
                oShippedMoreThanRelQty = oShippedMoreThanRelQty.OrderBy(oItem => oItem.JobNum).ToList();
                iNumberOfRows = 1;
                iNumberOfColumns = 1;
                if (bFirstWorksheet == true)
                {
                    oSLBOMDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "Shipped Too Much");
                    bFirstWorksheet = false;
                }
                else
                {
                    oSLBOMDocument.AddWorksheet("Shipped Too Much");
                }
                //set column header
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "JobNum");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Part Num");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Part Rev Num");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Order Num");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Order Line");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Order Rel");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Rel Qty");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Shipped Qty");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Shipped More Than Required On Release - Check Sales Order And Job");
                oSLBOMDocument.SetCellStyle(iNumberOfRows++, iNumberOfColumns, oHighlightHeaderStyle);
                oSLBOMDocument.SetColumnWidth(iNumberOfColumns++, 60);

                foreach (HSJob oJob in oShippedMoreThanRelQty)
                {
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 1, oJob.JobNum);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 2, oJob.PartNum);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 3, oJob.PartRevNum);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 4, oJob.OrderNum);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 5, oJob.OrderLine);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 6, oJob.OrderRel);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 7, oJob.Qty);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 8, oJob.ShippedQty);

                    iNumberOfRows++;
                    bDataInReport = true;
                }
            }

            // jobs that are stocking but tied to sales order
            List<HSJob> oStockJobsTiedToSalesOrder = m_oAllJobs.Where(oItem => (oItem.OrderNum != 0) && (oItem.PartMaster != null) && (oItem.PartMaster.PartNonStock == false)).ToList();
            if (oStockJobsTiedToSalesOrder.Count > 0)
            {
                // sort these by job number
                oStockJobsTiedToSalesOrder = oStockJobsTiedToSalesOrder.OrderBy(oItem => oItem.JobNum).ToList();
                iNumberOfRows = 1;
                iNumberOfColumns = 1;
                if (bFirstWorksheet == true)
                {
                    oSLBOMDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "Stock MFG Tied To SO");
                    bFirstWorksheet = false;
                }
                else
                {
                    oSLBOMDocument.AddWorksheet("Stock MFG Tied To SO");
                }
                //set column header
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "JobNum");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Part Num");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Part Rev Num");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Order Num");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Order Line");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Order Release");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Stocking Job Is Tied To SO");
                oSLBOMDocument.SetCellStyle(iNumberOfRows++, iNumberOfColumns, oHighlightHeaderStyle);
                oSLBOMDocument.SetColumnWidth(iNumberOfColumns++, 60);

                foreach (HSJob oJob in oStockJobsTiedToSalesOrder)
                {
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 1, oJob.JobNum);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 2, oJob.PartNum);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 3, oJob.PartRevNum);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 4, oJob.OrderNum);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 5, oJob.OrderLine);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 6, oJob.OrderRel);

                    iNumberOfRows++;
                    bDataInReport = true;
                }
            }

            // stock job that has shipped not received to stock
            List<HSJob> oStockJobShipped = m_oAllJobs.Where(oItem => (oItem.OrderNum == 0) && (oItem.ShippedQty > 0)).ToList();
            if (oStockJobShipped.Count > 0)
            {
                // sort these by job number
                oStockJobShipped = oStockJobShipped.OrderBy(oItem => oItem.JobNum).ToList();
                iNumberOfRows = 1;
                iNumberOfColumns = 1;
                if (bFirstWorksheet == true)
                {
                    oSLBOMDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "Stk Job Shipped");
                    bFirstWorksheet = false;
                }
                else
                {
                    oSLBOMDocument.AddWorksheet("Stk Job Shipped");
                }
                //set column header
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Job Num");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Part Num");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Part Rev Num");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Shipped Qty");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Stock Job Shipped Not Recieved To Stock");
                oSLBOMDocument.SetCellStyle(iNumberOfRows++, iNumberOfColumns, oHighlightHeaderStyle);
                oSLBOMDocument.SetColumnWidth(iNumberOfColumns++, 60);

                foreach (HSJob oJob in oStockJobShipped)
                {
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 1, oJob.JobNum);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 2, oJob.PartNum);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 3, oJob.PartRevNum);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 4, oJob.ShippedQty);

                    iNumberOfRows++;
                    bDataInReport = true;
                }
            }

            // production qty is less than what was shipped and received
            List<HSJob> oReceivedMoreThanProdQty = m_oAllJobs.Where(oItem => (oItem.ProductionQty < oItem.ReceivedQty + oItem.ShippedQty)).ToList();
            if (oReceivedMoreThanProdQty.Count > 0)
            {
                // sort these by job number
                oReceivedMoreThanProdQty = oReceivedMoreThanProdQty.OrderBy(oItem => oItem.JobNum).ToList();
                iNumberOfRows = 1;
                iNumberOfColumns = 1;
                if (bFirstWorksheet == true)
                {
                    oSLBOMDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "Prod Qty < Recvd And Ship");
                    bFirstWorksheet = false;
                }
                else
                {
                    oSLBOMDocument.AddWorksheet("Prod Qty < Recvd And Ship");
                }
                //set column header
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "JobNum");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Part Num");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Part Rev Num");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Production Qty");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Received Qty");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Shipped Qty");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Received And Shipped More Than Job Qty");
                oSLBOMDocument.SetCellStyle(iNumberOfRows++, iNumberOfColumns, oHighlightHeaderStyle);
                oSLBOMDocument.SetColumnWidth(iNumberOfColumns++, 60);

                foreach (HSJob oJob in oReceivedMoreThanProdQty)
                {
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 1, oJob.JobNum);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 2, oJob.PartNum);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 3, oJob.PartRevNum);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 4, oJob.ProductionQty);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 5, oJob.ReceivedQty);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 6, oJob.ShippedQty);

                    iNumberOfRows++;
                    bDataInReport = true;
                }
            }

            // qty complete but still have operations open / mtls not issued
            List<HSJob> oQtyCompleteButOperationsOpen = new List<HSJob>();
            List<HSJob> oTmpJobs = m_oAllJobs.Where(oItem => (oItem.ProductionQty <= oItem.ShippedQty + oItem.ReceivedQty)).ToList();
            // check to see if we have any materials not completely issued
            foreach (HSJob oTmp in oTmpJobs)
            {
                List<JobMaterial> oTmpMaterials = oTmp.AllMaterials.Where(OItem => OItem.IssuedComplete == false).ToList();
                List<JobOperation> oTmpOperations = oTmp.AllOperations.Where(oItem => oItem.OpComplete == false).ToList();
                if ( (oTmpMaterials.Count > 0) || (oTmpOperations.Count > 0) )
                {
                    oQtyCompleteButOperationsOpen.Add(oTmp);
                }
            }
            if (oQtyCompleteButOperationsOpen.Count > 0)
            {
                // sort these by job number
                oQtyCompleteButOperationsOpen = oQtyCompleteButOperationsOpen.OrderBy(oItem => oItem.JobNum).ToList();
                iNumberOfRows = 1;
                iNumberOfColumns = 1;
                if (bFirstWorksheet == true)
                {
                    oSLBOMDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "Qty Complete");
                    bFirstWorksheet = false;
                }
                else
                {
                    oSLBOMDocument.AddWorksheet("Qty Complete");
                }
                //set column header
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Job Num");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Part Num");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Part Rev Num");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Production Qty");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Shipped Qty");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Received Qty");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Job Qty Complete But Operations Or Materials Not Complete");
                oSLBOMDocument.SetCellStyle(iNumberOfRows++, iNumberOfColumns, oHighlightHeaderStyle);
                oSLBOMDocument.SetColumnWidth(iNumberOfColumns++, 90);

                foreach (HSJob oJob in oQtyCompleteButOperationsOpen)
                {
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 1, oJob.JobNum);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 2, oJob.PartNum);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 3, oJob.PartRevNum);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 4, oJob.ProductionQty);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 5, oJob.ShippedQty);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 6, oJob.ReceivedQty);

                    iNumberOfRows++;
                    bDataInReport = true;
                }
            }

            //ops not completed in order / holes in process

            // jobs that were started over 2 years ago and have not been clocked into in over 1 year
            DateTime dtOneYearAgo = DateTime.Now;
            dtOneYearAgo = dtOneYearAgo.AddYears(-1);
            DateTime dtTwoYearsOld = DateTime.Now;
            dtTwoYearsOld = dtTwoYearsOld.AddYears(-2);
            List<HSJob> oOldJobs = m_oAllJobs.Where(oItem => ((oItem.StartDate == DateTime.MinValue) || (oItem.StartDate <= dtTwoYearsOld)) && (oItem.LastLoginDate <= dtOneYearAgo)).ToList();
            if (oOldJobs.Count > 0)
            {
                // sort these by job number
                oOldJobs = oOldJobs.OrderBy(oItem => oItem.JobNum).ToList();
                iNumberOfRows = 1;
                iNumberOfColumns = 1;
                if (bFirstWorksheet == true)
                {
                    oSLBOMDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "old Jobs");
                    bFirstWorksheet = false;
                }
                else
                {
                    oSLBOMDocument.AddWorksheet("Old Jobs");
                }
                //set column header
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "JobNum");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Part Num");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Part Rev Num");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Start Date");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Last Login Date");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Due Date");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Jobs Started More Than 2 Years Ago Not Logged Into Recently");
                oSLBOMDocument.SetCellStyle(iNumberOfRows++, iNumberOfColumns, oHighlightHeaderStyle);
                oSLBOMDocument.SetColumnWidth(iNumberOfColumns++, 90);

                foreach (HSJob oJob in oOldJobs)
                {
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 1, oJob.JobNum);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 2, oJob.PartNum);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 3, oJob.PartRevNum);
                    if (oJob.StartDate == DateTime.MinValue)
                    {
                        oSLBOMDocument.SetCellValue(iNumberOfRows, 4, "");
                    }
                    else
                    {
                        oSLBOMDocument.SetCellValue(iNumberOfRows, 4, oJob.StartDate.ToShortDateString());
                    }
                    if (oJob.LastLoginDate == DateTime.MinValue)
                    {
                        oSLBOMDocument.SetCellValue(iNumberOfRows, 5, "");
                    }
                    else
                    {
                        oSLBOMDocument.SetCellValue(iNumberOfRows, 5, oJob.LastLoginDate.ToShortDateString());
                    }
                    if (oJob.DueDate == DateTime.MinValue)
                    {
                        oSLBOMDocument.SetCellValue(iNumberOfRows, 6, "");
                    }
                    else
                    {
                        oSLBOMDocument.SetCellValue(iNumberOfRows, 6, oJob.DueDate.ToShortDateString());
                    }
                    iNumberOfRows++;
                    bDataInReport = true;
                }
            }

            #endregion

            #region Material Attributes

            // jobs where the material qty should not be zero
            List< JobMaterial> oMaterialsWithZeroQty = m_oJobMaterials.Where(oItem => (oItem.RequiredQty == 0) && (oItem.QtyBearing == true) && ((oItem.PartMaster == null) || (m_oPartClassesToIgnore.Contains(oItem.PartMaster.ClassID) == false))).ToList();
            if (oMaterialsWithZeroQty.Count > 0)
            {
                // sort these by job number
                oMaterialsWithZeroQty = oMaterialsWithZeroQty.OrderBy(oItem => oItem.JobNum).ToList();
                iNumberOfRows = 1;
                iNumberOfColumns = 1;
                if (bFirstWorksheet == true)
                {
                    oSLBOMDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "Mtl Qty Zero");
                    bFirstWorksheet = false;
                }
                else
                {
                    oSLBOMDocument.AddWorksheet("Mtl Qty Zero");
                }
                //set column header
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "JobNum");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Parent Part Num");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Parent Part Rev Num");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Assembly Seq");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Mtl Seq");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Mtl Part Num");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Mtl Part Rev Num");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Class");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Mtl On Job Has Qty Of Zero");
                oSLBOMDocument.SetCellStyle(iNumberOfRows++, iNumberOfColumns, oHighlightHeaderStyle);
                oSLBOMDocument.SetColumnWidth(iNumberOfColumns++, 60);

                foreach (JobMaterial oJobMateiral in oMaterialsWithZeroQty)
                {
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 1, oJobMateiral.JobNum);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 2, oJobMateiral.ParentPartNum);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 3, oJobMateiral.ParentRevNum);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 4, oJobMateiral.AssemblySeq.ToString());
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 5, oJobMateiral.MtlSeq.ToString());
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 6, oJobMateiral.MaterialPartNum);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 7, oJobMateiral.MaterialPartRevNum);
                    if (oJobMateiral.PartMaster != null)
                    {
                        oSLBOMDocument.SetCellValue(iNumberOfRows, 8, oJobMateiral.PartMaster.ClassID);
                    }

                    iNumberOfRows++;
                    bDataInReport = true;
                }
            }

            #region IGNORING THESE ISSUES
            // materials not tied to operatiopn
            //////////List<JobMaterial> oMaterialsNotAssociatedToOp = m_oJobMaterials.Where(oItem => (oItem.RelatedOperation == 0) && (oItem.PartMaster != null) && (m_oPartClassesToIgnore.Contains(oItem.PartMaster.ClassID) == false)).ToList();
            //////////if (oMaterialsNotAssociatedToOp.Count > 0)
            //////////{
            //////////    // sort these by job number
            //////////    oMaterialsNotAssociatedToOp = oMaterialsNotAssociatedToOp.OrderBy(oItem => oItem.JobNum).ToList();
            //////////    iNumberOfRows = 1;
            //////////    iNumberOfColumns = 1;
            //////////    if (bFirstWorksheet == true)
            //////////    {
            //////////        oSLBOMDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "Mtl No Op");
            //////////        bFirstWorksheet = false;
            //////////    }
            //////////    else
            //////////    {
            //////////        oSLBOMDocument.AddWorksheet("Mtl No Op");
            //////////    }
            //////////    //set column header
            //////////    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "JobNum");
            //////////    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
            //////////    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Parent Part Num");
            //////////    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
            //////////    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Parent Part Rev Num");
            //////////    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
            //////////    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Assembly Seq");
            //////////    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
            //////////    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Mtl Seq");
            //////////    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
            //////////    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Mtl Part Num");
            //////////    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
            //////////    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Mtl Part Rev Num");
            //////////    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
            //////////    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Class");
            //////////    oSLBOMDocument.SetColumnWidth(iNumberOfRows++, iNumberOfColumns++, iStandardColumnWidth);

            //////////    foreach (JobMaterial oJobMateiral in oMaterialsNotAssociatedToOp)
            //////////    {
            //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, 1, oJobMateiral.JobNum);
            //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, 2, oJobMateiral.ParentPartNum);
            //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, 3, oJobMateiral.ParentRevNum);
            //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, 4, oJobMateiral.AssemblySeq.ToString());
            //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, 5, oJobMateiral.MtlSeq.ToString());
            //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, 6, oJobMateiral.MaterialPartNum);
            //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, 7, oJobMateiral.MaterialPartRevNum);
            //////////        if (oJobMateiral.PartMaster != null)
            //////////        {
            //////////            oSLBOMDocument.SetCellValue(iNumberOfRows, 8, oJobMateiral.PartMaster.ClassID);
            //////////        }

            //////////        iNumberOfRows++;
            //////////        bDataInReport = true;
            //////////    }
            //////////}

            // materials no lead time
            //////////List<JobMaterial> oMaterialsNoLeadTime = m_oJobMaterials.Where(oItem => (oItem.LeadTime == 0) && (oItem.PartMaster != null) && (m_oPartClassesToIgnore.Contains(oItem.PartMaster.ClassID) == false)).ToList();
            //////////if (oMaterialsNoLeadTime.Count > 0)
            //////////{
            //////////    // sort these by job number
            //////////    oMaterialsNoLeadTime = oMaterialsNoLeadTime.OrderBy(oItem => oItem.JobNum).ToList();
            //////////    iNumberOfRows = 1;
            //////////    iNumberOfColumns = 1;
            //////////    if (bFirstWorksheet == true)
            //////////    {
            //////////        oSLBOMDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "Mtl No Lead Time");
            //////////        bFirstWorksheet = false;
            //////////    }
            //////////    else
            //////////    {
            //////////        oSLBOMDocument.AddWorksheet("Mtl No Lead Time");
            //////////    }
            //////////    //set column header
            //////////    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "JobNum");
            //////////    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
            //////////    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Parent Part Num");
            //////////    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
            //////////    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Parent Part Rev Num");
            //////////    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
            //////////    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Assembly Seq");
            //////////    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
            //////////    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Mtl Seq");
            //////////    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
            //////////    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Mtl Part Num");
            //////////    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
            //////////    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Mtl Part Rev Num");
            //////////    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
            //////////    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Class");
            //////////    oSLBOMDocument.SetColumnWidth(iNumberOfRows++, iNumberOfColumns++, iStandardColumnWidth);

            //////////    foreach (JobMaterial oJobMateiral in oMaterialsNoLeadTime)
            //////////    {
            //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, 1, oJobMateiral.JobNum);
            //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, 2, oJobMateiral.ParentPartNum);
            //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, 3, oJobMateiral.ParentRevNum);
            //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, 4, oJobMateiral.AssemblySeq.ToString());
            //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, 5, oJobMateiral.MtlSeq.ToString());
            //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, 6, oJobMateiral.MaterialPartNum);
            //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, 7, oJobMateiral.MaterialPartRevNum);
            //////////        if (oJobMateiral.PartMaster != null)
            //////////        {
            //////////            oSLBOMDocument.SetCellValue(iNumberOfRows, 8, oJobMateiral.PartMaster.ClassID);
            //////////        }

            //////////        iNumberOfRows++;
            //////////        bDataInReport = true;
            //////////    }
            //////////}

            // materials no required date
            //////////List<JobMaterial> oMaterialsNoRequiredDate = m_oJobMaterials.Where(oItem => (oItem.RequiredDate == null) || (oItem.RequiredDate == DateTime.MinValue) && (oItem.PartMaster != null) && (m_oPartClassesToIgnore.Contains(oItem.PartMaster.ClassID) == false)).ToList();
            //////////if (oMaterialsNoRequiredDate.Count > 0)
            //////////{
            //////////    // sort these by job number
            //////////    oMaterialsNoRequiredDate = oMaterialsNoRequiredDate.OrderBy(oItem => oItem.JobNum).ToList();
            //////////    iNumberOfRows = 1;
            //////////    iNumberOfColumns = 1;
            //////////    if (bFirstWorksheet == true)
            //////////    {
            //////////        oSLBOMDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "Mtl No Required Date");
            //////////        bFirstWorksheet = false;
            //////////    }
            //////////    else
            //////////    {
            //////////        oSLBOMDocument.AddWorksheet("Mtl No Required Date");
            //////////    }
            //////////    //set column header
            //////////    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "JobNum");
            //////////    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
            //////////    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Parent Part Num");
            //////////    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
            //////////    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Parent Part Rev Num");
            //////////    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
            //////////    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Assembly Seq");
            //////////    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
            //////////    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Mtl Seq");
            //////////    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
            //////////    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Mtl Part Num");
            //////////    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
            //////////    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Mtl Part Rev Num");
            //////////    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
            //////////    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Class");
            //////////    oSLBOMDocument.SetColumnWidth(iNumberOfRows++, iNumberOfColumns++, iStandardColumnWidth);

            //////////    foreach (JobMaterial oJobMateiral in oMaterialsNoRequiredDate)
            //////////    {
            //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, 1, oJobMateiral.JobNum);
            //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, 2, oJobMateiral.ParentPartNum);
            //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, 3, oJobMateiral.ParentRevNum);
            //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, 4, oJobMateiral.AssemblySeq.ToString());
            //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, 5, oJobMateiral.MtlSeq.ToString());
            //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, 6, oJobMateiral.MaterialPartNum);
            //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, 7, oJobMateiral.MaterialPartRevNum);
            //////////        if (oJobMateiral.PartMaster != null)
            //////////        {
            //////////            oSLBOMDocument.SetCellValue(iNumberOfRows, 8, oJobMateiral.PartMaster.ClassID);
            //////////        }

            //////////        iNumberOfRows++;
            //////////        bDataInReport = true;
            //////////    }
            //////////}

            // materials set to backflush
            //////////List<JobMaterial> oMaterialsSetToBackflush = m_oJobMaterials.Where(oItem => oItem.Backflush == true).ToList();
            //////////if (oMaterialsSetToBackflush.Count > 0)
            //////////{
            //////////    // sort these by job number
            //////////    oMaterialsSetToBackflush = oMaterialsSetToBackflush.OrderBy(oItem => oItem.JobNum).ToList();
            //////////    iNumberOfRows = 1;
            //////////    iNumberOfColumns = 1;
            //////////    if (bFirstWorksheet == true)
            //////////    {
            //////////        oSLBOMDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "Mtl Set To Backflush");
            //////////        bFirstWorksheet = false;
            //////////    }
            //////////    else
            //////////    {
            //////////        oSLBOMDocument.AddWorksheet("Mtl Set To Backflush");
            //////////    }
            //////////    //set column header
            //////////    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "JobNum");
            //////////    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
            //////////    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Parent Part Num");
            //////////    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
            //////////    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Parent Part Rev Num");
            //////////    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
            //////////    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Assembly Seq");
            //////////    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
            //////////    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Mtl Seq");
            //////////    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
            //////////    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Mtl Part Num");
            //////////    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
            //////////    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Mtl Part Rev Num");
            //////////    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
            //////////    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Class");
            //////////    oSLBOMDocument.SetColumnWidth(iNumberOfRows++, iNumberOfColumns++, iStandardColumnWidth);

            //////////    foreach (JobMaterial oJobMateiral in oMaterialsSetToBackflush)
            //////////    {
            //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, 1, oJobMateiral.JobNum);
            //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, 2, oJobMateiral.ParentPartNum);
            //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, 3, oJobMateiral.ParentRevNum);
            //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, 4, oJobMateiral.AssemblySeq.ToString());
            //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, 5, oJobMateiral.MtlSeq.ToString());
            //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, 6, oJobMateiral.MaterialPartNum);
            //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, 7, oJobMateiral.MaterialPartRevNum);
            //////////        if (oJobMateiral.PartMaster != null)
            //////////        {
            //////////            oSLBOMDocument.SetCellValue(iNumberOfRows, 8, oJobMateiral.PartMaster.ClassID);
            //////////        }

            //////////        iNumberOfRows++;
            //////////        bDataInReport = true;
            //////////    }
            //////////}

            // materials on hold

            // materials inactive

            // materials bad rev

            // materials MRP turned off

            // group code not set

            // class code not set
            #endregion

            // jobs where we issued too much material or not enough material
            List<JobMaterial> oExcessMaterialIssued = m_oJobMaterials.Where(oItem => (oItem.IssuedQty > oItem.RequiredQty) || ((oItem.IssuedQty < oItem.RequiredQty) && (oItem.QtyBearing == true) && (oItem.OpComplete == true) && ((oItem.PartMaster == null) || (m_oPartClassesToIgnore.Contains(oItem.PartMaster.ClassID) == false)))).ToList();
            if (oExcessMaterialIssued.Count > 0)
            {
                // sort these by job number
                oExcessMaterialIssued = oExcessMaterialIssued.OrderBy(oItem => oItem.JobNum).ToList();
                iNumberOfRows = 1;
                iNumberOfColumns = 1;
                if (bFirstWorksheet == true)
                {
                    oSLBOMDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "Mtl Issued");
                    bFirstWorksheet = false;
                }
                else
                {
                    oSLBOMDocument.AddWorksheet("Mtl Issued");
                }
                //set column header
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "JobNum");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Parent Part Num");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Parent Part Rev Num");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Assembly Seq");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Mtl Seq");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Mtl Part Num");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Mtl Part Rev Num");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Class");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Mtl Req");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Mtl Issued");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Variance Qty");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Variance %");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Mtl Issued Does Not Match Mtl Required");
                oSLBOMDocument.SetCellStyle(iNumberOfRows++, iNumberOfColumns, oHighlightHeaderStyle);
                oSLBOMDocument.SetColumnWidth(iNumberOfColumns++, 60);

                foreach (JobMaterial oJobMateiral in oExcessMaterialIssued)
                {
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 1, oJobMateiral.JobNum);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 2, oJobMateiral.ParentPartNum);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 3, oJobMateiral.ParentRevNum);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 4, oJobMateiral.AssemblySeq.ToString());
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 5, oJobMateiral.MtlSeq.ToString());
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 6, oJobMateiral.MaterialPartNum);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 7, oJobMateiral.MaterialPartRevNum);
                    if (oJobMateiral.PartMaster != null)
                    {
                        oSLBOMDocument.SetCellValue(iNumberOfRows, 8, oJobMateiral.PartMaster.ClassID);
                    }
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 9, oJobMateiral.RequiredQty);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 10, oJobMateiral.IssuedQty);
                    decimal dDelta = oJobMateiral.IssuedQty - oJobMateiral.RequiredQty;
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 11, dDelta);
                    oSLBOMDocument.SetCellStyle(iNumberOfRows, 11, oDecimalStyle);
                    decimal dPercentageError = 0M;
                    if (oJobMateiral.RequiredQty != 0)
                    {
                        dPercentageError = dDelta / oJobMateiral.RequiredQty * 100.0M;
                    }
                    else
                    {
                        dPercentageError = 100.0M;
                    }
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 12, dPercentageError);
                    oSLBOMDocument.SetCellStyle(iNumberOfRows, 12, oDecimalStyle);

                    iNumberOfRows++;
                    bDataInReport = true;
                }
            }

            #endregion

            #region Material Costs
            // materials without cost
            List<JobMaterial> oMaterialsWithoutCost = m_oJobMaterials.Where(oItem => (oItem.EstUnitCost == 0) && (oItem.QtyBearing == true) && ((oItem.PartMaster == null) || ((oItem.PartMaster != null) && (m_oPartClassesToIgnore.Contains(oItem.PartMaster.ClassID) == false)))).ToList();
            if (oMaterialsWithoutCost.Count > 0)
            {
                // sort these by job number
                oMaterialsWithoutCost = oMaterialsWithoutCost.OrderBy(oItem => oItem.JobNum).ToList();
                iNumberOfRows = 1;
                iNumberOfColumns = 1;
                if (bFirstWorksheet == true)
                {
                    oSLBOMDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "Mtl Cost Is Zero");
                    bFirstWorksheet = false;
                }
                else
                {
                    oSLBOMDocument.AddWorksheet("Mtl Cost Is Zero");
                }
                //set column header
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "JobNum");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Parent Part Num");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Parent Part Rev Num");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Assembly Seq");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Mtl Seq");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Mtl Part Num");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Mtl Part Rev Num");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Class");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Material Cost Is Not Set -- POC ISSUE PLEASE SET COST FOR MATERIAL");
                oSLBOMDocument.SetCellStyle(iNumberOfRows++, iNumberOfColumns, oHighlightHeaderStyle);
                oSLBOMDocument.SetColumnWidth(iNumberOfColumns++, 90);

                foreach (JobMaterial oJobMateiral in oMaterialsWithoutCost)
                {
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 1, oJobMateiral.JobNum);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 2, oJobMateiral.ParentPartNum);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 3, oJobMateiral.ParentRevNum);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 4, oJobMateiral.AssemblySeq.ToString());
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 5, oJobMateiral.MtlSeq.ToString());
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 6, oJobMateiral.MaterialPartNum);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 7, oJobMateiral.MaterialPartRevNum);
                    if (oJobMateiral.PartMaster != null)
                    {
                        oSLBOMDocument.SetCellValue(iNumberOfRows, 8, oJobMateiral.PartMaster.ClassID);
                    }

                    iNumberOfRows++;
                    bDataInReport = true;
                }
            }

            #region IGNORING THESE ISSUES
            //////////if (m_bJustMissingCosts == false)
            //////////{
            //////////    // materials with unexpected material cost
            //////////    List<JobMaterial> oMtlWithUnexpectedMtlCost = m_oJobMaterials.Where(oItem => (oItem.MaterialCost > (oItem.EstMtlUnitCost * oItem.RequiredQty)) || ((oItem.MaterialCost < (oItem.EstMtlUnitCost * oItem.RequiredQty)) && (oItem.OpComplete == true))).ToList();
            //////////    if (oMtlWithUnexpectedMtlCost.Count > 0)
            //////////    {
            //////////        // sort these by job number
            //////////        oMtlWithUnexpectedMtlCost = oMtlWithUnexpectedMtlCost.OrderBy(oItem => oItem.JobNum).ToList();
            //////////        iNumberOfRows = 1;
            //////////        iNumberOfColumns = 1;
            //////////        if (bFirstWorksheet == true)
            //////////        {
            //////////            oSLBOMDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "Mtl Unexpected Mtl Cost");
            //////////            bFirstWorksheet = false;
            //////////        }
            //////////        else
            //////////        {
            //////////            oSLBOMDocument.AddWorksheet("Mtl Unexpected Mtl Cost");
            //////////        }
            //////////        //set column header
            //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "JobNum");
            //////////        oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
            //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Parent Part Num");
            //////////        oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
            //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Parent Part Rev Num");
            //////////        oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
            //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Assembly Seq");
            //////////        oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
            //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Mtl Seq");
            //////////        oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
            //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Mtl Part Num");
            //////////        oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
            //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Mtl Part Rev Num");
            //////////        oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
            //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Class");
            //////////        oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
            //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Est Mtl Cost");
            //////////        oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
            //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Act Mtl Cost");
            //////////        oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
            //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Variance $");
            //////////        oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
            //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Variance %");
            //////////        oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
            //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Material Actual Cost Differs From Estimate -- POC ISSUE PLEASE REVIEW");
            //////////        oSLBOMDocument.SetCellStyle(iNumberOfRows++, iNumberOfColumns, oHighlightHeaderStyle);
            //////////        oSLBOMDocument.SetColumnWidth(iNumberOfColumns++, 90);

            //////////        foreach (JobMaterial oJobMaterial in oMtlWithUnexpectedMtlCost)
            //////////        {
            //////////            decimal dEstTotal = oJobMaterial.EstMtlUnitCost * oJobMaterial.RequiredQty;
            //////////            decimal dDelta = oJobMaterial.MaterialCost - dEstTotal;
            //////////            decimal dPercentageError = 0M;
            //////////            decimal dMarginToDate = 0M;
            //////////            if (dEstTotal != 0)
            //////////            {
            //////////                dPercentageError = dDelta / dEstTotal * 100.0M;
            //////////            }
            //////////            else
            //////////            {
            //////////                dPercentageError = 100.0M;
            //////////            }
            //////////            HSJob oTmpJob = m_oAllJobs.FirstOrDefault(oItem => string.Compare(oJobMaterial.JobNum, oItem.JobNum, true) == 0);
            //////////            if (oTmpJob != null)
            //////////            {
            //////////                dMarginToDate = oTmpJob.MarginToDate * 100.0M;
            //////////            }

            //////////            bool bExceedsThreholds = false;
            //////////            if ((Math.Abs(dDelta) > m_dJobEstimateAbsoluteError) && (Math.Abs(dPercentageError) > m_dJobEstimatePercentError) && (dMarginToDate < m_dJobMarginThreshold))
            //////////            {
            //////////                // check to see if it exceeds both thresholds
            //////////                bExceedsThreholds = true;
            //////////            }
            //////////            if (m_bForceJobAnalysis == true)
            //////////            {
            //////////                // force it to report
            //////////                bExceedsThreholds = true;
            //////////            }

            //////////            if (bExceedsThreholds == true)
            //////////            {

            //////////                oSLBOMDocument.SetCellValue(iNumberOfRows, 1, oJobMaterial.JobNum);
            //////////                oSLBOMDocument.SetCellValue(iNumberOfRows, 2, oJobMaterial.ParentPartNum);
            //////////                oSLBOMDocument.SetCellValue(iNumberOfRows, 3, oJobMaterial.ParentRevNum);
            //////////                oSLBOMDocument.SetCellValue(iNumberOfRows, 4, oJobMaterial.AssemblySeq.ToString());
            //////////                oSLBOMDocument.SetCellValue(iNumberOfRows, 5, oJobMaterial.MtlSeq.ToString());
            //////////                oSLBOMDocument.SetCellValue(iNumberOfRows, 6, oJobMaterial.MaterialPartNum);
            //////////                oSLBOMDocument.SetCellValue(iNumberOfRows, 7, oJobMaterial.MaterialPartRevNum);
            //////////                if (oJobMaterial.PartMaster != null)
            //////////                {
            //////////                    oSLBOMDocument.SetCellValue(iNumberOfRows, 8, oJobMaterial.PartMaster.ClassID);
            //////////                }
            //////////                oSLBOMDocument.SetCellValue(iNumberOfRows, 9, oJobMaterial.EstMtlUnitCost * oJobMaterial.RequiredQty);
            //////////                oSLBOMDocument.SetCellStyle(iNumberOfRows, 9, oCurrencyStyle);
            //////////                oSLBOMDocument.SetCellValue(iNumberOfRows, 10, oJobMaterial.MaterialCost);
            //////////                oSLBOMDocument.SetCellStyle(iNumberOfRows, 10, oCurrencyStyle);

            //////////                oSLBOMDocument.SetCellValue(iNumberOfRows, 11, dDelta);
            //////////                oSLBOMDocument.SetCellStyle(iNumberOfRows, 11, oCurrencyStyle);

            //////////                oSLBOMDocument.SetCellValue(iNumberOfRows, 12, dPercentageError);
            //////////                oSLBOMDocument.SetCellStyle(iNumberOfRows, 12, oDecimalStyle);

            //////////                iNumberOfRows++;
            //////////                bDataInReport = true;
            //////////            }
            //////////        }
            //////////    }

            //////////    // materials with unexpected burden
            //////////    List<JobMaterial> oMtlWithUnexpectedBurden = m_oJobMaterials.Where(oItem => (oItem.BurdenCost > (oItem.EstBurdenUnitCost * oItem.RequiredQty)) || ((oItem.BurdenCost < (oItem.EstBurdenUnitCost * oItem.RequiredQty)) && (oItem.OpComplete == true))).ToList();
            //////////    if (oMtlWithUnexpectedBurden.Count > 0)
            //////////    {
            //////////        // sort these by job number
            //////////        oMtlWithUnexpectedBurden = oMtlWithUnexpectedBurden.OrderBy(oItem => oItem.JobNum).ToList();
            //////////        iNumberOfRows = 1;
            //////////        iNumberOfColumns = 1;
            //////////        if (bFirstWorksheet == true)
            //////////        {
            //////////            oSLBOMDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "Mtl Unexpected Burden");
            //////////            bFirstWorksheet = false;
            //////////        }
            //////////        else
            //////////        {
            //////////            oSLBOMDocument.AddWorksheet("Mtl Unexpected Burden");
            //////////        }
            //////////        //set column header
            //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "JobNum");
            //////////        oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
            //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Parent Part Num");
            //////////        oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
            //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Parent Part Rev Num");
            //////////        oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
            //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Assembly Seq");
            //////////        oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
            //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Mtl Seq");
            //////////        oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
            //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Mtl Part Num");
            //////////        oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
            //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Mtl Part Rev Num");
            //////////        oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
            //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Class");
            //////////        oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
            //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Est Burden");
            //////////        oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
            //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Act Burden");
            //////////        oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
            //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Variance $");
            //////////        oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
            //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Variance %");
            //////////        oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
            //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Material Burden Cost Job Differs From Estimate -- POC ISSUE PLEASE REVIEW");
            //////////        oSLBOMDocument.SetCellStyle(iNumberOfRows++, iNumberOfColumns, oHighlightHeaderStyle);
            //////////        oSLBOMDocument.SetColumnWidth(iNumberOfColumns++, 90);

            //////////        foreach (JobMaterial oJobMaterial in oMtlWithUnexpectedBurden)
            //////////        {
            //////////            decimal dEstTotal = oJobMaterial.EstBurdenUnitCost * oJobMaterial.RequiredQty;
            //////////            decimal dDelta = oJobMaterial.BurdenCost - dEstTotal;
            //////////            decimal dPercentageError = 0M;
            //////////            decimal dMarginToDate = 0M;
            //////////            if (dEstTotal != 0)
            //////////            {
            //////////                dPercentageError = dDelta / dEstTotal * 100.0M;
            //////////            }
            //////////            else
            //////////            {
            //////////                dPercentageError = 100.0M;
            //////////            }
            //////////            HSJob oTmpJob = m_oAllJobs.FirstOrDefault(oItem => string.Compare(oJobMaterial.JobNum, oItem.JobNum, true) == 0);
            //////////            if (oTmpJob != null)
            //////////            {
            //////////                dMarginToDate = oTmpJob.MarginToDate * 100.0M;
            //////////            }

            //////////            bool bExceedsThreholds = false;
            //////////            if ((Math.Abs(dDelta) > m_dJobEstimateAbsoluteError) && (Math.Abs(dPercentageError) > m_dJobEstimatePercentError) && (dMarginToDate < m_dJobMarginThreshold))
            //////////            {
            //////////                // check to see if it exceeds both thresholds
            //////////                bExceedsThreholds = true;
            //////////            }
            //////////            if (m_bForceJobAnalysis == true)
            //////////            {
            //////////                // force it to report
            //////////                bExceedsThreholds = true;
            //////////            }

            //////////            if (bExceedsThreholds == true)
            //////////            {
            //////////                oSLBOMDocument.SetCellValue(iNumberOfRows, 1, oJobMaterial.JobNum);
            //////////                oSLBOMDocument.SetCellValue(iNumberOfRows, 2, oJobMaterial.ParentPartNum);
            //////////                oSLBOMDocument.SetCellValue(iNumberOfRows, 3, oJobMaterial.ParentRevNum);
            //////////                oSLBOMDocument.SetCellValue(iNumberOfRows, 4, oJobMaterial.AssemblySeq.ToString());
            //////////                oSLBOMDocument.SetCellValue(iNumberOfRows, 5, oJobMaterial.MtlSeq.ToString());
            //////////                oSLBOMDocument.SetCellValue(iNumberOfRows, 6, oJobMaterial.MaterialPartNum);
            //////////                oSLBOMDocument.SetCellValue(iNumberOfRows, 7, oJobMaterial.MaterialPartRevNum);
            //////////                if (oJobMaterial.PartMaster != null)
            //////////                {
            //////////                    oSLBOMDocument.SetCellValue(iNumberOfRows, 8, oJobMaterial.PartMaster.ClassID);
            //////////                }
            //////////                oSLBOMDocument.SetCellValue(iNumberOfRows, 9, oJobMaterial.EstBurdenUnitCost * oJobMaterial.RequiredQty);
            //////////                oSLBOMDocument.SetCellStyle(iNumberOfRows, 9, oCurrencyStyle);
            //////////                oSLBOMDocument.SetCellValue(iNumberOfRows, 10, oJobMaterial.BurdenCost);
            //////////                oSLBOMDocument.SetCellStyle(iNumberOfRows, 10, oCurrencyStyle);

            //////////                oSLBOMDocument.SetCellValue(iNumberOfRows, 11, dDelta);
            //////////                oSLBOMDocument.SetCellStyle(iNumberOfRows, 11, oCurrencyStyle);

            //////////                oSLBOMDocument.SetCellValue(iNumberOfRows, 12, dPercentageError);
            //////////                oSLBOMDocument.SetCellStyle(iNumberOfRows, 12, oDecimalStyle);

            //////////                iNumberOfRows++;
            //////////                bDataInReport = true;
            //////////            }
            //////////        }
            //////////    }

            //////////    // materials with unexpected labor
            //////////    List<JobMaterial> oMtlWithUnexpectedLabor = m_oJobMaterials.Where(oItem => (oItem.LaborCost > (oItem.EstLaborUnitCost * oItem.RequiredQty)) || ((oItem.LaborCost < (oItem.EstLaborUnitCost * oItem.RequiredQty)) && (oItem.OpComplete == true))).ToList();
            //////////    if (oMtlWithUnexpectedLabor.Count > 0)
            //////////    {
            //////////        // sort these by job number
            //////////        oMtlWithUnexpectedLabor = oMtlWithUnexpectedLabor.OrderBy(oItem => oItem.JobNum).ToList();
            //////////        iNumberOfRows = 1;
            //////////        iNumberOfColumns = 1;
            //////////        if (bFirstWorksheet == true)
            //////////        {
            //////////            oSLBOMDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "Mtl Unexpected Labor");
            //////////            bFirstWorksheet = false;
            //////////        }
            //////////        else
            //////////        {
            //////////            oSLBOMDocument.AddWorksheet("Mtl Unexpected Labor");
            //////////        }
            //////////        //set column header
            //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "JobNum");
            //////////        oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
            //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Parent Part Num");
            //////////        oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
            //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Parent Part Rev Num");
            //////////        oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
            //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Assembly Seq");
            //////////        oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
            //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Mtl Seq");
            //////////        oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
            //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Mtl Part Num");
            //////////        oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
            //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Mtl Part Rev Num");
            //////////        oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
            //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Class");
            //////////        oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
            //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Est Labor");
            //////////        oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
            //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Act Labor");
            //////////        oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
            //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Variance $");
            //////////        oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
            //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Variance %");
            //////////        oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
            //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Material Labor Cost Differs From Estimate -- POC ISSUE PLEASE REVIEW");
            //////////        oSLBOMDocument.SetCellStyle(iNumberOfRows++, iNumberOfColumns, oHighlightHeaderStyle);
            //////////        oSLBOMDocument.SetColumnWidth(iNumberOfColumns++, 90);

            //////////        foreach (JobMaterial oJobMaterial in oMtlWithUnexpectedLabor)
            //////////        {
            //////////            decimal dEstTotal = oJobMaterial.EstLaborUnitCost * oJobMaterial.RequiredQty;
            //////////            decimal dDelta = oJobMaterial.LaborCost - dEstTotal;
            //////////            decimal dPercentageError = 0M;
            //////////            decimal dMarginToDate = 0M;
            //////////            if (dEstTotal != 0)
            //////////            {
            //////////                dPercentageError = dDelta / dEstTotal * 100.0M;
            //////////            }
            //////////            else
            //////////            {
            //////////                dPercentageError = 100.0M;
            //////////            }
            //////////            HSJob oTmpJob = m_oAllJobs.FirstOrDefault(oItem => string.Compare(oJobMaterial.JobNum, oItem.JobNum, true) == 0);
            //////////            if (oTmpJob != null)
            //////////            {
            //////////                dMarginToDate = oTmpJob.MarginToDate * 100.0M;
            //////////            }

            //////////            bool bExceedsThreholds = false;
            //////////            if ((Math.Abs(dDelta) > m_dJobEstimateAbsoluteError) && (Math.Abs(dPercentageError) > m_dJobEstimatePercentError) && (dMarginToDate < m_dJobMarginThreshold))
            //////////            {
            //////////                // check to see if it exceeds both thresholds
            //////////                bExceedsThreholds = true;
            //////////            }
            //////////            if (m_bForceJobAnalysis == true)
            //////////            {
            //////////                // force it to report
            //////////                bExceedsThreholds = true;
            //////////            }

            //////////            if (bExceedsThreholds == true)
            //////////            {
            //////////                oSLBOMDocument.SetCellValue(iNumberOfRows, 1, oJobMaterial.JobNum);
            //////////                oSLBOMDocument.SetCellValue(iNumberOfRows, 2, oJobMaterial.ParentPartNum);
            //////////                oSLBOMDocument.SetCellValue(iNumberOfRows, 3, oJobMaterial.ParentRevNum);
            //////////                oSLBOMDocument.SetCellValue(iNumberOfRows, 4, oJobMaterial.AssemblySeq.ToString());
            //////////                oSLBOMDocument.SetCellValue(iNumberOfRows, 5, oJobMaterial.MtlSeq.ToString());
            //////////                oSLBOMDocument.SetCellValue(iNumberOfRows, 6, oJobMaterial.MaterialPartNum);
            //////////                oSLBOMDocument.SetCellValue(iNumberOfRows, 7, oJobMaterial.MaterialPartRevNum);
            //////////                if (oJobMaterial.PartMaster != null)
            //////////                {
            //////////                    oSLBOMDocument.SetCellValue(iNumberOfRows, 8, oJobMaterial.PartMaster.ClassID);
            //////////                }
            //////////                oSLBOMDocument.SetCellValue(iNumberOfRows, 9, oJobMaterial.EstLaborUnitCost * oJobMaterial.RequiredQty);
            //////////                oSLBOMDocument.SetCellStyle(iNumberOfRows, 9, oCurrencyStyle);
            //////////                oSLBOMDocument.SetCellValue(iNumberOfRows, 10, oJobMaterial.LaborCost);
            //////////                oSLBOMDocument.SetCellStyle(iNumberOfRows, 10, oCurrencyStyle);

            //////////                oSLBOMDocument.SetCellValue(iNumberOfRows, 11, dDelta);
            //////////                oSLBOMDocument.SetCellStyle(iNumberOfRows, 11, oCurrencyStyle);

            //////////                oSLBOMDocument.SetCellValue(iNumberOfRows, 12, dPercentageError);
            //////////                oSLBOMDocument.SetCellStyle(iNumberOfRows, 12, oDecimalStyle);

            //////////                iNumberOfRows++;
            //////////                bDataInReport = true;
            //////////            }
            //////////        }
            //////////    }

            //////////    // materials wtih unepected subcontract costs
            //////////    List<JobMaterial> oMtlWithUnexpectedSubcontractCosts = m_oJobMaterials.Where(oItem => (oItem.SubcontractCost > (oItem.EstSubcontractUnitCost * oItem.RequiredQty)) || ((oItem.SubcontractCost < (oItem.EstSubcontractUnitCost * oItem.RequiredQty)) && (oItem.OpComplete == true))).ToList();
            //////////    if (oMtlWithUnexpectedSubcontractCosts.Count > 0)
            //////////    {
            //////////        // sort these by job number
            //////////        oMtlWithUnexpectedSubcontractCosts = oMtlWithUnexpectedSubcontractCosts.OrderBy(oItem => oItem.JobNum).ToList();
            //////////        iNumberOfRows = 1;
            //////////        iNumberOfColumns = 1;
            //////////        if (bFirstWorksheet == true)
            //////////        {
            //////////            oSLBOMDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "Mtl Unexpected Subcontract");
            //////////            bFirstWorksheet = false;
            //////////        }
            //////////        else
            //////////        {
            //////////            oSLBOMDocument.AddWorksheet("Mtl Unexpected Subcontract");
            //////////        }
            //////////        //set column header
            //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "JobNum");
            //////////        oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
            //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Parent Part Num");
            //////////        oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
            //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Parent Part Rev Num");
            //////////        oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
            //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Assembly Seq");
            //////////        oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
            //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Mtl Seq");
            //////////        oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
            //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Mtl Part Num");
            //////////        oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
            //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Mtl Part Rev Num");
            //////////        oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
            //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Class");
            //////////        oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
            //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Est Subcontract");
            //////////        oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
            //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Act Subcontract");
            //////////        oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
            //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Variance $");
            //////////        oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
            //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Variance %");
            //////////        oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
            //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Material Subcontract Cost Differs From Estimate -- POC ISSUE PLEASE REVIEW");
            //////////        oSLBOMDocument.SetCellStyle(iNumberOfRows++, iNumberOfColumns, oHighlightHeaderStyle);
            //////////        oSLBOMDocument.SetColumnWidth(iNumberOfColumns++, 90);

            //////////        foreach (JobMaterial oJobMaterial in oMtlWithUnexpectedLabor)
            //////////        {
            //////////            decimal dEstTotal = oJobMaterial.EstSubcontractUnitCost * oJobMaterial.RequiredQty;
            //////////            decimal dDelta = oJobMaterial.SubcontractCost - dEstTotal;
            //////////            decimal dPercentageError = 0M;
            //////////            decimal dMarginToDate = 0M;
            //////////            if (dEstTotal != 0)
            //////////            {
            //////////                dPercentageError = dDelta / dEstTotal * 100.0M;
            //////////            }
            //////////            else
            //////////            {
            //////////                dPercentageError = 100.0M;
            //////////            }
            //////////            HSJob oTmpJob = m_oAllJobs.FirstOrDefault(oItem => string.Compare(oJobMaterial.JobNum, oItem.JobNum, true) == 0);
            //////////            if (oTmpJob != null)
            //////////            {
            //////////                dMarginToDate = oTmpJob.MarginToDate * 100.0M;
            //////////            }

            //////////            bool bExceedsThreholds = false;
            //////////            if ((Math.Abs(dDelta) > m_dJobEstimateAbsoluteError) && (Math.Abs(dPercentageError) > m_dJobEstimatePercentError) && (dMarginToDate < m_dJobMarginThreshold))
            //////////            {
            //////////                // check to see if it exceeds both thresholds
            //////////                bExceedsThreholds = true;
            //////////            }
            //////////            if (m_bForceJobAnalysis == true)
            //////////            {
            //////////                // force it to report
            //////////                bExceedsThreholds = true;
            //////////            }

            //////////            if (bExceedsThreholds == true)
            //////////            {
            //////////                oSLBOMDocument.SetCellValue(iNumberOfRows, 1, oJobMaterial.JobNum);
            //////////                oSLBOMDocument.SetCellValue(iNumberOfRows, 2, oJobMaterial.ParentPartNum);
            //////////                oSLBOMDocument.SetCellValue(iNumberOfRows, 3, oJobMaterial.ParentRevNum);
            //////////                oSLBOMDocument.SetCellValue(iNumberOfRows, 4, oJobMaterial.AssemblySeq.ToString());
            //////////                oSLBOMDocument.SetCellValue(iNumberOfRows, 5, oJobMaterial.MtlSeq.ToString());
            //////////                oSLBOMDocument.SetCellValue(iNumberOfRows, 6, oJobMaterial.MaterialPartNum);
            //////////                oSLBOMDocument.SetCellValue(iNumberOfRows, 7, oJobMaterial.MaterialPartRevNum);
            //////////                if (oJobMaterial.PartMaster != null)
            //////////                {
            //////////                    oSLBOMDocument.SetCellValue(iNumberOfRows, 8, oJobMaterial.PartMaster.ClassID);
            //////////                }
            //////////                oSLBOMDocument.SetCellValue(iNumberOfRows, 9, oJobMaterial.EstSubcontractUnitCost * oJobMaterial.RequiredQty);
            //////////                oSLBOMDocument.SetCellStyle(iNumberOfRows, 9, oCurrencyStyle);
            //////////                oSLBOMDocument.SetCellValue(iNumberOfRows, 10, oJobMaterial.SubcontractCost);
            //////////                oSLBOMDocument.SetCellStyle(iNumberOfRows, 10, oCurrencyStyle);

            //////////                oSLBOMDocument.SetCellValue(iNumberOfRows, 11, dDelta);
            //////////                oSLBOMDocument.SetCellStyle(iNumberOfRows, 11, oCurrencyStyle);

            //////////                oSLBOMDocument.SetCellValue(iNumberOfRows, 12, dPercentageError);
            //////////                oSLBOMDocument.SetCellStyle(iNumberOfRows, 12, oDecimalStyle);

            //////////                iNumberOfRows++;
            //////////                bDataInReport = true;
            //////////            }
            //////////        }
            //////////    }
            //////////}
            #endregion

            #endregion

            #region Operation Attributes
            // resource id set on operation
            // we should not specify a resource id on the operation -- too specific
            List<JobOperation> oOperationsWithResourceIdSet = m_oJobOperations.Where(oItem => string.IsNullOrEmpty(oItem.ResourceId) == false).ToList();
            if (oOperationsWithResourceIdSet.Count > 0)
            {
                // sort these by job number
                oOperationsWithResourceIdSet = oOperationsWithResourceIdSet.OrderBy(oItem => oItem.JobNum).ToList();
                iNumberOfRows = 1;
                iNumberOfColumns = 1;
                if (bFirstWorksheet == true)
                {
                    oSLBOMDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "Resource Id Set On Op");
                    bFirstWorksheet = false;
                }
                else
                {
                    oSLBOMDocument.AddWorksheet("Resource Id Set On Op");
                }
                //set column header
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "JobNum");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Parent Part Num");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Parent Part Rev Num");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Assembly Seq");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Opr Seq");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Op Code");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Resource");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Resource Is Set On Operation -- SOP IS TO ONLY SET THE RESOURCE GROUP");
                oSLBOMDocument.SetCellStyle(iNumberOfRows++, iNumberOfColumns, oHighlightHeaderStyle);
                oSLBOMDocument.SetColumnWidth(iNumberOfColumns++, 90);

                foreach (JobOperation oJobOperation in oOperationsWithResourceIdSet)
                {
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 1, oJobOperation.JobNum);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 2, oJobOperation.ParentPartNum);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 3, oJobOperation.ParentRevNum);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 4, oJobOperation.AssemblySeq.ToString());
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 5, oJobOperation.OperationSeq.ToString());
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 6, oJobOperation.OpCode);
                    oSLBOMDocument.SetCellValue(iNumberOfRows++, 7, oJobOperation.ResourceId);

                    bDataInReport = true;
                }
            }

            List<JobOperation> oOperationsWithoutDueDate = m_oJobOperations.Where(oItem => oItem.DueDate == DateTime.MinValue).ToList();
            if (oOperationsWithoutDueDate.Count > 0)
            {
                // sort these by job number
                oOperationsWithoutDueDate = oOperationsWithoutDueDate.OrderBy(oItem => oItem.JobNum).ToList();
                iNumberOfRows = 1;
                iNumberOfColumns = 1;
                if (bFirstWorksheet == true)
                {
                    oSLBOMDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "Op No Due Date");
                    bFirstWorksheet = false;
                }
                else
                {
                    oSLBOMDocument.AddWorksheet("Op No Due Date");
                }
                //set column header
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "JobNum");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Parent Part Num");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Parent Part Rev Num");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Assembly Seq");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Opr Seq");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Op Code");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Operation Due Date Not Set On Operation -- SOP IS TO SET OP DUE DATE FOR SCHEUDLING AND PLANNING");
                oSLBOMDocument.SetCellStyle(iNumberOfRows++, iNumberOfColumns, oHighlightHeaderStyle);
                oSLBOMDocument.SetColumnWidth(iNumberOfColumns++, 90);

                foreach (JobOperation oJobOperation in oOperationsWithoutDueDate)
                {
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 1, oJobOperation.JobNum);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 2, oJobOperation.ParentPartNum);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 3, oJobOperation.ParentRevNum);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 4, oJobOperation.AssemblySeq.ToString());
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 5, oJobOperation.OperationSeq.ToString());
                    oSLBOMDocument.SetCellValue(iNumberOfRows++, 6, oJobOperation.OpCode);

                    bDataInReport = true;
                }
            }

            List<JobOperation> oOperationsWithoutStartDate = m_oJobOperations.Where(oItem => oItem.StartDate == DateTime.MinValue).ToList();
            if (oOperationsWithoutStartDate.Count > 0)
            {
                // sort these by job number
                oOperationsWithoutStartDate = oOperationsWithoutStartDate.OrderBy(oItem => oItem.JobNum).ToList();
                iNumberOfRows = 1;
                iNumberOfColumns = 1;
                if (bFirstWorksheet == true)
                {
                    oSLBOMDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "Op No Start Date");
                    bFirstWorksheet = false;
                }
                else
                {
                    oSLBOMDocument.AddWorksheet("Op No Start Date");
                }
                //set column header
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "JobNum");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Parent Part Num");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Parent Part Rev Num");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Assembly Seq");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Opr Seq");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Op Code");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Operation Start Date Not Set On Operation -- SOP IS TO SET OP START DATE FOR SCHEUDLING AND PLANNING");
                oSLBOMDocument.SetCellStyle(iNumberOfRows++, iNumberOfColumns, oHighlightHeaderStyle);
                oSLBOMDocument.SetColumnWidth(iNumberOfColumns++, 90);

                foreach (JobOperation oJobOperation in oOperationsWithoutStartDate)
                {
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 1, oJobOperation.JobNum);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 2, oJobOperation.ParentPartNum);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 3, oJobOperation.ParentRevNum);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 4, oJobOperation.AssemblySeq.ToString());
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 5, oJobOperation.OperationSeq.ToString());
                    oSLBOMDocument.SetCellValue(iNumberOfRows++, 6, oJobOperation.OpCode);

                    bDataInReport = true;
                }
            }

            // check to make sure the resource group is active
            List<JobOperation> oOperationsWithResourceGroup = m_oJobOperations.Where(oItem => (string.IsNullOrEmpty(oItem.ResourceGroupId) == false)).ToList();
            List<JobOperation> oOpsWithInactiveResourceGroups = new List<JobOperation>();
            foreach (JobOperation oJobOp in oOperationsWithResourceGroup)
            {
                if (m_oBOMSupport.IsResourceGroupActive(oJobOp.ResourceGroupId) == false)
                {
                    oOpsWithInactiveResourceGroups.Add(oJobOp);
                }
            }
            if (oOpsWithInactiveResourceGroups.Count > 0)
            {
                // sort these by job number
                oOpsWithInactiveResourceGroups = oOpsWithInactiveResourceGroups.OrderBy(oItem => oItem.JobNum).ToList();
                iNumberOfRows = 1;
                iNumberOfColumns = 1;
                if (bFirstWorksheet == true)
                {
                    oSLBOMDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "Op With Inactive Res Grp");
                    bFirstWorksheet = false;
                }
                else
                {
                    oSLBOMDocument.AddWorksheet("Op With Inactive Res Grp");
                }
                //set column header
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "JobNum");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Parent Part Num");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Parent Part Rev Num");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Assembly Seq");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Opr Seq");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Op Code");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Resource Group");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Operation Uses Inactive Resource Group -- POC ISSUE THIS JOB CANNOT BE SCHEUDLED WITH INACTIVE RESOURCE GROUP");
                oSLBOMDocument.SetCellStyle(iNumberOfRows++, iNumberOfColumns, oHighlightHeaderStyle);
                oSLBOMDocument.SetColumnWidth(iNumberOfColumns++, 90);

                foreach (JobOperation oJobOperation in oOpsWithInactiveResourceGroups)
                {
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 1, oJobOperation.JobNum);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 2, oJobOperation.ParentPartNum);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 3, oJobOperation.ParentRevNum);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 4, oJobOperation.AssemblySeq.ToString());
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 5, oJobOperation.OperationSeq.ToString());
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 6, oJobOperation.OpCode);
                    oSLBOMDocument.SetCellValue(iNumberOfRows++, 7, oJobOperation.ResourceGroupId);

                    bDataInReport = true;
                }
            }

            // check to make sure the resource is active
            List<JobOperation> oOperationsWithResources = m_oJobOperations.Where(oItem => (string.IsNullOrEmpty(oItem.ResourceId) == false)).ToList();
            List<JobOperation> oOpsWithInactiveResources = new List<JobOperation>();
            foreach (JobOperation oJobOp in oOperationsWithResources)
            {
                if (m_oBOMSupport.IsResourceActive(oJobOp.ResourceId) == false)
                {
                    oOpsWithInactiveResources.Add(oJobOp);
                }
            }
            if (oOpsWithInactiveResources.Count > 0)
            {
                // sort these by job number
                oOpsWithInactiveResources = oOpsWithInactiveResources.OrderBy(oItem => oItem.JobNum).ToList();
                iNumberOfRows = 1;
                iNumberOfColumns = 1;
                if (bFirstWorksheet == true)
                {
                    oSLBOMDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "Op With Inactive Res");
                    bFirstWorksheet = false;
                }
                else
                {
                    oSLBOMDocument.AddWorksheet("Op With Inactive Res");
                }
                //set column header
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "JobNum");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Parent Part Num");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Parent Part Rev Num");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Assembly Seq");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Opr Seq");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Op Code");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Resource Grouo");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Resource");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);

                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Operation Uses Inactive Resource -- POC ISSUE THIS JOB CANNOT BE SCHEUDLED WITH INACTIVE RESOURCES");
                oSLBOMDocument.SetCellStyle(iNumberOfRows++, iNumberOfColumns, oHighlightHeaderStyle);
                oSLBOMDocument.SetColumnWidth(iNumberOfColumns++, 90);

                foreach (JobOperation oJobOperation in oOpsWithInactiveResources)
                {
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 1, oJobOperation.JobNum);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 2, oJobOperation.ParentPartNum);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 3, oJobOperation.ParentRevNum);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 4, oJobOperation.AssemblySeq.ToString());
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 5, oJobOperation.OperationSeq.ToString());
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 6, oJobOperation.OpCode);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 7, oJobOperation.ResourceGroupId);
                    oSLBOMDocument.SetCellValue(iNumberOfRows++, 8, oJobOperation.ResourceId);

                    bDataInReport = true;
                }
            }

            // prod standard is zero
            List<JobOperation> oOperationsWithZeroEstimatedTime = m_oJobOperations.Where(oItem => (oItem.ProdStandard == 0) && (oItem.Subcontract == false)).ToList();
            if (oOperationsWithZeroEstimatedTime.Count > 0)
            {
                // sort these by job number
                oOperationsWithZeroEstimatedTime = oOperationsWithZeroEstimatedTime.OrderBy(oItem => oItem.JobNum).ToList();
                iNumberOfRows = 1;
                iNumberOfColumns = 1;
                if (bFirstWorksheet == true)
                {
                    oSLBOMDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "Op No Est Time");
                    bFirstWorksheet = false;
                }
                else
                {
                    oSLBOMDocument.AddWorksheet("Op No Est Time");
                }
                //set column header
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "JobNum");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Parent Part Num");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Parent Part Rev Num");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Assembly Seq");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Opr Seq");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Op Code");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Op Has Time Requirement Set To Zero");
                oSLBOMDocument.SetCellStyle(iNumberOfRows++, iNumberOfColumns, oHighlightHeaderStyle);
                oSLBOMDocument.SetColumnWidth(iNumberOfColumns++, 60);

                foreach (JobOperation oJobOperation in oOperationsWithZeroEstimatedTime)
                {
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 1, oJobOperation.JobNum);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 2, oJobOperation.ParentPartNum);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 3, oJobOperation.ParentRevNum);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 4, oJobOperation.AssemblySeq.ToString());
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 5, oJobOperation.OperationSeq.ToString());
                    oSLBOMDocument.SetCellValue(iNumberOfRows++, 6, oJobOperation.OpCode);

                    bDataInReport = true;
                }
            }

            // test if labor rate is set to zero
            // first get list of operations with time set and a run qty
            List<JobOperation> oOperationsWithHours = m_oJobOperations.Where(oItem => (oItem.ProdStandard != 0) && (oItem.RunQty != 0) && (oItem.Subcontract == false)).ToList();
            // get a list of operations where labor cost is zero
            List<JobOpsEstVsActualCosts> oOperationsWithNoLaborCost = m_oJobOpsEstVsActualCosts.Where(oItem => oItem.EstLaborCost == 0).ToList();
            // order by job num
            oOperationsWithNoLaborCost = oOperationsWithNoLaborCost.OrderBy(oItem => oItem.JobNum).ToList();
            // walk through this list of ops with no labor cost and if they have operational time then this means the labor rate is zero
            bool bSetHeader = false;
            foreach (JobOpsEstVsActualCosts oOpCost in oOperationsWithNoLaborCost)
            {
                JobOperation oJobOperation = oOperationsWithHours.FirstOrDefault(oItem => (string.Compare(oItem.JobNum, oOpCost.JobNum, true) == 0) && (oItem.AssemblySeq == oOpCost.AssemblySeq) && (oItem.OperationSeq == oOpCost.OprSeq) && (string.Compare(oItem.OpCode, oOpCost.OpCode, true) == 0));
                if (oJobOperation != null)
                {
                    // this operation has hours but the labor comes out to $0 so that means we must have a labor rate of $0
                    if (bSetHeader == false)
                    {
                        bSetHeader = true;
                        iNumberOfRows = 1;
                        iNumberOfColumns = 1;
                        if (bFirstWorksheet == true)
                        {
                            oSLBOMDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "Ops With No Labor Rate");
                            bFirstWorksheet = false;
                        }
                        else
                        {
                            oSLBOMDocument.AddWorksheet("Ops With No Labor Rate");
                        }
                        //set column header
                        oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "JobNum");
                        oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                        oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Parent Part Num");
                        oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                        oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Parent Part Rev Num");
                        oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                        oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Assembly Seq");
                        oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                        oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Opr Seq");
                        oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                        oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Op Code");
                        oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                        oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Labor Rate Is $0 For Operation -- POC ISSUE PLEASE SET LABOR RATE ON OPERATION");
                        oSLBOMDocument.SetCellStyle(iNumberOfRows++, iNumberOfColumns, oHighlightHeaderStyle);
                        oSLBOMDocument.SetColumnWidth(iNumberOfColumns++, 90);
                    }

                    // indicate which operation has a zero dollar labor rate
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 1, oJobOperation.JobNum);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 2, oJobOperation.ParentPartNum);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 3, oJobOperation.ParentRevNum);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 4, oJobOperation.AssemblySeq.ToString());
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 5, oJobOperation.OperationSeq.ToString());
                    oSLBOMDocument.SetCellValue(iNumberOfRows++, 6, oJobOperation.OpCode);

                    bDataInReport = true;
                }
            }

            // test if burden rate is set to zero
            // we will use the list of operations with time set and a run qty from above
            // get a list of operations where burden cost is zero
            List<JobOpsEstVsActualCosts> oOperationsWithNoBurdenCost = m_oJobOpsEstVsActualCosts.Where(oItem => oItem.EstBurCost == 0).ToList();
            // walk through this list of ops with no burden cost and if they have operational time then this means the burden rate is zero
            bSetHeader = false;
            foreach (JobOpsEstVsActualCosts oOpCost in oOperationsWithNoBurdenCost)
            {
                JobOperation oJobOperation = oOperationsWithHours.FirstOrDefault(oItem => (string.Compare(oItem.JobNum, oOpCost.JobNum, true) == 0) && (oItem.AssemblySeq == oOpCost.AssemblySeq) && (oItem.OperationSeq == oOpCost.OprSeq) && (string.Compare(oItem.OpCode, oOpCost.OpCode, true) == 0));
                if (oJobOperation != null)
                {
                    // this operation has hours but the burden comes out to $0 so that means we must have a burden rate of $0
                    if (bSetHeader == false)
                    {
                        bSetHeader = true;
                        iNumberOfRows = 1;
                        iNumberOfColumns = 1;
                        if (bFirstWorksheet == true)
                        {
                            oSLBOMDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "Ops With No Burden Rate");
                            bFirstWorksheet = false;
                        }
                        else
                        {
                            oSLBOMDocument.AddWorksheet("Ops With No Burden Rate");
                        }
                        //set column header
                        oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "JobNum");
                        oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                        oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Parent Part Num");
                        oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                        oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Parent Part Rev Num");
                        oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                        oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Assembly Seq");
                        oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                        oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Opr Seq");
                        oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                        oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Op Code");
                        oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                        oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Burden Rate Is $0 For Operation -- POC ISSUE PLEASE SET BURDEN RATE ON OPERATION");
                        oSLBOMDocument.SetCellStyle(iNumberOfRows++, iNumberOfColumns, oHighlightHeaderStyle);
                        oSLBOMDocument.SetColumnWidth(iNumberOfColumns++, 90);
                    }

                    // indicate which operation has a zero dollar burden rate
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 1, oJobOperation.JobNum);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 2, oJobOperation.ParentPartNum);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 3, oJobOperation.ParentRevNum);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 4, oJobOperation.AssemblySeq.ToString());
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 5, oJobOperation.OperationSeq.ToString());
                    oSLBOMDocument.SetCellValue(iNumberOfRows++, 6, oJobOperation.OpCode);

                    bDataInReport = true;
                }
            }

            #region IGNORING THESE ISSUES
            // std format is wrong
            // standard format should be "HP" hours / piece
            //////////List<JobOperation> oOperationsWithBadStandardFormat = m_oJobOperations.Where(oItem => (string.Compare(oItem.StdFormat, "HP", true) != 0)).ToList();
            //////////if (oOperationsWithBadStandardFormat.Count > 0)
            //////////{
            //////////    // sort these by job number
            //////////    oOperationsWithBadStandardFormat = oOperationsWithBadStandardFormat.OrderBy(oItem => oItem.JobNum).ToList();
            //////////    iNumberOfRows = 1;
            //////////    iNumberOfColumns = 1;
            //////////    if (bFirstWorksheet == true)
            //////////    {
            //////////        oSLBOMDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "Op Bad Std Format");
            //////////        bFirstWorksheet = false;
            //////////    }
            //////////    else
            //////////    {
            //////////        oSLBOMDocument.AddWorksheet("Op Bad Std Format");
            //////////    }
            //////////    //set column header
            //////////    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "JobNum");
            //////////    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
            //////////    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Parent Part Num");
            //////////    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
            //////////    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Parent Part Rev Num");
            //////////    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
            //////////    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Assembly Seq");
            //////////    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
            //////////    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Opr Seq");
            //////////    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
            //////////    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Op Code");
            //////////    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
            //////////    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Std Format");
            //////////    oSLBOMDocument.SetColumnWidth(iNumberOfRows++, iNumberOfColumns++, iStandardColumnWidth);

            //////////    foreach (JobOperation oJobOperation in oOperationsWithBadStandardFormat)
            //////////    {
            //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, 1, oJobOperation.JobNum);
            //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, 2, oJobOperation.ParentPartNum);
            //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, 3, oJobOperation.ParentRevNum);
            //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, 4, oJobOperation.AssemblySeq.ToString());
            //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, 5, oJobOperation.OperationSeq.ToString());
            //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, 6, oJobOperation.OpCode);
            //////////        oSLBOMDocument.SetCellValue(iNumberOfRows++, 7, oJobOperation.StdFormat);

            //////////        bDataInReport = true;
            //////////    }
            //////////}
            #endregion

            // if std format is "OM" or "OH" then the OpsPerPart field must be > 0
            List<JobOperation> oOperationsPerPartSetToZero = m_oJobOperations.Where(oItem => ((string.Compare(oItem.StdFormat, "OM", true) == 0) || (string.Compare(oItem.StdFormat, "OH", true) == 0)) && (oItem.OperationsPerPart == 0)).ToList();
            if (oOperationsPerPartSetToZero.Count > 0)
            {
                // sort these by job number
                oOperationsPerPartSetToZero = oOperationsPerPartSetToZero.OrderBy(oItem => oItem.JobNum).ToList();
                iNumberOfRows = 1;
                iNumberOfColumns = 1;
                if (bFirstWorksheet == true)
                {
                    oSLBOMDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "Ops Per Part 0");
                    bFirstWorksheet = false;
                }
                else
                {
                    oSLBOMDocument.AddWorksheet("Ops Per Part 0");
                }
                //set column header
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "JobNum");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Parent Part Num");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Parent Part Rev Num");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Assembly Seq");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Opr Seq");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Op Code");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Std Format");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Ops Per Part Should Be > Zero When Std Format Is OM or OH");
                oSLBOMDocument.SetCellStyle(iNumberOfRows++, iNumberOfColumns, oHighlightHeaderStyle);
                oSLBOMDocument.SetColumnWidth(iNumberOfColumns++, 90);

                foreach (JobOperation oJobOperation in oOperationsPerPartSetToZero)
                {
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 1, oJobOperation.JobNum);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 2, oJobOperation.ParentPartNum);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 3, oJobOperation.ParentRevNum);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 4, oJobOperation.AssemblySeq.ToString());
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 5, oJobOperation.OperationSeq.ToString());
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 6, oJobOperation.OpCode);
                    oSLBOMDocument.SetCellValue(iNumberOfRows++, 7, oJobOperation.StdFormat);

                    bDataInReport = true;
                }
            }

            // if std format is anything other than "OM" or "OH" then the OpsPerPart field should be zero
            List<JobOperation> oOperationsPerPartNotSetToZero = m_oJobOperations.Where(oItem => (string.Compare(oItem.StdFormat, "OM", true) != 0) && (string.Compare(oItem.StdFormat, "OH", true) != 0) && (oItem.OperationsPerPart != 0)).ToList();
            if (oOperationsPerPartNotSetToZero.Count > 0)
            {
                // sort these by job number
                oOperationsPerPartNotSetToZero = oOperationsPerPartNotSetToZero.OrderBy(oItem => oItem.JobNum).ToList();
                iNumberOfRows = 1;
                iNumberOfColumns = 1;
                if (bFirstWorksheet == true)
                {
                    oSLBOMDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "Ops Per Part Not Zero");
                    bFirstWorksheet = false;
                }
                else
                {
                    oSLBOMDocument.AddWorksheet("Ops Per Part Not Zero");
                }
                //set column header
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "JobNum");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Parent Part Num");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Parent Part Rev Num");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Assembly Seq");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Opr Seq");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Op Code");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Std Format");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Ops Per Part");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Op Std Format Is Not OM or OH So Ops Per Part Should Be Zero");
                oSLBOMDocument.SetCellStyle(iNumberOfRows++, iNumberOfColumns, oHighlightHeaderStyle);
                oSLBOMDocument.SetColumnWidth(iNumberOfColumns++, 90);

                foreach (JobOperation oJobOperation in oOperationsPerPartNotSetToZero)
                {
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 1, oJobOperation.JobNum);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 2, oJobOperation.ParentPartNum);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 3, oJobOperation.ParentRevNum);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 4, oJobOperation.AssemblySeq.ToString());
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 5, oJobOperation.OperationSeq.ToString());
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 6, oJobOperation.OpCode);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 7, oJobOperation.StdFormat);
                    oSLBOMDocument.SetCellValue(iNumberOfRows++, 8, oJobOperation.OperationsPerPart);

                    bDataInReport = true;
                }
            }

            // std basis is wrong -- ignore capabilites and subcontracts
            List<JobOperation> oOperationsStdBasisWrong = m_oJobOperations.Where(oItem => (string.Compare(oItem.StdBasis, "E", true) != 0) && (oItem.Subcontract == false) && (string.IsNullOrEmpty(oItem.CapabilityId) == true)).ToList();
            if (oOperationsStdBasisWrong.Count > 0)
            {
                // sort these by job number
                oOperationsStdBasisWrong = oOperationsStdBasisWrong.OrderBy(oItem => oItem.JobNum).ToList();
                iNumberOfRows = 1;
                iNumberOfColumns = 1;
                if (bFirstWorksheet == true)
                {
                    oSLBOMDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "Op Bad Std Basis");
                    bFirstWorksheet = false;
                }
                else
                {
                    oSLBOMDocument.AddWorksheet("Op Bad Std Basis");
                }
                //set column header
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "JobNum");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Parent Part Num");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Parent Part Rev Num");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Assembly Seq");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Opr Seq");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Op Code");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Std Basis");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Ops Std Basis Should Be Set To Each");
                oSLBOMDocument.SetCellStyle(iNumberOfRows++, iNumberOfColumns, oHighlightHeaderStyle);
                oSLBOMDocument.SetColumnWidth(iNumberOfColumns++, 60);

                foreach (JobOperation oJobOperation in oOperationsStdBasisWrong)
                {
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 1, oJobOperation.JobNum);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 2, oJobOperation.ParentPartNum);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 3, oJobOperation.ParentRevNum);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 4, oJobOperation.AssemblySeq.ToString());
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 5, oJobOperation.OperationSeq.ToString());
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 6, oJobOperation.OpCode);
                    oSLBOMDocument.SetCellValue(iNumberOfRows++, 7, oJobOperation.StdBasis);

                    bDataInReport = true;
                }
            }

            // labor entry method not correct
            List<JobOperation> oOperationsBadLaborEntryMethod = m_oJobOperations.Where(oItem => (string.Compare(oItem.LaborEntryMethod, "T", true) != 0)).ToList();
            if (oOperationsBadLaborEntryMethod.Count > 0)
            {
                // sort these by job number
                oOperationsBadLaborEntryMethod = oOperationsBadLaborEntryMethod.OrderBy(oItem => oItem.JobNum).ToList();
                iNumberOfRows = 1;
                iNumberOfColumns = 1;
                if (bFirstWorksheet == true)
                {
                    oSLBOMDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "Op Bad Labor Entry");
                    bFirstWorksheet = false;
                }
                else
                {
                    oSLBOMDocument.AddWorksheet("Op Bad Labor Entry");
                }
                //set column header
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "JobNum");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Parent Part Num");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Parent Part Rev Num");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Assembly Seq");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Opr Seq");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Op Code");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Labor Entry");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Labor Entry Should Be Set To Time And Qty");
                oSLBOMDocument.SetCellStyle(iNumberOfRows++, iNumberOfColumns, oHighlightHeaderStyle);
                oSLBOMDocument.SetColumnWidth(iNumberOfColumns++, 60);

                foreach (JobOperation oJobOperation in oOperationsBadLaborEntryMethod)
                {
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 1, oJobOperation.JobNum);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 2, oJobOperation.ParentPartNum);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 3, oJobOperation.ParentRevNum);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 4, oJobOperation.AssemblySeq.ToString());
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 5, oJobOperation.OperationSeq.ToString());
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 6, oJobOperation.OpCode);
                    oSLBOMDocument.SetCellValue(iNumberOfRows++, 7, oJobOperation.LaborEntryMethod);

                    bDataInReport = true;
                }
            }

            #region IGNORING THESE ISSUES
            // if this is a subcontract we should include the days out
            //////////List<JobOperation> oSubcontractOperationNoDaysOutSet = m_oJobOperations.Where(oItem => (oItem.Subcontract == true) && (oItem.DaysOut == 0)).ToList();
            //////////if (oSubcontractOperationNoDaysOutSet.Count > 0)
            //////////{
            //////////    // sort these by job number
            //////////    oSubcontractOperationNoDaysOutSet = oSubcontractOperationNoDaysOutSet.OrderBy(oItem => oItem.JobNum).ToList();
            //////////    iNumberOfRows = 1;
            //////////    iNumberOfColumns = 1;
            //////////    if (bFirstWorksheet == true)
            //////////    {
            //////////        oSLBOMDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "Sub Op No Days Out Set");
            //////////        bFirstWorksheet = false;
            //////////    }
            //////////    else
            //////////    {
            //////////        oSLBOMDocument.AddWorksheet("Sub Op No Days Out Set");
            //////////    }
            //////////    //set column header
            //////////    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "JobNum");
            //////////    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
            //////////    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Parent Part Num");
            //////////    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
            //////////    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Parent Part Rev Num");
            //////////    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
            //////////    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Assembly Seq");
            //////////    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
            //////////    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Opr Seq");
            //////////    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
            //////////    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Op Code");
            //////////    oSLBOMDocument.SetColumnWidth(iNumberOfRows++, iNumberOfColumns++, iStandardColumnWidth);

            //////////    foreach (JobOperation oJobOperation in oOperationsBadLaborEntryMethod)
            //////////    {
            //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, 1, oJobOperation.JobNum);
            //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, 2, oJobOperation.ParentPartNum);
            //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, 3, oJobOperation.ParentRevNum);
            //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, 4, oJobOperation.AssemblySeq.ToString());
            //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, 5, oJobOperation.OperationSeq.ToString());
            //////////        oSLBOMDocument.SetCellValue(iNumberOfRows++, 6, oJobOperation.OpCode);

            //////////        bDataInReport = true;
            //////////    }
            //////////}

            // sub op with no vendor set
            //////////List<JobOperation> oSubcontractOperationNoVendor = m_oJobOperations.Where(oItem => (oItem.Subcontract == true) && (oItem.VendorNum == 0)).ToList();
            //////////if (oSubcontractOperationNoVendor.Count > 0)
            //////////{
            //////////    // sort these by job number
            //////////    oSubcontractOperationNoVendor = oSubcontractOperationNoVendor.OrderBy(oItem => oItem.JobNum).ToList();
            //////////    iNumberOfRows = 1;
            //////////    iNumberOfColumns = 1;
            //////////    if (bFirstWorksheet == true)
            //////////    {
            //////////        oSLBOMDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "Sub Op No Vendor");
            //////////        bFirstWorksheet = false;
            //////////    }
            //////////    else
            //////////    {
            //////////        oSLBOMDocument.AddWorksheet("Sub Op No Vendor");
            //////////    }
            //////////    //set column header
            //////////    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "JobNum");
            //////////    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
            //////////    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Parent Part Num");
            //////////    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
            //////////    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Parent Part Rev Num");
            //////////    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
            //////////    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Assembly Seq");
            //////////    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
            //////////    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Opr Seq");
            //////////    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
            //////////    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Op Code");
            //////////    oSLBOMDocument.SetColumnWidth(iNumberOfRows++, iNumberOfColumns++, iStandardColumnWidth);

            //////////    foreach (JobOperation oJobOperation in oSubcontractOperationNoVendor)
            //////////    {
            //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, 1, oJobOperation.JobNum);
            //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, 2, oJobOperation.ParentPartNum);
            //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, 3, oJobOperation.ParentRevNum);
            //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, 4, oJobOperation.AssemblySeq.ToString());
            //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, 5, oJobOperation.OperationSeq.ToString());
            //////////        oSLBOMDocument.SetCellValue(iNumberOfRows++, 6, oJobOperation.OpCode);

            //////////        bDataInReport = true;
            //////////    }
            //////////}
            #endregion

            // qty per should be set for subcontract
            List<JobOperation> oSubcontractOperationZeroQtyPer = m_oJobOperations.Where(oItem => (oItem.Subcontract == true) && (oItem.QtyPer == 0)).ToList();
            if (oSubcontractOperationZeroQtyPer.Count > 0)
            {
                // sort these by job number
                oSubcontractOperationZeroQtyPer = oSubcontractOperationZeroQtyPer.OrderBy(oItem => oItem.JobNum).ToList();
                iNumberOfRows = 1;
                iNumberOfColumns = 1;
                if (bFirstWorksheet == true)
                {
                    oSLBOMDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "Sub Op Zero Qty Per");
                    bFirstWorksheet = false;
                }
                else
                {
                    oSLBOMDocument.AddWorksheet("Sub Op Zero Qty Per");
                }
                //set column header
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "JobNum");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Parent Part Num");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Parent Part Rev Num");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Assembly Seq");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Opr Seq");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Op Code");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Subcontract Has Qty Set To Zero");
                oSLBOMDocument.SetCellStyle(iNumberOfRows++, iNumberOfColumns, oHighlightHeaderStyle);
                oSLBOMDocument.SetColumnWidth(iNumberOfColumns++, 60);

                foreach (JobOperation oJobOperation in oSubcontractOperationZeroQtyPer)
                {
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 1, oJobOperation.JobNum);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 2, oJobOperation.ParentPartNum);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 3, oJobOperation.ParentRevNum);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 4, oJobOperation.AssemblySeq.ToString());
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 5, oJobOperation.OperationSeq.ToString());
                    oSLBOMDocument.SetCellValue(iNumberOfRows++, 6, oJobOperation.OpCode);

                    bDataInReport = true;
                }
            }

            // subcontract estimate is zero
            List<JobOperation> oSubOperationZeroEst = m_oJobOperations.Where(oItem => (oItem.Subcontract == true) && (oItem.EstUnitCost == 0)).ToList();
            if (oSubOperationZeroEst.Count > 0)
            {
                // sort these by job number
                oSubOperationZeroEst = oSubOperationZeroEst.OrderBy(oItem => oItem.JobNum).ToList();
                iNumberOfRows = 1;
                iNumberOfColumns = 1;
                if (bFirstWorksheet == true)
                {
                    oSLBOMDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "Sub Op Zero Est");
                    bFirstWorksheet = false;
                }
                else
                {
                    oSLBOMDocument.AddWorksheet("Sub Op Zero Est");
                }
                //set column header
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "JobNum");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Parent Part Num");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Parent Part Rev Num");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Assembly Seq");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Opr Seq");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Op Code");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Subcontract Estimate Is Zero");
                oSLBOMDocument.SetCellStyle(iNumberOfRows++, iNumberOfColumns, oHighlightHeaderStyle);
                oSLBOMDocument.SetColumnWidth(iNumberOfColumns++, 60);

                foreach (JobOperation oJobOperation in oSubOperationZeroEst)
                {
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 1, oJobOperation.JobNum);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 2, oJobOperation.ParentPartNum);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 3, oJobOperation.ParentRevNum);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 4, oJobOperation.AssemblySeq.ToString());
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 5, oJobOperation.OperationSeq.ToString());
                    oSLBOMDocument.SetCellValue(iNumberOfRows++, 6, oJobOperation.OpCode);

                    bDataInReport = true;
                }
            }

            #region IGNORING THESE ISSUES
            // burden does not equals labor
            //////////List<JobOperation> oOperationWithBurdenNotEqualToLabor = m_oJobOperations.Where(oItem => (oItem.BurdenEqualsLabor == false)).ToList();
            //////////if (oOperationWithBurdenNotEqualToLabor.Count > 0)
            //////////{
            //////////    // sort these by job number
            //////////    oOperationWithBurdenNotEqualToLabor = oOperationWithBurdenNotEqualToLabor.OrderBy(oItem => oItem.JobNum).ToList();
            //////////    iNumberOfRows = 1;
            //////////    iNumberOfColumns = 1;
            //////////    if (bFirstWorksheet == true)
            //////////    {
            //////////        oSLBOMDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "Burden Not Eq Labor");
            //////////        bFirstWorksheet = false;
            //////////    }
            //////////    else
            //////////    {
            //////////        oSLBOMDocument.AddWorksheet("Burden Not Eq Labor");
            //////////    }
            //////////    //set column header
            //////////    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "JobNum");
            //////////    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
            //////////    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Parent Part Num");
            //////////    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
            //////////    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Parent Part Rev Num");
            //////////    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
            //////////    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Assembly Seq");
            //////////    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
            //////////    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Opr Seq");
            //////////    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
            //////////    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Op Code");
            //////////    oSLBOMDocument.SetColumnWidth(iNumberOfRows++, iNumberOfColumns++, iStandardColumnWidth);

            //////////    foreach (JobOperation oJobOperation in oOperationWithBurdenNotEqualToLabor)
            //////////    {
            //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, 1, oJobOperation.JobNum);
            //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, 2, oJobOperation.ParentPartNum);
            //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, 3, oJobOperation.ParentRevNum);
            //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, 4, oJobOperation.AssemblySeq.ToString());
            //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, 5, oJobOperation.OperationSeq.ToString());
            //////////        oSLBOMDocument.SetCellValue(iNumberOfRows++, 6, oJobOperation.OpCode);

            //////////        bDataInReport = true;
            //////////    }
            //////////}
            #endregion

            #region IGNORING THESE ISSUES

            //////////// split burden
            //////////List<JobOperation> oSplitBurdenSet = m_oJobOperations.Where(oItem => (oItem.SplitBurden == true)).ToList();
            //////////if (oSplitOperationsSet.Count > 0)
            //////////{
            //////////    // sort these by job number
            //////////    oSplitBurdenSet = oSplitBurdenSet.OrderBy(oItem => oItem.JobNum).ToList();
            //////////    iNumberOfRows = 1;
            //////////    iNumberOfColumns = 1;
            //////////    if (bFirstWorksheet == true)
            //////////    {
            //////////        oSLBOMDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "Split Burden Set");
            //////////        bFirstWorksheet = false;
            //////////    }
            //////////    else
            //////////    {
            //////////        oSLBOMDocument.AddWorksheet("Split Burden Set");
            //////////    }
            //////////    //set column header
            //////////    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "JobNum");
            //////////    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
            //////////    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Parent Part Num");
            //////////    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
            //////////    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Parent Part Rev Num");
            //////////    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
            //////////    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Assembly Seq");
            //////////    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
            //////////    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Opr Seq");
            //////////    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
            //////////    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Op Code");
            //////////    oSLBOMDocument.SetColumnWidth(iNumberOfRows++, iNumberOfColumns++, iStandardColumnWidth);

            //////////    foreach (JobOperation oJobOperation in oSplitBurdenSet)
            //////////    {
            //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, 1, oJobOperation.JobNum);
            //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, 2, oJobOperation.ParentPartNum);
            //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, 3, oJobOperation.ParentRevNum);
            //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, 4, oJobOperation.AssemblySeq.ToString());
            //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, 5, oJobOperation.OperationSeq.ToString());
            //////////        oSLBOMDocument.SetCellValue(iNumberOfRows++, 6, oJobOperation.OpCode);

            //////////        bDataInReport = true;
            //////////    }
            //////////}


            // std format is wrong
            // standard format should be "HP" hours / piece
            //////////List<JobOperation> oOperationsWithBadStandardFormat = m_oJobOperations.Where(oItem => (string.Compare(oItem.StdFormat, "HP", true) != 0)).ToList();
            //////////if (oOperationsWithBadStandardFormat.Count > 0)
            //////////{
            //////////    // sort these by job number
            //////////    oOperationsWithBadStandardFormat = oOperationsWithBadStandardFormat.OrderBy(oItem => oItem.JobNum).ToList();
            //////////    iNumberOfRows = 1;
            //////////    iNumberOfColumns = 1;
            //////////    if (bFirstWorksheet == true)
            //////////    {
            //////////        oSLBOMDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "Op Bad Std Format");
            //////////        bFirstWorksheet = false;
            //////////    }
            //////////    else
            //////////    {
            //////////        oSLBOMDocument.AddWorksheet("Op Bad Std Format");
            //////////    }
            //////////    //set column header
            //////////    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "JobNum");
            //////////    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
            //////////    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Parent Part Num");
            //////////    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
            //////////    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Parent Part Rev Num");
            //////////    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
            //////////    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Assembly Seq");
            //////////    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
            //////////    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Opr Seq");
            //////////    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
            //////////    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Op Code");
            //////////    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
            //////////    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Std Format");
            //////////    oSLBOMDocument.SetColumnWidth(iNumberOfRows++, iNumberOfColumns++, iStandardColumnWidth);

            //////////    foreach (JobOperation oJobOperation in oOperationsWithBadStandardFormat)
            //////////    {
            //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, 1, oJobOperation.JobNum);
            //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, 2, oJobOperation.ParentPartNum);
            //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, 3, oJobOperation.ParentRevNum);
            //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, 4, oJobOperation.AssemblySeq.ToString());
            //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, 5, oJobOperation.OperationSeq.ToString());
            //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, 6, oJobOperation.OpCode);
            //////////        oSLBOMDocument.SetCellValue(iNumberOfRows++, 7, oJobOperation.StdFormat);

            //////////        bDataInReport = true;
            //////////    }
            //////////}


            // burden does not equals labor
            //////////List<JobOperation> oOperationWithBurdenNotEqualToLabor = m_oJobOperations.Where(oItem => (oItem.BurdenEqualsLabor == false)).ToList();
            //////////if (oOperationWithBurdenNotEqualToLabor.Count > 0)
            //////////{
            //////////    // sort these by job number
            //////////    oOperationWithBurdenNotEqualToLabor = oOperationWithBurdenNotEqualToLabor.OrderBy(oItem => oItem.JobNum).ToList();
            //////////    iNumberOfRows = 1;
            //////////    iNumberOfColumns = 1;
            //////////    if (bFirstWorksheet == true)
            //////////    {
            //////////        oSLBOMDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "Burden Not Eq Labor");
            //////////        bFirstWorksheet = false;
            //////////    }
            //////////    else
            //////////    {
            //////////        oSLBOMDocument.AddWorksheet("Burden Not Eq Labor");
            //////////    }
            //////////    //set column header
            //////////    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "JobNum");
            //////////    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
            //////////    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Parent Part Num");
            //////////    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
            //////////    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Parent Part Rev Num");
            //////////    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
            //////////    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Assembly Seq");
            //////////    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
            //////////    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Opr Seq");
            //////////    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
            //////////    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Op Code");
            //////////    oSLBOMDocument.SetColumnWidth(iNumberOfRows++, iNumberOfColumns++, iStandardColumnWidth);

            //////////    foreach (JobOperation oJobOperation in oOperationWithBurdenNotEqualToLabor)
            //////////    {
            //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, 1, oJobOperation.JobNum);
            //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, 2, oJobOperation.ParentPartNum);
            //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, 3, oJobOperation.ParentRevNum);
            //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, 4, oJobOperation.AssemblySeq.ToString());
            //////////        oSLBOMDocument.SetCellValue(iNumberOfRows, 5, oJobOperation.OperationSeq.ToString());
            //////////        oSLBOMDocument.SetCellValue(iNumberOfRows++, 6, oJobOperation.OpCode);

            //////////        bDataInReport = true;
            //////////    }
            //////////}
            #endregion

            // use estimates
            List<JobOperation> oOperationUseEstimatesSet = m_oJobOperations.Where(oItem => (oItem.UseEstimates == true)).ToList();
            if (oOperationUseEstimatesSet.Count > 0)
            {
                // sort these by job number
                oOperationUseEstimatesSet = oOperationUseEstimatesSet.OrderBy(oItem => oItem.JobNum).ToList();
                iNumberOfRows = 1;
                iNumberOfColumns = 1;
                if (bFirstWorksheet == true)
                {
                    oSLBOMDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "Use Estimates Set");
                    bFirstWorksheet = false;
                }
                else
                {
                    oSLBOMDocument.AddWorksheet("Use Estimates Set");
                }
                //set column header
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "JobNum");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Parent Part Num");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Parent Part Rev Num");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Assembly Seq");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Opr Seq");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Op Code");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Use Estimates Is Set On Resource Group For Op");
                oSLBOMDocument.SetCellStyle(iNumberOfRows++, iNumberOfColumns, oHighlightHeaderStyle);
                oSLBOMDocument.SetColumnWidth(iNumberOfColumns++, 60);

                foreach (JobOperation oJobOperation in oOperationUseEstimatesSet)
                {
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 1, oJobOperation.JobNum);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 2, oJobOperation.ParentPartNum);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 3, oJobOperation.ParentRevNum);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 4, oJobOperation.AssemblySeq.ToString());
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 5, oJobOperation.OperationSeq.ToString());
                    oSLBOMDocument.SetCellValue(iNumberOfRows++, 6, oJobOperation.OpCode);

                    bDataInReport = true;
                }
            }

            // split operations
            List<JobOperation> oSplitOperationsSet = m_oJobOperations.Where(oItem => (oItem.SplitOperations == true)).ToList();
            if (oSplitOperationsSet.Count > 0)
            {
                // sort these by job number
                oSplitOperationsSet = oSplitOperationsSet.OrderBy(oItem => oItem.JobNum).ToList();
                iNumberOfRows = 1;
                iNumberOfColumns = 1;
                if (bFirstWorksheet == true)
                {
                    oSLBOMDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "Split Operations Set");
                    bFirstWorksheet = false;
                }
                else
                {
                    oSLBOMDocument.AddWorksheet("Split Operations Set");
                }
                //set column header
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "JobNum");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Parent Part Num");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Parent Part Rev Num");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Assembly Seq");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Opr Seq");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Op Code");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Op Allows Work To Be Split");
                oSLBOMDocument.SetCellStyle(iNumberOfRows++, iNumberOfColumns, oHighlightHeaderStyle);
                oSLBOMDocument.SetColumnWidth(iNumberOfColumns++, 60);

                foreach (JobOperation oJobOperation in oSplitOperationsSet)
                {
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 1, oJobOperation.JobNum);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 2, oJobOperation.ParentPartNum);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 3, oJobOperation.ParentRevNum);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 4, oJobOperation.AssemblySeq.ToString());
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 5, oJobOperation.OperationSeq.ToString());
                    oSLBOMDocument.SetCellValue(iNumberOfRows++, 6, oJobOperation.OpCode);

                    bDataInReport = true;
                }
            }

            #endregion

            if (bDataInReport == true)
            {
                oSLBOMDocument.SaveAs(sDestinationFileName);
                // Check to see if we created a file and if so email it
                if (File.Exists(sDestinationFileName) == true)
                {
                    List<string> oAttachments = new List<string>();
                    oAttachments.Add(sDestinationFileName);

                    HSEmailHelper.SendEmail(oToAddresses, sCompany + " Job Validation Report", sCompany + " Job Validation Report for " + sDate, oAttachments);
                }
            }
        }

        #endregion

        #region Properties
        #endregion

        #region Data Members

        private HSValidateParts m_oValidateParts = new HSValidateParts();
        private BOMSupport m_oBOMSupport;

        private List<JobMaterial> m_oJobMaterials = new List<JobMaterial>();
        private List<JobOperation> m_oJobOperations = new List<JobOperation>();
        private List<JobOpsEstVsActualCosts> m_oJobOpsEstVsActualCosts = new List<JobOpsEstVsActualCosts>();
        private List<JobEstVsActualCostsQty> m_oJobEstVsActualCostsQtys = new List<JobEstVsActualCostsQty>();

        Dictionary<string, List<JobMaterial>> m_oFastJobMaterials = new Dictionary<string, List<JobMaterial>>();
        Dictionary<string, List<JobOperation>> m_oFastJobOperations = new Dictionary<string, List<JobOperation>>();
        Dictionary<string, List<JobOpsEstVsActualCosts>> m_oFastJobOperationCosts = new Dictionary<string, List<JobOpsEstVsActualCosts>>();
        Dictionary<string, JobEstVsActualCostsQty> m_oFastJobEstVsActualCosts = new Dictionary<string, JobEstVsActualCostsQty>();

        private string m_sCompany;

        private List<string> m_oJobNums;
        private List<HSJob> m_oAllJobs = new List<HSJob>();

        private List<string> m_oPartClassesToIgnore = new List<string>();

        #endregion
    }
}
