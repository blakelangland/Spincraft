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
//using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Drawing.Spreadsheet;
using SpreadsheetLight;
using SpreadsheetLight.Charts;
using Ice.Lib.Searches;
using System.IO;


namespace HorizonScientific
{
    public class JobOpsEstVsActualCosts
    {
        #region Constructors
        public JobOpsEstVsActualCosts(HSJob oJob, HSAssembly oParentAssembly)
        {
            m_sCompany = oParentAssembly.Company;
            m_sJobNum = oJob.JobNum;
            m_iOrderNum = oJob.OrderNum;
            m_iOrderLine = oJob.OrderLine;
            m_iOrderRel = oJob.OrderRel;
            m_sWarehouseCode = oJob.WarehouseCode;
            m_sPartNum = oParentAssembly.PartNum;
            m_sPartRevNum = oParentAssembly.PartRevNum;
            m_iAssemblySeq = oParentAssembly.AssemblySeq;
            if (oParentAssembly.ParentAssembly != null)
            {
                m_iParentAssemblySeq = oParentAssembly.ParentAssembly.AssemblySeq;
            }
            m_iOprSeq = 0;
            m_sOpCode = "No Op";
            m_bNoOp = true;
            // a no op is always complete
            m_bOpComplete = true;
            m_dEstimatedRemainingCost = 0M;
        }

        public JobOpsEstVsActualCosts(DataRow oRow, bool bAccepActualsForMissingEstimates)
        {
            if (oRow["JobOper_Company"] != DBNull.Value)
            {
                m_sCompany = (string)oRow["JobOper_Company"];
            }
            if (oRow["JobOper_JobNum"] != DBNull.Value)
            {
                m_sJobNum = (string)oRow["JobOper_JobNum"];
            }
            if (oRow["JobAsmbl1_PartNum"] != DBNull.Value)
            {
                m_sPartNum = (string)oRow["JobAsmbl1_PartNum"];
            }
            if (oRow["JobAsmbl1_RevisionNum"] != DBNull.Value)
            {
                m_sPartRevNum = (string)oRow["JobAsmbl1_RevisionNum"];
            }
            if (oRow["JobOper_AssemblySeq"] != DBNull.Value)
            {
                m_iAssemblySeq = (int)oRow["JobOper_AssemblySeq"];
            }
            if (oRow["JobAsmbl1_Parent"] != DBNull.Value)
            {
                m_iParentAssemblySeq = (int)oRow["JobAsmbl1_Parent"];
            }
            if (oRow["JobOper_OprSeq"] != DBNull.Value)
            {
                m_iOprSeq = (int)oRow["JobOper_OprSeq"];
            }
            if (oRow["JobOper_OpCode"] != DBNull.Value)
            {
                m_sOpCode = (string)oRow["JobOper_OpCode"];
            }
            if (oRow["JobOper_OpComplete"] != DBNull.Value)
            {
                m_bOpComplete = (bool)oRow["JobOper_OpComplete"];
            }
            if (oRow["JobOper_RunQty"] != DBNull.Value)
            {
                m_dRunQty = (decimal)oRow["JobOper_RunQty"];
            }
            if (oRow["JobOper_QtyCompleted"] != DBNull.Value)
            {
                m_dCompletedQty = (decimal)oRow["JobOper_QtyCompleted"];
            }
            if (oRow["Calculated_EstOprHours"] != DBNull.Value)
            {
                m_dEstOprHours = (decimal)oRow["Calculated_EstOprHours"];
            }
            if (oRow["Calculated_ActOprHours"] != DBNull.Value)
            {
                m_dActOprHours = (decimal)oRow["Calculated_ActOprHours"];
            }
            if (oRow["Calculated_EstLbrCostFromOprs"] != DBNull.Value)
            {
                m_dEstLaborCost = (decimal)oRow["Calculated_EstLbrCostFromOprs"];
            }
            if (oRow["Calculated_ActLbrCostFromOprs"] != DBNull.Value)
            {
                m_dActLaborCost = (decimal)oRow["Calculated_ActLbrCostFromOprs"];
            }
            if (oRow["Calculated_EstBurCostFromOprs"] != DBNull.Value)
            {
                m_dEstBurCost = (decimal)oRow["Calculated_EstBurCostFromOprs"];
            }
            if (oRow["Calculated_ActBurCostFromOprs"] != DBNull.Value)
            {
                m_dActBurdenCost = (decimal)oRow["Calculated_ActBurCostFromOprs"];
            }
            if (oRow["Calculated_EstSubCostFromOprs"] != DBNull.Value)
            {
                m_dEstSubCost = (decimal)oRow["Calculated_EstSubCostFromOprs"];
            }
            if (oRow["Calculated_ActSubCostFromOprs"] != DBNull.Value)
            {
                m_dActSubCost = (decimal)oRow["Calculated_ActSubCostFromOprs"];
            }
            if (oRow["Calculated_EstCostFromOprs"] != DBNull.Value)
            {
                m_dEstCost = (decimal)oRow["Calculated_EstCostFromOprs"];
            }
            if (oRow["Calculated_ActCostFromOprs"] != DBNull.Value)
            {
                m_dActCost = (decimal)oRow["Calculated_ActCostFromOprs"];
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
                m_iOrderRel = (int)oRow["JobProd_OrderRelNum"];
            }
            if (oRow["JobProd_WarehouseCode"] != DBNull.Value)
            {
                m_sWarehouseCode = (string)oRow["JobProd_WarehouseCode"];
            }
            if (oRow["JobProd_ProdQty"] != DBNull.Value)
            {
                m_dProductionQty = (decimal)oRow["JobProd_ProdQty"];
            }
            if (oRow["JobProd_ShippedQty"] != DBNull.Value)
            {
                m_dShippedQty = (decimal)oRow["JobProd_ShippedQty"];
            }
            if (oRow["JobProd_ReceivedQty"] != DBNull.Value)
            {
                m_dReceivedQty = (decimal)oRow["JobProd_ReceivedQty"];
            }
            if (oRow["JobOper_StartDate"] != DBNull.Value)
            {
                m_dtStartDate = (DateTime)oRow["JobOper_StartDate"];
            }
            if (oRow["JobOper_DueDate"] != DBNull.Value)
            {
                m_dtDueDate = (DateTime)oRow["JobOper_DueDate"];
            }

            // set percent complete and remaining estimated cost
            if ( (RunQty <= 0) || (CompletedQty > RunQty) || (OpComplete == true) )
            {
                m_dPercentComplete = 1.0M;
            }
            else
            {
                m_dPercentComplete = CompletedQty / RunQty;
            }

            // if the user wants to utilize actuals where we have zero dollar estimates then we need to adjust
            if (bAccepActualsForMissingEstimates == true)
            {
                if ( (m_dEstCost == 0) && (m_dActCost != 0) && (m_dPercentComplete != 0) )
                {
                    m_bUsedActualsForMissingEstimate = true;

                    m_dEstOprHours = m_dActOprHours / m_dPercentComplete;
                    m_dEstLaborCost = m_dActLaborCost / m_dPercentComplete;
                    m_dEstBurCost = m_dActBurdenCost / m_dPercentComplete;
                    m_dEstSubCost = m_dActSubCost / m_dPercentComplete;
                    m_dEstCost = m_dActCost / m_dPercentComplete;
                }
            }

            m_dEstimatedRemainingCost = (1 - m_dPercentComplete) * EstCost;
        }
        #endregion

        #region Methods
        #endregion

        #region Properties
        public string Company
        {
            get { return m_sCompany; }
            set { m_sCompany = value; }
        }
        public string JobNum
        {
            get { return m_sJobNum; }
            set { m_sJobNum = value; }
        }
        public string PartNum
        {
            get { return m_sPartNum; }
            set { m_sPartNum = value; }
        }
        public string PartRevNum
        {
            get { return m_sPartRevNum; }
            set { m_sPartRevNum = value; }
        }
        public int AssemblySeq
        {
            get { return m_iAssemblySeq; }
            set { m_iAssemblySeq = value; }
        }
        public int ParentAssemblySeq
        {
            get { return m_iParentAssemblySeq; }
            set { m_iParentAssemblySeq = value; }
        }
        public int OprSeq
        {
            get { return m_iOprSeq; }
            set { m_iOprSeq = value; }
        }
        public string OpCode
        {
            get { return m_sOpCode; }
            set { m_sOpCode = value; }
        }
        public bool OpComplete
        {
            get { return m_bOpComplete; }
            set { m_bOpComplete = value; }
        }
        public decimal RunQty
        {
            get { return m_dRunQty; }
            set { m_dRunQty = value; }
        }
        public decimal CompletedQty
        {
            get { return m_dCompletedQty; }
            set { m_dCompletedQty = value; }
        }

        public decimal EstOprHours
        {
            get { return m_dEstOprHours; }
            set { m_dEstOprHours = value; }
        }
        public decimal ActOprHours
        {
            get { return m_dActOprHours; }
            set { m_dActOprHours = value; }
        }
        public decimal EstLaborCost
        {
            get { return m_dEstLaborCost; }
            set { m_dEstLaborCost = value; }
        }
        public decimal ActLaborCost
        {
            get { return m_dActLaborCost; }
            set { m_dActLaborCost = value; }
        }
        public decimal EstBurCost
        {
            get { return m_dEstBurCost; }
            set { m_dEstBurCost = value; }
        }
        public decimal ActBurdenCost
        {
            get { return m_dActBurdenCost; }
            set { m_dActBurdenCost = value; }
        }
        public decimal EstSubCost
        {
            get { return m_dEstSubCost; }
            set { m_dEstSubCost = value; }
        }
        public decimal ActSubCost
        {
            get { return m_dActSubCost; }
            set { m_dActSubCost = value; }
        }
        public decimal EstCost
        {
            get { return m_dEstCost; }
            set { m_dEstCost = value; }
        }
        public decimal ActCost
        {
            get { return m_dActCost; }
            set { m_dActCost = value; }
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
        public string WarehouseCode
        {
            get { return m_sWarehouseCode; }
            set { m_sWarehouseCode = value; }
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

        public DateTime StartDate
        {
            get { return m_dtStartDate; }
            set { m_dtStartDate = value; }
        }
        public DateTime DueDate
        {
            get { return m_dtDueDate; }
            set { m_dtDueDate = value; }
        }

        public bool NoOp
        {
            get { return m_bNoOp; }
        }
        public decimal EstimatedRemainingCost
        {
            get { return m_dEstimatedRemainingCost;  }
            set { m_dEstimatedRemainingCost = value; }
        }
        public decimal PercentComplete
        {
            get { return m_dPercentComplete; }
            set { m_dPercentComplete = value; }
        }
        public bool UsedActualsForMissingEstimate
        {
            get { return m_bUsedActualsForMissingEstimate; }
            set { m_bUsedActualsForMissingEstimate = value; }
        }
        #endregion

        #region Data Members
        private string m_sCompany;
        private string m_sJobNum;
        private string m_sPartNum;
        private string m_sPartRevNum;
        private int m_iAssemblySeq;
        private int m_iParentAssemblySeq;
        private int m_iOprSeq;
        private string m_sOpCode;
        private bool m_bOpComplete;
        private decimal m_dRunQty;
        private decimal m_dCompletedQty;
        private decimal m_dEstOprHours;
        private decimal m_dActOprHours;
        private decimal m_dEstLaborCost;
        private decimal m_dActLaborCost;
        private decimal m_dEstBurCost;
        private decimal m_dActBurdenCost;
        private decimal m_dEstSubCost;
        private decimal m_dActSubCost;
        private decimal m_dEstCost;
        private decimal m_dActCost;
        private int m_iOrderNum;
        private int m_iOrderLine;
        private int m_iOrderRel;
        private string m_sWarehouseCode;
        private decimal m_dProductionQty;
        private decimal m_dShippedQty;
        private decimal m_dReceivedQty;
        private DateTime m_dtStartDate;
        private DateTime m_dtDueDate;

        private bool m_bNoOp;
        private decimal m_dEstimatedRemainingCost;
        private decimal m_dPercentComplete;
        private bool m_bUsedActualsForMissingEstimate;
        #endregion
    }
}
