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
    public class JobMaterial
    {
        #region Constructors

        public JobMaterial(DataRow oRow, HSValidateParts oAllParts, bool bAcceptActualsForMissingEstimates)
        {
            if (oRow["JobMtl_Company"] != DBNull.Value)
            {
                m_sCompany = (string)oRow["JobMtl_Company"];
            }
            if (oRow["JobMtl_JobNum"] != DBNull.Value)
            {
                m_sJobNum = (string)oRow["JobMtl_JobNum"];
            }
            if (oRow["JobHead_PartNum"] != DBNull.Value)
            {
                m_sParentPartNum = (string)oRow["JobHead_PartNum"];
            }
            if (oRow["JobHead_RevisionNum"] != DBNull.Value)
            {
                m_sParentRevNum = (string)oRow["JobHead_RevisionNum"];
            }
            if (oRow["JobAsmbl_Parent"] != DBNull.Value)
            {
                m_iParentAssemblySeq = (int)oRow["JobAsmbl_Parent"];
            }
            if (oRow["JobMtl_AssemblySeq"] != DBNull.Value)
            {
                m_iAssemblySeq = (int)oRow["JobMtl_AssemblySeq"];
            }
            if (oRow["JobMtl_MtlSeq"] != DBNull.Value)
            {
                m_iMtlSeq = (int)oRow["JobMtl_MtlSeq"];
            }
            if (oRow["JobOper_OprSeq"] != DBNull.Value)
            {
                m_iRelatedOperation = (int)oRow["JobOper_OprSeq"];
            }
            if (oRow["JobOper_OpCode"] != DBNull.Value)
            {
                m_sOpCode = (string)oRow["JobOper_OpCode"];
            }
            if (oRow["JobOper_OpComplete"] != DBNull.Value)
            {
                m_bOpComplete = (bool)oRow["JobOper_OpComplete"];
            }
            if (oRow["JobMtl_PartNum"] != DBNull.Value)
            {
                m_sMaterialPartNum = (string)oRow["JobMtl_PartNum"];
            }
            if (oRow["JobMtl_RevisionNum"] != DBNull.Value)
            {
                m_sMaterialPartRevNum = (string)oRow["JobMtl_RevisionNum"];
            }
            if (oRow["JobMtl_IUM"] != DBNull.Value)
            {
                m_sUOMCode = (string)oRow["JobMtl_IUM"];
            }
            if (oRow["JobMtl_LeadTime"] != DBNull.Value)
            {
                m_iLeadTime = (int)oRow["JobMtl_LeadTime"];
            }
            if (oRow["JobMtl_FixedQty"] != DBNull.Value)
            {
                m_bFixedQty = (bool)oRow["JobMtl_FixedQty"];
            }
            if (oRow["JobMtl_ReqDate"] != DBNull.Value)
            {
                m_dRequiredDate = (DateTime)oRow["JobMtl_ReqDate"];
            }
            if (oRow["JobMtl_Direct"] != DBNull.Value)
            {
                m_bMakeDirect = (bool)oRow["JobMtl_Direct"];
            }
            if (oRow["JobMtl_VendorNum"] != DBNull.Value)
            {
                m_iVendorNum = (int)oRow["JobMtl_VendorNum"];
            }
            if (oRow["JobMtl_BuyIt"] != DBNull.Value)
            {
                m_bBuyIt = (bool)oRow["JobMtl_BuyIt"];
            }
            if (oRow["JobMtl_Ordered"] != DBNull.Value)
            {
                m_bOrdered = (bool)oRow["JobMtl_Ordered"];
            }
            if (oRow["JobMtl_BackFlush"] != DBNull.Value)
            {
                m_bBackflush = (bool)oRow["JobMtl_BackFlush"];
            }
            if (oRow["JobMtl_EstScrap"] != DBNull.Value)
            {
                m_dEstScrap = (decimal)oRow["JobMtl_EstScrap"];
            }
            if (oRow["JobMtl_EstScrapType"] != DBNull.Value)
            {
                m_sScrapType = (string)oRow["JobMtl_EstScrapType"];
            }
            if (oRow["JobMtl_EstMtlBurUnitCost"] != DBNull.Value)
            {
                m_dEstMaterialBurdenUnitCost = (decimal)oRow["JobMtl_EstMtlBurUnitCost"];
            }
            if (oRow["JobMtl_MtlBurRate"] != DBNull.Value)
            {
                m_dMaterialBurdenRate = (decimal)oRow["JobMtl_MtlBurRate"];
            }
            if (oRow["JobMtl_AddedMtl"] != DBNull.Value)
            {
                m_bAddedMaterial = (bool)oRow["JobMtl_AddedMtl"];
            }
            if (oRow["JobMtl_QtyPer"] != DBNull.Value)
            {
                m_dQuantityPer = (decimal)oRow["JobMtl_QtyPer"];
            }
            if (oRow["JobMtl_RequiredQty"] != DBNull.Value)
            {
                m_dRequiredQty = (decimal)oRow["JobMtl_RequiredQty"];
            }
            if (oRow["JobMtl_IssuedQty"] != DBNull.Value)
            {
                m_dIssuedQty = (decimal)oRow["JobMtl_IssuedQty"];
            }
            if (oRow["JobMtl_EstUnitCost"] != DBNull.Value)
            {
                m_dEstUnitCost = (decimal)oRow["JobMtl_EstUnitCost"];
            }
            if (oRow["JobMtl_TotalCost"] != DBNull.Value)
            {
                m_dActTotalCost = (decimal)oRow["JobMtl_TotalCost"];
            }
            if (oRow["JobMtl_EstMtlUnitCost"] != DBNull.Value)
            {
                m_dEstMtlUnitCost = (decimal)oRow["JobMtl_EstMtlUnitCost"];
            }
            if (oRow["JobMtl_MaterialMtlCost"] != DBNull.Value)
            {
                m_dMaterialCost = (decimal)oRow["JobMtl_MaterialMtlCost"];
            }
            if (oRow["JobMtl_EstBurUnitCost"] != DBNull.Value)
            {
                m_dEstBurdenUnitCost = (decimal)oRow["JobMtl_EstBurUnitCost"];
            }
            if (oRow["JobMtl_MtlBurCost"] != DBNull.Value)
            {
                m_dBurdenCost = (decimal)oRow["JobMtl_MtlBurCost"];
            }
            if (oRow["JobMtl_EstLbrUnitCost"] != DBNull.Value)
            {
                m_dEstLaborUnitCost = (decimal)oRow["JobMtl_EstLbrUnitCost"];
            }
            if (oRow["JobMtl_MaterialLabCost"] != DBNull.Value)
            {
                m_dLaborCost = (decimal)oRow["JobMtl_MaterialLabCost"];
            }
            if (oRow["JobMtl_EstSubUnitCost"] != DBNull.Value)
            {
                m_dEstSubcontractUnitCost = (decimal)oRow["JobMtl_EstSubUnitCost"];
            }
            if (oRow["JobMtl_MaterialSubCost"] != DBNull.Value)
            {
                m_dSubcontractCost = (decimal)oRow["JobMtl_MaterialSubCost"];
            }
            if (oRow["JobMtl_IssuedComplete"] != DBNull.Value)
            {
                m_bIssuedComplete = (bool)oRow["JobMtl_IssuedComplete"];
            }
            if (oRow["JobMtl_MiscCharge"] != DBNull.Value)
            {
                m_bMiscCharge = (bool)oRow["JobMtl_MiscCharge"];
            }
            if (oRow["JobMtl_MiscCode"] != DBNull.Value)
            {
                m_sChargeCode = (string)oRow["JobMtl_MiscCode"];
            }

            // find the reference to the part in the part master table
            m_oPartMaster = oAllParts.GetPart(m_sMaterialPartNum);
            if (m_oPartMaster != null)
            {
                // we see if the part master is qty bearing
                m_bQtyBearing = m_oPartMaster.PartQtyBearing;
            }
            else
            {
                // since this is a part on the fly we assume it is qty bearing
                m_bQtyBearing = true;
            }

            // set total est cost
            m_dEstTotalCost = EstUnitCost * RequiredQty;
            //Calculated_EstTotalCost Northfield added this to the query

            // set the remaining estimate cost
            if ( (m_dRequiredQty == 0) || (m_dRequiredQty - m_dIssuedQty <= 0) || (m_bIssuedComplete == true) )
            {
                m_dPercentComplete = 1.0M;
            }
            else
            {
                m_dPercentComplete = m_dIssuedQty / m_dRequiredQty;
            }

            // if the user wants to utilize actuals where we have zero dollar estimates then we need to adjust
            if (bAcceptActualsForMissingEstimates == true)
            {
                if ((m_dEstUnitCost == 0) && (m_dActTotalCost != 0) && (m_dPercentComplete != 0))
                {
                    m_bUsedActualsForMissingEstimate = true;

                    m_dEstMtlUnitCost = (m_dMaterialCost / m_dPercentComplete) / m_dRequiredQty;
                    m_dEstBurdenUnitCost = m_dBurdenCost / m_dPercentComplete / m_dRequiredQty;
                    m_dEstLaborUnitCost = m_dLaborCost / m_dPercentComplete / m_dRequiredQty;
                    m_dEstSubcontractUnitCost = m_dSubcontractCost / m_dPercentComplete / m_dRequiredQty;
                    m_dEstUnitCost = m_dActTotalCost / m_dPercentComplete / m_dRequiredQty;
                    m_dEstTotalCost = m_dActTotalCost / m_dPercentComplete;
                }
            }

            m_dEstRemainingCost = (1 - m_dPercentComplete) * m_dEstTotalCost;
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
        public string ParentPartNum
        {
            get { return m_sParentPartNum; }
            set { m_sParentPartNum = value; }
        }
        public string ParentRevNum
        {
            get { return m_sParentRevNum; }
            set { m_sParentRevNum = value; }
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
        public int MtlSeq
        {
            get { return m_iMtlSeq; }
            set { m_iMtlSeq = value; }
        }
        public int RelatedOperation
        {
            get { return m_iRelatedOperation; }
            set { m_iRelatedOperation = value; }
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
        public string MaterialPartNum
        {
            get { return m_sMaterialPartNum; }
            set { m_sMaterialPartNum = value; }
        }
        public string MaterialPartRevNum
        {
            get { return m_sMaterialPartRevNum; }
            set { m_sMaterialPartRevNum = value; }
        }
        public string UOMCode
        {
            get { return m_sUOMCode; }
            set { m_sUOMCode = value; }
        }
        public int LeadTime
        {
            get { return m_iLeadTime; }
            set { m_iLeadTime = value; }
        }
        public bool FixedQty
        {
            get { return m_bFixedQty; }
            set { m_bFixedQty = value; }
        }
        public DateTime RequiredDate
        {
            get { return m_dRequiredDate; }
            set { m_dRequiredDate = value; }
        }
        public bool MakeDirect
        {
            get { return m_bMakeDirect; }
            set { m_bMakeDirect = value; }
        }
        public int VendorNum
        {
            get { return m_iVendorNum; }
            set { m_iVendorNum = value; }
        }
        public bool BuyIt
        {
            get { return m_bBuyIt; }
            set { m_bBuyIt = value; }
        }
        public bool Ordered
        {
            get { return m_bOrdered; }
            set { m_bOrdered = value; }
        }
        public bool Backflush
        {
            get { return m_bBackflush; }
            set { m_bBackflush = value; }
        }
        public decimal EstScrap
        {
            get { return m_dEstScrap; }
            set { m_dEstScrap = value; }
        }
        public string ScrapType
        {
            get { return m_sScrapType; }
            set { m_sScrapType = value; }
        }
        public decimal EstMaterialBurdenUnitCost
        {
            get { return m_dEstMaterialBurdenUnitCost; }
            set { m_dEstMaterialBurdenUnitCost = value; }
        }
        public decimal MaterialBurdenRate
        {
            get { return m_dMaterialBurdenRate; }
            set { m_dMaterialBurdenRate = value; }
        }
        public bool AddedMaterial
        {
            get { return m_bAddedMaterial; }
            set { m_bAddedMaterial = value; }
        }
        public decimal QuantityPer
        {
            get { return m_dQuantityPer; }
            set { m_dQuantityPer = value; }
        }
        public decimal RequiredQty
        {
            get { return m_dRequiredQty; }
            set { m_dRequiredQty = value; }
        }
        public decimal IssuedQty
        {
            get { return m_dIssuedQty; }
            set { m_dIssuedQty = value; }
        }
        public decimal EstUnitCost
        {
            get { return m_dEstUnitCost; }
            set { m_dEstUnitCost = value; }
        }
        public decimal ActTotalCost
        {
            get { return m_dActTotalCost; }
            set { m_dActTotalCost = value; }
        }
        public decimal EstMtlUnitCost
        {
            get { return m_dEstMtlUnitCost; }
            set { m_dEstMtlUnitCost = value; }
        }
        public decimal MaterialCost
        {
            get { return m_dMaterialCost; }
            set { m_dMaterialCost = value; }
        }
        public decimal EstBurdenUnitCost
        {
            get { return m_dEstBurdenUnitCost; }
            set { m_dEstBurdenUnitCost = value; }
        }
        public decimal BurdenCost
        {
            get { return m_dBurdenCost; }
            set { m_dBurdenCost = value; }
        }
        public decimal EstLaborUnitCost
        {
            get { return m_dEstLaborUnitCost; }
            set { m_dEstLaborUnitCost = value; }
        }
        public decimal LaborCost
        {
            get { return m_dLaborCost; }
            set { m_dLaborCost = value; }
        }
        public decimal EstSubcontractUnitCost
        {
            get { return m_dEstSubcontractUnitCost; }
            set { m_dEstSubcontractUnitCost = value; }
        }
        public decimal SubcontractCost
        {
            get { return m_dSubcontractCost; }
            set { m_dSubcontractCost = value; }
        }
        public bool IssuedComplete
        {
            get { return m_bIssuedComplete; }
            set { m_bIssuedComplete = value; }
        }
        public bool MiscCharge
        {
            get { return m_bMiscCharge; }
            set { m_bMiscCharge = value; }
        }
        public string ChargeCode
        {
            get { return m_sChargeCode; }
            set { m_sChargeCode = value; }
        }


        public HSPartData PartMaster
        {
            get { return m_oPartMaster; }
            set { m_oPartMaster = value; }
        }

        public decimal EstTotalCost
        {
            get { return m_dEstTotalCost; }
            set { m_dEstTotalCost = value; }
        }
        public decimal EstRemainingCost
        {
            get { return m_dEstRemainingCost; }
            set { m_dEstRemainingCost = value; }
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
        public bool QtyBearing
        {
            get { return m_bQtyBearing; }
            set { m_bQtyBearing = value; }
        }
        #endregion

        #region Data Members

        private string m_sCompany;
        private string m_sJobNum;
        private string m_sParentPartNum;
        private string m_sParentRevNum;
        private int m_iAssemblySeq;
        private int m_iParentAssemblySeq;
        private int m_iMtlSeq;
        private int m_iRelatedOperation;
        private string m_sOpCode;
        private bool m_bOpComplete;
        private string m_sMaterialPartNum;
        private string m_sMaterialPartRevNum;
        private string m_sUOMCode;
        private int m_iLeadTime;
        private bool m_bFixedQty;
        private DateTime m_dRequiredDate;
        private bool m_bMakeDirect;
        private int m_iVendorNum;
        private bool m_bBuyIt;
        private bool m_bOrdered;
        private bool m_bBackflush;
        private decimal m_dEstScrap;
        private string m_sScrapType;
        private decimal m_dEstMaterialBurdenUnitCost;
        private decimal m_dMaterialBurdenRate;
        private bool m_bAddedMaterial;
        private decimal m_dQuantityPer;
        private decimal m_dRequiredQty;
        private decimal m_dIssuedQty;
        private decimal m_dEstUnitCost;
        private decimal m_dActTotalCost;
        private decimal m_dEstMtlUnitCost;
        private decimal m_dMaterialCost;
        private decimal m_dEstBurdenUnitCost;
        private decimal m_dBurdenCost;
        private decimal m_dEstLaborUnitCost;
        private decimal m_dLaborCost;
        private decimal m_dEstSubcontractUnitCost;
        private decimal m_dSubcontractCost;
        private bool m_bIssuedComplete;
        private bool m_bMiscCharge;
        private string m_sChargeCode;

        private HSPartData m_oPartMaster;
        private decimal m_dEstTotalCost;
        private decimal m_dEstRemainingCost;
        private decimal m_dPercentComplete;

        private bool m_bUsedActualsForMissingEstimate;
        private bool m_bQtyBearing;
        #endregion
    }
}