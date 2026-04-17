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
    public class PartMaterial
    {
        #region Constructors

        public PartMaterial(DataRow oRow, Dictionary<string, List<MfgPart>> oAllMfgParts, HSValidateParts oAllParts)
        {
            if (oRow["PartMtl_Company"] != DBNull.Value)
            {
                m_sCompany = (string)oRow["PartMtl_Company"];
            }
            if (oRow["PartMtl_PartNum"] != DBNull.Value)
            {
                m_sParentPartNum = (string)oRow["PartMtl_PartNum"];
            }
            if (oRow["PartMtl_RevisionNum"] != DBNull.Value)
            {
                m_sParentRevNum = (string)oRow["PartMtl_RevisionNum"];
            }
            if (oRow["PartRev_Approved"] != DBNull.Value)
            {
                m_bRevApproved = (bool)oRow["PartRev_Approved"];
            }
            if (oRow["PartRev_ApprovedDate"] != DBNull.Value)
            {
                m_dtApprovedDate = (DateTime)oRow["PartRev_ApprovedDate"];
            }
            if (oRow["PartRev_EffectiveDate"] != DBNull.Value)
            {
                m_dtEffectiveDate = (DateTime)oRow["PartRev_EffectiveDate"];
            }
            if (oRow["PartMtl_MtlSeq"] != DBNull.Value)
            {
                m_iMtlSeq = (int)oRow["PartMtl_MtlSeq"];
            }
            if (oRow["PartOpr_OprSeq"] != DBNull.Value)
            {
                m_iRelatedOperation = (int)oRow["PartOpr_OprSeq"];
            }
            if (oRow["PartOpr_OpCode"] != DBNull.Value)
            {
                m_sOpCode = (string)oRow["PartOpr_OpCode"];
            }
            if (oRow["OpMaster_OpDesc"] != DBNull.Value)
            {
                m_sOpDescription = (string)oRow["OpMaster_OpDesc"];
            }
            if (oRow["PartMtl_MtlPartNum"] != DBNull.Value)
            {
                m_sMaterialPartNum = (string)oRow["PartMtl_MtlPartNum"];
            }
            if (oRow["Part_PartDescription"] != DBNull.Value)
            {
                m_sDescription = (string)oRow["Part_PartDescription"];
            }
            if (oRow["Part_TypeCode"] != DBNull.Value)
            {
                m_sTypeCode = (string)oRow["Part_TypeCode"];
            }
            if (oRow["PartMtl_UOMCode"] != DBNull.Value)
            {
                m_sUOMCode = (string)oRow["PartMtl_UOMCode"];
            }
            if (oRow["PartMtl_PullAsAsm"] != DBNull.Value)
            {
                m_bPullAsAsm = (bool)oRow["PartMtl_PullAsAsm"];
            }
            if (oRow["Part_PhantomBOM"] != DBNull.Value)
            {
                m_bPhantomBOM = (bool)oRow["Part_PhantomBOM"];
            }
            if (oRow["PartMtl_ViewAsAsm"] != DBNull.Value)
            {
                m_bViewAsAsm = (bool)oRow["PartMtl_ViewAsAsm"];
            }
            if (oRow["PartMtl_PlanAsAsm"] != DBNull.Value)
            {
                m_bPlanAsAsm = (bool)oRow["PartMtl_PlanAsAsm"];
            }
            if (oRow["Part_QtyBearing"] != DBNull.Value)
            {
                m_bQtyBearing = (bool)oRow["Part_QtyBearing"];
            }
            if (oRow["PartMtl_QtyPer"] != DBNull.Value)
            {
                m_dQuantityPer = (decimal)oRow["PartMtl_QtyPer"];
            }
            if (oRow["Calculated_Cost"] != DBNull.Value)
            {
                m_dCostPerUnit = (decimal)oRow["Calculated_Cost"];
            }
            if (oRow["PartMtl_FixedQty"] != DBNull.Value)
            {
                m_bFixedQty = (bool)oRow["PartMtl_FixedQty"];
            }
            if (oRow["PartMtl_MfgComment"] != DBNull.Value)
            {
                m_sMfgComments = (string)oRow["PartMtl_MfgComment"];
            }
            if (oRow["PartMtl_OverRideMfgComments"] != DBNull.Value)
            {
                m_bOverrideMfgComments = (bool)oRow["PartMtl_OverRideMfgComments"];
            }
            if (oRow["PartMtl_PurComment"] != DBNull.Value)
            {
                m_sPurchasingComments = (string)oRow["PartMtl_PurComment"];
            }
            if (oRow["PartMtl_OverRidePurComments"] != DBNull.Value)
            {
                m_bOverridePurchasingComments = (bool)oRow["PartMtl_OverRidePurComments"];
            }
            if (oRow["PartMtl_EstScrap"] != DBNull.Value)
            {
                m_dScrapFactor = (decimal)oRow["PartMtl_EstScrap"];
            }
            if (oRow["PartMtl_EstScrapType"] != DBNull.Value)
            {
                m_sScrapType = (string)oRow["PartMtl_EstScrapType"];
            }
            if (oRow["PartMtl_MtlBurRate"] != DBNull.Value)
            {
                m_dMaterialBurdenRate = (decimal)oRow["PartMtl_MtlBurRate"];
            }
            if (oRow["PartMtl_EstMtlBurUnitCost"] != DBNull.Value)
            {
                m_dEstMaterialBurdenCost = (decimal)oRow["PartMtl_EstMtlBurUnitCost"];
            }
            if (oRow["PartMtl_AltMethod"] != DBNull.Value)
            {
                m_aAltMethod = (string)oRow["PartMtl_AltMethod"];
            }
            if (oRow["PartMtl_BaseMethodOverridden"] != DBNull.Value)
            {
                m_bBaseMethodOverriden = (bool)oRow["PartMtl_BaseMethodOverridden"];
            }
            if (oRow["PartMtl_ParentAltMethod"] != DBNull.Value)
            {
                m_sParentAltMethod = (string)oRow["PartMtl_ParentAltMethod"];
            }
            if (oRow["PartMtl_ParentMtlSeq"] != DBNull.Value)
            {
                m_iParentMtlSeq = (int)oRow["PartMtl_ParentMtlSeq"];
            }
            if (oRow["PartMtl_ReqRefDes"] != DBNull.Value)
            {
                m_iReferenceRequiredDesignators = (int)oRow["PartMtl_ReqRefDes"];
            }

            // these should be ordered old to new with respect to the effective date and then alphabetically where the effective dates are the same
            m_oPotentialMfgParts.Clear();
            if (oAllMfgParts.ContainsKey(m_sMaterialPartNum) == true)
            {
                List<MfgPart> oPotentialMfgParts = oAllMfgParts[m_sMaterialPartNum];
                m_oPotentialMfgParts = oPotentialMfgParts.OrderBy(x => x.EffectiveDate).ThenBy(x => x.RevNum).ToList();
            }

            // find the reference to the part in the part master table
            m_oPartMaster = oAllParts.AllParts.FirstOrDefault(oItem => string.Compare(oItem.PartNum, m_sMaterialPartNum, true) == 0);
        }

        #endregion

        #region Methods

        public MfgPart GetMfgPart(bool bApprovedBOM, DateTime dtEffectiveDate)
        {
            MfgPart oMfgPart = null;
            if (m_oPotentialMfgParts.Count != 0)
            {
                // get the first one in the list in case they pass in an effective date that is too early
                oMfgPart = m_oPotentialMfgParts[0];
            }
            // get the correct Mfg Part based on the effective date passed in
            // if there are multiple revs with the same effective date we will get the last one in the list
            // which is the latest rev - sorted alphanumerically after effective date
            foreach (MfgPart oTmp in m_oPotentialMfgParts)
            {
                if (oTmp.EffectiveDate <= dtEffectiveDate)
                {
                    if (bApprovedBOM == true)
                    {
                        if (oTmp.RevApproved == true)
                        {
                            oMfgPart = oTmp;
                        }
                    }
                    else
                    {
                        oMfgPart = oTmp;
                    }
                }
            }
            return oMfgPart;
        }

        public MfgPart GetMfgPart(string sRevNum)
        {
            // get the correct Parent based on the rev num passed in
            MfgPart oParent = m_oPotentialMfgParts.FirstOrDefault(x => string.Compare(x.RevNum, sRevNum, true) == 0);
            return oParent;
        }

        #endregion

        #region Properties
        public string Company
        {
            get { return m_sCompany; }
            set { m_sCompany = value; }
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
        public string OpDescription
        {
            get { return m_sOpDescription; }
            set { m_sOpDescription = value; }
        }
        public string MaterialPartNum
        {
            get { return m_sMaterialPartNum; }
            set { m_sMaterialPartNum = value; }
        }
        public string Description
        {
            get { return m_sDescription; }
            set { m_sDescription = value; }
        }
        public string TypeCode
        {
            get { return m_sTypeCode; }
            set { m_sTypeCode = value; }
        }
        public string UOMCode
        {
            get { return m_sUOMCode; }
            set { m_sUOMCode = value; }
        }
        public bool ViewAsAsm
        {
            get { return m_bViewAsAsm; }
            set { m_bViewAsAsm = value; }
        }
        public bool PullAsAsm
        {
            get { return m_bPullAsAsm; }
            set { m_bPullAsAsm = value; }
        }
        public bool PhantomBOM
        {
            get { return m_bPhantomBOM; }
            set { m_bPhantomBOM = value; }
        }
        public bool PlanAsAsm
        {
            get { return m_bPlanAsAsm; }
            set { m_bPlanAsAsm = value; }
        }

        public bool QtyBearing
        {
            get { return m_bQtyBearing; }
            set { m_bQtyBearing = value; }
        }

        public decimal QuantityPer
        {
            get { return m_dQuantityPer; }
            set { m_dQuantityPer = value; }
        }
        public decimal CostPerUnit
        {
            get { return m_dCostPerUnit; }
            set { m_dCostPerUnit = value; }
        }

        public bool FixedQty
        {
            get { return m_bFixedQty; }
            set { m_bFixedQty = value; }
        }

        public string MfgComments
        {
            get { return m_sMfgComments; }
            set { m_sMfgComments = value; }
        }

        public bool OverrideMfgComments
        {
            get { return m_bOverrideMfgComments; }
            set { m_bOverrideMfgComments = value; }
        }

        public string PurchasingComments
        {
            get { return m_sPurchasingComments; }
            set { m_sPurchasingComments = value; }
        }

        public bool OverridePurchasingComments
        {
            get { return m_bOverridePurchasingComments; }
            set { m_bOverridePurchasingComments = value; }
        }

        public decimal ScrapFactor
        {
            get { return m_dScrapFactor; }
            set { m_dScrapFactor = value; }
        }

        public string ScrapType
        {
            get { return m_sScrapType; }
            set { m_sScrapType = value; }
        }

        public decimal MaterialBurdenRate
        {
            get { return m_dMaterialBurdenRate; }
            set { m_dMaterialBurdenRate = value; }
        }

        public decimal EstMaterialBurdenCost
        {
            get { return m_dEstMaterialBurdenCost; }
            set { m_dEstMaterialBurdenCost = value; }
        }

        public string AltMethod
        {
            get { return m_aAltMethod; }
            set { m_aAltMethod = value; }
        }

        public bool BaseMethodOverriden
        {
            get { return m_bBaseMethodOverriden; }
            set { m_bBaseMethodOverriden = value; }
        }

        public string ParentAltMethod
        {
            get { return m_sParentAltMethod; }
            set { m_sParentAltMethod = value; }
        }

        public int ParentMtlSeq
        {
            get { return m_iParentMtlSeq; }
            set { m_iParentMtlSeq = value; }
        }

        public int ReferenceRequiredDesignators
        {
            get { return m_iReferenceRequiredDesignators; }
            set { m_iReferenceRequiredDesignators = value; }
        }

        public List<PartMaterial> MyPartMaterials
        {
            get { return m_oPartMaterials; }
            set { m_oPartMaterials = value; }
        }

        public HSPartData PartMaster
        {
            get { return m_oPartMaster; }
            set { m_oPartMaster = value; }
        }
        #endregion

        #region Data Members

        private string m_sCompany;
        private string m_sParentPartNum;
        private string m_sParentRevNum;
        private bool m_bRevApproved;
        private DateTime m_dtApprovedDate;
        private DateTime m_dtEffectiveDate;
        private int m_iMtlSeq;
        private int m_iRelatedOperation;
        private string m_sOpCode;
        private string m_sOpDescription;
        private string m_sMaterialPartNum;
        private string m_sDescription;
        private string m_sTypeCode;
        private string m_sUOMCode;
        private bool m_bViewAsAsm;
        private bool m_bPullAsAsm;
        private bool m_bPhantomBOM;
        private bool m_bPlanAsAsm;
        private bool m_bQtyBearing;
        private decimal m_dQuantityPer;
        private decimal m_dCostPerUnit;
        private bool m_bFixedQty;
        private string m_sMfgComments;
        private bool m_bOverrideMfgComments;
        private string m_sPurchasingComments;
        private bool m_bOverridePurchasingComments;
        private decimal m_dScrapFactor;
        private string m_sScrapType;
        private decimal m_dMaterialBurdenRate;
        private decimal m_dEstMaterialBurdenCost;
        private string m_aAltMethod;
        private bool m_bBaseMethodOverriden;
        private string m_sParentAltMethod;
        private int m_iParentMtlSeq;
        private int m_iReferenceRequiredDesignators;

        private List<PartMaterial> m_oPartMaterials = new List<PartMaterial>();
        private List<MfgPart> m_oPotentialMfgParts = new List<MfgPart>();
        private HSPartData m_oPartMaster;

        #endregion
    }
}

