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
    public class PartOperation
    {
        #region Constructors
        public PartOperation(DataRow oRow)
        {
            if (oRow["PartOpr_Company"] != DBNull.Value)
            {
                m_sCompany = (string)oRow["PartOpr_Company"];
            }
            if (oRow["PartOpr_PartNum"] != DBNull.Value)
            {
                m_sPartNum = (string)oRow["PartOpr_PartNum"];
            }
            if (oRow["PartOpr_RevisionNum"] != DBNull.Value)
            {
                m_sRevNum = (string)oRow["PartOpr_RevisionNum"];
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
            if (oRow["PartOpr_OprSeq"] != DBNull.Value)
            {
                m_iOprSeq = (int)oRow["PartOpr_OprSeq"];
            }
            if (oRow["PartOpr_OpCode"] != DBNull.Value)
            {
                m_sOpCode = (string)oRow["PartOpr_OpCode"];
            }
            if (oRow["PartOpr_OpDesc"] != DBNull.Value)
            {
                m_sOpDescription = (string)oRow["PartOpr_OpDesc"];
            }
            if (oRow["PartOpr_CommentText"] != DBNull.Value)
            {
                m_sCommentText = (string)oRow["PartOpr_CommentText"];
            }
            if (oRow["PartOpDtl_OpDtlSeq"] != DBNull.Value)
            {
                m_iOprDtlSeq = (int)oRow["PartOpDtl_OpDtlSeq"];
            }
            if (oRow["PartOpDtl_CapabilityID"] != DBNull.Value)
            {
                m_sCapabilityId = (string)oRow["PartOpDtl_CapabilityID"];
            }
            if (oRow["PartOpDtl_ResourceGrpID"] != DBNull.Value)
            {
                m_sResourceGroupId = (string)oRow["PartOpDtl_ResourceGrpID"];
            }
            if (oRow["PartOpDtl_ResourceID"] != DBNull.Value)
            {
                m_sResourceId = (string)oRow["PartOpDtl_ResourceID"];
            }
            if (oRow["PartOpDtl_SetupHours"] != DBNull.Value)
            {
                m_dOprDtlSetupHours = (decimal)oRow["PartOpDtl_SetupHours"];
            }
            if (oRow["PartOpr_EstSetHours"] != DBNull.Value)
            {
                m_dOrpEstSetupHours = (decimal)oRow["PartOpr_EstSetHours"];
            }
            if (oRow["PartOpDtl_ProdHours"] != DBNull.Value)
            {
                m_dOprDtlProdHours = (decimal)oRow["PartOpDtl_ProdHours"];
            }
            if (oRow["PartOpr_EstProdHours"] != DBNull.Value)
            {
                m_dOprEstProdHours = (decimal)oRow["PartOpr_EstProdHours"];
            }
            if (oRow["PartOpDtl_NumResources"] != DBNull.Value)
            {
                m_iNumberOfResources = (int)oRow["PartOpDtl_NumResources"];
            }
            if (oRow["PartOpDtl_ProdCrewSize"] != DBNull.Value)
            {
                m_dOprDtlProdCrewSize = (decimal)oRow["PartOpDtl_ProdCrewSize"];
            }
            if (oRow["PartOpr_ProdCrewSize"] != DBNull.Value)
            {
                m_dOprProdCrewSize = (decimal)oRow["PartOpr_ProdCrewSize"];
            }
            if (oRow["PartOpDtl_SetUpCrewSize"] != DBNull.Value)
            {
                m_dOprDtlSetupCrewSize = (decimal)oRow["PartOpDtl_SetUpCrewSize"];
            }
            if (oRow["PartOpr_SetUpCrewSize"] != DBNull.Value)
            {
                m_dOprSetupCrewSize = (decimal)oRow["PartOpr_SetUpCrewSize"];
            }
            if (oRow["PartOpr_AddlSetupHours"] != DBNull.Value)
            {
                m_dAdditonalSetupHours = (decimal)oRow["PartOpr_AddlSetupHours"];
            }
            if (oRow["PartOpr_AddlSetUpQty"] != DBNull.Value)
            {
                m_dAdditionalSetQty = (decimal)oRow["PartOpr_AddlSetUpQty"];
            }
            if (oRow["PartOpr_SetupGroup"] != DBNull.Value)
            {
                m_sSetupGroup = (string)oRow["PartOpr_SetupGroup"];
            }
            if (oRow["PartOpDtl_SetupOrProd"] != DBNull.Value)
            {
                m_sSetupOrProduction = (string)oRow["PartOpDtl_SetupOrProd"];
            }
            if (oRow["PartOpDtl_AltMethod"] != DBNull.Value)
            {
                m_sAlternateMethod = (string)oRow["PartOpDtl_AltMethod"];
            }
            if (oRow["PartOpr_ParentAltMethod"] != DBNull.Value)
            {
                m_sParentAlternateMethod = (string)oRow["PartOpr_ParentAltMethod"];
            }
            if (oRow["PartOpr_BaseMethodOverridden"] != DBNull.Value)
            {
                m_bBaseMethodOverriden = (bool)oRow["PartOpr_BaseMethodOverridden"];
            }
            if (oRow["PartOpr_ProdStandard"] != DBNull.Value)
            {
                m_dProductionStandard = (decimal)oRow["PartOpr_ProdStandard"];
            }
            if (oRow["PartOpr_StdFormat"] != DBNull.Value)
            {
                m_sStandardFormat = (string)oRow["PartOpr_StdFormat"];
            }
            if (oRow["PartOpr_LaborEntryMethod"] != DBNull.Value)
            {
                m_sLaborEntryMethod = (string)oRow["PartOpr_LaborEntryMethod"];
            }
            if (oRow["PartOpr_OpStdID"] != DBNull.Value)
            {
                m_sOpStandardId = (string)oRow["PartOpr_OpStdID"];
            }
            if (oRow["PartOpr_StdBasis"] != DBNull.Value)
            {
                m_sStandardBasis = (string)oRow["PartOpr_StdBasis"];
            }
            if (oRow["PartOpr_Machines"] != DBNull.Value)
            {
                m_iMachines = (int)oRow["PartOpr_Machines"];
            }
            if (oRow["PartOpr_OpsPerPart"] != DBNull.Value)
            {
                m_iOperationsPerPart = (int)oRow["PartOpr_OpsPerPart"];
            }
            if (oRow["PartOpr_QtyPer"] != DBNull.Value)
            {
                m_dQtyPerParent = (decimal)oRow["PartOpr_QtyPer"];
            }
            if (oRow["PartOpr_EstScrap"] != DBNull.Value)
            {
                m_dEstScrap = (decimal)oRow["PartOpr_EstScrap"];
            }
            if (oRow["PartOpr_EstScrapType"] != DBNull.Value)
            {
                m_sScrapType = (string)oRow["PartOpr_EstScrapType"];
            }
            if (oRow["PartOpr_EstUnitCost"] != DBNull.Value)
            {
                m_dEstUnitCost = (decimal)oRow["PartOpr_EstUnitCost"];
            }
            if (oRow["PartOpr_SubContract"] != DBNull.Value)
            {
                m_bSubcontract = (bool)oRow["PartOpr_SubContract"];
            }
            if (oRow["PartOpr_DaysOut"] != DBNull.Value)
            {
                m_dDaysOut = (decimal)oRow["PartOpr_DaysOut"];
            }
            if (oRow["PartOpr_VendorNum"] != DBNull.Value)
            {
                m_iVendorNum = (int)oRow["PartOpr_VendorNum"];
            }
            if (oRow["PartOpr_RunQty"] != DBNull.Value)
            {
                m_dRunQty = (decimal)oRow["PartOpr_RunQty"];
            }
            if (oRow["PartOpr_SNRequiredOpr"] != DBNull.Value)
            {
                m_bSNRequiredOnOperation = (bool)oRow["PartOpr_SNRequiredOpr"];
            }
            if (oRow["PartOpr_SchedRelation"] != DBNull.Value)
            {
                m_sScheduleRelation = (string)oRow["PartOpr_SchedRelation"];
            }
            if (oRow["PartOpDtl_ConcurrentCapacity"] != DBNull.Value)
            {
                m_dConcurrentCapacity = (decimal)oRow["PartOpDtl_ConcurrentCapacity"];
            }
            if (oRow["PartOpr_SendAheadType"] != DBNull.Value)
            {
                m_sSendAheadType = (string)oRow["PartOpr_SendAheadType"];
            }
            if (oRow["PartOpr_SendAheadOffset"] != DBNull.Value)
            {
                m_dSendAheadOffset = (decimal)oRow["PartOpr_SendAheadOffset"];
            }
            if (oRow["PartOpr_PBrkCost01"] != DBNull.Value)
            {
                m_dPriceBreakCost1 = (decimal)oRow["PartOpr_PBrkCost01"];
            }
            if (oRow["PartOpr_PBrkCost02"] != DBNull.Value)
            {
                m_dPriceBreakCost2 = (decimal)oRow["PartOpr_PBrkCost02"];
            }
            if (oRow["PartOpr_PBrkCost03"] != DBNull.Value)
            {
                m_dPriceBreakCost3 = (decimal)oRow["PartOpr_PBrkCost03"];
            }
            if (oRow["PartOpr_PBrkCost04"] != DBNull.Value)
            {
                m_dPriceBreakCost4 = (decimal)oRow["PartOpr_PBrkCost04"];
            }
            if (oRow["PartOpr_PBrkCost05"] != DBNull.Value)
            {
                m_dPriceBreakCost5 = (decimal)oRow["PartOpr_PBrkCost05"];
            }
            if (oRow["PartOpr_PBrkCost06"] != DBNull.Value)
            {
                m_dPriceBreakCost6 = (decimal)oRow["PartOpr_PBrkCost06"];
            }
            if (oRow["PartOpr_PBrkCost07"] != DBNull.Value)
            {
                m_dPriceBreakCost7 = (decimal)oRow["PartOpr_PBrkCost07"];
            }
            if (oRow["PartOpr_PBrkCost08"] != DBNull.Value)
            {
                m_dPriceBreakCost8 = (decimal)oRow["PartOpr_PBrkCost08"];
            }
            if (oRow["PartOpr_PBrkCost09"] != DBNull.Value)
            {
                m_dPriceBreakCost9 = (decimal)oRow["PartOpr_PBrkCost09"];
            }
            if (oRow["PartOpr_PBrkCost10"] != DBNull.Value)
            {
                m_dPriceBreakCost10 = (decimal)oRow["PartOpr_PBrkCost10"];
            }
            if (oRow["PartOpr_PBrkStdRate01"] != DBNull.Value)
            {
                m_dProductionStandardRate1 = (decimal)oRow["PartOpr_PBrkStdRate01"];
            }
            if (oRow["PartOpr_PBrkStdRate02"] != DBNull.Value)
            {
                m_dProductionStandardRate2 = (decimal)oRow["PartOpr_PBrkStdRate02"];
            }
            if (oRow["PartOpr_PBrkStdRate03"] != DBNull.Value)
            {
                m_dProductionStandardRate3 = (decimal)oRow["PartOpr_PBrkStdRate03"];
            }
            if (oRow["PartOpr_PBrkStdRate04"] != DBNull.Value)
            {
                m_dProductionStandardRate4 = (decimal)oRow["PartOpr_PBrkStdRate04"];
            }
            if (oRow["PartOpr_PBrkStdRate05"] != DBNull.Value)
            {
                m_dProductionStandardRate5 = (decimal)oRow["PartOpr_PBrkStdRate05"];
            }
            if (oRow["PartOpr_PBrkStdRate06"] != DBNull.Value)
            {
                m_dProductionStandardRate6 = (decimal)oRow["PartOpr_PBrkStdRate06"];
            }
            if (oRow["PartOpr_PBrkStdRate07"] != DBNull.Value)
            {
                m_dProductionStandardRate7 = (decimal)oRow["PartOpr_PBrkStdRate07"];
            }
            if (oRow["PartOpr_PBrkStdRate08"] != DBNull.Value)
            {
                m_dProductionStandardRate8 = (decimal)oRow["PartOpr_PBrkStdRate08"];
            }
            if (oRow["PartOpr_PBrkStdRate09"] != DBNull.Value)
            {
                m_dProductionStandardRate9 = (decimal)oRow["PartOpr_PBrkStdRate09"];
            }
            if (oRow["PartOpr_PBrkStdRate10"] != DBNull.Value)
            {
                m_dProductionStandardRate10 = (decimal)oRow["PartOpr_PBrkStdRate10"];
            }
            if (oRow["PartOpr_BrkQty01"] != DBNull.Value)
            {
                m_dBreakQuantity1 = (decimal)oRow["PartOpr_BrkQty01"];
            }
            if (oRow["PartOpr_BrkQty02"] != DBNull.Value)
            {
                m_dBreakQuantity2 = (decimal)oRow["PartOpr_BrkQty02"];
            }
            if (oRow["PartOpr_BrkQty03"] != DBNull.Value)
            {
                m_dBreakQuantity3 = (decimal)oRow["PartOpr_BrkQty03"];
            }
            if (oRow["PartOpr_BrkQty04"] != DBNull.Value)
            {
                m_dBreakQuantity4 = (decimal)oRow["PartOpr_BrkQty04"];
            }
            if (oRow["PartOpr_BrkQty05"] != DBNull.Value)
            {
                m_dBreakQuantity5 = (decimal)oRow["PartOpr_BrkQty05"];
            }
            if (oRow["PartOpr_BrkQty06"] != DBNull.Value)
            {
                m_dBreakQuantity6 = (decimal)oRow["PartOpr_BrkQty06"];
            }
            if (oRow["PartOpr_BrkQty07"] != DBNull.Value)
            {
                m_dBreakQuantity7 = (decimal)oRow["PartOpr_BrkQty07"];
            }
            if (oRow["PartOpr_BrkQty08"] != DBNull.Value)
            {
                m_dBreakQuantity8 = (decimal)oRow["PartOpr_BrkQty08"];
            }
            if (oRow["PartOpr_BrkQty09"] != DBNull.Value)
            {
                m_dBreakQuantity9 = (decimal)oRow["PartOpr_BrkQty09"];
            }
            if (oRow["PartOpr_BrkQty10"] != DBNull.Value)
            {
                m_dBreakQuantity10 = (decimal)oRow["PartOpr_BrkQty10"];
            }
        }

        #endregion

        #region Properties
        public string Company
        {
            get { return m_sCompany; }
            set { m_sCompany = value; }
        }

        public string PartNum
        {
            get { return m_sPartNum; }
            set { m_sPartNum = value; }
        }

        public string RevNum
        {
            get { return m_sRevNum; }
            set { m_sRevNum = value; }
        }

        public bool RevApproved
        {
            get { return m_bRevApproved; }
            set { m_bRevApproved = value; }
        }

        public DateTime ApprovedDate
        {
            get { return m_dtApprovedDate; }
            set { m_dtApprovedDate = value; }
        }
        
        public DateTime EffectiveDate
        {
            get { return m_dtEffectiveDate; }
            set { m_dtEffectiveDate = value; }
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
        
        public string OpDescription
        {
            get { return m_sOpDescription; }
            set { m_sOpDescription = value; }
        }
        
        public string CommentText
        {
            get { return m_sCommentText; }
            set { m_sCommentText = value; }
        }
        
        public int OprDtlSeq
        {
            get { return m_iOprDtlSeq; }
            set { m_iOprDtlSeq = value; }
        }
        
        public string CapabilityId
        {
            get { return m_sCapabilityId; }
            set { m_sCapabilityId = value; }
        }
        
        public string ResourceGroupId
        {
            get { return m_sResourceGroupId; }
            set { m_sResourceGroupId = value; }
        }
        
        public string ResourceId
        {
            get { return m_sResourceId; }
            set { m_sResourceId = value; }
        }
        
        public decimal OprDtlSetupHours
        {
            get { return m_dOprDtlSetupHours; }
            set { m_dOprDtlSetupHours = value; }
        }
        
        public decimal OrpEstSetupHours
        {
            get { return m_dOrpEstSetupHours; }
            set { m_dOrpEstSetupHours = value; }
        }
        
        public decimal OprDtlProdHours
        {
            get { return m_dOprDtlProdHours; }
            set { m_dOprDtlProdHours = value; }
        }
        
        public decimal OprEstProdHours
        {
            get { return m_dOprEstProdHours; }
            set { m_dOprEstProdHours = value; }
        }
        
        public int NumberOfResources
        {
            get { return m_iNumberOfResources; }
            set { m_iNumberOfResources = value; }
        }

        public decimal OprDtlProdCrewSize
        {
            get { return m_dOprDtlProdCrewSize; }
            set { m_dOprDtlProdCrewSize = value; }
        }

        public decimal OprProdCrewSize
        {
            get { return m_dOprProdCrewSize; }
            set { m_dOprProdCrewSize = value; }
        }

        public decimal OprDtlSetupCrewSize
        {
            get { return m_dOprDtlSetupCrewSize; }
            set { m_dOprDtlSetupCrewSize = value; }
        }

        public decimal OprSetupCrewSize
        {
            get { return m_dOprSetupCrewSize; }
            set { m_dOprSetupCrewSize = value; }
        }

        public decimal AdditonalSetupHours
        {
            get { return m_dAdditonalSetupHours; }
            set { m_dAdditonalSetupHours = value; }
        }

        public decimal AdditionalSetQty
        {
            get { return m_dAdditionalSetQty; }
            set { m_dAdditionalSetQty = value; }
        }

        public string SetupGroup
        {
            get { return m_sSetupGroup; }
            set { m_sSetupGroup = value; }
        }

        public string SetupOrProduction
        {
            get { return m_sSetupOrProduction; }
            set { m_sSetupOrProduction = value; }
        }

        public string AlternateMethod
        {
            get { return m_sAlternateMethod; }
            set { m_sAlternateMethod = value; }
        }

        public string ParentAlternateMethod
        {
            get { return m_sParentAlternateMethod; }
            set { m_sParentAlternateMethod = value; }
        }

        public bool BaseMethodOverriden
        {
            get { return m_bBaseMethodOverriden; }
            set { m_bBaseMethodOverriden = value; }
        }

        public decimal ProductionStandard
        {
            get { return m_dProductionStandard; }
            set { m_dProductionStandard = value; }
        }

        public string StandardFormat
        {
            get { return m_sStandardFormat; }
            set { m_sStandardFormat = value; }
        }

        public string LaborEntryMethod
        {
            get { return m_sLaborEntryMethod; }
            set { m_sLaborEntryMethod = value; }
        }

        public string OpStandardId
        {
            get { return m_sOpStandardId; }
            set { m_sOpStandardId = value; }
        }

        public string StandardBasis
        {
            get { return m_sStandardBasis; }
            set { m_sStandardBasis = value; }
        }

        public int Machines
        {
            get { return m_iMachines; }
            set { m_iMachines = value; }
        }

        public int OperationsPerPart
        {
            get { return m_iOperationsPerPart; }
            set { m_iOperationsPerPart = value; }
        }

        public decimal QtyPerParent
        {
            get { return m_dQtyPerParent; }
            set { m_dQtyPerParent = value; }

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

        public decimal EstUnitCost
        {
            get { return m_dEstUnitCost; }
            set { m_dEstUnitCost = value; }
        }

        public bool Subcontract
        {
            get { return m_bSubcontract; }
            set { m_bSubcontract = value; }
        }

        public decimal DaysOut
        {
            get { return m_dDaysOut; }
            set { m_dDaysOut = value; }
        }

        public int VendorNum
        {
            get { return m_iVendorNum; }
            set { m_iVendorNum = value; }
        }

        public decimal RunQty
        {
            get { return m_dRunQty; }
            set { m_dRunQty = value; }

        }

        public bool SNRequiredOnOperation
        {
            get { return m_bSNRequiredOnOperation; }
            set { m_bSNRequiredOnOperation = value; }
        }

        public string ScheduleRelation
        {
            get { return m_sScheduleRelation; }
            set { m_sScheduleRelation = value; }
        }

        public decimal ConcurrentCapacity
        {
            get { return m_dConcurrentCapacity; }
            set { m_dConcurrentCapacity = value; }
        }

        public string SendAheadType
        {
            get { return m_sSendAheadType; }
            set { m_sSendAheadType = value; }
        }

        public decimal SendAheadOffset
        {
            get { return m_dSendAheadOffset; }
            set { m_dSendAheadOffset = value; }

        }
        public decimal PriceBreakCost1
        {
            get { return m_dPriceBreakCost1; }
            set { m_dPriceBreakCost1 = value; }
        }
        public decimal PriceBreakCost2
        {
            get { return m_dPriceBreakCost2; }
            set { m_dPriceBreakCost2 = value; }
        }
        public decimal PriceBreakCost3
        {
            get { return m_dPriceBreakCost3; }
            set { m_dPriceBreakCost3 = value; }
        }
        public decimal PriceBreakCost4
        {
            get { return m_dPriceBreakCost4; }
            set { m_dPriceBreakCost4 = value; }
        }
        public decimal PriceBreakCost5
        {
            get { return m_dPriceBreakCost5; }
            set { m_dPriceBreakCost5 = value; }
        }
        public decimal PriceBreakCost6
        {
            get { return m_dPriceBreakCost6; }
            set { m_dPriceBreakCost6 = value; }
        }
        public decimal PriceBreakCost7
        {
            get { return m_dPriceBreakCost7; }
            set { m_dPriceBreakCost7 = value; }
        }
        public decimal PriceBreakCost8
        {
            get { return m_dPriceBreakCost8; }
            set { m_dPriceBreakCost8 = value; }
        }
        public decimal PriceBreakCost9
        {
            get { return m_dPriceBreakCost9; }
            set { m_dPriceBreakCost9 = value; }
        }

        public decimal PriceBreakCost10
        {
            get { return m_dPriceBreakCost10; }
            set { m_dPriceBreakCost10 = value; }
        }

        public decimal ProductionStandardRate1
        {
            get { return m_dProductionStandardRate1; }
            set { m_dProductionStandardRate1 = value; }
        }
        public decimal ProductionStandardRate2
        {
            get { return m_dProductionStandardRate2; }
            set { m_dProductionStandardRate2 = value; }
        }
        public decimal ProductionStandardRate3
        {
            get { return m_dProductionStandardRate3; }
            set { m_dProductionStandardRate3 = value; }
        }
        public decimal ProductionStandardRate4
        {
            get { return m_dProductionStandardRate4; }
            set { m_dProductionStandardRate4 = value; }
        }
        public decimal ProductionStandardRate5
        {
            get { return m_dProductionStandardRate5; }
            set { m_dProductionStandardRate5 = value; }
        }
        public decimal ProductionStandardRate6
        {
            get { return m_dProductionStandardRate6; }
            set { m_dProductionStandardRate6 = value; }
        }
        public decimal ProductionStandardRate7
        {
            get { return m_dProductionStandardRate7; }
            set { m_dProductionStandardRate7 = value; }
        }
        public decimal ProductionStandardRate8
        {
            get { return m_dProductionStandardRate8; }
            set { m_dProductionStandardRate8 = value; }
        }
        public decimal ProductionStandardRate9
        {
            get { return m_dProductionStandardRate9; }
            set { m_dProductionStandardRate9 = value; }
        }
        public decimal ProductionStandardRate10
        {
            get { return m_dProductionStandardRate10; }
            set { m_dProductionStandardRate10 = value; }
        }

        public decimal BreakQuantity1
        {
            get { return m_dBreakQuantity1; }
            set { m_dBreakQuantity1 = value; }
        }
        public decimal BreakQuantity2
        {
            get { return m_dBreakQuantity2; }
            set { m_dBreakQuantity2 = value; }
        }
        public decimal BreakQuantity3
        {
            get { return m_dBreakQuantity3; }
            set { m_dBreakQuantity3 = value; }
        }
        public decimal BreakQuantity4
        {
            get { return m_dBreakQuantity4; }
            set { m_dBreakQuantity4 = value; }
        }
        public decimal BreakQuantity5
        {
            get { return m_dBreakQuantity5; }
            set { m_dBreakQuantity5 = value; }
        }
        public decimal BreakQuantity6
        {
            get { return m_dBreakQuantity6; }
            set { m_dBreakQuantity6 = value; }
        }
        public decimal BreakQuantity7
        {
            get { return m_dBreakQuantity7; }
            set { m_dBreakQuantity7 = value; }
        }
        public decimal BreakQuantity8
        {
            get { return m_dBreakQuantity8; }
            set { m_dBreakQuantity8 = value; }
        }
        public decimal BreakQuantity9
        {
            get { return m_dBreakQuantity9; }
            set { m_dBreakQuantity9 = value; }
        }
        public decimal BreakQuantity10
        {
            get { return m_dBreakQuantity10; }
            set { m_dBreakQuantity10 = value; }
        }
        #endregion

        #region Data Members

        private string m_sCompany;
        private string m_sPartNum;
        private string m_sRevNum;
        private bool m_bRevApproved;
        private DateTime m_dtApprovedDate;
        private DateTime m_dtEffectiveDate;
        private int m_iOprSeq;
        private string m_sOpCode;
        private string m_sOpDescription;
        private string m_sCommentText;
        private int m_iOprDtlSeq;
        private string m_sCapabilityId;
        private string m_sResourceGroupId;
        private string m_sResourceId;
        private decimal m_dOprDtlSetupHours;
        private decimal m_dOrpEstSetupHours;
        private decimal m_dOprDtlProdHours;
        private decimal m_dOprEstProdHours;
        private int m_iNumberOfResources;
        private decimal m_dOprDtlProdCrewSize;
        private decimal m_dOprProdCrewSize;
        private decimal m_dOprDtlSetupCrewSize;
        private decimal m_dOprSetupCrewSize;
        private decimal m_dAdditonalSetupHours;
        private decimal m_dAdditionalSetQty;
        private string m_sSetupGroup;
        private string m_sSetupOrProduction;
        private string m_sAlternateMethod;
        private string m_sParentAlternateMethod;
        private bool m_bBaseMethodOverriden;
        private decimal m_dProductionStandard;
        private string m_sStandardFormat;
        private string m_sLaborEntryMethod;
        private string m_sOpStandardId;
        private string m_sStandardBasis;
        private int m_iMachines;
        private int m_iOperationsPerPart;
        private decimal m_dQtyPerParent;
        private decimal m_dEstScrap;
        private string m_sScrapType;
        private decimal m_dEstUnitCost;
        private bool m_bSubcontract;
        private decimal m_dDaysOut;
        private int m_iVendorNum;
        private decimal m_dRunQty;
        private bool m_bSNRequiredOnOperation;
        private string m_sScheduleRelation;
        private decimal m_dConcurrentCapacity;
        private string m_sSendAheadType;
        private decimal m_dSendAheadOffset;
        private decimal m_dPriceBreakCost1;
        private decimal m_dPriceBreakCost2;
        private decimal m_dPriceBreakCost3;
        private decimal m_dPriceBreakCost4;
        private decimal m_dPriceBreakCost5;
        private decimal m_dPriceBreakCost6;
        private decimal m_dPriceBreakCost7;
        private decimal m_dPriceBreakCost8;
        private decimal m_dPriceBreakCost9;
        private decimal m_dPriceBreakCost10;
        private decimal m_dProductionStandardRate1;
        private decimal m_dProductionStandardRate2;
        private decimal m_dProductionStandardRate3;
        private decimal m_dProductionStandardRate4;
        private decimal m_dProductionStandardRate5;
        private decimal m_dProductionStandardRate6;
        private decimal m_dProductionStandardRate7;
        private decimal m_dProductionStandardRate8;
        private decimal m_dProductionStandardRate9;
        private decimal m_dProductionStandardRate10;
        private decimal m_dBreakQuantity1;
        private decimal m_dBreakQuantity2;
        private decimal m_dBreakQuantity3;
        private decimal m_dBreakQuantity4;
        private decimal m_dBreakQuantity5;
        private decimal m_dBreakQuantity6;
        private decimal m_dBreakQuantity7;
        private decimal m_dBreakQuantity8;
        private decimal m_dBreakQuantity9;
        private decimal m_dBreakQuantity10;

        #endregion
    }
}
