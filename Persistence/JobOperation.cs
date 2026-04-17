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
    public class JobOperation
    {
        #region Constructors
        public JobOperation(HSJob oJob, HSAssembly oParentAssembly)
        {
            m_sCompany = oParentAssembly.Company;
            m_sJobNum = oJob.JobNum;
            m_sParentPartNum = oParentAssembly.PartNum;
            m_sParentRevNum = oParentAssembly.PartRevNum;
            m_iAssemblySeq = oParentAssembly.AssemblySeq;
            if (oParentAssembly.ParentAssembly != null)
            {
                m_iParentAssemblySeq = oParentAssembly.ParentAssembly.AssemblySeq;
            }
            m_iOperationSeq = 0;
            m_sOpCode = "No Op";
            m_bNoOp = true;

            // a no op is always complete
            m_bOpComplete = true;
        }

        public JobOperation(DataRow oRow)
        {
            if (oRow["JobOpDtl_Company"] != DBNull.Value)
            {
                m_sCompany = (string)oRow["JobOpDtl_Company"];
            }
            if (oRow["JobOpDtl_JobNum"] != DBNull.Value)
            {
                m_sJobNum = (string)oRow["JobOpDtl_JobNum"];
            }
            if (oRow["JobAsmbl_PartNum"] != DBNull.Value)
            {
                m_sParentPartNum = (string)oRow["JobAsmbl_PartNum"];
            }
            if (oRow["JobAsmbl_RevisionNum"] != DBNull.Value)
            {
                m_sParentRevNum = (string)oRow["JobAsmbl_RevisionNum"];
            }
            if (oRow["JobOpDtl_AssemblySeq"] != DBNull.Value)
            {
                m_iAssemblySeq = (int)oRow["JobOpDtl_AssemblySeq"];
            }
            if (oRow["JobAsmbl_Parent"] != DBNull.Value)
            {
                m_iParentAssemblySeq = (int)oRow["JobAsmbl_Parent"];
            }
            if (oRow["JobOpDtl_OprSeq"] != DBNull.Value)
            {
                m_iOperationSeq = (int)oRow["JobOpDtl_OprSeq"];
            }
            if (oRow["JobOper_OpCode"] != DBNull.Value)
            {
                m_sOpCode = (string)oRow["JobOper_OpCode"];
            }
            if (oRow["JobOpDtl_OpDtlSeq"] != DBNull.Value)
            {
                m_iOpDetailSeq = (int)oRow["JobOpDtl_OpDtlSeq"];
            }
            if (oRow["JobOper_OpComplete"] != DBNull.Value)
            {
                m_bOpComplete = (bool)oRow["JobOper_OpComplete"];
            }
            if (oRow["JobOpDtl_CapabilityID"] != DBNull.Value)
            {
                m_sCapabilityId = (string)oRow["JobOpDtl_CapabilityID"];
            }
            if (oRow["JobOpDtl_ResourceGrpID"] != DBNull.Value)
            {
                m_sResourceGroupId = (string)oRow["JobOpDtl_ResourceGrpID"];
            }
            if (oRow["JobOpDtl_ResourceID"] != DBNull.Value)
            {
                m_sResourceId = (string)oRow["JobOpDtl_ResourceID"];
            }
            if (oRow["JobOpDtl_EstSetHours"] != DBNull.Value)
            {
                m_dEstSetupHours = (decimal)oRow["JobOpDtl_EstSetHours"];
            }
            if (oRow["JobOpDtl_ActSetupHours"] != DBNull.Value)
            {
                m_dActSetupHours = (decimal)oRow["JobOpDtl_ActSetupHours"];
            }
            if (oRow["JobOpDtl_EstProdHours"] != DBNull.Value)
            {
                m_dEstProdHours = (decimal)oRow["JobOpDtl_EstProdHours"];
            }
            if (oRow["JobOpDtl_ActProdHours"] != DBNull.Value)
            {
                m_dActProdHours = (decimal)oRow["JobOpDtl_ActProdHours"];
            }
            if (oRow["JobOpDtl_NumResources"] != DBNull.Value)
            {
                m_iNumResources = (int)oRow["JobOpDtl_NumResources"];
            }
            if (oRow["JobOpDtl_ProdCrewSize"] != DBNull.Value)
            {
                m_dProdCrewSize = (decimal)oRow["JobOpDtl_ProdCrewSize"];
            }
            if (oRow["JobOpDtl_SetUpCrewSize"] != DBNull.Value)
            {
                m_dSetupCrewSize = (decimal)oRow["JobOpDtl_SetUpCrewSize"];
            }
            if (oRow["JobOper_SetupGroup"] != DBNull.Value)
            {
                m_sSetupGroup = (string)oRow["JobOper_SetupGroup"];
            }
            if (oRow["JobOpDtl_ProdStandard"] != DBNull.Value)
            {
                m_dProdStandard = (decimal)oRow["JobOpDtl_ProdStandard"];
            }
            if (oRow["JobOpDtl_StdFormat"] != DBNull.Value)
            {
                m_sStdFormat = (string)oRow["JobOpDtl_StdFormat"];
            }
            if (oRow["JobOper_LaborEntryMethod"] != DBNull.Value)
            {
                m_sLaborEntryMethod = (string)oRow["JobOper_LaborEntryMethod"];
            }
            if (oRow["JobOper_OpStdID"] != DBNull.Value)
            {
                m_sOpStandardId = (string)oRow["JobOper_OpStdID"];
            }
            if (oRow["JobOpDtl_StdBasis"] != DBNull.Value)
            {
                m_sStdBasis = (string)oRow["JobOpDtl_StdBasis"];
            }
            if (oRow["JobOper_Machines"] != DBNull.Value)
            {
                m_iNumberOfMachines = (int)oRow["JobOper_Machines"];
            }
            if (oRow["JobOpDtl_EstSetHoursPerMch"] != DBNull.Value)
            {
                m_dEstSetupHoursPerMachine = (decimal)oRow["JobOpDtl_EstSetHoursPerMch"];
            }
            if (oRow["JobOpDtl_OverrideRates"] != DBNull.Value)
            {
                m_bOverrideRates = (bool)oRow["JobOpDtl_OverrideRates"];
            }
            if (oRow["JobOper_OpsPerPart"] != DBNull.Value)
            {
                m_iOperationsPerPart = (int)oRow["JobOper_OpsPerPart"];
            }
            if (oRow["JobOper_QtyPer"] != DBNull.Value)
            {
                m_dQtyPer = (decimal)oRow["JobOper_QtyPer"];
            }
            if (oRow["JobOper_EstScrap"] != DBNull.Value)
            {
                m_dEstScrap = (decimal)oRow["JobOper_EstScrap"];
            }
            if (oRow["JobOper_EstScrapType"] != DBNull.Value)
            {
                m_sScrapType = (string)oRow["JobOper_EstScrapType"];
            }
            if (oRow["JobOper_EstUnitCost"] != DBNull.Value)
            {
                m_dEstUnitCost = (decimal)oRow["JobOper_EstUnitCost"];
            }
            if (oRow["JobOper_SubContract"] != DBNull.Value)
            {
                m_bSubcontract = (bool)oRow["JobOper_SubContract"];
            }
            if (oRow["JobOper_DaysOut"] != DBNull.Value)
            {
                m_dDaysOut = (decimal)oRow["JobOper_DaysOut"];
            }
            if (oRow["JobOper_VendorNum"] != DBNull.Value)
            {
                m_iVendorNum = (int)oRow["JobOper_VendorNum"];
            }
            if (oRow["JobOper_RunQty"] != DBNull.Value)
            {
                m_dRunQty = (decimal)oRow["JobOper_RunQty"];
            }
            if (oRow["JobOper_SchedRelation"] != DBNull.Value)
            {
                m_sScheduleRelation = (string)oRow["JobOper_SchedRelation"];
            }
            if (oRow["JobOpDtl_ConcurrentCapacity"] != DBNull.Value)
            {
                m_dConcurrentCapacity = (decimal)oRow["JobOpDtl_ConcurrentCapacity"];
            }
            if (oRow["JobOper_SendAheadType"] != DBNull.Value)
            {
                m_sSendAheadType = (string)oRow["JobOper_SendAheadType"];
            }
            if (oRow["JobOper_SendAheadOffset"] != DBNull.Value)
            {
                m_dSendAheadOffset = (decimal)oRow["JobOper_SendAheadOffset"];
            }
            if (oRow["ResourceGroup_BurdenEQLabor"] != DBNull.Value)
            {
                m_bBurdenEqualsLabor = (bool) oRow["ResourceGroup_BurdenEQLabor"];
            }
            if (oRow["ResourceGroup_SplitOperations"] != DBNull.Value)
            {
                m_bSplitOperations = (bool)oRow["ResourceGroup_SplitOperations"];
            }
            if (oRow["ResourceGroup_UseEstimates"] != DBNull.Value)
            {
                m_bUseEstimates = (bool)oRow["ResourceGroup_UseEstimates"];
            }
            if (oRow["ResourceGroup_SplitBurden"] != DBNull.Value)
            {
                m_bSplitBurden = (bool)oRow["ResourceGroup_SplitBurden"];
            }
            if (oRow["ResourceGroup_BurdenType"] != DBNull.Value)
            {
                m_sBurdenType = (string)oRow["ResourceGroup_BurdenType"];
            }
            if (oRow["JobOper_StartDate"] != DBNull.Value)
            {
                m_dtStartDate = (DateTime)oRow["JobOper_StartDate"];
            }
            if (oRow["JobOper_DueDate"] != DBNull.Value)
            {
                m_dtDueDate = (DateTime)oRow["JobOper_DueDate"];
            }
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
        public int OperationSeq
        {
            get { return m_iOperationSeq; }
            set { m_iOperationSeq = value; }
        }
        public string OpCode
        {
            get { return m_sOpCode; }
            set { m_sOpCode = value; }
        }
        public int OpDetailSeq
        {
            get { return m_iOpDetailSeq; }
            set { m_iOpDetailSeq = value; }
        }
        public bool OpComplete
        {
            get { return m_bOpComplete; }
            set { m_bOpComplete = value; }
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
        public decimal EstSetupHours
        {
            get { return m_dEstSetupHours; }
            set { m_dEstSetupHours = value; }
        }
        public decimal ActSetupHours
        {
            get { return m_dActSetupHours; }
            set { m_dActSetupHours = value; }
        }
        public decimal EstProdHours
        {
            get { return m_dEstProdHours; }
            set { m_dEstProdHours = value; }
        }
        public decimal ActProdHours
        {
            get { return m_dActProdHours; }
            set { m_dActProdHours = value; }
        }
        public int NumResources
        {
            get { return m_iNumResources; }
            set { m_iNumResources = value; }
        }
        public decimal ProdCrewSize
        {
            get { return m_dProdCrewSize; }
            set { m_dProdCrewSize = value; }
        }
        public decimal SetupCrewSize
        {
            get { return m_dSetupCrewSize; }
            set { m_dSetupCrewSize = value; }
        }
        public string SetupGroup
        {
            get { return m_sSetupGroup; }
            set { m_sSetupGroup = value; }
        }
        public decimal ProdStandard
        {
            get { return m_dProdStandard; }
            set { m_dProdStandard = value; }
        }
        public string StdFormat
        {
            get { return m_sStdFormat; }
            set { m_sStdFormat = value; }
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
        public string StdBasis
        {
            get { return m_sStdBasis; }
            set { m_sStdBasis = value; }
        }
        public int NumberOfMachines
        {
            get { return m_iNumberOfMachines; }
            set { m_iNumberOfMachines = value; }
        }
        public decimal EstSetupHoursPerMachine
        {
            get { return m_dEstSetupHoursPerMachine; }
            set { m_dEstSetupHoursPerMachine = value; }
        }
        public bool OverrideRates
        {
            get { return m_bOverrideRates; }
            set { m_bOverrideRates = value; }
        }
        public int OperationsPerPart
        {
            get { return m_iOperationsPerPart; }
            set { m_iOperationsPerPart = value; }
        }
        public decimal QtyPer
        {
            get { return m_dQtyPer; }
            set { m_dQtyPer = value; }
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
        public bool BurdenEqualsLabor
        {
            get { return m_bBurdenEqualsLabor; }
            set { m_bBurdenEqualsLabor = value; }
        }
        public bool SplitOperations
        {
            get { return m_bSplitOperations; }
            set { m_bSplitOperations = value; }
        }
        public bool UseEstimates
        {
            get { return m_bUseEstimates; }
            set { m_bUseEstimates = value; }
        }        
        public bool SplitBurden
        {
            get { return m_bSplitBurden; }
            set { m_bSplitBurden = value; }
        }        
        public string BurdenType
        {
            get { return m_sBurdenType; }
            set { m_sBurdenType = value; }
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

        private bool NoOp
        {
            get { return m_bNoOp; }
        }
        #endregion

        #region Data Members
        private string m_sCompany;
        private string m_sJobNum;
        private string m_sParentPartNum;
        private string m_sParentRevNum;
        private int m_iAssemblySeq;
        private int m_iParentAssemblySeq;
        private int m_iOperationSeq;
        private string m_sOpCode;
        private int m_iOpDetailSeq;
        private bool m_bOpComplete;
        private string m_sCapabilityId;
        private string m_sResourceGroupId;
        private string m_sResourceId;
        private decimal m_dEstSetupHours;
        private decimal m_dActSetupHours;
        private decimal m_dEstProdHours;
        private decimal m_dActProdHours;
        private int m_iNumResources;
        private decimal m_dProdCrewSize;
        private decimal m_dSetupCrewSize;
        private string m_sSetupGroup;
        private decimal m_dProdStandard;
        private string m_sStdFormat;
        private string m_sLaborEntryMethod;
        private string m_sOpStandardId;
        private string m_sStdBasis;
        private int m_iNumberOfMachines;
        private decimal m_dEstSetupHoursPerMachine;
        private bool m_bOverrideRates;
        private int m_iOperationsPerPart;
        private decimal m_dQtyPer;
        private decimal m_dEstScrap;
        private string m_sScrapType;
        private decimal m_dEstUnitCost;
        private bool m_bSubcontract;
        private decimal m_dDaysOut;
        private int m_iVendorNum;
        private decimal m_dRunQty;
        private string m_sScheduleRelation;
        private decimal m_dConcurrentCapacity;
        private string m_sSendAheadType;
        private decimal m_dSendAheadOffset;
        private bool m_bBurdenEqualsLabor;
        private bool m_bSplitOperations;
        private bool m_bUseEstimates;
        private bool m_bSplitBurden;
        private string m_sBurdenType;
        private DateTime m_dtStartDate;
        private DateTime m_dtDueDate;

        private bool m_bNoOp;
        #endregion
    }
}
