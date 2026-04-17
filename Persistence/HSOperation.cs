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
    public class HSOperation
    {
        #region Constructors

        public HSOperation(DataRow oRow)
        {
            if (oRow["OpMaster_Company"] != DBNull.Value)
            {
                m_sCompany = (string)oRow["OpMaster_Company"];
            }
            if (oRow["OpMaster_OpCode"] != DBNull.Value)
            {
                m_sOpCode = (string)oRow["OpMaster_OpCode"];
            }
            if (oRow["OpMaster_OpDesc"] != DBNull.Value)
            {
                m_sDescription = (string)oRow["OpMaster_OpDesc"];
            }
            if (oRow["OpMaster_CommentText"] != DBNull.Value)
            {
                m_sCommentText = (string)oRow["OpMaster_CommentText"];
            }
            if (oRow["OpMaster_BillLaborRate"] != DBNull.Value)
            {
                m_dBilLLaborRate = (decimal)oRow["OpMaster_BillLaborRate"];
            }
            if (oRow["OpMaster_EstLabHours"] != DBNull.Value)
            {
                m_dEstLaborHours = (decimal)oRow["OpMaster_EstLabHours"];
            }
            if (oRow["OpMaster_VendorNum"] != DBNull.Value)
            {
                m_iVendorNum = (int)oRow["OpMaster_VendorNum"];
            }
            if (oRow["OpMaster_Subcontract"] != DBNull.Value)
            {
                m_bSubcontract = (bool)oRow["OpMaster_Subcontract"];
            }
            if (oRow["OpMaster_SendAheadType"] != DBNull.Value)
            {
                m_sSendAheadType = (string)oRow["OpMaster_SendAheadType"];
            }
            if (oRow["OpMaster_SendAheadOffset"] != DBNull.Value)
            {
                m_dSendAheadOffset = (decimal)oRow["OpMaster_SendAheadOffset"];
            }
            if (oRow["OpMasDtl_OpDtlSeq"] != DBNull.Value)
            {
                m_iOpDetailSeq = (int)oRow["OpMasDtl_OpDtlSeq"];
            }
            if (oRow["OpMasDtl_SetupOrProd"] != DBNull.Value)
            {
                m_sSetupOrProduction = (string)oRow["OpMasDtl_SetupOrProd"];
            }
            if (oRow["OpMasDtl_CapabilityID"] != DBNull.Value)
            {
                m_sCapabilityId = (string)oRow["OpMasDtl_CapabilityID"];
            }
            if (oRow["OpMasDtl_ResourceGrpID"] != DBNull.Value)
            {
                m_sResourceGroupId = (string)oRow["OpMasDtl_ResourceGrpID"];
            }
            if (oRow["OpMasDtl_ResourceID"] != DBNull.Value)
            {
                m_sResourceId = (string)oRow["OpMasDtl_ResourceID"];
            }
            if (oRow["OpMasDtl_SetupHours"] != DBNull.Value)
            {
                m_dSetupHours = (decimal)oRow["OpMasDtl_SetupHours"];
            }
            if (oRow["OpMasDtl_ProdHours"] != DBNull.Value)
            {
                m_dProductionHours = (decimal)oRow["OpMasDtl_ProdHours"];
            }
            if (oRow["OpMasDtl_NumResources"] != DBNull.Value)
            {
                m_iSchedulingBlocks = (int)oRow["OpMasDtl_NumResources"];
            }
            if (oRow["OpMasDtl_OpDtlDesc"] != DBNull.Value)
            {
                m_sOpDtlSeqDescription = (string)oRow["OpMasDtl_OpDtlDesc"];
            }
            if (oRow["OpMasDtl_ConcurrentCapacity"] != DBNull.Value)
            {
                m_dConcurrentCapacity = (decimal)oRow["OpMasDtl_ConcurrentCapacity"];
            }
            if (oRow["OpMasDtl_DailyProdRate"] != DBNull.Value)
            {
                m_dDailyProductionRate = (decimal)oRow["OpMasDtl_DailyProdRate"];
            }
            if (oRow["OpMasDtl_ProdCrewSize"] != DBNull.Value)
            {
                m_dProductionCrewSize = (decimal)oRow["OpMasDtl_ProdCrewSize"];
            }
            if (oRow["OpMasDtl_SetUpCrewSize"] != DBNull.Value)
            {
                m_dSetupCrewSize = (decimal)oRow["OpMasDtl_SetUpCrewSize"];
            }
            if (oRow["OpMasDtl_PrimaryProd"] != DBNull.Value)
            {
                m_bPrimaryProduction = (bool)oRow["OpMasDtl_PrimaryProd"];
            }
            if (oRow["OpMasDtl_PrimarySetup"] != DBNull.Value)
            {
                m_bPrimarySetup = (bool)oRow["OpMasDtl_PrimarySetup"];
            }
            if (oRow["OpStd_ProdStandard"] != DBNull.Value)
            {
                m_dProductionStandard = (decimal)oRow["OpStd_ProdStandard"];
            }
            if (oRow["OpStd_StdBasis"] != DBNull.Value)
            {
                m_sStandardBasis = (string)oRow["OpStd_StdBasis"];
            }
            if (oRow["OpStd_SetupGroup"] != DBNull.Value)
            {
                m_sSetupGroup = (string)oRow["OpStd_SetupGroup"];
            }
            if (oRow["OpMaster_SchedPrecedence"] != DBNull.Value)
            {
                m_iSchedulePrecedence = (int)oRow["OpMaster_SchedPrecedence"];
            }
            if (oRow["OpMaster_PrimarySetupOpDtl"] != DBNull.Value)
            {
                m_iPrimarySetupOpDetail = (int)oRow["OpMaster_PrimarySetupOpDtl"];
            }
            if (oRow["OpMaster_PrimaryProdOpDtl"] != DBNull.Value)
            {
                m_iPrimarProductionOpDetail = (int)oRow["OpMaster_PrimaryProdOpDtl"];
            }
            if (oRow["OpStd_OpStdID"] != DBNull.Value)
            {
                m_sOpStandardId = (string)oRow["OpStd_OpStdID"];
            }
            if (oRow["OpStd_SetupHours"] != DBNull.Value)
            {
                m_dOpStandardSetupHours = (decimal)oRow["OpStd_SetupHours"];
            }
            if (oRow["OpStd_StdFormat"] != DBNull.Value)
            {
                m_sStandardFormat = (string)oRow["OpStd_StdFormat"];
            }
            if (oRow["OpStd_ResourceGrpID"] != DBNull.Value)
            {
                m_sOpStandardResourceGroupId = (string)oRow["OpStd_ResourceGrpID"];
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
        public string OpCode
        {
            get { return m_sOpCode; }
            set { m_sOpCode = value; }
        }
        public string Description
        {
            get { return m_sDescription; }
            set { m_sDescription = value; }
        }
        public string CommentText
        {
            get { return m_sCommentText; }
            set { m_sCommentText = value; }
        }
        public decimal BilLLaborRate
        {
            get { return m_dBilLLaborRate; }
            set { m_dBilLLaborRate = value; }
        }
        public decimal EstLaborHours
        {
            get { return m_dEstLaborHours; }
            set { m_dEstLaborHours = value; }
        }
        public int VendorNum
        {
            get { return m_iVendorNum; }
            set { m_iVendorNum = value; }
        }
        public bool Subcontract
        {
            get { return m_bSubcontract; }
            set { m_bSubcontract = value; }
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
        public int OpDetailSeq
        {
            get { return m_iOpDetailSeq; }
            set { m_iOpDetailSeq = value; }
        }
        public string SetupOrProduction
        {
            get { return m_sSetupOrProduction; }
            set { m_sSetupOrProduction = value; }
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
        public decimal SetupHours
        {
            get { return m_dSetupHours; }
            set { m_dSetupHours = value; }
        }
        public decimal ProductionHours
        {
            get { return m_dProductionHours; }
            set { m_dProductionHours = value; }
        }
        public int SchedulingBlocks
        {
            get { return m_iSchedulingBlocks; }
            set { m_iSchedulingBlocks = value; }
        }
        public string OpDtlSeqDescription
        {
            get { return m_sOpDtlSeqDescription; }
            set { m_sOpDtlSeqDescription = value; }
        }
        public decimal ConcurrentCapacity
        {
            get { return m_dConcurrentCapacity; }
            set { m_dConcurrentCapacity = value; }
        }
        public decimal DailyProductionRate
        {
            get { return m_dDailyProductionRate; }
            set { m_dDailyProductionRate = value; }
        }
        public decimal ProductionCrewSize
        {
            get { return m_dProductionCrewSize; }
            set { m_dProductionCrewSize = value; }
        }
        public decimal SetupCrewSize
        {
            get { return m_dSetupCrewSize; }
            set { m_dSetupCrewSize = value; }
        }
        public bool PrimaryProduction
        {
            get { return m_bPrimaryProduction; }
            set { m_bPrimaryProduction = value; }
        }
        public bool PrimarySetup
        {
            get { return m_bPrimarySetup; }
            set { m_bPrimarySetup = value; }
        }
        public decimal ProductionStandard
        {
            get { return m_dProductionStandard; }
            set { m_dProductionStandard = value; }
        }
        
        public string StandardBasis
        {
            get { return m_sStandardBasis; }
            set { m_sStandardBasis = value; }
        }
        public string SetupGroup
        {
            get { return m_sSetupGroup; }
            set { m_sSetupGroup = value; }
        }
        public int SchedulePrecedence
        {
            get { return m_iSchedulePrecedence; }
            set { m_iSchedulePrecedence = value; }
        }
        public int PrimarySetupOpDetail
        {
            get { return m_iPrimarySetupOpDetail; }
            set { m_iPrimarySetupOpDetail = value; }
        }
        public int PrimarProductionOpDetail
        {
            get { return m_iPrimarProductionOpDetail; }
            set { m_iPrimarProductionOpDetail = value; }
        }        
        public string OpStandardId
        {
            get { return m_sOpStandardId; }
            set { m_sOpStandardId = value; }
        }
        public decimal OpStandardSetupHours
        {
            get { return m_dOpStandardSetupHours; }
            set { m_dOpStandardSetupHours = value; }
        }
        public string StandardFormat
        {
            get { return m_sStandardFormat; }
            set { m_sStandardFormat = value; }
        }
        public string OpStandardResourceGroupId
        {
            get { return m_sOpStandardResourceGroupId; }
            set { m_sOpStandardResourceGroupId = value; }
        }
        #endregion

        #region Data Members
        private string m_sCompany;
        private string m_sOpCode;
        private string m_sDescription;
        private string m_sCommentText;
        private decimal m_dBilLLaborRate;
        private decimal m_dEstLaborHours;
        private int m_iVendorNum;
        private bool m_bSubcontract;
        private string m_sSendAheadType;
        private decimal m_dSendAheadOffset;
        private int m_iOpDetailSeq;
        private string m_sSetupOrProduction;
        private string m_sCapabilityId;
        private string m_sResourceGroupId;
        private string m_sResourceId;
        private decimal m_dSetupHours;
        private decimal m_dProductionHours;
        private int m_iSchedulingBlocks;
        private string m_sOpDtlSeqDescription;
        private decimal m_dConcurrentCapacity;
        private decimal m_dDailyProductionRate;
        private decimal m_dProductionCrewSize;
        private decimal m_dSetupCrewSize;
        private bool m_bPrimaryProduction;
        private bool m_bPrimarySetup;
        private decimal m_dProductionStandard;
        private string m_sStandardBasis;
        private string m_sSetupGroup;
        private int m_iSchedulePrecedence;
        private int m_iPrimarySetupOpDetail;
        private int m_iPrimarProductionOpDetail;
        private string m_sOpStandardId;
        private decimal m_dOpStandardSetupHours;
        private string m_sStandardFormat;
        private string m_sOpStandardResourceGroupId;
        #endregion
    }
}
