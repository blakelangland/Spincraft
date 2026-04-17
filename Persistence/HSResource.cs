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
    public class HSResource
    {
        #region Constructors

        public HSResource(DataRow oRow)
        {
            if (oRow["Resource_Inactive"] != DBNull.Value)
            {
                m_bInactive = (bool)oRow["Resource_Inactive"];
            }
            if (oRow["Resource_Company"] != DBNull.Value)
            {
                m_sCompany = (string)oRow["Resource_Company"];
            }
            if (oRow["Resource_ResourceGrpID"] != DBNull.Value)
            {
                m_sResourceGroupId = (string)oRow["Resource_ResourceGrpID"];
            }
            if (oRow["ResourceGroup_Description"] != DBNull.Value)
            {
                m_sResourceGroupDescription = (string)oRow["ResourceGroup_Description"];
            }
            if (oRow["Resource_ResourceID"] != DBNull.Value)
            {
                m_sResourceId = (string)oRow["Resource_ResourceID"];
            }
            if (oRow["Resource_Description"] != DBNull.Value)
            {
                m_sDescription = (string)oRow["Resource_Description"];
            }
            if (oRow["Resource_ResourceType"] != DBNull.Value)
            {
                m_sResourceType = (string)oRow["Resource_ResourceType"];
            }
            if (oRow["Resource_Location"] != DBNull.Value)
            {
                m_bResourceLocation = (bool)oRow["Resource_Location"];
            }
            if (oRow["Resource_Finite"] != DBNull.Value)
            {
                m_bResourceFinite = (bool)oRow["Resource_Finite"];
            }
            if (oRow["Resource_OpCode"] != DBNull.Value)
            {
                m_sOpCode = (string)oRow["Resource_OpCode"];
            }
            if (oRow["Resource_OpStdID"] != DBNull.Value)
            {
                m_sOpStandardId = (string)oRow["Resource_OpStdID"];
            }
            if (oRow["Resource_AllowManualOverride"] != DBNull.Value)
            {
                m_bAllowManualOverride = (bool)oRow["Resource_AllowManualOverride"];
            }
            if (oRow["Resource_GetDefaultLaborFromGroup"] != DBNull.Value)
            {
                m_bGetDefaultLaborFromGroup = (bool)oRow["Resource_GetDefaultLaborFromGroup"];
            }
            if (oRow["Resource_GetDefaultBurdenFromGroup"] != DBNull.Value)
            {
                m_bGetDefaultBurdenFromGroup = (bool)oRow["Resource_GetDefaultBurdenFromGroup"];
            }
            if (oRow["Resource_BurdenType"] != DBNull.Value)
            {
                m_sBurdenType = (string)oRow["Resource_BurdenType"];
            }
            if (oRow["Resource_ProdBurRate"] != DBNull.Value)
            {
                m_dProductionBurdenRate = (decimal)oRow["Resource_ProdBurRate"];
            }
            if (oRow["Resource_ProdLabRate"] != DBNull.Value)
            {
                m_dProductionLaborRate = (decimal)oRow["Resource_ProdLabRate"];
            }
            if (oRow["Resource_SetupBurRate"] != DBNull.Value)
            {
                m_dSetupBurdenRate = (decimal)oRow["Resource_SetupBurRate"];
            }
            if (oRow["Resource_SetupLabRate"] != DBNull.Value)
            {
                m_dSetupLaborRate = (decimal)oRow["Resource_SetupLabRate"];
            }
            if (oRow["Resource_QBurdenType"] != DBNull.Value)
            {
                m_sQuoteBurdenType = (string)oRow["Resource_QBurdenType"];
            }
            if (oRow["Resource_QProdBurRate"] != DBNull.Value)
            {
                m_dQuoteProductionBurdenRate = (decimal)oRow["Resource_QProdBurRate"];
            }
            if (oRow["Resource_QProdLabRate"] != DBNull.Value)
            {
                m_dQuoteProductionLaborRate = (decimal)oRow["Resource_QProdLabRate"];
            }
            if (oRow["Resource_QSetupBurRate"] != DBNull.Value)
            {
                m_dQuoteSetupBurdenRate = (decimal)oRow["Resource_QSetupBurRate"];
            }
            if (oRow["Resource_QSetupLabRate"] != DBNull.Value)
            {
                m_dQuoteSetupLaborRate = (decimal)oRow["Resource_QSetupLabRate"];
            }
            if (oRow["Resource_GetDefaultWhseFromGroup"] != DBNull.Value)
            {
                m_bGetDefaultWarehouseFromGroup = (bool)oRow["Resource_GetDefaultWhseFromGroup"];
            }
            if (oRow["Resource_OutputWhse"] != DBNull.Value)
            {
                m_sOutputWarehouse = (string)oRow["Resource_OutputWhse"];
            }
            if (oRow["Resource_OutputBinNum"] != DBNull.Value)
            {
                m_sOutputBin = (string)oRow["Resource_OutputBinNum"];
            }
            if (oRow["Resource_BackflushWhse"] != DBNull.Value)
            {
                m_sBackflushWarehouse = (string)oRow["Resource_BackflushWhse"];
            }
            if (oRow["Resource_BackflushBinNum"] != DBNull.Value)
            {
                m_sBackflushBin = (string)oRow["Resource_BackflushBinNum"];
            }
            if (oRow["Resource_InputWhse"] != DBNull.Value)
            {
                m_sInputWarehouse = (string)oRow["Resource_InputWhse"];
            }
            if (oRow["Resource_InputBinNum"] != DBNull.Value)
            {
                m_sInputBin = (string)oRow["Resource_InputBinNum"];
            }
            if (oRow["Resource_VendorNum"] != DBNull.Value)
            {
                m_iVendorNum = (int)oRow["Resource_VendorNum"];
            }
            if (oRow["Resource_CalendarID"] != DBNull.Value)
            {
                m_sCalendarId = (string)oRow["Resource_CalendarID"];
            }
            if (oRow["Resource_GetDefaultMQFromGroup"] != DBNull.Value)
            {
                m_bGetDefaultMoveQueueFromGroup = (bool)oRow["Resource_GetDefaultMQFromGroup"];
            }
            if (oRow["Resource_AutoMove"] != DBNull.Value)
            {
                m_bAutoMove = (bool)oRow["Resource_AutoMove"];
            }
            if (oRow["Resource_MoveHours"] != DBNull.Value)
            {
                m_dMoveHours = (decimal)oRow["Resource_MoveHours"];
            }
            if (oRow["Resource_QueHours"] != DBNull.Value)
            {
                m_dQueueHours = (decimal)oRow["Resource_QueHours"];
            }
            if (oRow["Resource_FiniteHorizon"] != DBNull.Value)
            {
                m_iFiniteHorizon = (int)oRow["Resource_FiniteHorizon"];
            }
            if (oRow["Resource_SplitOperations"] != DBNull.Value)
            {
                m_bSplitOperations = (bool)oRow["Resource_SplitOperations"];
            }
            if (oRow["Resource_ConcurrentCapacity"] != DBNull.Value)
            {
                m_dConcurrentCapacity = (decimal)oRow["Resource_ConcurrentCapacity"];
            }
            if (oRow["Resource_InformOverload"] != DBNull.Value)
            {
                m_bInformOverload = (bool)oRow["Resource_InformOverload"];
            }
        }

        #endregion

        #region Methods
        #endregion

        #region Properties
        public bool Inactive
        {
            get { return m_bInactive; }
            set { m_bInactive = value; }
        }
        public string Company
        {
            get { return m_sCompany; }
            set { m_sCompany = value; }
        }
        public string ResourceGroupId
        {
            get { return m_sResourceGroupId; }
            set { m_sResourceGroupId = value; }
        }
        public string ResourceGroupDescription
        {
            get { return m_sResourceGroupDescription; }
            set { m_sResourceGroupDescription = value; }
        }
        public string ResourceId
        {
            get { return m_sResourceId; }
            set { m_sResourceId = value; }
        }
        public string Description
        {
            get { return m_sDescription; }
            set { m_sDescription = value; }
        }
        public string ResourceType
        {
            get { return m_sResourceType; }
            set { m_sResourceType = value; }
        }
        public bool ResourceLocation
        {
            get { return m_bResourceLocation; }
            set { m_bResourceLocation = value; }
        }
        public bool ResourceFinite
        {
            get { return m_bResourceFinite; }
            set { m_bResourceFinite = value; }
        }
        public string OpCode
        {
            get { return m_sOpCode; }
            set { m_sOpCode = value; }
        }
        public string OpStandardId
        {
            get { return m_sOpStandardId; }
            set { m_sOpStandardId = value; }
        }
        public bool AllowManualOverride
        {
            get { return m_bAllowManualOverride; }
            set { m_bAllowManualOverride = value; }
        }
        public bool GetDefaultLaborFromGroup
        {
            get { return m_bGetDefaultLaborFromGroup; }
            set { m_bGetDefaultLaborFromGroup = value; }
        }
        public bool GetDefaultBurdenFromGroup
        {
            get { return m_bGetDefaultBurdenFromGroup; }
            set { m_bGetDefaultBurdenFromGroup = value; }
        }
        public string BurdenType
        {
            get { return m_sBurdenType; }
            set { m_sBurdenType = value; }
        }
        public decimal ProductionBurdenRate
        {
            get { return m_dProductionBurdenRate; }
            set { m_dProductionBurdenRate = value; }
        }
        public decimal ProductionLaborRate
        {
            get { return m_dProductionLaborRate; }
            set { m_dProductionLaborRate = value; }
        }
        public decimal SetupBurdenRate
        {
            get { return m_dSetupBurdenRate; }
            set { m_dSetupBurdenRate = value; }
        }
        public decimal SetupLaborRate
        {
            get { return m_dSetupLaborRate; }
            set { m_dSetupLaborRate = value; }
        }
        public string QuoteBurdenType
        {
            get { return m_sQuoteBurdenType; }
            set { m_sQuoteBurdenType = value; }
        }
        public decimal QuoteProductionBurdenRate
        {
            get { return m_dQuoteProductionBurdenRate; }
            set { m_dQuoteProductionBurdenRate = value; }
        }
        public decimal QuoteProductionLaborRate
        {
            get { return m_dQuoteProductionLaborRate; }
            set { m_dQuoteProductionLaborRate = value; }
        }
        public decimal QuoteSetupBurdenRate
        {
            get { return m_dQuoteSetupBurdenRate; }
            set { m_dQuoteSetupBurdenRate = value; }
        }
        public decimal QuoteSetupLaborRate
        {
            get { return m_dQuoteSetupLaborRate; }
            set { m_dQuoteSetupLaborRate = value; }
        }
        public bool GetDefaultWarehouseFromGroup
        {
            get { return m_bGetDefaultWarehouseFromGroup; }
            set { m_bGetDefaultWarehouseFromGroup = value; }
        }
        public string OutputWarehouse
        {
            get { return m_sOutputWarehouse; }
            set { m_sOutputWarehouse = value; }
        }
        public string OutputBin
        {
            get { return m_sOutputBin; }
            set { m_sOutputBin = value; }
        }
        public string BackflushWarehouse
        {
            get { return m_sBackflushWarehouse; }
            set { m_sBackflushWarehouse = value; }
        }
        public string BackflushBin
        {
            get { return m_sBackflushBin; }
            set { m_sBackflushBin = value; }
        }
        public string InputWarehouse
        {
            get { return m_sInputWarehouse; }
            set { m_sInputWarehouse = value; }
        }
        public string InputBin
        {
            get { return m_sInputBin; }
            set { m_sInputBin = value; }
        }
        public int VendorNum
        {
            get { return m_iVendorNum; }
            set { m_iVendorNum = value; }
        }
        public string CalendarId
        {
            get { return m_sCalendarId; }
            set { m_sCalendarId = value; }
        }
        public bool GetDefaultMoveQueueFromGroup
        {
            get { return m_bGetDefaultMoveQueueFromGroup; }
            set { m_bGetDefaultMoveQueueFromGroup = value; }
        }
        public bool AutoMove
        {
            get { return m_bAutoMove; }
            set { m_bAutoMove = value; }
        }
        public decimal MoveHours
        {
            get { return m_dMoveHours; }
            set { m_dMoveHours = value; }
        }
        public decimal QueueHours
        {
            get { return m_dQueueHours; }
            set { m_dQueueHours = value; }
        }
        public int FiniteHorizon
        {
            get { return m_iFiniteHorizon; }
            set { m_iFiniteHorizon = value; }
        }
        public bool SplitOperations
        {
            get { return m_bSplitOperations; }
            set { m_bSplitOperations = value; }
        }
        public decimal ConcurrentCapacity
        {
            get { return m_dConcurrentCapacity; }
            set { m_dConcurrentCapacity = value; }
        }

        public bool InformOverload
        {
            get { return m_bInformOverload; }
            set { m_bInformOverload = value; }
        }
        #endregion

        #region Data Members
        private bool m_bInactive;
        private string m_sCompany;
        private string m_sResourceGroupId;
        private string m_sResourceGroupDescription;
        private string m_sResourceId;
        private string m_sDescription;
        private string m_sResourceType;
        private bool m_bResourceLocation;
        private bool m_bResourceFinite;
        private string m_sOpCode;
        private string m_sOpStandardId;
        private bool m_bAllowManualOverride;
        private bool m_bGetDefaultLaborFromGroup;
        private bool m_bGetDefaultBurdenFromGroup;
        private string m_sBurdenType;
        private decimal m_dProductionBurdenRate;
        private decimal m_dProductionLaborRate;
        private decimal m_dSetupBurdenRate;
        private decimal m_dSetupLaborRate;
        private string m_sQuoteBurdenType;
        private decimal m_dQuoteProductionBurdenRate;
        private decimal m_dQuoteProductionLaborRate;
        private decimal m_dQuoteSetupBurdenRate;
        private decimal m_dQuoteSetupLaborRate;
        private bool m_bGetDefaultWarehouseFromGroup;
        private string m_sOutputWarehouse;
        private string m_sOutputBin;
        private string m_sBackflushWarehouse;
        private string m_sBackflushBin;
        private string m_sInputWarehouse;
        private string m_sInputBin;
        private int m_iVendorNum;
        private string m_sCalendarId;
        private bool m_bGetDefaultMoveQueueFromGroup;
        private bool m_bAutoMove;
        private decimal m_dMoveHours;
        private decimal m_dQueueHours;
        private int m_iFiniteHorizon;
        private bool m_bSplitOperations;
        private decimal m_dConcurrentCapacity;
        private bool m_bInformOverload;
        #endregion
    }
}
