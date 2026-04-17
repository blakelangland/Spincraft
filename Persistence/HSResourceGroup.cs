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
    public class HSResourceGroup
    {
        #region Constructors

        public HSResourceGroup(DataRow oRow)
        {
            if (oRow["ResourceGroup_Inactive"] != DBNull.Value)
            {
                m_bInactive = (bool)oRow["ResourceGroup_Inactive"];
            }
            if (oRow["ResourceGroup_Company"] != DBNull.Value)
            {
                m_sCompany = (string)oRow["ResourceGroup_Company"];
            }
            if (oRow["ResourceGroup_Plant"] != DBNull.Value)
            {
                m_sSite = (string)oRow["ResourceGroup_Plant"];
            }
            if (oRow["ResourceGroup_ResourceGrpID"] != DBNull.Value)
            {
                m_sResourceGroupId = (string)oRow["ResourceGroup_ResourceGrpID"];
            }
            if (oRow["ResourceGroup_Description"] != DBNull.Value)
            {
                m_sDescription = (string)oRow["ResourceGroup_Description"];
            }
            if (oRow["ResourceGroup_Location"] != DBNull.Value)
            {
                m_bLocation = (bool)oRow["ResourceGroup_Location"];
            }
            if (oRow["ResourceGroup_OpCode"] != DBNull.Value)
            {
                m_sOpCode = (string)oRow["ResourceGroup_OpCode"];
            }
            if (oRow["ResourceGroup_OpStdID"] != DBNull.Value)
            {
                m_sOpStandardId = (string)oRow["ResourceGroup_OpStdID"];
            }
            if (oRow["ResourceGroup_ResourceType"] != DBNull.Value)
            {
                m_sResourceType = (string)oRow["ResourceGroup_ResourceType"];
            }
            if (oRow["ResourceGroup_CalendarID"] != DBNull.Value)
            {
                m_sCalendarId = (string)oRow["ResourceGroup_CalendarID"];
            }
            if (oRow["ResourceGroup_JCDept"] != DBNull.Value)
            {
                m_sJCDepartment = (string)oRow["ResourceGroup_JCDept"];
            }
            if (oRow["ResourceGroup_SubContract"] != DBNull.Value)
            {
                m_bSubcontract = (bool)oRow["ResourceGroup_SubContract"];
            }
            if (oRow["ResourceGroup_UseEstimates"] != DBNull.Value)
            {
                m_bUseEstimates = (bool)oRow["ResourceGroup_UseEstimates"];
            }
            if (oRow["ResourceGroup_SplitBurden"] != DBNull.Value)
            {
                m_bSplitBurden = (bool)oRow["ResourceGroup_SplitBurden"];
            }
            if (oRow["ResourceGroup_BurdenEQLabor"] != DBNull.Value)
            {
                m_bBurdenEqualsLabor = (bool)oRow["ResourceGroup_BurdenEQLabor"];
            }
            if (oRow["ResourceGroup_BurdenType"] != DBNull.Value)
            {
                m_sBurdenType = (string)oRow["ResourceGroup_BurdenType"];
            }
            if (oRow["ResourceGroup_ProdCrewSize"] != DBNull.Value)
            {
                m_dProductionCrewSize = (decimal)oRow["ResourceGroup_ProdCrewSize"];
            }
            if (oRow["ResourceGroup_SetupCrewSize"] != DBNull.Value)
            {
                m_dSetupCrewSize = (decimal)oRow["ResourceGroup_SetupCrewSize"];
            }
            if (oRow["ResourceGroup_ProdBurRate"] != DBNull.Value)
            {
                m_dProductionBurdenRate = (decimal)oRow["ResourceGroup_ProdBurRate"];
            }
            if (oRow["ResourceGroup_ProdLabRate"] != DBNull.Value)
            {
                m_dProductionLaborRate = (decimal)oRow["ResourceGroup_ProdLabRate"];
            }
            if (oRow["ResourceGroup_SetupBurRate"] != DBNull.Value)
            {
                m_dSetupBurdenRate = (decimal)oRow["ResourceGroup_SetupBurRate"];
            }
            if (oRow["ResourceGroup_SetupLabRate"] != DBNull.Value)
            {
                m_dSetupLaborRate = (decimal)oRow["ResourceGroup_SetupLabRate"];
            }
            if (oRow["ResourceGroup_QBurdenType"] != DBNull.Value)
            {
                m_sQuoteBurdenType = (string)oRow["ResourceGroup_QBurdenType"];
            }
            if (oRow["ResourceGroup_QProdBurRate"] != DBNull.Value)
            {
                m_dQuoteProductionBurdenRate = (decimal)oRow["ResourceGroup_QProdBurRate"];
            }
            if (oRow["ResourceGroup_QProdLabRate"] != DBNull.Value)
            {
                m_dQuoteProductionLaborRate = (decimal)oRow["ResourceGroup_QProdLabRate"];
            }
            if (oRow["ResourceGroup_QSetupBurRate"] != DBNull.Value)
            {
                m_dQuoteSetupBurdenRate = (decimal)oRow["ResourceGroup_QSetupBurRate"];
            }
            if (oRow["ResourceGroup_QSetupLabRate"] != DBNull.Value)
            {
                m_dQuoteSetupLaborRate = (decimal)oRow["ResourceGroup_QSetupLabRate"];
            }
            if (oRow["ResourceGroup_AllowManualOverride"] != DBNull.Value)
            {
                m_bAllowManualOverride = (bool)oRow["ResourceGroup_AllowManualOverride"];
            }
            if (oRow["ResourceGroup_NumberOfMachines"] != DBNull.Value)
            {
                m_iNumberOfMachines = (int)oRow["ResourceGroup_NumberOfMachines"];
            }
            if (oRow["ResourceGroup_SchMachine"] != DBNull.Value)
            {
                m_iSchedulingMachines = (int)oRow["ResourceGroup_SchMachine"];
            }
            if (oRow["ResourceGroup_FiniteHorizon"] != DBNull.Value)
            {
                m_iFiniteHorizon = (int)oRow["ResourceGroup_FiniteHorizon"];
            }
            if (oRow["ResourceGroup_SplitOperations"] != DBNull.Value)
            {
                m_bSplitOperations = (bool)oRow["ResourceGroup_SplitOperations"];
            }
            if (oRow["ResourceGroup_UseCalendarForMove"] != DBNull.Value)
            {
                m_bUseCalendarForMove = (bool)oRow["ResourceGroup_UseCalendarForMove"];
            }
            if (oRow["ResourceGroup_MoveHours"] != DBNull.Value)
            {
                m_dMoveHours = (decimal)oRow["ResourceGroup_MoveHours"];
            }
            if (oRow["ResourceGroup_UseCalendarForQueue"] != DBNull.Value)
            {
                m_bUseCalendarForQueue = (bool)oRow["ResourceGroup_UseCalendarForQueue"];
            }
            if (oRow["ResourceGroup_QueHours"] != DBNull.Value)
            {
                m_dQueueHours = (decimal)oRow["ResourceGroup_QueHours"];
            }
            if (oRow["ResourceGroup_InputWhse"] != DBNull.Value)
            {
                m_sInputWarehouse = (string)oRow["ResourceGroup_InputWhse"];
            }
            if (oRow["ResourceGroup_InputBinNum"] != DBNull.Value)
            {
                m_sInputBin = (string)oRow["ResourceGroup_InputBinNum"];
            }
            if (oRow["ResourceGroup_OutputWhse"] != DBNull.Value)
            {
                m_sOutputWarehouse = (string)oRow["ResourceGroup_OutputWhse"];
            }
            if (oRow["ResourceGroup_OutputBinNum"] != DBNull.Value)
            {
                m_sOutputBin = (string)oRow["ResourceGroup_OutputBinNum"];
            }
            if (oRow["ResourceGroup_BackflushWhse"] != DBNull.Value)
            {
                m_sBackflushWarehouse = (string)oRow["ResourceGroup_BackflushWhse"];
            }
            if (oRow["ResourceGroup_BackflushBinNum"] != DBNull.Value)
            {
                m_sBackflushBin = (string)oRow["ResourceGroup_BackflushBinNum"];
            }
            if (oRow["ResourceGroup_InformOverload"] != DBNull.Value)
            {
                m_bInformOverload = (bool)oRow["ResourceGroup_InformOverload"];
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
        public string Site
        {
            get { return m_sSite; }
            set { m_sSite = value; }
        }
        public string ResourceGroupId
        {
            get { return m_sResourceGroupId; }
            set { m_sResourceGroupId = value; }
        }
        public string Description
        {
            get { return m_sDescription; }
            set { m_sDescription = value; }
        }
        public bool Location
        {
            get { return m_bLocation; }
            set { m_bLocation = value; }
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
        public string ResourceType
        {
            get { return m_sResourceType; }
            set { m_sResourceType = value; }
        }
        public string CalendarId
        {
            get { return m_sCalendarId; }
            set { m_sCalendarId = value; }
        }
        public string JCDepartment
        {
            get { return m_sJCDepartment; }
            set { m_sJCDepartment = value; }
        }
        public bool Subcontract
        {
            get { return m_bSubcontract; }
            set { m_bSubcontract = value; }
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
        public bool BurdenEqualsLabor
        {
            get { return m_bBurdenEqualsLabor; }
            set { m_bBurdenEqualsLabor = value; }
        }
        public string BurdenType
        {
            get { return m_sBurdenType; }
            set { m_sBurdenType = value; }
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
        public bool AllowManualOverride
        {
            get { return m_bAllowManualOverride; }
            set { m_bAllowManualOverride = value; }
        }
        public int NumberOfMachines
        {
            get { return m_iNumberOfMachines; }
            set { m_iNumberOfMachines = value; }
        }
        public int SchedulingMachines
        {
            get { return m_iSchedulingMachines; }
            set { m_iSchedulingMachines = value; }
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
        public bool UseCalendarForMove
        {
            get { return m_bUseCalendarForMove; }
            set { m_bUseCalendarForMove = value; }
        }
        public decimal MoveHours
        {
            get { return m_dMoveHours; }
            set { m_dMoveHours = value; }
        }
        public bool UseCalendarForQueue
        {
            get { return m_bUseCalendarForQueue; }
            set { m_bUseCalendarForQueue = value; }
        }
        public decimal QueueHours
        {
            get { return m_dQueueHours; }
            set { m_dQueueHours = value; }
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
        public bool InformOverload
        {
            get { return m_bInformOverload; }
            set { m_bInformOverload = value; }
        }
        #endregion

        #region Data Members
        private bool m_bInactive;
        private string m_sCompany;
        private string m_sSite;
        private string m_sResourceGroupId;
        private string m_sDescription;
        private bool m_bLocation;
        private string m_sOpCode;
        private string m_sOpStandardId;
        private string m_sResourceType;
        private string m_sCalendarId;
        private string m_sJCDepartment;
        private bool m_bSubcontract;
        private bool m_bUseEstimates;
        private bool m_bSplitBurden;
        private bool m_bBurdenEqualsLabor;
        private string m_sBurdenType;
        private decimal m_dProductionCrewSize;
        private decimal m_dSetupCrewSize;
        private decimal m_dProductionBurdenRate;
        private decimal m_dProductionLaborRate;
        private decimal m_dSetupBurdenRate;
        private decimal m_dSetupLaborRate;
        private string m_sQuoteBurdenType;
        private decimal m_dQuoteProductionBurdenRate;
        private decimal m_dQuoteProductionLaborRate;
        private decimal m_dQuoteSetupBurdenRate;
        private decimal m_dQuoteSetupLaborRate;
        private bool m_bAllowManualOverride;
        private int m_iNumberOfMachines;
        private int m_iSchedulingMachines;
        private int m_iFiniteHorizon;
        private bool m_bSplitOperations;
        private bool m_bUseCalendarForMove;
        private decimal m_dMoveHours;
        private bool m_bUseCalendarForQueue;
        private decimal m_dQueueHours;
        private string m_sInputWarehouse;
        private string m_sInputBin;
        private string m_sOutputWarehouse;
        private string m_sOutputBin;
        private string m_sBackflushWarehouse;
        private string m_sBackflushBin;
        private bool m_bInformOverload;
        #endregion
    }
}
