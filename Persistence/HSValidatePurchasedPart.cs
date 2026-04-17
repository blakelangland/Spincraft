using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using System.IO;

using Erp.BO;
using Erp.Proxy.BO;
using Erp.Adapters;
using Ice.Adapters;
using Ice.Lib.Framework;
using Ice.BO;
using Ice.Core;
using Ice.Tablesets;
using Ice.Lib.Searches;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Drawing;
//using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Drawing.Spreadsheet;
using SpreadsheetLight;
using SpreadsheetLight.Charts;



namespace HorizonScientific
{
    public class HSPurchasedPartCost
    {
        #region Constructor

        public HSPurchasedPartCost(DataRow oDataRow)
        {
            m_sPart = (string)oDataRow["PartCost_PartNum"];
            if (oDataRow["PartCost_StdMaterialCost"] != DBNull.Value)
            {
                m_dStdCost = (decimal)oDataRow["PartCost_StdMaterialCost"];
            }
            if (oDataRow["PartCost_LastMaterialCost"] != DBNull.Value)
            {
                m_dLastPurchasePrice = (decimal)oDataRow["PartCost_LastMaterialCost"];
            }
            if (oDataRow["Part_CreateDate_c"] != DBNull.Value)
            {
                m_dtCreateDate = (DateTime)oDataRow["Part_CreateDate_c"];
            }
        }

        #endregion

        #region Properties

        public string PartNum
        {
            get { return m_sPart; }
        }

        public decimal StandardCost
        {
            get { return m_dStdCost; }
        }

        public decimal LastPurchasePrice
        {
            get { return m_dLastPurchasePrice; }
        }

        public DateTime CreateDate
        {
            get { return m_dtCreateDate; }
        }
        #endregion

        #region Data Members

        private string m_sPart;
        private decimal m_dStdCost;
        private decimal m_dLastPurchasePrice;
        private DateTime m_dtCreateDate;
        #endregion
    }

    public class HSPart
    {
        #region Constructors
        public HSPart(DataRow oDataRow)
        {
            if ((oDataRow["Part_Company"] != DBNull.Value) && (string.IsNullOrEmpty((string)oDataRow["Part_Company"]) == false))
            {
                m_sCompany = (string)oDataRow["Part_Company"];
            }
            if ((oDataRow["Part_PartNum"] != DBNull.Value) && (string.IsNullOrEmpty((string)oDataRow["Part_PartNum"]) == false))
            {
                m_sPart = (string)oDataRow["Part_PartNum"];
            }
            if ((oDataRow["Part_PartDescription"] != DBNull.Value) && (string.IsNullOrEmpty((string)oDataRow["Part_PartDescription"]) == false))
            {
                m_sDescription = (string)oDataRow["Part_PartDescription"];
            }
            if (oDataRow["Part_InActive"] != DBNull.Value)
            {
                m_bInactive = (Boolean)oDataRow["Part_InActive"];
            }
            if ((oDataRow["Part_TypeCode"] != DBNull.Value) && (string.IsNullOrEmpty((string)oDataRow["Part_TypeCode"]) == false))
            {
                m_sPartTypeCode = (string)oDataRow["Part_TypeCode"];
            }
            if ((oDataRow["Part_ClassID"] != DBNull.Value) && (string.IsNullOrEmpty((string)oDataRow["Part_ClassID"]) == false))
            {
                m_sClassID = (string)oDataRow["Part_ClassID"];
            }
            if ((oDataRow["Part_ProdCode"] != DBNull.Value) && (string.IsNullOrEmpty((string)oDataRow["Part_ProdCode"]) == false))
            {
                m_sGroup = (string)oDataRow["Part_ProdCode"];
            }
            if ((oDataRow["Part_SearchWord"] != DBNull.Value) && (string.IsNullOrEmpty((string)oDataRow["Part_SearchWord"]) == false))
            {
                m_sSearch = (string)oDataRow["Part_SearchWord"];
            }
            if ((oDataRow["Part_UOMClassID"] != DBNull.Value) && (string.IsNullOrEmpty((string)oDataRow["Part_UOMClassID"]) == false))
            {
                m_sUOMClassID = (string)oDataRow["Part_UOMClassID"];
            }
            if ((oDataRow["Part_IUM"] != DBNull.Value) && (string.IsNullOrEmpty((string)oDataRow["Part_IUM"]) == false))
            {
                m_sIUOM = (string)oDataRow["Part_IUM"];
            }
            if ((oDataRow["Part_PUM"] != DBNull.Value) && (string.IsNullOrEmpty((string)oDataRow["Part_PUM"]) == false))
            {
                m_sPUOM = (string)oDataRow["Part_PUM"];
            }
            if ((oDataRow["Part_SalesUM"] != DBNull.Value) && (string.IsNullOrEmpty((string)oDataRow["Part_SalesUM"]) == false))
            {
                m_sSUOM = (string)oDataRow["Part_SalesUM"];
            }
            if (oDataRow["Part_NetWeight"] != DBNull.Value)
            {
                m_dUnitNetWeight = (Decimal)oDataRow["Part_NetWeight"];
            }
            if ((oDataRow["Part_NetWeightUOM"] != DBNull.Value) && (string.IsNullOrEmpty((string)oDataRow["Part_NetWeightUOM"]) == false))
            {
                m_sWeightUOM = (string)oDataRow["Part_NetWeightUOM"];
            }
            if ((oDataRow["Part_ISOrigCountry"] != DBNull.Value) && (string.IsNullOrEmpty((string)oDataRow["Part_ISOrigCountry"]) == false))
            {
                m_sCountryofOrigin = (string)oDataRow["Part_ISOrigCountry"];
            }
            if ((oDataRow["Part_HTS"] != DBNull.Value) && (string.IsNullOrEmpty((string)oDataRow["Part_HTS"]) == false))
            {
                m_sHTS = (string)oDataRow["Part_HTS"];
            }
            if ((oDataRow["Part_CommodityCode"] != DBNull.Value) && (string.IsNullOrEmpty((string)oDataRow["Part_CommodityCode"]) == false))
            {
                m_sCommodityCode = (string)oDataRow["Part_CommodityCode"];
            }
            if (oDataRow["Calculated_CommodityCodeLength"] != DBNull.Value)
            {
                m_iCommodityCodeLength = (int)oDataRow["Calculated_CommodityCodeLength"];
            }
            if ((oDataRow["Part_ProductPortfolio_c"] != DBNull.Value) && (string.IsNullOrEmpty((string)oDataRow["Part_ProductPortfolio_c"]) == false))
            {
                m_sProductPortfolio = (string)oDataRow["Part_ProductPortfolio_c"];
            }
            if ((oDataRow["Part_WarrantyCode"] != DBNull.Value) && (string.IsNullOrEmpty((string)oDataRow["Part_WarrantyCode"]) == false))
            {
                m_sWarrantyCode = (string)oDataRow["Part_WarrantyCode"];
            }  
            if ((oDataRow["Part_CreatedBy"] != DBNull.Value) && (string.IsNullOrEmpty((string)oDataRow["Part_CreatedBy"]) == false))
            {
                m_sCreatedBy = (string)oDataRow["Part_CreatedBy"];
            }
            if (oDataRow["Part_CreatedOn"] != DBNull.Value)
            {
                m_dtCreatedOn = (DateTime)oDataRow["Part_CreatedOn"];
            }
            if (oDataRow["Part_RunOut"] != DBNull.Value)
            {
                m_bPartRunOut = (bool)oDataRow["Part_RunOut"];
            }
            if (oDataRow["Part_OnHold"] != DBNull.Value)
            {
                m_bPartOnHold = (bool)oDataRow["Part_OnHold"];
            }
            if (oDataRow["Part_TrackLots"] != DBNull.Value)
            {
                m_bPartTrackLots = (bool)oDataRow["Part_TrackLots"];
            }
            if (oDataRow["Part_TrackSerialNum"] != DBNull.Value)
            {
                m_bTrackSerial = (Boolean)oDataRow["Part_TrackSerialNum"];
            }
            if (oDataRow["Part_UsePartRev"] != DBNull.Value)
            {
                m_bUsePartRev = (Boolean)oDataRow["Part_UsePartRev"];
            }
            if (oDataRow["Part_Constrained"] != DBNull.Value)
            {
                m_bConstrained = (Boolean)oDataRow["Part_Constrained"];
            }
            if (oDataRow["Part_PhantomBOM"] != DBNull.Value)
            {
                m_bPartPhantomBOM = (Boolean)oDataRow["Part_PhantomBOM"];
            }
            if (oDataRow["PartPlant_PhantomBOM"] != DBNull.Value)
            {
                m_bPlantPhantomBOM = (Boolean)oDataRow["PartPlant_PhantomBOM"];
            }
            if (oDataRow["Part_RcvInspectionReq"] != DBNull.Value)
            {
                m_bInspectionRequired = (Boolean)oDataRow["Part_RcvInspectionReq"];
            }
            if ((oDataRow["PlantWhse_Plant"] != DBNull.Value) && (string.IsNullOrEmpty((string)oDataRow["PlantWhse_Plant"]) == false))
            {
                m_sSite = (string)oDataRow["PlantWhse_Plant"];
            }
            if (oDataRow["PartPlant_ProcessMRP"] != DBNull.Value)
            {
                m_bPlantProcessMRP = (bool)oDataRow["PartPlant_ProcessMRP"];
            }
            if (oDataRow["PartPlant_GenerateSugg"] != DBNull.Value)
            {
                m_bGenerateSuggestions = (Boolean)oDataRow["PartPlant_GenerateSugg"];
            }
            if (oDataRow["PartPlant_BackFlush"] != DBNull.Value)
            {
                m_bPlantBackFlush = (Boolean)oDataRow["PartPlant_BackFlush"];
            }
            if (oDataRow["PartPlant_KitBackFlush"] != DBNull.Value)
            {
                m_bBackflushKitComponents = (bool)oDataRow["PartPlant_KitBackFlush"];
            }
            if (oDataRow["PartPlant_KitTime"] != DBNull.Value)
            {
                m_iKitTime = (int)oDataRow["PartPlant_KitTime"];
            }
            if (oDataRow["PartPlant_GetFromLocalWhse"] != DBNull.Value)
            {
                m_bGetFromLocalWarehouse = (bool)oDataRow["PartPlant_GetFromLocalWhse"];
            }
            if (oDataRow["PartPlant_MaximumQty"] != DBNull.Value)
            {
                m_dMaximum = (Decimal)oDataRow["PartPlant_MaximumQty"];
            }
            if (oDataRow["PartPlant_MinimumQty"] != DBNull.Value)
            {
                m_dMinimum = (Decimal)oDataRow["PartPlant_MinimumQty"];
            }
            if (oDataRow["PartPlant_SafetyQty"] != DBNull.Value)
            {
                m_dSafety = (Decimal)oDataRow["PartPlant_SafetyQty"];
            }
            if (oDataRow["PartPlant_MaxMfgLotSize"] != DBNull.Value)
            {
                m_dCostingLotSize = (Decimal)oDataRow["PartPlant_MaxMfgLotSize"];
            }
            if (oDataRow["PartPlant_MinMfgLotSize"] != DBNull.Value)
            {
                m_dMinMfgLotSize = (Decimal)oDataRow["PartPlant_MinMfgLotSize"];
            }
            if (oDataRow["PartPlant_MfgLotSize"] != DBNull.Value)
            {
                m_dMfgLotSize = (Decimal)oDataRow["PartPlant_MfgLotSize"];
            }
            if (oDataRow["PartPlant_MfgLotMultiple"] != DBNull.Value)
            {
                m_dMfgLotMultiple = (decimal)oDataRow["PartPlant_MfgLotMultiple"];
            }
            if (oDataRow["PartPlant_PrepTime"] != DBNull.Value)
            {
                m_iPreparationTime = (int)oDataRow["PartPlant_PrepTime"];
            }
            if (oDataRow["PartPlant_DaysOfSupply"] != DBNull.Value)
            {
                m_iDaysofSupply = (int)oDataRow["PartPlant_DaysOfSupply"];
            }
            if ((oDataRow["PartPlant_PersonID"] != DBNull.Value) && (string.IsNullOrEmpty((string)oDataRow["PartPlant_PersonID"]) == false))
            {
                m_sPlannerID = (string)oDataRow["PartPlant_PersonID"];
            }
            if ((oDataRow["PartPlant_PrimWhse"] != DBNull.Value) && (string.IsNullOrEmpty((string)oDataRow["PartPlant_PrimWhse"]) == false))
            {
                m_sPrimWhse = (string)oDataRow["PartPlant_PrimWhse"];
            }
            if ((oDataRow["Part_MfgComment"] != DBNull.Value) && (string.IsNullOrEmpty((string)oDataRow["Part_MfgComment"]) == false))
            {
                m_sMfgComments = (string)oDataRow["Part_MfgComment"];
            }
            if ((oDataRow["Part_TypeCode"] != DBNull.Value) && (string.IsNullOrEmpty((string)oDataRow["Part_TypeCode"]) == false))
            {
                m_sPartTypeCode = (string)oDataRow["Part_TypeCode"];
            }
            if ((oDataRow["PartPlant_SourceType"] != DBNull.Value) && (string.IsNullOrEmpty((string)oDataRow["PartPlant_SourceType"]) == false))
            {
                m_sPlantSourceType = (string)oDataRow["PartPlant_SourceType"];
            }
            if (oDataRow["PartPlant_NonStock"] != DBNull.Value)
            {
                m_bPlantNonStockItem = (Boolean)oDataRow["PartPlant_NonStock"];
            }
            if (oDataRow["Part_NonStock"] != DBNull.Value)
            {
                m_bPartNonStock = (bool)oDataRow["Part_NonStock"];
            }
            if (oDataRow["PartPlant_QtyBearing"] != DBNull.Value)
            {
                m_bPlantQtyBearing = (Boolean)oDataRow["PartPlant_QtyBearing"];
            }
            if (oDataRow["Part_QtyBearing"] != DBNull.Value)
            {
                m_bPartQtyBearing = (Boolean)oDataRow["Part_QtyBearing"];
            }
            if (oDataRow["PartPlant_BuyToOrder"] != DBNull.Value)
            {
                m_bPlantBuyToOrder = (Boolean)oDataRow["PartPlant_BuyToOrder"];
            }
            if (oDataRow["Part_BuyToOrder"] != DBNull.Value)
            {
                m_bPartBuyToOrder = (bool)oDataRow["Part_BuyToOrder"];
            }
            if (oDataRow["PartPlant_DropShip"] != DBNull.Value)
            {
                m_bPlantDropShip = (Boolean)oDataRow["PartPlant_DropShip"];
            }
            if (oDataRow["Part_DropShip"] != DBNull.Value)
            {
                m_bPartDropShip = (bool)oDataRow["Part_DropShip"];
            }
            if ((oDataRow["Part_CostMethod"] != DBNull.Value) && (string.IsNullOrEmpty((string)oDataRow["Part_CostMethod"]) == false))
            {
                m_sPartCostingMethod = (string)oDataRow["Part_CostMethod"];
            }
            if ((oDataRow["PartPlant_CostMethod"] != DBNull.Value) && (string.IsNullOrEmpty((string)oDataRow["PartPlant_CostMethod"]) == false))
            {
                m_sPlantCostingMethod = (string)oDataRow["PartPlant_CostMethod"];
            }
            if (oDataRow["Calculated_Cost"] != DBNull.Value)
            {
                m_dCost = (decimal)oDataRow["Calculated_Cost"];
            }
            if ((oDataRow["Part_PricePerCode"] != DBNull.Value) && (string.IsNullOrEmpty((string)oDataRow["Part_PricePerCode"]) == false))
            {
                m_sPricePer = (string)oDataRow["Part_PricePerCode"];
            }
            if ((oDataRow["Part_InternalPricePerCode"] != DBNull.Value) && (string.IsNullOrEmpty((string)oDataRow["Part_InternalPricePerCode"]) == false))
            {
                m_sInternalDivisionalPricePer = (string)oDataRow["Part_InternalPricePerCode"];
            }
            if (oDataRow["Part_UnitPrice"] != DBNull.Value)
            {
                m_dUnitPrice = (Decimal)oDataRow["Part_UnitPrice"];
            }
            if ((oDataRow["Part_PurComment"] != DBNull.Value) && (string.IsNullOrEmpty((string)oDataRow["Part_PurComment"]) == false))
            {
                m_sPurchaseComments = (string)oDataRow["Part_PurComment"];
            }
            if ((oDataRow["PartPlant_BuyerID"] != DBNull.Value) && (string.IsNullOrEmpty((string)oDataRow["PartPlant_BuyerID"]) == false))
            {
                m_sBuyerID = (string)oDataRow["PartPlant_BuyerID"];
            }
            if ((oDataRow["Vendor_Name"] != DBNull.Value) && (string.IsNullOrEmpty((string)oDataRow["Vendor_Name"]) == false))
            {
                m_sVendorName = (string)oDataRow["Vendor_Name"];
            }
            if ((oDataRow["Vendor_VendorID"] != DBNull.Value) && (string.IsNullOrEmpty((string)oDataRow["Vendor_VendorID"]) == false))
            {
                m_sSupplierID = (string)oDataRow["Vendor_VendorID"];
            }
            if (oDataRow["PartPlant_MinOrderQty"] != DBNull.Value)
            {
                m_dMinOrderQty = (Decimal)oDataRow["PartPlant_MinOrderQty"];
            }
            if (oDataRow["PartPlant_VendorNum"] != DBNull.Value)
            {
                m_iSupplierNum = (int)oDataRow["PartPlant_VendorNum"];
            }
            if (oDataRow["PartPlant_LeadTime"] != DBNull.Value)
            {
                m_iLead = (int)oDataRow["PartPlant_LeadTime"];
            }
            if (oDataRow["PartCost_LastMaterialCost"] != DBNull.Value)
            {
                m_dLastMaterialCost = (decimal)oDataRow["PartCost_LastMaterialCost"];
            }
            if (oDataRow["VendPart_BaseUnitPrice"] != DBNull.Value)
            {
                m_dVendorBaseUnitPrice = (decimal)oDataRow["VendPart_BaseUnitPrice"];
            }
            if (oDataRow["VendPart_EffectiveDate"] != DBNull.Value)
            {
                m_dtVendorPriceEffectiveDate = (DateTime)oDataRow["VendPart_EffectiveDate"];
            }
            if (oDataRow["VendPart_ExpirationDate"] != DBNull.Value)
            {
                m_dtVendorPriceExpirationDate = (DateTime)oDataRow["VendPart_ExpirationDate"];
            }
            if ((oDataRow["PartWhse_MinAbc"] != DBNull.Value) && (string.IsNullOrEmpty((string)oDataRow["PartWhse_MinAbc"]) == false))
            {
                m_sMinABC = (string)oDataRow["PartWhse_MinAbc"];
            }
            if ((oDataRow["PartWhse_SystemAbc"] != DBNull.Value) && (string.IsNullOrEmpty((string)oDataRow["PartWhse_SystemAbc"]) == false))
            {
                m_sActualABC = (string)oDataRow["PartWhse_SystemAbc"];
            }
            if (oDataRow["PartWhse_OvrrideCountFreq"] != DBNull.Value)
            {
                m_bOverrideCountFrequency = (bool)oDataRow["PartWhse_OvrrideCountFreq"];
            }
            if (oDataRow["PartWhse_CountFreq"] != DBNull.Value)
            {
                m_iCountFrequency = (int)oDataRow["PartWhse_CountFreq"];
            }
            if (oDataRow["PartWhse_LastCCDate"] != DBNull.Value)
            {
                m_dtLastCycleCountDate = (DateTime)oDataRow["PartWhse_LastCCDate"];
            }
            if (oDataRow["PartWhse_ManualABC"] != DBNull.Value)
            {
                m_bManualABCCode = (bool)oDataRow["PartWhse_ManualABC"];
            }
        }
        #endregion

        #region Properties

        public String Company
        {
            get { return m_sCompany; }
            set { m_sCompany = value; }
        }

        public String Part
        {
            get { return m_sPart; }
            set { m_sPart = value; }
        }

        public String Description
        {
            get { return m_sDescription; }
            set { m_sDescription = value; }
        }

        public Boolean Inactive
        {
            get { return m_bInactive; }
            set { m_bInactive = value; }
        }

        public String ClassID
        {
            get { return m_sClassID; }
            set { m_sClassID = value; }
        }

        public String Group
        {
            get { return m_sGroup; }
            set { m_sGroup = value; }
        }

        public String Search
        {
            get { return m_sSearch; }
            set { m_sSearch = value; }
        }

        public String UOMClassID
        {
            get { return m_sUOMClassID; }
            set { m_sUOMClassID = value; }
        }

        public String IUOM
        {
            get { return m_sIUOM; }
            set { m_sIUOM = value; }
        }

        public String PUOM
        {
            get { return m_sPUOM; }
            set { m_sPUOM = value; }
        }

        public String SUOM
        {
            get { return m_sSUOM; }
            set { m_sSUOM = value; }
        }

        public Decimal UnitNetWeight
        {
            get { return m_dUnitNetWeight; }
            set { m_dUnitNetWeight = value; }
        }

        public String WeightUOM
        {
            get { return m_sWeightUOM; }
            set { m_sWeightUOM = value; }
        }

        public String CountryofOrigin
        {
            get { return m_sCountryofOrigin; }
            set { m_sCountryofOrigin = value; }
        }

        public String HTS
        {
            get { return m_sHTS; }
            set { m_sHTS = value; }
        }

        public String CommodityCode
        {
            get { return m_sCommodityCode; }
            set { m_sCommodityCode = value; }
        }

        public int  CommodityCodeLength
        {
            get { return m_iCommodityCodeLength; }
            set { m_iCommodityCodeLength = value; }
        }
        public String ProductPortfolio
        {
            get { return m_sProductPortfolio; }
            set { m_sProductPortfolio = value; }
        }

        public String WarrantyCode
        {
            get { return m_sWarrantyCode; }
            set { m_sWarrantyCode = value; }
        }

        public String CreatedBy
        {
            get { return m_sCreatedBy; }
            set { m_sCreatedBy = value; }
        }

        public DateTime CreatedOn
        {
            get { return m_dtCreatedOn; }
            set { m_dtCreatedOn = value; }
        }

        public bool PartRunOut
        {
            get { return m_bPartRunOut; }
            set { m_bPartRunOut = value; }
        }

        public bool PartOnHold
        {
            get { return m_bPartOnHold; }
            set { m_bPartOnHold = value; }
        }

        public bool PartTrackLots
        {
            get { return m_bPartTrackLots; }
            set { m_bPartTrackLots = value; }
        }

        public bool TrackSerial
        {
            get { return m_bTrackSerial; }
            set { m_bTrackSerial = value; }
        }

        public bool UsePartRev
        {
            get { return m_bUsePartRev; }
            set { m_bUsePartRev = value; }
        }

        public bool Constrained
        {
            get { return m_bConstrained; }
            set { m_bConstrained = value; }
        }

        public bool InspectionRequired
        {
            get { return m_bInspectionRequired; }
            set { m_bInspectionRequired = value; }
        }

        public String Site
        {
            get { return m_sSite; }
            set { m_sSite = value; }
        }

        public bool PlantProcessMRP
        {
            get { return m_bPlantProcessMRP; }
            set { m_bPlantProcessMRP = value; }

        }

        public bool GenerateSuggestions
        {
            get { return m_bGenerateSuggestions; }
            set { m_bGenerateSuggestions = value; }

        }

        public bool PlantBackFlush
        {
            get { return m_bPlantBackFlush; }
            set { m_bPlantBackFlush = value; }

        }

        public bool BackflushKitComponents
        {
            get { return m_bBackflushKitComponents; }
            set { m_bBackflushKitComponents = value; }

        }

        public int KitTime
        {
            get { return m_iKitTime; }
            set { m_iKitTime = value; }

        }

        public bool GetFromLocalWarehouse
        {
            get { return m_bGetFromLocalWarehouse; }
            set { m_bGetFromLocalWarehouse = value; }

        }

        public decimal Maximum
        {
            get { return m_dMaximum; }
            set { m_dMaximum = value; }

        }

        public decimal Minimum
        {
            get { return m_dMinimum; }
            set { m_dMinimum = value; }

        }

        public decimal Safety
        {
            get { return m_dSafety; }
            set { m_dSafety = value; }

        }

        public decimal MinMfgLotSize
        {
            get { return m_dMinMfgLotSize; }
            set { m_dMinMfgLotSize = value; }

        }

        public decimal CostingLotSize
        {
            get { return m_dCostingLotSize; }
            set { m_dCostingLotSize = value; }

        }

        public decimal MfgLotSize
        {
            get { return m_dMfgLotSize; }
            set { m_dMfgLotSize = value; }

        }

        public decimal MfgLotMultiple
        {
            get { return m_dMfgLotMultiple; }
            set { m_dMfgLotMultiple = value; }

        }

        public int PreparationTime
        {
            get { return m_iPreparationTime; }
            set { m_iPreparationTime = value; }

        }

        public int DaysofSupply
        {
            get { return m_iDaysofSupply; }
            set { m_iDaysofSupply = value; }

        }

        public string PlannerID
        {
            get { return m_sPlannerID; }
            set { m_sPlannerID = value; }

        }

        public string PrimWhse
        {
            get { return m_sPrimWhse; }
            set { m_sPrimWhse = value; }

        }

        public string MfgComments
        {
            get { return m_sMfgComments; }
            set { m_sMfgComments = value; }

        }

        public string PartTypeCode
        {
            get { return m_sPartTypeCode; }
            set { m_sPartTypeCode = value; }

        }

        public string PlantSourceType
        {
            get { return m_sPlantSourceType; }
            set { m_sPlantSourceType = value; }

        }

        public bool PartPhantomBOM
        {
            get { return m_bPartPhantomBOM; }
            set { m_bPartPhantomBOM = value; }
        }

        public bool PlantPhantomBOM
        {
            get { return m_bPlantPhantomBOM; }
            set { m_bPlantPhantomBOM = value; }
        }

        public bool PlantNonStockItem
        {
            get { return m_bPlantNonStockItem; }
            set { m_bPlantNonStockItem = value; }

        }

        public bool PartNonStock
        {
            get { return m_bPartNonStock; }
            set { m_bPartNonStock = value; }

        }

        public bool PlantQtyBearing
        {
            get { return m_bPlantQtyBearing; }
            set { m_bPlantQtyBearing = value; }

        }

        public bool PartQtyBearing
        {
            get { return m_bPartQtyBearing; }
            set { m_bPartQtyBearing = value; }

        }

        public bool PlantBuyToOrder
        {
            get { return m_bPlantBuyToOrder; }
            set { m_bPlantBuyToOrder = value; }

        }

        public bool PartBuyToOrder
        {
            get { return m_bPartBuyToOrder; }
            set { m_bPartBuyToOrder = value; }

        }

        public bool PlantDropShip
        {
            get { return m_bPlantDropShip; }
            set { m_bPlantDropShip = value; }

        }

        public bool PartDropShip
        {
            get { return m_bPartDropShip; }
            set { m_bPartDropShip = value; }

        }

        public string PartCostingMethod
        {
            get { return m_sPartCostingMethod; }
            set { m_sPartCostingMethod = value; }

        }

        public string PlantCostingMethod
        {
            get { return m_sPlantCostingMethod; }
            set { m_sPlantCostingMethod = value; }

        }

        public decimal Cost
        {
            get { return m_dCost; }
            set { m_dCost = value; }

        }

        public string PricePer
        {
            get { return m_sPricePer; }
            set { m_sPricePer = value; }

        }

        public string InternalDivisionalPricePer
        {
            get { return m_sInternalDivisionalPricePer; }
            set { m_sInternalDivisionalPricePer = value; }
        }

        public Decimal UnitPrice
        {
            get { return m_dUnitPrice; }
            set { m_dUnitPrice = value; }

        }

        public string PurchaseComments
        {
            get { return m_sPurchaseComments; }
            set { m_sPurchaseComments = value; }
        }

        public string BuyerID
        {
            get { return m_sBuyerID; }
            set { m_sBuyerID = value; }
        }

        public string VendorName
        {
            get { return m_sVendorName; }
            set { m_sVendorName = value; }
        }

        public string SupplierID
        {
            get { return m_sSupplierID; }
            set { m_sSupplierID = value; }
        }

        public decimal MinOrderQty
        {
            get { return m_dMinOrderQty; }
            set { m_dMinOrderQty = value; }
        }

        public int SupplierNum
        {
            get { return m_iSupplierNum; }
            set { m_iSupplierNum = value; }
        }

        public int Lead
        {
            get { return m_iLead; }
            set { m_iLead = value; }
        }

        public decimal LastMaterialCost
        {
            get { return m_dLastMaterialCost; }
            set { m_dLastMaterialCost = value; }
        }

        public decimal VendorBaseUnitPrice
        {
            get { return m_dVendorBaseUnitPrice; }
            set { m_dVendorBaseUnitPrice = value; }
        }

        public DateTime VendorPriceEffectiveDate
        {
            get { return m_dtVendorPriceEffectiveDate; }
            set { m_dtVendorPriceEffectiveDate = value; }
        }

        public DateTime VendorPriceExpirationDate
        {
            get { return m_dtVendorPriceExpirationDate; }
            set { m_dtVendorPriceExpirationDate = value; }
        }

        public string MinABC
        {
            get { return m_sMinABC; }
            set { m_sMinABC = value; }
        }

        public string ActualABC
        {
            get { return m_sActualABC; }
            set { m_sActualABC = value; }
        }

        public bool OverrideCountFrequency
        {
            get { return m_bOverrideCountFrequency; }
            set { m_bOverrideCountFrequency = value; }
        }

        public int CountFrequency
        {
            get { return m_iCountFrequency; }
            set { m_iCountFrequency = value; }
        }

        public DateTime LastCycleCountDate
        {
            get { return m_dtLastCycleCountDate; }
            set { m_dtLastCycleCountDate = value; }
        }

        public bool ManualABCCode
        {
            get { return m_bManualABCCode; }
            set { m_bManualABCCode = value; }
        }

        #endregion

        #region Data Members
        private string m_sCompany = "";
        private string m_sPart = "";
        private string m_sDescription = "";
        private bool m_bInactive;
        private string m_sClassID = "";
        private string m_sGroup = "";
        private string m_sSearch = "";
        private string m_sUOMClassID = "";
        private string m_sIUOM = "";
        private string m_sPUOM = "";
        private string m_sSUOM = "";
        private Decimal m_dUnitNetWeight;
        private string m_sWeightUOM = "";
        private string m_sCountryofOrigin = "";
        private string m_sHTS = "";
        private string m_sCommodityCode = "";
        private int m_iCommodityCodeLength;
        private string m_sProductPortfolio = "";
        private string m_sWarrantyCode = "";
        private string m_sCreatedBy = "";
        private DateTime m_dtCreatedOn;

        private bool m_bPartRunOut;
        private bool m_bPartOnHold;
        private bool m_bPartTrackLots;
        private bool m_bTrackSerial;
        private bool m_bUsePartRev;
        private bool m_bConstrained;
        private bool m_bInspectionRequired;

        private string m_sSite = "";
        private bool m_bPlantProcessMRP;
        private bool m_bGenerateSuggestions;
        private bool m_bPlantBackFlush;
        private bool m_bBackflushKitComponents;
        private int m_iKitTime;
        private bool m_bGetFromLocalWarehouse;
        private Decimal m_dMaximum;
        private Decimal m_dMinimum;
        private Decimal m_dSafety;
        private Decimal m_dCostingLotSize;
        private Decimal m_dMinMfgLotSize;
        private Decimal m_dMfgLotSize;
        private decimal m_dMfgLotMultiple;
        private int m_iPreparationTime;
        private int m_iDaysofSupply;
        private string m_sPlannerID = "";
        private string m_sPrimWhse = "";
        private string m_sMfgComments = "";

        private string m_sPartTypeCode = "";
        private string m_sPlantSourceType;

        private bool m_bPartPhantomBOM;
        private bool m_bPlantPhantomBOM;

        private bool m_bPlantNonStockItem;
        private bool m_bPartNonStock;

        private bool m_bPlantQtyBearing;
        private bool m_bPartQtyBearing;

        private bool m_bPlantBuyToOrder;
        private bool m_bPartBuyToOrder;

        private bool m_bPlantDropShip;
        private bool m_bPartDropShip;

        private string m_sPartCostingMethod = "";
        private string m_sPlantCostingMethod;
        private decimal m_dCost;
        private string m_sPricePer = "";
        private string m_sInternalDivisionalPricePer = "";
        private Decimal m_dUnitPrice;

        private string m_sPurchaseComments = "";
        private string m_sBuyerID = "";
        private string m_sVendorName = "";
        private string m_sSupplierID = "";
        private Decimal m_dMinOrderQty;
        private int m_iSupplierNum;
        private int m_iLead;
        private decimal m_dLastMaterialCost;
        private decimal m_dVendorBaseUnitPrice;
        private DateTime m_dtVendorPriceEffectiveDate;
        private DateTime m_dtVendorPriceExpirationDate;

        private string m_sMinABC = "";
        private string m_sActualABC = "";
        private bool m_bOverrideCountFrequency;
        private int m_iCountFrequency;
        private DateTime m_dtLastCycleCountDate;
        private bool m_bManualABCCode;
        #endregion
    }

    public class PartCOO
    {
        #region Constructors

        public PartCOO()
        {
        }

        public PartCOO(DataRow oDataRow)
        {
            if (oDataRow["Part_Company"] != DBNull.Value)
            {
                m_sCompany = (string)oDataRow["Part_Company"];
            }
            if (oDataRow["Part_PartNum"] != DBNull.Value)
            {
                m_sPartNum = (string)oDataRow["Part_PartNum"];
            }
            if (oDataRow["Part_PartDescription"] != DBNull.Value)
            {
                m_sPartDescription = (string)oDataRow["Part_PartDescription"];
            }
            if (oDataRow["Part_TypeCode"] != DBNull.Value)
            {
                m_sTypeCode = (string)oDataRow["Part_TypeCode"];
            }
            if (oDataRow["Part_ClassID"] != DBNull.Value)
            {
                m_sClassId = (string)oDataRow["Part_ClassID"];
            }
            if (oDataRow["Calculated_NoPartCountriesSet"] != DBNull.Value)
            {
                m_bNoPartCountriesSet = (bool)oDataRow["Calculated_NoPartCountriesSet"];
            }
            if (oDataRow["Calculated_BadQtyPerc"] != DBNull.Value)
            {
                m_bBadQtyPercent = (bool)oDataRow["Calculated_BadQtyPerc"];
            }
            if (oDataRow["Calculated_BadValuePerc"] != DBNull.Value)
            {
                m_bBadValuePercent = (bool)oDataRow["Calculated_BadValuePerc"];
            }
            if (oDataRow["Calculated_NoPrimaryCountriesSet"] != DBNull.Value)
            {
                m_bNoPrimaryCountrySet = (bool)oDataRow["Calculated_NoPrimaryCountriesSet"];
            }
            if (oDataRow["Calculated_MultiplePrimaryCountriesSet"] != DBNull.Value)
            {
                m_bMultiplePrimaryCountriesSet = (bool)oDataRow["Calculated_MultiplePrimaryCountriesSet"];
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

        public string PartDescription
        {
            get { return m_sPartDescription; }
            set { m_sPartDescription = value; }
        }

        public string TypeCode
        {
            get { return m_sTypeCode; }
            set { m_sTypeCode = value; }
        }

        public string ClassId
        {
            get { return m_sClassId; }
            set { m_sClassId = value; }
        }

        public bool NoPartCountriesSet
        {
            get { return m_bNoPartCountriesSet; }
            set { m_bNoPartCountriesSet = value; }
        }

        public bool BadQtyPercent
        {
            get { return m_bBadQtyPercent; }
            set { m_bBadQtyPercent = value; }
        }

        public bool BadValuePercent
        {
            get { return m_bBadValuePercent; }
            set { m_bBadValuePercent = value; }
        }

        public bool NoPrimaryCountrySet
        {
            get { return m_bNoPrimaryCountrySet; }
            set { m_bNoPrimaryCountrySet = value; }
        }

        public bool MultiplePrimaryCountriesSet
        {
            get { return m_bMultiplePrimaryCountriesSet; }
            set { m_bMultiplePrimaryCountriesSet = value; }
        }

        #endregion

        #region Data Members
        private string m_sCompany;
        private string m_sPartNum;
        private string m_sPartDescription;
        private string m_sTypeCode;
        private string m_sClassId;
        private bool m_bNoPartCountriesSet;
        private bool m_bBadQtyPercent;
        private bool m_bBadValuePercent;
        private bool m_bNoPrimaryCountrySet;
        private bool m_bMultiplePrimaryCountriesSet;
        #endregion
    }

    public class HSValidateParts
    {
        #region Methods

        public bool Initialize(Session oSession)
        {
            bool bSuccess = true;

            //****Set a parameter Value***** 
            Ice.Proxy.BO.DynamicQueryImpl oDynamicQuery = WCFServiceSupport.CreateImpl<Ice.Proxy.BO.DynamicQueryImpl>(oSession, Ice.Proxy.BO.DynamicQueryImpl.UriPath);
            QueryExecutionDataSet oQueryExecutionDataSet = oDynamicQuery.GetQueryExecutionParametersByID(BAQConstants.QUERY_LIST_ALL_PARTS);
            oQueryExecutionDataSet.ExecutionParameter.Clear();
            DataSet oDataSet = oDynamicQuery.ExecuteByID(BAQConstants.QUERY_LIST_ALL_PARTS, oQueryExecutionDataSet);

            m_oAllParts.Clear();
            foreach (DataRow oRow in oDataSet.Tables[0].Rows)
            {
                HSPartData oTmpPart = new HSPartData(oRow);
                m_oAllParts.Add(oTmpPart);
                if (m_oFastParts.ContainsKey(oTmpPart.PartNum) == false)
                {
                    m_oFastParts[oTmpPart.PartNum] = oTmpPart;
                }
            }

            return bSuccess;
        }

        public HSPartData GetPart(string sPartNum)
        {
            HSPartData oPart = null;
            if (m_oFastParts.ContainsKey(sPartNum) == true)
            {
                oPart = m_oFastParts[sPartNum];
            }
            return oPart;
        }

        public void PerformPurchasePartValidation(string sCompany, string sTmpFileDirectory)
        {
            // get the file name
            DateTime dtToday = DateTime.Now;
            string sDate = dtToday.Month.ToString() + "-" + dtToday.Day.ToString() + "-" + dtToday.Year.ToString();

            // only get purchased parts that are active
            List<HSPartData> oAllActivePurchasedParts = m_oAllParts.Where(oItem => (string.Compare(oItem.PartTypeCode, "P", true) == 0) && (oItem.Inactive == false)).ToList();

            HSUser oServiceAccount = HSUser.GetUserById(HSUser.SPINCRAFT_SERVICE_ACCOUNT_ID);

            List<string> oToAddresses = new List<string>();
            // get users in the purchasing, production
            HSUser.GetEmailsForUsersInGroup(oToAddresses, HSUser.REPORT_ON_PURCHASED_PART_ISSUES);
            oToAddresses.Add(HSUser.SPINCRAFT_ROOT_USER);

            System.Drawing.Color oHighlightYellow = System.Drawing.Color.FromArgb(255, 255, 0);

            SLStyle oHighlightHeaderStyle = new SLStyle();
            oHighlightHeaderStyle.SetFontBold(true);
            oHighlightHeaderStyle.SetFont(FontSchemeValues.Major, 12);
            oHighlightHeaderStyle.SetHorizontalAlignment(HorizontalAlignmentValues.Center);
            oHighlightHeaderStyle.SetPatternFill(PatternValues.Solid, oHighlightYellow, oHighlightYellow);
            oHighlightHeaderStyle.SetFontColor(SLThemeColorIndexValues.Dark2Color);

            int iNumberOfRows = 1;
            int iNumberOfColumns = 1;
            bool bDataInReport = false;
            bool bFirstWorksheet = true;

            #region Purchased Part 
            string sDestinationFileName = sTmpFileDirectory + "\\" +sCompany +  "-PurchasedPartReport-" + sDate + ".xlsx";
            SLDocument oSLPPDocument = new SLDocument();
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
            bDataInReport = false;
            bFirstWorksheet = true;
            oToAddresses = HSUser.GetEmailsForUsersInGroup(HSUser.REPORT_ON_PURCHASED_PART_ISSUES);
            oToAddresses.Add(HSUser.SPINCRAFT_ROOT_USER);

            //
            // no primary warehouse
            //
            iNumberOfRows = 1;
            iNumberOfColumns = 1;
            List<HSPartData> oNoPrimaryWarehouse = oAllActivePurchasedParts.Where(oItem => (string.IsNullOrEmpty(oItem.PrimWhse) == true)).ToList();
            if (oNoPrimaryWarehouse.Count > 0)
            {
                if (bFirstWorksheet == true)
                {
                    oSLPPDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "No Primary Warehouse");
                    bFirstWorksheet = false;
                }
                else
                {
                    oSLPPDocument.AddWorksheet("No Primary Warehouse");
                }
                //set column header
                oSLPPDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Part Num");
                oSLPPDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 30);
                oSLPPDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Description");
                oSLPPDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 60);
                oSLPPDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Class");
                oSLPPDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 30);
                oSLPPDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "The Purchased Part Should Be Assigned To A Primary Warehouse");
                oSLPPDocument.SetCellStyle(iNumberOfRows++, iNumberOfColumns, oHighlightHeaderStyle);
                oSLPPDocument.SetColumnWidth(iNumberOfColumns++, 90);

                foreach (HSPartData oTmpPart in oNoPrimaryWarehouse)
                {
                    oSLPPDocument.SetCellValue(iNumberOfRows, 1, oTmpPart.PartNum);
                    oSLPPDocument.SetCellValue(iNumberOfRows, 2, StringExt.Truncate(oTmpPart.Description, 60));
                    oSLPPDocument.SetCellValue(iNumberOfRows, 3, oTmpPart.ClassID);

                    iNumberOfRows++;
                    bDataInReport = true;
                }
            }

            //
            // no backflush
            //
            //iNumberOfRows = 1;
            //iNumberOfColumns = 1;
            //{
            //    List<HSPart> oNoBackflush = oAllActivePurchasedParts.Where(oItem => (oItem.PlantBackFlush == false) && (oItem.PartDropShip == false) && (oItem.PartBuyToOrder == false) && (oItem.TrackSerial == false)).ToList();
            //    if (oNoBackflush.Count > 0)
            //    {
            //        if (bFirstWorksheet == true)
            //        {
            //            oSLPPDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "No Backflush");
            //            bFirstWorksheet = false;
            //        }
            //        else
            //        {
            //            oSLPPDocument.AddWorksheet("No Backflush");
            //        }
            //        //set column header
            //        oSLPPDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Part Num");
            //        oSLPPDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 30);
            //        oSLPPDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Description");
            //        oSLPPDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 60);
            //        oSLPPDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "The Purchased Part Should Be Set To Be Backflushed");
            //        oSLPPDocument.SetCellStyle(iNumberOfRows++, iNumberOfColumns, oHighlightHeaderStyle);
            //        oSLPPDocument.SetColumnWidth(iNumberOfColumns++, 90);
            //
            //        foreach (HSPart oTmpPart in oNoBackflush)
            //        {
            //          oSLPPDocument.SetCellValue(iNumberOfRows, 1, oTmpPart.Part);
            //          oSLPPDocument.SetCellValue(iNumberOfRows, 2,  StringExt.Truncate(oTmpPart.Description, 60));
            //          oSLPPDocument.SetCellValue(iNumberOfRows, 3, oTmpPart.ClassID);
            //
            //              iNumberOfRows++;
            //            bDataInReport = true;
            //        }
            //    }
            //}

            //
            // no ABC code set
            //
            iNumberOfRows = 1;
            iNumberOfColumns = 1;
            List<HSPartData> oNoABCCode = oAllActivePurchasedParts.Where(oItem => string.IsNullOrEmpty(oItem.ActualABC) == true).ToList();
            if (oNoABCCode.Count > 0)
            {
                if (bFirstWorksheet == true)
                {
                    oSLPPDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "No ABC Code");
                    bFirstWorksheet = false;
                }
                else
                {
                    oSLPPDocument.AddWorksheet("No ABC Code");
                }
                //set column header
                oSLPPDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Part Num");
                oSLPPDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 30);
                oSLPPDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Description");
                oSLPPDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 60);
                oSLPPDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Class");
                oSLPPDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 30);
                oSLPPDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "The Purchased Part Should Have An ABC Code Assigned");
                oSLPPDocument.SetCellStyle(iNumberOfRows++, iNumberOfColumns, oHighlightHeaderStyle);
                oSLPPDocument.SetColumnWidth(iNumberOfColumns++, 90);

                foreach (HSPartData oTmpPart in oNoABCCode)
                {
                    oSLPPDocument.SetCellValue(iNumberOfRows, 1, oTmpPart.PartNum);
                    oSLPPDocument.SetCellValue(iNumberOfRows, 2, StringExt.Truncate(oTmpPart.Description, 60));
                    oSLPPDocument.SetCellValue(iNumberOfRows, 3, oTmpPart.ClassID);

                    iNumberOfRows++;
                    bDataInReport = true;
                }
            }

            ////
            //// no buyer
            ////
            //iNumberOfRows = 1;
            //iNumberOfColumns = 1;
            //List<HSPart> oNoBuyer = oAllActivePurchasedParts.Where(oItem => string.IsNullOrEmpty(oItem.BuyerID) == true).ToList();
            //if (oNoBuyer.Count > 0)
            //{
            //    if (bFirstWorksheet == true)
            //    {
            //        oSLPPDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "No Buyer");
            //        bFirstWorksheet = false;
            //    }
            //    else
            //    {
            //        oSLPPDocument.AddWorksheet("No Buyer");
            //    }
            //    //set column header
            //    oSLPPDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Part Num");
            //    oSLPPDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 30);
            //    oSLPPDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Description");
            //    oSLPPDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 60);
            //    oSLPPDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Class");
            //    oSLPPDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 30);
            //    oSLPPDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "The Purchased Part Needs To Have A Primary Buyer Set");
            //    oSLPPDocument.SetCellStyle(iNumberOfRows++, iNumberOfColumns, oHighlightHeaderStyle);
            //    oSLPPDocument.SetColumnWidth(iNumberOfColumns++, 90);
            //
            //    foreach (HSPart oTmpPart in oNoBuyer)
            //    {
            //        oSLPPDocument.SetCellValue(iNumberOfRows, 1, oTmpPart.Part);
            //        oSLPPDocument.SetCellValue(iNumberOfRows, 2,  StringExt.Truncate(oTmpPart.Description, 60));
            //        oSLPPDocument.SetCellValue(iNumberOfRows, 3, oTmpPart.ClassID);
            //
            //        iNumberOfRows++;
            //        bDataInReport = true;
            //    }
            //}

            //
            // no suggestions
            //
            iNumberOfRows = 1;
            iNumberOfColumns = 1;
            List<HSPartData> oNoSuggestions = oAllActivePurchasedParts.Where(oItem => (oItem.GenerateSuggestions == false) && (oItem.PartQtyBearing == true)).ToList();
            if (oNoSuggestions.Count > 0)
            {
                if (bFirstWorksheet == true)
                {
                    oSLPPDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "Qty Bearing No Suggestions");
                    bFirstWorksheet = false;
                }
                else
                {
                    oSLPPDocument.AddWorksheet("Qty Bearing No Suggestions");
                }
                //set column header
                oSLPPDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Part Num");
                oSLPPDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 30);
                oSLPPDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Description");
                oSLPPDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 60);
                oSLPPDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Class");
                oSLPPDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 30);
                oSLPPDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "The Purchased Part Is Quantity Bearing But Has PO Suggestions Turned Off");
                oSLPPDocument.SetCellStyle(iNumberOfRows++, iNumberOfColumns, oHighlightHeaderStyle);
                oSLPPDocument.SetColumnWidth(iNumberOfColumns++, 90);

                foreach (HSPartData oTmpPart in oNoSuggestions)
                {
                    oSLPPDocument.SetCellValue(iNumberOfRows, 1, oTmpPart.PartNum);
                    oSLPPDocument.SetCellValue(iNumberOfRows, 2, StringExt.Truncate(oTmpPart.Description, 60));
                    oSLPPDocument.SetCellValue(iNumberOfRows, 3, oTmpPart.ClassID);

                    iNumberOfRows++;
                    bDataInReport = true;
                }
            }

            //
            // no primary bin
            //

            //
            // negative qty
            //

            //
            // no vendor
            //
            iNumberOfRows = 1;
            iNumberOfColumns = 1;
            List<HSPartData> oNoVendor = oAllActivePurchasedParts.Where(oItem => (oItem.SupplierNum == 0)).ToList();
            if (oNoVendor.Count > 0)
            {
                if (bFirstWorksheet == true)
                {
                    oSLPPDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "No Vendor");
                    bFirstWorksheet = false;
                }
                else
                {
                    oSLPPDocument.AddWorksheet("No Vendor");
                }
                //set column header
                oSLPPDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Part Num");
                oSLPPDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 30);
                oSLPPDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Description");
                oSLPPDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 60);
                oSLPPDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Class");
                oSLPPDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 30);
                oSLPPDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "The Purchased Part Needs To Have A Primary Supplier Set");
                oSLPPDocument.SetCellStyle(iNumberOfRows++, iNumberOfColumns, oHighlightHeaderStyle);
                oSLPPDocument.SetColumnWidth(iNumberOfColumns++, 90);

                foreach (HSPartData oTmpPart in oNoVendor)
                {
                    oSLPPDocument.SetCellValue(iNumberOfRows, 1, oTmpPart.PartNum);
                    oSLPPDocument.SetCellValue(iNumberOfRows, 2, StringExt.Truncate(oTmpPart.Description, 60));
                    oSLPPDocument.SetCellValue(iNumberOfRows, 3, oTmpPart.ClassID);

                    iNumberOfRows++;
                    bDataInReport = true;
                }
            }


            //
            // no approved vendor
            //


            //
            // no vendor pricing
            //
            iNumberOfRows = 1;
            iNumberOfColumns = 1;
            List<HSPartData> oNoVendorPricing = oAllActivePurchasedParts.Where(oItem => (oItem.VendorBaseUnitPrice == 0)).ToList();
            if (oNoVendorPricing.Count > 0)
            {
                if (bFirstWorksheet == true)
                {
                    oSLPPDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "No Vendor Pricing");
                    bFirstWorksheet = false;
                }
                else
                {
                    oSLPPDocument.AddWorksheet("No Vendor Pricing");
                }
                //set column header
                oSLPPDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Part Num");
                oSLPPDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 30);
                oSLPPDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Description");
                oSLPPDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 60);
                oSLPPDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Class");
                oSLPPDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 30);
                oSLPPDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "The Purchased Part Has No Vendor Pricing Set");
                oSLPPDocument.SetCellStyle(iNumberOfRows++, iNumberOfColumns, oHighlightHeaderStyle);
                oSLPPDocument.SetColumnWidth(iNumberOfColumns++, 90);

                foreach (HSPartData oTmpPart in oNoVendorPricing)
                {
                    oSLPPDocument.SetCellValue(iNumberOfRows, 1, oTmpPart.PartNum);
                    oSLPPDocument.SetCellValue(iNumberOfRows, 2, StringExt.Truncate(oTmpPart.Description, 60));
                    oSLPPDocument.SetCellValue(iNumberOfRows, 3, oTmpPart.ClassID);

                    iNumberOfRows++;
                    bDataInReport = true;
                }
            }


            //
            // no established cost
            //
            iNumberOfRows = 1;
            iNumberOfColumns = 1;
            List<HSPartData> oNoEstablishedCost = oAllActivePurchasedParts.Where(oItem => (oItem.PartQtyBearing == true) && (oItem.Cost == 0)).ToList();
            if (oNoEstablishedCost.Count > 0)
            {
                if (bFirstWorksheet == true)
                {
                    oSLPPDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "No Estblished Cost");
                    bFirstWorksheet = false;
                }
                else
                {
                    oSLPPDocument.AddWorksheet("No Estblished Cost");
                }
                //set column header
                oSLPPDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Part Num");
                oSLPPDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 30);
                oSLPPDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Description");
                oSLPPDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 60);
                oSLPPDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Class");
                oSLPPDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 30);
                oSLPPDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "The Purchased Part Does Not Have An Established Cost");
                oSLPPDocument.SetCellStyle(iNumberOfRows++, iNumberOfColumns, oHighlightHeaderStyle);
                oSLPPDocument.SetColumnWidth(iNumberOfColumns++, 90);

                foreach (HSPartData oTmpPart in oNoEstablishedCost)
                {
                    oSLPPDocument.SetCellValue(iNumberOfRows, 1, oTmpPart.PartNum);
                    oSLPPDocument.SetCellValue(iNumberOfRows, 2, StringExt.Truncate(oTmpPart.Description, 60));
                    oSLPPDocument.SetCellValue(iNumberOfRows, 3, oTmpPart.ClassID);

                    iNumberOfRows++;
                    bDataInReport = true;
                }
            }


            //
            // no approved vendor pricing
            //


            //
            // no lead time
            //
            iNumberOfRows = 1;
            iNumberOfColumns = 1;
            List<HSPartData> oNoLeadTime = oAllActivePurchasedParts.Where(oItem => (oItem.Lead == 0)).ToList();
            if (oNoLeadTime.Count > 0)
            {
                if (bFirstWorksheet == true)
                {
                    oSLPPDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "No Lead Time");
                    bFirstWorksheet = false;
                }
                else
                {
                    oSLPPDocument.AddWorksheet("No Lead Time");
                }
                //set column header
                oSLPPDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Part Num");
                oSLPPDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 30);
                oSLPPDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Description");
                oSLPPDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 60);
                oSLPPDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Class");
                oSLPPDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 30);
                oSLPPDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "The Purchased Part Needs To Have A Lead Time Set");
                oSLPPDocument.SetCellStyle(iNumberOfRows++, iNumberOfColumns, oHighlightHeaderStyle);
                oSLPPDocument.SetColumnWidth(iNumberOfColumns++, 90);

                foreach (HSPartData oTmpPart in oNoLeadTime)
                {
                    oSLPPDocument.SetCellValue(iNumberOfRows, 1, oTmpPart.PartNum);
                    oSLPPDocument.SetCellValue(iNumberOfRows, 2, StringExt.Truncate(oTmpPart.Description, 60));
                    oSLPPDocument.SetCellValue(iNumberOfRows, 3, oTmpPart.ClassID);

                    iNumberOfRows++;
                    bDataInReport = true;
                }
            }


            //
            // inconsistent drop ship
            //
            iNumberOfRows = 1;
            iNumberOfColumns = 1;
            List<HSPartData> oInconsistentDropShip = oAllActivePurchasedParts.Where(oItem => (oItem.PartDropShip != oItem.PlantDropShip)).ToList();
            if (oInconsistentDropShip.Count > 0)
            {
                if (bFirstWorksheet == true)
                {
                    oSLPPDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "Inconsistent Drop Ship");
                    bFirstWorksheet = false;
                }
                else
                {
                    oSLPPDocument.AddWorksheet("Inconsistent Drop Ship");
                }
                //set column header
                oSLPPDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Part Num");
                oSLPPDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 30);
                oSLPPDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Description");
                oSLPPDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 60);
                oSLPPDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Class");
                oSLPPDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 30);
                oSLPPDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "The Purchased Part And Part Site Have Different Drop Ship Settings");
                oSLPPDocument.SetCellStyle(iNumberOfRows++, iNumberOfColumns, oHighlightHeaderStyle);
                oSLPPDocument.SetColumnWidth(iNumberOfColumns++, 90);

                foreach (HSPartData oTmpPart in oInconsistentDropShip)
                {
                    oSLPPDocument.SetCellValue(iNumberOfRows, 1, oTmpPart.PartNum);
                    oSLPPDocument.SetCellValue(iNumberOfRows, 2, StringExt.Truncate(oTmpPart.Description, 60));
                    oSLPPDocument.SetCellValue(iNumberOfRows, 3, oTmpPart.ClassID);

                    iNumberOfRows++;
                    bDataInReport = true;
                }
            }


            //
            // inconsistent buy to order
            //
            iNumberOfRows = 1;
            iNumberOfColumns = 1;
            List<HSPartData> oInconsistentBuyToOrder = oAllActivePurchasedParts.Where(oItem => (oItem.PartBuyToOrder != oItem.PlantBuyToOrder)).ToList();
            if (oInconsistentBuyToOrder.Count > 0)
            {
                if (bFirstWorksheet == true)
                {
                    oSLPPDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "Inconsistent Buy To Order");
                    bFirstWorksheet = false;
                }
                else
                {
                    oSLPPDocument.AddWorksheet("Inconsistent Buy To Order");
                }
                //set column header
                oSLPPDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Part Num");
                oSLPPDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 30);
                oSLPPDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Description");
                oSLPPDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 60);
                oSLPPDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Class");
                oSLPPDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 30);
                oSLPPDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "The Purchased Part And Part Site Have Different Buy To Order Settings");
                oSLPPDocument.SetCellStyle(iNumberOfRows++, iNumberOfColumns, oHighlightHeaderStyle);
                oSLPPDocument.SetColumnWidth(iNumberOfColumns++, 90);

                foreach (HSPartData oTmpPart in oInconsistentBuyToOrder)
                {
                    oSLPPDocument.SetCellValue(iNumberOfRows, 1, oTmpPart.PartNum);
                    oSLPPDocument.SetCellValue(iNumberOfRows, 2, StringExt.Truncate(oTmpPart.Description, 60));
                    oSLPPDocument.SetCellValue(iNumberOfRows, 3, oTmpPart.ClassID);

                    iNumberOfRows++;
                    bDataInReport = true;
                }
            }


            //
            // inconsistent qty bearing
            //
            iNumberOfRows = 1;
            iNumberOfColumns = 1;
            List<HSPartData> oInconsistentQtyBearing = oAllActivePurchasedParts.Where(oItem => (oItem.PartQtyBearing != oItem.PlantQtyBearing)).ToList();
            if (oInconsistentQtyBearing.Count > 0)
            {
                if (bFirstWorksheet == true)
                {
                    oSLPPDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "Inconsistent Qty Bearing");
                    bFirstWorksheet = false;
                }
                else
                {
                    oSLPPDocument.AddWorksheet("Inconsistent Qty Bearing");
                }
                //set column header
                oSLPPDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Part Num");
                oSLPPDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 30);
                oSLPPDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Description");
                oSLPPDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 60);
                oSLPPDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Class");
                oSLPPDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 30);
                oSLPPDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "The Purchased Part And Part Site Have Different Quantity Bearing Settings");
                oSLPPDocument.SetCellStyle(iNumberOfRows++, iNumberOfColumns, oHighlightHeaderStyle);
                oSLPPDocument.SetColumnWidth(iNumberOfColumns++, 90);

                foreach (HSPartData oTmpPart in oInconsistentQtyBearing)
                {
                    oSLPPDocument.SetCellValue(iNumberOfRows, 1, oTmpPart.PartNum);
                    oSLPPDocument.SetCellValue(iNumberOfRows, 2, StringExt.Truncate(oTmpPart.Description, 60));
                    oSLPPDocument.SetCellValue(iNumberOfRows, 3, oTmpPart.ClassID);

                    iNumberOfRows++;
                    bDataInReport = true;
                }
            }


            //
            // buy to order - no suggestions
            //
            iNumberOfRows = 1;
            iNumberOfColumns = 1;
            List<HSPartData> oBuyToOrderNoSuggestions = oAllActivePurchasedParts.Where(oItem => (oItem.PartBuyToOrder == true) && (oItem.GenerateSuggestions == false)).ToList();
            if (oBuyToOrderNoSuggestions.Count > 0)
            {
                if (bFirstWorksheet == true)
                {
                    oSLPPDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "Buy To Order - No Suggestions");
                    bFirstWorksheet = false;
                }
                else
                {
                    oSLPPDocument.AddWorksheet("Buy To Order - No Suggestions");
                }
                //set column header
                oSLPPDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Part Num");
                oSLPPDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 30);
                oSLPPDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Description");
                oSLPPDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 60);
                oSLPPDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Class");
                oSLPPDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 30);
                oSLPPDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "The Purchased Part Is Set To Buy-To-Order But Generate Suggestions Is Off");
                oSLPPDocument.SetCellStyle(iNumberOfRows++, iNumberOfColumns, oHighlightHeaderStyle);
                oSLPPDocument.SetColumnWidth(iNumberOfColumns++, 90);

                foreach (HSPartData oTmpPart in oBuyToOrderNoSuggestions)
                {
                    oSLPPDocument.SetCellValue(iNumberOfRows, 1, oTmpPart.PartNum);
                    oSLPPDocument.SetCellValue(iNumberOfRows, 2, StringExt.Truncate(oTmpPart.Description, 60));
                    oSLPPDocument.SetCellValue(iNumberOfRows, 3, oTmpPart.ClassID);

                    iNumberOfRows++;
                    bDataInReport = true;
                }
            }

            //
            // no supplier price established
            //

            if (bDataInReport == true)
            {
                oSLPPDocument.SaveAs(sDestinationFileName);
                // Check to see if we created a file and if so email it
                if (File.Exists(sDestinationFileName) == true)
                {
                    List<string> oAttachments = new List<string>();
                    oAttachments.Add(sDestinationFileName);

                    HSEmailHelper.SendEmail(oToAddresses, "Purchased Parts Report", "Purchased Parts Report for " + sDate, oAttachments);
                }
            }
            #endregion

            #region Accounting Issues
            sDestinationFileName = sTmpFileDirectory + "\\" + sCompany + "-PurchasedPartAccountingReport-" + sDate + ".xlsx";
            SLDocument oSLActDocument = new SLDocument();
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
            bDataInReport = false;
            bFirstWorksheet = true;
            oToAddresses = HSUser.GetEmailsForUsersInGroup(HSUser.REPORT_ON_PURCHASED_PART_ISSUES);
            oToAddresses.Add(HSUser.SPINCRAFT_ROOT_USER);

            //
            // inconsistent cost method
            //
            iNumberOfRows = 1;
            iNumberOfColumns = 1;
            List<HSPartData> oInconsistentCostMethod = oAllActivePurchasedParts.Where(oItem => (string.Compare(oItem.PartCostingMethod, oItem.PlantCostingMethod, true) != 0)).ToList();
            if (oInconsistentCostMethod.Count > 0)
            {
                if (bFirstWorksheet == true)
                {
                    oSLActDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "Inconsistent Costing Method");
                    bFirstWorksheet = false;
                }
                else
                {
                    oSLActDocument.AddWorksheet("Inconsistent Costing Method");
                }
                //set column header
                oSLActDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Part Num");
                oSLActDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 30);
                oSLActDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Part Cost Method");
                oSLActDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 30);
                oSLActDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Plant Cost Method");
                oSLActDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 30);
                oSLActDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "The Purchased Part And Part Site Have Different Costing Methods");
                oSLActDocument.SetCellStyle(iNumberOfRows++, iNumberOfColumns, oHighlightHeaderStyle);
                oSLActDocument.SetColumnWidth(iNumberOfColumns++, 90);

                foreach (HSPartData oTmpPart in oInconsistentCostMethod)
                {
                    oSLActDocument.SetCellValue(iNumberOfRows, 1, oTmpPart.PartNum);
                    oSLActDocument.SetCellValue(iNumberOfRows, 2, oTmpPart.PartCostingMethod);
                    oSLActDocument.SetCellValue(iNumberOfRows, 3, oTmpPart.PlantCostingMethod);

                    iNumberOfRows++;
                    bDataInReport = true;
                }
            }

            //
            // check costing method -- could be different for each site
            //
            iNumberOfRows = 1;
            iNumberOfColumns = 1;
            List<HSPartData> oBadCostingMethod = new List<HSPartData>();
            if (string.Compare(sCompany, CompanyConfiguration.SPINCRAFT_MA_COMPANY_ID, true) == 0)
            {
                // ETG MA SHOULD BE LAST
                oBadCostingMethod = oAllActivePurchasedParts.Where(oItem => (string.Compare(oItem.PartCostingMethod, "L", true) != 0)).ToList();
            }
            else if (string.Compare(sCompany, CompanyConfiguration.SPINCRAFT_WI_COMPANY_ID, true) == 0)
            {
                // ETG WI SHOULD BE AVERAGE OR STANDARD
                oBadCostingMethod = oAllActivePurchasedParts.Where(oItem => (string.Compare(oItem.PartCostingMethod, "A", true) != 0) && (string.Compare(oItem.PartCostingMethod, "S", true) != 0)).ToList();
            }
            else if (string.Compare(sCompany, CompanyConfiguration.SPINCRAFT_UK_COMPANY_ID, true) == 0)
            {
                // ETG UK SHOULD BE FIFO
                oBadCostingMethod = oAllActivePurchasedParts.Where(oItem => (string.Compare(oItem.PartCostingMethod, "F", true) != 0)).ToList();
            }
            if (oBadCostingMethod.Count > 0)
            {
                if (bFirstWorksheet == true)
                {
                    oSLActDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "Bad Costing Method");
                    bFirstWorksheet = false;
                }
                else
                {
                    oSLActDocument.AddWorksheet("Bad Costing Method");
                }
                //set column header
                oSLActDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Part Num");
                oSLActDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 30);
                oSLActDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Cost Method");
                oSLActDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 30);
                oSLActDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "The Purchased Part Has A Costing Method That Is Not Appropriate");
                oSLActDocument.SetCellStyle(iNumberOfRows++, iNumberOfColumns, oHighlightHeaderStyle);
                oSLActDocument.SetColumnWidth(iNumberOfColumns++, 90);

                foreach (HSPartData oTmpPart in oBadCostingMethod)
                {
                    oSLActDocument.SetCellValue(iNumberOfRows, 1, oTmpPart.PartNum);
                    oSLActDocument.SetCellValue(iNumberOfRows, 2, oTmpPart.PartCostingMethod);

                    iNumberOfRows++;
                    bDataInReport = true;
                }
            }

            //
            // no product group
            //
            iNumberOfRows = 1;
            iNumberOfColumns = 1;
            List<HSPartData> oMissingProductCodes = oAllActivePurchasedParts.Where(oItem => string.IsNullOrEmpty(oItem.Group) == true).ToList();
            if (oMissingProductCodes.Count > 0)
            {
                if (bFirstWorksheet == true)
                {
                    oSLActDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "No Product Group");
                    bFirstWorksheet = false;
                }
                else
                {
                    oSLActDocument.AddWorksheet("No Product Group");
                }
                //set column header
                oSLActDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Part Num");
                oSLActDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 30);
                oSLActDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Description");
                oSLActDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 60);
                oSLActDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "The Purchased Part Should Have The Product Group Set");
                oSLActDocument.SetCellStyle(iNumberOfRows++, iNumberOfColumns, oHighlightHeaderStyle);
                oSLActDocument.SetColumnWidth(iNumberOfColumns++, 90);

                foreach (HSPartData oTmpPart in oMissingProductCodes)
                {
                    oSLActDocument.SetCellValue(iNumberOfRows, 1, oTmpPart.PartNum);
                    oSLActDocument.SetCellValue(iNumberOfRows, 2, StringExt.Truncate(oTmpPart.Description, 60));

                    iNumberOfRows++;
                    bDataInReport = true;
                }
            }


            //
            // no class code
            //
            iNumberOfRows = 1;
            iNumberOfColumns = 1;
            List<HSPartData> oMissingClassCodes = oAllActivePurchasedParts.Where(oItem => string.IsNullOrEmpty(oItem.ClassID) == true).ToList();
            if (oMissingClassCodes.Count > 0)
            {
                if (bFirstWorksheet == true)
                {
                    oSLActDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "No Class Code");
                    bFirstWorksheet = false;
                }
                else
                {
                    oSLActDocument.AddWorksheet("No Class Code");
                }
                //set column header
                oSLActDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Part Num");
                oSLActDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 30);
                oSLActDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Description");
                oSLActDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 60);
                oSLActDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "The Purchased Part Must Have The Class Code Set");
                oSLActDocument.SetCellStyle(iNumberOfRows++, iNumberOfColumns, oHighlightHeaderStyle);
                oSLActDocument.SetColumnWidth(iNumberOfColumns++, 90);

                foreach (HSPartData oTmpPart in oMissingClassCodes)
                {
                    oSLActDocument.SetCellValue(iNumberOfRows, 1, oTmpPart.PartNum);
                    oSLActDocument.SetCellValue(iNumberOfRows, 2, StringExt.Truncate(oTmpPart.Description, 60));

                    iNumberOfRows++;
                    bDataInReport = true;
                }
            }


            //
            // no portfolio code
            //
            iNumberOfRows = 1;
            iNumberOfColumns = 1;
            List<HSPartData> oMissingProductPortfolioCode = oAllActivePurchasedParts.Where(oItem => string.IsNullOrEmpty(oItem.ProductPortfolio) == true).ToList();
            if (oMissingProductPortfolioCode.Count > 0)
            {
                if (bFirstWorksheet == true)
                {
                    oSLActDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "No Portfolio Code");
                    bFirstWorksheet = false;
                }
                else
                {
                    oSLActDocument.AddWorksheet("No Portfolio Code");
                }
                //set column header
                oSLActDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Part Num");
                oSLActDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 30);
                oSLActDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Description");
                oSLActDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 60);
                oSLActDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "The Purchased Part Should Have The Portfolio Code Set");
                oSLActDocument.SetCellStyle(iNumberOfRows++, iNumberOfColumns, oHighlightHeaderStyle);
                oSLActDocument.SetColumnWidth(iNumberOfColumns++, 90);

                foreach (HSPartData oTmpPart in oMissingProductPortfolioCode)
                {
                    oSLActDocument.SetCellValue(iNumberOfRows, 1, oTmpPart.PartNum);
                    oSLActDocument.SetCellValue(iNumberOfRows, 2, StringExt.Truncate(oTmpPart.Description, 60));

                    iNumberOfRows++;
                    bDataInReport = true;
                }
            }

            if (bDataInReport == true)
            {
                oSLActDocument.SaveAs(sDestinationFileName);
                // Check to see if we created a file and if so email it
                if (File.Exists(sDestinationFileName) == true)
                {
                    List<string> oAttachments = new List<string>();
                    oAttachments.Add(sDestinationFileName);

                    HSEmailHelper.SendEmail(oToAddresses, "Purchased Parts Accounting Report", "Purchased Parts Accounting Report for " + sDate, oAttachments);
                }
            }
            #endregion
        }

        public void PerformMfgPartValidation(string sCompany, string sTmpFileDirectory)
        {
            // get the file name
            DateTime dtToday = DateTime.Now;
            string sDate = dtToday.Month.ToString() + "-" + dtToday.Day.ToString() + "-" + dtToday.Year.ToString();

            // only get purchased parts that are active
            List<HSPartData> oAllActiveMfgParts = m_oAllParts.Where(oItem => (string.Compare(oItem.PartTypeCode, "M", true) == 0) && (oItem.Inactive == false)).ToList();

            #region BOM Issues

            string sDestinationFileName = sTmpFileDirectory + "\\" + sCompany + "-MfgPartBOMReport-" + sDate + ".xlsx";
            SLDocument oSLBOMDocument = new SLDocument();
            int iNumberOfRows = 1;
            int iNumberOfColumns = 1;
            bool bDataInReport = false;
            bool bFirstWorksheet = true;
            HSUser oServiceAccount = HSUser.GetUserById(HSUser.SPINCRAFT_SERVICE_ACCOUNT_ID);

            List<string> oToAddresses = new List<string>();
            oToAddresses.Add(HSUser.SPINCRAFT_ROOT_USER);

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

            //
            // no materials
            //

            //
            // inactive materials
            //

            //
            // materials on hold
            //

            //
            // materials on run out
            //

            //
            // material not assigned to operation
            //

            //
            // duplicate materials
            //

            //
            // SHOULD HAVE NON-STOCK TURNED ON!!!
            //

            //
            // SHOULD HAVE NON-STOCK TURNED OFF!!!
            //

            //
            // no operations
            //

            //
            // no revisions
            //

            //
            // no approved revs
            //

            //
            // multiple approved revs
            //

            //
            // not use part rev
            //
            iNumberOfRows = 1;
            iNumberOfColumns = 1;
            {
                List<HSPartData> oNotUsingPartRev = oAllActiveMfgParts.Where(oItem => (oItem.UsePartRev == false)).ToList();
                if (oNotUsingPartRev.Count > 0)
                {
                    if (bFirstWorksheet == true)
                    {
                        oSLBOMDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "Not Using Part Rev");
                        bFirstWorksheet = false;
                    }
                    else
                    {
                        oSLBOMDocument.AddWorksheet("Not Using Part Rev");
                    }
                    //set column header
                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "PartNum");
                    oSLBOMDocument.SetColumnWidth(iNumberOfRows++, iNumberOfColumns++, 30);

                    foreach (HSPartData oTmpPart in oNotUsingPartRev)
                    {
                        oSLBOMDocument.SetCellValue(iNumberOfRows++, 1, oTmpPart.PartNum);
                        bDataInReport = true;
                    }
                }
            }


            //
            // not qty bearing
            //
            iNumberOfRows = 1;
            iNumberOfColumns = 1;
            {
                List<HSPartData> oNotQtyBearing = oAllActiveMfgParts.Where(oItem => (oItem.PartQtyBearing == false)).ToList();
                if (oNotQtyBearing.Count > 0)
                {
                    if (bFirstWorksheet == true)
                    {
                        oSLBOMDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "Not Qty Bearing");
                        bFirstWorksheet = false;
                    }
                    else
                    {
                        oSLBOMDocument.AddWorksheet("Not Qty Bearing");
                    }
                    //set column header
                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "PartNum");
                    oSLBOMDocument.SetColumnWidth(iNumberOfRows++, iNumberOfColumns++, 30);

                    foreach (HSPartData oTmpPart in oNotQtyBearing)
                    {
                        oSLBOMDocument.SetCellValue(iNumberOfRows++, 1, oTmpPart.PartNum);
                        bDataInReport = true;
                    }
                }
            }


            //
            // incorrect UOM
            //
            iNumberOfRows = 1;
            iNumberOfColumns = 1;
            {
                List<HSPartData> oIncorrectUOM = oAllActiveMfgParts.Where(oItem => (string.Compare(oItem.IUOM, "EA", true) != 0)).ToList();
                if (oIncorrectUOM.Count > 0)
                {
                    if (bFirstWorksheet == true)
                    {
                        oSLBOMDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "IUM Incorrect");
                        bFirstWorksheet = false;
                    }
                    else
                    {
                        oSLBOMDocument.AddWorksheet("IUM Incorrect");
                    }
                    //set column header
                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "PartNum");
                    oSLBOMDocument.SetColumnWidth(iNumberOfRows++, iNumberOfColumns++, 30);

                    foreach (HSPartData oTmpPart in oIncorrectUOM)
                    {
                        oSLBOMDocument.SetCellValue(iNumberOfRows++, 1, oTmpPart.PartNum);
                        bDataInReport = true;
                    }
                }
            }


            //
            // serial or lot tracking?
            //

            //
            // inconsistent phantom BOM
            iNumberOfRows = 1;
            iNumberOfColumns = 1;
            {
                List<HSPartData> oInconsistentPhantomBOM = oAllActiveMfgParts.Where(oItem => (oItem.PartPhantomBOM != oItem.PlantPhantomBOM)).ToList();
                if (oInconsistentPhantomBOM.Count > 0)
                {
                    if (bFirstWorksheet == true)
                    {
                        oSLBOMDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "Inconsistent Phantom BOM");
                        bFirstWorksheet = false;
                    }
                    else
                    {
                        oSLBOMDocument.AddWorksheet("Inconsistent Phantom BOM");
                    }
                    //set column header
                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "PartNum");
                    oSLBOMDocument.SetColumnWidth(iNumberOfRows++, iNumberOfColumns++, 30);

                    foreach (HSPartData oTmpPart in oInconsistentPhantomBOM)
                    {
                        oSLBOMDocument.SetCellValue(iNumberOfRows++, 1, oTmpPart.PartNum);
                        bDataInReport = true;
                    }
                }
            }

            //
            // no primary warehouse
            //
            iNumberOfRows = 1;
            iNumberOfColumns = 1;
            {
                List<HSPartData> oNoPrimaryWarehouse = oAllActiveMfgParts.Where(oItem => (string.IsNullOrEmpty(oItem.PrimWhse) == true) && (oItem.PartPhantomBOM == false)).ToList();
                if (oNoPrimaryWarehouse.Count > 0)
                {
                    if (bFirstWorksheet == true)
                    {
                        oSLBOMDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "No Prim Warehouse");
                        bFirstWorksheet = false;
                    }
                    else
                    {
                        oSLBOMDocument.AddWorksheet("No Prim Warehouse");
                    }
                    //set column header
                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "PartNum");
                    oSLBOMDocument.SetColumnWidth(iNumberOfRows++, iNumberOfColumns++, 30);

                    foreach (HSPartData oTmpPart in oNoPrimaryWarehouse)
                    {
                        oSLBOMDocument.SetCellValue(iNumberOfRows++, 1, oTmpPart.PartNum);
                        bDataInReport = true;
                    }
                }
            }

            //
            // no primary bin
            //

            //
            // inconsistent qty bearing
            //
            iNumberOfRows = 1;
            iNumberOfColumns = 1;
            {
                List<HSPartData> oInconsistentQtyBearing = oAllActiveMfgParts.Where(oItem => (oItem.PartQtyBearing != oItem.PlantQtyBearing)).ToList();
                if (oInconsistentQtyBearing.Count > 0)
                {
                    if (bFirstWorksheet == true)
                    {
                        oSLBOMDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "Inconsistent Qty Bearing");
                        bFirstWorksheet = false;
                    }
                    else
                    {
                        oSLBOMDocument.AddWorksheet("Inconsistent Qty Bearing");
                    }
                    //set column header
                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "PartNum");
                    oSLBOMDocument.SetColumnWidth(iNumberOfRows++, iNumberOfColumns++, 30);

                    foreach (HSPartData oTmpPart in oInconsistentQtyBearing)
                    {
                        oSLBOMDocument.SetCellValue(iNumberOfRows++, 1, oTmpPart.PartNum);
                        bDataInReport = true;
                    }
                }
            }

            //
            // inconsistent non stock
            //
            iNumberOfRows = 1;
            iNumberOfColumns = 1;
            {
                List<HSPartData> oInconsistentNonStock = oAllActiveMfgParts.Where(oItem => (oItem.PartNonStock != oItem.PlantNonStockItem)).ToList();
                if (oInconsistentNonStock.Count > 0)
                {
                    if (bFirstWorksheet == true)
                    {
                        oSLBOMDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "Inconsistent Non-Stock");
                        bFirstWorksheet = false;
                    }
                    else
                    {
                        oSLBOMDocument.AddWorksheet("Inconsistent Non-Stock");
                    }
                    //set column header
                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "PartNum");
                    oSLBOMDocument.SetColumnWidth(iNumberOfRows++, iNumberOfColumns++, 30);

                    foreach (HSPartData oTmpPart in oInconsistentNonStock)
                    {
                        oSLBOMDocument.SetCellValue(iNumberOfRows++, 1, oTmpPart.PartNum);
                        bDataInReport = true;
                    }
                }
            }


            //
            // inconsistent part type
            //
            iNumberOfRows = 1;
            iNumberOfColumns = 1;
            {
                List<HSPartData> oInconsistentPartType = oAllActiveMfgParts.Where(oItem => (string.Compare(oItem.PartTypeCode, oItem.PlantSourceType, true) != 0)).ToList();
                if (oInconsistentPartType.Count > 0)
                {
                    if (bFirstWorksheet == true)
                    {
                        oSLBOMDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "Inconsistent Part Type");
                        bFirstWorksheet = false;
                    }
                    else
                    {
                        oSLBOMDocument.AddWorksheet("Inconsistent Part Type");
                    }
                    //set column header
                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "PartNum");
                    oSLBOMDocument.SetColumnWidth(iNumberOfRows++, iNumberOfColumns++, 30);

                    foreach (HSPartData oTmpPart in oInconsistentPartType)
                    {
                        oSLBOMDocument.SetCellValue(iNumberOfRows++, 1, oTmpPart.PartNum);
                        bDataInReport = true;
                    }
                }
            }

            //
            // mfg material qty zero
            //


            //
            // zero dollar cost
            //
            iNumberOfRows = 1;
            iNumberOfColumns = 1;
            {
                List<HSPartData> oZeroCost = oAllActiveMfgParts.Where(oItem => (oItem.Cost == 0)).ToList();
                if (oZeroCost.Count > 0)
                {
                    if (bFirstWorksheet == true)
                    {
                        oSLBOMDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "Zero Cost");
                        bFirstWorksheet = false;
                    }
                    else
                    {
                        oSLBOMDocument.AddWorksheet("Zero Cost");
                    }
                    //set column header
                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "PartNum");
                    oSLBOMDocument.SetColumnWidth(iNumberOfRows++, iNumberOfColumns++, 30);

                    foreach (HSPartData oTmpPart in oZeroCost)
                    {
                        oSLBOMDocument.SetCellValue(iNumberOfRows++, 1, oTmpPart.PartNum);
                        bDataInReport = true;
                    }
                }
            }


            //
            // MRP turned off
            //
            iNumberOfRows = 1;
            iNumberOfColumns = 1;
            {
                List<HSPartData> oMRPOff = oAllActiveMfgParts.Where(oItem => (oItem.PlantProcessMRP == false)).ToList();
                if (oMRPOff.Count > 0)
                {
                    if (bFirstWorksheet == true)
                    {
                        oSLBOMDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "MRP Off");
                        bFirstWorksheet = false;
                    }
                    else
                    {
                        oSLBOMDocument.AddWorksheet("MRP Off");
                    }
                    //set column header
                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "PartNum");
                    oSLBOMDocument.SetColumnWidth(iNumberOfRows++, iNumberOfColumns++, 30);

                    foreach (HSPartData oTmpPart in oMRPOff)
                    {
                        oSLBOMDocument.SetCellValue(iNumberOfRows++, 1, oTmpPart.PartNum);
                        bDataInReport = true;
                    }
                }
            }

            //
            // negative inventory
            //


            //
            // Buy To Order Set
            //
            iNumberOfRows = 1;
            iNumberOfColumns = 1;
            {
                List<HSPartData> oBuyToOrderSet = oAllActiveMfgParts.Where(oItem => ((oItem.PartBuyToOrder == true) || (oItem.PlantBuyToOrder == true))).ToList();
                if (oBuyToOrderSet.Count > 0)
                {
                    if (bFirstWorksheet == true)
                    {
                        oSLBOMDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "Buy-To-Order Set");
                        bFirstWorksheet = false;
                    }
                    else
                    {
                        oSLBOMDocument.AddWorksheet("Buy-To-Order Set");
                    }
                    //set column header
                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "PartNum");
                    oSLBOMDocument.SetColumnWidth(iNumberOfRows++, iNumberOfColumns++, 30);

                    foreach (HSPartData oTmpPart in oBuyToOrderSet)
                    {
                        oSLBOMDocument.SetCellValue(iNumberOfRows++, 1, oTmpPart.PartNum);
                        bDataInReport = true;
                    }
                }
            }


            //
            // Drop Ship Set
            //
            iNumberOfRows = 1;
            iNumberOfColumns = 1;
            {
                List<HSPartData> oDropShipSet = oAllActiveMfgParts.Where(oItem => ((oItem.PartDropShip == true) || (oItem.PlantDropShip == true))).ToList();
                if (oDropShipSet.Count > 0)
                {
                    if (bFirstWorksheet == true)
                    {
                        oSLBOMDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "Drop Ship Set");
                        bFirstWorksheet = false;
                    }
                    else
                    {
                        oSLBOMDocument.AddWorksheet("Drop Ship Set");
                    }
                    //set column header
                    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "PartNum");
                    oSLBOMDocument.SetColumnWidth(iNumberOfRows++, iNumberOfColumns++, 30);

                    foreach (HSPartData oTmpPart in oDropShipSet)
                    {
                        oSLBOMDocument.SetCellValue(iNumberOfRows++, 1, oTmpPart.PartNum);
                        bDataInReport = true;
                    }
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

                    HSEmailHelper.SendEmail(oToAddresses, "Mfg Parts BOM Report", "Mfg Parts BOM Report for " + sDate, oAttachments);
                }
            }
            #endregion

            #region Accounting Issues
            sDestinationFileName = sTmpFileDirectory + "\\" + sCompany + "-MfgPartAccountingReport-" + sDate + ".xlsx";
            SLDocument oSLActDocument = new SLDocument();
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
            bDataInReport = false;
            bFirstWorksheet = true;
            oToAddresses = HSUser.GetEmailsForUsersInGroup(HSUser.REPORT_ON_PURCHASED_PART_ISSUES);
            oToAddresses.Add(HSUser.SPINCRAFT_ROOT_USER);

            //
            // inconsistent cost method
            //
            iNumberOfRows = 1;
            iNumberOfColumns = 1;
            {
                List<HSPartData> oInconsistentCostMethod = oAllActiveMfgParts.Where(oItem => (string.Compare(oItem.PartCostingMethod, oItem.PlantCostingMethod, true) != 0)).ToList();
                if (oInconsistentCostMethod.Count > 0)
                {
                    if (bFirstWorksheet == true)
                    {
                        oSLActDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "Inconsistent Costing Method");
                        bFirstWorksheet = false;
                    }
                    else
                    {
                        oSLActDocument.AddWorksheet("Inconsistent Costing Method");
                    }
                    //set column header
                    oSLActDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "PartNum");
                    oSLActDocument.SetColumnWidth(iNumberOfRows++, iNumberOfColumns++, 30);

                    foreach (HSPartData oTmpPart in oInconsistentCostMethod)
                    {
                        oSLActDocument.SetCellValue(iNumberOfRows++, 1, oTmpPart.PartNum);
                        bDataInReport = true;
                    }
                }
            }

            //
            // no product group
            //
            iNumberOfRows = 1;
            iNumberOfColumns = 1;
            {
                List<HSPartData> oMissingProductCodes = oAllActiveMfgParts.Where(oItem => string.IsNullOrEmpty(oItem.Group) == true).ToList();
                if (oMissingProductCodes.Count > 0)
                {
                    if (bFirstWorksheet == true)
                    {
                        oSLActDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "No Product Group");
                        bFirstWorksheet = false;
                    }
                    else
                    {
                        oSLActDocument.AddWorksheet("No Product Group");
                    }
                    //set column header
                    oSLActDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "PartNum");
                    oSLActDocument.SetColumnWidth(iNumberOfRows++, iNumberOfColumns++, 30);

                    foreach (HSPartData oTmpPart in oMissingProductCodes)
                    {
                        oSLActDocument.SetCellValue(iNumberOfRows++, 1, oTmpPart.PartNum);
                        bDataInReport = true;
                    }
                }
            }


            //
            // no class code
            //
            iNumberOfRows = 1;
            iNumberOfColumns = 1;
            {
                List<HSPartData> oMissingClassCodes = oAllActiveMfgParts.Where(oItem => string.IsNullOrEmpty(oItem.ClassID) == true).ToList();
                if (oMissingClassCodes.Count > 0)
                {
                    if (bFirstWorksheet == true)
                    {
                        oSLActDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "No Class Code");
                        bFirstWorksheet = false;
                    }
                    else
                    {
                        oSLActDocument.AddWorksheet("No Class Code");
                    }
                    //set column header
                    oSLActDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "PartNum");
                    oSLActDocument.SetColumnWidth(iNumberOfRows++, iNumberOfColumns++, 30);

                    foreach (HSPartData oTmpPart in oMissingClassCodes)
                    {
                        oSLActDocument.SetCellValue(iNumberOfRows++, 1, oTmpPart.PartNum);
                        bDataInReport = true;
                    }
                }
            }


            //
            // no portfolio code
            //
            iNumberOfRows = 1;
            iNumberOfColumns = 1;
            {
                List<HSPartData> oMissingProductPortfolioCode = oAllActiveMfgParts.Where(oItem => string.IsNullOrEmpty(oItem.ProductPortfolio) == true).ToList();
                if (oMissingProductPortfolioCode.Count > 0)
                {
                    if (bFirstWorksheet == true)
                    {
                        oSLActDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "No Portfolio Code");
                        bFirstWorksheet = false;
                    }
                    else
                    {
                        oSLActDocument.AddWorksheet("No Portfolio Code");
                    }
                    //set column header
                    oSLActDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "PartNum");
                    oSLActDocument.SetColumnWidth(iNumberOfRows++, iNumberOfColumns++, 30);

                    foreach (HSPartData oTmpPart in oMissingProductPortfolioCode)
                    {
                        oSLActDocument.SetCellValue(iNumberOfRows++, 1, oTmpPart.PartNum);
                        bDataInReport = true;
                    }
                }
            }

            if (bDataInReport == true)
            {
                oSLActDocument.SaveAs(sDestinationFileName);
                // Check to see if we created a file and if so email it
                if (File.Exists(sDestinationFileName) == true)
                {
                    List<string> oAttachments = new List<string>();
                    oAttachments.Add(sDestinationFileName);

                    HSEmailHelper.SendEmail(oToAddresses, "Mfg Parts Accounting Report", "Mfg Parts Accounting Report for " + sDate, oAttachments);
                }
            }
            #endregion
        }

        #endregion

        #region Properties
        public List<HSPartData> AllParts
        {
            get { return m_oAllParts; }
        }
        #endregion

        #region Data Member
        private List<HSPartData> m_oAllParts = new List<HSPartData>();
        private Dictionary<string, HSPartData> m_oFastParts = new Dictionary<string, HSPartData>();
        #endregion
    }

}
