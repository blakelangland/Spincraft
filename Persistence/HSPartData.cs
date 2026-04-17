using HorizonScientific;
using Ice.BO;
using Ice.Core;
using Ice.Lib.Framework;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace HorizonScientific
{
    public class HSPartData
    {
        #region Constructors

        public HSPartData(HSPartData oOriginal)
        {
            this.m_sCompany = oOriginal.m_sCompany;
            // we always will force the part num to be in upper case for comparisons
            if (string.IsNullOrEmpty(oOriginal.m_sPartNum) == false)
            {
                this.m_sPartNum = oOriginal.m_sPartNum.ToUpper();
            }
            this.m_sDescription = oOriginal.m_sDescription;
            this.m_dtCreatedOn = oOriginal.m_dtCreatedOn;
            this.m_bInactive = oOriginal.m_bInactive;
            this.m_sPartTypeCode = oOriginal.m_sPartTypeCode;
            this.m_sClassID = oOriginal.m_sClassID;
            this.m_sGroup = oOriginal.m_sGroup;
            this.m_bGenerateSuggestions = oOriginal.m_bGenerateSuggestions;
            this.m_bPartNonStock = oOriginal.m_bPartNonStock;
            this.m_bPartQtyBearing = oOriginal.m_bPartQtyBearing;
            this.m_bPlantBackFlush = oOriginal.m_bPlantBackFlush;
            this.m_bPartDropShip = oOriginal.m_bPartDropShip;
            this.m_bPartBuyToOrder = oOriginal.m_bPartBuyToOrder;
            this.m_sBuyerID = oOriginal.m_sBuyerID;
            this.m_sVendorName = oOriginal.m_sVendorName;
            this.m_sSupplierID = oOriginal.m_sSupplierID;
            this.m_iSupplierNum = oOriginal.m_iSupplierNum;
            this.m_sPrimWhse = oOriginal.m_sPrimWhse;
            this.m_sIUOM = oOriginal.m_sIUOM;
            this.m_sSUOM = oOriginal.m_sSUOM;
            this.m_dMaximum = oOriginal.m_dMaximum;
            this.m_dMinimum = oOriginal.m_dMinimum;
            this.m_dSafety = oOriginal.m_dSafety;
            this.m_dLowestQty = oOriginal.m_dLowestQty;
            this.m_dHighestQty = oOriginal.m_dHighestQty;
            this.m_dCurrentQty = oOriginal.m_dCurrentQty;
            this.m_dMinimumCost = oOriginal.m_dMinimumCost;
            this.m_dMaximumCost = oOriginal.m_dMaximumCost;
            this.m_dCurrentCost = oOriginal.m_dCurrentCost;
            this.m_dMinOrderQty = oOriginal.m_dMinOrderQty;
        }

        public HSPartData(DataRow oDataRow)
        {
            if ((oDataRow["Part_Company"] != DBNull.Value) && (string.IsNullOrEmpty((string)oDataRow["Part_Company"]) == false))
            {
                m_sCompany = (string)oDataRow["Part_Company"];
            }
            if ((oDataRow["Part_PartNum"] != DBNull.Value) && (string.IsNullOrEmpty((string)oDataRow["Part_PartNum"]) == false))
            {
                m_sPartNum = (string)oDataRow["Part_PartNum"];
                // we always will force the part num to be in upper case for comparisons
                if (string.IsNullOrEmpty(m_sPartNum) == false)
                {
                    m_sPartNum = m_sPartNum.ToUpper();
                }
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
            if ((oDataRow["Vendor1_Name"] != DBNull.Value) && (string.IsNullOrEmpty((string)oDataRow["Vendor1_Name"]) == false))
            {
                m_sVendorName = (string)oDataRow["Vendor1_Name"];
            }
            if ((oDataRow["Vendor1_VendorID"] != DBNull.Value) && (string.IsNullOrEmpty((string)oDataRow["Vendor1_VendorID"]) == false))
            {
                m_sSupplierID = (string)oDataRow["Vendor1_VendorID"];
            }
            if (oDataRow["PartPlant_MinOrderQty"] != DBNull.Value)
            {
                m_dMinOrderQty = (Decimal)oDataRow["PartPlant_MinOrderQty"];
            }
            if (oDataRow["Vendor1_VendorNum"] != DBNull.Value)
            {
                m_iSupplierNum = (int)oDataRow["Vendor1_VendorNum"];
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
            if (oDataRow["Calculated_MaxEffectiveDate"] != DBNull.Value)
            {
                m_dtVendorPriceEffectiveDate = (DateTime)oDataRow["Calculated_MaxEffectiveDate"];
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

        #region Methods

        public static bool Initialize(Session oSession, string sPartNum)
        {
            Ice.Proxy.BO.DynamicQueryImpl oDynamicQuery = WCFServiceSupport.CreateImpl<Ice.Proxy.BO.DynamicQueryImpl>(oSession, Ice.Proxy.BO.DynamicQueryImpl.UriPath);
            QueryExecutionDataSet oQueryExecutionDataSet = null;
            DataSet oDataSet = null;

            #region Date Range
            // we will pull the data back for the cost history one month at a time
            List<Tuple<DateTime, DateTime>> oAllDateRanges = new List<Tuple<DateTime, DateTime>>();
            // first date covers everything up to the go-live date
            oAllDateRanges.Add(new Tuple<DateTime, DateTime>(new DateTime(2017, 1, 1), new DateTime(2017, 12, 31)));
            DateTime dtStartDate = new DateTime(2018, 1, 1);
            DateTime dtEndDate = DateTime.Now;
            dtEndDate = new DateTime(dtEndDate.Year, dtEndDate.Month, DateTime.DaysInMonth(dtEndDate.Year, dtEndDate.Month));
            for (DateTime dtFirstDayOfMonth = dtStartDate; dtFirstDayOfMonth <= dtEndDate;)
            {
                DateTime dtLastDayOfMonth = new DateTime(dtFirstDayOfMonth.Year, dtFirstDayOfMonth.Month, DateTime.DaysInMonth(dtFirstDayOfMonth.Year, dtFirstDayOfMonth.Month));

                oAllDateRanges.Add(new Tuple<DateTime, DateTime>(dtFirstDayOfMonth, dtLastDayOfMonth));
                dtFirstDayOfMonth = dtFirstDayOfMonth.AddMonths(1);
            }
            #endregion

            #region Get All Parts
            if (string.IsNullOrEmpty(sPartNum) == true)
            {
                // get the list of all parts

                // get all parts that have ever existed
                oQueryExecutionDataSet = oDynamicQuery.GetQueryExecutionParametersByID(BAQConstants.QUERY_LIST_ALL_PARTS);
                oQueryExecutionDataSet.ExecutionParameter.Clear();
                oDataSet = oDynamicQuery.ExecuteByID(BAQConstants.QUERY_LIST_ALL_PARTS, oQueryExecutionDataSet);
                g_oAllPartData.Clear();
                foreach (DataRow oDataRow in oDataSet.Tables[0].Rows)
                {
                    HSPartData oPartData = new HSPartData(oDataRow);
                    g_oAllPartData.Add(oPartData.PartNum, oPartData);
                }
            }
            else
            {
                // we just get one part
                oQueryExecutionDataSet = oDynamicQuery.GetQueryExecutionParametersByID(BAQConstants.QUERY_GET_PART);
                foreach (DataRow oParameter in oQueryExecutionDataSet.ExecutionParameter)
                {
                    if (string.Compare(oParameter["ParameterID"].ToString(), "PartNum") == 0)
                    {
                        oParameter["ParameterValue"] = sPartNum;
                    }
                }
                oQueryExecutionDataSet.AcceptChanges(); oDataSet = oDynamicQuery.ExecuteByID(BAQConstants.QUERY_GET_PART, oQueryExecutionDataSet);
                oDataSet = oDynamicQuery.ExecuteByID(BAQConstants.QUERY_GET_PART, oQueryExecutionDataSet);
                g_oAllPartData.Clear();
                foreach (DataRow oDataRow in oDataSet.Tables[0].Rows)
                {
                    HSPartData oPartData = new HSPartData(oDataRow);
                    g_oAllPartData.Add(oPartData.PartNum, oPartData);
                }
            }
            #endregion

            #region Get Cost History
            if (string.IsNullOrEmpty(sPartNum) == true)
            {
                // get all changes to part costs through time
                oQueryExecutionDataSet = oDynamicQuery.GetQueryExecutionParametersByID(BAQConstants.QUERY_PART_COST_HISTORY_IN_DATE_RANGE);
                foreach (Tuple<DateTime, DateTime> oDateRange in oAllDateRanges)
                {
                    foreach (DataRow oParameter in oQueryExecutionDataSet.ExecutionParameter)
                    {
                        if (string.Compare(oParameter["ParameterID"].ToString(), "StartDate") == 0)
                        {
                            oParameter["ParameterValue"] = oDateRange.Item1;
                        }
                        if (string.Compare(oParameter["ParameterID"].ToString(), "EndDate") == 0)
                        {
                            oParameter["ParameterValue"] = oDateRange.Item2;
                        }
                    }
                    oQueryExecutionDataSet.AcceptChanges();
                    oDataSet = oDynamicQuery.ExecuteByID(BAQConstants.QUERY_PART_COST_HISTORY_IN_DATE_RANGE, oQueryExecutionDataSet);
                    foreach (DataRow oDataRow in oDataSet.Tables[0].Rows)
                    {
                        HSPartCostHistory oPartCostHistory = new HSPartCostHistory(oDataRow);
                        if (g_oAllPartData.ContainsKey(oPartCostHistory.PartNum) == true)
                        {
                            HSPartData oPartData = g_oAllPartData[oPartCostHistory.PartNum];
                            oPartData.AddPartCostHistory(oPartCostHistory);
                        }
                    }
                }
            }
            else
            {
                // we only get the part history for one part
                oQueryExecutionDataSet = oDynamicQuery.GetQueryExecutionParametersByID(BAQConstants.QUERY_PART_COST_HISTORY_IN_DATE_RANGE_BY_PART2);
                foreach (Tuple<DateTime, DateTime> oDateRange in oAllDateRanges)
                {
                    foreach (DataRow oParameter in oQueryExecutionDataSet.ExecutionParameter)
                    {
                        if (string.Compare(oParameter["ParameterID"].ToString(), "StartDate") == 0)
                        {
                            oParameter["ParameterValue"] = oDateRange.Item1;
                        }
                        if (string.Compare(oParameter["ParameterID"].ToString(), "EndDate") == 0)
                        {
                            oParameter["ParameterValue"] = oDateRange.Item2;
                        }
                        if (string.Compare(oParameter["ParameterID"].ToString(), "PartNum") == 0)
                        {
                            oParameter["ParameterValue"] = sPartNum;
                        }
                    }
                    oQueryExecutionDataSet.AcceptChanges();
                    oDataSet = oDynamicQuery.ExecuteByID(BAQConstants.QUERY_PART_COST_HISTORY_IN_DATE_RANGE_BY_PART2, oQueryExecutionDataSet);
                    foreach (DataRow oDataRow in oDataSet.Tables[0].Rows)
                    {
                        HSPartCostHistory oPartCostHistory = new HSPartCostHistory(oDataRow);
                        if (g_oAllPartData.ContainsKey(oPartCostHistory.PartNum) == true)
                        {
                            HSPartData oPartData = g_oAllPartData[oPartCostHistory.PartNum];
                            oPartData.AddPartCostHistory(oPartCostHistory);
                        }
                    }
                }
            }
            #endregion

            #region Get Inventory Adjustments
            if (string.IsNullOrEmpty(sPartNum) == true)
            {
                // get all adustments made after the go-live date month by month
                oQueryExecutionDataSet = oDynamicQuery.GetQueryExecutionParametersByID(BAQConstants.QUERY_INVENTORY_ADJUSTMENTS_IN_DATE_RANGE);
                foreach (Tuple<DateTime, DateTime> oDateRange in oAllDateRanges)
                {
                    foreach (DataRow oParameter in oQueryExecutionDataSet.ExecutionParameter)
                    {
                        if (string.Compare(oParameter["ParameterID"].ToString(), "StartDate") == 0)
                        {
                            oParameter["ParameterValue"] = oDateRange.Item1;
                        }
                        if (string.Compare(oParameter["ParameterID"].ToString(), "EndDate") == 0)
                        {
                            oParameter["ParameterValue"] = oDateRange.Item2;
                        }
                    }
                    oQueryExecutionDataSet.AcceptChanges();
                    oDataSet = oDynamicQuery.ExecuteByID(BAQConstants.QUERY_INVENTORY_ADJUSTMENTS_IN_DATE_RANGE, oQueryExecutionDataSet);
                    foreach (DataRow oDataRow in oDataSet.Tables[0].Rows)
                    {
                        HSInventoryAdjustment oInventoryAdjustment = new HSInventoryAdjustment(oDataRow);
                        if (g_oAllPartData.ContainsKey(oInventoryAdjustment.PartNum) == true)
                        {
                            HSPartData oPartData = g_oAllPartData[oInventoryAdjustment.PartNum];
                            oPartData.AddInventoryAdjustment(oInventoryAdjustment);
                        }
                    }
                }
            }
            else
            {
                // we only get the inventory adjustments for one part
                // get all adustments made after the go-live date month by month
                oQueryExecutionDataSet = oDynamicQuery.GetQueryExecutionParametersByID(BAQConstants.QUERY_INVENTORY_ADJUSTMENTS_IN_DATE_RANGE_BY_PART);
                foreach (Tuple<DateTime, DateTime> oDateRange in oAllDateRanges)
                {
                    foreach (DataRow oParameter in oQueryExecutionDataSet.ExecutionParameter)
                    {
                        if (string.Compare(oParameter["ParameterID"].ToString(), "StartDate") == 0)
                        {
                            oParameter["ParameterValue"] = oDateRange.Item1;
                        }
                        if (string.Compare(oParameter["ParameterID"].ToString(), "EndDate") == 0)
                        {
                            oParameter["ParameterValue"] = oDateRange.Item2;
                        }
                        if (string.Compare(oParameter["ParameterID"].ToString(), "PartNum") == 0)
                        {
                            oParameter["ParameterValue"] = sPartNum;
                        }
                    }
                    oQueryExecutionDataSet.AcceptChanges();
                    oDataSet = oDynamicQuery.ExecuteByID(BAQConstants.QUERY_INVENTORY_ADJUSTMENTS_IN_DATE_RANGE_BY_PART, oQueryExecutionDataSet);
                    foreach (DataRow oDataRow in oDataSet.Tables[0].Rows)
                    {
                        HSInventoryAdjustment oInventoryAdjustment = new HSInventoryAdjustment(oDataRow);
                        if (g_oAllPartData.ContainsKey(oInventoryAdjustment.PartNum) == true)
                        {
                            HSPartData oPartData = g_oAllPartData[oInventoryAdjustment.PartNum];
                            oPartData.AddInventoryAdjustment(oInventoryAdjustment);
                        }
                    }
                }
            }
            #endregion

            #region Compute Quantities And Value
            // order the data and fix running totals and valuation
            foreach (HSPartData oPartData in g_oAllPartData.Values)
            {
                oPartData.ComputeRunningTotalsAndValue();
            }

            #endregion

            #region PO Detail History
            //PODetailHistory
            if (string.IsNullOrEmpty(sPartNum) == true)
            {
                // get all changes to po detail history through time
                oQueryExecutionDataSet = oDynamicQuery.GetQueryExecutionParametersByID(BAQConstants.QUERY_PO_DETAIL_HISTORY);
                foreach (Tuple<DateTime, DateTime> oDateRange in oAllDateRanges)
                {
                    foreach (DataRow oParameter in oQueryExecutionDataSet.ExecutionParameter)
                    {
                        if (string.Compare(oParameter["ParameterID"].ToString(), "StartDate") == 0)
                        {
                            oParameter["ParameterValue"] = oDateRange.Item1;
                        }
                        if (string.Compare(oParameter["ParameterID"].ToString(), "EndDate") == 0)
                        {
                            oParameter["ParameterValue"] = oDateRange.Item2;
                        }
                    }
                    oQueryExecutionDataSet.AcceptChanges();
                    oDataSet = oDynamicQuery.ExecuteByID(BAQConstants.QUERY_PO_DETAIL_HISTORY, oQueryExecutionDataSet);
                    foreach (DataRow oDataRow in oDataSet.Tables[0].Rows)
                    {
                        PODetailHistory oPODetailHistory = new PODetailHistory(oDataRow);
                        if (g_oAllPartData.ContainsKey(oPODetailHistory.PartNum) == true)
                        {
                            HSPartData oPartData = g_oAllPartData[oPODetailHistory.PartNum];
                            oPartData.AddPODetailHistory(oPODetailHistory);
                        }
                    }
                }
            }
            else
            {
                // we only get the part history for one part
                oQueryExecutionDataSet = oDynamicQuery.GetQueryExecutionParametersByID(BAQConstants.QUERY_PO_DETAIL_HISTORY_BY_PART);
                foreach (Tuple<DateTime, DateTime> oDateRange in oAllDateRanges)
                {
                    foreach (DataRow oParameter in oQueryExecutionDataSet.ExecutionParameter)
                    {
                        if (string.Compare(oParameter["ParameterID"].ToString(), "StartDate") == 0)
                        {
                            oParameter["ParameterValue"] = oDateRange.Item1;
                        }
                        if (string.Compare(oParameter["ParameterID"].ToString(), "EndDate") == 0)
                        {
                            oParameter["ParameterValue"] = oDateRange.Item2;
                        }
                        if (string.Compare(oParameter["ParameterID"].ToString(), "PartNum") == 0)
                        {
                            oParameter["ParameterValue"] = sPartNum;
                        }
                    }
                    oQueryExecutionDataSet.AcceptChanges();
                    oDataSet = oDynamicQuery.ExecuteByID(BAQConstants.QUERY_PO_DETAIL_HISTORY_BY_PART, oQueryExecutionDataSet);
                    foreach (DataRow oDataRow in oDataSet.Tables[0].Rows)
                    {
                        PODetailHistory oPODetailHistory = new PODetailHistory(oDataRow);
                        if (g_oAllPartData.ContainsKey(oPODetailHistory.PartNum) == true)
                        {
                            HSPartData oPartData = g_oAllPartData[oPODetailHistory.PartNum];
                            oPartData.AddPODetailHistory(oPODetailHistory);
                        }
                    }
                }
            }
            #endregion

            #region Get Purchased Parts Consumed On Job Or Sold Directly To Customer
            //PurPartsConsumed
            if (string.IsNullOrEmpty(sPartNum) == true)
            {
                // get all adustments made after the go-live date month by month
                oQueryExecutionDataSet = oDynamicQuery.GetQueryExecutionParametersByID(BAQConstants.QUERY_PURCHASE_PARTS_CONSUMED);
                foreach (Tuple<DateTime, DateTime> oDateRange in oAllDateRanges)
                {
                    foreach (DataRow oParameter in oQueryExecutionDataSet.ExecutionParameter)
                    {
                        if (string.Compare(oParameter["ParameterID"].ToString(), "StartDate") == 0)
                        {
                            oParameter["ParameterValue"] = oDateRange.Item1;
                        }
                        if (string.Compare(oParameter["ParameterID"].ToString(), "EndDate") == 0)
                        {
                            oParameter["ParameterValue"] = oDateRange.Item2;
                        }
                    }
                    oQueryExecutionDataSet.AcceptChanges();
                    oDataSet = oDynamicQuery.ExecuteByID(BAQConstants.QUERY_PURCHASE_PARTS_CONSUMED, oQueryExecutionDataSet);
                    foreach (DataRow oDataRow in oDataSet.Tables[0].Rows)
                    {
                        PurchasedPartsConsumed oPurchasedPartConsumed = new PurchasedPartsConsumed(oDataRow);
                        if (g_oAllPartData.ContainsKey(oPurchasedPartConsumed.PartNum) == true)
                        {
                            HSPartData oPartData = g_oAllPartData[oPurchasedPartConsumed.PartNum];
                            oPartData.AddPurchasedPartConsumed(oPurchasedPartConsumed);
                        }
                    }
                }
            }
            else
            {
                // we only get the inventory adjustments for one part
                // get all adustments made after the go-live date month by month
                oQueryExecutionDataSet = oDynamicQuery.GetQueryExecutionParametersByID(BAQConstants.QUERY_PURCHASED_PARTS_CONSUMED_BY_PART);
                foreach (Tuple<DateTime, DateTime> oDateRange in oAllDateRanges)
                {
                    foreach (DataRow oParameter in oQueryExecutionDataSet.ExecutionParameter)
                    {
                        if (string.Compare(oParameter["ParameterID"].ToString(), "StartDate") == 0)
                        {
                            oParameter["ParameterValue"] = oDateRange.Item1;
                        }
                        if (string.Compare(oParameter["ParameterID"].ToString(), "EndDate") == 0)
                        {
                            oParameter["ParameterValue"] = oDateRange.Item2;
                        }
                        if (string.Compare(oParameter["ParameterID"].ToString(), "PartNum") == 0)
                        {
                            oParameter["ParameterValue"] = sPartNum;
                        }
                    }
                    oQueryExecutionDataSet.AcceptChanges();
                    oDataSet = oDynamicQuery.ExecuteByID(BAQConstants.QUERY_PURCHASED_PARTS_CONSUMED_BY_PART, oQueryExecutionDataSet);
                    foreach (DataRow oDataRow in oDataSet.Tables[0].Rows)
                    {
                        PurchasedPartsConsumed oPurchasedPartConsumed = new PurchasedPartsConsumed(oDataRow);
                        if (g_oAllPartData.ContainsKey(oPurchasedPartConsumed.PartNum) == true)
                        {
                            HSPartData oPartData = g_oAllPartData[oPurchasedPartConsumed.PartNum];
                            oPartData.AddPurchasedPartConsumed(oPurchasedPartConsumed);
                        }
                    }
                }
            }
            #endregion
            return true;
        }

        public void AddPartCostHistory(HSPartCostHistory oPartCostHistory)
        {
            m_oPartCostHistory.Add(oPartCostHistory);
        }

        public void AddInventoryAdjustment(HSInventoryAdjustment oInventoryAdjustment)
        {
            m_oPartInventoryAdjustments.Add(oInventoryAdjustment);
        }

        public void AddPurchasedPartConsumed(PurchasedPartsConsumed oPurchasedPartConsumed)
        {
            m_oAllPurchasedPartsConsumed.Add(oPurchasedPartConsumed);
        }

        public void GetStdCostForDate(DateTime dtValuationDate, out decimal dStdCost)
        {
            dStdCost = 0;
            foreach (HSPartCostHistory oPartCostHistory in m_oPartCostHistory)
            {
                if (oPartCostHistory.TranDate <= dtValuationDate)
                {
                    dStdCost = oPartCostHistory.StdCost;
                }
                else
                {
                    break;
                }
            }
            return;
        }

        public void GetQuantityOnHandAndValuationForDate(DateTime dtValuationDate, out decimal dValuation, out int iOnHandQty)
        {
            dValuation = 0;
            iOnHandQty = 0;
            foreach (HSInventoryAdjustment oAdjustment in m_oPartInventoryAdjustments)
            {
                if (oAdjustment.TranDate <= dtValuationDate)
                {
                    iOnHandQty = (int)oAdjustment.RunningTotal;
                    dValuation = oAdjustment.TotalValue;
                }
                else
                {
                    break;
                }
            }
            return;
        }

        public void ComputeRunningTotalsAndValue()
        {
            // get all data in correct datetime order
            m_oPartCostHistory = m_oPartCostHistory.OrderBy(oItem => oItem.TranDate).ToList();
            m_oPartInventoryAdjustments = m_oPartInventoryAdjustments.OrderBy(oItem => oItem.TranDate).ToList();
            m_oAllPurchasedPartsConsumed = m_oAllPurchasedPartsConsumed.OrderBy(oItem => oItem.TranDate).ToList();
            // all of the inventory adjustment dates become an event
            m_oAllInventoryAdjustmentEvents = m_oPartInventoryAdjustments.Select(oItem => oItem.TranDate).ToList();
            m_oAllPartCostAdjustmentEvents = m_oPartCostHistory.Select(oItem => oItem.TranDate).ToList();
            m_dHighestQty = 0;
            m_dLowestQty = 0;

            decimal dRunningTotal = 0;
            decimal dLastStdCost = 0;
            foreach (HSInventoryAdjustment oAdjustment in m_oPartInventoryAdjustments)
            {
                foreach (HSPartCostHistory oPartCost in m_oPartCostHistory)
                {
                    if (oPartCost.TranDate <= oAdjustment.TranDate)
                    {
                        dLastStdCost = oPartCost.StdCost;
                    }
                    else
                    {
                        break;
                    }
                }

                // fix the adjustment to reflect the cost at the time of the transaction
                oAdjustment.StdCost = dLastStdCost;
                // postive number for added to stock mean the unit was put into stock
                dRunningTotal += oAdjustment.AddedToStock;
                // negative numbers for taken from stock mean the unit was taken out of stock
                dRunningTotal += oAdjustment.TakenFromStock;
                oAdjustment.RunningTotal = dRunningTotal;

                if (dRunningTotal < m_dLowestQty)
                {
                    m_dLowestQty = dRunningTotal;
                }
                if (dRunningTotal > m_dHighestQty)
                {
                    m_dHighestQty = dRunningTotal;
                }
                if (dLastStdCost < m_dMinimumCost)
                {
                    m_dMinimumCost = dLastStdCost;
                }
                if (dLastStdCost > m_dMaximumCost)
                {
                    m_dMaximumCost = dLastStdCost;
                }
            }
            m_dCurrentQty = dRunningTotal;

            m_dMinimumCost = 0;
            m_dMaximumCost = 0;
            m_dCurrentCost = 0;
            if (m_oPartCostHistory.Count > 0)
            {
                m_dMinimumCost = m_oPartCostHistory.Min(oItem => oItem.StdCost);
                m_dMaximumCost = m_oPartCostHistory.Max(oItem => oItem.StdCost);
                m_dCurrentCost = m_oPartCostHistory.Last().StdCost;
            }
        }

        public void AddPODetailHistory(PODetailHistory oPODetailHistory)
        {
            m_oPODetailHistory.Add(oPODetailHistory);
        }

        public void ComputeRawMaterialsInventoryTurnover(DateTime dtStartDate, DateTime dtEndDate, out decimal dRawMaterialsInventoryTurnover)
        {
            dRawMaterialsInventoryTurnover = 0M;
            if (dtStartDate <= dtEndDate)
            {
                decimal dValuationOfPartsConsumed = 0M;
                foreach (PurchasedPartsConsumed oPurchasedPartsConsumed in m_oAllPurchasedPartsConsumed)
                {
                    if (oPurchasedPartsConsumed.TranDate >= dtStartDate)
                    {
                        if (oPurchasedPartsConsumed.TranDate <= dtEndDate)
                        {
                            decimal dStdCostForDate = 0M;
                            GetStdCostForDate(oPurchasedPartsConsumed.TranDate, out dStdCostForDate);
                            dValuationOfPartsConsumed += dStdCostForDate * oPurchasedPartsConsumed.TakenFromStock * -1;
                        }
                    }
                    if (oPurchasedPartsConsumed.TranDate > dtEndDate)
                    {
                        break;
                    }
                }

                // get the average raw materials inventory balance over this period of time
                decimal dInventoryValuation = 0M;
                int iNumberOfDaysInPeriod = 0;
                for (DateTime dtCurrentDate = dtStartDate; dtCurrentDate <= dtEndDate;)
                {
                    decimal dCurrentValuation = 0M;
                    int iCurrentOnHandQty = 0;
                    GetQuantityOnHandAndValuationForDate(dtCurrentDate, out dCurrentValuation, out iCurrentOnHandQty);
                    iNumberOfDaysInPeriod++;
                    dInventoryValuation += dCurrentValuation;
                    dtCurrentDate = dtCurrentDate.AddDays(1);
                }
                decimal dAverageInventoryBalance = dInventoryValuation / (decimal)iNumberOfDaysInPeriod;

                // now we can compute the raw materials inventory turnover
                // INVENTORY_TURNOVER = MTL_CONSUMED / INV_BALANCE
                if (dAverageInventoryBalance != 0)
                {
                    dRawMaterialsInventoryTurnover = dValuationOfPartsConsumed / dAverageInventoryBalance;
                }
            }
        }

        #endregion

        #region Properties

        public String Company
        {
            get { return m_sCompany; }
            set { m_sCompany = value; }
        }

        public String PartNum
        {
            get { return m_sPartNum; }
            set { m_sPartNum = value; }
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

        public int CommodityCodeLength
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

        public decimal LowestQuantity
        {
            get { return m_dLowestQty; }
            set { m_dLowestQty = value; }
        }

        public decimal HighestQuantity
        {
            get { return m_dHighestQty; }
            set { m_dHighestQty = value; }
        }

        public decimal CurrentQuantity
        {
            get { return m_dCurrentQty; }
            set { m_dCurrentQty = value; }
        }

        #endregion

        #region Data Members
        private string m_sCompany = "";
        private string m_sPartNum = "";
        private string m_sDescription = "";
        private bool m_bInactive;
        private string m_sClassID = "";
        private string m_sGroup = "";
        private string m_sSearch = "";
        private string m_sUOMClassID = "";
        private string m_sIUOM = "";
        private string m_sPUOM = "";
        private string m_sSUOM = "";
        private decimal m_dUnitNetWeight;
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
        private decimal m_dMaximum;
        private decimal m_dMinimum;
        private decimal m_dSafety;
        private decimal m_dCostingLotSize;
        private decimal m_dMinMfgLotSize;
        private decimal m_dMfgLotSize;
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
        private decimal m_dUnitPrice;

        private string m_sPurchaseComments = "";
        private string m_sBuyerID = "";
        private string m_sVendorName = "";
        private string m_sSupplierID = "";
        private decimal m_dMinOrderQty;
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

        private decimal m_dLowestQty;
        private decimal m_dHighestQty;
        private decimal m_dCurrentQty;
        private decimal m_dMinimumCost;
        private decimal m_dMaximumCost;
        private decimal m_dCurrentCost;

        private static Dictionary<string, HSPartData> g_oAllPartData = new Dictionary<string, HSPartData>();
        private List<HSPartCostHistory> m_oPartCostHistory = new List<HSPartCostHistory>();
        private List<HSInventoryAdjustment> m_oPartInventoryAdjustments = new List<HSInventoryAdjustment>();
        private List<PurchasedPartsConsumed> m_oAllPurchasedPartsConsumed = new List<PurchasedPartsConsumed>();
        private List<DateTime> m_oAllInventoryAdjustmentEvents = new List<DateTime>();
        private List<DateTime> m_oAllPartCostAdjustmentEvents = new List<DateTime>();
        private List<PODetailHistory> m_oPODetailHistory = new List<PODetailHistory>();
        #endregion
    }

}
