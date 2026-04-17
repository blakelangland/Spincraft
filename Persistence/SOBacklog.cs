using System;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using System.Runtime.InteropServices.WindowsRuntime;
using System.IO;

using Ice.Core;
using Erp.BO;
using Erp.Proxy.BO;
using Ice.Lib.Framework;
using Ice.Adapters;
using Ice.BO;
using Ice.Lib.PerformanceCanvasXmla;
using Ice.Lib.Searches;

using SpreadsheetLight;
using SpreadsheetLight.Charts;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Drawing.Spreadsheet;
using DocumentFormat.OpenXml.Drawing.Charts;


namespace HorizonScientific
{
	public class SOBacklog
	{
		#region Constructors
		public SOBacklog()
		{
		}
		public SOBacklog(DataRow oDataRow)
		{
			if (oDataRow["OrderRel_Company"] != DBNull.Value)
			{
				m_sCompany = (string)oDataRow["OrderRel_Company"];
			}
			if (oDataRow["OrderRel_OrderNum"] != DBNull.Value)
			{
				m_iOrderNum = (int)oDataRow["OrderRel_OrderNum"];
			}
			if (oDataRow["OrderRel_OrderLine"] != DBNull.Value)
			{
				m_iOrderLine = (int)oDataRow["OrderRel_OrderLine"];
			}
			if (oDataRow["OrderRel_OrderRelNum"] != DBNull.Value)
			{
				m_iOrderRelNum = (int)oDataRow["OrderRel_OrderRelNum"];
			}
			if (oDataRow["OrderRel_ReqDate"] != DBNull.Value)
			{
				m_dtRequiredBy = (DateTime)oDataRow["OrderRel_ReqDate"];
			}
			if (oDataRow["OrderRel_OurReqQty"] != DBNull.Value)
			{
				m_dOurRequestedQty = (decimal)oDataRow["OrderRel_OurReqQty"];
			}
			if (oDataRow["OrderRel_OpenRelease"] != DBNull.Value)
			{
				m_bOpenRelease = (bool)oDataRow["OrderRel_OpenRelease"];
			}
            if (oDataRow["OrderRel_VoidRelease"] != DBNull.Value)
            {
                m_bVoidRelease = (bool)oDataRow["OrderRel_VoidRelease"];
            }
            if (oDataRow["OrderRel_FirmRelease"] != DBNull.Value)
			{
				m_bFirmRelease = (bool)oDataRow["OrderRel_FirmRelease"];
			}
			if (oDataRow["OrderRel_Make"] != DBNull.Value)
			{
				m_bMakeDirect = (bool)oDataRow["OrderRel_Make"];
			}
			if (oDataRow["OrderRel_OurJobShippedQty"] != DBNull.Value)
			{
				m_dOurJobShippedQty = (decimal)oDataRow["OrderRel_OurJobShippedQty"];
			}
			if (oDataRow["OrderRel_OurStockShippedQty"] != DBNull.Value)
			{
				m_dOurStockShippedQty = (decimal)oDataRow["OrderRel_OurStockShippedQty"];
			}
			if ((oDataRow["OrderRel_Plant"] != DBNull.Value) && (string.IsNullOrEmpty((string)oDataRow["OrderRel_Plant"]) == false))
			{
				m_sPlant = (string)oDataRow["OrderRel_Plant"];
			}
			if ((oDataRow["OrderRel_WarehouseCode"] != DBNull.Value) && (string.IsNullOrEmpty((string)oDataRow["OrderRel_WarehouseCode"]) == false))
			{
				m_sWarehouse = (string)oDataRow["OrderRel_WarehouseCode"];
			}
			if ((oDataRow["OrderRel_PartNum"] != DBNull.Value) && (string.IsNullOrEmpty((string)oDataRow["OrderRel_PartNum"]) == false))
			{
				m_sPartNum = (string)oDataRow["OrderRel_PartNum"];
			}
			if (oDataRow["OrderRel_RevisionNum"] != DBNull.Value)
			{
				m_sPartRev = (string)oDataRow["OrderRel_RevisionNum"];
			}
			if (oDataRow["OrderRel_NeedByDate"] != DBNull.Value)
			{
				m_dtNeedBy = (DateTime)oDataRow["OrderRel_NeedByDate"];
			}
			if (oDataRow["OrderRel_RelStatus"] != DBNull.Value)
			{
				m_sReleaseStatus = (string)oDataRow["OrderRel_RelStatus"];
			}
			if (oDataRow["OrderRel_DropShip"] != DBNull.Value)
			{
				m_bDropShip = (bool)oDataRow["OrderRel_DropShip"];
			}
			if (oDataRow["OrderRel_BuyToOrder"] != DBNull.Value)
			{
				m_bBuyToOrder = (bool)oDataRow["OrderRel_BuyToOrder"];
			}
            if (oDataRow["OrderRel_POLine"] != DBNull.Value)
			{
				m_iPONum = (int) oDataRow["OrderRel_POLine"];
            }
            if (oDataRow["OrderRel_POLine"] != DBNull.Value)
            {
                m_iPOLine = (int)oDataRow["OrderRel_POLine"];
            }
            if (oDataRow["OrderRel_PORelNum"] != DBNull.Value)
            {
                m_iPORelNum = (int)oDataRow["OrderRel_PORelNum"];
            }
            if (oDataRow["OrderRel_ChangedBy"] != DBNull.Value)
            {
                m_sChangedBy = (string)oDataRow["OrderRel_ChangedBy"];
            }
            if (oDataRow["OrderRel_ChangeDate"] != DBNull.Value)
            {
                m_dtChangeDate = (DateTime)oDataRow["OrderRel_ChangeDate"];
            }
            if (oDataRow["OrderHed_OrderHeld"] != DBNull.Value)
			{
				m_bHoldOrder = (bool)oDataRow["OrderHed_OrderHeld"];
			}
            if (oDataRow["OrderHed_VoidOrder"] != DBNull.Value)
            {
                m_bVoidOrder = (bool)oDataRow["OrderHed_VoidOrder"];
            }
            if ((oDataRow["OrderHed_EntryPerson"] != DBNull.Value) && (string.IsNullOrEmpty((string)oDataRow["OrderHed_EntryPerson"]) == false))
			{
				m_sEntryPerson = (string)oDataRow["OrderHed_EntryPerson"];
			}
			if (oDataRow["OrderHed_OrderDate"] != DBNull.Value)
			{
				m_dtOrderDate = (DateTime)oDataRow["OrderHed_OrderDate"];
			}
			if (oDataRow["OrderHed_TermsCode"] != DBNull.Value)
			{
				m_sTermsCode = (string)oDataRow["OrderHed_TermsCode"];
			}
			if (oDataRow["OrderHed_DiscountPercent"] != DBNull.Value)
			{
				m_dDiscountPercent = (decimal)oDataRow["OrderHed_DiscountPercent"];
			}
			if ((oDataRow["OrderHed_SalesRepList"] != DBNull.Value) && (string.IsNullOrEmpty((string)oDataRow["OrderHed_SalesRepList"]) == false))
			{
				m_sSalesReps = (string)oDataRow["OrderHed_SalesRepList"];
			}
			if ((oDataRow["OrderHed_OrderComment"] != DBNull.Value) && (string.IsNullOrEmpty((string)oDataRow["OrderHed_OrderComment"]) == false))
			{
				m_sOrderComment = (string)oDataRow["OrderHed_OrderComment"];
			}
			if ((oDataRow["OrderHed_ShipComment"] != DBNull.Value) && (string.IsNullOrEmpty((string)oDataRow["OrderHed_ShipComment"]) == false))
			{
				m_sShipComment = (string)oDataRow["OrderHed_ShipComment"];
			}
			if ((oDataRow["OrderHed_InvoiceComment"] != DBNull.Value) && (string.IsNullOrEmpty((string)oDataRow["OrderHed_InvoiceComment"]) == false))
			{
				m_sInvoiceComment = (string)oDataRow["OrderHed_InvoiceComment"];
			}
			if ((oDataRow["OrderHed_PickListComment"] != DBNull.Value) && (string.IsNullOrEmpty((string)oDataRow["OrderHed_PickListComment"]) == false))
			{
				m_sPickListComment = (string)oDataRow["OrderHed_PickListComment"];
			}
			if (oDataRow["OrderHed_ExchangeRate"] != DBNull.Value)
			{
				m_dExchangeRate = (decimal)oDataRow["OrderHed_ExchangeRate"];
			}
			if ((oDataRow["OrderHed_CurrencyCode"] != DBNull.Value) && (string.IsNullOrEmpty((string)oDataRow["OrderHed_CurrencyCode"]) == false))
			{
				m_sCurrencyCode = (string)oDataRow["OrderHed_CurrencyCode"];
			}
			if (oDataRow["OrderHed_LockRate"] != DBNull.Value)
			{
				m_bLockRate = (bool)oDataRow["OrderHed_LockRate"];
			}
			if (oDataRow["OrderHed_RepSplit1"] != DBNull.Value)
			{
				m_iRepSplit1 = (int)oDataRow["OrderHed_RepSplit1"];
			}
			if (oDataRow["OrderHed_RepSplit2"] != DBNull.Value)
			{
				m_iRepSplit2 = (int)oDataRow["OrderHed_RepSplit2"];
			}
			if (oDataRow["OrderHed_ReadyToCalc"] != DBNull.Value)
			{
				m_bReadyToCalc = (bool)oDataRow["OrderHed_ReadyToCalc"];
			}
			if (oDataRow["OrderHed_OrderStatus"] != DBNull.Value)
			{
				m_sOrderStatus = (string)oDataRow["OrderHed_OrderStatus"];
			}
			if (oDataRow["OrderHed_ReadyToFulfill"] != DBNull.Value)
			{
				m_bReadyToFulfill = (bool)oDataRow["OrderHed_ReadyToFulfill"];
			}
			if ((oDataRow["UDCodes1_CodeID"] != DBNull.Value) && (string.IsNullOrEmpty((string)oDataRow["UDCodes1_CodeID"]) == false))
			{
				m_sMarketSegmentID = (string)oDataRow["UDCodes1_CodeID"];
			}
			if ((oDataRow["UDCodes1_CodeDesc"] != DBNull.Value) && (string.IsNullOrEmpty((string)oDataRow["UDCodes1_CodeDesc"]) == false))
			{
				m_sMarketSegmentDescription = (string)oDataRow["UDCodes1_CodeDesc"];
			}
            if (oDataRow["OrderDtl_VoidLine"] != DBNull.Value)
            {
                m_bVoidLine = (bool)oDataRow["OrderDtl_VoidLine"];
            }
            if ((oDataRow["OrderDtl_LineType"] != DBNull.Value) && (string.IsNullOrEmpty((string)oDataRow["OrderDtl_LineType"]) == false))
			{
				m_sLineType = (string)oDataRow["OrderDtl_LineType"];
			}
			if ((oDataRow["OrderDtl_LineDesc"] != DBNull.Value) && (string.IsNullOrEmpty((string)oDataRow["OrderDtl_LineDesc"]) == false))
			{
				m_sPartDescription = (string)oDataRow["OrderDtl_LineDesc"];
			}
			if (oDataRow["OrderDtl_UnitPrice"] != DBNull.Value)
			{
				m_dUnitPrice = (decimal)oDataRow["OrderDtl_UnitPrice"];
			}
			if (oDataRow["OrderDtl_DocUnitPrice"] != DBNull.Value)
			{
				m_dDocUnitPrice = (decimal)oDataRow["OrderDtl_DocUnitPrice"];
			}
			if (oDataRow["OrderDtl_QuoteNum"] != DBNull.Value)
			{
				m_iQuoteNum = (int)oDataRow["OrderDtl_QuoteNum"];
			}
			if (oDataRow["OrderDtl_QuoteLine"] != DBNull.Value)
			{
				m_iQuoteLine = (int)oDataRow["OrderDtl_QuoteLine"];
			}
			if ((oDataRow["OrderDtl_LineStatus"] != DBNull.Value) && (string.IsNullOrEmpty((string)oDataRow["OrderDtl_LineStatus"]) == false))
			{
				m_sLineStatus = (string)oDataRow["OrderDtl_LineStatus"];
			}
			if (oDataRow["OrderDtl_MfgJobType"] != DBNull.Value)
			{
				m_sMfgJobType = (string)oDataRow["OrderDtl_MfgJobType"];
			}
			if ((oDataRow["UDCodes_CodeID"] != DBNull.Value) && (string.IsNullOrEmpty((string)oDataRow["UDCodes_CodeID"]) == false))
			{
				m_sPortfolioID = (string)oDataRow["UDCodes_CodeID"];
			}
			if ((oDataRow["UDCodes_CodeDesc"] != DBNull.Value) && (string.IsNullOrEmpty((string)oDataRow["UDCodes_CodeDesc"]) == false))
			{
				m_sPortfolioDescription = (string)oDataRow["UDCodes_CodeDesc"];
			}
			if ((oDataRow["Part_ClassID"] != DBNull.Value) && (string.IsNullOrEmpty((string)oDataRow["Part_ClassID"]) == false))
			{
				m_sClassID = (string)oDataRow["Part_ClassID"];
			}
			if ((oDataRow["PartClass_Description"] != DBNull.Value) && (string.IsNullOrEmpty((string)oDataRow["PartClass_Description"]) == false))
			{
				m_sClassDescription = (string)oDataRow["PartClass_Description"];
			}
			if ((oDataRow["ProdGrup_ProdCode"] != DBNull.Value) && (string.IsNullOrEmpty((string)oDataRow["ProdGrup_ProdCode"]) == false))
			{
				m_sGroupID = (string)oDataRow["ProdGrup_ProdCode"];
			}
			if ((oDataRow["ProdGrup_Description"] != DBNull.Value) && (string.IsNullOrEmpty((string)oDataRow["ProdGrup_Description"]) == false))
			{
				m_sGroupDescription = (string)oDataRow["ProdGrup_Description"];
			}
			if ((oDataRow["Part_TypeCode"] != DBNull.Value) && (string.IsNullOrEmpty((string)oDataRow["Part_TypeCode"]) == false))
			{
				m_sPartType = (string)oDataRow["Part_TypeCode"];
			}
			if (oDataRow["Part_NonStock"] != DBNull.Value)
			{
				m_bNonStockItem = (bool)oDataRow["Part_NonStock"];
			}
			if ((oDataRow["Part_IUM"] != DBNull.Value) && (string.IsNullOrEmpty((string)oDataRow["Part_IUM"]) == false))
			{
				m_sIUOM = (string)oDataRow["Part_IUM"];
			}
			if ((oDataRow["Part_PUM"] != DBNull.Value) && (string.IsNullOrEmpty((string)oDataRow["Part_PUM"]) == false))
			{
				m_sPUOM = (string)oDataRow["Part_PUM"];
			}
			if (oDataRow["Part_UnitPrice"] != DBNull.Value)
			{
				m_dPartUnitPrice = (decimal)oDataRow["Part_UnitPrice"];
			}
			if ((oDataRow["Part_PricePerCode"] != DBNull.Value) && (string.IsNullOrEmpty((string)oDataRow["Part_PricePerCode"]) == false))
			{
				m_sPricePer = (string)oDataRow["Part_PricePerCode"];
			}
			if ((oDataRow["Part_CostMethod"] != DBNull.Value) && (string.IsNullOrEmpty((string)oDataRow["Part_CostMethod"]) == false))
			{
				m_sCostingMethod = (string)oDataRow["Part_CostMethod"];
			}
			if (oDataRow["Part_InActive"] != DBNull.Value)
			{
				m_bInactive = (bool)oDataRow["Part_InActive"];
			}
			if (oDataRow["Part_TrackLots"] != DBNull.Value)
			{
				m_bTrackLots = (bool)oDataRow["Part_TrackLots"];
			}
			if (oDataRow["Part_TrackSerialNum"] != DBNull.Value)
			{
				m_bTrackSerial = (bool)oDataRow["Part_TrackSerialNum"];
			}
			if ((oDataRow["Part_SalesUM"] != DBNull.Value) && (string.IsNullOrEmpty((string)oDataRow["Part_SalesUM"]) == false))
			{
				m_sSUOM = (string)oDataRow["Part_SalesUM"];
			}
			if (oDataRow["Part_SellingFactor"] != DBNull.Value)
			{
				m_dSellingFactor = (decimal)oDataRow["Part_SellingFactor"];
			}
			if (oDataRow["Part_UsePartRev"] != DBNull.Value)
			{
				m_bUsePartRev = (bool)oDataRow["Part_UsePartRev"];
			}
			if (oDataRow["Part_OnHold"] != DBNull.Value)
			{
				m_bPartOnHold = (bool)oDataRow["Part_OnHold"];
			}
			if (oDataRow["Part_QtyBearing"] != DBNull.Value)
			{
				m_bQtyBearing = (bool)oDataRow["Part_QtyBearing"];
			}
			if (oDataRow["Calculated_ExtPrice"] != DBNull.Value)
			{
				m_dExtPrice = (decimal)oDataRow["Calculated_ExtPrice"];
			}
			if (oDataRow["Calculated_RemainingQty"] != DBNull.Value)
			{
				m_dRemainingQty = (decimal)oDataRow["Calculated_RemainingQty"];
			}
			if ((oDataRow["Calculated_ShipToName"] != DBNull.Value) && (string.IsNullOrEmpty((string)oDataRow["Calculated_ShipToName"]) == false))
			{
				m_sShipToName = (string)oDataRow["Calculated_ShipToName"];
			}
			if ((oDataRow["Calculated_Address1"] != DBNull.Value) && (string.IsNullOrEmpty((string)oDataRow["Calculated_Address1"]) == false))
			{
				m_sAddress1 = (string)oDataRow["Calculated_Address1"];
			}
			if ((oDataRow["Calculated_Address2"] != DBNull.Value) && (string.IsNullOrEmpty((string)oDataRow["Calculated_Address2"]) == false))
			{
				m_sAddress2 = (string)oDataRow["Calculated_Address2"];
			}
			if ((oDataRow["Calculated_Address3"] != DBNull.Value) && (string.IsNullOrEmpty((string)oDataRow["Calculated_Address3"]) == false))
			{
				m_sAddress3 = (string)oDataRow["Calculated_Address3"];
			}
			if ((oDataRow["Calculated_City"] != DBNull.Value) && (string.IsNullOrEmpty((string)oDataRow["Calculated_City"]) == false))
			{
				m_sCity = (string)oDataRow["Calculated_City"];
			}
			if ((oDataRow["Calculated_State"] != DBNull.Value) && (string.IsNullOrEmpty((string)oDataRow["Calculated_State"]) == false))
			{
				m_sState = (string)oDataRow["Calculated_State"];
			}
			if ((oDataRow["Calculated_Zip"] != DBNull.Value) && (string.IsNullOrEmpty((string)oDataRow["Calculated_Zip"]) == false))
			{
				m_sZip = (string)oDataRow["Calculated_Zip"];
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

		public int OrderRelNum
		{
			get { return m_iOrderRelNum; }
			set { m_iOrderRelNum = value; }
		}

		public DateTime RequiredByDate
		{
			get { return m_dtRequiredBy; }
			set { m_dtRequiredBy = value; }
		}

		public decimal OurRequestedQty
		{
			get { return m_dOurRequestedQty; }
			set { m_dOurRequestedQty = value; }
		}

		public bool OpenRelease
		{
			get { return m_bOpenRelease; }
			set { m_bOpenRelease = value; }
		}

        public bool VoidRelease
        {
            get { return m_bVoidRelease; }
            set { m_bVoidRelease = value; }
        }
        public bool FirmRelease
		{
			get { return m_bFirmRelease; }
			set { m_bFirmRelease = value; }
		}

		public bool MakeDirect
		{
			get { return m_bMakeDirect; }
			set { m_bMakeDirect = value; }
		}

		public decimal OurJobShippedQty
		{
			get { return m_dOurJobShippedQty; }
			set { m_dOurJobShippedQty = value; }
		}

		public decimal OurStockShippedQty
		{
			get { return m_dOurStockShippedQty; }
			set { m_dOurStockShippedQty = value; }
		}

		public string Plant
		{
			get { return m_sPlant; }
			set { m_sPlant = value; }
		}

		public string Warehouse
		{
			get { return m_sWarehouse; }
			set { m_sWarehouse = value; }
		}

		public string PartNum
		{
			get { return m_sPartNum; }
			set { m_sPartNum = value; }
		}

		public string PartRev
		{
			get { return m_sPartRev; }
			set { m_sPartRev = value; }
		}

		public DateTime NeedBy
		{
			get { return m_dtNeedBy; }
			set { m_dtNeedBy = value; }
		}

		public string ReleaseStatus
		{
			get { return m_sReleaseStatus; }
			set { m_sReleaseStatus = value; }
		}

		public bool BuyToOrder
		{
			get { return m_bBuyToOrder; }
			set { m_bBuyToOrder = value; }
		}

		public bool DropShip
		{
			get { return m_bDropShip; }
			set { m_bDropShip = value; }
		}

        public int PONum
        {
            get { return m_iPONum; }
            set { m_iPONum = value; }
        }
        public int POLine
        {
            get { return m_iPOLine; }
            set { m_iPOLine = value; }
        }
        public int PORelNum
        {
            get { return m_iPORelNum; }
            set { m_iPORelNum = value; }
        }
        public string ChangedBy
        {
            get { return m_sChangedBy; }
            set { m_sChangedBy = value; }
        }
        public DateTime ChangeDate
        {
            get { return m_dtChangeDate; }
            set { m_dtChangeDate = value; }
        }
        public bool HoldOrder
		{
			get { return m_bHoldOrder; }
			set { m_bHoldOrder = value; }
		}
        public bool VoidOrder
        {
            get { return m_bVoidOrder; }
            set { m_bVoidOrder = value; }
        }
        public string EntryPerson
		{
			get { return m_sEntryPerson; }
			set { m_sEntryPerson = value; }
		}

		public DateTime OrderDate
		{
			get { return m_dtOrderDate; }
			set { m_dtOrderDate = value; }
		}

		public string TermsCode
		{
			get { return m_sTermsCode; }
			set { m_sTermsCode = value; }
		}

		public decimal DiscountPercent
		{
			get { return m_dDiscountPercent; }
			set { m_dDiscountPercent = value; }
		}

		public string SalesReps
		{
			get { return m_sSalesReps; }
			set { m_sSalesReps = value; }
		}

		public string OrderComment
		{
			get { return m_sOrderComment; }
			set { m_sOrderComment = value; }
		}

		public string ShipComment
		{
			get { return m_sShipComment; }
			set { m_sShipComment = value; }
		}

		public string InvoiceComment
		{
			get { return m_sInvoiceComment; }
			set { m_sInvoiceComment = value; }
		}

		public string PickListComment
		{
			get { return m_sPickListComment; }
			set { m_sPickListComment = value; }
		}

		public decimal ExchangeRate
		{
			get { return m_dExchangeRate; }
			set { m_dExchangeRate = value; }
		}

		public string CurrencyCode
		{
			get { return m_sCurrencyCode; }
			set { m_sCurrencyCode = value; }
		}

		public bool LockRate
		{
			get { return m_bLockRate; }
			set { m_bLockRate = value; }
		}

		public int RepSplit1
		{
			get { return m_iRepSplit1; }
			set { m_iRepSplit1 = value; }
		}

		public int RepSplit2
		{
			get { return m_iRepSplit2; }
			set { m_iRepSplit2 = value; }
		}

		public bool ReadyToCalc
		{
			get { return m_bReadyToCalc; }
			set { m_bReadyToCalc = value; }
		}

		public string OrderStatus
		{
			get { return m_sOrderStatus; }
			set { m_sOrderStatus = value; }
		}

		public bool ReadyToFulfill
		{
			get { return m_bReadyToFulfill; }
			set { m_bReadyToFulfill = value; }
		}

		public string MarketSegmentID
		{
			get { return m_sMarketSegmentID; }
			set { m_sMarketSegmentID = value; }
		}

		public string MarketSegmentDescription
		{
			get { return m_sMarketSegmentDescription; }
			set { m_sMarketSegmentDescription = value; }
		}
        public bool VoidLine
        {
            get { return m_bVoidLine; }
            set { m_bVoidLine = value; }
        }
        public string LineType
		{
			get { return m_sLineType; }
			set { m_sLineType = value; }
		}

		public string PartDescription
		{
			get { return m_sPartDescription; }
			set { m_sPartDescription = value; }
		}

		public decimal UnitPrice
		{
			get { return m_dUnitPrice; }
			set { m_dUnitPrice = value; }
		}

		public decimal DocUnitPrice
		{
			get { return m_dDocUnitPrice; }
			set { m_dDocUnitPrice = value; }
		}

		public int QuoteNum
		{
			get { return m_iQuoteNum; }
			set { m_iQuoteNum = value; }
		}

		public int QuoteLine
		{
			get { return m_iQuoteLine; }
			set { m_iQuoteLine = value; }
		}

		public string LineStatus
		{
			get { return m_sLineStatus; }
			set { m_sLineStatus = value; }
		}

		public string MfgJobType
		{
			get { return m_sMfgJobType; }
			set { m_sMfgJobType = value; }
		}

		public string PortfolioID
		{
			get { return m_sPortfolioID; }
			set { m_sPortfolioID = value; }
		}

		public string PortfolioDescription
		{
			get { return m_sPortfolioDescription; }
			set { m_sPortfolioDescription = value; }
		}

		public string ClassID
		{
			get { return m_sClassID; }
			set { m_sClassID = value; }
		}

		public string ClassDescription
		{
			get { return m_sClassDescription; }
			set { m_sClassDescription = value; }
		}

		public string GroupID
		{
			get { return m_sGroupID; }
			set { m_sGroupID = value; }
		}

		public string GroupDescription
		{
			get { return m_sGroupDescription; }
			set { m_sGroupDescription = value; }
		}

		public string PartType
		{
			get { return m_sPartType; }
			set { m_sPartType = value; }
		}

		public bool NonStockItem
		{
			get { return m_bNonStockItem; }
			set { m_bNonStockItem = value; }
		}

		public string IUOM
		{
			get { return m_sIUOM; }
			set { m_sIUOM = value; }
		}

		public string PUOM
		{
			get { return m_sPUOM; }
			set { m_sPUOM = value; }
		}

		public decimal PartUnitPrice
		{
			get { return m_dPartUnitPrice; }
			set { m_dPartUnitPrice = value; }
		}

		public string PricePer
		{
			get { return m_sPricePer; }
			set { m_sPricePer = value; }
		}

		public string CostingMethod
		{
			get { return m_sCostingMethod; }
			set { m_sCostingMethod = value; }
		}

		public bool Inactive
		{
			get { return m_bInactive; }
			set { m_bInactive = value; }
		}

		public Boolean TrackLots
		{
			get { return m_bTrackLots; }
			set { m_bTrackLots = value; }
		}

		public bool TrackSerial
		{
			get { return m_bTrackSerial; }
			set { m_bTrackSerial = value; }
		}

		public string SUOM
		{
			get { return m_sSUOM; }
			set { m_sSUOM = value; }
		}

		public decimal SellingFactor
		{
			get { return m_dSellingFactor; }
			set { m_dSellingFactor = value; }
		}

		public bool UsePartRev
		{
			get { return m_bUsePartRev; }
			set { m_bUsePartRev = value; }
		}

		public bool PartOnHold
		{
			get { return m_bPartOnHold; }
			set { m_bPartOnHold = value; }
		}

		public bool QtyBearing
		{
			get { return m_bQtyBearing; }
			set { m_bQtyBearing = value; }
		}

		public decimal ExtPrice
		{
			get { return m_dExtPrice; }
			set { m_dExtPrice = value; }
		}

		public decimal RemainingQty
		{
			get { return m_dRemainingQty; }
			set { m_dRemainingQty = value; }
		}

		public string ShipToName
		{
			get { return m_sShipToName; }
			set { m_sShipToName = value; }
		}

		public string Address1
		{
			get { return m_sAddress1; }
			set { m_sAddress1 = value; }
		}

		public string Address2
		{
			get { return m_sAddress2; }
			set { m_sAddress2 = value; }
		}

		public string Address3
		{
			get { return m_sAddress3; }
			set { m_sAddress3 = value; }
		}

		public string City
		{
			get { return m_sCity; }
			set { m_sCity = value; }
		}

		public string State
		{
			get { return m_sState; }
			set { m_sState = value; }
		}

		public string Zip
		{
			get { return m_sZip; }
			set { m_sZip = value; }
		}

		#endregion

		#region Data Members
		private string m_sCompany;
		private int m_iOrderNum;
		private int m_iOrderLine;
		private int m_iOrderRelNum;
		private DateTime m_dtRequiredBy;
		private decimal m_dOurRequestedQty;
		private bool m_bOpenRelease;
        private bool m_bVoidRelease;
		private bool m_bFirmRelease;
		private bool m_bMakeDirect;
		private decimal m_dOurJobShippedQty;
		private decimal m_dOurStockShippedQty;
		private string m_sPlant;
		private string m_sWarehouse;
		private string m_sPartNum;
		private string m_sPartRev;
		private DateTime m_dtNeedBy;
		private string m_sReleaseStatus;
		private bool m_bBuyToOrder;
		private bool m_bDropShip;
        private int m_iPONum;
        private int m_iPOLine;
        private int m_iPORelNum;
        private string m_sChangedBy;
        private DateTime m_dtChangeDate;
		private bool m_bHoldOrder;
        private bool m_bVoidOrder;
		private string m_sEntryPerson;
		private DateTime m_dtOrderDate;
		private string m_sTermsCode;
		private decimal m_dDiscountPercent;
		private string m_sSalesReps;
		private string m_sOrderComment;
		private string m_sShipComment;
		private string m_sInvoiceComment;
		private string m_sPickListComment;
		private decimal m_dExchangeRate;
		private string m_sCurrencyCode;
		private bool m_bLockRate;
		private int m_iRepSplit1;
		private int m_iRepSplit2;
		private bool m_bReadyToCalc;
		private string m_sOrderStatus;
		private bool m_bReadyToFulfill;
		private string m_sMarketSegmentID;
		private string m_sMarketSegmentDescription;
        private bool m_bVoidLine;
		private string m_sLineType;
		private string m_sPartDescription;
		private decimal m_dUnitPrice;
		private decimal m_dDocUnitPrice;
		private int m_iQuoteNum;
		private int m_iQuoteLine;
		private string m_sLineStatus;
		private string m_sMfgJobType;
		private string m_sPortfolioID;
		private string m_sPortfolioDescription;
		private string m_sClassID;
		private string m_sClassDescription;
		private string m_sGroupID;
		private string m_sGroupDescription;
		private string m_sPartType;
		private bool m_bNonStockItem;
		private string m_sIUOM;
		private string m_sPUOM;
		private decimal m_dPartUnitPrice;
		private string m_sPricePer;
		private string m_sCostingMethod;
		private bool m_bInactive;
		private bool m_bTrackLots;
		private bool m_bTrackSerial;
		private string m_sSUOM;
		private decimal m_dSellingFactor;
		private bool m_bUsePartRev;
		private bool m_bPartOnHold;
		private bool m_bQtyBearing;
		private decimal m_dExtPrice;
		private decimal m_dRemainingQty;
		private string m_sShipToName;
		private string m_sAddress1;
		private string m_sAddress2;
		private string m_sAddress3;
		private string m_sCity;
		private string m_sState;
		private string m_sZip;
		#endregion
	}

	public class ValidateSalesOrders
	{
		#region Constructors
		public ValidateSalesOrders()
		{

		}
		#endregion

		#region Methods
		public bool Initialize(Session oSession, string sCompany, HSValidateParts oValidateParts)
		{
			bool bSuccess = true;

            m_oValidateParts = oValidateParts;

            m_oBOMSupport = new BOMSupport(sCompany);
            if (m_oBOMSupport.Initialize(oSession, m_oValidateParts) == false)
            {
                Console.WriteLine("Failed to load the BOM Support object!");
            }

            //****Set a parameter Value***** 
            Ice.Proxy.BO.DynamicQueryImpl oDynamicQuery = WCFServiceSupport.CreateImpl<Ice.Proxy.BO.DynamicQueryImpl>(oSession, Ice.Proxy.BO.DynamicQueryImpl.UriPath);
			QueryExecutionDataSet oQueryExecutionDataSet = oDynamicQuery.GetQueryExecutionParametersByID(BAQConstants.QUERY_SALES_ORDER_BACKLOG);
			oQueryExecutionDataSet.ExecutionParameter.Clear();
			DataSet oDataSet = oDynamicQuery.ExecuteByID(BAQConstants.QUERY_SALES_ORDER_BACKLOG, oQueryExecutionDataSet);

			m_oAllOpenOrders.Clear();
			foreach (DataRow oRow in oDataSet.Tables[0].Rows)
			{
				m_oAllOpenOrders.Add(new SOBacklog(oRow));
			}

			return bSuccess;
		}
        public void PerformValidation(string sCompany, string sTmpFileDirectory)
        {
            DateTime dtToday = DateTime.Now;
            string sDestinationFileName = sTmpFileDirectory + sCompany + "-SO BOM Validation-" + dtToday.Month.ToString() + "-" + dtToday.Day.ToString() + "-" + dtToday.Year.ToString() + ".xlsx";

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

            HSUser oServiceAccount = HSUser.GetUserById(HSUser.SPINCRAFT_SERVICE_ACCOUNT_ID);
            List<string> oToAddresses = new List<string>();
            bool bFirstWorksheet = true;

            #region BOM Management Issues
            SLDocument oSLBOMMgrDocument = new SLDocument();

            System.Drawing.Color oHighlightYellow = System.Drawing.Color.FromArgb(255, 255, 0);

            SLStyle oHighlightHeaderStyle = new SLStyle();
            oHighlightHeaderStyle.SetFontBold(true);
            oHighlightHeaderStyle.SetFont(FontSchemeValues.Major, 12);
            oHighlightHeaderStyle.SetHorizontalAlignment(HorizontalAlignmentValues.Center);
            oHighlightHeaderStyle.SetPatternFill(PatternValues.Solid, oHighlightYellow, oHighlightYellow);
            oHighlightHeaderStyle.SetFontColor(SLThemeColorIndexValues.Dark2Color);

            //
            // BOM management issues
            //       
            oToAddresses = HSUser.GetEmailsForUsersInGroup(HSUser.REPORT_ON_MANUFACTURED_PART_ISSUES);
            oToAddresses.Add(HSUser.SPINCRAFT_ROOT_USER);

            List<SOBacklog> oPartsWithoutRevision = m_oAllOpenOrders.Where(oItem => (string.IsNullOrEmpty(oItem.PartRev) == true)).ToList();
            if (oPartsWithoutRevision.Count != 0)
            {
                // we do not need to repeat every release as we only need to indicate which line on the sales order does not have the rev specified
                List<SOBacklog> oDistinctSalesOrderLines = new List<SOBacklog>();
                foreach (SOBacklog oTmp in oPartsWithoutRevision)
                {
                    SOBacklog oFound = oDistinctSalesOrderLines.FirstOrDefault(oItem => (oItem.OrderNum == oTmp.OrderNum) && (oItem.OrderLine == oTmp.OrderLine));
                    if (oFound == null)
                    {
                        // not in the list yet
                        oDistinctSalesOrderLines.Add(oTmp);
                    }
                }

                if (bFirstWorksheet == true)
                {
                    oSLBOMMgrDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "Parts No Rev");
                    bFirstWorksheet = false;
                }
                else
                {
                    oSLBOMMgrDocument.AddWorksheet("Parts No Rev");
                }

                //set up column headers
                int iNumOfColumns = 0;
                iNumOfColumns++;
                oSLBOMMgrDocument.SetCellValue(1, iNumOfColumns, "SO Num");
                iNumOfColumns++;
                oSLBOMMgrDocument.SetCellValue(1, iNumOfColumns, "Line Num");
                iNumOfColumns++;
                oSLBOMMgrDocument.SetCellValue(1, iNumOfColumns, "Order Date");
                oSLBOMMgrDocument.SetColumnWidth(iNumOfColumns, 15);
                iNumOfColumns++;
                oSLBOMMgrDocument.SetCellValue(1, iNumOfColumns, "PartNum");
                oSLBOMMgrDocument.SetColumnWidth(iNumOfColumns, 15);
                iNumOfColumns++;
                oSLBOMMgrDocument.SetCellValue(1, iNumOfColumns, "Part Orders Without Revisions");
                oSLBOMMgrDocument.SetCellStyle(1, iNumOfColumns, oHighlightHeaderStyle);
                oSLBOMMgrDocument.SetColumnWidth(iNumOfColumns, 60);

                int iNumOfRows = 1;
                foreach (SOBacklog oSalesOrder in oDistinctSalesOrderLines)
                {
                    iNumOfRows++;
                    oSLBOMMgrDocument.SetCellValue(iNumOfRows, 1, oSalesOrder.OrderNum);
                    oSLBOMMgrDocument.SetCellValue(iNumOfRows, 2, oSalesOrder.OrderLine);
                    oSLBOMMgrDocument.SetCellValue(iNumOfRows, 3, oSalesOrder.OrderDate.ToShortDateString());
                    oSLBOMMgrDocument.SetCellValue(iNumOfRows, 4, oSalesOrder.PartNum);
                }
            }

            // we need to check the part num for all sales orders to see if it is a valid part
            MfgPart oFinishedGood = null;

            List<SOBacklog> oOrdersWithUnapprovedRevs = new List<SOBacklog>();
            List<SOBacklog> oOrdersWithMfgGoodsWithNoMaterials = new List<SOBacklog>();
            List<SOBacklog> oOrdersWithMfgGoodsWithNoOperations = new List<SOBacklog>();

            List<SOBacklog> oOrdersForFinishedGoods = m_oAllOpenOrders.Where(oItem => string.Compare(oItem.PartType, "M", true) == 0).ToList();
            foreach (SOBacklog oSalesOrder in m_oAllOpenOrders)
            {
                // find the finished good
                oFinishedGood = m_oBOMSupport.GetPMfgPart(oSalesOrder.PartNum, oSalesOrder.PartRev);
                if (oFinishedGood != null)
                {
                    if (oFinishedGood.RevApproved == false)
                    {
                        oOrdersWithUnapprovedRevs.Add(oSalesOrder);
                    }

                    if (oFinishedGood.MyPartMaterials.Count == 0)
                    {
                        oOrdersWithMfgGoodsWithNoMaterials.Add(oSalesOrder);
                    }

                    if (oFinishedGood.MyPartOperations.Count == 0)
                    {
                        oOrdersWithMfgGoodsWithNoOperations.Add(oSalesOrder);
                    }
                }
            }

            // orders without an apoproved rev
            if (oOrdersWithUnapprovedRevs.Count != 0)
            {
                // we do not need to repeat every release as we only need to indicate which line on the sales order does not have the rev specified
                List<SOBacklog> oDistinctSalesOrderLines = new List<SOBacklog>();
                foreach (SOBacklog oTmp in oOrdersWithUnapprovedRevs)
                {
                    SOBacklog oFound = oDistinctSalesOrderLines.FirstOrDefault(oItem => (oItem.OrderNum == oTmp.OrderNum) && (oItem.OrderLine == oTmp.OrderLine));
                    if (oFound == null)
                    {
                        // not in the list yet
                        oDistinctSalesOrderLines.Add(oTmp);
                    }
                }

                if (bFirstWorksheet == true)
                {
                    oSLBOMMgrDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "No Appr Rev");
                    bFirstWorksheet = false;
                }
                else
                {
                    oSLBOMMgrDocument.AddWorksheet("No Appr Rev");
                }

                //set up column headers
                int iNumOfColumns = 0;
                iNumOfColumns++;
                oSLBOMMgrDocument.SetCellValue(1, iNumOfColumns, "SO Num");
                iNumOfColumns++;
                oSLBOMMgrDocument.SetCellValue(1, iNumOfColumns, "Line Num");
                iNumOfColumns++;
                oSLBOMMgrDocument.SetCellValue(1, iNumOfColumns, "Order Date");
                oSLBOMMgrDocument.SetColumnWidth(iNumOfColumns, 15);
                iNumOfColumns++;
                oSLBOMMgrDocument.SetCellValue(1, iNumOfColumns, "Part Num");
                oSLBOMMgrDocument.SetColumnWidth(iNumOfColumns, 15);
                iNumOfColumns++;
                oSLBOMMgrDocument.SetCellValue(1, iNumOfColumns, "Rev Num");
                oSLBOMMgrDocument.SetColumnWidth(iNumOfColumns, 15);
                iNumOfColumns++;
                oSLBOMMgrDocument.SetCellValue(1, iNumOfColumns, "Order For Parts Without Approved Revision");
                oSLBOMMgrDocument.SetCellStyle(1, iNumOfColumns, oHighlightHeaderStyle);
                oSLBOMMgrDocument.SetColumnWidth(iNumOfColumns, 60);

                int iNumOfRows = 1;
                foreach (SOBacklog oSalesOrder in oDistinctSalesOrderLines)
                {
                    iNumOfRows++;
                    oSLBOMMgrDocument.SetCellValue(iNumOfRows, 1, oSalesOrder.OrderNum);
                    oSLBOMMgrDocument.SetCellValue(iNumOfRows, 2, oSalesOrder.OrderLine);
                    oSLBOMMgrDocument.SetCellValue(iNumOfRows, 3, oSalesOrder.OrderDate.ToShortDateString());
                    oSLBOMMgrDocument.SetCellValue(iNumOfRows, 4, oSalesOrder.PartNum);
                    oSLBOMMgrDocument.SetCellValue(iNumOfRows, 5, oSalesOrder.PartRev);
                }
            }

            // orders where the finished good has no materials
            if (oOrdersWithMfgGoodsWithNoMaterials.Count != 0)
            {
                // we do not need to repeat every release as we only need to indicate which line on the sales order do not have materials
                List<SOBacklog> oDistinctSalesOrderLines = new List<SOBacklog>();
                foreach (SOBacklog oTmp in oOrdersWithMfgGoodsWithNoMaterials)
                {
                    SOBacklog oFound = oDistinctSalesOrderLines.FirstOrDefault(oItem => (oItem.OrderNum == oTmp.OrderNum) && (oItem.OrderLine == oTmp.OrderLine));
                    if (oFound == null)
                    {
                        // not in the list yet
                        oDistinctSalesOrderLines.Add(oTmp);
                    }
                }

                if (bFirstWorksheet == true)
                {
                    oSLBOMMgrDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "Mfg Part No Materials");
                    bFirstWorksheet = false;
                }
                else
                {
                    oSLBOMMgrDocument.AddWorksheet("Mfg Part No Materials");
                }

                //set up column headers
                int iNumOfColumns = 0;
                iNumOfColumns++;
                oSLBOMMgrDocument.SetCellValue(1, iNumOfColumns, "SO Num");
                iNumOfColumns++;
                oSLBOMMgrDocument.SetCellValue(1, iNumOfColumns, "Line Num");
                iNumOfColumns++;
                oSLBOMMgrDocument.SetCellValue(1, iNumOfColumns, "Order Date");
                oSLBOMMgrDocument.SetColumnWidth(iNumOfColumns, 15);
                iNumOfColumns++;
                oSLBOMMgrDocument.SetCellValue(1, iNumOfColumns, "PartNum");
                oSLBOMMgrDocument.SetColumnWidth(iNumOfColumns, 15);
                iNumOfColumns++;
                oSLBOMMgrDocument.SetCellValue(1, iNumOfColumns, "Orders For Mfg Parts Having No Materials");
                oSLBOMMgrDocument.SetCellStyle(1, iNumOfColumns, oHighlightHeaderStyle);
                oSLBOMMgrDocument.SetColumnWidth(iNumOfColumns, 60);

                int iNumOfRows = 1;
                foreach (SOBacklog oSalesOrder in oDistinctSalesOrderLines)
                {
                    iNumOfRows++;
                    oSLBOMMgrDocument.SetCellValue(iNumOfRows, 1, oSalesOrder.OrderNum);
                    oSLBOMMgrDocument.SetCellValue(iNumOfRows, 2, oSalesOrder.OrderLine);
                    oSLBOMMgrDocument.SetCellValue(iNumOfRows, 3, oSalesOrder.OrderDate.ToShortDateString());
                    oSLBOMMgrDocument.SetCellValue(iNumOfRows, 4, oSalesOrder.PartNum);
                }
            }
      
            // orders where the finished good has no operations
            if (oOrdersWithMfgGoodsWithNoOperations.Count != 0)
            {
                // we do not need to repeat every release as we only need to indicate which line on the sales order do not have operations
                List<SOBacklog> oDistinctSalesOrderLines = new List<SOBacklog>();
                foreach (SOBacklog oTmp in oOrdersWithMfgGoodsWithNoOperations)
                {
                    SOBacklog oFound = oDistinctSalesOrderLines.FirstOrDefault(oItem => (oItem.OrderNum == oTmp.OrderNum) && (oItem.OrderLine == oTmp.OrderLine));
                    if (oFound == null)
                    {
                        // not in the list yet
                        oDistinctSalesOrderLines.Add(oTmp);
                    }
                }

                if (bFirstWorksheet == true)
                {
                    oSLBOMMgrDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "Mfg Part No Operations");
                    bFirstWorksheet = false;
                }
                else
                {
                    oSLBOMMgrDocument.AddWorksheet("Mfg Part No Operations");
                }

                //set up column headers
                int iNumOfColumns = 0;
                iNumOfColumns++;
                oSLBOMMgrDocument.SetCellValue(1, iNumOfColumns, "SO Num");
                iNumOfColumns++;
                oSLBOMMgrDocument.SetCellValue(1, iNumOfColumns, "Line Num");
                iNumOfColumns++;
                oSLBOMMgrDocument.SetCellValue(1, iNumOfColumns, "Order Date");
                oSLBOMMgrDocument.SetColumnWidth(iNumOfColumns, 15);
                iNumOfColumns++;
                oSLBOMMgrDocument.SetCellValue(1, iNumOfColumns, "PartNum");
                oSLBOMMgrDocument.SetColumnWidth(iNumOfColumns, 15);
                iNumOfColumns++;
                oSLBOMMgrDocument.SetCellValue(1, iNumOfColumns, "Orders For Mfg Parts Having No Operations");
                oSLBOMMgrDocument.SetCellStyle(1, iNumOfColumns, oHighlightHeaderStyle);
                oSLBOMMgrDocument.SetColumnWidth(iNumOfColumns, 60);

                int iNumOfRows = 1;
                foreach (SOBacklog oSalesOrder in oDistinctSalesOrderLines)
                {
                    iNumOfRows++;
                    oSLBOMMgrDocument.SetCellValue(iNumOfRows, 1, oSalesOrder.OrderNum);
                    oSLBOMMgrDocument.SetCellValue(iNumOfRows, 2, oSalesOrder.OrderLine);
                    oSLBOMMgrDocument.SetCellValue(iNumOfRows, 3, oSalesOrder.OrderDate.ToShortDateString());
                    oSLBOMMgrDocument.SetCellValue(iNumOfRows, 4, oSalesOrder.PartNum);
                }
            }
            
            // send the email to the BOM manager
            if (bFirstWorksheet == false)
            {
                // we create a worksheet so we have some data in the spreadsheet so we email the spreadsheet
                oSLBOMMgrDocument.SaveAs(sDestinationFileName);
                List<string> oAttachments = new List<string>();
                oAttachments.Add(sDestinationFileName);
                HSEmailHelper.SendEmail(oToAddresses, sCompany + " Sales Order Validation - BOM", sCompany + " Sales Order Validation - BOM", oAttachments);
            }
            #endregion

            #region Customer Service Issues
            //
            // Customer Service Issues
            //
            bFirstWorksheet = true;
            sDestinationFileName = sTmpFileDirectory + sCompany + "-SO Validation-" + dtToday.Month.ToString() + "-" + dtToday.Day.ToString() + "-" + dtToday.Year.ToString() + ".xlsx";
            SLDocument oSLSOMgrDocument = new SLDocument();
            oToAddresses = HSUser.GetEmailsForUsersInGroup(HSUser.REPORT_ON_SO_ISSUES);
            oToAddresses.Add(HSUser.SPINCRAFT_ROOT_USER);


            // parts on the fly
            List<SOBacklog> oPartsOnTheFly = m_oAllOpenOrders.Where(oItem => string.IsNullOrEmpty(oItem.PartType) == true).ToList();
            if (oPartsOnTheFly.Count != 0)
            {
                // we do not need to repeat every release as we only need to indicate which line has the issue
                List<SOBacklog> oDistinctSalesOrderLines = new List<SOBacklog>();
                foreach (SOBacklog oTmp in oPartsOnTheFly)
                {
                    SOBacklog oFound = oDistinctSalesOrderLines.FirstOrDefault(oItem => (oItem.OrderNum == oTmp.OrderNum) && (oItem.OrderLine == oTmp.OrderLine));
                    if (oFound == null)
                    {
                        // not in the list yet
                        oDistinctSalesOrderLines.Add(oTmp);
                    }
                }

                if (bFirstWorksheet == true)
                {
                    oSLSOMgrDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "Parts Not Found");
                    bFirstWorksheet = false;
                }
                else
                {
                    oSLSOMgrDocument.AddWorksheet("Parts Not Found");
                }

                //set up column headers
                int iNumOfColumns = 0;
                iNumOfColumns++;
                oSLSOMgrDocument.SetCellValue(1, iNumOfColumns, "SO Num");
                iNumOfColumns++;
                oSLSOMgrDocument.SetCellValue(1, iNumOfColumns, "Line Num");
                iNumOfColumns++;
                oSLSOMgrDocument.SetCellValue(1, iNumOfColumns, "Order Date");
                oSLSOMgrDocument.SetColumnWidth(iNumOfColumns, 15);
                iNumOfColumns++;
                oSLSOMgrDocument.SetCellValue(1, iNumOfColumns, "PartNum");
                oSLSOMgrDocument.SetColumnWidth(iNumOfColumns, 15);
                iNumOfColumns++;
                oSLSOMgrDocument.SetCellValue(1, iNumOfColumns, "Entry Person");
                oSLSOMgrDocument.SetColumnWidth(iNumOfColumns, 15);
                iNumOfColumns++;
                oSLSOMgrDocument.SetCellValue(1, iNumOfColumns, "Part Does Not Exist");
                oSLSOMgrDocument.SetCellStyle(1, iNumOfColumns, oHighlightHeaderStyle);
                oSLSOMgrDocument.SetColumnWidth(iNumOfColumns, 30);

                int iNumOfRows = 1;
                foreach (SOBacklog oSalesOrder in oDistinctSalesOrderLines)
                {
                    iNumOfRows++;
                    oSLSOMgrDocument.SetCellValue(iNumOfRows, 1, oSalesOrder.OrderNum);
                    oSLSOMgrDocument.SetCellValue(iNumOfRows, 2, oSalesOrder.OrderLine);
                    oSLSOMgrDocument.SetCellValue(iNumOfRows, 3, oSalesOrder.OrderDate.ToShortDateString());
                    oSLSOMgrDocument.SetCellValue(iNumOfRows, 4, oSalesOrder.PartNum);
                    oSLSOMgrDocument.SetCellValue(iNumOfRows, 5, oSalesOrder.EntryPerson);
                }
            }

            // orders on hold
            List<SOBacklog> oOrdersOnHold = m_oAllOpenOrders.Where(oItem => oItem.HoldOrder == true).ToList();
            if (oOrdersOnHold.Count != 0)
            {
                // we do not need to repeat every release as we only need to indicate which order has the issue
                List<SOBacklog> oDistinctSalesOrders = new List<SOBacklog>();
                foreach (SOBacklog oTmp in oPartsOnTheFly)
                {
                    SOBacklog oFound = oDistinctSalesOrders.FirstOrDefault(oItem => (oItem.OrderNum == oTmp.OrderNum));
                    if (oFound == null)
                    {
                        // not in the list yet
                        oDistinctSalesOrders.Add(oTmp);
                    }
                }

                if (bFirstWorksheet == true)
                {
                    oSLSOMgrDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "Order On Hold");
                    bFirstWorksheet = false;
                }
                else
                {
                    oSLSOMgrDocument.AddWorksheet("Order On Hold");
                }

                //set up column headers
                int iNumOfColumns = 0;
                iNumOfColumns++;
                oSLSOMgrDocument.SetCellValue(1, iNumOfColumns, "SO Num");
                iNumOfColumns++;
                oSLSOMgrDocument.SetCellValue(1, iNumOfColumns, "Order Date");
                oSLSOMgrDocument.SetColumnWidth(iNumOfColumns, 15);
                iNumOfColumns++;
                oSLSOMgrDocument.SetCellValue(1, iNumOfColumns, "Entered By");
                oSLSOMgrDocument.SetColumnWidth(iNumOfColumns, 15);
                iNumOfColumns++;
                oSLSOMgrDocument.SetCellValue(1, iNumOfColumns, "Changed By");
                oSLSOMgrDocument.SetColumnWidth(iNumOfColumns, 15);
                iNumOfColumns++;
                oSLSOMgrDocument.SetCellValue(1, iNumOfColumns, "Changed Date");
                oSLSOMgrDocument.SetColumnWidth(iNumOfColumns, 15);
                iNumOfColumns++;
                oSLSOMgrDocument.SetCellValue(1, iNumOfColumns, "Orders On Hold");
                oSLSOMgrDocument.SetCellStyle(1, iNumOfColumns, oHighlightHeaderStyle);
                oSLSOMgrDocument.SetColumnWidth(iNumOfColumns, 30);

                int iNumOfRows = 1;
                foreach (SOBacklog oSalesOrder in oDistinctSalesOrders)
                {
                    iNumOfRows++;
                    oSLSOMgrDocument.SetCellValue(iNumOfRows, 1, oSalesOrder.OrderNum);
                    oSLSOMgrDocument.SetCellValue(iNumOfRows, 2, oSalesOrder.OrderDate.ToShortDateString());
                    oSLSOMgrDocument.SetCellValue(iNumOfRows, 3, oSalesOrder.EntryPerson);
                    oSLSOMgrDocument.SetCellValue(iNumOfRows, 4, oSalesOrder.ChangedBy);
                    oSLSOMgrDocument.SetCellValue(iNumOfRows, 5, oSalesOrder.ChangeDate.ToShortDateString());
                }
            }

            // orders with zero qty
            List<SOBacklog> oOrdersWithZeroQuantity = m_oAllOpenOrders.Where(oItem => oItem.OurRequestedQty == 0).ToList();
            if (oOrdersWithZeroQuantity.Count != 0)
            {
                // we do not need to repeat every release as we only need to indicate which line has the issue
                List<SOBacklog> oDistinctSalesOrderLines = new List<SOBacklog>();
                foreach (SOBacklog oTmp in oOrdersWithZeroQuantity)
                {
                    SOBacklog oFound = oDistinctSalesOrderLines.FirstOrDefault(oItem => (oItem.OrderNum == oTmp.OrderNum) && (oItem.OrderLine == oTmp.OrderLine));
                    if (oFound == null)
                    {
                        // not in the list yet
                        oDistinctSalesOrderLines.Add(oTmp);
                    }
                }

                if (bFirstWorksheet == true)
                {
                    oSLSOMgrDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "Zero Qty");
                    bFirstWorksheet = false;
                }
                else
                {
                    oSLSOMgrDocument.AddWorksheet("Zero Qty");
                }

                //set up column headers
                int iNumOfColumns = 0;
                iNumOfColumns++;
                oSLSOMgrDocument.SetCellValue(1, iNumOfColumns, "SO Num");
                iNumOfColumns++;
                oSLSOMgrDocument.SetCellValue(1, iNumOfColumns, "Line Num");
                iNumOfColumns++;
                oSLSOMgrDocument.SetCellValue(1, iNumOfColumns, "Order Date");
                oSLSOMgrDocument.SetColumnWidth(iNumOfColumns, 15);
                iNumOfColumns++;
                oSLSOMgrDocument.SetCellValue(1, iNumOfColumns, "PartNum");
                oSLSOMgrDocument.SetColumnWidth(iNumOfColumns, 15);
                iNumOfColumns++;
                oSLSOMgrDocument.SetCellValue(1, iNumOfColumns, "Entry Person");
                oSLSOMgrDocument.SetColumnWidth(iNumOfColumns, 15);
                iNumOfColumns++;
                oSLSOMgrDocument.SetCellValue(1, iNumOfColumns, "Order Release Has A Zero Quantity");
                oSLSOMgrDocument.SetCellStyle(1, iNumOfColumns, oHighlightHeaderStyle);
                oSLSOMgrDocument.SetColumnWidth(iNumOfColumns, 60);

                int iNumOfRows = 1;
                foreach (SOBacklog oSalesOrder in oDistinctSalesOrderLines)
                {
                    iNumOfRows++;
                    oSLSOMgrDocument.SetCellValue(iNumOfRows, 1, oSalesOrder.OrderNum);
                    oSLSOMgrDocument.SetCellValue(iNumOfRows, 2, oSalesOrder.OrderLine);
                    oSLSOMgrDocument.SetCellValue(iNumOfRows, 3, oSalesOrder.OrderDate.ToShortDateString());
                    oSLSOMgrDocument.SetCellValue(iNumOfRows, 4, oSalesOrder.PartNum);
                    oSLSOMgrDocument.SetCellValue(iNumOfRows, 5, oSalesOrder.EntryPerson);
                }
            }

            // orders with no price
            List<SOBacklog> oOrdersWithNoPrice = m_oAllOpenOrders.Where(oItem => oItem.ExtPrice == 0).ToList();
            if (oOrdersWithNoPrice.Count != 0)
            {
                // we do not need to repeat every release as we only need to indicate which line has the issue
                List<SOBacklog> oDistinctSalesOrderLines = new List<SOBacklog>();
                foreach (SOBacklog oTmp in oOrdersWithNoPrice)
                {
                    SOBacklog oFound = oDistinctSalesOrderLines.FirstOrDefault(oItem => (oItem.OrderNum == oTmp.OrderNum) && (oItem.OrderLine == oTmp.OrderLine));
                    if (oFound == null)
                    {
                        // not in the list yet
                        oDistinctSalesOrderLines.Add(oTmp);
                    }
                }

                if (bFirstWorksheet == true)
                {
                    oSLSOMgrDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "Price Not Set");
                    bFirstWorksheet = false;
                }
                else
                {
                    oSLSOMgrDocument.AddWorksheet("Price Not Set");
                }

                //set up column headers
                int iNumOfColumns = 0;
                iNumOfColumns++;
                oSLSOMgrDocument.SetCellValue(1, iNumOfColumns, "SO Num");
                iNumOfColumns++;
                oSLSOMgrDocument.SetCellValue(1, iNumOfColumns, "Line Num");
                iNumOfColumns++;
                oSLSOMgrDocument.SetCellValue(1, iNumOfColumns, "Order Date");
                oSLSOMgrDocument.SetColumnWidth(iNumOfColumns, 15);
                iNumOfColumns++;
                oSLSOMgrDocument.SetCellValue(1, iNumOfColumns, "PartNum");
                oSLSOMgrDocument.SetColumnWidth(iNumOfColumns, 15);
                iNumOfColumns++;
                oSLSOMgrDocument.SetCellValue(1, iNumOfColumns, "Entry Person");
                oSLSOMgrDocument.SetColumnWidth(iNumOfColumns, 15);
                iNumOfColumns++;
                oSLSOMgrDocument.SetCellValue(1, iNumOfColumns, "Price  Has Not Been Set");
                oSLSOMgrDocument.SetCellStyle(1, iNumOfColumns, oHighlightHeaderStyle);
                oSLSOMgrDocument.SetColumnWidth(iNumOfColumns, 30);

                int iNumOfRows = 1;
                foreach (SOBacklog oSalesOrder in oDistinctSalesOrderLines)
                {
                    iNumOfRows++;
                    oSLSOMgrDocument.SetCellValue(iNumOfRows, 1, oSalesOrder.OrderNum);
                    oSLSOMgrDocument.SetCellValue(iNumOfRows, 2, oSalesOrder.OrderLine);
                    oSLSOMgrDocument.SetCellValue(iNumOfRows, 3, oSalesOrder.OrderDate.ToShortDateString());
                    oSLSOMgrDocument.SetCellValue(iNumOfRows, 4, oSalesOrder.PartNum);
                    oSLSOMgrDocument.SetCellValue(iNumOfRows, 5, oSalesOrder.EntryPerson);
                }
            }

            // stocking order with no price on part -- cant compute POC
            List<SOBacklog> oStockingOrdersWithoutPrice = new List<SOBacklog>();
            foreach (SOBacklog oSalesOrder in oOrdersForFinishedGoods)
            {
                // find the finished good
                oFinishedGood = m_oBOMSupport.GetPMfgPart(oSalesOrder.PartNum, oSalesOrder.PartRev);
                if (oFinishedGood != null)
                {
                    if ( (oFinishedGood.PartMaster.UnitPrice == 0) && (oSalesOrder.MakeDirect == false) )
                    {
                        oStockingOrdersWithoutPrice.Add(oSalesOrder);
                    }
                }
            }
            if (oStockingOrdersWithoutPrice.Count > 0)
            {
                // we do not need to repeat every release as we only need to indicate which line has the issue
                List<SOBacklog> oDistinctSalesOrderLines = new List<SOBacklog>();
                foreach (SOBacklog oTmp in oStockingOrdersWithoutPrice)
                {
                    SOBacklog oFound = oDistinctSalesOrderLines.FirstOrDefault(oItem => (oItem.OrderNum == oTmp.OrderNum) && (oItem.OrderLine == oTmp.OrderLine));
                    if (oFound == null)
                    {
                        // not in the list yet
                        oDistinctSalesOrderLines.Add(oTmp);
                    }
                }

                if (bFirstWorksheet == true)
                {
                    oSLSOMgrDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "Stock Order No Price On Part");
                    bFirstWorksheet = false;
                }
                else
                {
                    oSLSOMgrDocument.AddWorksheet("Stock Order No Price On Part");
                }

                //set up column headers
                int iNumOfColumns = 0;
                iNumOfColumns++;
                oSLSOMgrDocument.SetCellValue(1, iNumOfColumns, "SO Num");
                iNumOfColumns++;
                oSLSOMgrDocument.SetCellValue(1, iNumOfColumns, "Line Num");
                iNumOfColumns++;
                oSLSOMgrDocument.SetCellValue(1, iNumOfColumns, "Order Date");
                oSLSOMgrDocument.SetColumnWidth(iNumOfColumns, 15);
                iNumOfColumns++;
                oSLSOMgrDocument.SetCellValue(1, iNumOfColumns, "PartNum");
                oSLSOMgrDocument.SetColumnWidth(iNumOfColumns, 15);
                iNumOfColumns++;
                oSLSOMgrDocument.SetCellValue(1, iNumOfColumns, "Entry Person");
                oSLSOMgrDocument.SetColumnWidth(iNumOfColumns, 15);
                iNumOfColumns++;
                oSLSOMgrDocument.SetCellValue(1, iNumOfColumns, "Stocking Order But Price Not Set On Part");
                oSLSOMgrDocument.SetCellStyle(1, iNumOfColumns, oHighlightHeaderStyle);
                oSLSOMgrDocument.SetColumnWidth(iNumOfColumns, 60);

                int iNumOfRows = 1;
                foreach (SOBacklog oSalesOrder in oDistinctSalesOrderLines)
                {
                    iNumOfRows++;
                    oSLSOMgrDocument.SetCellValue(iNumOfRows, 1, oSalesOrder.OrderNum);
                    oSLSOMgrDocument.SetCellValue(iNumOfRows, 2, oSalesOrder.OrderLine);
                    oSLSOMgrDocument.SetCellValue(iNumOfRows, 3, oSalesOrder.OrderDate.ToShortDateString());
                    oSLSOMgrDocument.SetCellValue(iNumOfRows, 4, oSalesOrder.PartNum);
                    oSLSOMgrDocument.SetCellValue(iNumOfRows, 5, oSalesOrder.EntryPerson);
                }
            }

            // unfirm orders
            List<SOBacklog> oUnfirmOrders = m_oAllOpenOrders.Where(oItem => oItem.FirmRelease == false).ToList();
            if (oUnfirmOrders.Count != 0)
            {
                if (bFirstWorksheet == true)
                {
                    oSLSOMgrDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "Unfirm Orders");
                    bFirstWorksheet = false;
                }
                else
                {
                    oSLSOMgrDocument.AddWorksheet("Unfirm Orders");
                }

                //set up column headers
                int iNumOfColumns = 0;
                iNumOfColumns++;
                oSLSOMgrDocument.SetCellValue(1, iNumOfColumns, "SO Num");
                iNumOfColumns++;
                oSLSOMgrDocument.SetCellValue(1, iNumOfColumns, "Line Num");
                iNumOfColumns++;
                oSLSOMgrDocument.SetCellValue(1, iNumOfColumns, "Release Num");
                iNumOfColumns++;
                oSLSOMgrDocument.SetCellValue(1, iNumOfColumns, "Order Date");
                oSLSOMgrDocument.SetColumnWidth(iNumOfColumns, 15);
                iNumOfColumns++;
                oSLSOMgrDocument.SetCellValue(1, iNumOfColumns, "PartNum");
                oSLSOMgrDocument.SetColumnWidth(iNumOfColumns, 15);
                iNumOfColumns++;
                oSLSOMgrDocument.SetCellValue(1, iNumOfColumns, "Entry Person");
                oSLSOMgrDocument.SetColumnWidth(iNumOfColumns, 15);
                iNumOfColumns++;
                oSLSOMgrDocument.SetCellValue(1, iNumOfColumns, "Orders Are Unfirm");
                oSLSOMgrDocument.SetCellStyle(1, iNumOfColumns, oHighlightHeaderStyle);
                oSLSOMgrDocument.SetColumnWidth(iNumOfColumns, 30);

                int iNumOfRows = 1;
                foreach (SOBacklog oSalesOrder in oUnfirmOrders)
                {
                    iNumOfRows++;
                    oSLSOMgrDocument.SetCellValue(iNumOfRows, 1, oSalesOrder.OrderNum);
                    oSLSOMgrDocument.SetCellValue(iNumOfRows, 2, oSalesOrder.OrderLine);
                    oSLSOMgrDocument.SetCellValue(iNumOfRows, 3, oSalesOrder.OrderRelNum);
                    oSLSOMgrDocument.SetCellValue(iNumOfRows, 4, oSalesOrder.OrderDate.ToShortDateString());
                    oSLSOMgrDocument.SetCellValue(iNumOfRows, 5, oSalesOrder.PartNum);
                    oSLSOMgrDocument.SetCellValue(iNumOfRows, 6, oSalesOrder.EntryPerson);
                }
            }

            // orders that shipped but the order is still open
            List<SOBacklog> oOpenOrdersThatShipped = m_oAllOpenOrders.Where(oItem => (oItem.OurRequestedQty > 0) && (oItem.RemainingQty == 0)).ToList();
            if (oOpenOrdersThatShipped.Count != 0)
            {
                if (bFirstWorksheet == true)
                {
                    oSLSOMgrDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "Open Orer That Shipped");
                    bFirstWorksheet = false;
                }
                else
                {
                    oSLSOMgrDocument.AddWorksheet("Open Order That Shipped");
                }

                //set up column headers
                int iNumOfColumns = 0;
                iNumOfColumns++;
                oSLSOMgrDocument.SetCellValue(1, iNumOfColumns, "SO Num");
                iNumOfColumns++;
                oSLSOMgrDocument.SetCellValue(1, iNumOfColumns, "SO Line");
                iNumOfColumns++;
                oSLSOMgrDocument.SetCellValue(1, iNumOfColumns, "SO Rel");
                iNumOfColumns++;
                oSLSOMgrDocument.SetCellValue(1, iNumOfColumns, "Order Date");
                oSLSOMgrDocument.SetColumnWidth(iNumOfColumns, 15);
                iNumOfColumns++;
                oSLSOMgrDocument.SetCellValue(1, iNumOfColumns, "Entered By");
                oSLSOMgrDocument.SetColumnWidth(iNumOfColumns, 15);
                iNumOfColumns++;
                oSLSOMgrDocument.SetCellValue(1, iNumOfColumns, "Changed By");
                oSLSOMgrDocument.SetColumnWidth(iNumOfColumns, 15);
                iNumOfColumns++;
                oSLSOMgrDocument.SetCellValue(1, iNumOfColumns, "Changed Date");
                oSLSOMgrDocument.SetColumnWidth(iNumOfColumns, 15);
                iNumOfColumns++;
                oSLSOMgrDocument.SetCellValue(1, iNumOfColumns, "Order Shipped But Remains Open");
                oSLSOMgrDocument.SetCellStyle(1, iNumOfColumns, oHighlightHeaderStyle);
                oSLSOMgrDocument.SetColumnWidth(iNumOfColumns, 60);


                int iNumOfRows = 1;
                foreach (SOBacklog oSalesOrder in oOpenOrdersThatShipped)
                {
                    iNumOfRows++;
                    oSLSOMgrDocument.SetCellValue(iNumOfRows, 1, oSalesOrder.OrderNum);
                    oSLSOMgrDocument.SetCellValue(iNumOfRows, 2, oSalesOrder.OrderLine);
                    oSLSOMgrDocument.SetCellValue(iNumOfRows, 3, oSalesOrder.OrderRelNum);
                    oSLSOMgrDocument.SetCellValue(iNumOfRows, 4, oSalesOrder.OrderDate.ToShortDateString());
                    oSLSOMgrDocument.SetCellValue(iNumOfRows, 5, oSalesOrder.EntryPerson);
                    oSLSOMgrDocument.SetCellValue(iNumOfRows, 6, oSalesOrder.ChangedBy);
                    oSLSOMgrDocument.SetCellValue(iNumOfRows, 7, oSalesOrder.ChangeDate.ToShortDateString());
                }
            }

            // line order is open but no open releases
            List<SOBacklog> oLineNoOpenRelease = m_oAllOpenOrders.Where(oItem => oItem.OrderRelNum == 0).ToList();
            if (oLineNoOpenRelease.Count != 0)
            {
                // we do not need to repeat every release as we only need to indicate which line has the issue
                List<SOBacklog> oDistinctSalesOrderLines = new List<SOBacklog>();
                foreach (SOBacklog oTmp in oLineNoOpenRelease)
                {
                    SOBacklog oFound = oDistinctSalesOrderLines.FirstOrDefault(oItem => (oItem.OrderNum == oTmp.OrderNum) && (oItem.OrderLine == oTmp.OrderLine));
                    if (oFound == null)
                    {
                        // not in the list yet
                        oDistinctSalesOrderLines.Add(oTmp);
                    }
                }

                if (bFirstWorksheet == true)
                {
                    oSLSOMgrDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "Open Line No Open Release");
                    bFirstWorksheet = false;
                }
                else
                {
                    oSLSOMgrDocument.AddWorksheet("Open Line No Open Release");
                }

                //set up column headers
                int iNumOfColumns = 0;
                iNumOfColumns++;
                oSLSOMgrDocument.SetCellValue(1, iNumOfColumns, "SO Num");
                iNumOfColumns++;
                oSLSOMgrDocument.SetCellValue(1, iNumOfColumns, "Line Num");
                iNumOfColumns++;
                oSLSOMgrDocument.SetCellValue(1, iNumOfColumns, "Order Date");
                oSLSOMgrDocument.SetColumnWidth(iNumOfColumns, 15);
                iNumOfColumns++;
                oSLSOMgrDocument.SetCellValue(1, iNumOfColumns, "PartNum");
                oSLSOMgrDocument.SetColumnWidth(iNumOfColumns, 15);
                iNumOfColumns++;
                oSLSOMgrDocument.SetCellValue(1, iNumOfColumns, "Entry Person");
                oSLSOMgrDocument.SetColumnWidth(iNumOfColumns, 15);
                iNumOfColumns++;
                oSLSOMgrDocument.SetCellValue(1, iNumOfColumns, "Line Is Open But Release Is Closed");
                oSLSOMgrDocument.SetCellStyle(1, iNumOfColumns, oHighlightHeaderStyle);
                oSLSOMgrDocument.SetColumnWidth(iNumOfColumns, 60);

                int iNumOfRows = 1;
                foreach (SOBacklog oSalesOrder in oDistinctSalesOrderLines)
                {
                    iNumOfRows++;
                    oSLSOMgrDocument.SetCellValue(iNumOfRows, 1, oSalesOrder.OrderNum);
                    oSLSOMgrDocument.SetCellValue(iNumOfRows, 2, oSalesOrder.OrderLine);
                    oSLSOMgrDocument.SetCellValue(iNumOfRows, 3, oSalesOrder.OrderDate.ToShortDateString());
                    oSLSOMgrDocument.SetCellValue(iNumOfRows, 4, oSalesOrder.PartNum);
                    oSLSOMgrDocument.SetCellValue(iNumOfRows, 5, oSalesOrder.EntryPerson);
                }
            }
           
            // order is open but there are no open lines
            List<SOBacklog> oNoOpenLines = m_oAllOpenOrders.Where(oItem => oItem.OrderLine == 0).ToList();
            if (oNoOpenLines.Count != 0)
            {
                // we do not need to repeat every release as we only need to indicate which order has the issue
                List<SOBacklog> oDistinctSalesOrders = new List<SOBacklog>();
                foreach (SOBacklog oTmp in oNoOpenLines)
                {
                    SOBacklog oFound = oDistinctSalesOrders.FirstOrDefault(oItem => (oItem.OrderNum == oTmp.OrderNum));
                    if (oFound == null)
                    {
                        // not in the list yet
                        oDistinctSalesOrders.Add(oTmp);
                    }
                }

                if (bFirstWorksheet == true)
                {
                    oSLSOMgrDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "Open Order No Open Lines");
                    bFirstWorksheet = false;
                }
                else
                {
                    oSLSOMgrDocument.AddWorksheet("Open Order No Open Lines");
                }

                //set up column headers
                int iNumOfColumns = 0;
                iNumOfColumns++;
                oSLSOMgrDocument.SetCellValue(1, iNumOfColumns, "SO Num");
                iNumOfColumns++;
                oSLSOMgrDocument.SetCellValue(1, iNumOfColumns, "Order Date");
                oSLSOMgrDocument.SetColumnWidth(iNumOfColumns, 15);
                iNumOfColumns++;
                oSLSOMgrDocument.SetCellValue(1, iNumOfColumns, "Entry Person");
                oSLSOMgrDocument.SetColumnWidth(iNumOfColumns, 15);
                iNumOfColumns++;
                oSLSOMgrDocument.SetCellValue(1, iNumOfColumns, "Open Order But Lines Are Closed");
                oSLSOMgrDocument.SetCellStyle(1, iNumOfColumns, oHighlightHeaderStyle);
                oSLSOMgrDocument.SetColumnWidth(iNumOfColumns, 60);

                int iNumOfRows = 1;
                foreach (SOBacklog oSalesOrder in oDistinctSalesOrders)
                {
                    iNumOfRows++;
                    oSLSOMgrDocument.SetCellValue(iNumOfRows, 1, oSalesOrder.OrderNum);
                    oSLSOMgrDocument.SetCellValue(iNumOfRows, 2, oSalesOrder.OrderDate.ToShortDateString());
                    oSLSOMgrDocument.SetCellValue(iNumOfRows, 3, oSalesOrder.EntryPerson);
                }
            }

            // no rev number specified
            List<SOBacklog> oOrdersWithoutRevNumber = m_oAllOpenOrders.Where(oItem => string.IsNullOrEmpty(oItem.PartRev) == true).ToList();
            if (oOrdersWithoutRevNumber.Count != 0)
            {
                // we do not need to repeat every release as we only need to indicate which line has the issue
                List<SOBacklog> oDistinctSalesOrderLines = new List<SOBacklog>();
                foreach (SOBacklog oTmp in oOrdersWithoutRevNumber)
                {
                    SOBacklog oFound = oDistinctSalesOrderLines.FirstOrDefault(oItem => (oItem.OrderNum == oTmp.OrderNum) && (oItem.OrderLine == oTmp.OrderLine));
                    if (oFound == null)
                    {
                        // not in the list yet
                        oDistinctSalesOrderLines.Add(oTmp);
                    }
                }

                if (bFirstWorksheet == true)
                {
                    oSLSOMgrDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "No Rev Number");
                    bFirstWorksheet = false;
                }
                else
                {
                    oSLSOMgrDocument.AddWorksheet("No Rev Number");
                }

                //set up column headers
                int iNumOfColumns = 0;
                iNumOfColumns++;
                oSLSOMgrDocument.SetCellValue(1, iNumOfColumns, "SO Num");
                iNumOfColumns++;
                oSLSOMgrDocument.SetCellValue(1, iNumOfColumns, "Line Num");
                iNumOfColumns++;
                oSLSOMgrDocument.SetCellValue(1, iNumOfColumns, "Order Date");
                oSLSOMgrDocument.SetColumnWidth(iNumOfColumns, 15);
                iNumOfColumns++;
                oSLSOMgrDocument.SetCellValue(1, iNumOfColumns, "PartNum");
                oSLSOMgrDocument.SetColumnWidth(iNumOfColumns, 15);
                iNumOfColumns++;
                oSLSOMgrDocument.SetCellValue(1, iNumOfColumns, "Entry Person");
                oSLSOMgrDocument.SetColumnWidth(iNumOfColumns, 15);
                iNumOfColumns++;
                oSLSOMgrDocument.SetCellValue(1, iNumOfColumns, "Part Does Not Have Revision Specified");
                oSLSOMgrDocument.SetCellStyle(1, iNumOfColumns, oHighlightHeaderStyle);
                oSLSOMgrDocument.SetColumnWidth(iNumOfColumns, 60);

                int iNumOfRows = 1;
                foreach (SOBacklog oSalesOrder in oDistinctSalesOrderLines)
                {
                    iNumOfRows++;
                    oSLSOMgrDocument.SetCellValue(iNumOfRows, 1, oSalesOrder.OrderNum);
                    oSLSOMgrDocument.SetCellValue(iNumOfRows, 2, oSalesOrder.OrderLine);
                    oSLSOMgrDocument.SetCellValue(iNumOfRows, 3, oSalesOrder.OrderDate.ToShortDateString());
                    oSLSOMgrDocument.SetCellValue(iNumOfRows, 4, oSalesOrder.PartNum);
                    oSLSOMgrDocument.SetCellValue(iNumOfRows, 5, oSalesOrder.EntryPerson);
                }
            }
          
            // the rev is not approved -- already created this list above
            if (oOrdersWithUnapprovedRevs.Count != 0)
            {
                // we do not need to repeat every release as we only need to indicate which line has the issue
                List<SOBacklog> oDistinctSalesOrderLines = new List<SOBacklog>();
                foreach (SOBacklog oTmp in oOrdersWithUnapprovedRevs)
                {
                    SOBacklog oFound = oDistinctSalesOrderLines.FirstOrDefault(oItem => (oItem.OrderNum == oTmp.OrderNum) && (oItem.OrderLine == oTmp.OrderLine));
                    if (oFound == null)
                    {
                        // not in the list yet
                        oDistinctSalesOrderLines.Add(oTmp);
                    }
                }

                if (bFirstWorksheet == true)
                {
                    oSLSOMgrDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "No Appr Rev");
                    bFirstWorksheet = false;
                }
                else
                {
                    oSLSOMgrDocument.AddWorksheet("No Appr Rev");
                }

                //set up column headers
                int iNumOfColumns = 0;
                iNumOfColumns++;
                oSLSOMgrDocument.SetCellValue(1, iNumOfColumns, "SO Num");
                iNumOfColumns++;
                oSLSOMgrDocument.SetCellValue(1, iNumOfColumns, "Line Num");
                iNumOfColumns++;
                oSLSOMgrDocument.SetCellValue(1, iNumOfColumns, "Order Date");
                oSLSOMgrDocument.SetColumnWidth(iNumOfColumns, 15);
                iNumOfColumns++;
                oSLSOMgrDocument.SetCellValue(1, iNumOfColumns, "PartNum");
                oSLSOMgrDocument.SetColumnWidth(iNumOfColumns, 15);
                iNumOfColumns++;
                oSLSOMgrDocument.SetCellValue(1, iNumOfColumns, "Rev Num");
                oSLSOMgrDocument.SetColumnWidth(iNumOfColumns, 15);
                iNumOfColumns++;
                oSLSOMgrDocument.SetCellValue(1, iNumOfColumns, "Entry Person");
                oSLSOMgrDocument.SetColumnWidth(iNumOfColumns, 15);
                iNumOfColumns++;
                oSLSOMgrDocument.SetCellValue(1, iNumOfColumns, "Part Does Not Have Approved Revision");
                oSLSOMgrDocument.SetCellStyle(1, iNumOfColumns, oHighlightHeaderStyle);
                oSLSOMgrDocument.SetColumnWidth(iNumOfColumns, 60);

                int iNumOfRows = 1;
                foreach (SOBacklog oSalesOrder in oDistinctSalesOrderLines)
                {
                    iNumOfRows++;
                    oSLSOMgrDocument.SetCellValue(iNumOfRows, 1, oSalesOrder.OrderNum);
                    oSLSOMgrDocument.SetCellValue(iNumOfRows, 2, oSalesOrder.OrderLine);
                    oSLSOMgrDocument.SetCellValue(iNumOfRows, 3, oSalesOrder.OrderDate.ToShortDateString());
                    oSLSOMgrDocument.SetCellValue(iNumOfRows, 4, oSalesOrder.PartNum);
                    oSLSOMgrDocument.SetCellValue(iNumOfRows, 5, oSalesOrder.PartRev);
                    oSLSOMgrDocument.SetCellValue(iNumOfRows, 6, oSalesOrder.EntryPerson);
                }
            }

            #region IGNORING THESE
            // not readt to fulfill
            //List<SOBacklog> oOrderNotReadyToFulfil = m_oAllOpenOrders.Where(oItem => oItem.ReadyToFulfill == false).ToList();
            //if (oOrderNotReadyToFulfil.Count != 0)
            //{
            //    // we do not need to repeat every release as we only need to indicate which order has the issue
            //    List<SOBacklog> oDistinctSalesOrders = new List<SOBacklog>();
            //    foreach (SOBacklog oTmp in oOrderNotReadyToFulfil)
            //    {
            //        SOBacklog oFound = oDistinctSalesOrders.FirstOrDefault(oItem => (oItem.OrderNum == oTmp.OrderNum));
            //        if (oFound == null)
            //        {
            //            // not in the list yet
            //            oDistinctSalesOrders.Add(oTmp);
            //        }
            //    }

            //    if (bFirstWorksheet == true)
            //    {
            //        oSLSOMgrDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "Not Ready To Fulfill");
            //        bFirstWorksheet = false;
            //    }
            //    else
            //    {
            //        oSLSOMgrDocument.AddWorksheet("Not Ready To Fulfill");
            //    }

            //    //set up column headers
            //    int iNumOfColumns = 0;
            //    iNumOfColumns++;
            //    oSLSOMgrDocument.SetCellValue(1, iNumOfColumns, "SO Num");
            //    iNumOfColumns++;
            //    oSLSOMgrDocument.SetCellValue(1, iNumOfColumns, "Order Date");
            //    oSLSOMgrDocument.SetColumnWidth(iNumOfColumns, 15);
            //    iNumOfColumns++;
            //    oSLSOMgrDocument.SetCellValue(1, iNumOfColumns, "Entry Person");
            //    oSLSOMgrDocument.SetColumnWidth(iNumOfColumns, 15);
            //    iNumOfColumns++;
            //    oSLSOMgrDocument.SetCellValue(1, iNumOfColumns, "Order Not Ready To Fulfill");
            //    oSLSOMgrDocument.SetCellStyle(1, iNumOfColumns, oHighlightHeaderStyle);
            //    oSLSOMgrDocument.SetColumnWidth(iNumOfColumns, 60);

            //    int iNumOfRows = 1;
            //    foreach (SOBacklog oSalesOrder in oDistinctSalesOrders)
            //    {
            //        iNumOfRows++;
            //        oSLSOMgrDocument.SetCellValue(iNumOfRows, 1, oSalesOrder.OrderNum);
            //        oSLSOMgrDocument.SetCellValue(iNumOfRows, 2, oSalesOrder.OrderDate.ToShortDateString());
            //        oSLSOMgrDocument.SetCellValue(iNumOfRows, 3, oSalesOrder.EntryPerson);
            //    }
            //}
            #endregion

            // order header is canceled and we know the header is open or the BAQ would not return the item
            List<SOBacklog> oCanceledOrdersNotClosed = m_oAllOpenOrders.Where(oItem => (oItem.VoidOrder == true)).ToList();
            if (oCanceledOrdersNotClosed.Count != 0)
            {
                // we do not need to repeat every release as we only need to indicate which order has the issue
                List<SOBacklog> oDistinctSalesOrders = new List<SOBacklog>();
                foreach (SOBacklog oTmp in oCanceledOrdersNotClosed)
                {
                    SOBacklog oFound = oDistinctSalesOrders.FirstOrDefault(oItem => (oItem.OrderNum == oTmp.OrderNum));
                    if (oFound == null)
                    {
                        // not in the list yet
                        oDistinctSalesOrders.Add(oTmp);
                    }
                }

                if (bFirstWorksheet == true)
                {
                    oSLSOMgrDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "Canceled Orders Open");
                    bFirstWorksheet = false;
                }
                else
                {
                    oSLSOMgrDocument.AddWorksheet("Canceled Orders Open");
                }

                //set up column headers
                int iNumOfColumns = 0;
                iNumOfColumns++;
                oSLSOMgrDocument.SetCellValue(1, iNumOfColumns, "SO Num");
                iNumOfColumns++;
                oSLSOMgrDocument.SetCellValue(1, iNumOfColumns, "Order Date");
                oSLSOMgrDocument.SetColumnWidth(iNumOfColumns, 15);
                iNumOfColumns++;
                oSLSOMgrDocument.SetCellValue(1, iNumOfColumns, "Entry Person");
                oSLSOMgrDocument.SetColumnWidth(iNumOfColumns, 15);
                iNumOfColumns++;
                oSLSOMgrDocument.SetCellValue(1, iNumOfColumns, "Order Canceled But Is Not Closed");
                oSLSOMgrDocument.SetCellStyle(1, iNumOfColumns, oHighlightHeaderStyle);
                oSLSOMgrDocument.SetColumnWidth(iNumOfColumns, 60);

                int iNumOfRows = 1;
                foreach (SOBacklog oSalesOrder in oDistinctSalesOrders)
                {
                    iNumOfRows++;
                    oSLSOMgrDocument.SetCellValue(iNumOfRows, 1, oSalesOrder.OrderNum);
                    oSLSOMgrDocument.SetCellValue(iNumOfRows, 2, oSalesOrder.OrderDate.ToShortDateString());
                    oSLSOMgrDocument.SetCellValue(iNumOfRows, 3, oSalesOrder.EntryPerson);
                }
            }

            // order line is canceled and we know the line is open or the BAQ would not return the item
            List<SOBacklog> oCanceledLinesNotClosed = m_oAllOpenOrders.Where(oItem => (oItem.VoidLine == true)).ToList();
            if (oCanceledLinesNotClosed.Count != 0)
            {
                // we do not need to repeat every release as we only need to indicate which line has the issue
                List<SOBacklog> oDistinctSalesOrderLines = new List<SOBacklog>();
                foreach (SOBacklog oTmp in oCanceledLinesNotClosed)
                {
                    SOBacklog oFound = oDistinctSalesOrderLines.FirstOrDefault(oItem => (oItem.OrderNum == oTmp.OrderNum) && (oItem.OrderLine == oTmp.OrderLine));
                    if (oFound == null)
                    {
                        // not in the list yet
                        oDistinctSalesOrderLines.Add(oTmp);
                    }
                }

                if (bFirstWorksheet == true)
                {
                    oSLSOMgrDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "Canceled Line Open");
                    bFirstWorksheet = false;
                }
                else
                {
                    oSLSOMgrDocument.AddWorksheet("Canceled Line Open");
                }

                //set up column headers
                int iNumOfColumns = 0;
                iNumOfColumns++;
                oSLSOMgrDocument.SetCellValue(1, iNumOfColumns, "SO Num");
                iNumOfColumns++;
                oSLSOMgrDocument.SetCellValue(1, iNumOfColumns, "Line Num");
                iNumOfColumns++;
                oSLSOMgrDocument.SetCellValue(1, iNumOfColumns, "Order Date");
                oSLSOMgrDocument.SetColumnWidth(iNumOfColumns, 15);
                iNumOfColumns++;
                oSLSOMgrDocument.SetCellValue(1, iNumOfColumns, "PartNum");
                oSLSOMgrDocument.SetColumnWidth(iNumOfColumns, 15);
                iNumOfColumns++;
                oSLSOMgrDocument.SetCellValue(1, iNumOfColumns, "Entry Person");
                oSLSOMgrDocument.SetColumnWidth(iNumOfColumns, 15);
                iNumOfColumns++;
                oSLSOMgrDocument.SetCellValue(1, iNumOfColumns, "Line Is Canceled But Remains Open");
                oSLSOMgrDocument.SetCellStyle(1, iNumOfColumns, oHighlightHeaderStyle);
                oSLSOMgrDocument.SetColumnWidth(iNumOfColumns, 60);

                int iNumOfRows = 1;
                foreach (SOBacklog oSalesOrder in oDistinctSalesOrderLines)
                {
                    iNumOfRows++;
                    oSLSOMgrDocument.SetCellValue(iNumOfRows, 1, oSalesOrder.OrderNum);
                    oSLSOMgrDocument.SetCellValue(iNumOfRows, 2, oSalesOrder.OrderLine);
                    oSLSOMgrDocument.SetCellValue(iNumOfRows, 3, oSalesOrder.OrderDate.ToShortDateString());
                    oSLSOMgrDocument.SetCellValue(iNumOfRows, 4, oSalesOrder.PartNum);
                    oSLSOMgrDocument.SetCellValue(iNumOfRows, 5, oSalesOrder.EntryPerson);
                }
            }

            // the req date is set before the order date -- cant meet this deadline
            List<SOBacklog> oReqDateBeforeOrderDate = m_oAllOpenOrders.Where(oItem => oItem.RequiredByDate < oItem.OrderDate).ToList();
            if (oReqDateBeforeOrderDate.Count != 0)
            {
                if (bFirstWorksheet == true)
                {
                    oSLSOMgrDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "Req Date Before Order Date");
                    bFirstWorksheet = false;
                }
                else
                {
                    oSLSOMgrDocument.AddWorksheet("Req Date Before Order Date");
                }

                //set up column headers
                int iNumOfColumns = 0;
                iNumOfColumns++;
                oSLSOMgrDocument.SetCellValue(1, iNumOfColumns, "SO Num");
                iNumOfColumns++;
                oSLSOMgrDocument.SetCellValue(1, iNumOfColumns, "Line Num");
                iNumOfColumns++;
                oSLSOMgrDocument.SetCellValue(1, iNumOfColumns, "Release Num");
                iNumOfColumns++;
                oSLSOMgrDocument.SetCellValue(1, iNumOfColumns, "Order Date");
                oSLSOMgrDocument.SetColumnWidth(iNumOfColumns, 15);
                iNumOfColumns++;
                oSLSOMgrDocument.SetCellValue(1, iNumOfColumns, "Req Date");
                oSLSOMgrDocument.SetColumnWidth(iNumOfColumns, 15);
                iNumOfColumns++;
                oSLSOMgrDocument.SetCellValue(1, iNumOfColumns, "PartNum");
                oSLSOMgrDocument.SetColumnWidth(iNumOfColumns, 15);
                iNumOfColumns++;
                oSLSOMgrDocument.SetCellValue(1, iNumOfColumns, "Entry Person");
                oSLSOMgrDocument.SetColumnWidth(iNumOfColumns, 15);
                iNumOfColumns++;
                oSLSOMgrDocument.SetCellValue(1, iNumOfColumns, "Release Required Date Is Earlier Than Order Date");
                oSLSOMgrDocument.SetCellStyle(1, iNumOfColumns, oHighlightHeaderStyle);
                oSLSOMgrDocument.SetColumnWidth(iNumOfColumns, 60);

                int iNumOfRows = 1;
                foreach (SOBacklog oSalesOrder in oReqDateBeforeOrderDate)
                {
                    iNumOfRows++;
                    oSLSOMgrDocument.SetCellValue(iNumOfRows, 1, oSalesOrder.OrderNum);
                    oSLSOMgrDocument.SetCellValue(iNumOfRows, 2, oSalesOrder.OrderLine);
                    oSLSOMgrDocument.SetCellValue(iNumOfRows, 3, oSalesOrder.OrderRelNum);
                    oSLSOMgrDocument.SetCellValue(iNumOfRows, 4, oSalesOrder.OrderDate.ToShortDateString());
                    oSLSOMgrDocument.SetCellValue(iNumOfRows, 5, oSalesOrder.RequiredByDate.ToShortDateString());
                    oSLSOMgrDocument.SetCellValue(iNumOfRows, 6, oSalesOrder.PartNum);
                    oSLSOMgrDocument.SetCellValue(iNumOfRows, 7, oSalesOrder.EntryPerson);
                }
            }

            // did not indicate when this order must ship -- cant schedule
            List<SOBacklog> oOrdersWithNoReqDateSet = m_oAllOpenOrders.Where(oItem => oItem.RequiredByDate == DateTime.MinValue).ToList();
            if (oOrdersWithNoReqDateSet.Count != 0)
            {
                if (bFirstWorksheet == true)
                {
                    oSLSOMgrDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "No Req Date");
                    bFirstWorksheet = false;
                }
                else
                {
                    oSLSOMgrDocument.AddWorksheet("No Req Date");
                }

                //set up column headers
                int iNumOfColumns = 0;
                iNumOfColumns++;
                oSLSOMgrDocument.SetCellValue(1, iNumOfColumns, "SO Num");
                iNumOfColumns++;
                oSLSOMgrDocument.SetCellValue(1, iNumOfColumns, "Line Num");
                iNumOfColumns++;
                oSLSOMgrDocument.SetCellValue(1, iNumOfColumns, "Release Num");
                iNumOfColumns++;
                oSLSOMgrDocument.SetCellValue(1, iNumOfColumns, "Order Date");
                oSLSOMgrDocument.SetColumnWidth(iNumOfColumns, 15);
                iNumOfColumns++;
                oSLSOMgrDocument.SetCellValue(1, iNumOfColumns, "PartNum");
                oSLSOMgrDocument.SetColumnWidth(iNumOfColumns, 15);
                iNumOfColumns++;
                oSLSOMgrDocument.SetCellValue(1, iNumOfColumns, "Entry Person");
                oSLSOMgrDocument.SetColumnWidth(iNumOfColumns, 15);
                iNumOfColumns++;
                oSLSOMgrDocument.SetCellValue(1, iNumOfColumns, "Release Required Date Is Not Set");
                oSLSOMgrDocument.SetCellStyle(1, iNumOfColumns, oHighlightHeaderStyle);
                oSLSOMgrDocument.SetColumnWidth(iNumOfColumns, 60);

                int iNumOfRows = 1;
                foreach (SOBacklog oSalesOrder in oOrdersWithNoReqDateSet)
                {
                    iNumOfRows++;
                    oSLSOMgrDocument.SetCellValue(iNumOfRows, 1, oSalesOrder.OrderNum);
                    oSLSOMgrDocument.SetCellValue(iNumOfRows, 2, oSalesOrder.OrderLine);
                    oSLSOMgrDocument.SetCellValue(iNumOfRows, 3, oSalesOrder.OrderRelNum);
                    oSLSOMgrDocument.SetCellValue(iNumOfRows, 4, oSalesOrder.OrderDate.ToShortDateString());
                    oSLSOMgrDocument.SetCellValue(iNumOfRows, 5, oSalesOrder.PartNum);
                    oSLSOMgrDocument.SetCellValue(iNumOfRows, 6, oSalesOrder.EntryPerson);
                }
            }
     
            // send the email to the Customer Service team
            if (bFirstWorksheet == false)
            {
                // we create a worksheet so we have some data in the spreadsheet so we email the spreadsheet
                oSLSOMgrDocument.SaveAs(sDestinationFileName);
                List<string> oAttachments = new List<string>();
                oAttachments.Add(sDestinationFileName);
                HSEmailHelper.SendEmail(oToAddresses, sCompany + " Order Validation", sCompany + " Order Validation", oAttachments);
            }
            #endregion

            #region Procurement Issues
            //
            // Procurement Issues
            //
            bFirstWorksheet = true;
            sDestinationFileName = sTmpFileDirectory + sCompany + "-SO PO Validation-" + dtToday.Month.ToString() + "-" + dtToday.Day.ToString() + "-" + dtToday.Year.ToString() + ".xlsx";
            SLDocument oSLLogMgrDocument = new SLDocument();
            oToAddresses = HSUser.GetEmailsForUsersInGroup(HSUser.REPORT_ON_PROCUREMENT_ISSUES);
            oToAddresses.Add(HSUser.SPINCRAFT_ROOT_USER);
            List<SOBacklog> oBuyDirectMissingPO = m_oAllOpenOrders.Where(oItem => (oItem.BuyToOrder == true) && (oItem.PONum == 0)).ToList();

            if (oBuyDirectMissingPO.Count != 0)
            {
                if (bFirstWorksheet == true)
                {
                    oSLLogMgrDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "Missing PO");
                    bFirstWorksheet = false;
                }
                else
                {
                    oSLLogMgrDocument.AddWorksheet("Missing PO");
                }

                //set up column headers
                int iNumOfColumns = 0;
                iNumOfColumns++;
                oSLLogMgrDocument.SetCellValue(1, iNumOfColumns, "SO Num");
                iNumOfColumns++;
                oSLLogMgrDocument.SetCellValue(1, iNumOfColumns, "Line Num");
                iNumOfColumns++;
                oSLLogMgrDocument.SetCellValue(1, iNumOfColumns, "Release Num");
                iNumOfColumns++;
                oSLLogMgrDocument.SetCellValue(1, iNumOfColumns, "Order Date");
                oSLLogMgrDocument.SetColumnWidth(iNumOfColumns, 15);
                iNumOfColumns++;
                oSLLogMgrDocument.SetCellValue(1, iNumOfColumns, "PartNum");
                oSLLogMgrDocument.SetColumnWidth(iNumOfColumns, 15);
                iNumOfColumns++;
                oSLLogMgrDocument.SetCellValue(1, iNumOfColumns, "Entry Person");
                oSLLogMgrDocument.SetColumnWidth(iNumOfColumns, 15);
                iNumOfColumns++;
                oSLLogMgrDocument.SetCellValue(1, iNumOfColumns, "Release Requires A PO But None Is Linked");
                oSLLogMgrDocument.SetCellStyle(1, iNumOfColumns, oHighlightHeaderStyle);
                oSLLogMgrDocument.SetColumnWidth(iNumOfColumns, 60);

                int iNumOfRows = 1;
                foreach (SOBacklog oSalesOrder in oBuyDirectMissingPO)
                {
                    iNumOfRows++;
                    oSLLogMgrDocument.SetCellValue(iNumOfRows, 1, oSalesOrder.OrderNum);
                    oSLLogMgrDocument.SetCellValue(iNumOfRows, 2, oSalesOrder.OrderLine);
                    oSLLogMgrDocument.SetCellValue(iNumOfRows, 3, oSalesOrder.OrderRelNum);
                    oSLLogMgrDocument.SetCellValue(iNumOfRows, 4, oSalesOrder.OrderDate.ToShortDateString());
                    oSLLogMgrDocument.SetCellValue(iNumOfRows, 5, oSalesOrder.PartNum);
                    oSLLogMgrDocument.SetCellValue(iNumOfRows, 6, oSalesOrder.EntryPerson);
                }
            }
            
            // send the email to the Logistics manager
            if (bFirstWorksheet == false)
            {
                // we create a worksheet so we have some data in the spreadsheet so we email the spreadsheet
                oSLLogMgrDocument.SaveAs(sDestinationFileName);
                List<string> oAttachments = new List<string>();
                oAttachments.Add(sDestinationFileName);
                HSEmailHelper.SendEmail(oToAddresses, sCompany + " Sales Order Validation", sCompany + " Sales Order Validation", oAttachments);
            }
            #endregion
        }

        #endregion

        #region Properties
        #endregion

        #region Data Members
        private List<SOBacklog> m_oAllOpenOrders = new List<SOBacklog>();
        private HSValidateParts m_oValidateParts;
        private BOMSupport m_oBOMSupport;
        #endregion
    }
}