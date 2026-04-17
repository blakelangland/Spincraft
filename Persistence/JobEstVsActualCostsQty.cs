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
    public class JobEstVsActualCostsQty
    {
        #region Constructors

        public JobEstVsActualCostsQty()
        {
        }

        public JobEstVsActualCostsQty(DataRow oRow)
        {
            if (oRow["JobHead_Company"] != DBNull.Value)
            {
                m_sCompany = (string)oRow["JobHead_Company"];
            }
            if (oRow["JobHead_Plant"] != DBNull.Value)
            {
                m_sPlant = (string)oRow["JobHead_Plant"];
            }
            if (oRow["JobHead_JobNum"] != DBNull.Value)
            {
                m_sJobNum = (string)oRow["JobHead_JobNum"];
            }
            if (oRow["JobHead_StartDate"] != DBNull.Value)
            {
                m_dtStartDate = (DateTime)oRow["JobHead_StartDate"];
            }
            if (oRow["Calculated_LastJobClockInDate"] != DBNull.Value)
            {
                m_dtLastClockInDate = (DateTime)oRow["Calculated_LastJobClockInDate"];
            }
            if (oRow["JobHead_DueDate"] != DBNull.Value)
            {
                m_dtDueDate = (DateTime)oRow["JobHead_DueDate"];
            }
            if (oRow["JobHead_ReqDueDate"] != DBNull.Value)
            {
                m_dtRequiredDate = (DateTime)oRow["JobHead_ReqDueDate"];
            }
            if (oRow["JobHead_ProdQty"] != DBNull.Value)
            {
                m_dProdQty = (decimal)oRow["JobHead_ProdQty"];
            }
            if (oRow["JobHead_PartNum"] != DBNull.Value)
            {
                m_sPartNum = (string)oRow["JobHead_PartNum"];
            }
            if (oRow["JobHead_RevisionNum"] != DBNull.Value)
            {
                m_sPartRevNum = (string)oRow["JobHead_RevisionNum"];
            }
            if (oRow["Calculated_PercentageComplete"] != DBNull.Value)
            {
                m_dPercentageComplete = (decimal)oRow["Calculated_PercentageComplete"];
            }
            if (oRow["Calculated_EstTotalCost"] != DBNull.Value)
            {
                m_dEstTotalCost = (decimal)oRow["Calculated_EstTotalCost"];
            }
            if (oRow["Calculated_PercentEstTotalCost"] != DBNull.Value)
            {
                m_dPercentEstTotalCost = (decimal)oRow["Calculated_PercentEstTotalCost"];
            }
            if (oRow["Calculated_ActTotalCost"] != DBNull.Value)
            {
                m_dActTotalCost = (decimal)oRow["Calculated_ActTotalCost"];
            }
            if (oRow["Calculated_PercentHoursComplete"] != DBNull.Value)
            {
                m_dPercentHoursComplete = (decimal)oRow["Calculated_PercentHoursComplete"];
            }
            if (oRow["Calculated_EstOprHours"] != DBNull.Value)
            {
                m_dEstOperationalHours = (decimal)oRow["Calculated_EstOprHours"];
            }
            if (oRow["Calculated_PercentEstOprHours"] != DBNull.Value)
            {
                m_dPercentEstOperationalHours = (decimal)oRow["Calculated_PercentEstOprHours"];
            }
            if (oRow["Calculated_ActOprHours"] != DBNull.Value)
            {
                m_dActualOperationalHours = (decimal)oRow["Calculated_ActOprHours"];
            }
            if (oRow["Calculated_PercentageLaborComplete"] != DBNull.Value)
            {
                m_dPercentLaborComplete = (decimal)oRow["Calculated_PercentageLaborComplete"];
            }
            if (oRow["Calculated_EstLbrCost"] != DBNull.Value)
            {
                m_dEstLaborCost = (decimal)oRow["Calculated_EstLbrCost"];
            }
            if (oRow["Calculated_PercentEstLbrCost"] != DBNull.Value)
            {
                m_dPercentEstLaborCost = (decimal)oRow["Calculated_PercentEstLbrCost"];
            }
            if (oRow["Calculated_ActLbrCost"] != DBNull.Value)
            {
                m_dActualLaborCost = (decimal)oRow["Calculated_ActLbrCost"];
            }
            if (oRow["Calculated_EstBurCost"] != DBNull.Value)
            {
                m_dEstBurdenCost = (decimal)oRow["Calculated_EstBurCost"];
            }
            if (oRow["Calculated_PercentEstBurCost"] != DBNull.Value)
            {
                m_dPercentEstBurdenCost = (decimal)oRow["Calculated_PercentEstBurCost"];
            }
            if (oRow["Calculated_ActBurCost"] != DBNull.Value)
            {
                m_dActualBurdenCost = (decimal)oRow["Calculated_ActBurCost"];
            }
            if (oRow["Calculated_PercentageSubcontractComplete"] != DBNull.Value)
            {
                m_dPercentSubcontractComplete = (decimal)oRow["Calculated_PercentageSubcontractComplete"];
            }
            if (oRow["Calculated_EstSubCost"] != DBNull.Value)
            {
                m_dEstSubcontractCost = (decimal)oRow["Calculated_EstSubCost"];
            }
            if (oRow["Calculated_PercentEstSubCost"] != DBNull.Value)
            {
                m_dPercentEstSubcontractCost = (decimal)oRow["Calculated_PercentEstSubCost"];
            }
            if (oRow["Calculated_ActSubCost"] != DBNull.Value)
            {
                m_dActualSubcontractCost = (decimal)oRow["Calculated_ActSubCost"];
            }
            if (oRow["Calculated_PercentageMaterialsComplete"] != DBNull.Value)
            {
                m_dPercentMaterialComplete = (decimal)oRow["Calculated_PercentageMaterialsComplete"];
            }
            if (oRow["Calculated_EstMtlCost"] != DBNull.Value)
            {
                m_dEstMaterialCost = (decimal)oRow["Calculated_EstMtlCost"];
            }
            if (oRow["Calculated_PercentEstMtlCost"] != DBNull.Value)
            {
                m_dPercentEstMaterialCost = (decimal)oRow["Calculated_PercentEstMtlCost"];
            }
            if (oRow["Calculated_ActMtlCost"] != DBNull.Value)
            {
                m_dActualMaterialCost = (decimal)oRow["Calculated_ActMtlCost"];
            }
            if (oRow["Calculated_EstMtlBurCost"] != DBNull.Value)
            {
                m_dEstMaterialBurdenCost = (decimal)oRow["Calculated_EstMtlBurCost"];
            }
            if (oRow["Calculated_PercentEstMtlBurCost"] != DBNull.Value)
            {
                m_dPercentEstMaterialBurdenCost = (decimal)oRow["Calculated_PercentEstMtlBurCost"];
            }
            if (oRow["Calculated_ActMtlBurCost"] != DBNull.Value)
            {
                m_dActualMaterialBurdenCost = (decimal)oRow["Calculated_ActMtlBurCost"];
            }
            if (oRow["JobHead_JobClosed"] != DBNull.Value)
            {
                m_bJobClosed = (bool)oRow["JobHead_JobClosed"];
            }
            if (oRow["JobHead_JobComplete"] != DBNull.Value)
            {
                m_bJobComplete = (bool)oRow["JobHead_JobComplete"];
            }
            if (oRow["Customer_CustNum"] != DBNull.Value)
            {
                m_iCustNum = (int)oRow["Customer_CustNum"];
            }
            if (oRow["Customer_CustID"] != DBNull.Value)
            {
                m_sCustID = (string)oRow["Customer_CustID"];
            }
            if (oRow["Customer_Name"] != DBNull.Value)
            {
                m_sCustomerName = (string)oRow["Customer_Name"];
            }
            if (oRow["OrderHed_MarketSegment_c"] != DBNull.Value)
            {
                m_sMarketSegment = (string)oRow["OrderHed_MarketSegment_c"];
            }
            if (oRow["Calculated_MarketSegmentDesc"] != DBNull.Value)
            {
                m_sMarketSegmentDescription = (string)oRow["Calculated_MarketSegmentDesc"];
            }
            if (oRow["OrderRel_OrderNum"] != DBNull.Value)
            {
                m_iOrderNum = (int)oRow["OrderRel_OrderNum"];
            }
            if (oRow["OrderRel_OrderLine"] != DBNull.Value)
            {
                m_iOrderLine = (int)oRow["OrderRel_OrderLine"];
            }
            if (oRow["OrderRel_OrderRelNum"] != DBNull.Value)
            {
                m_iOrderRelNum = (int)oRow["OrderRel_OrderRelNum"];
            }
            if (oRow["OrderDtl_ProductPortfolio_c"] != DBNull.Value)
            {
                m_sProductPortfolio = (string)oRow["OrderDtl_ProductPortfolio_c"];
            }
            if (oRow["Calculated_ProductPortfolioDesc"] != DBNull.Value)
            {
                m_sProductPortfolioDescription = (string)oRow["Calculated_ProductPortfolioDesc"];
            }
            if (oRow["OrderRel_OurReqQty"] != DBNull.Value)
            {
                m_dOurReqQty = (decimal)oRow["OrderRel_OurReqQty"];
            }
            if (oRow["OrderDtl_UnitPrice"] != DBNull.Value)
            {
                m_dUnitPrice = (decimal)oRow["OrderDtl_UnitPrice"];
            }
            if (oRow["Calculated_ProjectedTotalCost"] != DBNull.Value)
            {
                m_dProjectedTotalCost = (decimal)oRow["Calculated_ProjectedTotalCost"];
            }
            if (oRow["Calculated_MarginToDate"] != DBNull.Value)
            {
                m_dMarginToDate = (decimal)oRow["Calculated_MarginToDate"];
            }
            if (oRow["Calculated_SalesDemandLink"] != DBNull.Value)
            {
                m_bSalesDemandLink = (bool)oRow["Calculated_SalesDemandLink"];
            }
            if (oRow["Calculated_PartUnitPrice"] != DBNull.Value)
            {
                m_dPartUnitPrice = (decimal)oRow["Calculated_PartUnitPrice"];
            }
            if (oRow["Calculated_PartOnFly"] != DBNull.Value)
            {
                m_bPartOnTheFly = (bool)oRow["Calculated_PartOnFly"];
            }
            if (oRow["JobProd_ShippedQty"] != DBNull.Value)
            {
                m_dShippedQty = (decimal)oRow["JobProd_ShippedQty"];
            }
            if (oRow["JobProd_ReceivedQty"] != DBNull.Value)
            {
                m_dReceivedQty = (decimal)oRow["JobProd_ReceivedQty"];
            }
            if (oRow["OrderRel_OpenRelease"] != DBNull.Value)
            {
                m_bOpenRelease = (bool)oRow["OrderRel_OpenRelease"];
            }
            if (oRow["JobProd_WarehouseCode"] != DBNull.Value)
            {
                m_sWarehouseCode = (string)oRow["JobProd_WarehouseCode"];
            }

            // compute the estimated remaining cost
            m_dEstimatedRemainingCost = 0M;
            

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

        public string Plant
        {
            get { return m_sPlant; }
            set { m_sPlant = value; }
        }

        public string JobNum
        {
            get { return m_sJobNum; }
            set { m_sJobNum = value; }
        }

        public DateTime StartDate
        {
            get { return m_dtStartDate; }
            set { m_dtStartDate = value; }
        }
        public DateTime LastClockInDate
        {
            get { return m_dtLastClockInDate; }
            set { m_dtLastClockInDate = value; }
        }
        public DateTime DueDate
        {
            get { return m_dtDueDate; }
            set { m_dtDueDate = value; }
        }
        public DateTime RequiredDate
        {
            get { return m_dtRequiredDate; }
            set { m_dtRequiredDate = value; }
        }
        public decimal ProdQty
        {
            get { return m_dProdQty; }
            set { m_dProdQty = value; }
        }
        public string PartNum
        {
            get { return m_sPartNum; }
            set { m_sPartNum = value; }
        }
        public string PartRevNum
        {
            get { return m_sPartRevNum; }
            set { m_sPartRevNum = value; }
        }
        public decimal PercentageComplete
        {
            get { return m_dPercentageComplete; }
            set { m_dPercentageComplete = value; }
        }
        public decimal EstTotalCost
        {
            get { return m_dEstTotalCost; }
            set { m_dEstTotalCost = value; }
        }
        public decimal PercentEstTotalCost
        {
            get { return m_dPercentEstTotalCost; }
            set { m_dPercentEstTotalCost = value; }
        }
        public decimal ActTotalCost
        {
            get { return m_dActTotalCost; }
            set { m_dActTotalCost = value; }
        }
        public decimal PercentHoursComplete
        {
            get { return m_dPercentHoursComplete; }
            set { m_dPercentHoursComplete = value; }
        }
        public decimal EstOperationalHours
        {
            get { return m_dEstOperationalHours; }
            set { m_dEstOperationalHours = value; }
        }
        public decimal PercentEstOperationalHours
        {
            get { return m_dPercentEstOperationalHours; }
            set { m_dPercentEstOperationalHours = value; }
        }
        public decimal ActualOperationalHours
        {
            get { return m_dActualOperationalHours; }
            set { m_dActualOperationalHours = value; }
        }
        public decimal PercentLaborComplete
        {
            get { return m_dPercentLaborComplete; }
            set { m_dPercentLaborComplete = value; }
        }
        public decimal EstLaborCost
        {
            get { return m_dEstLaborCost; }
            set { m_dEstLaborCost = value; }
        }
        public decimal PercentEstLaborCost
        {
            get { return m_dPercentEstLaborCost; }
            set { m_dPercentEstLaborCost = value; }
        }
        public decimal ActualLaborCost
        {
            get { return m_dActualLaborCost; }
            set { m_dActualLaborCost = value; }
        }
        public decimal EstBurdenCost
        {
            get { return m_dEstBurdenCost; }
            set { m_dEstBurdenCost = value; }
        }
        public decimal PercentEstBurdenCost
        {
            get { return m_dPercentEstBurdenCost; }
            set { m_dPercentEstBurdenCost = value; }
        }
        public decimal ActualBurdenCost
        {
            get { return m_dActualBurdenCost; }
            set { m_dActualBurdenCost = value; }
        }
        public decimal PercentSubcontractComplete
        {
            get { return m_dPercentSubcontractComplete; }
            set { m_dPercentSubcontractComplete = value; }
        }
        public decimal EstSubcontractCost
        {
            get { return m_dEstSubcontractCost; }
            set { m_dEstSubcontractCost = value; }
        }
        public decimal PercentEstSubcontractCost
        {
            get { return m_dPercentEstSubcontractCost; }
            set { m_dPercentEstSubcontractCost = value; }
        }
        public decimal ActualSubcontractCost
        {
            get { return m_dActualSubcontractCost; }
            set { m_dActualSubcontractCost = value; }
        }
        public decimal PercentMaterialComplete
        {
            get { return m_dPercentMaterialComplete; }
            set { m_dPercentMaterialComplete = value; }
        }
        public decimal EstMaterialCost
        {
            get { return m_dEstMaterialCost; }
            set { m_dEstMaterialCost = value; }
        }
        public decimal PercentEstMaterialCost
        {
            get { return m_dPercentEstMaterialCost; }
            set { m_dPercentEstMaterialCost = value; }
        }
        public decimal ActualMaterialCost
        {
            get { return m_dActualMaterialCost; }
            set { m_dActualMaterialCost = value; }
        }
        public decimal EstMaterialBurdenCost
        {
            get { return m_dEstMaterialBurdenCost; }
            set { m_dEstMaterialBurdenCost = value; }
        }
        public decimal PercentEstMaterialBurdenCost
        {
            get { return m_dPercentEstMaterialBurdenCost; }
            set { m_dPercentEstMaterialBurdenCost = value; }
        }
        public decimal ActualMaterialBurdenCost
        {
            get { return m_dActualMaterialBurdenCost; }
            set { m_dActualMaterialBurdenCost = value; }
        }
        public bool JobClosed
        {
            get { return m_bJobClosed; }
            set { m_bJobClosed = value; }
        }
        public bool JobComplete
        {
            get { return m_bJobComplete; }
            set { m_bJobComplete = value; }
        }
        public int CustNum
        {
            get { return m_iCustNum; }
            set { m_iCustNum = value; }
        }
        public string CustID
        {
            get { return m_sCustID; }
            set { m_sCustID = value; }
        }
        public string CustomerName
        {
            get { return m_sCustomerName; }
            set { m_sCustomerName = value; }
        }
        public string MarketSegment
        {
            get { return m_sMarketSegment; }
            set { m_sMarketSegment = value; }
        }
        public string MarketSegmentDescription
        {
            get { return m_sMarketSegmentDescription; }
            set { m_sMarketSegmentDescription = value; }
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
        public string ProductPortfolio
        {
            get { return m_sProductPortfolio; }
            set { m_sProductPortfolio = value; }
        }
        public string ProductPortfolioDescription
        {
            get { return m_sProductPortfolioDescription; }
            set { m_sProductPortfolioDescription = value; }
        }
        public decimal OurReqQty
        {
            get { return m_dOurReqQty; }
            set { m_dOurReqQty = value; }
        }
        public decimal UnitPrice
        {
            get { return m_dUnitPrice; }
            set { m_dUnitPrice = value; }
        }
        public decimal ProjectedTotalCost
        {
            get { return m_dProjectedTotalCost; }
            set { m_dProjectedTotalCost = value; }
        }
        public decimal MarginToDate
        {
            get { return m_dMarginToDate; }
            set { m_dMarginToDate = value; }
        }
        public bool SalesDemandLink
        {
            get { return m_bSalesDemandLink; }
            set { m_bSalesDemandLink = value; }
        }
        public decimal PartUnitPrice
        {
            get { return m_dPartUnitPrice; }
            set { m_dPartUnitPrice = value; }
        }
        public bool PartOnTheFly
        {
            get { return m_bPartOnTheFly; }
            set { m_bPartOnTheFly = value; }
        }
        public decimal ShippedQty
        {
            get { return m_dShippedQty; }
            set { m_dShippedQty = value; }
        }
        public decimal ReceivedQty
        {
            get { return m_dReceivedQty; }
            set { m_dReceivedQty = value; }
        }
        public bool OpenRelease
        {
            get { return m_bOpenRelease; }
            set { m_bOpenRelease = value; }
        }
        public string WarehouseCode
        {
            get { return m_sWarehouseCode; }
            set { m_sWarehouseCode = value; }
        }
        private decimal EstimatedRemainingCost
        {
            get { return m_dEstimatedRemainingCost; }
            set { m_dEstimatedRemainingCost = value; }
        }
        #endregion

        #region Data Members
        private string m_sCompany;
        private string m_sPlant;
        private string m_sJobNum;
        private DateTime m_dtStartDate;
        private DateTime m_dtLastClockInDate;
        private DateTime m_dtDueDate;
        private DateTime m_dtRequiredDate;
        private decimal m_dProdQty;
        private string m_sPartNum;
        private string m_sPartRevNum;
        private decimal m_dPercentageComplete;
        private decimal m_dEstTotalCost;
        private decimal m_dPercentEstTotalCost;
        private decimal m_dActTotalCost;
        private decimal m_dPercentHoursComplete;
        private decimal m_dEstOperationalHours;
        private decimal m_dPercentEstOperationalHours;
        private decimal m_dActualOperationalHours;
        private decimal m_dPercentLaborComplete;
        private decimal m_dEstLaborCost;
        private decimal m_dPercentEstLaborCost;
        private decimal m_dActualLaborCost;
        private decimal m_dEstBurdenCost;
        private decimal m_dPercentEstBurdenCost;
        private decimal m_dActualBurdenCost;
        private decimal m_dPercentSubcontractComplete;
        private decimal m_dEstSubcontractCost;
        private decimal m_dPercentEstSubcontractCost;
        private decimal m_dActualSubcontractCost;
        private decimal m_dPercentMaterialComplete;
        private decimal m_dEstMaterialCost;
        private decimal m_dPercentEstMaterialCost;
        private decimal m_dActualMaterialCost;
        private decimal m_dEstMaterialBurdenCost;
        private decimal m_dPercentEstMaterialBurdenCost;
        private decimal m_dActualMaterialBurdenCost;
        private bool m_bJobClosed;
        private bool m_bJobComplete;
        private int m_iCustNum;
        private string m_sCustID;
        private string m_sCustomerName;
        private string m_sMarketSegment;
        private string m_sMarketSegmentDescription;
        private int m_iOrderNum;
        private int m_iOrderLine;
        private int m_iOrderRelNum;
        private string m_sProductPortfolio;
        private string m_sProductPortfolioDescription;
        private decimal m_dOurReqQty;
        private decimal m_dUnitPrice;
        private decimal m_dProjectedTotalCost;
        private decimal m_dMarginToDate;
        private bool m_bSalesDemandLink;
        private decimal m_dPartUnitPrice;
        private bool m_bPartOnTheFly;
        private decimal m_dShippedQty;
        private decimal m_dReceivedQty;
        private bool m_bOpenRelease;
        private string m_sWarehouseCode;

        private decimal m_dEstimatedRemainingCost;

        #endregion
    }
}
