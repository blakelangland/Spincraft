using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace HorizonScientific
{
    public class PODetailHistory
    {
        #region Constructors
        public PODetailHistory(PODetailHistory oOriginal)
        {
            this.m_iPONum = oOriginal.m_iPONum;
            this.m_iPOLine = oOriginal.m_iPOLine;
            this.m_iPORelNum = oOriginal.m_iPORelNum;
            this.m_sPartNum = oOriginal.m_sPartNum;
            this.m_sDescription = oOriginal.m_sDescription;
            this.m_dRelQty = oOriginal.m_dRelQty;
            this.m_dUnitCost = oOriginal.m_dUnitCost;
            this.m_dExtPrice = oOriginal.m_dExtPrice;
            this.m_iRemainingQty = oOriginal.m_iRemainingQty;
            this.m_dtOrderDate = oOriginal.m_dtOrderDate;
            this.m_dtDueDate = oOriginal.m_dtDueDate;
            this.m_dtPromiseDate = oOriginal.m_dtPromiseDate;
            this.m_sVendorId = oOriginal.m_sVendorId;
            this.m_sVendorName = oOriginal.m_sVendorName;
            this.m_sBuyerId = oOriginal.m_sBuyerId;
            this.m_dtExpectedArrivalDate = oOriginal.m_dtExpectedArrivalDate;
            this.m_sPOComments = oOriginal.m_sPOComments;
        }

        public PODetailHistory(DataRow oDataRow)
        {
            if (oDataRow["PORel_PONum"] != DBNull.Value)
            {
                m_iPONum = (int)oDataRow["PORel_PONum"];
            }
            if (oDataRow["PORel_POLine"] != DBNull.Value)
            {
                m_iPOLine = (int)oDataRow["PORel_POLine"];
            }
            if (oDataRow["PORel_PORelNum"] != DBNull.Value)
            {
                m_iPORelNum = (int)oDataRow["PORel_PORelNum"];
            }
            if ((oDataRow["PODetail_PartNum"] != DBNull.Value) && (string.IsNullOrEmpty((string)oDataRow["PODetail_PartNum"]) == false))
            {
                m_sPartNum = (string)oDataRow["PODetail_PartNum"];
            }
            if ((oDataRow["Calculated_Description"] != DBNull.Value) && (string.IsNullOrEmpty((string)oDataRow["Calculated_Description"]) == false))
            {
                m_sDescription = (string)oDataRow["Calculated_Description"];
            }
            if (oDataRow["PORel_RelQty"] != DBNull.Value)
            {
                m_dRelQty = (decimal)oDataRow["PORel_RelQty"];
            }
            if (oDataRow["PODetail_UnitCost"] != DBNull.Value)
            {
                m_dUnitCost = (decimal)oDataRow["PODetail_UnitCost"];
            }
            if (oDataRow["Calculated_ExtCost"] != DBNull.Value)
            {
                m_dExtPrice = (decimal)oDataRow["Calculated_ExtCost"];
            }
            if (oDataRow["Calculated_RemainingQty"] != DBNull.Value)
            {
                m_iRemainingQty = (int)oDataRow["Calculated_RemainingQty"];
            }
            if (oDataRow["POHeader_OrderDate"] != DBNull.Value)
            {
                m_dtOrderDate = (DateTime)oDataRow["POHeader_OrderDate"];
            }
            if (oDataRow["PORel_DueDate"] != DBNull.Value)
            {
                m_dtDueDate = (DateTime)oDataRow["PORel_DueDate"];
            }
            if (oDataRow["PORel_PromiseDt"] != DBNull.Value)
            {
                m_dtPromiseDate = (DateTime)oDataRow["PORel_PromiseDt"];
            }
            if ((oDataRow["Vendor_VendorID"] != DBNull.Value) && (string.IsNullOrEmpty((string)oDataRow["Vendor_VendorID"]) == false))
            {
                m_sVendorId = (string)oDataRow["Vendor_VendorID"];
            }
            if ((oDataRow["Vendor_Name"] != DBNull.Value) && (string.IsNullOrEmpty((string)oDataRow["Vendor_Name"]) == false))
            {
                m_sVendorName = (string)oDataRow["Vendor_Name"];
            }
            if ((oDataRow["POHeader_BuyerID"] != DBNull.Value) && (string.IsNullOrEmpty((string)oDataRow["POHeader_BuyerID"]) == false))
            {
                m_sBuyerId = (string)oDataRow["POHeader_BuyerID"];
            }
            if (oDataRow["PORel_ExpectedArrivalDate_c"] != DBNull.Value)
            {
                m_dtExpectedArrivalDate = (DateTime)oDataRow["PORel_ExpectedArrivalDate_c"];
            }
            if ((oDataRow["PORel_PurchaseOrderComments_c"] != DBNull.Value) && (string.IsNullOrEmpty((string)oDataRow["PORel_PurchaseOrderComments_c"]) == false))
            {
                m_sPOComments = (string)oDataRow["PORel_PurchaseOrderComments_c"];
            }
        }

        #endregion

        #region Methods
        #endregion

        #region Properties

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

        public string PartNum
        {
            get { return m_sPartNum; }
            set { m_sPartNum = value; }
        }

        public string Description
        {
            get { return m_sDescription; }
            set { m_sDescription = value; }
        }

        public decimal RelQty
        {
            get { return m_dRelQty; }
            set { m_dRelQty = value; }
        }

        public decimal UnitCost
        {
            get { return m_dUnitCost; }
            set { m_dUnitCost = value; }
        }

        public decimal ExtPrice
        {
            get { return m_dExtPrice; }
            set { m_dExtPrice = value; }
        }

        public int RemainingQty
        {
            get { return m_iRemainingQty; }
            set { m_iRemainingQty = value; }
        }

        public DateTime OrderDate
        {
            get { return m_dtOrderDate; }
            set { m_dtOrderDate = value; }
        }

        public DateTime DueDate
        {
            get { return m_dtDueDate; }
            set { m_dtDueDate = value; }
        }

        public DateTime PromiseDate
        {
            get { return m_dtPromiseDate; }
            set { m_dtPromiseDate = value; }
        }

        public string VendorId
        {
            get { return m_sVendorId; }
            set { m_sVendorId = value; }
        }

        public string VendorName
        {
            get { return m_sVendorName; }
            set { m_sVendorName = value; }
        }

        public string BuyerId
        {
            get { return m_sBuyerId; }
            set { m_sBuyerId = value; }
        }

        public DateTime ExpectedArrivalDate
        {
            get { return m_dtExpectedArrivalDate; }
            set { m_dtExpectedArrivalDate = value; }
        }

        public string POComments
        {
            get { return m_sPOComments; }
            set { m_sPOComments = value; }
        }

        #endregion

        #region Data Members

        private int m_iPONum;
        private int m_iPOLine;
        private int m_iPORelNum;
        private string m_sPartNum;
        private string m_sDescription;
        private decimal m_dRelQty;
        private decimal m_dUnitCost;
        private decimal m_dExtPrice;
        private int m_iRemainingQty;
        private DateTime m_dtOrderDate = DateTime.MinValue;
        private DateTime m_dtDueDate = DateTime.MinValue;
        private DateTime m_dtPromiseDate = DateTime.MinValue;
        private string m_sVendorId;
        private string m_sVendorName;
        private string m_sBuyerId;
        private DateTime m_dtExpectedArrivalDate = DateTime.MinValue;
        private string m_sPOComments;


        #endregion
    }
}
