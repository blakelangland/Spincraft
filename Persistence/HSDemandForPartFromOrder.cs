using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace HorizonScientific
{
    public class HSDemandForPartFromOrder
    {
        //DemandForPartsFromOrders

        #region Constructors

        public HSDemandForPartFromOrder(DataRow oRow)
        {
            if ((oRow["PartDtl_PartNum"] != DBNull.Value) && (string.IsNullOrEmpty((string)oRow["PartDtl_PartNum"]) == false))
            {
                m_sPartNum = (string)oRow["PartDtl_PartNum"];
            }
            if ((oRow["PartDtl_RevisionNum"] != DBNull.Value) && (string.IsNullOrEmpty((string)oRow["PartDtl_RevisionNum"]) == false))
            {
                m_sRevisionNumber = (string)oRow["PartDtl_RevisionNum"];
            }
            if ((oRow["OrderDtl_LineDesc"] != DBNull.Value) && (string.IsNullOrEmpty((string)oRow["OrderDtl_LineDesc"]) == false))
            {
                m_sDescription = (string)oRow["OrderDtl_LineDesc"];
            }
            if (oRow["PartDtl_DueDate"] != DBNull.Value)
            {
                m_dtDueDate = (DateTime)oRow["PartDtl_DueDate"];
                m_bDueDateNotSet = false;
            }
            else
            {
                m_dtDueDate = DateTime.Now;
                m_bDueDateNotSet = true;
            }
            if (oRow["Calculated_NetDemandQuantity"] != DBNull.Value)
            {
                m_dNetDemandQuantity = (decimal)oRow["Calculated_NetDemandQuantity"];
            }
            if ((oRow["PartDtl_SourceFile"] != DBNull.Value) && (string.IsNullOrEmpty((string)oRow["PartDtl_SourceFile"]) == false))
            {
                m_sSourceFile = (string)oRow["PartDtl_SourceFile"];
            }
            if (oRow["OrderRel_OrderNum"] != DBNull.Value)
            {
                m_iOrderNumber = (int)oRow["OrderRel_OrderNum"];
            }
            if (oRow["OrderRel_OrderLine"] != DBNull.Value)
            {
                m_iLineNumber = (int)oRow["OrderRel_OrderLine"];
            }
            if (oRow["OrderRel_OrderRelNum"] != DBNull.Value)
            {
                m_iReleaseNumber = (int)oRow["OrderRel_OrderRelNum"];
            }
            if (oRow["OrderRel_OurReqQty"] != DBNull.Value)
            {
                m_dRequiredQuantity = (decimal)oRow["OrderRel_OurReqQty"];
            }
            if (oRow["Calculated_ShippedQuantity"] != DBNull.Value)
            {
                m_dShippedQuantity = (decimal)oRow["Calculated_ShippedQuantity"];
            }
            if (oRow["OrderRel_ReqDate"] != DBNull.Value)
            {
                m_dtReleaseRequiredDate = (DateTime)oRow["OrderRel_ReqDate"];
                m_bRequiredDateNotSet = false;
            }
            else
            {
                m_dtReleaseRequiredDate = DateTime.Now;
                m_bRequiredDateNotSet = true;
            }
            if (oRow["OrderRel_FirmRelease"] != DBNull.Value)
            {
                m_bFirm = (bool)oRow["OrderRel_FirmRelease"];
            }
            if (oRow["Calculated_RemainingQty"] != DBNull.Value)
            {
                m_dRemainingQuantity = (decimal)oRow["Calculated_RemainingQty"];
            }
            if (oRow["OrderDtl_DocUnitPrice"] != DBNull.Value)
            {
                m_dUnitPrice = (decimal)oRow["OrderDtl_DocUnitPrice"];
            }
            if ((oRow["Customer_CustID"] != DBNull.Value) && (string.IsNullOrEmpty((string)oRow["Customer_CustID"]) == false))
            {
                m_sCustomerId = (string)oRow["Customer_CustID"];
            }
            if (oRow["Customer_CustNum"] != DBNull.Value)
            {
                m_iCustNum = (int)oRow["Customer_CustNum"];
            }
            if ((oRow["Customer_Name"] != DBNull.Value) && (string.IsNullOrEmpty((string)oRow["Customer_Name"]) == false))
            {
                m_sCustomerName = (string)oRow["Customer_Name"];
            }
        }

        #endregion

        #region Properties

        public string PartNum
        {
            get { return m_sPartNum; }
        }
        public string RevisionNumber
        {
            get { return m_sRevisionNumber; }
        }
        public string Description
        {
            get { return m_sDescription; }
        }
        public DateTime DueDate
        {
            get { return m_dtDueDate; }
        }
        public decimal NetDemandQuantity
        {
            get { return m_dNetDemandQuantity; }
        }
        public string SourceFile
        {
            get { return m_sSourceFile; }
        }
        public int OrderNumber
        {
            get { return m_iOrderNumber; }
        }
        public int LineNumber
        {
            get { return m_iLineNumber; }
        }
        public int ReleaseNumber
        {
            get { return m_iReleaseNumber; }
        }
        public decimal RequiredQuantity
        {
            get { return m_dRequiredQuantity; }
        }
        public decimal ShippedQuantity
        {
            get { return m_dShippedQuantity; }
        }
        public DateTime ReleaseRequiredDate
        {
            get { return m_dtReleaseRequiredDate; }
        }
        public bool Firm
        {
            get { return m_bFirm; }
        }
        public decimal RemainingQuantity
        {
            get { return m_dRemainingQuantity; }
        }
        public decimal UnitPrice
        {
            get { return m_dUnitPrice; }
        }
        public string CustomerId
        {
            get { return m_sCustomerId; }
        }
        public int CustNum
        {
            get { return m_iCustNum; }
        }
        public string CustomerName
        {
            get { return m_sCustomerName; }
        }

        public bool DueDateNoteSet
        {
            get { return m_bDueDateNotSet; }
        }
        public bool RequiredDateNotSet
        {
            get { return m_bRequiredDateNotSet; }
        }
        #endregion

        #region Data Members

        private string m_sPartNum;
        private string m_sRevisionNumber;
        private string m_sDescription;
        private DateTime m_dtDueDate;
        private decimal m_dNetDemandQuantity;
        private string m_sSourceFile;
        private int m_iOrderNumber;
        private int m_iLineNumber;
        private int m_iReleaseNumber;
        private decimal m_dRequiredQuantity;
        private decimal m_dShippedQuantity;
        private DateTime m_dtReleaseRequiredDate;
        private bool m_bFirm;
        private decimal m_dRemainingQuantity;
        private decimal m_dUnitPrice;
        private string m_sCustomerId;
        private int m_iCustNum;
        private string m_sCustomerName;

        private bool m_bDueDateNotSet;
        private bool m_bRequiredDateNotSet;
        #endregion
    }

}
