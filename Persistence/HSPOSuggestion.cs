using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace HorizonScientific
{
    public class HSPOSuggestion
    {
        //SuggestedPOs

        #region Constructors

        public HSPOSuggestion(DataRow oDataRow)
        {
            if (oDataRow["SugPoDtl_SugNum"] != DBNull.Value)
            {
                m_iPOSuggestionNumber = (int)oDataRow["SugPoDtl_SugNum"];
            }
            if ((oDataRow["SugPoDtl_BuyerID"] != DBNull.Value) && (string.IsNullOrEmpty((string)oDataRow["SugPoDtl_BuyerID"]) == false))
            {
                m_sBuyerId = (string)oDataRow["SugPoDtl_BuyerID"];
            }
            if (oDataRow["SugPoDtl_DueDate"] != DBNull.Value)
            {
                m_dtDueDate = (DateTime)oDataRow["SugPoDtl_DueDate"];
            }
            else
            {
                m_dtDueDate = DateTime.Now;
            }
            if (oDataRow["SugPoDtl_RelQty"] != DBNull.Value)
            {
                m_dRequestedQuantity = (decimal)oDataRow["SugPoDtl_RelQty"];
            }
            if ((oDataRow["SugPoDtl_PartNum"] != DBNull.Value) && (string.IsNullOrEmpty((string)oDataRow["SugPoDtl_PartNum"]) == false))
            {
                m_sPart = (string)oDataRow["SugPoDtl_PartNum"];
            }
            if ((oDataRow["Calculated_ShortDescription"] != DBNull.Value) && (string.IsNullOrEmpty((string)oDataRow["Calculated_ShortDescription"]) == false))
            {
                m_sPartDescription = (string)oDataRow["Calculated_ShortDescription"];
            }
            if ((oDataRow["SugPoDtl_PUM"] != DBNull.Value) && (string.IsNullOrEmpty((string)oDataRow["SugPoDtl_PUM"]) == false))
            {
                m_sPurchaseUnitOfMeasure = (string)oDataRow["SugPoDtl_PUM"];
            }
            if ((oDataRow["SugPoDtl_IUM"] != DBNull.Value) && (string.IsNullOrEmpty((string)oDataRow["SugPoDtl_IUM"]) == false))
            {
                m_sInventoryUnitOfMeasure = (string)oDataRow["SugPoDtl_IUM"];
            }
            if ((oDataRow["SugPoDtl_VenPartNum"] != DBNull.Value) && (string.IsNullOrEmpty((string)oDataRow["SugPoDtl_VenPartNum"]) == false))
            {
                m_sVendorPart = (string)oDataRow["SugPoDtl_VenPartNum"];
            }
            if ((oDataRow["SugPoDtl_ClassID"] != DBNull.Value) && (string.IsNullOrEmpty((string)oDataRow["SugPoDtl_ClassID"]) == false))
            {
                m_sClassId = (string)oDataRow["SugPoDtl_ClassID"];
            }
            if ((oDataRow["SugPoDtl_VendorID"] != DBNull.Value) && (string.IsNullOrEmpty((string)oDataRow["SugPoDtl_VendorID"]) == false))
            {
                m_sVendorId = (string)oDataRow["SugPoDtl_VendorID"];
            }
            if (oDataRow["SugPoDtl_VendorNum"] != DBNull.Value)
            {
                m_iVendorNum = (int)oDataRow["SugPoDtl_VendorNum"];
            }
            if ((oDataRow["SugPoDtl_Name"] != DBNull.Value) && (string.IsNullOrEmpty((string)oDataRow["SugPoDtl_Name"]) == false))
            {
                m_sVendorName = (string)oDataRow["SugPoDtl_Name"];
            }
            if (oDataRow["SugPoDtl_LeadTime"] != DBNull.Value)
            {
                m_iActualLeadTime = (int)oDataRow["SugPoDtl_LeadTime"];
            }
            if (oDataRow["SugPoDtl_OrderByDate"] != DBNull.Value)
            {
                m_dtOrderByDate = (DateTime)oDataRow["SugPoDtl_OrderByDate"];
            }
            else
            {
                m_dtOrderByDate = DateTime.Now;
            }
            if (oDataRow["SugPoDtl_DropShip"] != DBNull.Value)
            {
                m_bDropShip = (bool)oDataRow["SugPoDtl_DropShip"];
            }
            if ((oDataRow["SugPoDtl_WarehouseCode"] != DBNull.Value) && (string.IsNullOrEmpty((string)oDataRow["SugPoDtl_WarehouseCode"]) == false))
            {
                m_sWarehouseCode = (string)oDataRow["SugPoDtl_WarehouseCode"];
            }
            if ((oDataRow["Calculated_BuyReason"] != DBNull.Value) && (string.IsNullOrEmpty((string)oDataRow["Calculated_BuyReason"]) == false))
            {
                m_sBuyReason = (string)oDataRow["Calculated_BuyReason"];
            }
            if ((oDataRow["SugPoDtl_UrgentPlanning"] != DBNull.Value) && (string.IsNullOrEmpty((string)oDataRow["SugPoDtl_UrgentPlanning"]) == false))
            {
                m_sUrgentPlanning = (string)oDataRow["SugPoDtl_UrgentPlanning"];
            }
            if (oDataRow["PartPlant_LeadTime"] != DBNull.Value)
            {
                m_iRequiredLeadTime = (int)oDataRow["PartPlant_LeadTime"];
            }
            if (oDataRow["Calculated_MinDueDate"] != DBNull.Value)
            {
                m_dtMinDueDate = (DateTime)oDataRow["Calculated_MinDueDate"];
            }
            else
            {
                m_dtMinDueDate = DateTime.Now;
            }
            if (oDataRow["Calculated_TotalOnHand"] != DBNull.Value)
            {
                m_dTotalPartsOnHand = (decimal)oDataRow["Calculated_TotalOnHand"];
            }
            if (oDataRow["Calculated_NetDemandQuantity"] != DBNull.Value)
            {
                m_dNetDemandQuantity = (decimal)oDataRow["Calculated_NetDemandQuantity"];
            }
            if (oDataRow["PartPlant_MinimumQty"] != DBNull.Value)
            {
                m_dMinimumQuantity = (decimal)oDataRow["PartPlant_MinimumQty"];
            }
            if (oDataRow["PartPlant_MaximumQty"] != DBNull.Value)
            {
                m_dMaximumQuantity = (decimal)oDataRow["PartPlant_MaximumQty"];
            }
            if (oDataRow["PartPlant_SafetyQty"] != DBNull.Value)
            {
                m_dSafetyQuantity = (decimal)oDataRow["PartPlant_SafetyQty"];
            }
            if (oDataRow["PartPlant_MinOrderQty"] != DBNull.Value)
            {
                m_dMinOrderQuantity = (decimal)oDataRow["PartPlant_MinOrderQty"];
            }
            if ((oDataRow["SugPoDtl_JobNum"] != DBNull.Value) && (string.IsNullOrEmpty((string)oDataRow["SugPoDtl_JobNum"]) == false))
            {
                m_sJobNumber = (string)oDataRow["SugPoDtl_JobNum"];
            }
            if (oDataRow["SugPoDtl_AssemblySeq"] != DBNull.Value)
            {
                m_iAssembly = (int)oDataRow["SugPoDtl_AssemblySeq"];
            }
            if (oDataRow["SugPoDtl_JobSeq"] != DBNull.Value)
            {
                m_iJobSequence = (int)oDataRow["SugPoDtl_JobSeq"];
            }
            if (oDataRow["SugPoDtl_OrderNum"] != DBNull.Value)
            {
                m_iOrderNum = (int)oDataRow["SugPoDtl_OrderNum"];
            }
            if (oDataRow["SugPoDtl_OrderLine"] != DBNull.Value)
            {
                m_iOrderLine = (int)oDataRow["SugPoDtl_OrderLine"];
            }
            if (oDataRow["SugPoDtl_OrderRelNum"] != DBNull.Value)
            {
                m_iOrderRel = (int)oDataRow["SugPoDtl_OrderRelNum"];
            }
            if (oDataRow["SugPoDtl_ReqNum"] != DBNull.Value)
            {
                m_iReqNum = (int)oDataRow["SugPoDtl_ReqNum"];
            }
            if (oDataRow["SugPoDtl_ReqLine"] != DBNull.Value)
            {
                m_iReqLine = (int)oDataRow["SugPoDtl_ReqLine"];
            }
            if ((oDataRow["Calculated_Comment"] != DBNull.Value) && (string.IsNullOrEmpty((string)oDataRow["Calculated_Comment"]) == false))
            {
                m_sComment = (string)oDataRow["Calculated_Comment"];
            }
            if ((oDataRow["Calculated_Status"] != DBNull.Value) && (string.IsNullOrEmpty((string)oDataRow["Calculated_Status"]) == false))
            {
                m_sStatus = (string)oDataRow["Calculated_Status"];
            }
        }

        #endregion

        #region Properties

        public int POSuggestionNumber
        {
            get { return m_iPOSuggestionNumber; }
        }
        public string BuyerId
        {
            get { return m_sBuyerId; }
        }
        public DateTime DueDate
        {
            get { return m_dtDueDate; }
        }
        public decimal RequestedQuantity
        {
            get { return m_dRequestedQuantity; }
        }
        public string Part
        {
            get { return m_sPart; }
        }
        public string PartDescription
        {
            get { return m_sPartDescription; }
        }
        public string PurchaseUnitOfMeasure
        {
            get { return m_sPurchaseUnitOfMeasure; }
        }
        public string InventoryUnitOfMeasure
        {
            get { return m_sInventoryUnitOfMeasure; }
        }
        public string VendorPart
        {
            get { return m_sVendorPart; }
        }
        public string ClassId
        {
            get { return m_sClassId; }
        }
        public string VendorId
        {
            get { return m_sVendorId; }
        }
        public int VendorNum
        {
            get { return m_iVendorNum; }
        }
        public string VendorName
        {
            get { return m_sVendorName; }
        }
        public int ActualLeadTime
        {
            get { return m_iActualLeadTime; }
        }
        public DateTime OrderByDate
        {
            get { return m_dtOrderByDate; }
        }
        public bool DropShip
        {
            get { return m_bDropShip; }
        }
        public string WarehouseCode
        {
            get { return m_sWarehouseCode; }
        }
        public string BuyReason
        {
            get { return m_sBuyReason; }
        }
        public string UrgentPlanning
        {
            get { return m_sUrgentPlanning; }
        }
        public int RequiredLeadTime
        {
            get { return m_iRequiredLeadTime; }
        }
        public DateTime MinDueDate
        {
            get { return m_dtMinDueDate; }
        }
        public decimal TotalPartsOnHand
        {
            get { return m_dTotalPartsOnHand; }
        }
        public decimal NetDemandQuantity
        {
            get { return m_dNetDemandQuantity; }
        }
        public decimal MinimumQuantity
        {
            get { return m_dMinimumQuantity; }
        }
        public decimal MaximumQuantity
        {
            get { return m_dMaximumQuantity; }
        }
        public decimal SafetyQuantity
        {
            get { return m_dSafetyQuantity; }
        }
        public decimal MinOrderQuantity
        {
            get { return m_dMinOrderQuantity; }
        }
        public string JobNumber
        {
            get { return m_sJobNumber; }
        }
        public int Assembly
        {
            get { return m_iAssembly; }
        }
        public int JobSequence
        {
            get { return m_iJobSequence; }
        }
        public int OrderNum
        {
            get { return m_iOrderNum; }
        }
        public int OrderLine
        {
            get { return m_iOrderLine; }
        }
        public int OrderRel
        {
            get { return m_iOrderRel; }
        }
        public int ReqNum
        {
            get { return m_iReqNum; }
        }
        public int ReqLine
        {
            get { return m_iReqLine; }
        }
        public string Comment
        {
            get { return m_sComment; }
        }
        public string Status
        {
            get { return m_sStatus; }
        }
        #endregion

        #region Data Members

        private int m_iPOSuggestionNumber;
        private string m_sBuyerId;
        private DateTime m_dtDueDate;
        private decimal m_dRequestedQuantity;
        private string m_sPart;
        private string m_sPartDescription;
        private string m_sPurchaseUnitOfMeasure;
        private string m_sInventoryUnitOfMeasure;
        private string m_sVendorPart;
        private string m_sClassId;
        private string m_sVendorId;
        private int m_iVendorNum;
        private string m_sVendorName;
        private int m_iActualLeadTime;
        private DateTime m_dtOrderByDate;
        private bool m_bDropShip;
        private string m_sWarehouseCode;
        private string m_sBuyReason;
        private string m_sUrgentPlanning;
        private int m_iRequiredLeadTime;
        private DateTime m_dtMinDueDate;
        private decimal m_dTotalPartsOnHand;
        private decimal m_dNetDemandQuantity;
        private decimal m_dMinimumQuantity;
        private decimal m_dMaximumQuantity;
        private decimal m_dSafetyQuantity;
        private decimal m_dMinOrderQuantity;
        private string m_sJobNumber;
        private int m_iAssembly;
        private int m_iJobSequence;
        private int m_iOrderNum;
        private int m_iOrderLine;
        private int m_iOrderRel;
        private int m_iReqNum;
        private int m_iReqLine;
        private string m_sComment;
        private string m_sStatus;

        #endregion
    }
}
