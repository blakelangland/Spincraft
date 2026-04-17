using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace HorizonScientific
{
    public class SOBacklogForPartAnalysis
    {
        #region Constructors
        public SOBacklogForPartAnalysis()
        {
        }

        public SOBacklogForPartAnalysis(DataRow oDataRow)
        {
            if ((oDataRow["OrderHed_Company"] != DBNull.Value) && (string.IsNullOrEmpty((string)oDataRow["OrderHed_Company"]) == false))
            {
                m_sCompany = (string)oDataRow["OrderHed_Company"];
            }
            if ((oDataRow["OrderHed_EntryPerson"] != DBNull.Value) && (string.IsNullOrEmpty((string)oDataRow["OrderHed_EntryPerson"]) == false))
            {
                m_sEntryPerson = (string)oDataRow["OrderHed_EntryPerson"];
            }
            if ((oDataRow["OrderHed_PONum"] != DBNull.Value) && (string.IsNullOrEmpty((string)oDataRow["OrderHed_PONum"]) == false))
            {
                m_sPONum = (string)oDataRow["OrderHed_PONum"];
            }
            if (oDataRow["OrderHed_OrderNum"] != DBNull.Value)
            {
                m_iOrderNum = (int)oDataRow["OrderHed_OrderNum"];
            }
            if (oDataRow["OrderRel_OrderLine"] != DBNull.Value)
            {
                m_iOrderLine = (int)oDataRow["OrderRel_OrderLine"];
            }
            if (oDataRow["OrderRel_OrderRelNum"] != DBNull.Value)
            {
                m_iOrderRelease = (int)oDataRow["OrderRel_OrderRelNum"];
            }
            if (oDataRow["OrderDtl_KitParentLine"] != DBNull.Value)
            {
                m_iKitParentLine = (int)oDataRow["OrderDtl_KitParentLine"];
            }
            if ((oDataRow["Customer_Name"] != DBNull.Value) && (string.IsNullOrEmpty((string)oDataRow["Customer_Name"]) == false))
            {
                m_sCustomer = (string)oDataRow["Customer_Name"];
                m_sCustomer = StringExt.CleanString(m_sCustomer);
            }
            if (oDataRow["Customer_CreditHold"] != DBNull.Value)
            {
                m_bCustomerCreditHold = (bool)oDataRow["Customer_CreditHold"];
            }
            if (oDataRow["OrderHed_ShipOrderComplete"] != DBNull.Value)
            {
                m_bShipOrderComplete = (bool)oDataRow["OrderHed_ShipOrderComplete"];
            }
            if (oDataRow["OrderDtl_ShipLineComplete"] != DBNull.Value)
            {
                m_bShipLineComplete = (bool)oDataRow["OrderDtl_ShipLineComplete"];
            }
            if (oDataRow["OrderRel_Make"] != DBNull.Value)
            {
                m_bMakeDirect = (bool)oDataRow["OrderRel_Make"];
            }
            if (oDataRow["OrderHed_OrderHeld"] != DBNull.Value)
            {
                m_bHoldOrder = (bool)oDataRow["OrderHed_OrderHeld"];
            }
            if (oDataRow["OrderHed_OrderDate"] != DBNull.Value)
            {
                m_dtOrderDate = (DateTime)oDataRow["OrderHed_OrderDate"];
            }
            if (oDataRow["OrderRel_ReqDate"] != DBNull.Value)
            {
                m_dtRequiredByDate = (DateTime)oDataRow["OrderRel_ReqDate"];
            }
            if ((oDataRow["OrderDtl_PartNum"] != DBNull.Value) && (string.IsNullOrEmpty((string)oDataRow["OrderDtl_PartNum"]) == false))
            {
                m_sPartNum = (string)oDataRow["OrderDtl_PartNum"];
            }
            if ((oDataRow["OrderDtl_RevisionNum"] != DBNull.Value) && (string.IsNullOrEmpty((string)oDataRow["OrderDtl_RevisionNum"]) == false))
            {
                m_sPartRev = (string)oDataRow["OrderDtl_RevisionNum"];
            }
            if (oDataRow["Calculated_ExtPrice"] != DBNull.Value)
            {
                m_dExtPrice = (decimal)oDataRow["Calculated_ExtPrice"];
            }
            if ((oDataRow["Part_TypeCode"] != DBNull.Value) && (string.IsNullOrEmpty((string)oDataRow["Part_TypeCode"]) == false))
            {
                m_sTypeCode = (string)oDataRow["Part_TypeCode"];
            }
            if ((oDataRow["PartPlant_PersonID"] != DBNull.Value) && (string.IsNullOrEmpty((string)oDataRow["PartPlant_PersonID"]) == false))
            {
                m_sPlanner = (string)oDataRow["PartPlant_PersonID"];
            }
            if (oDataRow["Part_QtyBearing"] != DBNull.Value)
            {
                m_bQuantityBearing = (bool)oDataRow["Part_QtyBearing"];
            }
            if ((oDataRow["Calculated_PartDescription"] != DBNull.Value) && (string.IsNullOrEmpty((string)oDataRow["Calculated_PartDescription"]) == false))
            {
                m_sDescription = (string)oDataRow["Calculated_PartDescription"];
                m_sDescription = StringExt.CleanString(m_sDescription);
            }
            if (oDataRow["Calculated_RemainingQty"] != DBNull.Value)
            {
                m_dRemainingQty = (decimal)oDataRow["Calculated_RemainingQty"];
            }
            if ((oDataRow["OrderRel_Plant"] != DBNull.Value) && (string.IsNullOrEmpty((string)oDataRow["OrderRel_Plant"]) == false))
            {
                m_sPlant = (string)oDataRow["OrderRel_Plant"];
            }
            if ((oDataRow["OrderRel_WarehouseCode"] != DBNull.Value) && (string.IsNullOrEmpty((string)oDataRow["OrderRel_WarehouseCode"]) == false))
            {
                m_sWarehouse = (string)oDataRow["OrderRel_WarehouseCode"];
            }
            if ((oDataRow["Calculated_ShipToName"] != DBNull.Value) && (string.IsNullOrEmpty((string)oDataRow["Calculated_ShipToName"]) == false))
            {
                m_sShipToName = (string)oDataRow["Calculated_ShipToName"];
                m_sShipToName = StringExt.CleanString(m_sShipToName);
            }
            if ((oDataRow["Calculated_Address1"] != DBNull.Value) && (string.IsNullOrEmpty((string)oDataRow["Calculated_Address1"]) == false))
            {
                m_sAddress1 = (string)oDataRow["Calculated_Address1"];
                m_sAddress1 = StringExt.CleanString(m_sAddress1);
            }
            if ((oDataRow["Calculated_Address2"] != DBNull.Value) && (string.IsNullOrEmpty((string)oDataRow["Calculated_Address2"]) == false))
            {
                m_sAddress2 = (string)oDataRow["Calculated_Address2"];
                m_sAddress2 = StringExt.CleanString(m_sAddress2);
            }
            if ((oDataRow["Calculated_Address3"] != DBNull.Value) && (string.IsNullOrEmpty((string)oDataRow["Calculated_Address3"]) == false))
            {
                m_sAddress3 = (string)oDataRow["Calculated_Address3"];
                m_sAddress3 = StringExt.CleanString(m_sAddress3);
            }
            if ((oDataRow["Calculated_City"] != DBNull.Value) && (string.IsNullOrEmpty((string)oDataRow["Calculated_City"]) == false))
            {
                m_sCity = (string)oDataRow["Calculated_City"];
                m_sCity = StringExt.CleanString(m_sCity);
            }
            if ((oDataRow["Calculated_State"] != DBNull.Value) && (string.IsNullOrEmpty((string)oDataRow["Calculated_State"]) == false))
            {
                m_sState = (string)oDataRow["Calculated_State"];
                m_sState = StringExt.CleanString(m_sState);
            }
            if ((oDataRow["Calculated_Zip"] != DBNull.Value) && (string.IsNullOrEmpty((string)oDataRow["Calculated_Zip"]) == false))
            {
                m_sZip = (string)oDataRow["Calculated_Zip"];
                m_sZip = StringExt.CleanString(m_sZip);
            }
            if ((oDataRow["JobProd_JobNum"] != DBNull.Value) && (string.IsNullOrEmpty((string)oDataRow["JobProd2_JobNum"]) == false))
            {
                m_sJobNum = (string)oDataRow["JobProd_JobNum"];
            }
            if ((oDataRow["JobProd_WarehouseCode"] != DBNull.Value) && (string.IsNullOrEmpty((string)oDataRow["JobProd_WarehouseCode"]) == false))
            {
                m_sWarehouseCode = (string)oDataRow["JobProd_WarehouseCode"];
            }
            if ((oDataRow["JobProd_TargetJobNum"] != DBNull.Value) && (string.IsNullOrEmpty((string)oDataRow["JobProd_TargetJobNum"]) == false))
            {
                m_sTargetJobNum = (string)oDataRow["JobProd_TargetJobNum"];
            }
            if (oDataRow["JobProd_TargetAssemblySeq"] != DBNull.Value) 
            {
                m_iTargetAssemblySeq = (int)oDataRow["JobProd_TargetAssemblySeq"];
            }
            if (oDataRow["JobProd_TargetMtlSeq"] != DBNull.Value)
            {
                m_iTargetMaterialSeq = (int)oDataRow["JobProd_TargetMtlSeq"];
            }
        }
        #endregion

        #region Methods
        public void AddRelatedJob(HSOpenJob oJob)
        {
            m_oRelatedJobs.Add(oJob);
        }
        #endregion

        #region Properties

        public string Company
        {
            get { return m_sCompany; }
            set { m_sCompany = value; }
        }

        public string EntryPerson
        {
            get { return m_sEntryPerson; }
            set { m_sEntryPerson = value; }
        }
        public string PONum
        {
            get { return m_sPONum; }
            set { m_sPONum = value; }
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

        public int OrderRelease
        {
            get { return m_iOrderRelease; }
            set { m_iOrderRelease = value; }
        }

        public int KitParentLine
        {
            get { return m_iKitParentLine; }
            set { m_iKitParentLine = value; }
        }

        public string Customer
        {
            get { return m_sCustomer; }
            set { m_sCustomer = value; }
        }

        public bool CustomerCreditHold
        {
            get { return m_bCustomerCreditHold; }
            set { m_bCustomerCreditHold = value; }
        }

        public bool ShipOrderComplete
        {
            get { return m_bShipOrderComplete; }
            set { m_bShipOrderComplete = value; }
        }

        public bool ShipLineComplete
        {
            get { return m_bShipLineComplete; }
            set { m_bShipLineComplete = value; }
        }

        public bool MakeDirect
        {
            get { return m_bMakeDirect; }
            set { m_bMakeDirect = value; }
        }

        public bool HoldOrder
        {
            get { return m_bHoldOrder; }
            set { m_bHoldOrder = value; }
        }

        public DateTime OrderDate
        {
            get { return m_dtOrderDate; }
            set { m_dtOrderDate = value; }
        }

        public DateTime RequiredByDate
        {
            get { return m_dtRequiredByDate; }
            set { m_dtRequiredByDate = value; }
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

        public decimal ExtPrice
        {
            get { return m_dExtPrice; }
            set { m_dExtPrice = value; }
        }

        public string TypeCode
        {
            get { return m_sTypeCode; }
            set { m_sTypeCode = value; }
        }

        public string Planner
        {
            get { return m_sPlanner; }
            set { m_sPlanner = value; }
        }
        public bool QuantityBearing
        {
            get { return m_bQuantityBearing; }
            set { m_bQuantityBearing = value; }
        }

        public string Description
        {
            get { return m_sDescription; }
            set { m_sDescription = value; }
        }

        public decimal RemainingQty
        {
            get { return m_dRemainingQty; }
            set { m_dRemainingQty = value; }
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

        public string JobNum
        {
            get { return m_sJobNum; }
            set { m_sJobNum = value; }
        }

        public string WarehouseCode
        {
            get { return m_sWarehouseCode; }
            set { m_sWarehouseCode = value; }
        }

        public string TargetJobNum
        {
            get { return m_sTargetJobNum; }
            set { m_sTargetJobNum = value; }
        }

        public int TargetAssemblySeq
        {
            get { return m_iTargetAssemblySeq; }
            set { m_iTargetAssemblySeq = value; }
        }

        public int TargetMaterialSeq
        {
            get { return m_iTargetMaterialSeq; }
            set { m_iTargetMaterialSeq = value; }
        }
        // used to help us track which orders we have already accounted for
        public bool Processed
        {
            get { return m_bProcessed; }
            set { m_bProcessed = value; }
        }

        public DateTime EffectiveShipDate
        {
            get { return m_dtEffectiveShipDate; }
            set { m_dtEffectiveShipDate = value; }
        }

        public bool PartShortage
        {
            get
            {
                bool bShortage = m_bPartShortage;
                if (m_bPartShortage == false)
                {
                    // if this order does not depend on a purchased part
                    // then we need to walk through all related jobs to 
                    // see if any of the jobs has a part shortage
                    foreach (HSOpenJob oJob in m_oRelatedJobs)
                    {
                        if (oJob.PartShortage == true)
                        {
                            bShortage = true;
                            break;
                        }
                    }

                    // if this order is make direct and there are no jobs created
                    // then we must assume we have a part shortage until we have 
                    // a job created that we can analyze
                    if ((m_bMakeDirect == true) && (m_oRelatedJobs.Count == 0))
                    {
                        bShortage = true;
                    }
                }
                return bShortage;
            }
            set
            {
                m_bPartShortage = value;
            }
        }
        public List<HSOpenJob> RelatedJobs
        {
            get { return m_oRelatedJobs; }
        }
        #endregion

        #region Data Members

        private string m_sCompany;
        public string m_sEntryPerson;
        private string m_sPONum;
        private int m_iOrderNum;
        private int m_iOrderLine;
        private int m_iOrderRelease;
        private int m_iKitParentLine;
        private string m_sCustomer;
        private bool m_bCustomerCreditHold;
        private bool m_bShipOrderComplete;
        private bool m_bShipLineComplete;
        private bool m_bMakeDirect;
        private bool m_bHoldOrder;
        private DateTime m_dtOrderDate;
        private DateTime m_dtRequiredByDate;
        private string m_sPartNum;
        private string m_sPartRev;
        private decimal m_dExtPrice;
        private string m_sTypeCode;
        private string m_sPlanner;
        private bool m_bQuantityBearing;
        private string m_sDescription;
        private decimal m_dRemainingQty;
        private string m_sPlant;
        private string m_sWarehouse;
        private string m_sShipToName;
        private string m_sAddress1;
        private string m_sAddress2;
        private string m_sAddress3;
        private string m_sCity;
        private string m_sState;
        private string m_sZip;
        private string m_sJobNum;
        private string m_sWarehouseCode;
        private string m_sTargetJobNum;
        private int m_iTargetAssemblySeq;
        private int m_iTargetMaterialSeq;

        private DateTime m_dtEffectiveShipDate;
        private bool m_bProcessed;

        private List<HSOpenJob> m_oRelatedJobs = new List<HSOpenJob>();
        private bool m_bPartShortage = false;
        #endregion
    }
}
