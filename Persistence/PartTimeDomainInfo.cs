using DocumentFormat.OpenXml.ExtendedProperties;
using HorizonScientific;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace HorizonScientific
{
    public enum PartInfoTimeLineStatus
    {
        PAST_DUE,
        THIS_WEEK,
        NEXT_WEEK,
        FUTURE,
        UNKNOWN
    }

    public class PartTimeDomainInfo
    {
        #region Constructors
        public PartTimeDomainInfo(string sCompany)
        {
            Company = sCompany;
        }

        public PartTimeDomainInfo(HSPartData oPartInfo, string sCompany)
        {
            Company = sCompany;

            if (oPartInfo != null)
            {
                m_sPartNum = oPartInfo.PartNum;
                m_sPartDescription = oPartInfo.Description;
                m_dMinimumInventoryLevel = oPartInfo.Minimum;
                m_dSafetyInventoryLevel = oPartInfo.Safety;
                m_dMaximumInventoryLevel = oPartInfo.Maximum;
                m_dMinimumOrderQuantity = oPartInfo.MinOrderQty;
                m_sBuyerId = oPartInfo.BuyerID;
                m_sClassId = oPartInfo.ClassID;
                m_sPurchaseUnitOfMeasure = oPartInfo.PUOM;
                m_sVendorName = oPartInfo.VendorName;
                m_sVendorId = oPartInfo.SupplierID;
                m_iVendorNum = oPartInfo.SupplierNum;
                m_iRequiredLeadTime = oPartInfo.Lead;
                m_bDropShip = oPartInfo.PartDropShip;
                m_dStartingInventoryOnHand = 0;
            }
        }

        public PartTimeDomainInfo(List<HSPOSuggestion> oPOSuggestions)
        {
            m_oPOSuggestionsForPart = oPOSuggestions;
            if ((m_oPOSuggestionsForPart != null) && (m_oPOSuggestionsForPart.Count > 0))
            {
                HSPOSuggestion oPOSuggestion = m_oPOSuggestionsForPart[0];
                m_sPartNum = oPOSuggestion.Part;
                m_sPartDescription = oPOSuggestion.PartDescription;
                m_dMinimumInventoryLevel = oPOSuggestion.MinimumQuantity;
                m_dSafetyInventoryLevel = oPOSuggestion.SafetyQuantity;
                m_dMaximumInventoryLevel = oPOSuggestion.MaximumQuantity;
                m_dMinimumOrderQuantity = oPOSuggestion.MinOrderQuantity;
                m_sBuyerId = oPOSuggestion.BuyerId;
                m_sClassId = oPOSuggestion.ClassId;
                m_sPurchaseUnitOfMeasure = oPOSuggestion.PurchaseUnitOfMeasure;
                m_sVendorName = oPOSuggestion.VendorName;
                m_sVendorId = oPOSuggestion.VendorId;
                m_iVendorNum = oPOSuggestion.VendorNum;
                m_iRequiredLeadTime = oPOSuggestion.RequiredLeadTime;
                m_bDropShip = oPOSuggestion.DropShip;
                m_dStartingInventoryOnHand = 0;
            }
        }

        #endregion

        #region Methods

        // Requisitions Without POs - Critical
        public List<HSPartInfoTimeLine> CheckForRequisitionsNeedingPO()
        {
            if (HasPOs == false)
            {
                DateTime dtCurrentDate = DateTime.Now;
                int iCurrentDayOfWeek = (int)dtCurrentDate.DayOfWeek;
                DateTime dtEndOfWeek = dtCurrentDate.AddDays(6 - iCurrentDayOfWeek);
                DateTime dtEndOfNextWeek = dtEndOfWeek.AddDays(7);
                bool bFirstCriticalDate = true;
               
                foreach (HSPartInfoTimeLine oTmpPartInfo in PartInfoTimeLineData)
                {
                    if (string.Compare(oTmpPartInfo.ActionType, "REQUISITION", true) == 0)
                    {
                        // we would need to order this the number of days required for lead time before this event happens
                        DateTime dtAdjustedForLeadTime = DefaultProductionCalendar.AddBusinessDaysToDate(oTmpPartInfo.DateOfChange, -RequiredLeadTime);
                        if ((bFirstCriticalDate == true) || (dtAdjustedForLeadTime < m_dtCriticalDate))
                        {
                            bFirstCriticalDate = false;
                            m_dtCriticalDate = dtAdjustedForLeadTime;
                        }

                        // check to see when this event falls
                        PartInfoTimeLineStatus oStatus = PartInfoTimeLineStatus.UNKNOWN;
                        if (dtAdjustedForLeadTime <= dtCurrentDate)
                        {
                            oStatus = PartInfoTimeLineStatus.PAST_DUE;
                        }
                        else if (dtAdjustedForLeadTime < dtEndOfWeek)
                        {
                            oStatus = PartInfoTimeLineStatus.THIS_WEEK;
                        }
                        else if (dtAdjustedForLeadTime < dtEndOfNextWeek)
                        {
                            oStatus = PartInfoTimeLineStatus.NEXT_WEEK;
                        }
                        else
                        {
                            oStatus = PartInfoTimeLineStatus.FUTURE;
                        }
                        if (Status > oStatus)
                        {
                            Status = oStatus;
                        }

                        m_oRequisitionsNeedingPOs.Add(oTmpPartInfo);
                        Processed = true;
                    }
                }
            }
            return m_oRequisitionsNeedingPOs;
        }

        // Critical Needs
        public List<HSPartInfoTimeLine> CheckForOrdersAndJobsThatAreCritical()
        {
            List<HSPartInfoTimeLine> oPotentialCandidate = new List<HSPartInfoTimeLine>();

            DateTime dtCurrentDate = DateTime.Now;
            int iCurrentDayOfWeek = (int)dtCurrentDate.DayOfWeek;
            DateTime dtEndOfWeek = dtCurrentDate.AddDays(6 - iCurrentDayOfWeek);
            DateTime dtEndOfNextWeek = dtEndOfWeek.AddDays(7);

            bool bPartBelowRequiredLevel = false;
            bool bFirstCriticalDate = true;
            foreach (HSPartInfoTimeLine oTmpPartInfo in PartInfoTimeLineData)
            {
                if ((oTmpPartInfo.CurrentInventoryLevel < 0) && ((string.Compare(oTmpPartInfo.ActionType, "ORDER", true) == 0) || (string.Compare(oTmpPartInfo.ActionType, "JOB", true) == 0)))
                {
                    // we would need to order this the number of days required for lead time before this event happens
                    DateTime dtAdjustedForLeadTime = DefaultProductionCalendar.AddBusinessDaysToDate(oTmpPartInfo.DateOfChange, -RequiredLeadTime);

                    bPartBelowRequiredLevel = true;
                    if ((bFirstCriticalDate == true) || (dtAdjustedForLeadTime < m_dtCriticalDate))
                    {
                        bFirstCriticalDate = false;
                        m_dtCriticalDate = dtAdjustedForLeadTime;
                    }

                    // check to see when this event falls
                    PartInfoTimeLineStatus oStatus = PartInfoTimeLineStatus.UNKNOWN;
                    if (dtAdjustedForLeadTime <= dtCurrentDate)
                    {
                        oStatus = PartInfoTimeLineStatus.PAST_DUE;
                    }
                    else if (dtAdjustedForLeadTime < dtEndOfWeek)
                    {
                        oStatus = PartInfoTimeLineStatus.THIS_WEEK;
                    }
                    else if (dtAdjustedForLeadTime < dtEndOfNextWeek)
                    {
                        oStatus = PartInfoTimeLineStatus.NEXT_WEEK;
                    }
                    else
                    {
                        oStatus = PartInfoTimeLineStatus.FUTURE;
                    }
                    if (Status > oStatus)
                    {
                        Status = oStatus;
                    }

                    oPotentialCandidate.Add(oTmpPartInfo);
                }
                if (bPartBelowRequiredLevel == true)
                {
                    // we went below the required inventory level due to a sale order or job so 
                    // we now need to check when a PO will come in and provide inventory
                    if ((oTmpPartInfo.CurrentInventoryLevel >= 0) && (string.Compare(oTmpPartInfo.ActionType, "PO", true) == 0))
                    {
                        bPartBelowRequiredLevel = false;
                        bFirstCriticalDate = true;
                        Status = PartInfoTimeLineStatus.UNKNOWN;
                        oPotentialCandidate.Clear();
                    }
                }
            }
            if ((oPotentialCandidate.Count > 0) && (bPartBelowRequiredLevel == true))
            {
                // we have run out of inventory
                m_oCriticalOrderAndJobs.AddRange(oPotentialCandidate);
                Processed = true;
            }
            return m_oCriticalOrderAndJobs;
        }

        public List<HSPartInfoTimeLine> CheckForInventoryLevelsThatAreCritical()
        {
            List<HSPartInfoTimeLine> oPotentialCandidate = new List<HSPartInfoTimeLine>();

            DateTime dtCurrentDate = DateTime.Now;
            int iCurrentDayOfWeek = (int)dtCurrentDate.DayOfWeek;
            DateTime dtEndOfWeek = dtCurrentDate.AddDays(6 - iCurrentDayOfWeek);
            DateTime dtEndOfNextWeek = dtEndOfWeek.AddDays(7);

            bool bPartBelowRequiredLevel = false;
            bool bFirstCriticalDate = true;
            foreach (HSPartInfoTimeLine oTmpPartInfo in PartInfoTimeLineData)
            {
                if (oTmpPartInfo.CurrentInventoryLevel < 0)
                {
                    // we would need to order this the number of days required for lead time before this event happens
                    DateTime dtAdjustedForLeadTime = DefaultProductionCalendar.AddBusinessDaysToDate(oTmpPartInfo.DateOfChange, -RequiredLeadTime);

                    // now we check the lead time to see if this event would happen before we have a chance to place a PO

                    bPartBelowRequiredLevel = true;
                    // something is driving the inventory level negative
                    if ((bFirstCriticalDate == true) || (dtAdjustedForLeadTime < m_dtCriticalDate))
                    {
                        bFirstCriticalDate = false;
                        m_dtCriticalDate = dtAdjustedForLeadTime;
                    }

                    // check to see when this event falls
                    PartInfoTimeLineStatus oStatus = PartInfoTimeLineStatus.UNKNOWN;
                    if (dtAdjustedForLeadTime <= dtCurrentDate)
                    {
                        oStatus = PartInfoTimeLineStatus.PAST_DUE;
                    }
                    else if (dtAdjustedForLeadTime < dtEndOfWeek)
                    {
                        oStatus = PartInfoTimeLineStatus.THIS_WEEK;
                    }
                    else if (dtAdjustedForLeadTime < dtEndOfNextWeek)
                    {
                        oStatus = PartInfoTimeLineStatus.NEXT_WEEK;
                    }
                    else
                    {
                        oStatus = PartInfoTimeLineStatus.FUTURE;
                    }
                    if (Status > oStatus)
                    {
                        Status = oStatus;
                    }

                    oPotentialCandidate.Add(oTmpPartInfo);
                }
                if (bPartBelowRequiredLevel == true)
                {
                    // we went below the required inventory level due to a sale order or job so 
                    // we now need to check when a PO will come in and provide inventory
                    if ((oTmpPartInfo.CurrentInventoryLevel >= 0) && (string.Compare(oTmpPartInfo.ActionType, "PO", true) == 0))
                    {
                        bPartBelowRequiredLevel = false;
                        bFirstCriticalDate = true;
                        Status = PartInfoTimeLineStatus.UNKNOWN;
                        oPotentialCandidate.Clear();
                    }
                }
            }

            if ((oPotentialCandidate.Count > 0) && (bPartBelowRequiredLevel == true))
            {
                // we have run out of inventory
                m_oCriticalInventoryLevels.AddRange(oPotentialCandidate);
                Processed = true;
            }
            return m_oCriticalInventoryLevels;
        }

        // Inventory Level Problem - Below Safety Or Min
        public List<HSPartInfoTimeLine> CheckForInventoryBelowSafetyOrMin()
        {
            List<HSPartInfoTimeLine> oPotentialCandidate = new List<HSPartInfoTimeLine>();

            DateTime dtCurrentDate = DateTime.Now;
            int iCurrentDayOfWeek = (int)dtCurrentDate.DayOfWeek;
            DateTime dtEndOfWeek = dtCurrentDate.AddDays(6 - iCurrentDayOfWeek);
            DateTime dtEndOfNextWeek = dtEndOfWeek.AddDays(7);

            bool bPartBelowRequiredLevel = false;
            bool bFirstCriticalDate = true;
            foreach (HSPartInfoTimeLine oTmpPartInfo in PartInfoTimeLineData)
            {
                if ((oTmpPartInfo.CurrentInventoryLevel < 0) || (BelowSafety(oTmpPartInfo.CurrentInventoryLevel) == true) || (BelowMinimum(oTmpPartInfo.CurrentInventoryLevel) == true))
                {
                    // we would need to order this the number of days required for lead time before this event happens
                    DateTime dtAdjustedForLeadTime = DefaultProductionCalendar.AddBusinessDaysToDate(oTmpPartInfo.DateOfChange, -RequiredLeadTime);

                    bPartBelowRequiredLevel = true;
                    if ((bFirstCriticalDate == true) || (dtAdjustedForLeadTime < m_dtCriticalDate))
                    {
                        bFirstCriticalDate = false;
                        m_dtCriticalDate = dtAdjustedForLeadTime;
                    }

                    // check to see when this event falls
                    PartInfoTimeLineStatus oStatus = PartInfoTimeLineStatus.UNKNOWN;
                    if (dtAdjustedForLeadTime <= dtCurrentDate)
                    {
                        oStatus = PartInfoTimeLineStatus.PAST_DUE;
                    }
                    else if (dtAdjustedForLeadTime < dtEndOfWeek)
                    {
                        oStatus = PartInfoTimeLineStatus.THIS_WEEK;
                    }
                    else if (dtAdjustedForLeadTime < dtEndOfNextWeek)
                    {
                        oStatus = PartInfoTimeLineStatus.NEXT_WEEK;
                    }
                    else
                    {
                        oStatus = PartInfoTimeLineStatus.FUTURE;
                    }
                    if (Status > oStatus)
                    {
                        Status = oStatus;
                    }

                    oPotentialCandidate.Add(oTmpPartInfo);
                }
                if (bPartBelowRequiredLevel == true)
                {
                    // we went below the required inventory level due to a sale order or job so 
                    // we now need to check when a PO will come in and provide inventory
                    if ((oTmpPartInfo.CurrentInventoryLevel >= 0) && (string.Compare(oTmpPartInfo.ActionType, "PO", true) == 0))
                    {
                        bPartBelowRequiredLevel = false;
                        bFirstCriticalDate = true;
                        Status = PartInfoTimeLineStatus.UNKNOWN;
                        oPotentialCandidate.Clear();
                    }
                }
            }
            if ((oPotentialCandidate.Count > 0) && (bPartBelowRequiredLevel == true))
            {
                // we are below safety or min
                m_oInventoryBelowSafetyOrMin.AddRange(oPotentialCandidate);
                Processed = true;
            }
            return m_oInventoryBelowSafetyOrMin;
        }

        public List<HSPartInfoTimeLine> CheckForExcessiveWaits(PartTimeDomainInfo oPartInfo)
        {
            //if (oPartInfo.NegativeForMoreThanOneMonth == true)
            //{
            //    m_oExcessiveWaits.Add(oPartInfo);
            //}
            return m_oExcessiveWaits;
        }

        // Delayed Orders
        public List<HSPartInfoTimeLine> CheckForDelayedOrders()
        {
            bool bOrderCausedNegativeInventory = false;
            if (HasPOs == true)
            {
                DateTime dtCurrentDate = DateTime.Now;
                int iCurrentDayOfWeek = (int)dtCurrentDate.DayOfWeek;
                DateTime dtEndOfWeek = dtCurrentDate.AddDays(6 - iCurrentDayOfWeek);
                DateTime dtEndOfNextWeek = dtEndOfWeek.AddDays(7);

                bool bFirstCriticalDate = true;
                foreach (HSPartInfoTimeLine oTmpPartInfo in PartInfoTimeLineData)
                {
                    if ((oTmpPartInfo.CurrentInventoryLevel < 0) && (string.Compare(oTmpPartInfo.ActionType, "ORDER", true) == 0))
                    {
                        bOrderCausedNegativeInventory = true;
                        if ((bFirstCriticalDate == true) || (oTmpPartInfo.DateOfChange < m_dtCriticalDate))
                        {
                            bFirstCriticalDate = false;
                            m_dtCriticalDate = oTmpPartInfo.DateOfChange;
                        }

                        m_oDelayedOrders.Add(oTmpPartInfo);
                        Processed = true;
                    }

                    // check to see when this event falls
                    PartInfoTimeLineStatus oStatus = PartInfoTimeLineStatus.UNKNOWN;
                    if (m_dtCriticalDate <= dtCurrentDate)
                    {
                        oStatus = PartInfoTimeLineStatus.PAST_DUE;
                    }
                    else if (m_dtCriticalDate < dtEndOfWeek)
                    {
                        oStatus = PartInfoTimeLineStatus.THIS_WEEK;
                    }
                    else if (m_dtCriticalDate < dtEndOfNextWeek)
                    {
                        oStatus = PartInfoTimeLineStatus.NEXT_WEEK;
                    }
                    else
                    {
                        oStatus = PartInfoTimeLineStatus.FUTURE;
                    }
                    if (Status > oStatus)
                    {
                        Status = oStatus;
                    }

                    if (bOrderCausedNegativeInventory == true)
                    {
                        // we went negative in the past due to a sale order so 
                        // we now need to check when a PO will come in and provide inventory
                        if ((oTmpPartInfo.CurrentInventoryLevel >= 0) && (string.Compare(oTmpPartInfo.ActionType, "PO", true) == 0))
                        {
                            m_oDelayedOrders.Add(oTmpPartInfo);
                            break;
                        }
                    }
                }
            }
            return m_oDelayedOrders;
        }

        // Delayed Jobs
        public List<HSPartInfoTimeLine> CheckForDelayedJobs()
        {
            bool bJobCausedNegativeInventory = false;
            if (HasPOs == true)
            {
                DateTime dtCurrentDate = DateTime.Now;
                int iCurrentDayOfWeek = (int)dtCurrentDate.DayOfWeek;
                DateTime dtEndOfWeek = dtCurrentDate.AddDays(6 - iCurrentDayOfWeek);
                DateTime dtEndOfNextWeek = dtEndOfWeek.AddDays(7);

                bool bFirstCriticalDate = true;
                foreach (HSPartInfoTimeLine oTmpPartInfo in PartInfoTimeLineData)
                {
                    if ((oTmpPartInfo.CurrentInventoryLevel < 0) && (string.Compare(oTmpPartInfo.ActionType, "JOB", true) == 0))
                    {
                        bJobCausedNegativeInventory = true;
                        if ((bFirstCriticalDate == true) || (oTmpPartInfo.DateOfChange < m_dtCriticalDate))
                        {
                            bFirstCriticalDate = false;
                            m_dtCriticalDate = oTmpPartInfo.DateOfChange;
                        }

                        m_oDelayedJobs.Add(oTmpPartInfo);
                        Processed = true;
                    }

                    // check to see when this event falls
                    PartInfoTimeLineStatus oStatus = PartInfoTimeLineStatus.UNKNOWN;
                    if (m_dtCriticalDate <= dtCurrentDate)
                    {
                        oStatus = PartInfoTimeLineStatus.PAST_DUE;
                    }
                    else if (m_dtCriticalDate < dtEndOfWeek)
                    {
                        oStatus = PartInfoTimeLineStatus.THIS_WEEK;
                    }
                    else if (m_dtCriticalDate < dtEndOfNextWeek)
                    {
                        oStatus = PartInfoTimeLineStatus.NEXT_WEEK;
                    }
                    else
                    {
                        oStatus = PartInfoTimeLineStatus.FUTURE;
                    }
                    if (Status > oStatus)
                    {
                        Status = oStatus;
                    }

                    if (bJobCausedNegativeInventory == true)
                    {
                        // we went negative in the past due to a sale order so 
                        // we now need to check when a PO will come in and provide inventory
                        if ((oTmpPartInfo.CurrentInventoryLevel >= 0) && (string.Compare(oTmpPartInfo.ActionType, "PO", true) == 0))
                        {
                            m_oDelayedJobs.Add(oTmpPartInfo);
                            break;
                        }
                    }
                }
            }
            return m_oDelayedJobs;
        }

        // Delayed Inventory
        public List<HSPartInfoTimeLine> CheckForDelayedInventoryLevels()
        {
            bool bPartBelowRequiredLevel = false;
            if (HasPOs == true)
            {
                DateTime dtCurrentDate = DateTime.Now;
                int iCurrentDayOfWeek = (int)dtCurrentDate.DayOfWeek;
                DateTime dtEndOfWeek = dtCurrentDate.AddDays(6 - iCurrentDayOfWeek);
                DateTime dtEndOfNextWeek = dtEndOfWeek.AddDays(7);

                bool bFirstCriticalDate = true;
                foreach (HSPartInfoTimeLine oTmpPartInfo in PartInfoTimeLineData)
                {
                    if ((oTmpPartInfo.CurrentInventoryLevel < 0) || (BelowSafety(oTmpPartInfo.CurrentInventoryLevel) == true) || (BelowMinimum(oTmpPartInfo.CurrentInventoryLevel) == true))
                    {
                        // we would need to order this the number of days required for lead time before this event happens
                        DateTime dtAdjustedForLeadTime = DefaultProductionCalendar.AddBusinessDaysToDate(oTmpPartInfo.DateOfChange, -RequiredLeadTime);

                        bPartBelowRequiredLevel = true;
                        if ((bFirstCriticalDate == true) || (dtAdjustedForLeadTime < m_dtCriticalDate))
                        {
                            bFirstCriticalDate = false;
                            m_dtCriticalDate = dtAdjustedForLeadTime;
                        }

                        // check to see when this event falls
                        PartInfoTimeLineStatus oStatus = PartInfoTimeLineStatus.UNKNOWN;
                        if (dtAdjustedForLeadTime <= dtCurrentDate)
                        {
                            oStatus = PartInfoTimeLineStatus.PAST_DUE;
                        }
                        else if (dtAdjustedForLeadTime < dtEndOfWeek)
                        {
                            oStatus = PartInfoTimeLineStatus.THIS_WEEK;
                        }
                        else if (dtAdjustedForLeadTime < dtEndOfNextWeek)
                        {
                            oStatus = PartInfoTimeLineStatus.NEXT_WEEK;
                        }
                        else
                        {
                            oStatus = PartInfoTimeLineStatus.FUTURE;
                        }
                        if (Status > oStatus)
                        {
                            Status = oStatus;
                        }

                        m_oInventoryLevelsNowDelayed.Add(oTmpPartInfo);
                        Processed = true;
                    }
                    if (bPartBelowRequiredLevel == true)
                    {
                        // we went below the required inventory level due to a sale order or job so 
                        // we now need to check when a PO will come in and provide inventory
                        if ((oTmpPartInfo.CurrentInventoryLevel >= 0) && (BelowSafety(oTmpPartInfo.CurrentInventoryLevel) == false) && (BelowMinimum(oTmpPartInfo.CurrentInventoryLevel) == false) && (string.Compare(oTmpPartInfo.ActionType, "PO", true) == 0))
                        {
                            m_oInventoryLevelsNowDelayed.Add(oTmpPartInfo);
                            break;
                        }
                    }
                }
            }
            return m_oInventoryLevelsNowDelayed;
        }

        public bool BelowMinimum(decimal dInventoryLevel)
        {
            bool bBelowMin = false;
            if ((m_dMinimumInventoryLevel != 0) && (dInventoryLevel < m_dMinimumInventoryLevel))
            {
                bBelowMin = true;
            }
            return bBelowMin;
        }

        public bool BelowSafety(decimal dInventoryLevel)
        {
            bool bBelowSafety = false;
            if ((m_dMinimumInventoryLevel != 0) || (m_dSafetyInventoryLevel != 0))
            {
                if (dInventoryLevel < m_dMinimumInventoryLevel + m_dSafetyInventoryLevel)
                {
                    bBelowSafety = true;
                }
            }
            return bBelowSafety;
        }

        public bool AboveMax(decimal dInventoryLevel)
        {
            bool bAboveMax = false;
            if ((m_dMaximumInventoryLevel != 0) && (dInventoryLevel > m_dMaximumInventoryLevel))
            {
                bAboveMax = true;
            }
            return bAboveMax;
        }

        #endregion

        #region Properties

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

        public decimal MinimumInventoryLevel
        {
            get { return m_dMinimumInventoryLevel; }
            set { m_dMinimumInventoryLevel = value; }
        }

        public decimal SafetyInventoryLevel
        {
            get { return m_dSafetyInventoryLevel; }
            set { m_dSafetyInventoryLevel = value; }
        }

        public decimal MaximumInventoryLevel
        {
            get { return m_dMaximumInventoryLevel; }
            set { m_dMaximumInventoryLevel = value; }
        }

        public decimal MinimumOrderQuantity
        {
            get { return m_dMinimumOrderQuantity; }
            set { m_dMinimumOrderQuantity = value; }
        }

        public string BuyerId
        {
            get { return m_sBuyerId; }
            set { m_sBuyerId = value; }
        }

        public string ClassId
        {
            get { return m_sClassId; }
            set { m_sClassId = value; }
        }

        public string PurchaseUnitOfMeasure
        {
            get { return m_sPurchaseUnitOfMeasure; }
            set { m_sPurchaseUnitOfMeasure = value; }
        }

        public string VendorName
        {
            get { return m_sVendorName; }
            set { m_sVendorName = value; }
        }

        public string VendorId
        {
            get { return m_sVendorId; }
            set { m_sVendorId = value; }
        }

        public int VendorNum
        {
            get { return m_iVendorNum; }
            set { m_iVendorNum = value; }
        }

        public int RequiredLeadTime
        {
            get { return m_iRequiredLeadTime; }
            set { m_iRequiredLeadTime = value; }
        }

        public bool DropShip
        {
            get { return m_bDropShip; }
            set { m_bDropShip = value; }
        }

        public decimal StartingInventoryOnHand
        {
            get { return m_dStartingInventoryOnHand; }
            set { m_dStartingInventoryOnHand = value; }
        }

        public DateTime StartDate
        {
            get { return m_dtStartDate; }
            set { m_dtStartDate = value; }
        }

        public DateTime EndDate
        {
            get { return m_dtEndDate; }
            set { m_dtEndDate = value; }
        }

        public decimal MinQuantity
        {
            get { return m_dMinQuantity; }
            set { m_dMinQuantity = value; }
        }

        public decimal MaxQuantity
        {
            get { return m_dMaxQuantity; }
            set { m_dMaxQuantity = value; }
        }


        public decimal EndingQuantity
        {
            get { return m_dEndingQuantity; }
            set { m_dEndingQuantity = value; }
        }

        public bool HasPOs
        {
            get { return m_oPOsForPart.Count > 0; }
        }

        public bool HasRequisitions
        {
            get { return m_oRequisitionsForPart.Count > 0; }
        }

        public bool HasJobss
        {
            get { return m_oJobsForPart.Count > 0; }
        }

        public bool HasOrders
        {
            get { return m_oOrdersForPart.Count > 0; }
        }

        public bool Processed
        {
            get { return m_bProcessed; }
            set { m_bProcessed = value; }
        }

        public PartInfoTimeLineStatus Status
        {
            get { return m_oStatus; }
            set { m_oStatus = value; }
        }

        public List<HSPartInfoTimeLine> CriticalOrdersAndJobs
        {
            get { return m_oCriticalOrderAndJobs; }
        }

        public List<HSPartInfoTimeLine> FutureOrderAndJobNeeds
        {
            get { return m_oFutureOrderAndJobNeeds; }
        }

        public List<HSPartInfoTimeLine> RequisitionsNeedingPOs
        {
            get { return m_oRequisitionsNeedingPOs; }
        }

        public List<HSPartInfoTimeLine> CriticalInventoryLevels
        {
            get { return m_oCriticalInventoryLevels; }
        }

        public List<HSPartInfoTimeLine> FutureInventoryLevelNeeds
        {
            get { return m_oInventoryBelowSafetyOrMin; }
        }

        public List<HSPartInfoTimeLine> ExcessiveWaits
        {
            get { return m_oExcessiveWaits; }
        }

        public List<HSPartInfoTimeLine> DelayedOrders
        {
            get { return m_oDelayedOrders; }
        }

        public List<HSPartInfoTimeLine> DelayedJobs
        {
            get { return m_oDelayedJobs; }
        }

        public List<HSPartInfoTimeLine> InventoryLevelsNowDelayed
        {
            get { return m_oInventoryLevelsNowDelayed; }
        }

        public DateTime CriticalDate
        {
            get { return m_dtCriticalDate; }
        }


        public List<HSPOSuggestion> POSuggestions
        {
            get { return m_oPOSuggestionsForPart; }
            set { m_oPOSuggestionsForPart = value; }
        }

        public List<HSDemandForPartFromJob> Jobs
        {
            get { return m_oJobsForPart; }
            set { m_oJobsForPart = value; }
        }

        public List<HSDemandForPartFromOrder> Orders
        {
            get { return m_oOrdersForPart; }
            set { m_oOrdersForPart = value; }
        }

        public List<HSDemandForPartsFromRequisition> Requisitions
        {
            get { return m_oRequisitionsForPart; }
            set { m_oRequisitionsForPart = value; }
        }

        public List<HSSourceForPartFromPO> POs
        {
            get { return m_oPOsForPart; }
            set { m_oPOsForPart = value; }
        }
        public List<HSOpenJob> JobsSupplyingParts
        {
            get { return m_oJobsSupplyingPart; }
            set { m_oJobsSupplyingPart = value; }
        }

        public List<HSPartsOnHand> PartsOnHand
        {
            get { return m_oPartsOnHand; }
            set { m_oPartsOnHand = value; }
        }

        public List<HSPartInfoTimeLine> PartInfoTimeLineData
        {
            get { return m_oPartInfoTimeLineData; }
            set { m_oPartInfoTimeLineData = value; }
        }

        public int ReqNum
        {
            get { return m_iReqNum; }
            set { m_iReqNum = value; }
        }

        public string RequestedBy
        {
            get { return m_sRequestedBy; }
            set { m_sRequestedBy = value; }
        }

        public DateTime DueDate
        {
            get { return m_dtDueDate; }
            set { m_dtDueDate = value; }
        }

        public DateTime OrderBy
        {
            get { return m_dtOrderBy; }
            set { m_dtOrderBy = value; }
        }

        public DateTime DateOfEvent
        {
            get { return m_dtDateOfEvent; }
            set { m_dtDateOfEvent = value; }
        }

        public DateTime EarliestDeliveryDate
        {
            get { return m_dtEarliestDeliveryDate; }
            set { m_dtEarliestDeliveryDate = value; }
        }

        public string PONum
        {
            get { return m_sPONum; }
            set { m_sPONum = value; }
        }

        public DateTime ExpectedDeliveryDate
        {
            get { return m_dtExpectedDeliveryDate; }
            set { m_dtExpectedDeliveryDate = value; }
        }

        public string Comments
        {
            get { return m_sComments; }
            set { m_sComments = value; }
        }

        public bool NegativeForMoreThanOneMonth
        {
            get { return m_bNegativeForMoreThanOneMonth; }
            set { m_bNegativeForMoreThanOneMonth = value; }
        }

        public List<string> OriginalRecord
        {
            get { return m_oOriginalRecord; }
            set { m_oOriginalRecord = value; }
        }


        public string Company
        {
            get { return m_sCompany; }
            set
            {
                m_sCompany = value;
                m_sDefaultProductionCalendarName = "";
                if (string.Compare(m_sCompany, CompanyConfiguration.SPINCRAFT_MA_COMPANY_ID, true) == 0)
                {
                    m_sDefaultProductionCalendarName = ProductionCalendarCollection.SPINCRAFT_MA_PRODUCTION_CALENDAR;
                }
                else if (string.Compare(m_sCompany, CompanyConfiguration.SPINCRAFT_WI_COMPANY_ID, true) == 0)
                {
                    m_sDefaultProductionCalendarName = ProductionCalendarCollection.SPINCRAFT_WI_PRODUCTION_CALENDAR;
                }
                else if (string.Compare(m_sCompany, CompanyConfiguration.SPINCRAFT_UK_COMPANY_ID, true) == 0)
                {
                    m_sDefaultProductionCalendarName = ProductionCalendarCollection.SPINCRAFT_UK_PRODUCTION_CALENDAR;
                }
                else
                {
                    throw new Exception("Invalid company name: " + value);
                }
            }
        }

        public ProductionCalendar DefaultProductionCalendar
        {
            get
            {
                if (m_oDefaultProductionCalendar == null)
                {
                    m_oDefaultProductionCalendar = ProductionCalendarCollection.GetByCalendarId(m_sDefaultProductionCalendarName);
                }
                return m_oDefaultProductionCalendar;
            }
        }
        #endregion

        #region Data Members
        private string m_sCompany;
        private string m_sDefaultProductionCalendarName;
        private ProductionCalendar m_oDefaultProductionCalendar;

        private string m_sPartNum;
        private string m_sPartDescription;
        private decimal m_dMinimumInventoryLevel;
        private decimal m_dSafetyInventoryLevel;
        private decimal m_dMaximumInventoryLevel;
        private decimal m_dMinimumOrderQuantity;
        private string m_sBuyerId;
        private string m_sClassId;
        private string m_sPurchaseUnitOfMeasure;
        private string m_sVendorName;
        private string m_sVendorId;
        private int m_iVendorNum;
        private int m_iRequiredLeadTime;
        private bool m_bDropShip;
        private decimal m_dStartingInventoryOnHand;

        private DateTime m_dtStartDate;
        private DateTime m_dtEndDate;
        private decimal m_dMinQuantity;
        private decimal m_dMaxQuantity;
        private decimal m_dEndingQuantity;

        private bool m_bProcessed;
        private PartInfoTimeLineStatus m_oStatus = PartInfoTimeLineStatus.UNKNOWN;

        private List<HSPartInfoTimeLine> m_oCriticalOrderAndJobs = new List<HSPartInfoTimeLine>();
        private List<HSPartInfoTimeLine> m_oFutureOrderAndJobNeeds = new List<HSPartInfoTimeLine>();
        private List<HSPartInfoTimeLine> m_oRequisitionsNeedingPOs = new List<HSPartInfoTimeLine>();
        private List<HSPartInfoTimeLine> m_oCriticalInventoryLevels = new List<HSPartInfoTimeLine>();
        private List<HSPartInfoTimeLine> m_oInventoryBelowSafetyOrMin = new List<HSPartInfoTimeLine>();
        private List<HSPartInfoTimeLine> m_oExcessiveWaits = new List<HSPartInfoTimeLine>();
        private List<HSPartInfoTimeLine> m_oDelayedOrders = new List<HSPartInfoTimeLine>();
        private List<HSPartInfoTimeLine> m_oDelayedJobs = new List<HSPartInfoTimeLine>();
        private List<HSPartInfoTimeLine> m_oInventoryLevelsNowDelayed = new List<HSPartInfoTimeLine>();
        private DateTime m_dtCriticalDate;

        private List<HSPOSuggestion> m_oPOSuggestionsForPart = new List<HSPOSuggestion>();
        private List<HSDemandForPartFromJob> m_oJobsForPart = new List<HSDemandForPartFromJob>();
        private List<HSDemandForPartFromOrder> m_oOrdersForPart = new List<HSDemandForPartFromOrder>();
        private List<HSDemandForPartsFromRequisition> m_oRequisitionsForPart = new List<HSDemandForPartsFromRequisition>();
        private List<HSSourceForPartFromPO> m_oPOsForPart = new List<HSSourceForPartFromPO>();
        private List<HSOpenJob> m_oJobsSupplyingPart = new List<HSOpenJob>();
        private List<HSPartsOnHand> m_oPartsOnHand = new List<HSPartsOnHand>();
        private List<HSPartInfoTimeLine> m_oPartInfoTimeLineData = new List<HSPartInfoTimeLine>();

        private int m_iReqNum;
        private string m_sRequestedBy;
        private DateTime m_dtDueDate;
        private DateTime m_dtOrderBy;
        private DateTime m_dtDateOfEvent;
        private DateTime m_dtEarliestDeliveryDate;
        private string m_sPONum;
        private DateTime m_dtExpectedDeliveryDate;
        private string m_sComments;
        private bool m_bNegativeForMoreThanOneMonth = false;
        private List<string> m_oOriginalRecord = new List<string>();

        #endregion
    }
}
