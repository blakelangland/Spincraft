using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using System.IO;

using Ice.Core;
using Erp.BO;
using Erp.Proxy.BO;
using Ice.Lib.Framework;
using Ice.Adapters;
using Ice.BO;
using Erp.Adapters;
using Ice.Tablesets;

using SpreadsheetLight;

namespace HorizonScientific
{
    public class ProductionCalendarCollection
    {
        #region Constructors
        public ProductionCalendarCollection()
        {

        }
        #endregion

        #region Methods
        public static bool Initialize(Session oSession, DateTime dtFirstProductionDay, DateTime dtLastProductionDay)
        {
            bool bSuccess = true;

            //****Set a parameter Value***** 
            Ice.Proxy.BO.DynamicQueryImpl oDynamicQuery = WCFServiceSupport.CreateImpl<Ice.Proxy.BO.DynamicQueryImpl>(oSession, Ice.Proxy.BO.DynamicQueryImpl.UriPath);
            QueryExecutionDataSet oQueryExecutionDataSet = oDynamicQuery.GetQueryExecutionParametersByID(BAQConstants.QUERY_PRODUCTION_CALENDAR);
            DateTime dtToday = DateTime.Today;
            foreach (DataRow oParameter in oQueryExecutionDataSet.ExecutionParameter)
            {
                if (string.Compare(oParameter["ParameterID"].ToString(), "StartDate") == 0)
                {
                    oParameter["ParameterValue"] = dtFirstProductionDay;
                }
                if (string.Compare(oParameter["ParameterID"].ToString(), "EndDate") == 0)
                {
                    oParameter["ParameterValue"] = dtLastProductionDay;
                }

            }
            oQueryExecutionDataSet.AcceptChanges();
            DataSet oDataSet = oDynamicQuery.ExecuteByID(BAQConstants.QUERY_PRODUCTION_CALENDAR, oQueryExecutionDataSet);

            g_oProductionCalendars.Clear();
            // this query returns all production calendars as well as any holidays
            // so when processing this we will just add one production calendar
            // and then have each production calendar process this data to add
            // their specific holdays
            List<ProductionCalendar> oHolidays = new List<ProductionCalendar>();
            foreach (DataRow oRow in oDataSet.Tables[0].Rows)
            {
                ProductionCalendar oTmpProductionCalendar = new ProductionCalendar(oRow);
                oHolidays.Add(oTmpProductionCalendar);
                // check to see if we already have this production calendar in the list
                ProductionCalendar oExists = g_oProductionCalendars.FirstOrDefault(oItem => string.Compare(oItem.CalendarId, oTmpProductionCalendar.CalendarId, true) == 0);
                if (oExists == null)
                {
                    g_oProductionCalendars.Add(oTmpProductionCalendar);
                }
            }
            // now we will have each production calendar extract its set of holidays
            foreach (ProductionCalendar oProductionCalendar in g_oProductionCalendars)
            {
                oProductionCalendar.Initialize(oHolidays);
            }

            return bSuccess;
        }

        public static ProductionCalendar GetByCalendarId(string sCalendarId)
        {
            ProductionCalendar oTmp = g_oProductionCalendars.FirstOrDefault(oItem => string.Compare(oItem.CalendarId, sCalendarId, true) == 0);
            return oTmp;
        }
        #endregion

        #region Properties
        public static List<ProductionCalendar> ProductionCalendars
        {
            get { return g_oProductionCalendars; }
        }
        #endregion

        #region Data Members
        private static List<ProductionCalendar> g_oProductionCalendars = new List<ProductionCalendar>();

        public static string SPINCRAFT_MA_PRODUCTION_CALENDAR = @"51504";
        public static string SPINCRAFT_WI_PRODUCTION_CALENDAR = @"51503";
        public static string SPINCRAFT_UK_PRODUCTION_CALENDAR = @"51515";
        public static string SPINCRAFT_CA_PRODUCTION_CALENDAR = @"MCS";
        #endregion
    }

    // Production calendar
    public class ProductionCalendar
    {
        #region Constructors

        public ProductionCalendar(DataRow oDataRow)
        {
            if (oDataRow["ProdCal_CalendarID"] != DBNull.Value)
            {
                m_sCalendarId = (string)oDataRow["ProdCal_CalendarID"];
            }
            if (oDataRow["ProdCal_Description"] != DBNull.Value)
            {
                m_sDescription = (string)oDataRow["ProdCal_Description"];
            }
            // these are really the holdays for the prodcuton calendar
            // which we will process separately later
            if (oDataRow["ProdCalDay_ModifiedDay"] != DBNull.Value)
            {
                m_dtModifiedDay = (DateTime)oDataRow["ProdCalDay_ModifiedDay"];
            }
            if (oDataRow["ProdCalDay_WorkingDay"] != DBNull.Value)
            {
                m_bWorkingDay = (bool)oDataRow["ProdCalDay_WorkingDay"];
            }

            // we are going to read and set the hours in this collection in the proper order
            // starting with ProdCal_Hour001 and going to ProdCal_Hour168
            bool bTmp;
            // there are 168 hours in the work week
            for (int i = 1; i <= 168; i++)
            {
                string sColumnName = "ProdCal_Hour" + string.Format("{0:000}", i);
                if (oDataRow[sColumnName] != DBNull.Value)
                {
                    bTmp = (bool)oDataRow[sColumnName];
                    m_oWorkingHoursForWeek.Add(bTmp);
                }
            }

            // review the standard work hours in the week to build up the list of standard work days in week
            m_oStandardWorkDays.Clear();
            for (int iIndex = 0; iIndex < 168; iIndex++)
            {
                if (m_oWorkingHoursForWeek[iIndex] == true)
                {
                    // this is a working day in the week
                    if (iIndex < 24)
                    {
                        // this is Sunday
                        if (m_oStandardWorkDays.Contains(DayOfWeek.Sunday) == false)
                        {
                            m_oStandardWorkDays.Add(DayOfWeek.Sunday);
                        }
                        continue;
                    }
                    if (iIndex < 48)
                    {
                        // this is Monday
                        if (m_oStandardWorkDays.Contains(DayOfWeek.Monday) == false)
                        {
                            m_oStandardWorkDays.Add(DayOfWeek.Monday);
                        }
                        continue;
                    }
                    if (iIndex < 72)
                    {
                        // this is Tuesday
                        if (m_oStandardWorkDays.Contains(DayOfWeek.Tuesday) == false)
                        {
                            m_oStandardWorkDays.Add(DayOfWeek.Tuesday);
                        }
                        continue;
                    }
                    if (iIndex < 96)
                    {
                        // this is Wednesday
                        if (m_oStandardWorkDays.Contains(DayOfWeek.Wednesday) == false)
                        {
                            m_oStandardWorkDays.Add(DayOfWeek.Wednesday);
                        }
                        continue;
                    }
                    if (iIndex < 120)
                    {
                        // this is Thursday
                        if (m_oStandardWorkDays.Contains(DayOfWeek.Thursday) == false)
                        {
                            m_oStandardWorkDays.Add(DayOfWeek.Thursday);
                        }
                        continue;
                    }
                    if (iIndex < 144)
                    {
                        // this is Friday
                        if (m_oStandardWorkDays.Contains(DayOfWeek.Friday) == false)
                        {
                            m_oStandardWorkDays.Add(DayOfWeek.Friday);
                        }
                        continue;
                    }
                    if (iIndex < 168)
                    {
                        // this is Saturday
                        if (m_oStandardWorkDays.Contains(DayOfWeek.Saturday) == false)
                        {
                            m_oStandardWorkDays.Add(DayOfWeek.Saturday);
                        }
                        continue;
                    }
                }
            }
        }

        #endregion

        #region Methods

        public void Initialize(List<ProductionCalendar> oHolidays)
        {
            // pull out all items tied to this calendar id
            List<ProductionCalendar> oMyHolidays = oHolidays.Where(oItem => string.Compare(oItem.CalendarId, CalendarId, true) == 0).ToList();
            // now add these in as a list of date that this calendar has as holidays
            foreach (ProductionCalendar oTmp in oMyHolidays)
            {
                if ((oTmp.m_dtModifiedDay != DateTime.MinValue) && (oTmp.m_bWorkingDay == false))
                {
                    m_oHolidays.Add(oTmp.m_dtModifiedDay);
                }
            }
        }

        public bool IsWorkDay(DateTime dtDay)
        {
            bool bIsWorkDay = true;

            // see if the day is one of our standard work days in the week
            if (m_oStandardWorkDays.Contains(dtDay.DayOfWeek) == false)
            {
                bIsWorkDay = false;
            }

            if (bIsWorkDay == true)
            {
                // this is a normal work day for this calendar so now we need to check it it is a holiday
                if (m_oHolidays.Contains(dtDay) == true)
                {
                    // this falls on a holiday
                    bIsWorkDay = false;
                }
            }
            return bIsWorkDay;
        }

        public int WorkDaysInMonth(DateTime dtMonthOfInterest)
        {
            DateTime dtLastDayInMonth = new DateTime(dtMonthOfInterest.Year, dtMonthOfInterest.Month, DateTime.DaysInMonth(dtMonthOfInterest.Year, dtMonthOfInterest.Month));
            DateTime dtFirstDayOfMonth = new DateTime(dtMonthOfInterest.Year, dtMonthOfInterest.Month, 1);
            DateTime dttmpDate = dtFirstDayOfMonth;
            int iWorkDays = 0;
            while (dttmpDate.Month == dtLastDayInMonth.Month)
            {
                if (IsWorkDay(dttmpDate))
                {
                    iWorkDays++;
                }
                dttmpDate = dttmpDate.AddDays(1);
            }
            return iWorkDays;
        }

        public int WorkDaysInCurrentMonth()
        {
            DateTime dtToday = DateTime.Today;
            DateTime dtFirstDayOfMonth = new DateTime(dtToday.Year, dtToday.Month, 1);
            DateTime dttmpDate = dtFirstDayOfMonth;
            int iWorkDays = 0;
            while (dttmpDate.Month == dtToday.Month)
            {
                if (IsWorkDay(dttmpDate))
                {
                    iWorkDays++;
                }
                dttmpDate = dttmpDate.AddDays(1);
            }
            return iWorkDays;
        }

        public int WorkDaysFromStartOfMonth()
        {
            DateTime dtToday = DateTime.Today;
            DateTime dtFirstDayOfMonth = new DateTime(dtToday.Year, dtToday.Month, 1);
            DateTime dtTmpDate = dtFirstDayOfMonth;
            int iWorkDays = 0;
            while (dtTmpDate.Month == dtToday.Month && dtTmpDate.Day <= dtToday.Day)
            {
                if (IsWorkDay(dtTmpDate))
                {
                    iWorkDays++;
                }
                dtTmpDate = dtTmpDate.AddDays(1);
            }
            if (iWorkDays == 0)
            {
                // protect from divide by zero when its 
                // the first day of the month and this is a work day
                iWorkDays = 1;
            }
            return iWorkDays;
        }

        public DateTime FirstDayOfLastWeek()
        {
            DateTime dtToday = DateTime.Today;
            DayOfWeek iWeekDay = dtToday.DayOfWeek;
            DateTime dtFirstDayOfLastWeek = dtToday.AddDays(-1 * ((int)iWeekDay + 7));
            return dtFirstDayOfLastWeek;
        }

        public DateTime LastDayOfLastWeek()
        {
            DateTime dtToday = DateTime.Today;
            DayOfWeek iWeekDay = dtToday.DayOfWeek;
            int iDaysToAdd = (int)DayOfWeek.Saturday - (int)iWeekDay;

            DateTime dtLastDayOfLastWeek = dtToday.AddDays(iDaysToAdd - 7);
            return dtLastDayOfLastWeek;
        }

        public DateTime FirstDayOfMonth()
        {
            int iCurrentYear = DateTime.Today.Year;
            int iCurrentMonth = DateTime.Today.Month;
            DateTime dtFirstDayOfMonth = new DateTime(iCurrentYear, iCurrentMonth, 1);
            return dtFirstDayOfMonth;
        }

        public DateTime LastDayOfMonth()
        {
            int iCurrentYear = DateTime.Today.Year;
            int iCurrentMonth = DateTime.Today.Month;
            DateTime dtLastDayOfThisMonth = new DateTime(iCurrentYear, iCurrentMonth, 1).AddMonths(1).AddDays(-1);
            return dtLastDayOfThisMonth;
        }

        public DateTime AddBusinessDaysToDate(DateTime dtStartDate, int iBusinessDays)
        {
            DateTime dtEndDate = dtStartDate;
            if (iBusinessDays < 0)
            {
                // we are counting backward from the start date
                for (int i = iBusinessDays; i < 0; i++)
                {
                    // keep subtracting days until we have a business day
                    dtEndDate = dtEndDate.AddDays(-1);
                    while (IsWorkDay(dtEndDate) == false)
                    {
                        dtEndDate = dtEndDate.AddDays(-1);
                    }
                }
            }
            else
            {
                // we are counting forward from the start date
                for (int i = 0; i < iBusinessDays; i++)
                {
                    // keep adding days until we have a business day
                    dtEndDate = dtEndDate.AddDays(1);
                    while (IsWorkDay(dtEndDate) == false)
                    {
                        dtEndDate = dtEndDate.AddDays(1);
                    }
                }
            }
            return dtEndDate;
        }

        public int BusinessDaysBetweenDates(DateTime dtStartDate, DateTime dtEndDate)
        {
            int iBusinessDaysBetweenDates = 0;
            if (dtStartDate < dtEndDate)
            {
                // we are counting forward from the start date
                for (DateTime dtCurrentDate = dtStartDate; dtCurrentDate < dtEndDate;)
                {
                    // keep adding days until we have a business day
                    if (IsWorkDay(dtCurrentDate) == true)
                    {
                        iBusinessDaysBetweenDates++;
                    }
                    dtCurrentDate = dtCurrentDate.AddDays(1);
                }
            }
            return iBusinessDaysBetweenDates;
        }

        #endregion

        #region Properties

        public string CalendarId
        {
            get { return m_sCalendarId; }
        }
        public string Description
        {
            get { return m_sDescription; }
        }
        public List<DayOfWeek> StandardWorkDaysInWeek
        {
            get { return m_oStandardWorkDays; }
        }
        public List<DateTime> Holidays
        {
            get { return m_oHolidays; }
        }


        #endregion

        #region Data Members

        private string m_sCalendarId;
        private string m_sDescription;
        private List<bool> m_oWorkingHoursForWeek = new List<bool>();
        private List<DayOfWeek> m_oStandardWorkDays = new List<DayOfWeek>();
        private List<DateTime> m_oHolidays = new List<DateTime>();

        // these are really ignored for the production calendar
        // and are eventually turned into a list of holidays for
        // the production calendar
        private DateTime m_dtModifiedDay;
        private bool m_bWorkingDay;

        #endregion
    }
}
