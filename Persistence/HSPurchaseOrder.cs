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
using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Drawing.Spreadsheet;
using SpreadsheetLight;
using SpreadsheetLight.Charts;
using Ice.Lib.Searches;
using System.IO;


namespace HorizonScientific
{
    public class HSPurchaseOrder
    {
        #region Constructors
        public HSPurchaseOrder(DataRow oRow)
        {
            if (oRow["POHeader_Company"] != DBNull.Value)
            {
                m_sCompany = (string)oRow["POHeader_Company"];
            }
            if (oRow["POHeader_PONum"] != DBNull.Value)
            {
                m_iPONum = (int)oRow["POHeader_PONum"];
            }
            if (oRow["PODetail_POLine"] != DBNull.Value)
            {
                m_iPOLine = (int)oRow["PODetail_POLine"];
            }
            if (oRow["PORel_PORelNum"] != DBNull.Value)
            {
                m_iPORelNun = (int)oRow["PORel_PORelNum"];
            }
            if (oRow["POHeader_OrderDate"] != DBNull.Value)
            {
                m_dtOrderDate = (DateTime) oRow["POHeader_OrderDate"];
            }
            if (oRow["PORel_DueDate"] != DBNull.Value)
            {
                m_dtDueDate = (DateTime)oRow["PORel_DueDate"];
            }
            if (oRow["PORel_PromiseDt"] != DBNull.Value)
            {
                m_dtPromiseDate = (DateTime)oRow["PORel_PromiseDt"];
            }
            if (oRow["POHeader_EntryPerson"] != DBNull.Value)
            {
                m_sEntryPerson = (string)oRow["POHeader_EntryPerson"];
            }
            if (oRow["PODetail_PartNum"] != DBNull.Value)
            {
                m_sPartNum = (string)oRow["PODetail_PartNum"];
                // there can text entered for a part number that is problematic
                m_sPartNum = StringExt.CleanString(m_sPartNum);
            }
            if (oRow["PODetail_RevisionNum"] != DBNull.Value)
            {
                m_sPartRevNum = (string)oRow["PODetail_RevisionNum"];
                // there can text entered for a part rev that is problematic
                m_sPartRevNum = StringExt.CleanString(m_sPartRevNum);
            }
            if (oRow["PODetail_ChangeDate"] != DBNull.Value)
            {
                m_dtChangeDate = (DateTime)oRow["PODetail_ChangeDate"];
            }
            if (oRow["PODetail_ChangedBy"] != DBNull.Value)
            {
                m_sChangedBy = (string)oRow["PODetail_ChangedBy"];
            }
            if (oRow["PORel_OpenRelease"] != DBNull.Value)
            {
                m_bOpenRelease = (bool)oRow["PORel_OpenRelease"];
            }
            if (oRow["PORel_VoidRelease"] != DBNull.Value)
            {
                m_bVoidRelease = (bool)oRow["PORel_VoidRelease"];
            }
            if (oRow["PODetail_OpenLine"] != DBNull.Value)
            {
                m_bOpenLine = (bool)oRow["PODetail_OpenLine"];
            }
            if (oRow["PODetail_VoidLine"] != DBNull.Value)
            {
                m_bVoidLine = (bool)oRow["PODetail_VoidLine"];
            }
            if (oRow["PORel_ReceivedQty"] != DBNull.Value)
            {
                m_dReceivedQty = (decimal)oRow["PORel_ReceivedQty"];
            }
            if (oRow["PORel_XRelQty"] != DBNull.Value)
            {
                m_dOurQty = (decimal)oRow["PORel_XRelQty"];
            }
            if (oRow["PORel_PurchasingFactor"] != DBNull.Value)
            {
                m_dPurchasingFactor = (decimal)oRow["PORel_PurchasingFactor"];
            }
            if (oRow["PORel_PurchasingFactorDirection"] != DBNull.Value)
            {
                m_sPurchasingDirection = (string)oRow["PORel_PurchasingFactorDirection"];
            }
            if (oRow["POHeader_BuyerID"] != DBNull.Value)
            {
                m_sBuyer = (string)oRow["POHeader_BuyerID"];
            }
            if (oRow["POHeader_ApprovedDate"] != DBNull.Value)
            {
                m_dtApprovedDate = (DateTime)oRow["POHeader_ApprovedDate"];
            }
            if (oRow["POHeader_ApprovedBy"] != DBNull.Value)
            {
                m_sApprovedBy = (string)oRow["POHeader_ApprovedBy"];
            }
            if (oRow["POHeader_Approve"] != DBNull.Value)
            {
                m_bApproved = (bool)oRow["POHeader_Approve"];
            }
            if (oRow["POHeader_Confirmed"] != DBNull.Value)
            {
                m_bConfirmed = (bool)oRow["POHeader_Confirmed"];
            }
            if (oRow["PORel_BTOOrderNum"] != DBNull.Value)
            {
                m_iBTOrderNum = (int)oRow["PORel_BTOOrderNum"];
            }
            if (oRow["PORel_BTOOrderLine"] != DBNull.Value)
            {
                m_iBTOrderLine = (int)oRow["PORel_BTOOrderLine"];
            }
            if (oRow["PORel_BTOOrderRelNum"] != DBNull.Value)
            {
                m_iBTOrderRel = (int)oRow["PORel_BTOOrderRelNum"];
            }
            if (oRow["PORel_DropShip"] != DBNull.Value)
            {
                m_bDropShip = (bool)oRow["PORel_DropShip"];
            }
            if (oRow["PORel_ArrivedQty"] != DBNull.Value)
            {
                m_dArrivedQty = (decimal)oRow["PORel_ArrivedQty"];
            }
            if (oRow["PORel_InvoicedQty"] != DBNull.Value)
            {
                m_dInvoicedQty = (decimal)oRow["PORel_InvoicedQty"];
            }
            if (oRow["PORel_NeedByDate"] != DBNull.Value)
            {
                m_dtNeedByDate = (DateTime)oRow["PORel_NeedByDate"];
            }
            if (oRow["PORel_InspectionQty"] != DBNull.Value)
            {
                m_dInspectionQty = (decimal)oRow["PORel_InspectionQty"];
            }
            if (oRow["PORel_FailedQty"] != DBNull.Value)
            {
                m_dFailedQty = (decimal)oRow["PORel_FailedQty"];
            }
            if (oRow["PORel_PassedQty"] != DBNull.Value)
            {
                m_dPassedQty = (decimal)oRow["PORel_PassedQty"];
            }
            if (oRow["PORel_JobNum"] != DBNull.Value)
            {
                m_sJobNum = (string)oRow["PORel_JobNum"];
            }
            if (oRow["PORel_AssemblySeq"] != DBNull.Value)
            {
                m_iAssemblySeq = (int)oRow["PORel_AssemblySeq"];
            }
            if (oRow["PORel_JobSeqType"] != DBNull.Value)
            {
                m_sJobSeqType = (string)oRow["PORel_JobSeqType"];
            }
            if (oRow["PORel_JobSeq"] != DBNull.Value)
            {
                m_iJobSeq = (int)oRow["PORel_JobSeq"];
            }
            if (oRow["PORel_WarehouseCode"] != DBNull.Value)
            {
                m_sWarehouseCode = (string)oRow["PORel_WarehouseCode"];
            }
            if (oRow["PODetail_QtyOption"] != DBNull.Value)
            {
                m_sQtyOption = (string)oRow["PODetail_QtyOption"];
            }
            if (oRow["PODetail_DocExtCost"] != DBNull.Value)
            {
                m_dDocExtCost = (decimal)oRow["PODetail_DocExtCost"];
            }
            if (oRow["PODetail_ExtCost"] != DBNull.Value)
            {
                m_dExtCost = (decimal)oRow["PODetail_ExtCost"];
            }
            if (oRow["PODetail_CostPerCode"] != DBNull.Value)
            {
                m_sCostPerCode = (string)oRow["PODetail_CostPerCode"];
            }
            if (oRow["PODetail_VendorNum"] != DBNull.Value)
            {
                m_iVendorNum = (int)oRow["PODetail_VendorNum"];
            }
            if (oRow["PODetail_PUM"] != DBNull.Value)
            {
                m_sPUM = (string)oRow["PODetail_PUM"];
            }
            if (oRow["JobHead_JobClosed"] != DBNull.Value)
            {
                m_bJobClosed = (bool)oRow["JobHead_JobClosed"];
            }
            if (oRow["JobHead_JobComplete"] != DBNull.Value)
            {
                m_bJobComplete = (bool)oRow["JobHead_JobComplete"];
            }
            if (oRow["OrderRel_OpenRelease"] != DBNull.Value)
            {
                m_bOpenSalesOrderRelease = (bool)oRow["OrderRel_OpenRelease"];
            }
            if (oRow["OrderRel_VoidRelease"] != DBNull.Value)
            {
                m_bVoidSalesOrderRelease = (bool)oRow["OrderRel_VoidRelease"];
            }

            // CER UD filed only in WI
            if (string.Compare(m_sCompany, CompanyConfiguration.SPINCRAFT_WI_COMPANY_ID, true) == 0)
            {
                if (oRow["POHeader_CERNumber_c"] != DBNull.Value)
                {
                    m_sCERNumber = (string)oRow["POHeader_CERNumber_c"];
                    // there can text entered for a CER number that is problematic
                    m_sCERNumber = StringExt.CleanString(m_sCERNumber);
                }
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
        public int PORelNun
        {
            get { return m_iPORelNun; }
            set { m_iPORelNun = value; }
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
        public string EntryPerson
        {
            get { return m_sEntryPerson; }
            set { m_sEntryPerson = value; }
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
        public DateTime ChangeDate
        {
            get { return m_dtChangeDate; }
            set { m_dtChangeDate = value; }
        }
        public string ChangedBy
        {
            get { return m_sChangedBy; }
            set { m_sChangedBy = value; }
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
        public bool OpenLine
        {
            get { return m_bOpenLine; }
            set { m_bOpenLine = value; }
        }
        public bool VoidLine
        {
            get { return m_bVoidLine; }
            set { m_bVoidLine = value; }
        }
        public decimal ReceivedQty
        {
            get { return m_dReceivedQty; }
            set { m_dReceivedQty = value; }
        }
        public decimal OurQty
        {
            get { return m_dOurQty; }
            set { m_dOurQty = value; }
        }
        public decimal PurchasingFactor
        {
            get { return m_dPurchasingFactor; }
            set { m_dPurchasingFactor = value; }
        }
        public string PurchasingDirection
        {
            get { return m_sPurchasingDirection; }
            set { m_sPurchasingDirection = value; }
        }
        public string Buyer
        {
            get { return m_sBuyer; }
            set { m_sBuyer = value; }
        }
        public DateTime ApprovedDate
        {
            get { return m_dtApprovedDate; }
            set { m_dtApprovedDate = value; }
        }
        public string ApprovedBy
        {
            get { return m_sApprovedBy; }
            set { m_sApprovedBy = value; }
        }
        public bool Approved
        {
            get { return m_bApproved; }
            set { m_bApproved = value; }
        }
        public bool Confirmed
        {
            get { return m_bConfirmed; }
            set { m_bConfirmed = value; }
        }
        public int BTOrderNum
        {
            get { return m_iBTOrderNum; }
            set { m_iBTOrderNum = value; }
        }
        public int BTOrderLine
        {
            get { return m_iBTOrderLine; }
            set { m_iBTOrderLine = value; }
        }
        public int BTOrderRel
        {
            get { return m_iBTOrderRel; }
            set { m_iBTOrderRel = value; }
        }
        public bool DropShip
        {
            get { return m_bDropShip; }
            set { m_bDropShip = value; }
        }
        public decimal ArrivedQty
        {
            get { return m_dArrivedQty; }
            set { m_dArrivedQty = value; }
        }
        public decimal InvoicedQty
        {
            get { return m_dInvoicedQty; }
            set { m_dInvoicedQty = value; }
        }
        public DateTime NeedByDate
        {
            get { return m_dtNeedByDate; }
            set { m_dtNeedByDate = value; }
        }
        public decimal InspectionQty
        {
            get { return m_dInspectionQty; }
            set { m_dInspectionQty = value; }
        }
        public decimal FailedQty
        {
            get { return m_dFailedQty; }
            set { m_dFailedQty = value; }
        }
        public decimal PassedQty
        {
            get { return m_dPassedQty; }
            set { m_dPassedQty = value; }
        }
        public string JobNum
        {
            get { return m_sJobNum; }
            set { m_sJobNum = value; }
        }
        public int AssemblySeq
        {
            get { return m_iAssemblySeq; }
            set { m_iAssemblySeq = value; }
        }
        public string JobSeqType
        {
            get { return m_sJobSeqType; }
            set { m_sJobSeqType = value; }
        }
        public int JobSeq
        {
            get { return m_iJobSeq; }
            set { m_iJobSeq = value; }
        }
        public string WarehouseCode
        {
            get { return m_sWarehouseCode; }
            set { m_sWarehouseCode = value; }
        }
        public string QtyOption
        {
            get { return m_sQtyOption; }
            set { m_sQtyOption = value; }
        }
        public decimal DocExtCost
        {
            get { return m_dDocExtCost; }
            set { m_dDocExtCost = value; }
        }
        public decimal ExtCost
        {
            get { return m_dExtCost; }
            set { m_dExtCost = value; }
        }
        public string CostPerCode
        {
            get { return m_sCostPerCode; }
            set { m_sCostPerCode = value; }
        }
        public int VendorNum
        {
            get { return m_iVendorNum; }
            set { m_iVendorNum = value; }
        }
        public string PUM
        {
            get { return m_sPUM; }
            set { m_sPUM = value; }
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
        public bool OpenSalesOrderRelease
        {
            get { return m_bOpenSalesOrderRelease; }
            set { m_bOpenSalesOrderRelease = value; }
        }
        public bool VoidSalesOrderRelease
        {
            get { return m_bVoidSalesOrderRelease; }
            set { m_bVoidSalesOrderRelease = value; }
        }
        public string CERNumber
        {
            get { return m_sCERNumber; }
            set { m_sCERNumber = value; }
        }
        #endregion

        #region Data Members
        private string m_sCompany;
        private int m_iPONum;
        private int m_iPOLine;
        private int m_iPORelNun;
        private DateTime m_dtOrderDate;
        private DateTime m_dtDueDate;
        private DateTime m_dtPromiseDate;
        private string m_sEntryPerson;
        private string m_sPartNum;
        private string m_sPartRevNum;
        private DateTime m_dtChangeDate;
        private string m_sChangedBy;
        private bool m_bOpenRelease;
        private bool m_bVoidRelease;
        private bool m_bOpenLine;
        private bool m_bVoidLine;
        private decimal m_dReceivedQty;
        private decimal m_dOurQty;
        private decimal m_dPurchasingFactor;
        private string m_sPurchasingDirection;
        private string m_sBuyer;
        private DateTime m_dtApprovedDate;
        private string m_sApprovedBy;
        private bool m_bApproved;
        private bool m_bConfirmed;
        private int m_iBTOrderNum;
        private int m_iBTOrderLine;
        private int m_iBTOrderRel;
        private bool m_bDropShip;
        private decimal m_dArrivedQty;
        private decimal m_dInvoicedQty;
        private DateTime m_dtNeedByDate;
        private decimal m_dInspectionQty;
        private decimal m_dFailedQty;
        private decimal m_dPassedQty;
        private string m_sJobNum;
        private int m_iAssemblySeq;
        private string m_sJobSeqType;
        private int m_iJobSeq;
        private string m_sWarehouseCode;
        private string m_sQtyOption;
        private decimal m_dDocExtCost;
        private decimal m_dExtCost;
        private string m_sCostPerCode;
        private int m_iVendorNum;
        private string m_sPUM;
        private bool m_bJobClosed;
        private bool m_bJobComplete;
        private bool m_bOpenSalesOrderRelease;
        private bool m_bVoidSalesOrderRelease;
        private string m_sCERNumber;
        #endregion
    }

    public class HSPOValidation
    {
        #region constructors
        public HSPOValidation()
        {
        }
        #endregion

        #region Methods
        public bool Initialize(Session oSession, HSValidateParts oValidateParts)
        {
            bool bSuccess = true;

            // loading all parts from the part master
            if (oValidateParts == null)
            {
                if (m_oValidateParts.Initialize(oSession) == false)
                {
                    Console.WriteLine("Failed to load the validate parts!");
                }
                oValidateParts = m_oValidateParts;
            }
            else
            {
                m_oValidateParts = oValidateParts;
            }

            // get a list of all materials for open jobs
            m_oPOs.Clear();
            Ice.Proxy.BO.DynamicQueryImpl oDynamicQuery = WCFServiceSupport.CreateImpl<Ice.Proxy.BO.DynamicQueryImpl>(oSession, Ice.Proxy.BO.DynamicQueryImpl.UriPath);
            QueryExecutionDataSet oQueryExecutionDataSet = oDynamicQuery.GetQueryExecutionParametersByID(BAQConstants.QUERY_LIST_VALIDATE_PURCHASE_ORDERS);
            oQueryExecutionDataSet.Clear();
            oQueryExecutionDataSet.AcceptChanges();
            DataSet oDataSet = oDynamicQuery.ExecuteByID(BAQConstants.QUERY_LIST_VALIDATE_PURCHASE_ORDERS, oQueryExecutionDataSet);
            foreach (DataRow oRow in oDataSet.Tables[0].Rows)
            {
                HSPurchaseOrder oPurchaseOrder = new HSPurchaseOrder(oRow);
                m_oPOs.Add(oPurchaseOrder);
            }

            return bSuccess;
        }

        public void PerformValidation(string sCompany, string sTmpFileDirectory)
        {
            #region Setup

            // get the file name
            DateTime dtToday = DateTime.Now;
            string sDate = dtToday.Month.ToString() + "-" + dtToday.Day.ToString() + "-" + dtToday.Year.ToString();
            string sDestinationFileName = sTmpFileDirectory + "\\" + sCompany + "-PO Validation Report-" + sDate + ".xlsx";
            int iNumberOfRows = 1;
            int iNumberOfColumns = 1;

            HSUser oServiceAccount = HSUser.GetUserById(HSUser.SPINCRAFT_SERVICE_ACCOUNT_ID);
            List<string> oToAddresses = new List<string>();
            // get users in the engineering, production, and quoting groups
            HSUser.GetEmailsForUsersInGroup(oToAddresses, HSUser.REPORT_ON_PROCUREMENT_ISSUES);
            oToAddresses.Add(HSUser.SPINCRAFT_ROOT_USER);

            SLFill oSLFill = new SLFill();
            oSLFill.SetPatternBackgroundColor(SLThemeColorIndexValues.Accent1Color);
            oSLFill.SetPattern(PatternValues.Solid, SLThemeColorIndexValues.Light2Color, SLThemeColorIndexValues.Light2Color);

            SLAlignment oCenterAlignment = new SLAlignment();
            oCenterAlignment.Horizontal = HorizontalAlignmentValues.Center;

            SLAlignment oRightAlignment = new SLAlignment();
            oRightAlignment.Horizontal = HorizontalAlignmentValues.Right;

            SLAlignment oLeftAlignment = new SLAlignment();
            oLeftAlignment.Horizontal = HorizontalAlignmentValues.Left;

            SLStyle oBoldStyle = new SLStyle();
            oBoldStyle.SetFontBold(true);
            oBoldStyle.SetTopBorder(BorderStyleValues.Thin, SLThemeColorIndexValues.Dark2Color);
            oBoldStyle.SetBottomBorder(BorderStyleValues.Thin, SLThemeColorIndexValues.Dark2Color);
            oBoldStyle.SetLeftBorder(BorderStyleValues.Thin, SLThemeColorIndexValues.Dark2Color);
            oBoldStyle.SetRightBorder(BorderStyleValues.Thin, SLThemeColorIndexValues.Dark2Color);
            oBoldStyle.SetFontColor(SLThemeColorIndexValues.Dark2Color);

            SLStyle oBoldStyle2 = new SLStyle();
            oBoldStyle2.SetFontBold(true);
            oBoldStyle2.SetTopBorder(BorderStyleValues.Thick, SLThemeColorIndexValues.Dark2Color);
            oBoldStyle2.SetBottomBorder(BorderStyleValues.Thick, SLThemeColorIndexValues.Dark2Color);
            oBoldStyle2.SetLeftBorder(BorderStyleValues.Thick, SLThemeColorIndexValues.Dark2Color);
            oBoldStyle.SetRightBorder(BorderStyleValues.Thick, SLThemeColorIndexValues.Dark2Color);
            oBoldStyle2.SetFontColor(System.Drawing.Color.IndianRed);

            SLStyle oCurrencyStyle = new SLStyle();
            oCurrencyStyle.ApplyNamedCellStyle(SLNamedCellStyleValues.Currency);
            oCurrencyStyle.Alignment = oLeftAlignment;
            oCurrencyStyle.FormatCode = "$#,##0.00";

            SLStyle oDecimalStyle = new SLStyle();
            oDecimalStyle.Alignment = oLeftAlignment;
            oDecimalStyle.FormatCode = "###.00";

            System.Drawing.Color oHighlightYellow = System.Drawing.Color.FromArgb(255, 255, 0);

            SLStyle oHighlightHeaderStyle = new SLStyle();
            oHighlightHeaderStyle.SetFontBold(true);
            oHighlightHeaderStyle.SetFont(FontSchemeValues.Major, 12);
            oHighlightHeaderStyle.SetHorizontalAlignment(HorizontalAlignmentValues.Center);
            oHighlightHeaderStyle.SetPatternFill(PatternValues.Solid, oHighlightYellow, oHighlightYellow);
            oHighlightHeaderStyle.SetFontColor(SLThemeColorIndexValues.Dark2Color);

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
            SLDocument oSLDocument = new SLDocument();

            bool bDataInReport = false;
            bool bFirstWorksheet = true;

            int iStandardColumnWidth = 30;
            #endregion

            // we will use this to see if we need to display fields specific to WI
            bool bETGWisconsin = false;
            if (string.Compare(sCompany, CompanyConfiguration.SPINCRAFT_WI_COMPANY_ID, true) == 0)
            {
                bETGWisconsin = true;
            }

            List<HSPurchaseOrder> oOpenReleases = m_oPOs.Where(oItem => oItem.OpenRelease == true).ToList();

            // POs that were created over two years ago but are still open
            DateTime dtTwoYearsAgo = DateTime.Now.AddYears(-2);
            List<HSPurchaseOrder> oOldPOs = m_oPOs.Where(oItem => (oItem.OrderDate <= dtTwoYearsAgo)).ToList();
            if (oOldPOs.Count > 0)
            {
                // sort these by po number
                oOldPOs = oOldPOs.OrderBy(oItem => oItem.PONum).ToList();
                // we will just pull out the first PO as we do not need to repeat these PO numbers for each reelease in the PO
                List<int> oDistinctPONums = new List<int>();
                foreach (HSPurchaseOrder oTmpPO in oOldPOs)
                {
                    if (oDistinctPONums.Contains(oTmpPO.PONum) == false)
                    {
                        oDistinctPONums.Add(oTmpPO.PONum);
                    }
                }

                iNumberOfRows = 1;
                iNumberOfColumns = 1;
                if (bFirstWorksheet == true)
                {
                    oSLDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "Old POs");
                    bFirstWorksheet = false;
                }
                else
                {
                    oSLDocument.AddWorksheet("Old POs");
                }
                //set column header
                oSLDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "PO Num");
                oSLDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Order Date");
                oSLDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Entered By");
                oSLDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                if (bETGWisconsin == true)
                {
                    oSLDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "CER Number");
                    oSLDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                }
                oSLDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "PO Is Open But Over Two Years Old");
                oSLDocument.SetCellStyle(iNumberOfRows++, iNumberOfColumns, oHighlightHeaderStyle);
                oSLDocument.SetColumnWidth(iNumberOfColumns++, 60);

                foreach (int iPONum in oDistinctPONums)
                {
                    HSPurchaseOrder oPO = oOldPOs.FirstOrDefault(oItem => oItem.PONum == iPONum);
                    if (oPO != null)
                    {
                        oSLDocument.SetCellValue(iNumberOfRows, 1, oPO.PONum);
                        if (oPO.OrderDate != DateTime.MinValue)
                        {
                            oSLDocument.SetCellValue(iNumberOfRows, 2, oPO.OrderDate.ToShortDateString());
                        }
                        else
                        {
                            oSLDocument.SetCellValue(iNumberOfRows, 2, "");
                        }
                        oSLDocument.SetCellValue(iNumberOfRows, 3, oPO.EntryPerson);
                        if (bETGWisconsin == true)
                        {
                            oSLDocument.SetCellValue(iNumberOfRows, 4, oPO.CERNumber);
                        }
                        iNumberOfRows++;
                        bDataInReport = true;
                    }
                }
            }

            // no due date
            List<HSPurchaseOrder> oMissingDueDate = oOpenReleases.Where(oItem => oItem.DueDate == DateTime.MinValue).ToList();
            if (oMissingDueDate.Count > 0)
            {
                // sort these by po number
                oMissingDueDate = oMissingDueDate.OrderBy(oItem => oItem.PONum).ToList();
                iNumberOfRows = 1;
                iNumberOfColumns = 1;
                if (bFirstWorksheet == true)
                {
                    oSLDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "No Due Date");
                    bFirstWorksheet = false;
                }
                else
                {
                    oSLDocument.AddWorksheet("No Due Date");
                }
                //set column header
                oSLDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "PO Num");
                oSLDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "PO Line");
                oSLDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "PO Rel");
                oSLDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Entered By");
                oSLDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Part Num");
                oSLDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Part Rev Num");
                oSLDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                if (bETGWisconsin == true)
                {
                    oSLDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "CER Number");
                    oSLDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                }
                oSLDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "The PO Release Does Not Have The Due Date Set");
                oSLDocument.SetCellStyle(iNumberOfRows++, iNumberOfColumns, oHighlightHeaderStyle);
                oSLDocument.SetColumnWidth(iNumberOfColumns++, 60);

                foreach (HSPurchaseOrder oPO in oMissingDueDate)
                {
                    oSLDocument.SetCellValue(iNumberOfRows, 1, oPO.PONum);
                    oSLDocument.SetCellValue(iNumberOfRows, 2, oPO.POLine);
                    oSLDocument.SetCellValue(iNumberOfRows, 3, oPO.PORelNun);
                    oSLDocument.SetCellValue(iNumberOfRows, 4, oPO.EntryPerson);
                    oSLDocument.SetCellValue(iNumberOfRows, 5, oPO.PartNum);
                    oSLDocument.SetCellValue(iNumberOfRows, 6, oPO.PartRevNum);
                    if (bETGWisconsin == true)
                    {
                        oSLDocument.SetCellValue(iNumberOfRows, 7, oPO.CERNumber);
                    }

                    iNumberOfRows++;
                    bDataInReport = true;
                }
            }

            // no proimise date
            List<HSPurchaseOrder> oMissingPromiseDate = oOpenReleases.Where(oItem => oItem.PromiseDate == DateTime.MinValue).ToList();
            if (oMissingPromiseDate.Count > 0)
            {
                // sort these by po number
                oMissingPromiseDate = oMissingPromiseDate.OrderBy(oItem => oItem.PONum).ToList();
                iNumberOfRows = 1;
                iNumberOfColumns = 1;
                if (bFirstWorksheet == true)
                {
                    oSLDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "No Promise Date");
                    bFirstWorksheet = false;
                }
                else
                {
                    oSLDocument.AddWorksheet("No Promise Date");
                }
                //set column header
                oSLDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "PO Num");
                oSLDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "PO Line");
                oSLDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "PO Rel");
                oSLDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Entered By");
                oSLDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Part Num");
                oSLDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Part Rev Num");
                oSLDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                if (bETGWisconsin == true)
                {
                    oSLDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "CER Number");
                    oSLDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                }
                oSLDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "The PO Release Does Not Have The Promsie Date Set");
                oSLDocument.SetCellStyle(iNumberOfRows++, iNumberOfColumns, oHighlightHeaderStyle);
                oSLDocument.SetColumnWidth(iNumberOfColumns++, 60);

                foreach (HSPurchaseOrder oPO in oMissingPromiseDate)
                {
                    oSLDocument.SetCellValue(iNumberOfRows, 1, oPO.PONum);
                    oSLDocument.SetCellValue(iNumberOfRows, 2, oPO.POLine);
                    oSLDocument.SetCellValue(iNumberOfRows, 3, oPO.PORelNun);
                    oSLDocument.SetCellValue(iNumberOfRows, 4, oPO.EntryPerson);
                    oSLDocument.SetCellValue(iNumberOfRows, 5, oPO.PartNum);
                    oSLDocument.SetCellValue(iNumberOfRows, 6, oPO.PartRevNum);
                    if (bETGWisconsin == true)
                    {
                        if (string.IsNullOrEmpty(oPO.CERNumber) == false)
                        {
                            oSLDocument.SetCellValue(iNumberOfRows, 7, oPO.CERNumber);
                        }
                    }

                    iNumberOfRows++;
                    bDataInReport = true;
                }
            }

            // no part rev specified
            List<HSPurchaseOrder> oMissingPartRev = oOpenReleases.Where(oItem => (string.IsNullOrEmpty(oItem.PartRevNum) == true)).ToList();
            List<HSPurchaseOrder> oTmpPOs = new List<HSPurchaseOrder>();
            // need to see if these parts require part revs
            foreach (HSPurchaseOrder oTmpPO in oMissingPartRev)
            {
                HSPartData oTmpPart = m_oValidateParts.AllParts.FirstOrDefault(oItem => (string.Compare(oTmpPO.PartNum, oItem.PartNum, true) == 0));
                if (oTmpPart != null)
                {
                    if (oTmpPart.UsePartRev == true)
                    {
                        oTmpPOs.Add(oTmpPO);
                    }
                }
            }
            oMissingPartRev = oTmpPOs;
            if (oMissingPartRev.Count > 0)
            {
                // sort these by po number
                oMissingPartRev = oMissingPartRev.OrderBy(oItem => oItem.PONum).ToList();
                iNumberOfRows = 1;
                iNumberOfColumns = 1;
                if (bFirstWorksheet == true)
                {
                    oSLDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "No Part Rev");
                    bFirstWorksheet = false;
                }
                else
                {
                    oSLDocument.AddWorksheet("No Part Rev");
                }
                //set column header
                oSLDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "PO Num");
                oSLDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "PO Line");
                oSLDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "PO Rel");
                oSLDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Entered By");
                oSLDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Part Num");
                oSLDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                if (bETGWisconsin == true)
                {
                    oSLDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "CER Number");
                    oSLDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                }
                oSLDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "The Part Rev Is Not Set");
                oSLDocument.SetCellStyle(iNumberOfRows++, iNumberOfColumns, oHighlightHeaderStyle);
                oSLDocument.SetColumnWidth(iNumberOfColumns++, 60);

                foreach (HSPurchaseOrder oPO in oMissingPartRev)
                {
                    oSLDocument.SetCellValue(iNumberOfRows, 1, oPO.PONum);
                    oSLDocument.SetCellValue(iNumberOfRows, 2, oPO.POLine);
                    oSLDocument.SetCellValue(iNumberOfRows, 3, oPO.PORelNun);
                    oSLDocument.SetCellValue(iNumberOfRows, 4, oPO.EntryPerson);
                    oSLDocument.SetCellValue(iNumberOfRows, 5, oPO.PartNum);
                    if (bETGWisconsin == true)
                    {
                        oSLDocument.SetCellValue(iNumberOfRows, 6, oPO.CERNumber);
                    }

                    iNumberOfRows++;
                    bDataInReport = true;
                }
            }

            // not approved
            List<HSPurchaseOrder> oPONotApproved = oOpenReleases.Where(oItem => oItem.Approved == false).ToList();
            if (oPONotApproved.Count > 0)
            {
                // sort these by po number
                oPONotApproved = oPONotApproved.OrderBy(oItem => oItem.PONum).ToList();
                // we will just pull out the first PO as we do not need to repeat these PO numbers for each reelease in the PO
                List<int> oDistinctPONums = new List<int>();
                foreach (HSPurchaseOrder oTmpPO in oPONotApproved)
                {
                    if (oDistinctPONums.Contains(oTmpPO.PONum) == false)
                    {
                        oDistinctPONums.Add(oTmpPO.PONum);
                    }
                }

                iNumberOfRows = 1;
                iNumberOfColumns = 1;
                if (bFirstWorksheet == true)
                {
                    oSLDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "Not Approved");
                    bFirstWorksheet = false;
                }
                else
                {
                    oSLDocument.AddWorksheet("Not Approved");
                }
                //set column header
                oSLDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "PO Num");
                oSLDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Order Date");
                oSLDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Entered By");
                oSLDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                if (bETGWisconsin == true)
                {
                    oSLDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "CER Number");
                    oSLDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                }
                oSLDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "PO Is Not Approved");
                oSLDocument.SetCellStyle(iNumberOfRows++, iNumberOfColumns, oHighlightHeaderStyle);
                oSLDocument.SetColumnWidth(iNumberOfColumns++, 60);

                foreach (int iPONum in oDistinctPONums)
                {
                    HSPurchaseOrder oPO = oPONotApproved.FirstOrDefault(oItem => oItem.PONum == iPONum);
                    if (oPO != null)
                    {
                        oSLDocument.SetCellValue(iNumberOfRows, 1, oPO.PONum);
                        if (oPO.OrderDate != DateTime.MinValue)
                        {
                            oSLDocument.SetCellValue(iNumberOfRows, 2, oPO.OrderDate.ToShortDateString());
                        }
                        else
                        {
                            oSLDocument.SetCellValue(iNumberOfRows, 2, "");
                        }
                        oSLDocument.SetCellValue(iNumberOfRows, 3, oPO.EntryPerson);
                        if (bETGWisconsin == true)
                        {
                            oSLDocument.SetCellValue(iNumberOfRows, 4, oPO.CERNumber);
                        }

                        iNumberOfRows++;
                        bDataInReport = true;
                    }
                }
            }

            // not confirmed
            List<HSPurchaseOrder> oPONotConfirmed = oOpenReleases.Where(oItem => oItem.Confirmed == false).ToList();
            if (oPONotConfirmed.Count > 0)
            {
                // sort these by po number
                oPONotConfirmed = oPONotConfirmed.OrderBy(oItem => oItem.PONum).ToList();
                // we will just pull out the first PO as we do not need to repeat these PO numbers for each reelease in the PO
                List<int> oDistinctPONums = new List<int>();
                foreach (HSPurchaseOrder oTmpPO in oPONotConfirmed)
                {
                    if (oDistinctPONums.Contains(oTmpPO.PONum) == false)
                    {
                        oDistinctPONums.Add(oTmpPO.PONum);
                    }
                }

                iNumberOfRows = 1;
                iNumberOfColumns = 1;
                if (bFirstWorksheet == true)
                {
                    oSLDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "Not Confirmed");
                    bFirstWorksheet = false;
                }
                else
                {
                    oSLDocument.AddWorksheet("Not Confirmed");
                }
                //set column header
                oSLDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "PO Num");
                oSLDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Order Date");
                oSLDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Entered By");
                oSLDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                if (bETGWisconsin == true)
                {
                    oSLDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "CER Number");
                    oSLDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                }
                oSLDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "PO Is Not Confirmed");
                oSLDocument.SetCellStyle(iNumberOfRows++, iNumberOfColumns, oHighlightHeaderStyle);
                oSLDocument.SetColumnWidth(iNumberOfColumns++, 60);

                foreach (int iPONum in oDistinctPONums)
                {
                    HSPurchaseOrder oPO = oPONotConfirmed.FirstOrDefault(oItem => oItem.PONum == iPONum);
                    if (oPO != null)
                    {
                        oSLDocument.SetCellValue(iNumberOfRows, 1, oPO.PONum);
                        if (oPO.OrderDate != DateTime.MinValue)
                        {
                            oSLDocument.SetCellValue(iNumberOfRows, 2, oPO.OrderDate.ToShortDateString());
                        }
                        else
                        {
                            oSLDocument.SetCellValue(iNumberOfRows, 2, "");
                        }
                        oSLDocument.SetCellValue(iNumberOfRows, 3, oPO.EntryPerson);
                        if (bETGWisconsin == true)
                        {
                            oSLDocument.SetCellValue(iNumberOfRows, 4, oPO.CERNumber);
                        }

                        iNumberOfRows++;
                        bDataInReport = true;
                    }
                }
            }

            // header open but all lines closed
            List<HSPurchaseOrder> oPOsWithOpenLines = m_oPOs.Where(oItem => (oItem.OpenLine == true)).ToList();
            // see which POs do not have open lines
            List<int> oPONumsWithNoOpenLines = new List<int>();
            foreach (HSPurchaseOrder oTmpPO in m_oPOs)
            {
                HSPurchaseOrder oExists = oPOsWithOpenLines.FirstOrDefault(oItem => oItem.PONum == oTmpPO.PONum);
                if (oExists == null)
                {
                    // we could not find it in the list so this PO has no open lines but the header is open
                    if (oPONumsWithNoOpenLines.Contains(oTmpPO.PONum) == false)
                    {
                        // only add it once to the list
                        oPONumsWithNoOpenLines.Add(oTmpPO.PONum);
                    }
                }
            }
            if (oPONumsWithNoOpenLines.Count > 0)
            {
                iNumberOfRows = 1;
                iNumberOfColumns = 1;
                if (bFirstWorksheet == true)
                {
                    oSLDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "No Open Lines");
                    bFirstWorksheet = false;
                }
                else
                {
                    oSLDocument.AddWorksheet("No Open Lines");
                }
                //set column header
                oSLDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "PO Num");
                oSLDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Order Date");
                oSLDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Entered By");
                oSLDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                if (bETGWisconsin == true)
                {
                    oSLDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "CER Number");
                    oSLDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                }
                oSLDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "PO Is Open But All Lines Are Closed");
                oSLDocument.SetCellStyle(iNumberOfRows++, iNumberOfColumns, oHighlightHeaderStyle);
                oSLDocument.SetColumnWidth(iNumberOfColumns++, 60);

                foreach (int iPONum in oPONumsWithNoOpenLines)
                {
                    HSPurchaseOrder oPO = m_oPOs.FirstOrDefault(oItem => oItem.PONum == iPONum);
                    if (oPO != null)
                    {
                        oSLDocument.SetCellValue(iNumberOfRows, 1, oPO.PONum);
                        if (oPO.OrderDate != DateTime.MinValue)
                        {
                            oSLDocument.SetCellValue(iNumberOfRows, 2, oPO.OrderDate.ToShortDateString());
                        }
                        else
                        {
                            oSLDocument.SetCellValue(iNumberOfRows, 2, "");
                        }
                        oSLDocument.SetCellValue(iNumberOfRows, 3, oPO.EntryPerson);
                        if (bETGWisconsin == true)
                        {
                            oSLDocument.SetCellValue(iNumberOfRows, 4, oPO.CERNumber);
                        }

                        iNumberOfRows++;
                        bDataInReport = true;
                    }
                }
            }

            #region IGNORE FOR NOW
            // line open but all releases are closed
            //////////List<HSPurchaseOrder> oPOsWithOpenReleases = m_oPOs.Where(oItem => (oItem.OpenLine== false) && (oItem.OpenRelease == true)).ToList();
            //////////// see which PO Lines do not have open releaess
            //////////List<int> oPOLinesWithNoOpenLines = new List<int>();
            //////////foreach (HSPurchaseOrder oTmpPO in m_oPOs)
            //////////{
            //////////    HSPurchaseOrder oExists = oPOsWithOpenLines.FirstOrDefault(oItem => oItem.PONum == oTmpPO.PONum);
            //////////    if (oExists == null)
            //////////    {
            //////////        // we could not find it in the list so this PO has no open lines but the header is open
            //////////        oPONumsWithNoOpenLines.Add(oTmpPO.PONum);
            //////////    }
            //////////}
            //////////if (oPONumsWithNoOpenLines.Count > 0)
            //////////{
            //////////    iNumberOfRows = 1;
            //////////    iNumberOfColumns = 1;
            //////////    if (bFirstWorksheet == true)
            //////////    {
            //////////        oSLDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "No Open Lines");
            //////////        bFirstWorksheet = false;
            //////////    }
            //////////    else
            //////////    {
            //////////        oSLDocument.AddWorksheet("No Open Lines");
            //////////    }
            //////////    //set column header
            //////////    oSLDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "PO Num");
            //////////    oSLDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
            //////////    oSLDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Order Date");
            //////////    oSLDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
            //////////    oSLDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Entered By");
            //////////    oSLDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
            ///////////   if (bETGWisconsin == true)
            //////////{
            //////////    oSLDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "CER Number");
            //////////    oSLDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
            //////////}
            //////////    oSLDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "PO Is Open But All Lines Are Closed");
            //////////    oSLDocument.SetCellStyle(iNumberOfRows++, iNumberOfColumns, oHighlightHeaderStyle);
            //////////    oSLDocument.SetColumnWidth(iNumberOfColumns++, 60);

            //////////    foreach (int iPONum in oPONumsWithNoOpenLines)
            //////////    {
            //////////        HSPurchaseOrder oPO = m_oPOs.FirstOrDefault(oItem => oItem.PONum == iPONum);
            //////////        if (oPO != null)
            //////////        {
            //////////            oSLDocument.SetCellValue(iNumberOfRows, 1, oPO.PONum);
            //////////            if (oPO.OrderDate != DateTime.MinValue)
            //////////            {
            //////////                oSLDocument.SetCellValue(iNumberOfRows, 2, oPO.OrderDate.ToShortDateString());
            //////////            }
            //////////            else
            //////////            {
            //////////                oSLDocument.SetCellValue(iNumberOfRows, 2, "");
            //////////            }
            //////////            oSLDocument.SetCellValue(iNumberOfRows, 3, oPO.EntryPerson);
            //////////if (bETGWisconsin == true)
            //////////{
            //////////    oSLDocument.SetCellValue(iNumberOfRows, 4, oPO.CERNumber);
            //////////}
            //////////            iNumberOfRows++;
            //////////            bDataInReport = true;
            //////////        }
            //////////    }
            //////////}
            #endregion

            // recived in full but still open
            List<HSPurchaseOrder> oPOReleaseReceivedInFullButStilOpen = oOpenReleases.Where(oItem => (oItem.ReceivedQty >= oItem.OurQty)).ToList();
            if (oPOReleaseReceivedInFullButStilOpen.Count > 0)
            {
                // sort these by po number
                oPOReleaseReceivedInFullButStilOpen = oPOReleaseReceivedInFullButStilOpen.OrderBy(oItem => oItem.PONum).ToList();
                iNumberOfRows = 1;
                iNumberOfColumns = 1;
                if (bFirstWorksheet == true)
                {
                    oSLDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "Rcvd Still Open");
                    bFirstWorksheet = false;
                }
                else
                {
                    oSLDocument.AddWorksheet("Rcvd Still Open");
                }
                //set column header
                oSLDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "PO Num");
                oSLDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "PO Line");
                oSLDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "PO Rel");
                oSLDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Entered By");
                oSLDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Part Num");
                oSLDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Part Rev Num");
                oSLDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Our Qty");
                oSLDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Received Qty");
                oSLDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                if (bETGWisconsin == true)
                {
                    oSLDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "CER Number");
                    oSLDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                }
                oSLDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "The PO Release Is Open But Has Been Received In Full");
                oSLDocument.SetCellStyle(iNumberOfRows++, iNumberOfColumns, oHighlightHeaderStyle);
                oSLDocument.SetColumnWidth(iNumberOfColumns++, 60);

                foreach (HSPurchaseOrder oPO in oPOReleaseReceivedInFullButStilOpen)
                {
                    oSLDocument.SetCellValue(iNumberOfRows, 1, oPO.PONum);
                    oSLDocument.SetCellValue(iNumberOfRows, 2, oPO.POLine);
                    oSLDocument.SetCellValue(iNumberOfRows, 3, oPO.PORelNun);
                    oSLDocument.SetCellValue(iNumberOfRows, 4, oPO.EntryPerson);
                    oSLDocument.SetCellValue(iNumberOfRows, 5, oPO.PartNum);
                    oSLDocument.SetCellValue(iNumberOfRows, 6, oPO.PartRevNum);
                    oSLDocument.SetCellValue(iNumberOfRows, 7, oPO.OurQty);
                    oSLDocument.SetCellValue(iNumberOfRows, 8, oPO.ReceivedQty);
                    if (bETGWisconsin == true)
                    {
                        oSLDocument.SetCellValue(iNumberOfRows, 9, oPO.CERNumber);
                    }

                    iNumberOfRows++;
                    bDataInReport = true;
                }
            }

            // buy to order but sales order rel is closed
            List<HSPurchaseOrder> oBuyToOrderSalesRelIsClosed = oOpenReleases.Where(oItem => (oItem.BTOrderNum != 0) && ((oItem.OpenSalesOrderRelease == false) || (oItem.VoidSalesOrderRelease == true))).ToList();
            if (oBuyToOrderSalesRelIsClosed.Count > 0)
            {
                // sort these by po number
                oBuyToOrderSalesRelIsClosed = oBuyToOrderSalesRelIsClosed.OrderBy(oItem => oItem.PONum).ToList();
                iNumberOfRows = 1;
                iNumberOfColumns = 1;
                if (bFirstWorksheet == true)
                {
                    oSLDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "SO Rel Closed");
                    bFirstWorksheet = false;
                }
                else
                {
                    oSLDocument.AddWorksheet("SO Rel Closed");
                }
                //set column header
                oSLDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "PO Num");
                oSLDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "PO Line");
                oSLDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "PO Rel");
                oSLDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Entered By");
                oSLDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Part Num");
                oSLDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Part Rev Num");
                oSLDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Order Num");
                oSLDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Order Line");
                oSLDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Order Rel");
                oSLDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                if (bETGWisconsin == true)
                {
                    oSLDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "CER Number");
                    oSLDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                }

                oSLDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "The PO Release Is Buy To Order But The Sales Order Release Is Closed");
                oSLDocument.SetCellStyle(iNumberOfRows++, iNumberOfColumns, oHighlightHeaderStyle);
                oSLDocument.SetColumnWidth(iNumberOfColumns++, 90);

                foreach (HSPurchaseOrder oPO in oBuyToOrderSalesRelIsClosed)
                {
                    oSLDocument.SetCellValue(iNumberOfRows, 1, oPO.PONum);
                    oSLDocument.SetCellValue(iNumberOfRows, 2, oPO.POLine);
                    oSLDocument.SetCellValue(iNumberOfRows, 3, oPO.PORelNun);
                    oSLDocument.SetCellValue(iNumberOfRows, 4, oPO.EntryPerson);
                    oSLDocument.SetCellValue(iNumberOfRows, 5, oPO.PartNum);
                    oSLDocument.SetCellValue(iNumberOfRows, 6, oPO.PartRevNum);
                    oSLDocument.SetCellValue(iNumberOfRows, 7, oPO.BTOrderNum);
                    oSLDocument.SetCellValue(iNumberOfRows, 8, oPO.BTOrderLine);
                    oSLDocument.SetCellValue(iNumberOfRows, 9, oPO.BTOrderRel);
                    if (bETGWisconsin == true)
                    {
                        oSLDocument.SetCellValue(iNumberOfRows, 10, oPO.CERNumber);
                    }

                    iNumberOfRows++;
                    bDataInReport = true;
                }
            }

            // buy to job but job is closed
            List<HSPurchaseOrder> oBuyToJobButJobIsClosed = oOpenReleases.Where(oItem => (string.IsNullOrEmpty(oItem.JobNum) != true) && ((oItem.JobClosed == true) || (oItem.JobComplete == true))).ToList();
            if (oBuyToJobButJobIsClosed.Count > 0)
            {
                // sort these by po number
                oBuyToJobButJobIsClosed = oBuyToJobButJobIsClosed.OrderBy(oItem => oItem.PONum).ToList();
                iNumberOfRows = 1;
                iNumberOfColumns = 1;
                if (bFirstWorksheet == true)
                {
                    oSLDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "Job Closed");
                    bFirstWorksheet = false;
                }
                else
                {
                    oSLDocument.AddWorksheet("Job Closed");
                }
                //set column header
                oSLDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "PO Num");
                oSLDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "PO Line");
                oSLDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "PO Rel");
                oSLDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Entered By");
                oSLDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Part Num");
                oSLDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Part Rev Num");
                oSLDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Job Num");
                oSLDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                if (bETGWisconsin == true)
                {
                    oSLDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "CER Number");
                    oSLDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                }

                oSLDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "The PO Release Is Buy To Job But The Job Is Completed Or Closed");
                oSLDocument.SetCellStyle(iNumberOfRows++, iNumberOfColumns, oHighlightHeaderStyle);
                oSLDocument.SetColumnWidth(iNumberOfColumns++, 90);

                foreach (HSPurchaseOrder oPO in oBuyToJobButJobIsClosed)
                {
                    oSLDocument.SetCellValue(iNumberOfRows, 1, oPO.PONum);
                    oSLDocument.SetCellValue(iNumberOfRows, 2, oPO.POLine);
                    oSLDocument.SetCellValue(iNumberOfRows, 3, oPO.PORelNun);
                    oSLDocument.SetCellValue(iNumberOfRows, 4, oPO.EntryPerson);
                    oSLDocument.SetCellValue(iNumberOfRows, 5, oPO.PartNum);
                    oSLDocument.SetCellValue(iNumberOfRows, 6, oPO.PartRevNum);
                    oSLDocument.SetCellValue(iNumberOfRows, 7, oPO.JobNum);
                    if (bETGWisconsin == true)
                    {
                        oSLDocument.SetCellValue(iNumberOfRows, 8, oPO.CERNumber);
                    }

                    iNumberOfRows++;
                    bDataInReport = true;
                }
            }

            // po is pending waiting for approval

            // arrived but not recived
            List<HSPurchaseOrder> oPOArrivedButNotReceived = oOpenReleases.Where(oItem => (oItem.ArrivedQty > oItem.ReceivedQty)).ToList();
            if (oPOArrivedButNotReceived.Count > 0)
            {
                // sort these by po number
                oPOArrivedButNotReceived = oPOArrivedButNotReceived.OrderBy(oItem => oItem.PONum).ToList();
                iNumberOfRows = 1;
                iNumberOfColumns = 1;
                if (bFirstWorksheet == true)
                {
                    oSLDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "Arrived Not Rcvd");
                    bFirstWorksheet = false;
                }
                else
                {
                    oSLDocument.AddWorksheet("Arrived Not Rcvd");
                }
                //set column header
                oSLDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "PO Num");
                oSLDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "PO Line");
                oSLDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "PO Rel");
                oSLDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Entered By");
                oSLDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Part Num");
                oSLDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Part Rev Num");
                oSLDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Arrived Qty");
                oSLDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Received Qty");
                oSLDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                if (bETGWisconsin == true)
                {
                    oSLDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "CER Number");
                    oSLDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                }

                oSLDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "The Receipt For The PO Is In The Arrived State But Has Not Been Received");
                oSLDocument.SetCellStyle(iNumberOfRows++, iNumberOfColumns, oHighlightHeaderStyle);
                oSLDocument.SetColumnWidth(iNumberOfColumns++, 90);

                foreach (HSPurchaseOrder oPO in oPOArrivedButNotReceived)
                {
                    oSLDocument.SetCellValue(iNumberOfRows, 1, oPO.PONum);
                    oSLDocument.SetCellValue(iNumberOfRows, 2, oPO.POLine);
                    oSLDocument.SetCellValue(iNumberOfRows, 3, oPO.PORelNun);
                    oSLDocument.SetCellValue(iNumberOfRows, 4, oPO.EntryPerson);
                    oSLDocument.SetCellValue(iNumberOfRows, 5, oPO.PartNum);
                    oSLDocument.SetCellValue(iNumberOfRows, 6, oPO.PartRevNum);
                    oSLDocument.SetCellValue(iNumberOfRows, 7, oPO.ArrivedQty);
                    oSLDocument.SetCellValue(iNumberOfRows, 8, oPO.ReceivedQty);
                    if (bETGWisconsin == true)
                    {
                        oSLDocument.SetCellValue(iNumberOfRows, 9, oPO.CERNumber);
                    }

                    iNumberOfRows++;
                    bDataInReport = true;
                }
            }

            // cost not set
            List<HSPurchaseOrder> oCostNotSetOnPO = oOpenReleases.Where(oItem => (oItem.ExtCost == 0)).ToList();
            if (oCostNotSetOnPO.Count > 0)
            {
                // sort these by po number
                oCostNotSetOnPO = oCostNotSetOnPO.OrderBy(oItem => oItem.PONum).ToList();
                iNumberOfRows = 1;
                iNumberOfColumns = 1;
                if (bFirstWorksheet == true)
                {
                    oSLDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "No Cost");
                    bFirstWorksheet = false;
                }
                else
                {
                    oSLDocument.AddWorksheet("No Cost");
                }
                //set column header
                oSLDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "PO Num");
                oSLDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "PO Line");
                oSLDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "PO Rel");
                oSLDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Entered By");
                oSLDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Part Num");
                oSLDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Part Rev Num");
                oSLDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                if (bETGWisconsin == true)
                {
                    oSLDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "CER Number");
                    oSLDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                }

                oSLDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "The Ext Cost For This Part Has Not Been Set");
                oSLDocument.SetCellStyle(iNumberOfRows++, iNumberOfColumns, oHighlightHeaderStyle);
                oSLDocument.SetColumnWidth(iNumberOfColumns++, 60);

                foreach (HSPurchaseOrder oPO in oCostNotSetOnPO)
                {
                    oSLDocument.SetCellValue(iNumberOfRows, 1, oPO.PONum);
                    oSLDocument.SetCellValue(iNumberOfRows, 2, oPO.POLine);
                    oSLDocument.SetCellValue(iNumberOfRows, 3, oPO.PORelNun);
                    oSLDocument.SetCellValue(iNumberOfRows, 4, oPO.EntryPerson);
                    oSLDocument.SetCellValue(iNumberOfRows, 5, oPO.PartNum);
                    oSLDocument.SetCellValue(iNumberOfRows, 6, oPO.PartRevNum);
                    if (bETGWisconsin == true)
                    {
                        oSLDocument.SetCellValue(iNumberOfRows, 7, oPO.CERNumber);
                    }

                    iNumberOfRows++;
                    bDataInReport = true;
                }
            }

            // check po rel status

//////////            Open(O), 
//////////Arrived(A), 
//////////Inspection(I), 
//////////Received(R), 
//////////Consumed(U), 
//////////Drop Shipped(D), 
//////////Closed(C), 
//////////Voided(V).





            if (bDataInReport == true)
            {
                oSLDocument.SaveAs(sDestinationFileName);
                // Check to see if we created a file and if so email it
                if (File.Exists(sDestinationFileName) == true)
                {
                    List<string> oAttachments = new List<string>();
                    oAttachments.Add(sDestinationFileName);

                    HSEmailHelper.SendEmail(oToAddresses, sCompany + " PO Validation Report", sCompany + " PO Validation Report for " + sDate, oAttachments);
                }
            }
        }
        #endregion

        #region Properties
        #endregion

        #region Data Members
        private HSValidateParts m_oValidateParts = new HSValidateParts();
        private List<HSPurchaseOrder> m_oPOs = new List<HSPurchaseOrder>();
        #endregion

    }
}
