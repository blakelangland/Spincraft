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
    public class HSReceipt
    {
        #region Constructors

        public HSReceipt(DataRow oDataRow)
        {
            if ((oDataRow["RcvDtl_Company"] != DBNull.Value) && (string.IsNullOrEmpty((string)oDataRow["RcvDtl_Company"]) == false))
            {
                m_sCompany = (string)oDataRow["RcvDtl_Company"];
            }
            if (oDataRow["RcvDtl_PONum"] != DBNull.Value)
            {
                m_iPONum = (int)oDataRow["RcvDtl_PONum"];
            }
            if (oDataRow["RcvDtl_POLine"] != DBNull.Value)
            {
                m_iPOLine = (int)oDataRow["RcvDtl_POLine"];
            }
            if (oDataRow["RcvDtl_PORelNum"] != DBNull.Value)
            {
                m_iPORelNum = (int)oDataRow["RcvDtl_PORelNum"];
            }
            if ((oDataRow["Calculated_POType"] != DBNull.Value) && (string.IsNullOrEmpty((string)oDataRow["Calculated_POType"]) == false))
            {
                m_sPOType = (string)oDataRow["Calculated_POType"];
            }
            if ((oDataRow["Calculated_RelType"] != DBNull.Value) && (string.IsNullOrEmpty((string)oDataRow["Calculated_RelType"]) == false))
            {
                m_sRelType = (string)oDataRow["Calculated_RelType"];
            }
            if ((oDataRow["RcvHead_EntryPerson"] != DBNull.Value) && (string.IsNullOrEmpty((string)oDataRow["RcvHead_EntryPerson"]) == false))
            {
                m_sEntryPerson = (string)oDataRow["RcvHead_EntryPerson"];
            }
            if (oDataRow["RcvHead_EntryDate"] != DBNull.Value)
            {
                m_dtEntryDate = (DateTime)oDataRow["RcvHead_EntryDate"];
            }
            if ((oDataRow["RcvDtl_PackSlip"] != DBNull.Value) && (string.IsNullOrEmpty((string)oDataRow["RcvDtl_PackSlip"]) == false))
            {
                m_sPackSlip = (string)oDataRow["RcvDtl_PackSlip"];
            }
            if (oDataRow["RcvDtl_PackLine"] != DBNull.Value)
            {
                m_iPackLine = (int)oDataRow["RcvDtl_PackLine"];
            }
            if ((oDataRow["Vendor_VendorID"] != DBNull.Value) && (string.IsNullOrEmpty((string)oDataRow["Vendor_VendorID"]) == false))
            {
                m_sVendorId = (string)oDataRow["Vendor_VendorID"];
            }
            if ((oDataRow["Vendor_Name"] != DBNull.Value) && (string.IsNullOrEmpty((string)oDataRow["Vendor_Name"]) == false))
            {
                m_sVendorName = (string)oDataRow["Vendor_Name"];
            }
            if ((oDataRow["RcvDtl_PartNum"] != DBNull.Value) && (string.IsNullOrEmpty((string)oDataRow["RcvDtl_PartNum"]) == false))
            {
                m_sPartNum = (string)oDataRow["RcvDtl_PartNum"];
            }
            if ((oDataRow["RcvDtl_PartDescription"] != DBNull.Value) && (string.IsNullOrEmpty((string)oDataRow["RcvDtl_PartDescription"]) == false))
            {
                m_sPartDescription = (string)oDataRow["RcvDtl_PartDescription"];
            }
            if (oDataRow["Calculated_PartOnFly"] != DBNull.Value)
            {
                m_bPartOnFly = (bool)oDataRow["Calculated_PartOnFly"];
            }
            if (oDataRow["RcvDtl_VendorQty"] != DBNull.Value)
            {
                m_dVendorQuantity = (decimal)oDataRow["RcvDtl_VendorQty"];
            }
            if ((oDataRow["RcvDtl_JobNum"] != DBNull.Value) && (string.IsNullOrEmpty((string)oDataRow["RcvDtl_JobNum"]) == false))
            {
                m_sJobNum = (string)oDataRow["RcvDtl_JobNum"];
            }
            if (oDataRow["RcvDtl_AssemblySeq"] != DBNull.Value)
            {
                m_iAssemblySeq = (int)oDataRow["RcvDtl_AssemblySeq"];
            }
            if ((oDataRow["RcvDtl_JobSeqType"] != DBNull.Value) && (string.IsNullOrEmpty((string)oDataRow["RcvDtl_JobSeqType"]) == false))
            {
                m_sJobSeqType = (string)oDataRow["RcvDtl_JobSeqType"];
            }
            if (oDataRow["RcvDtl_JobSeq"] != DBNull.Value)
            {
                m_iJobSeq = (int)oDataRow["RcvDtl_JobSeq"];
            }
            if (oDataRow["RcvDtl_InspectionPending"] != DBNull.Value)
            {
                m_bInspectionPending = (bool)oDataRow["RcvDtl_InspectionPending"];
            }
            if (oDataRow["RcvDtl_PassedQty"] != DBNull.Value)
            {
                m_dPassedQty = (decimal)oDataRow["RcvDtl_PassedQty"];
            }
            if (oDataRow["RcvDtl_FailedQty"] != DBNull.Value)
            {
                m_dFailedQty = (decimal)oDataRow["RcvDtl_FailedQty"];
            }
        }

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
        public int PORelNum
        {
            get { return m_iPORelNum; }
            set { m_iPORelNum = value; }
        }
        public string POType
        {
            get { return m_sPOType; }
            set { m_sPOType = value; }
        }
        public string RelType
        {
            get { return m_sRelType; }
            set { m_sRelType = value; }
        }
        public string PackSlip
        {
            get { return m_sPackSlip; }
            set { m_sPackSlip = value; }
        }
        public string EntryPerson
        {
            get { return m_sEntryPerson; }
            set { m_sEntryPerson = value; }
        }
        public DateTime EntryDate
        {
            get { return m_dtEntryDate; }
            set { m_dtEntryDate = value; }
        }
        public int PackLine
        {
            get { return m_iPackLine; }
            set { m_iPackLine = value; }
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
        public bool PartOnFly
        {
            get { return m_bPartOnFly; }
            set { m_bPartOnFly = value; }
        }
        public decimal VendorQuantity
        {
            get { return m_dVendorQuantity; }
            set { m_dVendorQuantity = value; }
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
        public bool InspectionPending
        {
            get { return m_bInspectionPending; }
            set { m_bInspectionPending = value; }
        }
        public decimal PassedQty
        {
            get { return m_dPassedQty; }
            set { m_dPassedQty = value; }
        }
        public decimal FailedQty
        {
            get { return m_dFailedQty; }
            set { m_dFailedQty = value; }
        }
        #endregion

        #region Data Members

        private string m_sCompany;
        private int m_iPONum;
        private int m_iPOLine;
        private int m_iPORelNum;
        private string m_sPOType;
        private string m_sRelType;
        private string m_sPackSlip;
        private string m_sEntryPerson;
        private DateTime m_dtEntryDate;
        private int m_iPackLine;
        private string m_sVendorId;
        private string m_sVendorName;
        private string m_sPartNum;
        private string m_sPartDescription;
        private bool m_bPartOnFly;
        private decimal m_dVendorQuantity;
        private string m_sJobNum;
        private int m_iAssemblySeq;
        private string m_sJobSeqType;
        private int m_iJobSeq;
        private bool m_bInspectionPending;
        private decimal m_dPassedQty;
        private decimal m_dFailedQty;
        #endregion

    }

    public class HSReceiptNoAttachment
    {
        #region Constructors

        public HSReceiptNoAttachment(DataRow oDataRow)
        {
            if (oDataRow["RcvHead_ReceiptDate"] != DBNull.Value)
            {
                m_dtReceiptDate = (DateTime)oDataRow["RcvHead_ReceiptDate"];
            }
            if ((oDataRow["RcvHead_PackSlip"] != DBNull.Value) && (string.IsNullOrEmpty((string)oDataRow["RcvHead_PackSlip"]) == false))
            {
                m_sPackSlip = (string)oDataRow["RcvHead_PackSlip"];
            }
            if ((oDataRow["Vendor_VendorID"] != DBNull.Value) && (string.IsNullOrEmpty((string)oDataRow["Vendor_VendorID"]) == false))
            {
                m_sVendorId = (string)oDataRow["Vendor_VendorID"];
            }
            if ((oDataRow["Vendor_Name"] != DBNull.Value) && (string.IsNullOrEmpty((string)oDataRow["Vendor_Name"]) == false))
            {
                m_sVendorName = (string)oDataRow["Vendor_Name"];
            }
            if (oDataRow["RcvHead_PONum"] != DBNull.Value)
            {
                m_iPONum = (int)oDataRow["RcvHead_PONum"];
            }
            if ((oDataRow["RcvHead_EntryPerson"] != DBNull.Value) && (string.IsNullOrEmpty((string)oDataRow["RcvHead_EntryPerson"]) == false))
            {
                m_sEntryPerson = (string)oDataRow["RcvHead_EntryPerson"];
            }
            if ((oDataRow["RcvHead_ReceivePerson"] != DBNull.Value) && (string.IsNullOrEmpty((string)oDataRow["RcvHead_ReceivePerson"]) == false))
            {
                m_sReceivePerson = (string)oDataRow["RcvHead_ReceivePerson"];
            }
        }

        #endregion

        #region Properties


        public DateTime ReceiptDate
        {
            get { return m_dtReceiptDate; }
        }
        public string PackSlip
        {
            get { return m_sPackSlip; }
        }
        public string VendorId
        {
            get { return m_sVendorId; }
        }
        public string VendorName
        {
            get { return m_sVendorName; }
        }
        public int PONum
        {
            get { return m_iPONum; }
        }
        public string EntryPerson
        {
            get { return m_sEntryPerson; }
        }
        public string ReceivePerson
        {
            get { return m_sReceivePerson; }
        }
        #endregion

        #region Data Members

        private DateTime m_dtReceiptDate;
        private string m_sPackSlip;
        private string m_sVendorId;
        private string m_sVendorName;
        private int m_iPONum;
        private string m_sEntryPerson;
        private string m_sReceivePerson;

        #endregion
    }

    public class HSPOReceipt
    {
        #region Constructors
        public HSPOReceipt(DataRow oDataRow)
        {
            if ((oDataRow["RcvDtl_PackSlip"] != DBNull.Value) && (string.IsNullOrEmpty((string)oDataRow["RcvDtl_PackSlip"]) == false))
            {
                m_sPackSlip = (string)oDataRow["RcvDtl_PackSlip"];
            }
            if (oDataRow["RcvDtl_PackLine"] != DBNull.Value)
            {
                m_iLine = (int)oDataRow["RcvDtl_PackLine"];
            }
            if (oDataRow["RcvDtl_ReceiptDate"] != DBNull.Value)
            {
                m_dtReceiptDate = (DateTime)oDataRow["RcvDtl_ReceiptDate"];
            }
            if ((oDataRow["RcvHead_ReceivePerson"] != DBNull.Value) && (string.IsNullOrEmpty((string)oDataRow["RcvHead_ReceivePerson"]) == false))
            {
                m_sReceivedBy = (string)oDataRow["RcvHead_ReceivePerson"];
            }
            if ((oDataRow["RcvDtl_PartNum"] != DBNull.Value) && (string.IsNullOrEmpty((string)oDataRow["RcvDtl_PartNum"]) == false))
            {
                m_sPartNum = (string)oDataRow["RcvDtl_PartNum"];
            }
            if ((oDataRow["RcvDtl_PartDescription"] != DBNull.Value) && (string.IsNullOrEmpty((string)oDataRow["RcvDtl_PartDescription"]) == false))
            {
                m_sPartDescription = (string)oDataRow["RcvDtl_PartDescription"];
            }
            if (oDataRow["RcvDtl_OurQty"] != DBNull.Value)
            {
                m_dQuantity = (decimal)oDataRow["RcvDtl_OurQty"];
            }
            if ((oDataRow["RcvDtl_WareHouseCode"] != DBNull.Value) && (string.IsNullOrEmpty((string)oDataRow["RcvDtl_WareHouseCode"]) == false))
            {
                m_sWarehouseCode = (string)oDataRow["RcvDtl_WareHouseCode"];
            }
            if ((oDataRow["RcvDtl_BinNum"] != DBNull.Value) && (string.IsNullOrEmpty((string)oDataRow["RcvDtl_BinNum"]) == false))
            {
                m_sBinNum = (string)oDataRow["RcvDtl_BinNum"];
            }
            if ((oDataRow["Vendor_VendorID"] != DBNull.Value) && (string.IsNullOrEmpty((string)oDataRow["Vendor_VendorID"]) == false))
            {
                m_sVendorId = (string)oDataRow["Vendor_VendorID"];
            }
            if ((oDataRow["Vendor_Name"] != DBNull.Value) && (string.IsNullOrEmpty((string)oDataRow["Vendor_Name"]) == false))
            {
                m_sVendorName = (string)oDataRow["Vendor_Name"];
            }
            if (oDataRow["RcvDtl_PONum"] != DBNull.Value)
            {
                m_iPONum = (int)oDataRow["RcvDtl_PONum"];
            }
            if (oDataRow["RcvDtl_POLine"] != DBNull.Value)
            {
                m_iPOLine = (int)oDataRow["RcvDtl_POLine"];
            }
            if (oDataRow["RcvDtl_PORelNum"] != DBNull.Value)
            {
                m_iPORelease = (int)oDataRow["RcvDtl_PORelNum"];
            }
            if (oDataRow["RcvDtl_InspectionReq"] != DBNull.Value)
            {
                m_bInspectionRequired = (bool)oDataRow["RcvDtl_InspectionReq"];
            }
            if (oDataRow["RcvDtl_PassedQty"] != DBNull.Value)
            {
                m_dPassedQuantity = (decimal)oDataRow["RcvDtl_PassedQty"];
            }
            if (oDataRow["RcvDtl_FailedQty"] != DBNull.Value)
            {
                m_dFailedQuantity = (decimal)oDataRow["RcvDtl_FailedQty"];
            }
            if (oDataRow["RcvDtl_OrderNum"] != DBNull.Value)
            {
                m_iOrderNum = (int)oDataRow["RcvDtl_OrderNum"];
            }
            if (oDataRow["RcvDtl_OrderLine"] != DBNull.Value)
            {
                m_iOrderLine = (int)oDataRow["RcvDtl_OrderLine"];
            }
            if (oDataRow["RcvDtl_OrderRelNum"] != DBNull.Value)
            {
                m_iOrderRelease = (int)oDataRow["RcvDtl_OrderRelNum"];
            }
            if (oDataRow["RcvDtl_JobNum"] != DBNull.Value)
            {
                m_sJobNum = (string)oDataRow["RcvDtl_JobNum"];
            }
            if (oDataRow["RcvDtl_AssemblySeq"] != DBNull.Value)
            {
                m_iAssemblySeq = (int)oDataRow["RcvDtl_AssemblySeq"];
            }
            if (oDataRow["RcvDtl_JobSeqType"] != DBNull.Value)
            {
                m_sJobSeqType = (string)oDataRow["RcvDtl_JobSeqType"];
            }
            if (oDataRow["RcvDtl_JobSeq"] != DBNull.Value)
            {
                m_iJobSeq = (int)oDataRow["RcvDtl_JobSeq"];
            }
        }

        #endregion

        #region Methods

        #endregion

        #region Properties

        public string PackSlip
        {
            get { return m_sPackSlip; }
            set { m_sPackSlip = value; }
        }
        public int Line
        {
            get { return m_iLine; }
            set { m_iLine = value; }
        }
        public DateTime ReceiptDate
        {
            get { return m_dtReceiptDate; }
            set { m_dtReceiptDate = value; }
        }
        public string ReceivedBy
        {
            get { return m_sReceivedBy; }
            set { m_sReceivedBy = value; }
        }
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
        public decimal Quantity
        {
            get { return m_dQuantity; }
            set { m_dQuantity = value; }
        }
        public string WarehouseCode
        {
            get { return m_sWarehouseCode; }
            set { m_sWarehouseCode = value; }
        }
        public string BinNum
        {
            get { return m_sBinNum; }
            set { m_sBinNum = value; }
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
        public int PORelease
        {
            get { return m_iPORelease; }
            set { m_iPORelease = value; }
        }
        public bool InspectionRequired
        {
            get { return m_bInspectionRequired; }
            set { m_bInspectionRequired = value; }
        }
        public decimal PassedQuantity
        {
            get { return m_dPassedQuantity; }
            set { m_dPassedQuantity = value; }
        }
        public decimal FailedQuantity
        {
            get { return m_dFailedQuantity; }
            set { m_dFailedQuantity = value; }
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
        #endregion

        #region Data Members

        private string m_sPackSlip;
        private int m_iLine;
        private DateTime m_dtReceiptDate;
        private string m_sReceivedBy;
        private string m_sPartNum;
        private string m_sPartDescription;
        private decimal m_dQuantity;
        private string m_sWarehouseCode;
        private string m_sBinNum;
        private string m_sVendorId;
        private string m_sVendorName;
        private int m_iPONum;
        private int m_iPOLine;
        private int m_iPORelease;
        private bool m_bInspectionRequired;
        private decimal m_dPassedQuantity;
        private decimal m_dFailedQuantity;
        private int m_iOrderNum;
        private int m_iOrderLine;
        private int m_iOrderRelease;
        private string m_sJobNum;
        private int m_iAssemblySeq;
        private string m_sJobSeqType;
        private int m_iJobSeq;
        #endregion
    }

    public class HSPOReceiptValidation
    {
        #region Constructors
        public HSPOReceiptValidation()
        {
        }
        #endregion

        #region Methods
        public bool Initialize(Session oSession)
        {
            bool bSuccess = true;

            //****Set a parameter Value***** 
            Ice.Proxy.BO.DynamicQueryImpl oDynamicQuery = WCFServiceSupport.CreateImpl<Ice.Proxy.BO.DynamicQueryImpl>(oSession, Ice.Proxy.BO.DynamicQueryImpl.UriPath);
            QueryExecutionDataSet oQueryExecutionDataSet = oDynamicQuery.GetQueryExecutionParametersByID(BAQConstants.QUERY_VALIDATE_PO_RECEIPT_FOR_INSPECTION);
            oQueryExecutionDataSet.ExecutionParameter.Clear();
            DataSet oDataSet = oDynamicQuery.ExecuteByID(BAQConstants.QUERY_VALIDATE_PO_RECEIPT_FOR_INSPECTION, oQueryExecutionDataSet);
            m_oPOReceipts.Clear();
            foreach (DataRow oDataRow in oDataSet.Tables[0].Rows)
            {
                m_oPOReceipts.Add(new HSPOReceipt(oDataRow));
            }

            oQueryExecutionDataSet = oDynamicQuery.GetQueryExecutionParametersByID(BAQConstants.QUERY_RECEIVE_DETAIL_OPEN);
            oQueryExecutionDataSet.ExecutionParameter.Clear();
            oDataSet = oDynamicQuery.ExecuteByID(BAQConstants.QUERY_RECEIVE_DETAIL_OPEN, oQueryExecutionDataSet);
            m_oReceipts.Clear();
            foreach (DataRow oDataRow in oDataSet.Tables[0].Rows)
            {
                m_oReceipts.Add(new HSReceipt(oDataRow));
            }

            //////////oQueryExecutionDataSet = oDynamicQuery.GetQueryExecutionParametersByID(BAQConstants.QUERY_RECEIVED_PACKSLIPS_NO_ATTACHMENTS);
            //////////oQueryExecutionDataSet.ExecutionParameter.Clear();
            //////////oDataSet = oDynamicQuery.ExecuteByID(BAQConstants.QUERY_RECEIVED_PACKSLIPS_NO_ATTACHMENTS, oQueryExecutionDataSet);
            //////////m_oReceiptsNoAttachments.Clear();
            //////////foreach (DataRow oDataRow in oDataSet.Tables[0].Rows)
            //////////{
            //////////    m_oReceiptsNoAttachments.Add(new HSReceiptNoAttachment(oDataRow));
            //////////}

            return bSuccess;
        }

        public void PerformValidation(string sCompany, string sTmpFileDirectory)
        {
            #region Setup

            // get the file name
            DateTime dtToday = DateTime.Now;
            string sDate = dtToday.Month.ToString() + "-" + dtToday.Day.ToString() + "-" + dtToday.Year.ToString();
            string sDestinationFileName = sTmpFileDirectory + "\\" + sCompany + "-PO Receipt Report-" + sDate + ".xlsx";
            int iNumberOfRows = 1;
            int iNumberOfColumns = 1;

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
            int iStandardColumnWidth = 30;

            bool bDataInReport = false;
            bool bFirstWorksheet = true;

            HSUser oServiceAccount = HSUser.GetUserById(HSUser.SPINCRAFT_SERVICE_ACCOUNT_ID);
            #endregion

            #region Production Issues
            //
            // Production Manager issues
            //
            SLDocument oSLProductionDocument = new SLDocument();
            SLStyle oNotInspectedOnTime = oSLProductionDocument.CreateStyle();
            oNotInspectedOnTime.ApplyNamedCellStyle(SLNamedCellStyleValues.Bad);

            List<string> oToAddresses = HSUser.GetEmailsForUsersInGroup(HSUser.REPORTS_PARTS_IN_INSPECTION);
            oToAddresses.Add(HSUser.SPINCRAFT_ROOT_USER);
            if (m_oPOReceipts.Count != 0)
            {
                if (bFirstWorksheet == true)
                {
                    oSLProductionDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "Inspection");
                    bFirstWorksheet = false;
                }
                else
                {
                    oSLProductionDocument.AddWorksheet("Inspection");
                }

                //set up column headers
                iNumberOfRows = 1;
                iNumberOfColumns = 1;

                oSLProductionDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "PO Num");
                oSLProductionDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLProductionDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "PO Line");
                oSLProductionDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLProductionDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "PO Rel");
                oSLProductionDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLProductionDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Part Num");
                oSLProductionDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLProductionDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Description");
                oSLProductionDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 50);
                oSLProductionDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Qty");
                oSLProductionDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLProductionDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Date Received");
                oSLProductionDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLProductionDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Warehouse");
                oSLProductionDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLProductionDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Supplier");
                oSLProductionDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 40);
            
                oSLProductionDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Parts Received That Are Still In Inspection");
                oSLProductionDocument.SetCellStyle(iNumberOfRows++, iNumberOfColumns, oHighlightHeaderStyle);
                oSLProductionDocument.SetColumnWidth(iNumberOfColumns++, 60);

                foreach (HSPOReceipt oPOReceipt in m_oPOReceipts)
                {
                    oSLProductionDocument.SetCellValue(iNumberOfRows, 1, oPOReceipt.PONum);
                    oSLProductionDocument.SetCellValue(iNumberOfRows, 2, oPOReceipt.POLine);
                    oSLProductionDocument.SetCellValue(iNumberOfRows, 3, oPOReceipt.PORelease);
                    oSLProductionDocument.SetCellValue(iNumberOfRows, 4, oPOReceipt.PartNum);
                    if (oPOReceipt.PartDescription.Length > 50)
                    {
                        oSLProductionDocument.SetCellValue(iNumberOfRows, 5, oPOReceipt.PartDescription.Substring(0, 50));
                    }
                    else
                    {
                        oSLProductionDocument.SetCellValue(iNumberOfRows, 5, oPOReceipt.PartDescription);
                    }
                    oSLProductionDocument.SetCellValue(iNumberOfRows, 6, oPOReceipt.Quantity.ToString());
                    oSLProductionDocument.SetCellValue(iNumberOfRows, 7, oPOReceipt.ReceiptDate.ToShortDateString());
                    if (oPOReceipt.ReceiptDate < DateTime.Now.AddDays(-3))
                    {
                        oSLProductionDocument.SetCellStyle(iNumberOfRows, 7, oNotInspectedOnTime);
                    }
                    oSLProductionDocument.SetCellValue(iNumberOfRows, 8, oPOReceipt.WarehouseCode);
                    oSLProductionDocument.SetCellValue(iNumberOfRows, 9, oPOReceipt.VendorName);

                    iNumberOfRows++;
                    bDataInReport = true;
                }
            }

            if (bDataInReport == true)
            {
                // we create a worksheet so we have some data in the spreadsheet so we email the spreadsheet
                oSLProductionDocument.SaveAs(sDestinationFileName);
                // Check to see if we created a file and if so email it
                if (File.Exists(sDestinationFileName) == true)
                {
                    List<string> oAttachments = new List<string>();
                    oAttachments.Add(sDestinationFileName);

                    HSEmailHelper.SendEmail(oToAddresses, sCompany + " Parts Needing Inspection", sCompany + " Parts Needing Inspection for " + sDate, oAttachments);
                }
            }


            #endregion

            #region Receiving Issues
            //
            // Receiving issues
            //
            bFirstWorksheet = true;
            bDataInReport = false;
            sDestinationFileName = sTmpFileDirectory + "Part Receipt Validation-" + dtToday.Month.ToString() + "-" + dtToday.Day.ToString() + "-" + dtToday.Year.ToString() + ".xlsx";
            SLDocument oSLReceiveDocument = new SLDocument();
            oToAddresses = HSUser.GetEmailsForUsersInGroup(HSUser.REPORTS_PO_RECEIVE_ISSUES);
            oToAddresses.Add(HSUser.SPINCRAFT_ROOT_USER);

            if (m_oReceipts.Count != 0)
            {
                if (bFirstWorksheet == true)
                {
                    oSLReceiveDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "Not Received");
                    bFirstWorksheet = false;
                }
                else
                {
                    oSLReceiveDocument.AddWorksheet("Not Received");
                }

                //set up column headers
                iNumberOfColumns = 1;
                iNumberOfRows = 1;

                oSLReceiveDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "PO Num");
                oSLReceiveDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLReceiveDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "PO Line");
                oSLReceiveDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLReceiveDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "PO Rel");
                oSLReceiveDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLReceiveDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Pack Slip");
                oSLReceiveDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLReceiveDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Vendor");
                oSLReceiveDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 40);
                oSLReceiveDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Part Num");
                oSLReceiveDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLReceiveDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Description");
                oSLReceiveDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 50);
                oSLReceiveDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Qty");
                oSLReceiveDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLReceiveDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Entry Date");
                oSLReceiveDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLReceiveDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Entry Person");
                oSLReceiveDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);

                oSLReceiveDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Receipt Process Started But Not Finished");
                oSLReceiveDocument.SetCellStyle(iNumberOfRows++, iNumberOfColumns, oHighlightHeaderStyle);
                oSLReceiveDocument.SetColumnWidth(iNumberOfColumns++, 60);

                foreach (HSReceipt oReceipt in m_oReceipts)
                {
                    oSLReceiveDocument.SetCellValue(iNumberOfRows, 1, oReceipt.PONum.ToString());
                    oSLReceiveDocument.SetCellValue(iNumberOfRows, 2, oReceipt.POLine.ToString());
                    oSLReceiveDocument.SetCellValue(iNumberOfRows, 3, oReceipt.PORelNum.ToString());
                    oSLReceiveDocument.SetCellValue(iNumberOfRows, 4, oReceipt.PackSlip);
                    oSLReceiveDocument.SetCellValue(iNumberOfRows, 5, oReceipt.VendorName);
                    oSLReceiveDocument.SetCellValue(iNumberOfRows, 6, oReceipt.PartNum);
                    if (oReceipt.PartDescription.Length > 50)
                    {
                        oSLReceiveDocument.SetCellValue(iNumberOfRows, 7, oReceipt.PartDescription.Substring(0, 50));
                    }
                    else
                    {
                        oSLReceiveDocument.SetCellValue(iNumberOfRows, 7, oReceipt.PartDescription);
                    }
                    oSLReceiveDocument.SetCellValue(iNumberOfRows, 8, oReceipt.VendorQuantity.ToString());
                    oSLReceiveDocument.SetCellValue(iNumberOfRows, 9, oReceipt.EntryDate.ToShortDateString());
                    oSLReceiveDocument.SetCellValue(iNumberOfRows, 10, oReceipt.EntryPerson);

                    iNumberOfRows++;
                    bDataInReport = true;
                }
            }

            if (m_oReceiptsNoAttachments.Count != 0)
            {
                if (bFirstWorksheet == true)
                {
                    oSLReceiveDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "Missing Attachment");
                    bFirstWorksheet = false;
                }
                else
                {
                    oSLReceiveDocument.AddWorksheet("Missing Attachment");
                }

                //set up column headers
                iNumberOfRows = 1;
                iNumberOfColumns = 1;

                oSLReceiveDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "PO Num");
                oSLReceiveDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLReceiveDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Pack Slip");
                oSLReceiveDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLReceiveDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Vendor");
                oSLReceiveDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 40);
                oSLReceiveDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Entry Person");
                oSLReceiveDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);

                oSLReceiveDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Receipt Does Not Have Pack Slip Attachment");
                oSLReceiveDocument.SetCellStyle(iNumberOfRows++, iNumberOfColumns, oHighlightHeaderStyle);
                oSLReceiveDocument.SetColumnWidth(iNumberOfColumns++, 60);

                foreach (HSReceiptNoAttachment oReceipt in m_oReceiptsNoAttachments)
                {
                    oSLReceiveDocument.SetCellValue(iNumberOfRows, 1, oReceipt.PONum.ToString());
                    oSLReceiveDocument.SetCellValue(iNumberOfRows, 2, oReceipt.PackSlip);
                    oSLReceiveDocument.SetCellValue(iNumberOfRows, 3, oReceipt.VendorName);
                    oSLReceiveDocument.SetCellValue(iNumberOfRows, 4, oReceipt.EntryPerson);

                    iNumberOfRows++;
                    bDataInReport = true;
                }
            }
            if (bDataInReport == true)
            {
                // we create a worksheet so we have some data in the spreadsheet so we email the spreadsheet
                oSLReceiveDocument.SaveAs(sDestinationFileName);
                // Check to see if we created a file and if so email it
                if (File.Exists(sDestinationFileName) == true)
                {
                    List<string> oAttachments = new List<string>();
                    oAttachments.Add(sDestinationFileName);

                    HSEmailHelper.SendEmail(oToAddresses, sCompany + " PO Receipt Issues", sCompany + " PO Receipt Issues for " + sDate, oAttachments);
                }
            }
            #endregion
        }
        #endregion

        #region Properties
        public List<HSPOReceipt> POReceipts
        {
            get { return m_oPOReceipts; }
        }
        public List<HSReceipt> Receipts
        {
            get { return m_oReceipts; }
        }
        public List<HSReceiptNoAttachment> ReceiptsNoAttachments
        {
            get { return m_oReceiptsNoAttachments; }
        }
        #endregion

        #region Data Members
        private List<HSPOReceipt> m_oPOReceipts = new List<HSPOReceipt>();
        private List<HSReceipt> m_oReceipts = new List<HSReceipt>();
        private List<HSReceiptNoAttachment> m_oReceiptsNoAttachments = new List<HSReceiptNoAttachment>();

        #endregion
    }
}