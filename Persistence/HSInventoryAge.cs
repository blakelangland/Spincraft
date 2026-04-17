using DocumentFormat.OpenXml.Bibliography;
using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Spreadsheet;
using Erp.BO;
using Erp.Common.ContractInterfaces;
using Erp.Proxy.BO;
using Erp.Tablesets;
using HorizonScientific;
using Ice.Adapters;
using Ice.BO;
using Ice.Core;
using Ice.Lib.Customization.Dialogs;
using Ice.Lib.Framework;
using SpreadsheetLight;
using SpreadsheetLight.Charts;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;


namespace HorizonScientific
{
    public class HSCalculateInventoryAge
    {
        #region Constructors
        public HSCalculateInventoryAge()
        {
            // we will start looking at inventory age from a few years back
            // if this is not a big enough window then we will keep walking this back
            // until we have all the data we need to determine the age of inventory
            m_dtStartDate = DateTime.Now.AddYears(-LOOK_BACK_YEARS);
        }
        #endregion

        #region Methods
        public bool Initialize(Session oSession)
        {
            bool bSuccess = true;

            //****Set a parameter Value***** 
            Ice.Proxy.BO.DynamicQueryImpl oDynamicQuery = WCFServiceSupport.CreateImpl<Ice.Proxy.BO.DynamicQueryImpl>(oSession, Ice.Proxy.BO.DynamicQueryImpl.UriPath);

            // first we determine how much inventory we have on hand
            QueryExecutionDataSet oQueryExecutionDataSet = oDynamicQuery.GetQueryExecutionParametersByID(BAQConstants.QUERY_PART_INVENTORY_AGING);
            oQueryExecutionDataSet.ExecutionParameter.Clear();
            DataSet oDataSet = oDynamicQuery.ExecuteByID(BAQConstants.QUERY_PART_INVENTORY_AGING, oQueryExecutionDataSet);
            m_oAgeOfAllInventory.Clear();
            foreach (DataRow oRow in oDataSet.Tables[0].Rows)
            {
                HSInventoryAge oTmpInventoryAge = new HSInventoryAge(oRow);
                // we will filter out any zero or negative inventory
                if (oTmpInventoryAge.TotalOnHand > 0)
                {
                    m_oAgeOfAllInventory.Add(oTmpInventoryAge);
                }
            }

            oQueryExecutionDataSet = oDynamicQuery.GetQueryExecutionParametersByID(BAQConstants.QUERY_PART_RECEIPTS_IN_DATE_RANGE);
            oQueryExecutionDataSet.Clear();
            oQueryExecutionDataSet.ExecutionParameter.AddExecutionParameterRow("StartDate", m_dtStartDate.ToShortDateString(), "date", false, Guid.NewGuid(), "A");
            oQueryExecutionDataSet.AcceptChanges();
            oDataSet = oDynamicQuery.ExecuteByID(BAQConstants.QUERY_PART_RECEIPTS_IN_DATE_RANGE, oQueryExecutionDataSet);
            m_oPartReceiptsAndAdjustments.Clear();
            foreach (DataRow oRow in oDataSet.Tables[0].Rows)
            {
                HSPartReceiptsOrAdjustment oTmpPartReceiptOrAdjustment = new HSPartReceiptsOrAdjustment(oRow);
                m_oPartReceiptsAndAdjustments.Add(oTmpPartReceiptOrAdjustment);
            }

            // find all parts for finished goods
            oQueryExecutionDataSet = oDynamicQuery.GetQueryExecutionParametersByID(BAQConstants.QUERY_SIMPLE_WHERE_USED);
            oQueryExecutionDataSet.ExecutionParameter.Clear();
            oDataSet = oDynamicQuery.ExecuteByID(BAQConstants.QUERY_SIMPLE_WHERE_USED, oQueryExecutionDataSet);
            m_oSimpleWhereUsed.Clear();
            foreach (DataRow oRow in oDataSet.Tables[0].Rows)
            {
                SimpleWhereUsed oTmpSimpleWhereUsed = new SimpleWhereUsed(oRow);
                m_oSimpleWhereUsed.Add(oTmpSimpleWhereUsed);
            }

            CalculateBatchAverageAge();

            return bSuccess;
        }

        public void CreateReport(string sArchiveFileDirectory, string sTmpFileDirectory, string sCompany, HSUser oRequestingUser)
        {
            #region Setup
            DateTime dtToday = DateTime.Now;
            string sDate = dtToday.Month.ToString() + "-" + dtToday.Day.ToString() + "-" + dtToday.Year.ToString();
            string sDestinationFileName = sTmpFileDirectory + "\\" + sCompany + "-InventoryAging-" + sDate + ".xlsx";

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
            if (oRequestingUser != null)
            {
                // this report was requested by someone directly from Epicor
                oToAddresses.Add(oRequestingUser.Email);
            }
            oToAddresses.Add(HSUser.SPINCRAFT_ROOT_USER);

            SLDocument oSLDocument = new SLDocument();

            // create the styles we will use in this document
            SLStyle oMoney = oSLDocument.CreateStyle();
            oMoney.FormatCode = "$#,##0.00";
            SLStyle oPercentage = oSLDocument.CreateStyle();
            oPercentage.FormatCode = "0.00%";
            SLStyle oGood = oSLDocument.CreateStyle();
            oGood.ApplyNamedCellStyle(SLNamedCellStyleValues.Good);
            SLStyle oNeutral = oSLDocument.CreateStyle();
            oNeutral.ApplyNamedCellStyle(SLNamedCellStyleValues.Neutral);
            SLStyle oBad = oSLDocument.CreateStyle();
            oBad.ApplyNamedCellStyle(SLNamedCellStyleValues.Bad);

            System.Drawing.Color oRed = System.Drawing.Color.FromArgb(255, 32, 32);
            System.Drawing.Color oGreen = System.Drawing.Color.FromArgb(0, 100, 5);
            System.Drawing.Color oDRed = System.Drawing.Color.FromArgb(255, 100, 100);
            System.Drawing.Color oDGreen = System.Drawing.Color.FromArgb(50, 100, 32);
            System.Drawing.Color oDYellow = System.Drawing.Color.FromArgb(200, 200, 50);
            System.Drawing.Color oDBlue = System.Drawing.Color.FromArgb(0, 100, 200);
            System.Drawing.Color oDOrange = System.Drawing.Color.FromArgb(255, 128, 2);

            SLStyle oBelowSafetyStyle = new SLStyle();
            oBelowSafetyStyle.SetFontBold(true);
            oBelowSafetyStyle.SetFontColor(oDOrange);

            SLStyle oBelowMinimumStyle = new SLStyle();
            oBelowMinimumStyle.SetFontBold(true);
            oBelowMinimumStyle.SetFontColor(oDYellow);

            SLStyle oBelowZeroStyle = new SLStyle();
            oBelowZeroStyle.SetFontBold(true);
            oBelowZeroStyle.SetFontColor(oRed);

            SLStyle oLatePOStyle = new SLStyle();
            oLatePOStyle.SetFontBold(true);
            oLatePOStyle.SetFontColor(oRed);

            SLFill oSLFill = new SLFill();
            oSLFill.SetPatternBackgroundColor(SLThemeColorIndexValues.Dark2Color);

            SLAlignment oLeftAlignment = new SLAlignment();
            oLeftAlignment.Horizontal = DocumentFormat.OpenXml.Spreadsheet.HorizontalAlignmentValues.Left;

            SLStyle oLeftAlignmentStyle = new SLStyle();
            oLeftAlignmentStyle.Alignment = oLeftAlignment;

            SLAlignment oRightAlignment = new SLAlignment();
            oRightAlignment.Horizontal = DocumentFormat.OpenXml.Spreadsheet.HorizontalAlignmentValues.Right;

            SLStyle oRightAlignmentStyle = new SLStyle();
            oRightAlignmentStyle.Alignment = oRightAlignment;

            SLStyle oBoldStyle = new SLStyle();
            oBoldStyle.SetFontBold(true);
            oBoldStyle.SetTopBorder(BorderStyleValues.Thin, SLThemeColorIndexValues.Dark2Color);
            oBoldStyle.SetBottomBorder(BorderStyleValues.Thin, SLThemeColorIndexValues.Dark2Color);
            oBoldStyle.SetLeftBorder(BorderStyleValues.Thin, SLThemeColorIndexValues.Dark2Color);
            oBoldStyle.SetRightBorder(BorderStyleValues.Thin, SLThemeColorIndexValues.Dark2Color);
            oBoldStyle.SetFontColor(SLThemeColorIndexValues.Dark2Color);

            SLStyle oInventoryQuantityStyle = new SLStyle();
            oInventoryQuantityStyle.FormatCode = "######";
            #endregion

            #region Purchased Part Inventory Aging

            bool bDataInReport = false;
            bool bFirstWorksheet = true;
            int iNumberOfRows = 1;
            int iNumberOfColumns = 1;

            // order the inventory by part number, then lot number, and then serial number so that it is easier to read in the report
            m_oAgeOfAllInventory = m_oAgeOfAllInventory.OrderBy(x => x.PartNum).ThenBy(x => x.LotNum).ThenBy(x => x.SerialNumber).ToList();

            if (m_oAgeOfAllInventory.Count != 0)
            {
                if (bFirstWorksheet == true)
                {
                    oSLDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "Aged Inventory");
                    bFirstWorksheet = false;
                }
                else
                {
                    oSLDocument.AddWorksheet("Aged Inventory");
                }
                //set column header
                oSLDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "PartNum");
                oSLDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 50);
                oSLDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Description");
                oSLDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 50);
                oSLDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Type");
                oSLDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 10);
                oSLDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Class ID");
                oSLDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 10);
                oSLDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Class Description");
                oSLDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 20);
                oSLDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "IUM");
                oSLDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 10);
                oSLDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Cost Method");
                oSLDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 15);
                oSLDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Cost Per");
                oSLDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 10);
                oSLDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Qty");
                oSLDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 10);
                oSLDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Ext Cost");
                oSLDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 10);
                oSLDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Lot Tracked");
                oSLDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 20);
                oSLDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Lot Num");
                oSLDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 20);
                oSLDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Serial Tracked");
                oSLDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 20);
                oSLDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Serial Num");
                oSLDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 20);
                oSLDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Missing Receipts");
                oSLDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 20);
                oSLDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Average Age");
                oSLDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 20);
                oSLDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Where Used");
                oSLDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 20);
                iNumberOfRows++;

                foreach (HSInventoryAge oPartAge in m_oAgeOfAllInventory)
                {
                    oSLDocument.SetCellValue(iNumberOfRows, 1, oPartAge.PartNum);
                    oSLDocument.SetCellValue(iNumberOfRows, 2, oPartAge.PartDescription);
                    oSLDocument.SetCellValue(iNumberOfRows, 3, oPartAge.TypeCode);
                    oSLDocument.SetCellValue(iNumberOfRows, 4, oPartAge.ClassId);
                    oSLDocument.SetCellValue(iNumberOfRows, 5, oPartAge.ClassDescription);
                    oSLDocument.SetCellValue(iNumberOfRows, 6, oPartAge.IUM);
                    oSLDocument.SetCellValue(iNumberOfRows, 7, oPartAge.CostMethod);
                    oSLDocument.SetCellValue(iNumberOfRows, 8, oPartAge.CostPer);
                    oSLDocument.SetCellStyle(iNumberOfRows, 8, oMoney);
                    oSLDocument.SetCellValue(iNumberOfRows, 9, oPartAge.TotalOnHand);
                    oSLDocument.SetCellValue(iNumberOfRows, 10, oPartAge.ExtendedCost);
                    oSLDocument.SetCellStyle(iNumberOfRows, 10, oMoney);
                    oSLDocument.SetCellValue(iNumberOfRows, 11, oPartAge.TrackLots);
                    oSLDocument.SetCellValue(iNumberOfRows, 12, oPartAge.LotNum);
                    oSLDocument.SetCellValue(iNumberOfRows, 13, oPartAge.TrackSerialNum);
                    oSLDocument.SetCellValue(iNumberOfRows, 14, oPartAge.SerialNumber);
                    oSLDocument.SetCellValue(iNumberOfRows, 15, oPartAge.MissingReceipts);
                    if (oPartAge.MissingReceipts == true)
                    {
                        oSLDocument.SetCellStyle(iNumberOfRows, 15, oBad);
                    }
                    oSLDocument.SetCellValue(iNumberOfRows, 16, oPartAge.AverageAgeInDays);
                    // just use an int for days
                    oSLDocument.SetCellStyle(iNumberOfRows, 16, oInventoryQuantityStyle);
                    oSLDocument.SetCellValue(iNumberOfRows, 17, oPartAge.WhereUsed);

                    iNumberOfRows++;
                }

                bDataInReport = true;
            }

            if (bDataInReport == true)
            {
                oSLDocument.SaveAs(sDestinationFileName);
                // Check to see if we created a file and if so email it
                if (File.Exists(sDestinationFileName) == true)
                {
                    List<string> oAttachments = new List<string>();
                    oAttachments.Add(sDestinationFileName);

                    HSEmailHelper.SendEmail(oToAddresses, "Inventory Aging Report", "Inventory Aging Report for " + sDate, oAttachments);
                }
            }
            #endregion
        }

        public void CalculateBatchAverageAge()
        {
            // we need to group the part trans records by PartNum|Lot|Serial
            Dictionary<InventoryKey, List<HSPartReceiptsOrAdjustment>> oPartTransDictionary = new Dictionary<InventoryKey, List<HSPartReceiptsOrAdjustment>>();

            // create the keys for the records in the part tran list so we can extract them quickly for each inventory item
            foreach (HSPartReceiptsOrAdjustment oPartTran in m_oPartReceiptsAndAdjustments)
            {
                InventoryKey oKey = new InventoryKey(oPartTran.PartNum, oPartTran.LotNumber, oPartTran.SerialNumber);

                if (!oPartTransDictionary.ContainsKey(oKey))
                {
                    oPartTransDictionary[oKey] = new List<HSPartReceiptsOrAdjustment>();
                }
                oPartTransDictionary[oKey].Add(oPartTran);
            }

            // walk through each inventory item and compute its age
            foreach (HSInventoryAge oInventoryItem in m_oAgeOfAllInventory)
            {
                // get the key for this inventory item so we can lookup the transactions that match this item
                InventoryKey oKey = new InventoryKey(oInventoryItem.PartNum, oInventoryItem.LotNum, oInventoryItem.SerialNumber);

                List<HSPartReceiptsOrAdjustment> oItemTransactions = oPartTransDictionary.ContainsKey(oKey)
                    ? oPartTransDictionary[oKey]
                    : new List<HSPartReceiptsOrAdjustment>();

                // now we remove any offsetting transactions (oldest -> newest) to get a list of receipts that represent the layers that make up the current inventory balance
                // this helps to get better aging data because we can have adjustments that occur after the receipt of the inventory that can throw off the aging if not accounted for
                var oCleanedPartTrans = RemoveOffsettingTransactions(oItemTransactions);

                // Sort from newest to oldest for backward aging
                var oSortedNewsetToOldest = oCleanedPartTrans
                    .OrderByDescending(x => x.TranDate)
                    .ThenByDescending(x => x.TranNum)
                    .ToList();

                // now we can start computing the weighted average age of the inventory by walking through the transactions
                // from newest to oldest and applying the qty received against the remaining inventory balance until we have
                // accounted for the entire balance. If we run out of transactions before we run out of inventory balance,
                // then we will assume the remaining balance is at the age of the last transaction we looked at
                // (or a default max age if we didnt look at any transactions)
                decimal dRemainingQty = oInventoryItem.TotalOnHand;
                decimal dTotalWeightedDays = 0M;
                decimal dTotalUsedQty = 0M;
                decimal dLastAgeDays = 3650M; // default to 10 years if we have no transactions to look at

                foreach (HSPartReceiptsOrAdjustment oPartTran in oSortedNewsetToOldest)
                {
                    if (oPartTran.WorkingQty <= 0)
                    {
                        continue;
                    }

                    decimal dQtyUsed = Math.Min(oPartTran.WorkingQty, dRemainingQty);

                    decimal dAgeInDays = (decimal)(DateTime.Now - oPartTran.TranDate).TotalDays;
                    dLastAgeDays = dAgeInDays;

                    dTotalWeightedDays += dQtyUsed * dAgeInDays;
                    dTotalUsedQty += dQtyUsed;

                    dRemainingQty -= dQtyUsed;
                    if (dRemainingQty <= 0)
                    {
                        // exit the loop if we have accounted for the entire inventory balance
                        break;
                    }
                }

                // check to see if we have any remaining inventory balance that we couldnt account for with the transactions we looked at
                // if so we assume it is at least as old as the last transaction we looked at (or a default max age if we didnt look at any transactions)
                if (dRemainingQty > 0)
                {
                    dTotalWeightedDays += dRemainingQty * dLastAgeDays;
                    dTotalUsedQty += dRemainingQty;
                }

                decimal dAverageAge = dTotalUsedQty > 0 ? dTotalWeightedDays / dTotalUsedQty : 3650M;

                oInventoryItem.MissingReceipts = dRemainingQty > 0;
                oInventoryItem.AverageAgeInDays = dAverageAge;

                // now we need to check if this part is used in any finished goods and if so we need to make sure the age of the inventory is not greater than the age of the finished goods it is used in
                var oWhereUsed = m_oSimpleWhereUsed.Where(x => x.MtlPartNum == oInventoryItem.PartNum).ToList();
                if (oWhereUsed.Count > 0)
                {
                    // we just get the first finished good
                    oInventoryItem.WhereUsed = oWhereUsed[0].PartNum;
                }
            }
        }

        // Offset removal (oldest -> newest)
        // this helps us to get a more accurate age of the inventory by accounting for adjustments that occur after the receipt of the inventory
        // that can throw off the aging if not accounted for. For example, if we have a receipt of 10 items and then an adjustment that reduces
        // the inventory by 10 items, if we just look at the receipt transaction then we would think we have 10 items that are very old,
        // when in reality we have 0 items. By removing offsetting transactions we can get a more accurate picture of the age of the inventory.
        private static List<HSPartReceiptsOrAdjustment> RemoveOffsettingTransactions(List<HSPartReceiptsOrAdjustment> oPartTrans)
        {
            var oPositives = new Queue<HSPartReceiptsOrAdjustment>();
            var oOldestToNewest = oPartTrans
                .OrderBy(x => x.TranDate)
                .ThenBy(x => x.TranNum)
                .Select(x => new HSPartReceiptsOrAdjustment(x))
                .ToList();

            foreach (HSPartReceiptsOrAdjustment oPartTran in oOldestToNewest)
            {
                if (oPartTran.WorkingQty > 0)
                {
                    oPositives.Enqueue(oPartTran);
                }
                else
                {
                    decimal dQtyToOffset = Math.Abs(oPartTran.WorkingQty);
                    while (dQtyToOffset > 0 && oPositives.Count > 0)
                    {
                        var pos = oPositives.Peek();
                        if (pos.WorkingQty <= dQtyToOffset)
                        {
                            dQtyToOffset -= pos.WorkingQty;
                            oPositives.Dequeue();
                        }
                        else
                        {
                            pos.WorkingQty -= dQtyToOffset;
                            dQtyToOffset = 0;
                        }
                    }
                }
            }

            return oPositives.ToList();
        }

        #endregion

        #region Data Members
        private List<HSInventoryAge> m_oAgeOfAllInventory = new List<HSInventoryAge>();
        private List<HSPartReceiptsOrAdjustment> m_oPartReceiptsAndAdjustments = new List<HSPartReceiptsOrAdjustment>();
        private List<SimpleWhereUsed> m_oSimpleWhereUsed = new List<SimpleWhereUsed>();

        private DateTime m_dtStartDate;
        private const int LOOK_BACK_YEARS = 15;

        #endregion
    }

    public class HSPartReceiptsOrAdjustment
    {
        #region Constructors

        public HSPartReceiptsOrAdjustment(DataRow oDataRow)
        {
            if (oDataRow["PartTran_Company"] != DBNull.Value)
            {
                m_sCompany = (string)oDataRow["PartTran_Company"];
            }
            if (oDataRow["PartTran_TranNum"] != DBNull.Value)
            {
                m_iTranNum = (int)oDataRow["PartTran_TranNum"];
            }
            if (oDataRow["PartTran_PartNum"] != DBNull.Value)
            {
                m_sPartNum = (string)oDataRow["PartTran_PartNum"];
                m_sPartNum = m_sPartNum.ToUpper();
            }
            if (oDataRow["PartTran_TranClass"] != DBNull.Value)
            {
                m_sTranClass = (string)oDataRow["PartTran_TranClass"];
            }
            if (oDataRow["PartTran_TranType"] != DBNull.Value)
            {
                m_sTranType = (string)oDataRow["PartTran_TranType"];
            }
            if (oDataRow["PartTran_TranDate"] != DBNull.Value)
            {
                m_dtTranDate = (DateTime)oDataRow["PartTran_TranDate"];
            }
            if (oDataRow["PartTran_TranQty"] != DBNull.Value)
            {
                m_dTranQty = (decimal)oDataRow["PartTran_TranQty"];
                m_dWorkingQty = m_dTranQty;
            }
            if (oDataRow["PartTran_LotNum"] != DBNull.Value)
            {
                m_sLotNumber = (string)oDataRow["PartTran_LotNum"];
                m_sLotNumber = m_sLotNumber.ToUpper();
            }
            if (oDataRow["PartTranSNTran_SerialNumber"] != DBNull.Value)
            {
                m_sSerialNumber = (string)oDataRow["PartTranSNTran_SerialNumber"];
                m_sSerialNumber = m_sSerialNumber.ToUpper();
            }
        }

        public HSPartReceiptsOrAdjustment(HSPartReceiptsOrAdjustment source)
        {
            m_sCompany = source.Company;
            m_iTranNum = source.TranNum;
            m_sPartNum = source.PartNum;
            m_sTranClass = source.TranClass;
            m_sTranType = source.TranType;
            m_dtTranDate = source.TranDate;
            m_dTranQty = source.TranQty;
            m_dWorkingQty = source.TranQty;
            m_sLotNumber = source.LotNumber;
            m_sSerialNumber = source.SerialNumber;
        }

        #endregion

        #region Properties

        public decimal TranQty => m_dTranQty;

        public decimal WorkingQty
        {
            get { return m_dWorkingQty; }
            set { m_dWorkingQty = value; }
        }

        public DateTime TranDate => m_dtTranDate;
        public string Company => m_sCompany;
        public int TranNum => m_iTranNum;
        public string PartNum => m_sPartNum;
        public string TranClass => m_sTranClass;
        public string TranType => m_sTranType;
        public string LotNumber => m_sLotNumber;
        public string SerialNumber => m_sSerialNumber;

        #endregion

        #region Data Members

        private string m_sCompany;
        private int m_iTranNum;
        private string m_sPartNum;
        private string m_sTranClass;
        private string m_sTranType;
        private DateTime m_dtTranDate;
        private decimal m_dTranQty;
        private decimal m_dWorkingQty;
        private string m_sLotNumber;
        private string m_sSerialNumber;

        #endregion
    }

    public class HSInventoryAge
    {
        #region Constructors
        public HSInventoryAge(DataRow oDataRow)
        {
            if (oDataRow["Part_Company"] != DBNull.Value)
            {
                m_sCompany = (string)oDataRow["Part_Company"];
            }
            if (oDataRow["Part_PartNum"] != DBNull.Value)
            {
                m_sPartNum = (string)oDataRow["Part_PartNum"];
                m_sPartNum = m_sPartNum.ToUpper();
            }
            if (oDataRow["Part_PartDescription"] != DBNull.Value)
            {
                m_sPartDescription = (string)oDataRow["Part_PartDescription"];
            }
            if (oDataRow["Part_TypeCode"] != DBNull.Value)
            {
                m_sTypeCode = (string)oDataRow["Part_TypeCode"];
            }
            if (oDataRow["Part_ClassID"] != DBNull.Value)
            {
                m_sClassId = (string)oDataRow["Part_ClassID"];
            }
            if (oDataRow["PartClass_Description"] != DBNull.Value)
            {
                m_sClassDescription = (string)oDataRow["PartClass_Description"];
            }
            if (oDataRow["Part_IUM"] != DBNull.Value)
            {
                m_sIUM = (string)oDataRow["Part_IUM"];
            }
            if (oDataRow["Part_CostMethod"] != DBNull.Value)
            {
                m_sCostMethod = (string)oDataRow["Part_CostMethod"];
            }
            if (oDataRow["Calculated_CostPer"] != DBNull.Value)
            {
                m_dCostPer = (decimal)oDataRow["Calculated_CostPer"];
            }
            if (oDataRow["Calculated_QtyOnHand"] != DBNull.Value)
            {
                m_dTotalOnHand = (decimal)oDataRow["Calculated_QtyOnHand"];
            }
            if (oDataRow["Calculated_ExtCost"] != DBNull.Value)
            {
                m_dExtendedCost = (decimal)oDataRow["Calculated_ExtCost"];
            }
            if (oDataRow["Part_TrackLots"] != DBNull.Value)
            {
                m_bTrackLots = (bool)oDataRow["Part_TrackLots"];
            }
            if (oDataRow["PartBin_LotNum"] != DBNull.Value)
            {
                m_sLotNum = (string)oDataRow["PartBin_LotNum"];
                m_sLotNum = m_sLotNum.ToUpper();
            }
            if (oDataRow["Part_TrackSerialNum"] != DBNull.Value)
            {
                m_bTrackSerialNum = (bool)oDataRow["Part_TrackSerialNum"];
            }
            if (oDataRow["Calculated_SerialNumber"] != DBNull.Value)
            {
                m_sSerialNumber = (string)oDataRow["Calculated_SerialNumber"];
                m_sSerialNumber = m_sSerialNumber.ToUpper();
            }
            if (oDataRow["Calculated_AgeInDays"] != DBNull.Value)
            {
                m_dAverageNumberOfDays = (decimal)oDataRow["Calculated_AgeInDays"];
            }
        }
        #endregion

        #region Properties
        public string Company
        {
            get { return m_sCompany; }
        }
        public string PartNum
        {
            get { return m_sPartNum; }
        }
        public string PartDescription
        {
            get { return m_sPartDescription; }
        }
        public string TypeCode
        {
            get { return m_sTypeCode; }
        }
        public string ClassId
        {
            get { return m_sClassId; }
        }
        public string ClassDescription
        {
            get { return m_sClassDescription; }
        }
        public string IUM
        {
            get { return m_sIUM; }
        }
        public string CostMethod
        {
            get { return m_sCostMethod; }
        }
        public decimal CostPer
        {
            get { return m_dCostPer; }
        }
        public decimal TotalOnHand
        {
            get { return m_dTotalOnHand; }
        }
        public decimal ExtendedCost
        {
            get { return m_dExtendedCost; }
        }
        public bool TrackLots
        {
            get { return m_bTrackLots; }
        }
        public string LotNum
        {
            get { return m_sLotNum; }
        }
        public bool TrackSerialNum
        {
            get { return m_bTrackSerialNum; }
        }
        public string SerialNumber
        {
            get { return m_sSerialNumber; }
        }

        public decimal AverageAgeInDays
        {
            get { return m_dAverageNumberOfDays; }
            set { m_dAverageNumberOfDays = value; }
        }

        public bool MissingReceipts
        {
            get { return m_bMissingReceipts; }
            set { m_bMissingReceipts = value; }
        }

        public string WhereUsed
        {
            get { return m_sWhereUsed; }
            set { m_sWhereUsed = value; }
        }
        #endregion

        #region Data Members
        private string m_sCompany;
        private string m_sPartNum;
        private string m_sPartDescription;
        private string m_sTypeCode;
        private string m_sClassId;
        private string m_sClassDescription;
        private string m_sIUM;
        private string m_sCostMethod;
        private decimal m_dCostPer;
        private decimal m_dTotalOnHand;
        private decimal m_dExtendedCost;
        private bool m_bTrackLots;
        private string m_sLotNum;
        private bool m_bTrackSerialNum;
        private string m_sSerialNumber;
        private bool m_bMissingReceipts;
        private decimal m_dAverageNumberOfDays;
        private string m_sWhereUsed;

        private const decimal DEFAULT_AVERAGE_NUMBER_OF_DAYS = 365 * 10;

        #endregion
    }

    public class SimpleWhereUsed
    {
        #region Constructors
        public SimpleWhereUsed(DataRow oDataRow)
        {
            if (oDataRow["PartRev_PartNum"] != DBNull.Value)
            {
                m_sPartNum = (string)oDataRow["PartRev_PartNum"];
                m_sPartNum = m_sPartNum.ToUpper();
            }
            if (oDataRow["PartRev_RevisionNum"] != DBNull.Value)
            {
                m_sRevNum = (string)oDataRow["PartRev_RevisionNum"];
                m_sRevNum = m_sRevNum.ToUpper();
            }
            if (oDataRow["PartRev_Approved"] != DBNull.Value)
            {
                m_bApproved = (bool)oDataRow["PartRev_Approved"];
            }
            if (oDataRow["PartMtl_MtlSeq"] != DBNull.Value)
            {
                m_iMtlSeq = (int)oDataRow["PartMtl_MtlSeq"];
            }
            if (oDataRow["PartMtl_MtlPartNum"] != DBNull.Value)
            {
                m_sMtlPartNum = (string)oDataRow["PartMtl_MtlPartNum"];
                m_sMtlPartNum = m_sMtlPartNum.ToUpper();
            }
            if (oDataRow["PartMtl_QtyPer"] != DBNull.Value)
            {
                m_dQtyPer = (decimal)oDataRow["PartMtl_QtyPer"];
            }
        }
        #endregion

        #region Properties
 
        public string PartNum
        {
            get { return m_sPartNum; }
            set { m_sPartNum = value; }
        }
        public string RevNum
        {
            get { return m_sRevNum; }
                set { m_sRevNum = value; }
        }
        public bool Approved
        {
            get { return m_bApproved; }
            set { m_bApproved = value; }
        }
        public int MtlSeq
        {
            get { return m_iMtlSeq; }
            set { m_iMtlSeq = value; }
        }
        public string MtlPartNum
        {
            get { return m_sMtlPartNum; }
            set { m_sMtlPartNum = value; }
        }
        public decimal QtyPer
        {
            get { return m_dQtyPer; }
            set { m_dQtyPer = value; }
        }
        #endregion

        #region Data Members
        private string m_sPartNum;
        private string m_sRevNum;
        private bool m_bApproved;
        private int m_iMtlSeq;
        private string m_sMtlPartNum;
        private decimal m_dQtyPer;
        #endregion
    }

    public class InventoryKey
    {
        #region Constructors
        public InventoryKey(string part, string lot, string serial)
        {
            PartNum = part ?? "";
            LotNum = lot ?? "";
            SerialNumber = serial ?? "";
        }
        #endregion

        #region Methods
        public override bool Equals(object obj)
        {
            if (obj is InventoryKey k)
                return PartNum == k.PartNum && LotNum == k.LotNum && SerialNumber == k.SerialNumber;
            return false;
        }

        public override int GetHashCode()
        {
            return (PartNum + "|" + LotNum + "|" + SerialNumber).GetHashCode();
        }
        #endregion

        #region Data Memebers
        public string PartNum;
        public string LotNum;
        public string SerialNumber;
        #endregion
    }
}
