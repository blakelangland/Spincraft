using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml;
using System.Text.RegularExpressions;
using System.IO;


namespace HSPersistence
{
    public class SLExcelStatus
    {
        public string Message { get; set; }
        public bool Success
        {
            get { return string.IsNullOrWhiteSpace(Message); }
        }
    }

    public class SLExcelData
    {
        public SLExcelStatus Status { get; set; }
        public Columns ColumnConfigurations { get; set; }
        public List<string> Headers { get; set; }
        public List<List<string>> DataRows { get; set; }
        public string SheetName { get; set; }

        public SLExcelData()
        {
            Status = new SLExcelStatus();
            Headers = new List<string>();
            DataRows = new List<List<string>>();
        }
    }

    public class SLExcelReader
    {

        private string GetColumnName(string cellReference)
        {
            Regex oRegEx = new Regex("[A-Za-z]+");
            Match oMatch = oRegEx.Match(cellReference);

            return oMatch.Value;
        }

        private int ConvertColumnNameToNumber(string sColumnName)
        {
            Regex oAlpha = new Regex("^[A-Z]+$");
            if (!oAlpha.IsMatch(sColumnName))
            {
                throw new ArgumentException();
            }

            char[] oColumnLetters = sColumnName.ToCharArray();
            Array.Reverse(oColumnLetters);

            int iConvertedValue = 0;
            for (int i = 0; i < oColumnLetters.Length; i++)
            {
                char cLetter = oColumnLetters[i];
                // ASCII 'A' = 65
                int iCurrent = i == 0 ? cLetter - 65 : cLetter - 64;
                iConvertedValue += iCurrent * (int)Math.Pow(26, i);
            }

            return iConvertedValue;
        }

        private IEnumerator<Cell> GetExcelCellEnumerator(Row oRow)
        {
            int iCurrentCount = 0;
            foreach (Cell oCell in oRow.Descendants<Cell>())
            {
                string sColumnName = GetColumnName(oCell.CellReference);

                int iCurrentColumnIndex = ConvertColumnNameToNumber(sColumnName);

                for (; iCurrentCount < iCurrentColumnIndex; iCurrentCount++)
                {
                    var oEmptyCell = new Cell()
                    {
                        DataType = null,
                        CellValue = new CellValue(string.Empty)
                    };
                    yield return oEmptyCell;
                }

                yield return oCell;
                iCurrentCount++;
            }
        }

        private string ReadExcelCell(Cell oCell, WorkbookPart oWorkbookPart)
        {
            var oCellValue = oCell.CellValue;
            string sText = (oCellValue == null) ? oCell.InnerText : oCellValue.Text;
            if ((oCell.DataType != null) && (oCell.DataType == CellValues.SharedString))
            {
                sText = oWorkbookPart.SharedStringTablePart.SharedStringTable.Elements<SharedStringItem>().ElementAt(Convert.ToInt32(oCell.CellValue.Text)).InnerText;
            }

            return (sText ?? string.Empty).Trim();
        }

        public SLExcelData ReadFirstSheetInExcel(string sFileName)
        {
            SLExcelData oData = new SLExcelData();

            // Open the excel document
            WorkbookPart oWorkbookPart;
            List<Row> oRows;
            SpreadsheetDocument oSpreadsheetDocument = null;
            try
            {
                using (oSpreadsheetDocument = SpreadsheetDocument.Open(sFileName, false))
                {
                    oWorkbookPart = oSpreadsheetDocument.WorkbookPart;

                    IEnumerable<Sheet> oSheets = oWorkbookPart.Workbook.Descendants<Sheet>();
                    Sheet oSheet = oSheets.First();
                    oData.SheetName = oSheet.Name;

                    Worksheet oWorksheet = ((WorksheetPart)oWorkbookPart.GetPartById(oSheet.Id)).Worksheet;
                    Columns oColumns = oWorksheet.Descendants<Columns>().FirstOrDefault();
                    oData.ColumnConfigurations = oColumns;

                    var sheetData = oWorksheet.Elements<SheetData>().First();
                    oRows = sheetData.Elements<Row>().ToList();

                    // Read the header
                    if (oRows.Count > 0)
                    {
                        Row oRow = oRows[0];
                        IEnumerator<Cell> oCellEnumerator = GetExcelCellEnumerator(oRow);
                        while (oCellEnumerator.MoveNext())
                        {
                            Cell oCell = oCellEnumerator.Current;
                            string sText = ReadExcelCell(oCell, oWorkbookPart).Trim();
                            oData.Headers.Add(sText);
                        }
                    }
                    // Read the sheet data
                    if (oRows.Count > 1)
                    {
                        for (var i = 1; i < oRows.Count; i++)
                        {
                            List<string> oDataRows = new List<string>();
                            oData.DataRows.Add(oDataRows);
                            Row oRow = oRows[i];
                            IEnumerator<Cell> oCellEnumerator = GetExcelCellEnumerator(oRow);
                            while (oCellEnumerator.MoveNext())
                            {
                                Cell oCell = oCellEnumerator.Current;
                                string sText = ReadExcelCell(oCell, oWorkbookPart).Trim();
                                oDataRows.Add(sText);
                            }
                        }
                    }

                    if (oSpreadsheetDocument != null)
                    {
                        oSpreadsheetDocument.Close();
                        oSpreadsheetDocument = null;
                    }
                }
            }
            catch (Exception)
            {
                oData.Status.Message = "Unable to open the file";
                return oData;
            }

            return oData;
        }

        public List<SLExcelData> ReadAllExcelSheets(string sFileName)
        {
            List<SLExcelData> oAllData = new List<SLExcelData>();

            SLExcelData oData = new SLExcelData();
            // Open the excel document
            WorkbookPart oWorkbookPart;
            SpreadsheetDocument oSpreadsheetDocument = null;
            List<Row> oRows;
            try
            {
                oSpreadsheetDocument = SpreadsheetDocument.Open(sFileName, false);
                oWorkbookPart = oSpreadsheetDocument.WorkbookPart;

                IEnumerable<Sheet> oSheets = oWorkbookPart.Workbook.Descendants<Sheet>();
                foreach (Sheet oSheet in oSheets)
                {
                    oData = new SLExcelData();
                    oData.SheetName = oSheet.Name;

                    Worksheet oWorksheet = ((WorksheetPart)oWorkbookPart.GetPartById(oSheet.Id)).Worksheet;
                    Columns oColumns = oWorksheet.Descendants<Columns>().FirstOrDefault();
                    oData.ColumnConfigurations = oColumns;

                    var sheetData = oWorksheet.Elements<SheetData>().First();
                    oRows = sheetData.Elements<Row>().ToList();

                    // Read the header
                    if (oRows.Count > 0)
                    {
                        Row oRow = oRows[0];
                        IEnumerator<Cell> oCellEnumerator = GetExcelCellEnumerator(oRow);
                        while (oCellEnumerator.MoveNext())
                        {
                            Cell oCell = oCellEnumerator.Current;
                            string sText = ReadExcelCell(oCell, oWorkbookPart).Trim();
                            oData.Headers.Add(sText);
                        }
                    }
                    // Read the sheet data
                    if (oRows.Count > 1)
                    {
                        for (var i = 1; i < oRows.Count; i++)
                        {
                            List<string> oDataRows = new List<string>();
                            oData.DataRows.Add(oDataRows);
                            Row oRow = oRows[i];
                            IEnumerator<Cell> oCellEnumerator = GetExcelCellEnumerator(oRow);
                            while (oCellEnumerator.MoveNext())
                            {
                                Cell oCell = oCellEnumerator.Current;
                                string sText = ReadExcelCell(oCell, oWorkbookPart).Trim();
                                oDataRows.Add(sText);
                            }
                        }
                    }
                    oAllData.Add(oData);
                }
            }
            catch (Exception)
            {
                oData.Status.Message = "Unable to open the file";
            }
            finally
            {
                if (oSpreadsheetDocument != null)
                {
                    oSpreadsheetDocument.Close();
                    oSpreadsheetDocument = null;
                }
            }
            return oAllData;
        }
    }

    public class SLExcelWriter
    {
        private string ColumnLetter(int iColumn)
        {
            int iFirstLetter = ((iColumn) / 676) + 64;
            int iSecondLetter = ((iColumn % 676) / 26) + 64;
            int iThirdLetter = (iColumn % 26) + 65;

            char cFirstLetter = (iFirstLetter > 64) ? (char)iFirstLetter : ' ';
            char cSecondLetter = (iSecondLetter > 64) ? (char)iSecondLetter : ' ';
            char cThirdLetter = (char)iThirdLetter;

            return string.Concat(cFirstLetter, cSecondLetter, cThirdLetter).Trim();
        }

        private Cell CreateTextCell(string sHeader, UInt32 iIndex, string sText)
        {
            Cell oCell = new Cell
            {
                DataType = CellValues.InlineString,
                CellReference = sHeader + iIndex
            };

            InlineString oInlineString = new InlineString();
            Text oText = new Text { Text = sText };
            oInlineString.AppendChild(oText);
            oCell.AppendChild(oInlineString);
            return oCell;
        }

        public byte[] GenerateExcel(SLExcelData oData)
        {
            MemoryStream oStream = new MemoryStream();
            using (SpreadsheetDocument oDocument = SpreadsheetDocument.Create(oStream, SpreadsheetDocumentType.Workbook))
            {

                WorkbookPart oWorkbookpart = oDocument.AddWorkbookPart();
                oWorkbookpart.Workbook = new Workbook();
                WorksheetPart oWorksheetPart = oWorkbookpart.AddNewPart<WorksheetPart>();
                SheetData oSheetData = new SheetData();

                oWorksheetPart.Worksheet = new Worksheet(oSheetData);

                Sheets oSheets = oDocument.WorkbookPart.Workbook.AppendChild<Sheets>(new Sheets());

                var sheet = new Sheet()
                {
                    Id = oDocument.WorkbookPart.GetIdOfPart(oWorksheetPart),
                    SheetId = 1,
                    Name = oData.SheetName ?? "Sheet 1"
                };
                oSheets.AppendChild(sheet);

                // Add header
                UInt32 iRowIdex = 0;
                Row oRow = new Row { RowIndex = ++iRowIdex };
                oSheetData.AppendChild(oRow);
                int iCellIndex = 0;

                foreach (string sHeader in oData.Headers)
                {
                    oRow.AppendChild(CreateTextCell(ColumnLetter(iCellIndex++), iRowIdex, sHeader ?? string.Empty));
                }
                if (oData.Headers.Count > 0)
                {
                    // Add the column configuration if available
                    if (oData.ColumnConfigurations != null)
                    {
                        Columns oColumns = (Columns)oData.ColumnConfigurations.Clone();
                        oWorksheetPart.Worksheet.InsertAfter(oColumns, oWorksheetPart.Worksheet.SheetFormatProperties);
                    }
                }

                // Add sheet data
                foreach (List<string> oRowData in oData.DataRows)
                {
                    iCellIndex = 0;
                    oRow = new Row { RowIndex = ++iRowIdex };
                    oSheetData.AppendChild(oRow);
                    foreach (string sCellData in oRowData)
                    {
                        Cell oCell = CreateTextCell(ColumnLetter(iCellIndex++), iRowIdex, sCellData ?? string.Empty);
                        oRow.AppendChild(oCell);
                    }
                }

                oWorkbookpart.Workbook.Save();
                oDocument.Close();
            }
            return oStream.ToArray();
        }
    }

    public class SLExcelHelper
    {
        public static void AddStandardFormattingAndStyles(WorkbookPart oWorkbookPart)
        {
            // add the styles
            Stylesheet oStyleSheet = new Stylesheet();

            CellFormat integerFormatting = new CellFormat();
            integerFormatting.NumberFormatId = 1;
            integerFormatting.ApplyNumberFormat = true;

            CellFormat doubleFormatting = new CellFormat();
            doubleFormatting.NumberFormatId = 2;
            doubleFormatting.ApplyNumberFormat = true;

            CellFormat moneyFormatting = new CellFormat();
            moneyFormatting.NumberFormatId = 4;
            moneyFormatting.ApplyNumberFormat = true;

            CellFormat dateFormatting = new CellFormat();
            dateFormatting.NumberFormatId = 14;
            dateFormatting.ApplyNumberFormat = true;

            CellFormats oCellFormats = new CellFormats();
            oCellFormats.Append(integerFormatting);
            oCellFormats.Append(doubleFormatting);
            oCellFormats.Append(moneyFormatting);
            oCellFormats.Append(dateFormatting);

            oStyleSheet.CellFormats = oCellFormats;

            oStyleSheet.Borders = new Borders();
            oStyleSheet.Borders.Append(new Border());
            oStyleSheet.Fills = new Fills();
            oStyleSheet.Fills.Append(new DocumentFormat.OpenXml.Spreadsheet.Fill());
            oStyleSheet.Fonts = new DocumentFormat.OpenXml.Spreadsheet.Fonts();
            oStyleSheet.Fonts.Append(new Font());

            oWorkbookPart.AddNewPart<WorkbookStylesPart>();
            oWorkbookPart.WorkbookStylesPart.Stylesheet = oStyleSheet;

            CellStyles oCellStyles = new CellStyles();
            CellStyle oCellStyle = new CellStyle();
            oCellStyle.FormatId = 0;
            oCellStyle.BuiltinId = 0;
            oCellStyles.Append(oCellStyle);
            oCellStyles.Count = UInt32Value.FromUInt32((uint)oCellStyles.ChildElements.Count);
            oStyleSheet.Append(oCellStyles);
        }

        public static Cell ConstructCell(string sValue)
        {
            Cell oCell = new Cell();
            oCell.DataType = CellValues.InlineString;
            oCell.InlineString = new InlineString
            {
                Text = new Text(sValue)
            };
            return oCell;
        }

        public static Cell ConstructCell(int iValue)
        {
            Cell oCell = new Cell();
            oCell.DataType = CellValues.Number;
            oCell.CellValue = new CellValue(iValue.ToString());
            oCell.StyleIndex = 0;
            return oCell;
        }

        public static Cell ConstructCell(double dValue)
        {
            Cell oCell = new Cell();
            oCell.DataType = CellValues.Number;
            oCell.CellValue = new CellValue(dValue.ToString());
            oCell.StyleIndex = 1;
            return oCell;
        }

        public static Cell ConstructCell(decimal dValue)
        {
            Cell oCell = new Cell();
            oCell.DataType = CellValues.Number;
            oCell.CellValue = new CellValue(dValue.ToString());
            oCell.StyleIndex = 2;
            return oCell;
        }

        public static Cell ConstructCell(bool bValue)
        {
            Cell oCell = new Cell();
            oCell.DataType = CellValues.String;
            oCell.CellValue = new CellValue();
            oCell.CellValue.Text = Convert.ToString(bValue);
            return oCell;
        }

        public static Cell ConstructCell(DateTime dtValue)
        {
            Cell oCell = new Cell();
            oCell.DataType = new EnumValue<CellValues>(CellValues.Date);
            oCell.CellValue = new CellValue(dtValue.ToString("yyyy-MM-dd"));
            oCell.StyleIndex = 3;
            return oCell;
        }
    }
}
