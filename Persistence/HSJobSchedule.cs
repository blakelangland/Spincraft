using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using System.IO;

using Erp.BO;
using Erp.Proxy.BO;
using Erp.Adapters;

using Ice.Core;
using Ice.Lib.Framework;
using Ice.Lib.Searches;
using Ice.Adapters;
using Ice.Proxy.BO;
using Ice.BO;

using System.Drawing;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Drawing.Spreadsheet;
using SpreadsheetLight;
using SpreadsheetLight.Charts;

using HSPersistence;

namespace HorizonScientific
{
    public class HSOperationDetail
    {
        public HSOperationDetail()
        {
        }

        public HSOperationDetail(DataRow oDataRow)
        {
            if (oDataRow["JobOper_Company"] != DBNull.Value)
            {
                m_sCompany = (string)oDataRow["JobOper_Company"];
            }
            if (oDataRow["JobOper_JobNum"] != DBNull.Value)
            {
                m_sJobNum = (string)oDataRow["JobOper_JobNum"];
            }
            if (oDataRow["JobHead1_PartNum"] != DBNull.Value)
            {
                m_sPartNum = (string)oDataRow["JobHead1_PartNum"];
            }
            if (oDataRow["JobOper_JobComplete"] != DBNull.Value)
            {
                m_bJobComplete = (bool)oDataRow["JobOper_JobComplete"];
            }
            if (oDataRow["JobOper_AssemblySeq"] != DBNull.Value)
            {
                m_iAssemblySequence = (int)oDataRow["JobOper_AssemblySeq"];
            }
            if (oDataRow["JobOper_OprSeq"] != DBNull.Value)
            {
                m_iOperationSequence = (int)oDataRow["JobOper_OprSeq"];
            }
            if (oDataRow["JobOper_OpCode"] != DBNull.Value)
            {
                m_sOperationCode = (string)oDataRow["JobOper_OpCode"];
            }
            if (oDataRow["JobOper_OpComplete"] != DBNull.Value)
            {
                m_bOperationComplete = (bool)oDataRow["JobOper_OpComplete"];
            }
            if (oDataRow["JobOper_RunQty"] != DBNull.Value)
            {
                m_dRunQuantity = (decimal)oDataRow["JobOper_RunQty"];
            }
            if (oDataRow["JobOper_QtyCompleted"] != DBNull.Value)
            {
                m_dQuantityCompleted = (decimal)oDataRow["JobOper_QtyCompleted"];
            }
            if (oDataRow["JobOper_StartDate"] != DBNull.Value)
            {
                m_dtStartDate = (DateTime)oDataRow["JobOper_StartDate"];
            }
            if (oDataRow["JobOper_DueDate"] != DBNull.Value)
            {
                m_dtDueDate = (DateTime)oDataRow["JobOper_DueDate"];
            }
        }

        #region Properties

        public string Company
        {
            get { return m_sCompany; }
            set { m_sCompany = value; }
        }

        public string JobNum
        {
            get { return m_sJobNum; }
            set { m_sJobNum = value; }
        }

        public string PartNum
        {
            get { return m_sPartNum; }
            set { m_sPartNum = value; }
        }

        public bool JobComplete
        {
            get { return m_bJobComplete; }
            set { m_bJobComplete = value; }
        }

        public int AssemblySequence
        {
            get { return m_iAssemblySequence; }
            set { m_iAssemblySequence = value; }
        }

        public int OperationSequence
        {
            get { return m_iOperationSequence; }
            set { m_iOperationSequence = value; }
        }
        public string OperationCode
        {
            get { return m_sOperationCode; }
            set { m_sOperationCode = value; }
        }
        public bool OperationComplete
        {
            get { return m_bOperationComplete; }
            set { m_bOperationComplete = value; }
        }
        public decimal RunQuantity
        {
            get { return m_dRunQuantity; }
            set { m_dRunQuantity = value; }
        }
        public decimal QuantityCompleted
        {
            get { return m_dQuantityCompleted; }
            set { m_dQuantityCompleted = value; }
        }
        public DateTime StartDate
        {
            get { return m_dtStartDate; }
            set { m_dtStartDate = value; }
        }
        public DateTime DueDate
        {
            get { return m_dtDueDate; }
            set { m_dtDueDate = value; }
        }
        #endregion

        #region Data Members

        private string m_sCompany;
        private string m_sJobNum;
        private string m_sPartNum;
        private bool m_bJobComplete;
        private int m_iAssemblySequence;
        private int m_iOperationSequence;
        private string m_sOperationCode;
        private bool m_bOperationComplete;
        private decimal m_dRunQuantity;
        private decimal m_dQuantityCompleted;
        private DateTime m_dtStartDate;
        private DateTime m_dtDueDate;

        #endregion
    }

    public class HSOpCode
    {
        public HSOpCode(int iAssemblySequence, int iOperationSequence, string sOperationCode)
        {
            m_iAssemblySequence = iAssemblySequence;
            m_iOperationSequence = iOperationSequence;
            m_sOperationCode = sOperationCode;
        }

        public int AssemblySequence
        {
            get { return m_iAssemblySequence; }
            set { m_iAssemblySequence = value; }
        }

        public int OperationSequence
        {
            get { return m_iOperationSequence; }
            set { m_iOperationSequence = value; }
        }
        public string OperationCode
        {
            get { return m_sOperationCode; }
            set { m_sOperationCode = value; }
        }

        public int ColumnPosition
        {
            get { return m_iColumnPosition; }
            set { m_iColumnPosition = value; }
        }
        public string Expression
        {
            get { return m_iAssemblySequence.ToString() + "-" + m_iOperationSequence.ToString() + "-" + m_sOperationCode; }
        }

        private int m_iAssemblySequence;
        private int m_iOperationSequence;
        private string m_sOperationCode;
        private int m_iColumnPosition;
    }

    public class HSJobSchedule
    {
        public HSJobSchedule()
        {
        }

        static public List<HSOperationDetail> Initialize(Session oSession, string sPartNum)
        {
            List<HSOperationDetail> oOperationDetails = new List<HSOperationDetail>();

            //****Set a parameter Value***** 
            Ice.Proxy.BO.DynamicQueryImpl oDynamicQuery = WCFServiceSupport.CreateImpl<Ice.Proxy.BO.DynamicQueryImpl>(oSession, Ice.Proxy.BO.DynamicQueryImpl.UriPath);
            QueryExecutionDataSet oQueryExecutionDataSet = oDynamicQuery.GetQueryExecutionParametersByID(BAQConstants.QUERY_JOB_OPERATION_DETAILS_BY_PART);
            foreach (DataRow oParameter in oQueryExecutionDataSet.ExecutionParameter)
            {
                if (string.Compare(oParameter["ParameterID"].ToString(), "PartNum") == 0)
                {
                    oParameter["ParameterValue"] = sPartNum;
                }
            }
            oQueryExecutionDataSet.AcceptChanges();
            DataSet oDataSet = oDynamicQuery.ExecuteByID(BAQConstants.QUERY_JOB_OPERATION_DETAILS_BY_PART, oQueryExecutionDataSet);
            foreach (DataRow oRow in oDataSet.Tables[0].Rows)
            {
                oOperationDetails.Add(new HSOperationDetail(oRow));
            }

            return oOperationDetails;
        }

        public static void CreateAndSendReport(string sTmpFileDirectory, HSUser oRequestingUser, List<HSOperationDetail> oOperationDetails)
        {
            // get the file name
            DateTime dtToday = DateTime.Now;
            string sNameOfMonth = dtToday.ToString("MMMM");

            // get the first part number -- should be the same for everything as this list is generated by part num
            string sPartNum = "";
            if (oOperationDetails.Count > 0)
            {
                sPartNum = oOperationDetails[0].PartNum;
            }

            // these are all for the same type of job so we just get the first operation detail in the list
            HSOperationDetail oFirstOperationDetail = oOperationDetails[0];
            string sDestinationFileName = sTmpFileDirectory + "\\JobSchedule-" + oFirstOperationDetail.PartNum.ToString() + ".xlsx";

            // if the file is already there then get rid of it as we are reprocessing for some reason
            if (File.Exists(sDestinationFileName) == true)
            {
                try
                {
                    File.Delete(sDestinationFileName);
                }
                catch (Exception)
                {
                    // another process may be accessing it so just skip this for now
                }
            }

            SLDocument oSLDocument = new SLDocument();

            // set up the style of cells
            SLStyle oGood = oSLDocument.CreateStyle();
            oGood.ApplyNamedCellStyle(SLNamedCellStyleValues.Good);

            SLStyle oNeutrual = oSLDocument.CreateStyle();
            oNeutrual.ApplyNamedCellStyle(SLNamedCellStyleValues.Neutral);

            SLStyle oBad = oSLDocument.CreateStyle();
            oBad.ApplyNamedCellStyle(SLNamedCellStyleValues.Bad);

            SLStyle oBold = oSLDocument.CreateStyle();
            oBold.SetFontBold(true);

            SLStyle oBoldRed = oSLDocument.CreateStyle();
            oBoldRed.SetFontBold(true);
            oBoldRed.SetFontColor(System.Drawing.Color.Red);

            SLStyle oBoldUnderline = oSLDocument.CreateStyle();
            oBoldUnderline.SetFontBold(true);
            oBoldUnderline.SetFontUnderline(UnderlineValues.Single);

            SLStyle oUnderline = oSLDocument.CreateStyle();
            oUnderline.SetFontUnderline(UnderlineValues.Single);

            SLStyle oBoldCenter = oSLDocument.CreateStyle();
            oBoldCenter.SetFontBold(true);
            SLAlignment oCenterAlignment = new SLAlignment();
            oCenterAlignment.Horizontal = HorizontalAlignmentValues.Center;
            oBoldCenter.Alignment = oCenterAlignment;

            SLStyle oCenter = oSLDocument.CreateStyle();
            oCenter.Alignment = oCenterAlignment;

            SLStyle oUnderlineCenetered = oSLDocument.CreateStyle();
            oUnderlineCenetered.SetFontUnderline(UnderlineValues.Single);
            oUnderlineCenetered.Alignment = oCenterAlignment;

            SLStyle oRight = oSLDocument.CreateStyle();
            SLAlignment oRightAlignment = new SLAlignment();
            oRightAlignment.Horizontal = HorizontalAlignmentValues.Right;
            oRight.Alignment = oRightAlignment;

            SLFill oSLFill = new SLFill();
            oSLFill.SetPatternBackgroundColor(SLThemeColorIndexValues.Dark2Color);

            // create a box style
            SLStyle oBoxStyleCentered = new SLStyle();
            oBoxStyleCentered.SetFontBold(true);
            oBoxStyleCentered.Alignment = oCenterAlignment;
            oBoxStyleCentered.SetTopBorder(BorderStyleValues.Thin, SLThemeColorIndexValues.Dark2Color);
            oBoxStyleCentered.SetBottomBorder(BorderStyleValues.Thin, SLThemeColorIndexValues.Dark2Color);
            oBoxStyleCentered.SetLeftBorder(BorderStyleValues.Thin, SLThemeColorIndexValues.Dark2Color);
            oBoxStyleCentered.SetRightBorder(BorderStyleValues.Thin, SLThemeColorIndexValues.Dark2Color);
            oBoxStyleCentered.SetFontColor(SLThemeColorIndexValues.Dark2Color);
            oBoxStyleCentered.FormatCode = "$#,##0.00";
            oBoxStyleCentered.Fill = oSLFill;

            SLStyle oBoldCurrencyStyle = new SLStyle();
            oBoldCurrencyStyle.SetFontBold(true);
            oBoldCurrencyStyle.ApplyNamedCellStyle(SLNamedCellStyleValues.Currency);
            oBoldCurrencyStyle.Alignment = oRightAlignment;
            oBoldCurrencyStyle.SetTopBorder(BorderStyleValues.Thick, SLThemeColorIndexValues.Dark2Color);
            oBoldCurrencyStyle.SetBottomBorder(BorderStyleValues.Thick, SLThemeColorIndexValues.Dark2Color);
            oBoldCurrencyStyle.SetLeftBorder(BorderStyleValues.Thick, SLThemeColorIndexValues.Dark2Color);
            oBoldCurrencyStyle.SetRightBorder(BorderStyleValues.Thick, SLThemeColorIndexValues.Dark2Color);
            oBoldCurrencyStyle.FormatCode = "$#,##0.00";
            oBoldCurrencyStyle.Fill = oSLFill;

            SLStyle oCurrencyStyleCentered = new SLStyle();
            oCurrencyStyleCentered.ApplyNamedCellStyle(SLNamedCellStyleValues.Currency);
            oCurrencyStyleCentered.Alignment = oCenterAlignment;
            oCurrencyStyleCentered.FormatCode = "$#,##0.00";

            SLStyle oCurrencyStyleBoldCentered = new SLStyle();
            oCurrencyStyleBoldCentered.SetFontBold(true);
            oCurrencyStyleBoldCentered.ApplyNamedCellStyle(SLNamedCellStyleValues.Currency);
            oCurrencyStyleBoldCentered.Alignment = oCenterAlignment;
            oCurrencyStyleBoldCentered.FormatCode = "$#,##0.00";

            SLStyle oDecimalBoxStyleBoldCentered = new SLStyle();
            oDecimalBoxStyleBoldCentered.SetFontBold(true);
            oDecimalBoxStyleBoldCentered.ApplyNamedCellStyle(SLNamedCellStyleValues.Currency);
            oDecimalBoxStyleBoldCentered.SetTopBorder(BorderStyleValues.Thin, SLThemeColorIndexValues.Dark2Color);
            oDecimalBoxStyleBoldCentered.SetBottomBorder(BorderStyleValues.Thin, SLThemeColorIndexValues.Dark2Color);
            oDecimalBoxStyleBoldCentered.SetLeftBorder(BorderStyleValues.Thin, SLThemeColorIndexValues.Dark2Color);
            oDecimalBoxStyleBoldCentered.SetRightBorder(BorderStyleValues.Thin, SLThemeColorIndexValues.Dark2Color);
            oDecimalBoxStyleBoldCentered.Alignment = oCenterAlignment;
            oDecimalBoxStyleBoldCentered.FormatCode = "###.00";

            SLStyle oPercentStyleBoldCentered = new SLStyle();
            oPercentStyleBoldCentered.SetFontBold(true);
            oPercentStyleBoldCentered.ApplyNamedCellStyle(SLNamedCellStyleValues.Currency);
            oPercentStyleBoldCentered.Alignment = oCenterAlignment;
            oPercentStyleBoldCentered.FormatCode = "###.00%";

            bool bFirstWorksheet = true;
            if (oOperationDetails.Count > 0)
            {
                if (bFirstWorksheet == true)
                {
                    oSLDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, sPartNum);
                    bFirstWorksheet = false;
                }
                else
                {
                    oSLDocument.AddWorksheet(sPartNum);
                }

                // take all operations tied to a job and put them in a list where the job num has a list of all operation details
                var oAllOperationsForJobs = oOperationDetails.GroupBy(p => p.JobNum);

                // get a list of all operations that exist for all of these jobs -- generally its the same set of operations but there can be ops added for rework etc.
                List<HSOpCode> oAllOperations = new List<HSOpCode>();
                foreach (HSOperationDetail oOp in oOperationDetails)
                {
                    // the unique name for the operation is the assembly sequence and the operation sequence
                    HSOpCode oOpCode = new HSOpCode(oOp.AssemblySequence, oOp.OperationSequence, oOp.OperationCode);
                    HSOpCode oTmp = oAllOperations.FirstOrDefault(z => (z.AssemblySequence == oOpCode.AssemblySequence) && (z.OperationSequence == oOpCode.OperationSequence) && (string.Compare(z.OperationCode, oOpCode.OperationCode, true) == 0));
                    if (oTmp == null)
                    {
                        oAllOperations.Add(oOpCode);
                    }
                }
                // now we put this in order by assembly sequence and then operation sequence
                oAllOperations = oAllOperations.OrderBy(x => x.AssemblySequence).ThenBy(y => y.OperationSequence).ToList();
                // now we will set the column number for each operation
                int iColumnPosition = 4;
                foreach (HSOpCode oTmp in oAllOperations)
                {
                    oTmp.ColumnPosition = iColumnPosition++;
                }
                //set up column headers
                int iAColumn = 1;
                int iBColumn = 2;
                int iCColumn = 3;
                int iDColumn = 4;
                int iNumOfRows = 1;

                oSLDocument.SetCellValue(iNumOfRows, iAColumn, "Company");
                oSLDocument.SetCellStyle(iNumOfRows, iAColumn, oBoldCenter);
                oSLDocument.SetColumnWidth(iNumOfRows, iAColumn, 15);
                oSLDocument.SetCellValue(iNumOfRows, iBColumn, "Job");
                oSLDocument.SetCellStyle(iNumOfRows, iBColumn, oBoldCenter);
                oSLDocument.SetColumnWidth(iNumOfRows, iBColumn, 15);
                oSLDocument.SetCellValue(iNumOfRows, iCColumn, "Part");
                oSLDocument.SetCellStyle(iNumOfRows, iCColumn, oBoldCenter);
                oSLDocument.SetColumnWidth(iNumOfRows, iCColumn, 15);
                oSLDocument.SetCellValue(iNumOfRows, iDColumn, "Job Complete");
                oSLDocument.SetCellStyle(iNumOfRows, iDColumn, oBoldCenter);
                oSLDocument.SetColumnWidth(iNumOfRows, iDColumn, 15);
                // we will add a column per unique operation
                int iCurrentColumn = 5;
                foreach (HSOpCode oTmp in oAllOperations)
                {
                    oSLDocument.SetCellValue(iNumOfRows, iCurrentColumn, oTmp.Expression);
                    oSLDocument.SetCellStyle(iNumOfRows, iCurrentColumn, oBoldCenter);
                    oSLDocument.SetColumnWidth(iNumOfRows, iCurrentColumn, 15);
                    // SET CELL WIDTH TO 15
                    oTmp.ColumnPosition = iCurrentColumn++;
                }

                // lets freeze the job info and the op info
                oSLDocument.FreezePanes(1, 4);

                // walk through each job
                foreach (var oJob in oAllOperationsForJobs)
                {
                    iNumOfRows++;

                    foreach (HSOpCode oTmp in oAllOperations)
                    {
                        // we will set each column to NA and then overwrite this below if the operation has a value for this op code
                        oSLDocument.SetCellValue(iNumOfRows, oTmp.ColumnPosition, OPERATION_NOT_APPLICABLE_TO_JOB);
                    }

                    // just get any operation in this list to fill in the first few columns
                    bool bFirstOperation = true;

                    // walk through each operation
                    foreach (var oOperation in oJob)
                    {
                        if (bFirstOperation == true)
                        {
                            oSLDocument.SetCellValue(iNumOfRows, iAColumn, oOperation.Company);
                            oSLDocument.SetCellStyle(iNumOfRows, iAColumn, oCenter);
                            oSLDocument.SetCellValue(iNumOfRows, iBColumn, oOperation.JobNum);
                            oSLDocument.SetCellStyle(iNumOfRows, iBColumn, oCenter);
                            oSLDocument.SetCellValue(iNumOfRows, iCColumn, oOperation.PartNum);
                            oSLDocument.SetCellStyle(iNumOfRows, iCColumn, oCenter);
                            oSLDocument.SetCellValue(iNumOfRows, iDColumn, oOperation.JobComplete);
                            oSLDocument.SetCellStyle(iNumOfRows, iDColumn, oCenter);
                            bFirstOperation = false;
                        }

                        // now we will try to extract the appropriate operation for each column
                        HSOpCode oTmp = oAllOperations.FirstOrDefault(z => (z.AssemblySequence == oOperation.AssemblySequence) && (z.OperationSequence == oOperation.OperationSequence) && (string.Compare(z.OperationCode, oOperation.OperationCode, true) == 0));
                        if (oTmp != null)
                        {
                            oSLDocument.SetCellValue(iNumOfRows, oTmp.ColumnPosition, oOperation.DueDate.ToShortDateString());
                            if (oOperation.OperationComplete == true)
                            {
                                // check to see if quantity is correct
                                if (oOperation.QuantityCompleted == oOperation.RunQuantity)
                                {
                                    oSLDocument.SetCellStyle(iNumOfRows, oTmp.ColumnPosition, oGood);
                                }
                                else
                                {
                                    // quantity not right even though operation is completed
                                    oSLDocument.SetCellStyle(iNumOfRows, oTmp.ColumnPosition, oNeutrual);
                                }
                            }
                            else
                            {
                                oSLDocument.SetCellStyle(iNumOfRows, oTmp.ColumnPosition, oBad);
                            }
                        }
                    }
                }

                // save off the excel spreadsheet
                oSLDocument.SaveAs(sDestinationFileName);

                // email customer the statement
                List<string> oToAddresses = new List<string>();
                oToAddresses.Add(HSUser.SPINCRAFT_ROOT_USER);
                List<string> oAttachments = new List<string>();
                if (File.Exists(sDestinationFileName) == true)
                {
                    oAttachments.Add(sDestinationFileName);
                }

                if ((oRequestingUser != null) && (File.Exists(sDestinationFileName) == true))
                {
                    // we are emailing the file to a user of Epicor
                    oToAddresses.Add(oRequestingUser.Email);
                    HSEmailHelper.SendEmail(oToAddresses, "Job Schedule For Part " + oFirstOperationDetail.PartNum.ToString(), "Job Schedule For Part " + oFirstOperationDetail.PartNum.ToString(), oAttachments);
                }
            }
        }

        static public List<HSOperationDetail> UpdateJobOperations(Session oSession, string sOriginalFileName, HSUser oRequestingUser, string sTmpFileFolder, List<HSOperationDetail> oOriginalOperationDetails)
        {
            List<HSOperationDetail> oOperationsDetails = new List<HSOperationDetail>();

            // delete any existing files with JobScheduleUpdates in the name that are in the temp folder
            string[] oAllFiles = System.IO.Directory.GetFiles(sTmpFileFolder);
            foreach (string sFile in oAllFiles)
            {
                if (sFile.Contains("JobScheduleUpdates") == true)
                {
                    File.Delete(sFile);
                }
            }

            if (File.Exists(sOriginalFileName) == true)
            {
                // we will make a copy of this in the temp folder so we can delete the original file
                FileInfo oTmpFile = new FileInfo(sOriginalFileName);

                string sFileName = sTmpFileFolder + "\\" + oTmpFile.Name;
                File.Move(sOriginalFileName, sFileName);

                oOperationsDetails = ReadOperationDetails(sFileName);

                // now we will group these by job as we need to process them this way for efficiency
                // take all operations tied to a job and put them in a list where the job num has a list of all operation details
                var oAllOperationsForJobs = oOperationsDetails.GroupBy(p => p.JobNum);

                List<string> oErrorMessages = new List<string>();
                foreach (var oJob in oAllOperationsForJobs)
                {
                    List<string> oErrors = UpdateAllOperationsForJob(oSession, oJob, oOriginalOperationDetails);
                    if (oErrors.Count != 0)
                    {
                        oErrorMessages.AddRange(oErrors);
                    }
                }

                // report out to user whether the job schedule was updated successfully or not
                List<string> oToAddresses = new List<string>();
                oToAddresses.Add(HSUser.SPINCRAFT_ROOT_USER);
                if (oRequestingUser != null)
                {
                    oToAddresses.Add(oRequestingUser.Email);
                }
                StringBuilder oStringBuilder = new StringBuilder();
                oStringBuilder.Clear();

                if (oErrorMessages.Count > 0)
                {
                    oStringBuilder.Append("The file named " + sOriginalFileName + " was had erorrs while processing\n");
                    foreach (string sError in oErrorMessages)
                    {
                        oStringBuilder.Append(sError + "\n");
                    }
                }
                else
                {
                    oStringBuilder.Append("The file named " + sOriginalFileName + " was processed successfully\n");
                }
                HSEmailHelper.SendEmail(oToAddresses, "Job Schedule Updates Processed", oStringBuilder.ToString());

            }
            return oOperationsDetails;
        }

        private static List<HSOperationDetail> ReadOperationDetails(string sFileName)
        {
            List<HSOperationDetail> oOperationsDetails = new List<HSOperationDetail>();
            List<HSOpCode> oOpCodes = new List<HSOpCode>();

            SLExcelData oData = (new SLExcelReader()).ReadFirstSheetInExcel(sFileName);

            int iCounter = 0;
            foreach (List<string> sRowData in oData.DataRows)
            {
                int iColumn = 0;

                // if this is the first row then we need to first read in the header with the op codes
                if (iCounter == 0)
                {
                    iColumn = 0;
                    foreach (string sCell in oData.Headers)
                    {
                        if (iColumn >= 4)
                        {
                            string sOpCodeInfo = sCell.Trim();
                            // now we pull appart this string
                            string[] OpCodeParts = sOpCodeInfo.Split('-');
                            if (OpCodeParts.Length == 3)
                            {
                                int iAssemblySequence;
                                if (int.TryParse(OpCodeParts[0], out iAssemblySequence) == true)
                                {
                                    int iOperationSequence;
                                    if (int.TryParse(OpCodeParts[1], out iOperationSequence) == true)
                                    {
                                        string sOpCode = OpCodeParts[2];
                                        HSOpCode oOpCode = new HSOpCode(iAssemblySequence, iOperationSequence, sOpCode);
                                        oOpCode.ColumnPosition = iColumn;
                                        oOpCodes.Add(oOpCode);
                                    }
                                }

                            }
                        }
                        iColumn++;
                    }
                }

                // we are reading in operation details
                iColumn = 0;
                string sCompany = "";
                string sJobNum = "";
                string sPartNum = "";
                bool bJobComplete = false;
                DateTime dtDueDate = DateTime.MinValue;
                foreach (string sCell in sRowData)
                {
                    if (string.IsNullOrEmpty(sCell) == false)
                    {
                        // there are an unlimited number of columns
                        // the first column will be company, job num, then part num, then job complete
                        // after that there is a column per operation .. unsure how many operations there may be
                        if (iColumn == 0)
                        {
                            sCompany = sCell.Trim();
                        }
                        else if (iColumn == 1)
                        {
                            sJobNum = sCell.Trim();
                        }
                        else if (iColumn == 2)
                        {
                            sPartNum = sCell.Trim();
                        }
                        else if (iColumn == 3)
                        {
                            bool.TryParse(sCell.Trim(), out bJobComplete);
                        }
                        else
                        {
                            // if we are here then we are adding a new operation

                            // this column is the operation due date or it has the text "N/A"
                            if (string.Compare(sCell.Trim(), OPERATION_NOT_APPLICABLE_TO_JOB, true) == 0)
                            {
                                // we can throw this operation out as it does not technically exist
                                iColumn++;
                                continue;
                            }
                            else
                            {
                                dtDueDate = DateTime.MinValue;
                                try
                                {
                                    double dDateTime = Convert.ToDouble(sCell);
                                    dtDueDate = DateTime.FromOADate(dDateTime);
                                }
                                catch (Exception e)
                                {
                                    dtDueDate = DateTime.MinValue;
                                    if (DateTime.TryParse(sCell, out dtDueDate) == false)
                                    {
                                        // couldn't get a good date so we skip this
                                        iColumn++;
                                        continue;
                                    }
                                }
                            }

                            // if we get here we can add a new operation
                            HSOperationDetail oTmp = new HSOperationDetail();
                            oTmp.Company = sCompany;
                            oTmp.JobNum = sJobNum;
                            oTmp.PartNum = sPartNum;
                            oTmp.JobComplete = bJobComplete;
                            // we get the assembly sequence, op sequence, and op code from the name of the column
                            HSOpCode oTmpOpCode = oOpCodes.FirstOrDefault(x => x.ColumnPosition == iColumn);
                            if (oTmp != null)
                            {
                                oTmp.AssemblySequence = oTmpOpCode.AssemblySequence;
                                oTmp.OperationSequence = oTmpOpCode.OperationSequence;
                                oTmp.OperationCode = oTmpOpCode.OperationCode;
                                oTmp.DueDate = dtDueDate;

                                oOperationsDetails.Add(oTmp);
                            }
                        }
                    }
                    iColumn++;
                }
                iCounter++;
            }

            return oOperationsDetails;
        }

        static public List<string> UpdateAllOperationsForJob(Session oSession, IGrouping<string,  HSOperationDetail> oAllOperationsForJob, List<HSOperationDetail> oOriginalOperationDetails)
        {
            List<string> oErrorMessages = new List<string>();
            // get all operations for this job
            List<HSOperationDetail> oAllOperations = oAllOperationsForJob.ToList();
            List<HSOperationDetail> oAllOriginalOperations = new List<HSOperationDetail>();
            bool bAnyChanges = false;
            string sJobNum = "";
            foreach (HSOperationDetail oOperationDetail in oAllOperations)
            {
                // all job numbers are the same but we just need to capture the job number somewhere for later reference
                sJobNum = oOperationDetail.JobNum;
                HSOperationDetail oOriginalOperation = oOriginalOperationDetails.FirstOrDefault(x => (string.Compare(x.Company, oOperationDetail.Company, true) == 0) && (string.Compare(x.JobNum, oOperationDetail.JobNum, true) == 0) &&
                          (x.AssemblySequence == oOperationDetail.AssemblySequence) && (x.OperationSequence == oOperationDetail.OperationSequence) && (string.Compare(x.OperationCode, oOperationDetail.OperationCode, true) == 0));
                if (oOriginalOperation != null)
                {
                    oAllOriginalOperations.Add(oOriginalOperation);
                    // see if any changes have beenn mde
                    if (oOperationDetail.DueDate != oOriginalOperation.DueDate)
                    {
                        bAnyChanges = true;
                    }
                }
                else
                {
                    oErrorMessages.Add("Could not find Job: " + oOperationDetail.JobNum + " Assembly Sequence: " + oOperationDetail.AssemblySequence.ToString() + " Operation Sequence: " + oOperationDetail.OperationSequence.ToString() + " Operation Code: " + oOperationDetail.OperationCode);
                }
            }

            // first see if there are any changes required for this job
            if (bAnyChanges == true)
            {
                // on job head we need to unrelease and unegineer the job
                bool bReleased;
                bool bEngineered;
                string sError = OpenJobHeadForChange(oSession, sJobNum, out bReleased, out bEngineered);
                if (string.IsNullOrEmpty(sError) == false)
                {
                    oErrorMessages.Add(sError);
                }

                // we can now change any opertions that need to be updated
                List<string> oTmpErrorMessages = ManuallyUpdateOperationsForJob(oSession, oAllOperations, oAllOriginalOperations);
                if (oTmpErrorMessages.Count!= 0)
                {
                    oErrorMessages.AddRange(oTmpErrorMessages);
                }

                // we now set the engineer and release status of the job back to what it was before we updated the operations
                sError = CloseJobHeadForChange(oSession, sJobNum, bReleased, bEngineered);
                if (string.IsNullOrEmpty(sError) == false)
                {
                    oErrorMessages.Add(sError);
                }
            }

            return oErrorMessages;
        }

        static public string UpdateOperationViaScheduler(Session oSession, HSOperationDetail oOperationDetail, HSOperationDetail oOriginalOperationDetail)
        {
            // update the job operations and use the scheduler to set the due dates
            string sErrorMessages = "";

            // see if there was any changes made to the due date
            if (oOperationDetail.DueDate != oOriginalOperationDetail.DueDate)
            {
                try
                {
                    // now update the job operation due date and possibly start date
                    ScheduleEngineImpl oScheduleEngineImpl = WCFServiceSupport.CreateImpl<ScheduleEngineImpl>(oSession, Erp.Proxy.BO.ScheduleEngineImpl.UriPath);
                    ScheduleEngineDataSet ds = new ScheduleEngineDataSet();
                    ScheduleEngineDataSet.ScheduleEngineRow row = ds.ScheduleEngine.NewScheduleEngineRow();
                    row.Company = oSession.CompanyID;
                    row.JobNum = oOperationDetail.JobNum;
                    row.AssemblySeq = oOperationDetail.AssemblySequence;
                    row.OprSeq = oOperationDetail.OperationSequence;
                    row.OpDtlSeq = 0;
                    // we may change the start date if it needs to be moved to be before the due date
                    if (oOriginalOperationDetail.StartDate > oOperationDetail.DueDate)
                    {
                        row.StartDate = oOperationDetail.DueDate;
                        row.StartTime = 0;
                    }
                    else
                    {
                        // leave original start date
                        row.StartDate = oOriginalOperationDetail.StartDate;
                        row.StartTime = 0;

                    }
                    row.EndDate = oOperationDetail.DueDate;
                    row.EndTime = 0;
                    row.WhatIf = false;
                    row.Finite = false;
                    row.SchedTypeCode = "oo"; // move operation only
                    row.ScheduleDirection = "End";
                    row.SetupComplete = false;
                    row.ProductionComplete = false;
                    row.OverrideMtlCon = true;
                    row.OverRideHistDateSetting = 2;
                    row.RecalcExpProdYld = false;
                    ds.ScheduleEngine.AddScheduleEngineRow(row);
                    bool l_finished;
                    string c_WarnLogTxt;
                    oScheduleEngineImpl.MoveJobItem(ds, out l_finished, out c_WarnLogTxt);
                }
                catch (Exception e)
                {
                    sErrorMessages = e.Message;
                }
            }
            return sErrorMessages;
        }

        static public string OpenJobHeadForChange(Session oSession, string sJobNum, out bool bJobReleased, out bool bJobEngineered)
        {
            // update the job operations
            string sErrorMessages = "";
            bJobReleased = false;
            bJobEngineered = false;

            JobEntryImpl oJobEntryImpl = WCFServiceSupport.CreateImpl<JobEntryImpl>(oSession, Erp.Proxy.BO.JobEntryImpl.UriPath);
            try
            {
                // load up the job
                var oDataSet = oJobEntryImpl.GetByID(sJobNum);
                if ((oDataSet != null) && (oDataSet.JobHead != null) && (oDataSet.JobHead.Count == 1))
                {
                    DataRow dr = oDataSet.JobHead[0];

                    // retain if the job was laready engineered and released or not
                    bJobReleased = (bool)dr["JobReleased"];
                    bJobEngineered = (bool)dr["JobEngineered"];

                    // force the job to be unreleased and unengineered
                    dr.BeginEdit();
                    dr["JobReleased"] = false;
                    dr["JobEngineered"] = false;
                    dr.EndEdit();
                    oJobEntryImpl.Update(oDataSet);
                }
            }
            catch (Exception e)
            {
                sErrorMessages = e.Message;
            }
            finally
            {
                oJobEntryImpl.Dispose();
            }
            return sErrorMessages;
        }

        static public string CloseJobHeadForChange(Session oSession, string sJobNum, bool bJobReleased, bool bJobEngineered)
        {
            // update the job operations
            string sErrorMessages = "";

            try
            {
                JobEntryImpl oJobEntryImpl = WCFServiceSupport.CreateImpl<JobEntryImpl>(oSession, Erp.Proxy.BO.JobEntryImpl.UriPath);
                // load up the job
                var oDataSet = oJobEntryImpl.GetByID(sJobNum);
                if ((oDataSet != null) && (oDataSet.JobHead != null) && (oDataSet.JobHead.Count == 1))
                {
                    DataRow dr = oDataSet.JobHead[0];
                    dr.BeginEdit();
                    dr["JobReleased"] = bJobReleased;
                    dr["JobEngineered"] = bJobEngineered;
                    dr["ChangeDescription"] = "Changed due date per job scheduling tool.";
                    dr.EndEdit();
                    oJobEntryImpl.Update(oDataSet);
                    oJobEntryImpl.Dispose();
                }
            }
            catch (Exception e)
            {
                sErrorMessages = e.Message;
            }

            return sErrorMessages;
        }

        static public List<string> ManuallyUpdateOperationsForJob(Session oSession, List<HSOperationDetail> oOperationDetails, List<HSOperationDetail> oOriginalOperationDetails)
        {
            // update the job operations
            List<string> oErrorMessages = new List<string>();

            // see if there was any changes made to the due date
            foreach (HSOperationDetail oOperationDetail in oOperationDetails)
            {
                // get the original operation detail to see if it has been modified
                HSOperationDetail oOriginalOperationDetail = oOriginalOperationDetails.FirstOrDefault(x => (string.Compare(x.Company, oOperationDetail.Company, true) == 0) && (string.Compare(x.JobNum, oOperationDetail.JobNum, true) == 0) &&
                    (x.AssemblySequence == oOperationDetail.AssemblySequence) && (x.OperationSequence == oOperationDetail.OperationSequence) && (string.Compare(x.OperationCode, oOperationDetail.OperationCode, true) == 0));
                if ((oOriginalOperationDetail != null) && (oOperationDetail.DueDate != oOriginalOperationDetail.DueDate))
                {
                    try
                    {
                        JobEntryImpl oJobEntryImpl = WCFServiceSupport.CreateImpl<JobEntryImpl>(oSession, Erp.Proxy.BO.JobEntryImpl.UriPath);
                        // load up the job
                        var oDataSet = oJobEntryImpl.GetByID(oOriginalOperationDetail.JobNum);
                        if ((oDataSet != null) && (oDataSet.JobOper != null))
                        {
                            // get the correct operation
                            foreach (DataRow dr in oDataSet.JobOper.Rows)
                            {
                                //AssemblySeq - int
                                //OpCode - string
                                //OprSeq - int
                                int iAssemblySeq = (int)dr["AssemblySeq"];
                                int iOperationSequence = (int)dr["OprSeq"];
                                string sOpCode = (string)dr["OpCode"];
                                if ((iAssemblySeq == oOperationDetail.AssemblySequence) && (iOperationSequence == oOperationDetail.OperationSequence) && (string.Compare(sOpCode, oOperationDetail.OperationCode, true) == 0))
                                {
                                    dr.BeginEdit();

                                    // check to see if we need to reset the start date
                                    if (oOriginalOperationDetail.StartDate > oOperationDetail.DueDate)
                                    {
                                        dr["StartDate"] = oOperationDetail.DueDate;
                                        dr["StartTime"] = 0;
                                    }
                                    dr["DueDate"] = oOperationDetail.DueDate;

                                    dr.EndEdit();
                                    oJobEntryImpl.Update(oDataSet);
                                    oJobEntryImpl.Dispose();
                                    // found the one and only operation to update so we get out of loop
                                    // and move to the next operation to update
                                    break;
                                }
                            }
                        }
                    }
                    catch (Exception e)
                    {
                        oErrorMessages.Add(e.Message);
                    }
                }
            }
            return oErrorMessages;
        }

        private static string OPERATION_NOT_APPLICABLE_TO_JOB = "N/A";
    }

    public class HSMasterSchedule
    {
        public HSMasterSchedule()
        {
        }

        public HSMasterSchedule(DataRow oDataRow)
        {
            if (oDataRow["JobOper_JobNum"] != DBNull.Value)
            {
                m_sJobNum = (string)oDataRow["JobOper_JobNum"];
            }
            if (oDataRow["JobHead_PartNum"] != DBNull.Value)
            {
                m_sPartNum = (string)oDataRow["JobHead_PartNum"];
            }
            if (oDataRow["JobOper_OprSeq"] != DBNull.Value)
            {
                m_iOperationSequence = (int)oDataRow["JobOper_OprSeq"];
            }
            if (oDataRow["JobOper_OpCode"] != DBNull.Value)
            {
                m_sOperationCode = (string)oDataRow["JobOper_OpCode"];
            }
            if (oDataRow["JobOper_OpDesc"] != DBNull.Value)
            {
                m_sOperationDescription = (string)oDataRow["JobOper_OpDesc"];
            }
            if (oDataRow["JobOper_OpComplete"] != DBNull.Value)
            {
                m_bOperationComplete = (bool)oDataRow["JobOper_OpComplete"];
            }
            if (oDataRow["JobOper_EstProdHours"] != DBNull.Value)
            {
                m_dEstProductionHours = (decimal)oDataRow["JobOper_EstProdHours"];
            }
            if (oDataRow["JobOper_ActProdHours"] != DBNull.Value)
            {
                m_dActualProductionHours = (decimal)oDataRow["JobOper_ActProdHours"];
            }
            if (oDataRow["JobOper_SubContract"] != DBNull.Value)
            {
                m_bSubcontract = (bool)oDataRow["JobOper_SubContract"];
            }
            if (oDataRow["JobOper_VendorNum"] != DBNull.Value)
            {
                m_iVendorNum = (int)oDataRow["JobOper_VendorNum"];
            }
            if (oDataRow["Vendor_VendorID"] != DBNull.Value)
            {
                m_sVendorId = (string)oDataRow["Vendor_VendorID"];
            }
            if (oDataRow["JobHead_JobComplete"] != DBNull.Value)
            {
                m_bJobComplete = (bool)oDataRow["JobHead_JobComplete"];
            }
            if (oDataRow["JobHead_JobCompletionDate"] != DBNull.Value)
            {
                m_dtJobCompletionDate = (DateTime)oDataRow["JobHead_JobCompletionDate"];
            }
            if (oDataRow["JobHead_JobClosed"] != DBNull.Value)
            {
                m_bJobClosed = (bool)oDataRow["JobHead_JobClosed"];
            }
            if (oDataRow["JobHead_ClosedDate"] != DBNull.Value)
            {
                m_dtJobClosedDate = (DateTime)oDataRow["JobHead_ClosedDate"];
            }
            if (oDataRow["JobOper_QtyCompleted"] != DBNull.Value)
            {
                m_dQuantityCompleted = (decimal)oDataRow["JobOper_QtyCompleted"];
            }
        }

        #region Properties

        public string JobNum
        {
            get { return m_sJobNum; }
            set { m_sJobNum = value; }
        }

        public string PartNum
        {
            get { return m_sPartNum; }
            set { m_sPartNum = value; }
        }

        public int OperationSequence
        {
            get { return m_iOperationSequence; }
            set { m_iOperationSequence = value; }
        }
        public string OperationCode
        {
            get { return m_sOperationCode; }
            set { m_sOperationCode = value; }
        }

        public string OperationDescription
        {
            get { return m_sOperationDescription; }
            set { m_sOperationDescription = value; }
        }
        public bool OperationComplete
        {
            get { return m_bOperationComplete; }
            set { m_bOperationComplete = value; }
        }
        public decimal EstProductionHours
        {
            get { return m_dEstProductionHours; }
            set { m_dEstProductionHours = value; }
        }
        public decimal ActualProductionHours
        {
            get { return m_dActualProductionHours; }
            set { m_dActualProductionHours = value; }
        }
        public bool Subcontract
        {
            get { return m_bSubcontract; }
            set { m_bSubcontract = value; }
        }

        public int VendorNum
        {
            get { return m_iVendorNum; }
            set { m_iVendorNum = value; }
        }
        public string VendorId
        {
            get { return m_sVendorId; }
            set { m_sVendorId = value; }
        }
        public bool JobComplete
        {
            get { return m_bJobComplete; }
            set { m_bJobComplete = value; }
        }
        public DateTime JobCompletionDate
        {
            get { return m_dtJobCompletionDate; }
            set { m_dtJobCompletionDate = value; }
        }
        public bool JobClosed
        {
            get { return m_bJobClosed; }
            set { m_bJobClosed = value; }
        }
        public DateTime JobClosedDate
        {
            get { return m_dtJobClosedDate; }
            set { m_dtJobClosedDate = value; }
        }
        public decimal QuantityCompleted
        {
            get { return m_dQuantityCompleted; }
            set { m_dQuantityCompleted = value; }
        }
        #endregion

        #region Data Members

        private string m_sJobNum;
        private string m_sPartNum;
        private int m_iOperationSequence;
        private string m_sOperationCode;
        private string m_sOperationDescription;
        private bool m_bOperationComplete;
        private decimal m_dEstProductionHours;
        private decimal m_dActualProductionHours;
        private bool m_bSubcontract;
        private int m_iVendorNum;
        private string m_sVendorId;
        private bool m_bJobComplete;
        private DateTime m_dtJobCompletionDate;
        private bool m_bJobClosed;
        private DateTime m_dtJobClosedDate;
        private decimal m_dQuantityCompleted;

        #endregion
    }
}
