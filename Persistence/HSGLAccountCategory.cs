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
    public class HSGLAccountCategories
    {
        #region Constructors

        public HSGLAccountCategories()
        {
        }

        #endregion

        #region Methods
        public bool Initialize(Session oSession)
        {
            bool bSuccess = true;

            // read in all GL Segments
            Ice.Proxy.BO.DynamicQueryImpl oDynamicQuery = WCFServiceSupport.CreateImpl<Ice.Proxy.BO.DynamicQueryImpl>(oSession, Ice.Proxy.BO.DynamicQueryImpl.UriPath);
            QueryExecutionDataSet oQueryExecutionDataSet = oDynamicQuery.GetQueryExecutionParametersByID(BAQConstants.QUERY_GL_SEGMENTS);
            oQueryExecutionDataSet.ExecutionParameter.Clear();
            DataSet oDataSet = oDynamicQuery.ExecuteByID(BAQConstants.QUERY_GL_SEGMENTS, oQueryExecutionDataSet);
            m_oAllSegments.Clear();
            foreach (DataRow oDataRow in oDataSet.Tables[0].Rows)
            {
                m_oAllSegments.Add(new HSGLSegment(oDataRow));
            }

            //GLControls
            oDynamicQuery = WCFServiceSupport.CreateImpl<Ice.Proxy.BO.DynamicQueryImpl>(oSession, Ice.Proxy.BO.DynamicQueryImpl.UriPath);
            oQueryExecutionDataSet = oDynamicQuery.GetQueryExecutionParametersByID(BAQConstants.QUERY_GL_CONTROLS);
            oQueryExecutionDataSet.ExecutionParameter.Clear();
            oDataSet = oDynamicQuery.ExecuteByID(BAQConstants.QUERY_GL_CONTROLS, oQueryExecutionDataSet);
            m_oAllGLControls.Clear();
            foreach (DataRow oDataRow in oDataSet.Tables[0].Rows)
            {
                m_oAllGLControls.Add(new HSGLControl(oDataRow));
            }

            // read in all GL accounts
            oDynamicQuery = WCFServiceSupport.CreateImpl<Ice.Proxy.BO.DynamicQueryImpl>(oSession, Ice.Proxy.BO.DynamicQueryImpl.UriPath);
            oQueryExecutionDataSet = oDynamicQuery.GetQueryExecutionParametersByID(BAQConstants.QUERY_GL_ACCOUNTS);
            oQueryExecutionDataSet.ExecutionParameter.Clear();
            oDataSet = oDynamicQuery.ExecuteByID(BAQConstants.QUERY_GL_ACCOUNTS, oQueryExecutionDataSet);
            m_oAllGLAccounts.Clear();
            foreach (DataRow oDataRow in oDataSet.Tables[0].Rows)
            {
                m_oAllGLAccounts.Add(new HSGLAccount(oDataRow));
            }

            // read in all GL account categories
            oDynamicQuery = WCFServiceSupport.CreateImpl<Ice.Proxy.BO.DynamicQueryImpl>(oSession, Ice.Proxy.BO.DynamicQueryImpl.UriPath);
            oQueryExecutionDataSet = oDynamicQuery.GetQueryExecutionParametersByID(BAQConstants.QUERY_GL_ACCOUNT_CATEGORIES);
            oQueryExecutionDataSet.ExecutionParameter.Clear();
            oDataSet = oDynamicQuery.ExecuteByID(BAQConstants.QUERY_GL_ACCOUNT_CATEGORIES, oQueryExecutionDataSet);
            m_oAllAccountCategories.Clear();
            foreach (DataRow oDataRow in oDataSet.Tables[0].Rows)
            {
                m_oAllAccountCategories.Add(new HSGLAccountCategory(oDataRow, m_oAllGLAccounts));
            }
            // ensure each category has its list of child categories
            SetChildCategories();

            // get the account categories not tied to a parent
            m_oRootAccountCategories = m_oAllAccountCategories.Where(oItem => string.IsNullOrEmpty(oItem.ParentCategoryId)).ToList();
            // order these by sequence
            m_oRootAccountCategories.OrderBy(oItem => oItem.Sequence).ThenBy(x => x.CategoryId).ToList();

            // now we compute the max depth of child categories
            m_iNumberOfChildCategoryLevels = 0;
            if (m_oRootAccountCategories.Count > 0)
            {
                // we definitely have a child so we are at level 1
                m_iNumberOfChildCategoryLevels = 1;
            }
            string sPath = "";
            foreach (HSGLAccountCategory oTmp in m_oRootAccountCategories)
            {
                int iTmp = oTmp.GetMaxDepth(m_iNumberOfChildCategoryLevels, sPath);
                if (iTmp > m_iNumberOfChildCategoryLevels)
                {
                    m_iNumberOfChildCategoryLevels = iTmp;
                }
            }
            return bSuccess;
        }

        private void SetChildCategories()
        {
            // walk through each category and pull out any child categories
            foreach (HSGLAccountCategory oAccountCategory in m_oAllAccountCategories)
            {
                oAccountCategory.SetChildCategories(m_oAllAccountCategories);
            }
        }

        public void GenerateReport(string sCompany, string sTmpFileDirectory, HSUser oRequestingUser)
        {
            #region Setup

            // get the file name
            DateTime dtToday = DateTime.Now;
            string sDestinationFileName = sTmpFileDirectory + sCompany + "-GL Accounts-" + dtToday.Month.ToString() + "-" + dtToday.Day.ToString() + "-" + dtToday.Year.ToString() + ".xlsx";
            int iNumberOfRows = 1;
            int iNumberOfColumns = 1;

            HSUser oServiceAccount = HSUser.GetUserById(HSUser.SPINCRAFT_SERVICE_ACCOUNT_ID);
            List<string> oToAddresses = new List<string>();
            if (oRequestingUser != null)
            {
                // this report was requested by someone directly from Epicor
                oToAddresses.Add(oRequestingUser.Email);
            }
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

            SLStyle oGood = new SLStyle();
            oGood.ApplyNamedCellStyle(SLNamedCellStyleValues.Good);
            SLStyle oNeutral = new SLStyle();
            oNeutral.ApplyNamedCellStyle(SLNamedCellStyleValues.Neutral);
            SLStyle oBad = new SLStyle();
            oBad.ApplyNamedCellStyle(SLNamedCellStyleValues.Bad);

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

            int iStandardColumnWidth = 20;
            #endregion

            #region COA Categories
            if (m_oAllAccountCategories.Count > 0)
            {
                iNumberOfRows = 1;
                iNumberOfColumns = 1;
                if (bFirstWorksheet == true)
                {
                    oSLDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "Account Categories");
                    bFirstWorksheet = false;
                }
                else
                {
                    oSLDocument.AddWorksheet("Account Categories");
                }

                //set column header
                oSLDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Category Id");
                oSLDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Description");
                oSLDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Type");
                oSLDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Parent Category Id");
                oSLDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Net Income");
                oSLDocument.SetColumnWidth(iNumberOfRows++, iNumberOfColumns++, iStandardColumnWidth);

                foreach (HSGLAccountCategory oAccountCategory in m_oAllAccountCategories)
                {
                    oSLDocument.SetCellValue(iNumberOfRows, 1, oAccountCategory.CategoryId);
                    oSLDocument.SetCellValue(iNumberOfRows, 2, oAccountCategory.CategoryDescription);
                    oSLDocument.SetCellValue(iNumberOfRows, 3, oAccountCategory.Type);
                    oSLDocument.SetCellValue(iNumberOfRows, 4, oAccountCategory.ParentCategoryId);
                    oSLDocument.SetCellValue(iNumberOfRows, 5, oAccountCategory.NetIncome);

                    iNumberOfRows++;
                    bDataInReport = true;
                }
            }
            #endregion

            #region COA Segments
            if (m_oAllSegments.Count > 0)
            {
                iNumberOfRows = 1;
                iNumberOfColumns = 1;
                if (bFirstWorksheet == true)
                {
                    oSLDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "Segments");
                    bFirstWorksheet = false;
                }
                else
                {
                    oSLDocument.AddWorksheet("Segments");
                }

                //set column header
                oSLDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Segment Number");
                oSLDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Segment Name");
                oSLDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Segment Abbreviation");
                oSLDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Max Length");
                oSLDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Min Length");
                oSLDocument.SetColumnWidth(iNumberOfRows++, iNumberOfColumns++, iStandardColumnWidth);

                foreach (HSGLSegment oSegment in m_oAllSegments)
                {
                    oSLDocument.SetCellValue(iNumberOfRows, 1, oSegment.SegmentNumber);
                    oSLDocument.SetCellValue(iNumberOfRows, 2, oSegment.SegmentName);
                    oSLDocument.SetCellValue(iNumberOfRows, 3, oSegment.SegmentAbbreviation);
                    oSLDocument.SetCellValue(iNumberOfRows, 4, oSegment.MaximumLength);
                    oSLDocument.SetCellValue(iNumberOfRows, 5, oSegment.MinimumLength);

                    iNumberOfRows++;
                    bDataInReport = true;
                }
            }
            #endregion

            #region GL Controls
            if (m_oAllGLControls.Count > 0)
            {
                iNumberOfRows = 1;
                iNumberOfColumns = 1;
                if (bFirstWorksheet == true)
                {
                    oSLDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "GL Controls");
                    bFirstWorksheet = false;
                }
                else
                {
                    oSLDocument.AddWorksheet("GL Controls");
                }

                //set column header
                oSLDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Book Id");
                oSLDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Control Type");
                oSLDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Control Code");
                oSLDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Account Context");
                oSLDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "GL Account");
                oSLDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Required");
                oSLDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "GL Display");
                oSLDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Account Description");
                oSLDocument.SetColumnWidth(iNumberOfRows++, iNumberOfColumns++, iStandardColumnWidth);

                foreach (HSGLControl oControl in m_oAllGLControls)
                {
                    oSLDocument.SetCellValue(iNumberOfRows, 1, oControl.BookId);
                    oSLDocument.SetCellValue(iNumberOfRows, 2, oControl.ControlType);
                    oSLDocument.SetCellValue(iNumberOfRows, 3, oControl.ControlCode);
                    oSLDocument.SetCellValue(iNumberOfRows, 4, oControl.GLAccountContext);
                    oSLDocument.SetCellValue(iNumberOfRows, 5, oControl.GLAccount);
                    oSLDocument.SetCellValue(iNumberOfRows, 6, oControl.Required);
                    if ((oControl.Required == true) && (string.IsNullOrEmpty(oControl.GLAccount) == true))
                    {
                        oSLDocument.SetCellStyle(iNumberOfRows, 5, oBad);
                        oSLDocument.SetCellStyle(iNumberOfRows, 6, oBad);
                    }
                    else
                    {
                        oSLDocument.SetCellStyle(iNumberOfRows, 5, oGood);
                        oSLDocument.SetCellStyle(iNumberOfRows, 6, oGood);
                    }
                    oSLDocument.SetCellValue(iNumberOfRows, 7, oControl.GLAccountDisplay);
                    oSLDocument.SetCellValue(iNumberOfRows, 8, oControl.GLAccountDescription);

                    iNumberOfRows++;
                    bDataInReport = true;
                }
            }
            #endregion

            #region GL Accounts
            if (m_oRootAccountCategories.Count > 0)
            {
                iNumberOfRows = 1;
                iNumberOfColumns = 1;
                if (bFirstWorksheet == true)
                {
                    oSLDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "GL Accounts");
                    bFirstWorksheet = false;
                }
                else
                {
                    oSLDocument.AddWorksheet("GL Accounts");
                }

                //set column header
                oSLDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Category Id");
                oSLDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "GL Account");
                oSLDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Description");
                oSLDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Active");
                oSLDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Effective From");
                oSLDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Effective To");
                oSLDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Segment1");
                oSLDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Segment2");
                oSLDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, iStandardColumnWidth);
                oSLDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Segment3");
                oSLDocument.SetColumnWidth(iNumberOfRows++, iNumberOfColumns++, iStandardColumnWidth);

                foreach (HSGLAccountCategory oAccountCategory in m_oRootAccountCategories)
                {
                    // first we will print out all of our accounts in this category
                    foreach (HSGLAccount oGLAccount in oAccountCategory.AccountsInCategory)
                    {
                        oSLDocument.SetCellValue(iNumberOfRows, 1, oGLAccount.Category.Path);
                        oSLDocument.SetCellValue(iNumberOfRows, 2, oGLAccount.GLAccount);
                        oSLDocument.SetCellValue(iNumberOfRows, 3, oGLAccount.Description);
                        oSLDocument.SetCellValue(iNumberOfRows, 4, oGLAccount.Active);
                        if (oGLAccount.Active == true)
                        {
                            oSLDocument.SetCellStyle(iNumberOfRows, 4, oGood);
                        }
                        else
                        {
                            oSLDocument.SetCellStyle(iNumberOfRows, 4, oBad);
                        }
                        if (oGLAccount.EffectiveFrom != DateTime.MinValue)
                        {
                            oSLDocument.SetCellValue(iNumberOfRows, 5, oGLAccount.EffectiveFrom.ToShortDateString());
                        }
                        if ((oGLAccount.EffectiveFrom != DateTime.MinValue) && (oGLAccount.EffectiveFrom > DateTime.Now))
                        {
                            oSLDocument.SetCellStyle(iNumberOfRows, 5, oBad);
                        }
                        if (oGLAccount.EffectiveTo != DateTime.MinValue)
                        {
                            oSLDocument.SetCellValue(iNumberOfRows, 6, oGLAccount.EffectiveTo.ToShortDateString());
                        }
                        if ((oGLAccount.EffectiveTo != DateTime.MinValue) && (oGLAccount.EffectiveTo < DateTime.Now))
                        {
                            oSLDocument.SetCellStyle(iNumberOfRows, 6, oBad);
                        }
                        oSLDocument.SetCellValue(iNumberOfRows, 7, oGLAccount.SegmentValue1);
                        oSLDocument.SetCellValue(iNumberOfRows, 8, oGLAccount.SegmentValue2);
                        oSLDocument.SetCellValue(iNumberOfRows, 9, oGLAccount.SegmentValue3);

                        iNumberOfRows++;
                    }

                    // now we will dive into the child categories
                    foreach (HSGLAccountCategory oChildAccountCategory in oAccountCategory.ChildCategories)
                    {
                        oChildAccountCategory.GenerateReport(oSLDocument, ref iNumberOfRows, oGood, oBad);
                    }

                    bDataInReport = true;
                }
            }
            #endregion

            #region Send Report
            oSLDocument.SaveAs(sDestinationFileName);
            // Check to see if we created a file and if so email it to the shipping team to update
            if (File.Exists(sDestinationFileName) == true)
            {
                List<string> oAttachments = new List<string>();
                oAttachments.Add(sDestinationFileName);

                HSEmailHelper.SendEmail(oToAddresses, sCompany + " - GL Account Information", "GL Account Information\n", oAttachments);
            }

            #endregion

        }
        #endregion

        #region Properties
        public List<HSGLAccountCategory> AllAccountCategories
        {
            get { return m_oAllAccountCategories; }
        }
        public List<HSGLAccountCategory> RootAccountCategories
        {
            get { return m_oRootAccountCategories; }
        }
        public List<HSGLSegment> AllSegments
        {
            get { return m_oAllSegments; }
        }
        public List<HSGLAccount> AllGLAccounts
        {
            get { return m_oAllGLAccounts; }
        }
        #endregion

        #region Data Members
        private List<HSGLAccountCategory> m_oAllAccountCategories = new List<HSGLAccountCategory>();
        private List<HSGLAccountCategory> m_oRootAccountCategories = new List<HSGLAccountCategory>();
        private List<HSGLSegment> m_oAllSegments = new List<HSGLSegment>();
        private List<HSGLControl> m_oAllGLControls = new List<HSGLControl>();
        private List<HSGLAccount> m_oAllGLAccounts = new List<HSGLAccount>();

        private int m_iNumberOfChildCategoryLevels;
        #endregion
    }

    public class HSGLAccountCategory
    {
        #region Constructors

        public HSGLAccountCategory(DataRow oDataRow, List<HSGLAccount> oAllGLAccounts)
        {
            if ((oDataRow["COAActCat_Company"] != DBNull.Value) && (string.IsNullOrEmpty((string)oDataRow["COAActCat_Company"]) == false))
            {
                m_sCompany = (string)oDataRow["COAActCat_Company"];
            }
            if ((oDataRow["COAActCat_COACode"] != DBNull.Value) && (string.IsNullOrEmpty((string)oDataRow["COAActCat_COACode"]) == false))
            {
                m_sCOACode = (string)oDataRow["COAActCat_COACode"];
            }
            if ((oDataRow["COAActCat_CategoryID"] != DBNull.Value) && (string.IsNullOrEmpty((string)oDataRow["COAActCat_CategoryID"]) == false))
            {
                m_sCategoryId = (string)oDataRow["COAActCat_CategoryID"];
            }
            if ((oDataRow["COAActCat_Description"] != DBNull.Value) && (string.IsNullOrEmpty((string)oDataRow["COAActCat_Description"]) == false))
            {
                m_sCategoryDescription = (string)oDataRow["COAActCat_Description"];
            }
            if ((oDataRow["COAActCat_Type"] != DBNull.Value) && (string.IsNullOrEmpty((string)oDataRow["COAActCat_Type"]) == false))
            {
                m_sType = (string)oDataRow["COAActCat_Type"];
            }
            if ((oDataRow["COAActCat_NormalBalance"] != DBNull.Value) && (string.IsNullOrEmpty((string)oDataRow["COAActCat_NormalBalance"]) == false))
            {
                m_sNormalBalance = (string)oDataRow["COAActCat_NormalBalance"];
            }
            if ((oDataRow["COAActCat_ParentCategory"] != DBNull.Value) && (string.IsNullOrEmpty((string)oDataRow["COAActCat_ParentCategory"]) == false))
            {
                m_sParentCategoryId = (string)oDataRow["COAActCat_ParentCategory"];
            }
            if (oDataRow["COAActCat_Sequence"] != DBNull.Value)
            {
                m_iSequence = (int)oDataRow["COAActCat_Sequence"];
            }
            if ((oDataRow["COAActCat_ConsolidationType"] != DBNull.Value) && (string.IsNullOrEmpty((string)oDataRow["COAActCat_ConsolidationType"]) == false))
            {
                m_sConsolidationType = (string)oDataRow["COAActCat_ConsolidationType"];
            }
            if (oDataRow["COAActCat_NetIncome"] != DBNull.Value)
            {
                m_bNetIncome = (bool)oDataRow["COAActCat_NetIncome"];
            }

            // we need to pull out all of our GL Accounts
            m_oGLAccountsInCategory = oAllGLAccounts.Where(oItem => string.Compare(CategoryId, oItem.CategoryId, true) == 0).ToList();
            // list these aphabetically
            m_oGLAccountsInCategory = m_oGLAccountsInCategory.OrderBy(oItem => oItem.AccountDescription).ToList();
            // set these to point to this category
            foreach (HSGLAccount oTmp in m_oGLAccountsInCategory)
            {
                oTmp.Category = this;
            }
        }

        #endregion

        #region Methods
        public void SetChildCategories(List<HSGLAccountCategory> oAllCategories)
        {
            m_oChildCategories = oAllCategories.Where(oItem => string.Compare(CategoryId, oItem.ParentCategoryId, true) == 0).ToList();
            // put these in order by sequence number
            m_oChildCategories = m_oChildCategories.OrderBy(oItem => oItem.Sequence).ToList();
        }

        public int GetMaxDepth(int iCurrentDepth, string sPath)
        {
            if (string.IsNullOrEmpty(sPath) == false)
            {
                m_sPath = sPath + " -> " + m_sCategoryId;
            }
            else
            {
                m_sPath = m_sCategoryId;
            }

            if (m_oChildCategories.Count > 0)
            {
                // we have at least one child
                iCurrentDepth += 1;
            }

            foreach (HSGLAccountCategory oTmp in m_oChildCategories)
            {
                int iTmp = oTmp.GetMaxDepth(iCurrentDepth, m_sPath);
                if (iTmp > iCurrentDepth)
                {
                    iCurrentDepth = iTmp;
                }
            }

            return iCurrentDepth;
        }

        public void GenerateReport(SLDocument oSLDocument, ref int iNumberOfRows, SLStyle oGood, SLStyle oBad)
        {
            foreach (HSGLAccountCategory oAccountCategory in m_oChildCategories)
            {
                // first we will print out all of our accounts in this category
                foreach (HSGLAccount oGLAccount in oAccountCategory.AccountsInCategory)
                {
                    oSLDocument.SetCellValue(iNumberOfRows, 1, oGLAccount.Category.Path);
                    oSLDocument.SetCellValue(iNumberOfRows, 2, oGLAccount.GLAccount);
                    oSLDocument.SetCellValue(iNumberOfRows, 3, oGLAccount.Description);
                    oSLDocument.SetCellValue(iNumberOfRows, 4, oGLAccount.Active);
                    if (oGLAccount.Active == true)
                    {
                        oSLDocument.SetCellStyle(iNumberOfRows, 4, oGood);
                    }
                    else
                    {
                        oSLDocument.SetCellStyle(iNumberOfRows, 4, oBad);
                    }
                    if (oGLAccount.EffectiveFrom != DateTime.MinValue)
                    {
                        oSLDocument.SetCellValue(iNumberOfRows, 5, oGLAccount.EffectiveFrom.ToShortDateString());
                    }
                    if ((oGLAccount.EffectiveFrom != DateTime.MinValue) && (oGLAccount.EffectiveFrom > DateTime.Now))
                    {
                        oSLDocument.SetCellStyle(iNumberOfRows, 5, oBad);
                    }
                    if (oGLAccount.EffectiveTo != DateTime.MinValue)
                    {
                        oSLDocument.SetCellValue(iNumberOfRows, 6, oGLAccount.EffectiveTo.ToShortDateString());
                    }
                    if ((oGLAccount.EffectiveTo != DateTime.MinValue) && (oGLAccount.EffectiveTo < DateTime.Now))
                    {
                        oSLDocument.SetCellStyle(iNumberOfRows, 6, oBad);
                    }
                    oSLDocument.SetCellValue(iNumberOfRows, 7, oGLAccount.SegmentValue1);
                    oSLDocument.SetCellValue(iNumberOfRows, 8, oGLAccount.SegmentValue2);
                    oSLDocument.SetCellValue(iNumberOfRows, 9, oGLAccount.SegmentValue3);

                    iNumberOfRows++;
                }

                // now we will dive into the child categories
                foreach (HSGLAccountCategory oChildAccountCategory in oAccountCategory.ChildCategories)
                {
                    oChildAccountCategory.GenerateReport(oSLDocument, ref iNumberOfRows, oGood, oBad);
                }
            }
        }
        #endregion

        #region Properties
        public string Company
        {
            get { return m_sCompany; }
            set { m_sCompany = value; }
        }
        public string COACode
        {
            get { return m_sCOACode; }
            set { m_sCOACode = value; }
        }
        public string CategoryId
        {
            get { return m_sCategoryId; }
            set { m_sCategoryId = value; }
        }
        public string CategoryDescription
        {
            get { return m_sCategoryDescription; }
            set { m_sCategoryDescription = value; }
        }
        public string Type
        {
            get { return m_sType; }
            set { m_sType = value; }
        }
        public string NormalBalance
        {
            get { return m_sNormalBalance; }
            set { m_sNormalBalance = value; }
        }
        public string ParentCategoryId
        {
            get { return m_sParentCategoryId; }
            set { m_sParentCategoryId = value; }
        }
        public int Sequence
        {
            get { return m_iSequence; }
            set { m_iSequence = value; }
        }
        public string ConsolidationType
        {
            get { return m_sConsolidationType; }
            set { m_sConsolidationType = value; }
        }
        public bool NetIncome
        {
            get { return m_bNetIncome; }
            set { m_bNetIncome = value; }
        }

        public List<HSGLAccount> AccountsInCategory
        {
            get { return m_oGLAccountsInCategory; }
        }
        public List<HSGLAccountCategory> ChildCategories
        {
            get { return m_oChildCategories; }
        }
        public string Path
        {
            get { return m_sPath; }
        }
        #endregion

        #region Data Members
        private string m_sCompany;
        private string m_sCOACode;
        private string m_sCategoryId;
        private string m_sCategoryDescription;
        private string m_sType;
        private string m_sNormalBalance;
        private string m_sParentCategoryId;
        private int m_iSequence;
        private string m_sConsolidationType;
        private bool m_bNetIncome;

        private string m_sPath = "";
        private List<HSGLAccount> m_oGLAccountsInCategory = new List<HSGLAccount>();
        private List<HSGLAccountCategory> m_oChildCategories = new List<HSGLAccountCategory>();
        #endregion
    }
}

