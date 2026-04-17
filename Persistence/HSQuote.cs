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
    public class HSQuoteOperations
    {
        #region Constructors
        public HSQuoteOperations()
        {
        }

        public HSQuoteOperations(DataRow oDataRow)
        {
            if (oDataRow["QuoteDtl_Company"] != DBNull.Value)
            {
                m_sCompany = (string)oDataRow["QuoteDtl_Company"];
            }
            if (oDataRow["QuoteDtl_QuoteNum"] != DBNull.Value)
            {
                m_iQuoteNum = (int)oDataRow["QuoteDtl_QuoteNum"];
            }
            if (oDataRow["QuoteDtl_QuoteLine"] != DBNull.Value)
            {
                m_iQuoteLine = (int)oDataRow["QuoteDtl_QuoteLine"];
            }
            if (oDataRow["QuoteQty1_QtyNum"] != DBNull.Value)
            {
                m_iQuoteQtyNum = (int)oDataRow["QuoteQty1_QtyNum"];
            }
            if (oDataRow["QuoteDtl_PartNum"] != DBNull.Value)
            {
                m_sPartNum = (string)oDataRow["QuoteDtl_PartNum"];
            }
            if (oDataRow["QuoteDtl_LineDesc"] != DBNull.Value)
            {
                m_sPartDescription = (string)oDataRow["QuoteDtl_LineDesc"];
            }
            if (oDataRow["QuoteQty_SellingQuantity"] != DBNull.Value)
            {
                m_dSellingQty = (decimal)oDataRow["QuoteQty_SellingQuantity"];
            }
            if (oDataRow["QuoteOpr_OpCode"] != DBNull.Value)
            {
                m_sOpCode = (string)oDataRow["QuoteOpr_OpCode"];
            }
            if (oDataRow["Calculated_TotalProductionHours"] != DBNull.Value)
            {
                m_dTotalProductionHours = (decimal)oDataRow["Calculated_TotalProductionHours"];
            }
            if (oDataRow["Calculated_TotalSetupHours"] != DBNull.Value)
            {
                m_dTotalSetupHours = (decimal)oDataRow["Calculated_TotalSetupHours"];
            }
            if (oDataRow["Calculated_SubcontractCosts"] != DBNull.Value)
            {
                m_dSubcontractCosts = (decimal)oDataRow["Calculated_SubcontractCosts"];
            }
            if (oDataRow["Calculated_BurdenCosts"] != DBNull.Value)
            {
                m_dBurdenCosts = (decimal)oDataRow["Calculated_BurdenCosts"];
            }
            if (oDataRow["Calculated_LaborCosts"] != DBNull.Value)
            {
                m_dLaborCosts = (decimal)oDataRow["Calculated_LaborCosts"];
            }
        }

        #endregion

        #region Properties

        public string Company
        {
            get { return m_sCompany; }
            set { m_sCompany = value; }
        }

        public int QuoteNum
        {
            get { return m_iQuoteNum; }
            set { m_iQuoteNum = value; }
        }

        public int QuoteLine
        {
            get { return m_iQuoteLine; }
            set { m_iQuoteLine = value; }
        }

        public int QuoteQtyNum
        {
            get { return m_iQuoteQtyNum; }
            set { m_iQuoteQtyNum = value; }
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
        public decimal SellingQty
        {
            get { return m_dSellingQty; }
            set { m_dSellingQty = value; }
        }
        public string OpCode
        {
            get { return m_sOpCode; }
            set { m_sOpCode = value; }
        }
        public decimal TotalProductionHours
        {
            get { return m_dTotalProductionHours; }
            set { m_dTotalProductionHours = value; }
        }
        public decimal TotalSetupHours
        {
            get { return m_dTotalSetupHours; }
            set { m_dTotalSetupHours = value; }
        }
        public decimal SubcontractCosts
        {
            get { return m_dSubcontractCosts; }
            set { m_dSubcontractCosts = value; }
        }
        public decimal BurdenCosts
        {
            get { return m_dBurdenCosts; }
            set { m_dBurdenCosts = value; }
        }
        public decimal LaborCosts
        {
            get { return m_dLaborCosts; }
            set { m_dLaborCosts = value; }
        }
        #endregion

        #region Data Members

        private string m_sCompany;
        private int m_iQuoteNum;
        private int m_iQuoteLine;
        private int m_iQuoteQtyNum;
        private string m_sPartNum;
        private string m_sPartDescription;
        private decimal m_dSellingQty;
        private string m_sOpCode;
        private decimal m_dTotalProductionHours;
        private decimal m_dTotalSetupHours;
        private decimal m_dSubcontractCosts;
        private decimal m_dBurdenCosts;
        private decimal m_dLaborCosts;

        #endregion
    }

    public class SF1411Operation
    {
        #region Constructors
        public SF1411Operation()
        {
        }

        public SF1411Operation(DataRow oDataRow)
        {
            if (oDataRow["QuoteDtl_Company"] != DBNull.Value)
            {
                m_sCompany = (string)oDataRow["QuoteDtl_Company"];
            }
            if (oDataRow["QuoteDtl_QuoteNum"] != DBNull.Value)
            {
                m_iQuoteNum = (int)oDataRow["QuoteDtl_QuoteNum"];
            }
            if (oDataRow["QuoteDtl_QuoteLine"] != DBNull.Value)
            {
                m_iQuoteLine = (int)oDataRow["QuoteDtl_QuoteLine"];
            }
            if (oDataRow["QuoteQty1_QtyNum"] != DBNull.Value)
            {
                m_iQuoteQtyNum = (int)oDataRow["QuoteQty1_QtyNum"];
            }
            if (oDataRow["QuoteOpr_OpCode"] != DBNull.Value)
            {
                m_sOpCode = (string)oDataRow["QuoteOpr_OpCode"];
            }
            if (oDataRow["QuoteOpr_PartNum"] != DBNull.Value)
            {
                m_sPartNum = (string)oDataRow["QuoteOpr_PartNum"];
            }
            if (oDataRow["QuoteOpr_Description"] != DBNull.Value)
            {
                m_sPartDescription = (string)oDataRow["QuoteOpr_Description"];
            }
            if (oDataRow["QuoteOpr_AssemblySeq"] != DBNull.Value)
            {
                m_iAssemblySeq = (int)oDataRow["QuoteOpr_AssemblySeq"];
            }
            if (oDataRow["QuoteOpr_OprSeq"] != DBNull.Value)
            {
                m_iOpSeq = (int)oDataRow["QuoteOpr_OprSeq"];
            }
            if (oDataRow["OpMaster_OpDesc"] != DBNull.Value)
            {
                m_sOperationDesctiption = (string)oDataRow["OpMaster_OpDesc"];
            }
            if (oDataRow["Vendor_Name"] != DBNull.Value)
            {
                m_sVendor = (string)oDataRow["Vendor_Name"];
            }
            if (oDataRow["Calculated_ProductionHours"] != DBNull.Value)
            {
                m_dProductionHours = (decimal)oDataRow["Calculated_ProductionHours"];
            }
            if (oDataRow["Calculated_LaborCost"] != DBNull.Value)
            {
                m_dLaborCosts = (decimal)oDataRow["Calculated_LaborCost"];
            }
            if (oDataRow["Calculated_BurdenCost"] != DBNull.Value)
            {
                m_dBurdenCosts = (decimal)oDataRow["Calculated_BurdenCost"];
            }
            if (oDataRow["Calculated_SubcontractCost"] != DBNull.Value)
            {
                m_dSubcontractCosts = (decimal)oDataRow["Calculated_SubcontractCost"];
            }
            if (oDataRow["Calculated_ActualSubCost"] != DBNull.Value)
            {
                m_dActualSubcontractCosts = (decimal)oDataRow["Calculated_ActualSubCost"];
            }
            if (oDataRow["Calculated_ActualQty"] != DBNull.Value)
            {
                m_dActualQty = (decimal)oDataRow["Calculated_ActualQty"];
            }
            if (oDataRow["QuoteQty1_SellingQuantity"] != DBNull.Value)
            {
                m_dSellingQty = (decimal)oDataRow["QuoteQty1_SellingQuantity"];
            }
            if (oDataRow["Calculated_TotalSetupManHours"] != DBNull.Value)
            {
                m_dTotalSetupHours = (decimal)oDataRow["Calculated_TotalSetupManHours"];
            }
            if (oDataRow["Calculated_TotalProductionManHours"] != DBNull.Value)
            {
                m_dTotalProductionHours = (decimal)oDataRow["Calculated_TotalProductionManHours"];
            }
            if (oDataRow["QuoteOpr_ProdLabRate"] != DBNull.Value)
            {
                m_dProductionLaborRate = (decimal)oDataRow["QuoteOpr_ProdLabRate"];
            }
            if (oDataRow["QuoteOpr_SetupLabRate"] != DBNull.Value)
            {
                m_dSetupLaborRate = (decimal)oDataRow["QuoteOpr_SetupLabRate"];
            }
            if (oDataRow["QuoteOpr_EstScrap"] != DBNull.Value)
            {
                m_dEstScrap = (decimal)oDataRow["QuoteOpr_EstScrap"];
            }
            //%
            if (oDataRow["QuoteOpr_EstScrapType"] != DBNull.Value)
            {
                m_sScrapType = (string)oDataRow["QuoteOpr_EstScrapType"];
            }
        }

        #endregion

        #region Properties

        public string Company
        {
            get { return m_sCompany; }
            set { m_sCompany = value; }
        }

        public int QuoteNum
        {
            get { return m_iQuoteNum; }
            set { m_iQuoteNum = value; }
        }

        public int QuoteLine
        {
            get { return m_iQuoteLine; }
            set { m_iQuoteLine = value; }
        }

        public int QuoteQtyNum
        {
            get { return m_iQuoteQtyNum; }
            set { m_iQuoteQtyNum = value; }
        }
        public string OpCode
        {
            get { return m_sOpCode; }
            set { m_sOpCode = value; }
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

        public int  AssemblySeq
        {
            get { return m_iAssemblySeq; }
            set { m_iAssemblySeq = value; }
        }

        public int OpSeq
        {
            get { return m_iOpSeq; }
            set { m_iOpSeq = value; }
        }

        public string OperationDesctiption
        {
            get { return m_sOperationDesctiption; }
            set { m_sOperationDesctiption = value; }
        }

        public string Vendor
        {
            get { return m_sVendor; }
            set { m_sVendor = value; }
        }
        public decimal ProductionHours
        {
            get { return m_dProductionHours; }
            set { m_dProductionHours = value; }
        }
        public decimal LaborCosts
        {
            get { return m_dLaborCosts; }
            set { m_dLaborCosts = value; }
        }
        public decimal BurdenCosts
        {
            get { return m_dBurdenCosts; }
            set { m_dBurdenCosts = value; }
        }
        public decimal SubcontractCosts
        {
            get { return m_dSubcontractCosts; }
            set { m_dSubcontractCosts = value; }
        }
        public decimal ActualSubcontractCosts
        {
            get { return m_dActualSubcontractCosts; }
            set { m_dActualSubcontractCosts = value; }
        }
        public decimal ActualQty
        {
            get { return m_dActualQty; }
            set { m_dActualQty = value; }
        }
        public decimal SellingQty
        {
            get { return m_dSellingQty; }
            set { m_dSellingQty = value; }
        }
        public decimal TotalSetupHours
        {
            get { return m_dTotalSetupHours; }
            set { m_dTotalSetupHours = value; }
        }
        public decimal TotalProductionHours
        {
            get { return m_dTotalProductionHours; }
            set { m_dTotalProductionHours = value; }
        }
        public decimal ProductionLaborRate
        {
            get { return m_dProductionLaborRate; }
            set { m_dProductionLaborRate = value; }
        }
        public decimal SetupLaborRate
        {
            get { return m_dSetupLaborRate; }
            set { m_dSetupLaborRate = value; }
        }
        public decimal EstScrap
        {
            get { return m_dEstScrap; }
            set { m_dEstScrap = value; }
        }
        public string ScrapType
        {
            get { return m_sScrapType; }
            set { m_sScrapType = value; }
        }
        #endregion

        #region Data Members

        private string m_sCompany;
        private int m_iQuoteNum;
        private int m_iQuoteLine;
        private int m_iQuoteQtyNum;
        private string m_sOpCode;
        private string m_sPartNum;
        private string m_sPartDescription;
        private int m_iAssemblySeq;
        private int m_iOpSeq;
        private string m_sOperationDesctiption;
        private string m_sVendor;
        private decimal m_dProductionHours;
        private decimal m_dLaborCosts;
        private decimal m_dBurdenCosts;
        private decimal m_dSubcontractCosts;
        private decimal m_dActualSubcontractCosts;
        private decimal m_dActualQty;
        private decimal m_dSellingQty;
        private decimal m_dTotalSetupHours;
        private decimal m_dTotalProductionHours;
        private decimal m_dProductionLaborRate;
        private decimal m_dSetupLaborRate;
        private decimal m_dEstScrap;
        private string m_sScrapType;

        #endregion
    }

    public class SF1411Material
    {
        #region Constructors
        public SF1411Material()
        {
        }

        public SF1411Material(DataRow oDataRow)
        {
            if (oDataRow["QuoteDtl_Company"] != DBNull.Value)
            {
                m_sCompany = (string)oDataRow["QuoteDtl_Company"];
            }
            if (oDataRow["QuoteDtl_QuoteNum"] != DBNull.Value)
            {
                m_iQuoteNum = (int)oDataRow["QuoteDtl_QuoteNum"];
            }
            if (oDataRow["QuoteDtl_QuoteLine"] != DBNull.Value)
            {
                m_iQuoteLine = (int)oDataRow["QuoteDtl_QuoteLine"];
            }
            if (oDataRow["QuoteQty1_QtyNum"] != DBNull.Value)
            {
                m_iQuoteQtyNum = (int)oDataRow["QuoteQty1_QtyNum"];
            }
            if (oDataRow["QuoteMtl_AssemblySeq"] != DBNull.Value)
            {
                m_iAssemblySeq = (int)oDataRow["QuoteMtl_AssemblySeq"];
            }
            if (oDataRow["QuoteMtl_MtlSeq"] != DBNull.Value)
            {
                m_iMaterialSeq = (int)oDataRow["QuoteMtl_MtlSeq"];
            }
            if (oDataRow["QuoteMtl_PartNum"] != DBNull.Value)
            {
                m_sPartNum = (string)oDataRow["QuoteMtl_PartNum"];
            }
            if (oDataRow["QuoteMtl_Description"] != DBNull.Value)
            {
                m_sPartDescription = (string)oDataRow["QuoteMtl_Description"];
            }
            if (oDataRow["Vendor_Name"] != DBNull.Value)
            {
                m_sVendor = (string)oDataRow["Vendor_Name"];
            }
            if (oDataRow["QuoteMtl_EstUnitCost"] != DBNull.Value)
            {
                m_dEstUnitCost = (decimal)oDataRow["QuoteMtl_EstUnitCost"];
            }
            if (oDataRow["QuoteMtl_EstMtlBurUnitCost"] != DBNull.Value)
            {
                m_dEstMtlBurdenUnitCost = (decimal)oDataRow["QuoteMtl_EstMtlBurUnitCost"];
            }
            if (oDataRow["QuoteMtl_RequiredQty"] != DBNull.Value)
            {
                m_dRequiredQty = (decimal)oDataRow["QuoteMtl_RequiredQty"];
            }
            if (oDataRow["QuoteMtl_EstScrap"] != DBNull.Value)
            {
                m_dEstScrap = (decimal)oDataRow["QuoteMtl_EstScrap"];
            }
            if (oDataRow["QuoteMtl_EstScrapType"] != DBNull.Value)
            {
                m_sScrapType = (string)oDataRow["QuoteMtl_EstScrapType"];
            }
            if (oDataRow["QuoteMtl_EstMtlUnitCost"] != DBNull.Value)
            {
                m_dEstMtlUnitCost = (decimal)oDataRow["QuoteMtl_EstMtlUnitCost"];
            }
            if (oDataRow["QuoteMtl_MinimumCost"] != DBNull.Value)
            {
                m_dMinimumCost = (decimal)oDataRow["QuoteMtl_MinimumCost"];
            }
            if (oDataRow["QuoteMtl_MtlBurRate"] != DBNull.Value)
            {
                m_dMaterialBurdenRate = (decimal) oDataRow["QuoteMtl_MtlBurRate"];
            }
            if (oDataRow["QuoteMtl_QtyPer"] != DBNull.Value)
            {
                m_dQtyPer = (decimal)oDataRow["QuoteMtl_QtyPer"];
            }
            if (oDataRow["QuoteMtl_FixedQty"] != DBNull.Value)
            {
                m_bFixedQty = (bool)oDataRow["QuoteMtl_FixedQty"];
            }
            if (oDataRow["Calculated_TotalCost"] != DBNull.Value)
            {
                m_dTotalCost = (decimal)oDataRow["Calculated_TotalCost"];
            }
            decimal dTmp;
            if (oDataRow["QuoteMtl_PBrkQty01"] != DBNull.Value)
            {
                dTmp = (decimal)oDataRow["QuoteMtl_PBrkQty01"];
                m_oPriceBreakQuantities[0] = dTmp;
            }
            if (oDataRow["QuoteMtl_PBrkCost01"] != DBNull.Value)
            {
                dTmp = (decimal)oDataRow["QuoteMtl_PBrkCost01"];
                m_oPriceBreakCosts[0] = dTmp;
            }
            if (oDataRow["QuoteMtl_PBrkQty02"] != DBNull.Value)
            {
                dTmp = (decimal)oDataRow["QuoteMtl_PBrkQty02"];
                m_oPriceBreakQuantities[1] = dTmp;
            }
            if (oDataRow["QuoteMtl_PBrkCost02"] != DBNull.Value)
            {
                dTmp = (decimal)oDataRow["QuoteMtl_PBrkCost02"];
                m_oPriceBreakCosts[1] = dTmp;
            }
            if (oDataRow["QuoteMtl_PBrkQty03"] != DBNull.Value)
            {
                dTmp = (decimal)oDataRow["QuoteMtl_PBrkQty03"];
                m_oPriceBreakQuantities[2] = dTmp;
            }
            if (oDataRow["QuoteMtl_PBrkCost03"] != DBNull.Value)
            {
                dTmp = (decimal)oDataRow["QuoteMtl_PBrkCost03"];
                m_oPriceBreakCosts[2] = dTmp;
            }
            if (oDataRow["QuoteMtl_PBrkQty04"] != DBNull.Value)
            {
                dTmp = (decimal)oDataRow["QuoteMtl_PBrkQty04"];
                m_oPriceBreakQuantities[3] = dTmp;
            }
            if (oDataRow["QuoteMtl_PBrkCost04"] != DBNull.Value)
            {
                dTmp = (decimal)oDataRow["QuoteMtl_PBrkCost04"];
                m_oPriceBreakCosts[3] = dTmp;
            }
            if (oDataRow["QuoteMtl_PBrkQty05"] != DBNull.Value)
            {
                dTmp = (decimal)oDataRow["QuoteMtl_PBrkQty05"];
                m_oPriceBreakQuantities[4] = dTmp;
            }
            if (oDataRow["QuoteMtl_PBrkCost05"] != DBNull.Value)
            {
                dTmp = (decimal)oDataRow["QuoteMtl_PBrkCost05"];
                m_oPriceBreakCosts[4] = dTmp;
            }
            if (oDataRow["QuoteMtl_PBrkQty06"] != DBNull.Value)
            {
                dTmp = (decimal)oDataRow["QuoteMtl_PBrkQty06"];
                m_oPriceBreakQuantities[5] = dTmp;
            }
            if (oDataRow["QuoteMtl_PBrkCost06"] != DBNull.Value)
            {
                dTmp = (decimal)oDataRow["QuoteMtl_PBrkCost06"];
                m_oPriceBreakCosts[5] = dTmp;
            }
            if (oDataRow["QuoteMtl_PBrkQty07"] != DBNull.Value)
            {
                dTmp = (decimal)oDataRow["QuoteMtl_PBrkQty07"];
                m_oPriceBreakQuantities[6] = dTmp;
            }
            if (oDataRow["QuoteMtl_PBrkCost07"] != DBNull.Value)
            {
                dTmp = (decimal)oDataRow["QuoteMtl_PBrkCost07"];
                m_oPriceBreakCosts[6] = dTmp;
            }
            if (oDataRow["QuoteMtl_PBrkQty08"] != DBNull.Value)
            {
                dTmp = (decimal)oDataRow["QuoteMtl_PBrkQty08"];
                m_oPriceBreakQuantities[7] = dTmp;
            }
            if (oDataRow["QuoteMtl_PBrkCost08"] != DBNull.Value)
            {
                dTmp = (decimal)oDataRow["QuoteMtl_PBrkCost08"];
                m_oPriceBreakCosts[7] = dTmp;
            }
            if (oDataRow["QuoteMtl_PBrkQty09"] != DBNull.Value)
            {
                dTmp = (decimal)oDataRow["QuoteMtl_PBrkQty09"];
                m_oPriceBreakQuantities[8] = dTmp;
            }
            if (oDataRow["QuoteMtl_PBrkCost09"] != DBNull.Value)
            {
                dTmp = (decimal)oDataRow["QuoteMtl_PBrkCost09"];
                m_oPriceBreakCosts[8] = dTmp;
            }
            if (oDataRow["QuoteMtl_PBrkQty10"] != DBNull.Value)
            {
                dTmp = (decimal)oDataRow["QuoteMtl_PBrkQty10"];
                m_oPriceBreakQuantities[9] = dTmp;
            }
            if (oDataRow["QuoteMtl_PBrkCost10"] != DBNull.Value)
            {
                dTmp = (decimal)oDataRow["QuoteMtl_PBrkCost10"];
                m_oPriceBreakCosts[9] = dTmp;
            }
        }

        #endregion

        #region Properties

        public string Company
        {
            get { return m_sCompany; }
            set { m_sCompany = value; }
        }

        public int QuoteNum
        {
            get { return m_iQuoteNum; }
            set { m_iQuoteNum = value; }
        }

        public int QuoteLine
        {
            get { return m_iQuoteLine; }
            set { m_iQuoteLine = value; }
        }

        public int QuoteQtyNum
        {
            get { return m_iQuoteQtyNum; }
            set { m_iQuoteQtyNum = value; }
        }
        public int AssemblySeq
        {
            get { return m_iAssemblySeq; }
            set { m_iAssemblySeq = value; }
        }
        public int MaterialSeq
        {
            get { return m_iMaterialSeq; }
            set { m_iMaterialSeq = value; }
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
        public string Vendor
        {
            get { return m_sVendor; }
            set { m_sVendor = value; }
        }
        public decimal EstUnitCost
        {
            get { return m_dEstUnitCost; }
            set { m_dEstUnitCost = value; }
        }
        public decimal EstMtlBurdenUnitCost
        {
            get { return m_dEstMtlBurdenUnitCost; }
            set { m_dEstMtlBurdenUnitCost = value; }
        }
        public decimal RequiredQty
        {
            get { return m_dRequiredQty; }
            set { m_dRequiredQty = value; }
        }
        public decimal EstScrap
        {
            get { return m_dEstScrap; }
            set { m_dEstScrap = value; }
        }
        public string ScrapType
        {
            get { return m_sScrapType; }
            set { m_sScrapType = value; }
        }
        public decimal EstMtlUnitCost
        {
            get { return m_dEstMtlUnitCost; }
            set { m_dEstMtlUnitCost = value; }
        }
        public decimal MinimumCost
        {
            get { return m_dMinimumCost; }
            set { m_dMinimumCost = value; }
        }
        
        public decimal MaterialBurdenRate
        {
            get { return m_dMaterialBurdenRate; }
            set { m_dMaterialBurdenRate = value; }
        }

        public decimal QtyPer
        {
            get { return m_dQtyPer; }
            set { m_dQtyPer = value; }
        }
        public bool FixedQty
        {
            get { return m_bFixedQty; }
            set { m_bFixedQty = value; }
        }
        public decimal TotalCost
        {
            get { return m_dTotalCost; }
            set { m_dTotalCost = value; }
        }
        public decimal[] PriceBreakQuantities
        {
            get { return m_oPriceBreakQuantities; }
            set { m_oPriceBreakQuantities = value; }
        }
        public decimal[] PriceBreakCosts
        {
            get { return m_oPriceBreakCosts; }
            set { m_oPriceBreakCosts = value; }
        }
        #endregion

        #region Data Members

        private string m_sCompany;
        private int m_iQuoteNum;
        private int m_iQuoteLine;
        private int m_iQuoteQtyNum;
        private int m_iAssemblySeq;
        private int m_iMaterialSeq;
        private string m_sPartNum;
        private string m_sPartDescription;
        private string m_sVendor;
        private decimal m_dEstUnitCost;
        private decimal m_dEstMtlBurdenUnitCost;
        private decimal m_dRequiredQty;
        private decimal m_dEstScrap;
        private string m_sScrapType;
        private decimal m_dEstMtlUnitCost;
        private decimal m_dMinimumCost;
        private decimal m_dMaterialBurdenRate;
        private decimal m_dQtyPer;
        private bool m_bFixedQty;
        private decimal m_dTotalCost;

        private decimal[] m_oPriceBreakQuantities = new decimal[10];
        private decimal[] m_oPriceBreakCosts = new decimal[10];
        
        #endregion
    }

    public class HSQuote
    {
        #region Constructors

        public HSQuote()
        {
        }

        public HSQuote(DataRow oDataRow)
        {
            if (oDataRow["QuoteDtl_Company"] != DBNull.Value)
            {
                m_sCompany = (string)oDataRow["QuoteDtl_Company"];
            }
            if (oDataRow["QuoteDtl_QuoteNum"] != DBNull.Value)
            {
                m_iQuoteNum = (int)oDataRow["QuoteDtl_QuoteNum"];
            }
            if (oDataRow["QuoteDtl_QuoteLine"] != DBNull.Value)
            {
                m_iQuoteLine = (int)oDataRow["QuoteDtl_QuoteLine"];
            }
            if (oDataRow["QuoteQty_QtyNum"] != DBNull.Value)
            {
                m_iQuoteQtyNum = (int)oDataRow["QuoteQty_QtyNum"];
            }
            if (oDataRow["Customer_CustID"] != DBNull.Value)
            {
                m_sCustomerId = (string)oDataRow["Customer_CustID"];
            }
            if (oDataRow["Customer_Name"] != DBNull.Value)
            {
                m_sCustomerName = (string)oDataRow["Customer_Name"];
            }
            if (oDataRow["QuoteHed_EntryDate"] != DBNull.Value)
            {
                m_dtEntryDate = (DateTime)oDataRow["QuoteHed_EntryDate"];
            }
            if (oDataRow["QuoteHed_EntryDate"] != DBNull.Value)
            {
                m_dtDateQuoted = (DateTime)oDataRow["QuoteHed_EntryDate"];
            }
            if (oDataRow["QuoteHed_QuoteComment"] != DBNull.Value)
            {
                m_sQuoteComment = (string)oDataRow["QuoteHed_QuoteComment"];
            }
            if (oDataRow["QuoteDtl_PartNum"] != DBNull.Value)
            {
                m_sPartNum = (string)oDataRow["QuoteDtl_PartNum"];
            }
            if (oDataRow["QuoteDtl_LineDesc"] != DBNull.Value)
            {
                m_sPartDescription = (string)oDataRow["QuoteDtl_LineDesc"];
            }
            if (oDataRow["QuoteQty_SellingQuantity"] != DBNull.Value)
            {
                m_dSellingQty = (decimal)oDataRow["QuoteQty_SellingQuantity"];
            }
            if (oDataRow["Calculated_BurdenCost"] != DBNull.Value)
            {
                m_dBurdenCost = (decimal)oDataRow["Calculated_BurdenCost"];
            }
            if (oDataRow["Calculated_LaborCost"] != DBNull.Value)
            {
                m_dLaborCost = (decimal)oDataRow["Calculated_LaborCost"];
            }
            if (oDataRow["Calculated_MtlCost"] != DBNull.Value)
            {
                m_dMaterialCost = (decimal)oDataRow["Calculated_MtlCost"];
            }
            if (oDataRow["Calculated_SubcontractCost"] != DBNull.Value)
            {
                m_dSubcontractCost = (decimal)oDataRow["Calculated_SubcontractCost"];
            }
            if (oDataRow["Calculated_MtlBurdenCost"] != DBNull.Value)
            {
                m_dMaterialBurdenCost = (decimal)oDataRow["Calculated_MtlBurdenCost"];
            }
            if (oDataRow["QuoteQty_MiscCostDesc"] != DBNull.Value)
            {
                m_sMiscCostDescription = (string)oDataRow["QuoteQty_MiscCostDesc"];
            }
            if (oDataRow["QuoteQty_MiscCost"] != DBNull.Value)
            {
                m_dMiscCost = (decimal)oDataRow["QuoteQty_MiscCost"];
            }
            if (oDataRow["QuoteQty_PricePerCode"] != DBNull.Value)
            {
                m_sPricePerCode = (string)oDataRow["QuoteQty_PricePerCode"];
            }
            if (oDataRow["QuoteQty_PercentType"] != DBNull.Value)
            {
                m_sPercentType = (string)oDataRow["QuoteQty_PercentType"];
            }
            if (oDataRow["QuoteQty_BurdenMarkUp"] != DBNull.Value)
            {
                m_dBurdenMarkup = (decimal)oDataRow["QuoteQty_BurdenMarkUp"];
            }
            if (oDataRow["QuoteQty_LaborMarkUp"] != DBNull.Value)
            {
                m_dLaborMarkup = (decimal)oDataRow["QuoteQty_LaborMarkUp"];
            }
            if (oDataRow["QuoteQty_MaterialMarkUp"] != DBNull.Value)
            {
                m_dMaterialMarkup = (decimal)oDataRow["QuoteQty_MaterialMarkUp"];
            }
            if (oDataRow["QuoteQty_SubcontractMarkUp"] != DBNull.Value)
            {
                m_dSubcontractMarkup = (decimal)oDataRow["QuoteQty_SubcontractMarkUp"];
            }
            if (oDataRow["QuoteQty_MtlBurMarkUp"] != DBNull.Value)
            {
                m_dMaterialBurdenMarkup = (decimal)oDataRow["QuoteQty_MtlBurMarkUp"];
            }
            if (oDataRow["QuoteQty_MiscCostMarkUp"] != DBNull.Value)
            {
                m_dMiscCostMarkup = (decimal)oDataRow["QuoteQty_MiscCostMarkUp"];
            }
            if (oDataRow["QuoteQty_CommissionPct"] != DBNull.Value)
            {
                m_dCommissionPercent = (decimal)oDataRow["QuoteQty_CommissionPct"];
            }
            if (oDataRow["QuoteQty_SellingFactor"] != DBNull.Value)
            {
                m_dSellingFactor = (decimal)oDataRow["QuoteQty_SellingFactor"];
            }
            if (oDataRow["QuoteQty_SellingFactorDirection"] != DBNull.Value)
            {
                m_sSellingDirection = (string)oDataRow["QuoteQty_SellingFactorDirection"];
            }
            if (oDataRow["QuoteQty_SalesUM"] != DBNull.Value)
            {
                m_sSalesUOM = (string)oDataRow["QuoteQty_SalesUM"];
            }
            if (oDataRow["Calculated_PriceBurdenProfit"] != DBNull.Value)
            {
                m_dPriceBurdenProfit = (decimal)oDataRow["Calculated_PriceBurdenProfit"];
            }
            if (oDataRow["Calculated_PriceLaborProfit"] != DBNull.Value)
            {
                m_dPriceLaborProfit = (decimal)oDataRow["Calculated_PriceLaborProfit"];
            }
            if (oDataRow["Calculated_PriceMtlProfit"] != DBNull.Value)
            {
                m_dPriceMaterialProfit = (decimal)oDataRow["Calculated_PriceMtlProfit"];
            }
            if (oDataRow["Calculated_PriceSubcontractProfit"] != DBNull.Value)
            {
                m_dPriceSubcontractProfit = (decimal)oDataRow["Calculated_PriceSubcontractProfit"];
            }
            if (oDataRow["Calculated_PriceMtlBurdenProfit"] != DBNull.Value)
            {
                m_dPriceMaterialBurdenProfit = (decimal)oDataRow["Calculated_PriceMtlBurdenProfit"];
            }
            if (oDataRow["Calculated_PriceMiscProfit"] != DBNull.Value)
            {
                m_dPriceMiscProfit = (decimal)oDataRow["Calculated_PriceMiscProfit"];
            }
            if (oDataRow["Calculated_TotalCost"] != DBNull.Value)
            {
                m_dTotalCost = (decimal)oDataRow["Calculated_TotalCost"];
            }
            if (oDataRow["Calculated_TotalProfit"] != DBNull.Value)
            {
                m_dTotalProfit = (decimal)oDataRow["Calculated_TotalProfit"];
            }
            if (oDataRow["Calculated_PriceTotalMarkup"] != DBNull.Value)
            {
                m_dPriceTotalMarkup = (decimal)oDataRow["Calculated_PriceTotalMarkup"];
            }
            if (oDataRow["Calculated_TotalCommission"] != DBNull.Value)
            {
                m_dTotalCommission = (decimal)oDataRow["Calculated_TotalCommission"];
            }
            if (oDataRow["Calculated_PriceTotalCommissionMarkup"] != DBNull.Value)
            {
                m_dPriceTotalCommissionMarkup = (decimal)oDataRow["Calculated_PriceTotalCommissionMarkup"];
            }
            if (oDataRow["Calculated_UnitCost"] != DBNull.Value)
            {
                m_dUnitCost = (decimal)oDataRow["Calculated_UnitCost"];
            }
            if (oDataRow["Calculated_UnitPrice"] != DBNull.Value)
            {
                m_dUnitPrice = (decimal)oDataRow["Calculated_UnitPrice"];
            }
            if (oDataRow["Calculated_UnitPriceWithCommission"] != DBNull.Value)
            {
                m_dUnitPriceWithCommission = (decimal)oDataRow["Calculated_UnitPriceWithCommission"];
            }

            // we need to calculate this material price
            //if (string.Compare(m_sPercentType, "%", true) == 0)
            //{
            //    m_dPriceMaterialProfit = (m_dMaterialCost + (m_dMaterialCost * m_dMaterialMarkup / 100.0M));
            //}
            //else
            //{
            //    m_dPriceMaterialProfit = m_dMaterialCost / ((100.0M - m_dMaterialMarkup) / 100.0M);
            //}
        }
        #endregion

        static public List<HSQuote> Initialize(Session oSession, int iQuoteNum)
        {
            List<HSQuote> oQuotes = new List<HSQuote>();

            //****Set a parameter Value***** 
            Ice.Proxy.BO.DynamicQueryImpl oDynamicQuery = WCFServiceSupport.CreateImpl<Ice.Proxy.BO.DynamicQueryImpl>(oSession, Ice.Proxy.BO.DynamicQueryImpl.UriPath);
            QueryExecutionDataSet oQueryExecutionDataSet = oDynamicQuery.GetQueryExecutionParametersByID(BAQConstants.QUERY_QUOTE_COSTS);
            foreach (DataRow oParameter in oQueryExecutionDataSet.ExecutionParameter)
            {
                if (string.Compare(oParameter["ParameterID"].ToString(), "QuoteNum") == 0)
                {
                    oParameter["ParameterValue"] = iQuoteNum;
                }
            }
            oQueryExecutionDataSet.AcceptChanges();
            DataSet oDataSet = oDynamicQuery.ExecuteByID(BAQConstants.QUERY_QUOTE_COSTS, oQueryExecutionDataSet);
            foreach (DataRow oRow in oDataSet.Tables[0].Rows)
            {
                oQuotes.Add(new HSQuote(oRow));
            }

            // now get the operation details for this quote
            oQueryExecutionDataSet = oDynamicQuery.GetQueryExecutionParametersByID(BAQConstants.QUERY_QUOTE_OPERATIONS);
            foreach (DataRow oParameter in oQueryExecutionDataSet.ExecutionParameter)
            {
                if (string.Compare(oParameter["ParameterID"].ToString(), "QuoteNum") == 0)
                {
                    oParameter["ParameterValue"] = iQuoteNum;
                }
            }
            oQueryExecutionDataSet.AcceptChanges();
            oDataSet = oDynamicQuery.ExecuteByID(BAQConstants.QUERY_QUOTE_OPERATIONS, oQueryExecutionDataSet);
            List<HSQuoteOperations> oAllQuoteOperations = new List<HSQuoteOperations>();
            foreach (DataRow oRow in oDataSet.Tables[0].Rows)
            {
                oAllQuoteOperations.Add(new HSQuoteOperations(oRow));
            }

            // get the quote operation details
            oQueryExecutionDataSet = oDynamicQuery.GetQueryExecutionParametersByID(BAQConstants.QUERY_QUOTE_OPERATION_DETAILS);
            foreach (DataRow oParameter in oQueryExecutionDataSet.ExecutionParameter)
            {
                if (string.Compare(oParameter["ParameterID"].ToString(), "QuoteNum") == 0)
                {
                    oParameter["ParameterValue"] = iQuoteNum;
                }
            }
            oQueryExecutionDataSet.AcceptChanges();
            oDataSet = oDynamicQuery.ExecuteByID(BAQConstants.QUERY_QUOTE_OPERATION_DETAILS, oQueryExecutionDataSet);
            List<SF1411Operation> oAllSF1411Operations = new List<SF1411Operation>();
            foreach (DataRow oRow in oDataSet.Tables[0].Rows)
            {
                oAllSF1411Operations.Add(new SF1411Operation(oRow));
            }

            // get the quote material details
            oQueryExecutionDataSet = oDynamicQuery.GetQueryExecutionParametersByID(BAQConstants.QUERY_QUOTE_MATERIAL_DETAILS);
            foreach (DataRow oParameter in oQueryExecutionDataSet.ExecutionParameter)
            {
                if (string.Compare(oParameter["ParameterID"].ToString(), "QuoteNum") == 0)
                {
                    oParameter["ParameterValue"] = iQuoteNum;
                }
            }
            oQueryExecutionDataSet.AcceptChanges();
            oDataSet = oDynamicQuery.ExecuteByID(BAQConstants.QUERY_QUOTE_MATERIAL_DETAILS, oQueryExecutionDataSet);
            List<SF1411Material> oAllsf1411Materials = new List<SF1411Material>();
            foreach (DataRow oRow in oDataSet.Tables[0].Rows)
            {
                oAllsf1411Materials.Add(new SF1411Material(oRow));
            }

            // now we need to associate the correct operations with each quote line
            foreach (HSQuote oQuote in oQuotes)
            {
                oQuote.m_oOperations.Clear();
                oQuote.m_oSF1411Operations.Clear();
                oQuote.m_oSF1411Materials.Clear();

                // pull out all operations tied to this quote, quote line, and qty number
                List<HSQuoteOperations> oAssociatedOperations = oAllQuoteOperations.Where(x => (x.QuoteNum == oQuote.QuoteNum) && (x.QuoteLine == oQuote.QuoteLine) && (x.QuoteQtyNum == oQuote.QuoteQtyNum)).ToList();
                oQuote.m_oOperations.AddRange(oAssociatedOperations);

                List<SF1411Operation> oAssociatedOperationDetails = oAllSF1411Operations.Where(x => (x.QuoteNum == oQuote.QuoteNum) && (x.QuoteLine == oQuote.QuoteLine) && (x.QuoteQtyNum == oQuote.QuoteQtyNum)).ToList();
                oQuote.m_oSF1411Operations.AddRange(oAssociatedOperationDetails);

                List<SF1411Material> oAssociatedMaterialDetails = oAllsf1411Materials.Where(x => (x.QuoteNum == oQuote.QuoteNum) && (x.QuoteLine == oQuote.QuoteLine) && (x.QuoteQtyNum == oQuote.QuoteQtyNum)).ToList();
                oQuote.m_oSF1411Materials.AddRange(oAssociatedMaterialDetails);
            }

            return oQuotes;
        }

        #region Generate Quote Forms
        public static void CreateAndSendReport(string sTmpFileDirectory, HSUser oRequestingUser, List<HSQuote> oQuotes)
        {
            // get the file name
            DateTime dtToday = DateTime.Now;
            string sNameOfMonth = dtToday.ToString("MMMM");

            // these are all for the same quote so we just get the first quote in the list
            HSQuote oFirstQuote = oQuotes[0];
            string sDestinationFileName = sTmpFileDirectory + "\\QuoteSF1411-" + oFirstQuote.QuoteNum.ToString() + ".xlsx";

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

            SetStyles(oSLDocument);

            bool bFirstWorksheet = true;
            foreach (HSQuote oQuote in oQuotes)
            {
                string sWorksheetName = oQuote.QuoteNum.ToString() + "-" + oQuote.QuoteLine.ToString() + "-" + oQuote.m_iQuoteQtyNum.ToString();
                if (bFirstWorksheet == true)
                {
                    oSLDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, sWorksheetName);
                    bFirstWorksheet = false;
                }
                else
                {
                    oSLDocument.AddWorksheet(sWorksheetName);
                }

                //set up column headers
                int iAColumn = 1;
                int iBColumn = 2;
                int iCColumn = 3;
                int iDColumn = 4;
                int iEColumn = 5;
                int iFColumn = 6;
                int iGColumn = 7;
                //int iHColumn = 8;
                //int iIColumn = 9;

                int iNumOfRows = 1;

                oSLDocument.SetCellValue(iNumOfRows, iAColumn, "QUOTE");
                oSLDocument.SetCellStyle(iNumOfRows, iAColumn, g_oBoldCenter);
                oSLDocument.SetCellValue(iNumOfRows, iBColumn, "LINE");
                oSLDocument.SetCellStyle(iNumOfRows, iBColumn, g_oBoldCenter);
                oSLDocument.SetCellValue(iNumOfRows, iCColumn, "QTY");
                oSLDocument.SetCellStyle(iNumOfRows, iCColumn, g_oBoldCenter);
                iNumOfRows++;
                oSLDocument.SetCellValue(iNumOfRows, iAColumn, oQuote.QuoteNum);
                oSLDocument.SetCellStyle(iNumOfRows, iAColumn, g_oCenter);
                oSLDocument.SetCellValue(iNumOfRows, iBColumn, oQuote.QuoteLine);
                oSLDocument.SetCellStyle(iNumOfRows, iBColumn, g_oCenter);
                oSLDocument.SetCellValue(iNumOfRows, iCColumn, oQuote.SellingQty);
                oSLDocument.SetCellStyle(iNumOfRows, iCColumn, g_oCenter);

                iNumOfRows++;
                // blank row
                iNumOfRows++;
                // blank row

                oSLDocument.SetCellValue(iNumOfRows, iAColumn, "SPINCRAFT-WI");
                oSLDocument.SetCellStyle(iNumOfRows, iAColumn, g_oBold);
                oSLDocument.SetCellValue(iNumOfRows, iDColumn, oQuote.CustomerName);

                iNumOfRows++;
                oSLDocument.SetCellValue(iNumOfRows, iAColumn, "2455 COMMERCE DRIVE");
                oSLDocument.SetCellValue(iNumOfRows, iDColumn, oQuote.PartDescription);
                oSLDocument.SetCellStyle(iNumOfRows, iDColumn, g_oBoldRed);

                iNumOfRows++;
                oSLDocument.SetCellValue(iNumOfRows, iAColumn, "NEW BERLIN, WI 53151");
                // get the date the quote was entered on
                oSLDocument.SetCellValue(iNumOfRows, iBColumn, oQuote.EntryDate.ToShortDateString());
                oSLDocument.SetCellStyle(iNumOfRows, iBColumn, g_oBoldCenter);
                oSLDocument.SetCellValue(iNumOfRows, iDColumn, "P/N " + oQuote.PartNum);

                iNumOfRows++;
                oSLDocument.SetCellValue(iNumOfRows, iAColumn, "SF 1411 REFERENCE SHEET");
                oSLDocument.SetCellStyle(iNumOfRows, iAColumn, g_oBold);

                iNumOfRows++;
                // blank row

                iNumOfRows++;
                oSLDocument.SetCellValue(iNumOfRows, iBColumn, "REF");
                oSLDocument.SetCellStyle(iNumOfRows, iBColumn, g_oUnderlineCenetered);
                oSLDocument.SetCellValue(iNumOfRows, iCColumn, "TOTAL COST");
                oSLDocument.SetCellStyle(iNumOfRows, iCColumn, g_oUnderlineCenetered);
                oSLDocument.SetCellValue(iNumOfRows, iDColumn, "UNIT COST");
                oSLDocument.SetCellStyle(iNumOfRows, iDColumn, g_oUnderlineCenetered);


                iNumOfRows++;
                oSLDocument.SetCellValue(iNumOfRows, iAColumn, "COST ELEMENTS:");
                oSLDocument.SetCellStyle(iNumOfRows, iAColumn, g_oBoldCenter);

                iNumOfRows++;
                oSLDocument.SetCellValue(iNumOfRows, iAColumn, "RAW MATERIAL");
                oSLDocument.SetCellValue(iNumOfRows, iBColumn, "1A.");
                oSLDocument.SetCellStyle(iNumOfRows, iBColumn, g_oCenter);
                oSLDocument.SetCellValue(iNumOfRows, iCColumn, oQuote.MaterialCost + oQuote.MaterialBurdenCost);
                oSLDocument.SetCellStyle(iNumOfRows, iCColumn, g_oCurrencyStyleCentered);
                if (oQuote.SellingQty != 0)
                {
                    oSLDocument.SetCellValue(iNumOfRows, iDColumn, (oQuote.MaterialCost + oQuote.MaterialBurdenCost) / oQuote.SellingQty);
                }
                else
                {
                    oSLDocument.SetCellValue(iNumOfRows, iDColumn, 0);
                }
                oSLDocument.SetCellStyle(iNumOfRows, iDColumn, g_oCurrencyStyleCentered);

                iNumOfRows++;
                oSLDocument.SetCellValue(iNumOfRows, iAColumn, "SUB-CONTRACTED ITEMS");
                oSLDocument.SetCellValue(iNumOfRows, iBColumn, "2A.");
                oSLDocument.SetCellStyle(iNumOfRows, iBColumn, g_oCenter);
                oSLDocument.SetCellValue(iNumOfRows, iCColumn, oQuote.SubcontractCost);
                oSLDocument.SetCellStyle(iNumOfRows, iCColumn, g_oCurrencyStyleCentered);
                if (oQuote.SellingQty != 0)
                {
                    oSLDocument.SetCellValue(iNumOfRows, iDColumn, oQuote.SubcontractCost / oQuote.SellingQty);
                }
                else
                {
                    oSLDocument.SetCellValue(iNumOfRows, iDColumn, 0);
                }
                oSLDocument.SetCellStyle(iNumOfRows, iDColumn, g_oCurrencyStyleCentered);

                iNumOfRows++;
                oSLDocument.SetCellValue(iNumOfRows, iAColumn, "DIRECT LABOR");
                oSLDocument.SetCellValue(iNumOfRows, iBColumn, "3A.");
                oSLDocument.SetCellStyle(iNumOfRows, iBColumn, g_oCenter);
                oSLDocument.SetCellValue(iNumOfRows, iCColumn, oQuote.LaborCost);
                oSLDocument.SetCellStyle(iNumOfRows, iCColumn, g_oCurrencyStyleCentered);
                if (oQuote.SellingQty != 0)
                {
                    oSLDocument.SetCellValue(iNumOfRows, iDColumn, oQuote.LaborCost / oQuote.SellingQty);
                }
                else
                {
                    oSLDocument.SetCellValue(iNumOfRows, iDColumn, 0);
                }
                oSLDocument.SetCellStyle(iNumOfRows, iDColumn, g_oCurrencyStyleCentered);

                iNumOfRows++;
                oSLDocument.SetCellValue(iNumOfRows, iAColumn, "MFG. OVERHEAD");
                oSLDocument.SetCellValue(iNumOfRows, iBColumn, "4A.");
                oSLDocument.SetCellStyle(iNumOfRows, iBColumn, g_oCenter);
                oSLDocument.SetCellValue(iNumOfRows, iCColumn, oQuote.BurdenCost);
                oSLDocument.SetCellStyle(iNumOfRows, iCColumn, g_oCurrencyStyleCentered);
                if (oQuote.SellingQty != 0)
                {
                    oSLDocument.SetCellValue(iNumOfRows, iDColumn, oQuote.BurdenCost / oQuote.SellingQty);
                }
                else
                {
                    oSLDocument.SetCellValue(iNumOfRows, iDColumn, 0);
                }
                oSLDocument.SetCellStyle(iNumOfRows, iDColumn, g_oCurrencyStyleCentered);

                iNumOfRows++;
                oSLDocument.SetCellValue(iNumOfRows, iAColumn, "SUBTOTAL");
                oSLDocument.SetCellStyle(iNumOfRows, iAColumn, g_oBoldCenter);
                // SHOULD THIS INCLUDE MISC COSTS????????????
                decimal dTotalCosts = oQuote.MaterialCost + oQuote.MaterialBurdenCost + oQuote.SubcontractCost + oQuote.LaborCost + oQuote.BurdenCost;
                decimal dTotalPrice = oQuote.PriceMaterialProfit + oQuote.PriceMaterialBurdenProfit + oQuote.PriceSubcontractProfit + oQuote.PriceLaborProfit + oQuote.PriceBurdenProfit;
                oSLDocument.SetCellValue(iNumOfRows, iCColumn, dTotalCosts);
                oSLDocument.SetCellStyle(iNumOfRows, iCColumn, g_oCurrencyStyleBoldCentered);
                if (oQuote.SellingQty != 0)
                {
                    oSLDocument.SetCellValue(iNumOfRows, iDColumn, dTotalCosts / oQuote.SellingQty);
                }
                else
                {
                    oSLDocument.SetCellValue(iNumOfRows, iDColumn, 0);
                }
                oSLDocument.SetCellStyle(iNumOfRows, iDColumn, g_oCurrencyStyleBoldCentered);

                iNumOfRows++;
                // blank row

                iNumOfRows++;
                oSLDocument.SetCellValue(iNumOfRows, iAColumn, "SALES, G&A EXPENSES");
                oSLDocument.SetCellValue(iNumOfRows, iBColumn, "5A.");
                oSLDocument.SetCellStyle(iNumOfRows, iBColumn, g_oCenter);
                // SHOULD THIS BE COMMISSION + MISC???????????
                //decimal dGAExpenses = oQuote.TotalCommission + oQuote.MiscCost;
                decimal dGAExpenses = oQuote.TotalCommission + oQuote.PriceMiscProfit;
                oSLDocument.SetCellValue(iNumOfRows, iCColumn, dGAExpenses);
                oSLDocument.SetCellStyle(iNumOfRows, iCColumn, g_oCurrencyStyleCentered);
                if (oQuote.SellingQty != 0)
                {
                    oSLDocument.SetCellValue(iNumOfRows, iDColumn, dGAExpenses / oQuote.SellingQty);
                }
                else
                {
                    oSLDocument.SetCellValue(iNumOfRows, iDColumn, 0);
                }
                oSLDocument.SetCellStyle(iNumOfRows, iDColumn, g_oCurrencyStyleCentered);

                iNumOfRows++;
                oSLDocument.SetCellValue(iNumOfRows, iAColumn, "SUBTOTAL");
                oSLDocument.SetCellStyle(iNumOfRows, iAColumn, g_oBoldCenter);
                oSLDocument.SetCellValue(iNumOfRows, iCColumn, oQuote.PriceTotalCommissionMarkup - oQuote.TotalProfit);
                oSLDocument.SetCellStyle(iNumOfRows, iCColumn, g_oCurrencyStyleBoldCentered);
                if (oQuote.SellingQty != 0)
                {
                    oSLDocument.SetCellValue(iNumOfRows, iDColumn, (oQuote.PriceTotalCommissionMarkup - oQuote.TotalProfit) / oQuote.SellingQty);
                }
                else
                {
                    oSLDocument.SetCellValue(iNumOfRows, iDColumn, 0);
                }
                oSLDocument.SetCellStyle(iNumOfRows, iDColumn, g_oCurrencyStyleBoldCentered);

                iNumOfRows++;
                oSLDocument.SetCellValue(iNumOfRows, iAColumn, "PROFIT OR FEE");
                oSLDocument.SetCellValue(iNumOfRows, iBColumn, "6A.");
                oSLDocument.SetCellStyle(iNumOfRows, iBColumn, g_oCenter);
                oSLDocument.SetCellValue(iNumOfRows, iCColumn, oQuote.TotalProfit);
                oSLDocument.SetCellStyle(iNumOfRows, iCColumn, g_oCurrencyStyleCentered);
                if (oQuote.SellingQty != 0)
                {
                    oSLDocument.SetCellValue(iNumOfRows, iDColumn, oQuote.TotalProfit / oQuote.SellingQty);
                }
                else
                {
                    oSLDocument.SetCellValue(iNumOfRows, iDColumn, 0);
                }
                oSLDocument.SetCellStyle(iNumOfRows, iDColumn, g_oCurrencyStyleCentered);

                iNumOfRows++;
                oSLDocument.SetCellValue(iNumOfRows, iAColumn, "TOTAL PRICE");
                oSLDocument.SetCellStyle(iNumOfRows, iAColumn, g_oBoldCenter);
                oSLDocument.SetCellValue(iNumOfRows, iCColumn, oQuote.PriceTotalCommissionMarkup);
                oSLDocument.SetCellStyle(iNumOfRows, iCColumn, g_oBoxStyleCentered);
                oSLDocument.SetCellValue(iNumOfRows, iDColumn, oQuote.UnitPriceWithCommission);
                oSLDocument.SetCellStyle(iNumOfRows, iDColumn, g_oBoxStyleCentered);

                iNumOfRows++;
                // blank row

                iNumOfRows++;
                // blank row

                iNumOfRows++;
                oSLDocument.SetCellValue(iNumOfRows, iAColumn, "REFERENCES:");
                oSLDocument.SetCellStyle(iNumOfRows, iAColumn, g_oBold);

                iNumOfRows++;
                oSLDocument.SetCellValue(iNumOfRows, iAColumn, "1A.  RAW MATERIAL");
                oSLDocument.SetCellStyle(iNumOfRows, iAColumn, g_oBoldUnderline);
                oSLDocument.SetCellValue(iNumOfRows, iCColumn, oQuote.PriceMaterialProfit + oQuote.PriceMaterialBurdenProfit);
                oSLDocument.SetCellStyle(iNumOfRows, iCColumn, g_oBoxStyleCentered);
                //
                // we will list all materials that have a non-zero cost
                //
                // PART NUM, Description, Base Qty, Scrap, Unit Cost, Extended, Vendor, Current Purchase Cost, Reference - PONum
                List<SF1411Material> oAllMaterialsWithCost = oQuote.m_oSF1411Materials.Where(x => x.TotalCost != 0).ToList();
                decimal dTotalPurchaseCostOfAllParts = 0M;
                decimal dTotalExtendedPriceOfAllParts = 0M;
                if (oAllMaterialsWithCost.Count != 0)
                {
                    // first create the header
                    iNumOfRows++;
                    oSLDocument.SetCellValue(iNumOfRows, iAColumn, "Part Number");
                    oSLDocument.SetCellStyle(iNumOfRows, iAColumn, g_oBold);
                    oSLDocument.SetCellValue(iNumOfRows, iBColumn, "Description");
                    oSLDocument.SetCellStyle(iNumOfRows, iBColumn, g_oBold);
                    oSLDocument.SetCellValue(iNumOfRows, iCColumn, "Base Qty");
                    oSLDocument.SetCellStyle(iNumOfRows, iCColumn, g_oBold);
                    oSLDocument.SetCellValue(iNumOfRows, iDColumn, "Scrap");
                    oSLDocument.SetCellStyle(iNumOfRows, iDColumn, g_oBold);
                    oSLDocument.SetCellValue(iNumOfRows, iEColumn, "Vendor");
                    oSLDocument.SetCellStyle(iNumOfRows, iEColumn, g_oBold);
                    oSLDocument.SetCellValue(iNumOfRows, iFColumn, "Extended Cost");
                    oSLDocument.SetCellStyle(iNumOfRows, iFColumn, g_oBold);
                    oSLDocument.SetCellValue(iNumOfRows, iGColumn, "Reference");
                    oSLDocument.SetCellStyle(iNumOfRows, iGColumn, g_oBold);
                    foreach (SF1411Material oMaterialWithCost in oAllMaterialsWithCost)
                    {
                        iNumOfRows++;
                        // partnum
                        oSLDocument.SetCellValue(iNumOfRows, iAColumn, oMaterialWithCost.PartNum);
                        // description
                        oSLDocument.SetCellValue(iNumOfRows, iBColumn, oMaterialWithCost.PartDescription);

                        // check for fixed qty
                        decimal dBaseQty = 0M;
                        if (oMaterialWithCost.FixedQty == true)
                        {
                            dBaseQty = oMaterialWithCost.RequiredQty;
                        }
                        else
                        {
                            dBaseQty = oMaterialWithCost.RequiredQty * oQuote.SellingQty;
                        }
                        oSLDocument.SetCellValue(iNumOfRows, iCColumn, dBaseQty);
                        oSLDocument.SetCellStyle(iNumOfRows, iCColumn, g_oDecimalBoxStyleBoldCentered);

                        // variable that holds the total material quantity that will be required for this job -- affected by scrap
                        decimal dMaterialTotalQuantity = dBaseQty;

                        // consider scrap
                        if (oMaterialWithCost.ScrapType == "%")
                        {
                            oSLDocument.SetCellStyle(iNumOfRows, iDColumn, g_oPercentStyleBoldCentered);
                            oSLDocument.SetCellValue(iNumOfRows, iDColumn, oMaterialWithCost.EstScrap / 100.0M);
                            // if we have an estimated scrap percentage then we need to modify the total material required
                            if (oMaterialWithCost.EstScrap != 0)
                            {
                                dMaterialTotalQuantity += dMaterialTotalQuantity * (oMaterialWithCost.EstScrap / 100.0M);
                            }
                        }
                        else
                        {
                            oSLDocument.SetCellStyle(iNumOfRows, iDColumn, g_oDecimalBoxStyleBoldCentered);
                            oSLDocument.SetCellValue(iNumOfRows, iDColumn, oMaterialWithCost.EstScrap);
                            // if we have an estimated scrap quantity then we need to modify the total material required
                            if (oMaterialWithCost.EstScrap != 0)
                            {
                                dMaterialTotalQuantity += oMaterialWithCost.EstScrap;
                            }
                        }

                        // now see if we need to consider price breaks for all materials
                        // check to see if a price break is required -- there are exactly 10 price breaks present
                        for (int iCurrentPosition = 0; iCurrentPosition < 10; iCurrentPosition++)
                        {
                            // if the quantity is zero then it is not in effect
                            if (oMaterialWithCost.PriceBreakQuantities[iCurrentPosition] != 0)
                            {
                                if (dMaterialTotalQuantity >= oMaterialWithCost.PriceBreakQuantities[iCurrentPosition])
                                {
                                    oMaterialWithCost.EstMtlUnitCost = oMaterialWithCost.PriceBreakCosts[iCurrentPosition];
                                    oMaterialWithCost.TotalCost = oMaterialWithCost.EstMtlUnitCost * dMaterialTotalQuantity;
                                }
                            }
                        }

                        decimal dTotalCost = oMaterialWithCost.EstMtlUnitCost * dMaterialTotalQuantity;
                        // we need to check if minimum price is in effect
                        if (dTotalCost < oMaterialWithCost.MinimumCost)
                        {
                            dTotalCost = oMaterialWithCost.MinimumCost;
                        }

                        decimal dTotalExtendedPrice = 0M;
                        if (oQuote.PercentType == "%")
                        {
                            dTotalExtendedPrice = (dTotalCost + (dTotalCost * oQuote.MaterialMarkup / 100.0M));
                        }
                        else
                        {
                            dTotalExtendedPrice = dTotalCost / ((100.0M - oQuote.MaterialMarkup) / 100.0M);
                        }
                        dTotalExtendedPriceOfAllParts += dTotalExtendedPrice;
                        // vendor
                        oSLDocument.SetCellValue(iNumOfRows, iEColumn, oMaterialWithCost.Vendor);
                        // total cost
                        oSLDocument.SetCellValue(iNumOfRows, iFColumn, dTotalCost);
                        oSLDocument.SetCellStyle(iNumOfRows, iFColumn, g_oCurrencyStyleCentered);
                        dTotalPurchaseCostOfAllParts += dTotalCost;
                        // reference
                        oSLDocument.SetCellValue(iNumOfRows, iGColumn, "PO# YYYY");
                    }

                    // add a row for the total parts
                    iNumOfRows++;
                    oSLDocument.SetCellValue(iNumOfRows, iAColumn, "Total Raw Materials");
                    oSLDocument.SetCellStyle(iNumOfRows, iAColumn, g_oBoxStyleCentered);
                    // total extended price
                    oSLDocument.SetCellValue(iNumOfRows, iFColumn, oQuote.MaterialBurdenCost + oQuote.MaterialCost);
                    oSLDocument.SetCellStyle(iNumOfRows, iFColumn, g_oBoxStyleCentered);
                }

                iNumOfRows++;
                // blank row

                iNumOfRows++;
                oSLDocument.SetCellValue(iNumOfRows, iAColumn, "2A.  SUB-CONTRACTED ITEMS");
                oSLDocument.SetCellStyle(iNumOfRows, iAColumn, g_oBoldUnderline);
                oSLDocument.SetCellValue(iNumOfRows, iCColumn, oQuote.PriceSubcontractProfit);
                oSLDocument.SetCellStyle(iNumOfRows, iCColumn, g_oCurrencyStyleCentered);
                //
                // we will list all subcontracted items that have a non-zero cost
                //
                // Part Num, Descritpion, Base Qty, Scrap, Unit Price, Extended, Vendor, Cost
                List<SF1411Operation> oAllSubcontractsWithCost = oQuote.m_oSF1411Operations.Where(x => x.SubcontractCosts != 0).ToList();
                if (oAllSubcontractsWithCost.Count != 0)
                {
                    // first create the header
                    iNumOfRows++;
                    oSLDocument.SetCellValue(iNumOfRows, iAColumn, "Part Number");
                    oSLDocument.SetCellStyle(iNumOfRows, iAColumn, g_oBold);
                    oSLDocument.SetCellValue(iNumOfRows, iBColumn, "Description");
                    oSLDocument.SetCellStyle(iNumOfRows, iBColumn, g_oBold);
                    oSLDocument.SetCellValue(iNumOfRows, iCColumn, "Base Qty");
                    oSLDocument.SetCellStyle(iNumOfRows, iCColumn, g_oBold);
                    oSLDocument.SetCellValue(iNumOfRows, iDColumn, "Scrap");
                    oSLDocument.SetCellStyle(iNumOfRows, iDColumn, g_oBold);
                    oSLDocument.SetCellValue(iNumOfRows, iEColumn, "Vendor");
                    oSLDocument.SetCellStyle(iNumOfRows, iEColumn, g_oBold);
                    oSLDocument.SetCellValue(iNumOfRows, iFColumn, "Extended Cost");
                    oSLDocument.SetCellStyle(iNumOfRows, iFColumn, g_oBold);
                    foreach (SF1411Operation oSubcontractWithCost in oAllSubcontractsWithCost)
                    {
                        iNumOfRows++;
                        oSLDocument.SetCellValue(iNumOfRows, iAColumn, oSubcontractWithCost.PartNum);

                        oSLDocument.SetCellValue(iNumOfRows, iBColumn, oSubcontractWithCost.PartDescription);

                        // need to use selling qty here not the actual qty (scrap * selling qty)
                        oSLDocument.SetCellValue(iNumOfRows, iCColumn, oSubcontractWithCost.SellingQty);
                        oSLDocument.SetCellStyle(iNumOfRows, iCColumn, g_oDecimalBoxStyleBoldCentered);

                        if (string.Compare(oSubcontractWithCost.ScrapType, "%", true) == 0)
                        {
                            oSLDocument.SetCellStyle(iNumOfRows, iDColumn, g_oPercentStyleBoldCentered);
                            oSLDocument.SetCellValue(iNumOfRows, iDColumn, oSubcontractWithCost.EstScrap / 100.0M);
                        }
                        else
                        {
                            oSLDocument.SetCellStyle(iNumOfRows, iDColumn, g_oDecimalBoxStyleBoldCentered);
                            oSLDocument.SetCellValue(iNumOfRows, iDColumn, oSubcontractWithCost.EstScrap);
                        }

                        // include any scrap factor
                        decimal dTotalSubcontractCostPerUnit = 0M;
                        if (oQuote.PercentType == "P")
                        {
                            // profit calculation
                            dTotalSubcontractCostPerUnit = ((oSubcontractWithCost.ActualSubcontractCosts * oSubcontractWithCost.ActualQty) + (oSubcontractWithCost.ActualSubcontractCosts * oSubcontractWithCost.ActualQty * oQuote.SubcontractMarkup / 100.0M));
                        }
                        else
                        {
                            // markup calculation
                            dTotalSubcontractCostPerUnit = (oSubcontractWithCost.ActualSubcontractCosts * oSubcontractWithCost.ActualQty) / ((100.0M - oQuote.SubcontractMarkup) / 100.0M);
                        }

                        oSLDocument.SetCellValue(iNumOfRows, iEColumn, oSubcontractWithCost.Vendor);

                        oSLDocument.SetCellValue(iNumOfRows, iFColumn, oSubcontractWithCost.ActualSubcontractCosts * oSubcontractWithCost.ActualQty);
                        oSLDocument.SetCellStyle(iNumOfRows, iFColumn, g_oCurrencyStyleCentered);
                    }
                    // add a row for the total subcontract costs
                    iNumOfRows++;
                    oSLDocument.SetCellValue(iNumOfRows, iAColumn, "Total Subcontracts");
                    oSLDocument.SetCellStyle(iNumOfRows, iAColumn, g_oBoxStyleCentered);
                    // total extended cost
                    oSLDocument.SetCellValue(iNumOfRows, iFColumn, oQuote.SubcontractCost);
                    oSLDocument.SetCellStyle(iNumOfRows, iFColumn, g_oBoxStyleCentered);
                }

                iNumOfRows++;
                // blank row

                iNumOfRows++;
                oSLDocument.SetCellValue(iNumOfRows, iAColumn, "3A.  DIRECT LABOR");
                oSLDocument.SetCellStyle(iNumOfRows, iAColumn, g_oBoldUnderline);
                oSLDocument.SetCellValue(iNumOfRows, iFColumn, "4A.  MFG. OVERHEAD");
                oSLDocument.SetCellStyle(iNumOfRows, iFColumn, g_oBoldUnderline);

                iNumOfRows++;
                oSLDocument.SetCellValue(iNumOfRows, iAColumn, "COST CENTER");
                oSLDocument.SetCellStyle(iNumOfRows, iAColumn, g_oUnderline);
                oSLDocument.SetCellValue(iNumOfRows, iBColumn, "HOURS");
                oSLDocument.SetCellStyle(iNumOfRows, iBColumn, g_oUnderlineCenetered);
                oSLDocument.SetCellValue(iNumOfRows, iCColumn, "LABOR RATE");
                oSLDocument.SetCellStyle(iNumOfRows, iCColumn, g_oUnderline);
                oSLDocument.SetCellValue(iNumOfRows, iDColumn, "LABOR COST");
                oSLDocument.SetCellStyle(iNumOfRows, iDColumn, g_oUnderline);

                oSLDocument.SetCellValue(iNumOfRows, iFColumn, "OVERHEAD RATE");
                oSLDocument.SetCellStyle(iNumOfRows, iFColumn, g_oUnderline);
                oSLDocument.SetCellValue(iNumOfRows, iGColumn, "OVERHEAD COST");
                oSLDocument.SetCellStyle(iNumOfRows, iGColumn, g_oUnderline);

                decimal dTotalHours = 0;
                decimal dTotalLabor = 0;
                decimal dTotalBurden = 0;
                decimal dTotalLaborPrice = 0;
                decimal dTotalBurdenPrice = 0;
                foreach (HSQuoteOperations oOperation in oQuote.AllOperations)
                {
                    // we only include operations that are not subcontracts
                    if (oOperation.SubcontractCosts == 0)
                    {
                        decimal dTotalOpHours = oOperation.TotalProductionHours + oOperation.TotalSetupHours;

                        // list all operation names and details
                        iNumOfRows++;
                        oSLDocument.SetCellValue(iNumOfRows, iAColumn, oOperation.OpCode);
                        oSLDocument.SetCellValue(iNumOfRows, iBColumn, oOperation.TotalProductionHours + oOperation.TotalSetupHours);
                        oSLDocument.SetCellStyle(iNumOfRows, iBColumn, g_oCenter);
                        if (dTotalOpHours != 0)
                        {
                            oSLDocument.SetCellValue(iNumOfRows, iCColumn, oOperation.LaborCosts / (oOperation.TotalProductionHours + oOperation.TotalSetupHours));
                        }
                        else
                        {
                            oSLDocument.SetCellValue(iNumOfRows, iCColumn, 0);
                        }
                        oSLDocument.SetCellStyle(iNumOfRows, iCColumn, g_oCurrencyStyleCentered);
                        oSLDocument.SetCellValue(iNumOfRows, iDColumn, oOperation.LaborCosts);
                        oSLDocument.SetCellStyle(iNumOfRows, iDColumn, g_oCurrencyStyleCentered);

                        //
                        // DO WE NEED TO INCLUDE ANY SCRAP FACTORS ON OPERATIONS TO GET THE $$$ AMOUNT CORRECT?????
                        //
                        decimal dTmpLaborPrice = 0M;
                        if (oQuote.PercentType == "P")
                        {
                            // profit calculation
                            dTmpLaborPrice = (oOperation.LaborCosts + (oOperation.LaborCosts * oQuote.LaborMarkup / 100.0M));
                        }
                        else
                        {
                            // markup calculation
                            dTmpLaborPrice = oOperation.LaborCosts / ((100.0M - oQuote.LaborMarkup) / 100.0M);
                        }

                        if (dTotalOpHours != 0)
                        {
                            oSLDocument.SetCellValue(iNumOfRows, iFColumn, oOperation.BurdenCosts / (oOperation.TotalProductionHours + oOperation.TotalSetupHours));
                        }
                        else
                        {
                            oSLDocument.SetCellValue(iNumOfRows, iFColumn, 0);
                        }
                        oSLDocument.SetCellStyle(iNumOfRows, iFColumn, g_oCurrencyStyleCentered);
                        oSLDocument.SetCellValue(iNumOfRows, iGColumn, oOperation.BurdenCosts);
                        oSLDocument.SetCellStyle(iNumOfRows, iGColumn, g_oCurrencyStyleCentered);
                        decimal dTmpBurdenPrice = 0M;
                        if (oQuote.PercentType == "%")
                        {
                            dTmpBurdenPrice = (oOperation.BurdenCosts + (oOperation.BurdenCosts * oQuote.BurdenMarkup / 100.0M));
                        }
                        else
                        {
                            dTmpBurdenPrice = oOperation.BurdenCosts / ((100.0M - oQuote.BurdenMarkup) / 100.0M);
                        }

                        dTotalHours += oOperation.TotalProductionHours + oOperation.TotalSetupHours;
                        dTotalLabor += oOperation.LaborCosts;
                        dTotalLaborPrice += dTmpLaborPrice;
                        dTotalBurden += oOperation.BurdenCosts;
                        dTotalBurdenPrice += dTmpBurdenPrice;
                    }
                }

                // now compute the totals
                iNumOfRows++;
                oSLDocument.SetCellValue(iNumOfRows, iAColumn, "TOTAL");
                oSLDocument.SetCellValue(iNumOfRows, iBColumn, dTotalHours);
                oSLDocument.SetCellStyle(iNumOfRows, iBColumn, g_oDecimalBoxStyleBoldCentered);
                oSLDocument.SetCellValue(iNumOfRows, iDColumn, dTotalLabor);
                oSLDocument.SetCellStyle(iNumOfRows, iDColumn, g_oBoxStyleCentered);

                oSLDocument.SetCellValue(iNumOfRows, iGColumn, dTotalBurden);
                oSLDocument.SetCellStyle(iNumOfRows, iGColumn, g_oBoxStyleCentered);

                iNumOfRows++;
                // blank row

                iNumOfRows++;
                oSLDocument.SetCellValue(iNumOfRows, iAColumn, "5A.  SALES, G&A EXPENSES");
                oSLDocument.SetCellStyle(iNumOfRows, iAColumn, g_oBoldUnderline);
                decimal dPercentOfFactoryMfgCost = 0M;
                if (dTotalCosts != 0)
                {
                    dPercentOfFactoryMfgCost = dGAExpenses / dTotalCosts;
                }
                oSLDocument.SetCellValue(iNumOfRows, iCColumn, dPercentOfFactoryMfgCost);
                oSLDocument.SetCellStyle(iNumOfRows, iCColumn, g_oPercentStyleBoldCentered);
                oSLDocument.SetCellValue(iNumOfRows, iDColumn, "OF FACTORY MFG COST");

                iNumOfRows++;
                // blank row

                iNumOfRows++;
                oSLDocument.SetCellValue(iNumOfRows, iAColumn, "6A.  PROFIT OR FEE");
                oSLDocument.SetCellStyle(iNumOfRows, iAColumn, g_oBoldUnderline);
                decimal dPercentOfTotalCost = 0M;
                if (oQuote.TotalCost != 0)
                {
                    dPercentOfTotalCost = oQuote.TotalProfit / oQuote.TotalCost;
                }
                oSLDocument.SetCellValue(iNumOfRows, iCColumn, dPercentOfTotalCost);
                oSLDocument.SetCellStyle(iNumOfRows, iCColumn, g_oPercentStyleBoldCentered);
                oSLDocument.SetCellValue(iNumOfRows, iDColumn, "OF TOTAL COST");
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
                HSEmailHelper.SendEmail(oToAddresses, "SF 1411 For Quote  " + oFirstQuote.QuoteNum.ToString(), "SF 1411 For Quote  " + oFirstQuote.QuoteNum.ToString(), oAttachments);
            }
        }

        public static void CreateAndSendWhatIfReport(string sTmpFileDirectory, HSUser oRequestingUser, List<HSQuote> oQuotes)
        {
            // get the file name
            DateTime dtToday = DateTime.Now;
            string sNameOfMonth = dtToday.ToString("MMMM");

            // these are all for the same quote so we just get the first quote in the list
            HSQuote oFirstQuote = oQuotes[0];
            string sDestinationFileName = sTmpFileDirectory + "\\QuoteSF1411-" + oFirstQuote.QuoteNum.ToString() + ".xlsx";

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
            SetStyles(oSLDocument);

            bool bFirstWorksheet = true;
            foreach (HSQuote oQuote in oQuotes)
            {
                int iStartRowForParts = 0;
                int iEndRowForParts = 0;
                int iStartRowForSubcontracts = 0;
                int iEndRowForSubcontracts = 0;
                int iStartRowForLabor = 0;
                int iEndRowForLabor = 0;

                // we need to generate the basic workheet
                string sWorksheet = oQuote.QuoteNum.ToString() + "-" + oQuote.QuoteLine.ToString() + "-" + oQuote.m_iQuoteQtyNum.ToString();
                GenerateWorksheet(oSLDocument, sWorksheet, oQuote, bFirstWorksheet, out iStartRowForParts, out iEndRowForParts, out iStartRowForSubcontracts, out iEndRowForSubcontracts, out iStartRowForLabor, out iEndRowForLabor);
                bFirstWorksheet = false;

                // now we need to generate the what if worksheet for this same quote
                string sWorksheetWhatIf = oQuote.QuoteNum.ToString() + "-" + oQuote.QuoteLine.ToString() + "-" + oQuote.m_iQuoteQtyNum.ToString() + " (What If)";
                GenerateWorksheet(oSLDocument, sWorksheetWhatIf, oQuote, bFirstWorksheet, out iStartRowForParts, out iEndRowForParts, out iStartRowForSubcontracts, out iEndRowForSubcontracts, out iStartRowForLabor, out iEndRowForLabor);

                // now we need to generate the quote summary worksheet
                string sQuoteSummary = oQuote.QuoteNum.ToString() + "-" + oQuote.QuoteLine.ToString() + "-" + oQuote.m_iQuoteQtyNum.ToString() + " Summary";
                GenerateQuoteSummary(oSLDocument, sWorksheet, sWorksheetWhatIf, sQuoteSummary, oQuote, bFirstWorksheet, iStartRowForParts, iEndRowForParts, iStartRowForSubcontracts, iEndRowForSubcontracts, iStartRowForLabor, iEndRowForLabor);
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
                HSEmailHelper.SendEmail(oToAddresses, "SF 1411 For Quote  " + oFirstQuote.QuoteNum.ToString(), "SF 1411 For Quote  " + oFirstQuote.QuoteNum.ToString(), oAttachments);
            }
        }

        public static void CreateQuoteCostBreakdown(string sTemplateName, string sTmpFileDirectory, HSUser oRequestingUser, List<HSQuote> oQuotes)
        {
            // get the file name
            DateTime dtToday = DateTime.Now;
            string sNameOfMonth = dtToday.ToString("MMMM");

            int iQuoteNum = 0;
            int iLastLineNumber = 0;
            if (oQuotes.Count > 0)
            {
                // sort these by quote line number
                oQuotes = oQuotes.OrderBy(oItem => oItem.QuoteLine).ToList();

                // get the first one
                HSQuote oTmpQuote = oQuotes[0];
                iQuoteNum = oTmpQuote.m_iQuoteNum;
                iLastLineNumber = oTmpQuote.QuoteLine;
            }

            string sDestinationFileName = sTmpFileDirectory + "\\QuoteCostBreakdown-" + iQuoteNum.ToString() + ".xlsx";

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

            // we load the template
            SLDocument oSLDocument = new SLDocument(sTemplateName);

            SetStyles(oSLDocument);

            //set up column headers
            int iAColumn = 1;
            int iBColumn = 2;
            int iCColumn = 3;
            int iDColumn = 4;
            int iEColumn = 5;
            int iFColumn = 6;
            int iGColumn = 7;
            int iHColumn = 8;
            int iIColumn = 9;
            int iJColumn = 10;
            int iKColumn = 11;
            int iLColumn = 12;
            int iMColumn = 13;

            // we start at row 3 as the template already has the header filled in on the first two rows
            int iNumOfRows = 3;

            // we will start with style 1 and then rotate to style 2
            g_bUsingStyle1 = true;
            SLStyle oCurrentMoneyStyle = g_oCurrencyStyle1;
            SLStyle oCurrentStyle = g_oRowStyle1;

            string sHeader = @"Epicor Cost Breakdown for Q" + iQuoteNum.ToString();
            oSLDocument.SetCellValue(1, 1, sHeader);

            foreach (HSQuote oQuote in oQuotes)
            {
                // we alternate styles when the quote line changes
                if (oQuote.QuoteLine != iLastLineNumber)
                {
                    iLastLineNumber = oQuote.QuoteLine;
                    if (g_bUsingStyle1 == true)
                    {
                        oCurrentMoneyStyle = g_oCurrencyStyle2;
                        oCurrentStyle = g_oRowStyle2;
                        g_bUsingStyle1 = false;
                    }
                    else
                    {
                        oCurrentMoneyStyle = g_oCurrencyStyle1;
                        oCurrentStyle = g_oRowStyle1;
                        g_bUsingStyle1 = true;
                    }
                }

                oSLDocument.SetCellStyle(iNumOfRows, iAColumn, oCurrentStyle);
                oSLDocument.SetCellValue(iNumOfRows, iAColumn, oQuote.PartNum);

                oSLDocument.SetCellStyle(iNumOfRows, iBColumn, oCurrentStyle);
                oSLDocument.SetCellValue(iNumOfRows, iBColumn, oQuote.QuoteLine);

                oSLDocument.SetCellStyle(iNumOfRows, iCColumn, oCurrentStyle);
                oSLDocument.SetCellValue(iNumOfRows, iCColumn, oQuote.PartDescription);

                oSLDocument.SetCellStyle(iNumOfRows, iDColumn, oCurrentStyle);
                oSLDocument.SetCellValue(iNumOfRows, iDColumn, oQuote.SellingQty);

                //
                // sum all labor and burden costs
                //
                decimal dTotalProductionHours = 0;
                decimal dTotalSetupHours = 0;
                decimal dTotalLabor = 0;
                decimal dTotalBurden = 0;
                foreach (HSQuoteOperations oOperation in oQuote.AllOperations)
                {
                    // we only include operations that are not subcontracts
                    if (oOperation.SubcontractCosts == 0)
                    {
                        decimal dTotalOpHours = oOperation.TotalProductionHours + oOperation.TotalSetupHours;

                        dTotalProductionHours += oOperation.TotalProductionHours;
                        dTotalSetupHours += oOperation.TotalSetupHours;
                        dTotalLabor += oOperation.LaborCosts;
                        dTotalBurden += oOperation.BurdenCosts;
                    }
                }
                // set the cost for burden
                oSLDocument.SetCellStyle(iNumOfRows, iEColumn, oCurrentMoneyStyle);
                oSLDocument.SetCellValue(iNumOfRows, iEColumn, dTotalBurden);
                // set the cost for labor
                oSLDocument.SetCellStyle(iNumOfRows, iFColumn, oCurrentMoneyStyle);
                oSLDocument.SetCellValue(iNumOfRows, iFColumn, dTotalLabor);
                // set the number of production hours
                oSLDocument.SetCellStyle(iNumOfRows, iLColumn, oCurrentStyle);
                oSLDocument.SetCellValue(iNumOfRows, iLColumn, dTotalProductionHours);
                // set the number of setup hours
                oSLDocument.SetCellStyle(iNumOfRows, iMColumn, oCurrentStyle);
                oSLDocument.SetCellValue(iNumOfRows, iMColumn, dTotalSetupHours);

                //
                // sum all material costs
                //
                List<SF1411Material> oAllMaterialsWithCost = oQuote.m_oSF1411Materials.Where(x => x.TotalCost != 0).ToList();
                decimal dTotalPurchaseCostOfAllParts = 0M;
                foreach (SF1411Material oMaterialWithCost in oAllMaterialsWithCost)
                {
                    // check for fixed qty
                    decimal dBaseQty = 0M;
                    if (oMaterialWithCost.FixedQty == true)
                    {
                        dBaseQty = oMaterialWithCost.RequiredQty;
                    }
                    else
                    {
                        dBaseQty = oMaterialWithCost.RequiredQty * oQuote.SellingQty;
                    }

                    // variable that holds the total material quantity that will be required for this job -- affected by scrap
                    decimal dMaterialTotalQuantity = dBaseQty;

                    // consider scrap
                    if (oMaterialWithCost.ScrapType == "%")
                    {
                        // if we have an estimated scrap percentage then we need to modify the total material required
                        if (oMaterialWithCost.EstScrap != 0)
                        {
                            dMaterialTotalQuantity += dMaterialTotalQuantity * (oMaterialWithCost.EstScrap / 100.0M);
                        }
                    }
                    else
                    {
                        // if we have an estimated scrap quantity then we need to modify the total material required
                        if (oMaterialWithCost.EstScrap != 0)
                        {
                            dMaterialTotalQuantity += oMaterialWithCost.EstScrap;
                        }
                    }

                    // now see if we need to consider price breaks for all materials
                    // check to see if a price break is required -- there are exactly 10 price breaks present
                    for (int iCurrentPosition = 0; iCurrentPosition < 10; iCurrentPosition++)
                    {
                        // if the quantity is zero then it is not in effect
                        if (oMaterialWithCost.PriceBreakQuantities[iCurrentPosition] != 0)
                        {
                            if (dMaterialTotalQuantity >= oMaterialWithCost.PriceBreakQuantities[iCurrentPosition])
                            {
                                oMaterialWithCost.EstMtlUnitCost = oMaterialWithCost.PriceBreakCosts[iCurrentPosition];
                                oMaterialWithCost.TotalCost = oMaterialWithCost.EstMtlUnitCost * dMaterialTotalQuantity;
                            }
                        }
                    }

                    decimal dTotalCost = oMaterialWithCost.EstMtlUnitCost * dMaterialTotalQuantity;
                    // we need to check if minimum price is in effect
                    if (dTotalCost < oMaterialWithCost.MinimumCost)
                    {
                        dTotalCost = oMaterialWithCost.MinimumCost;
                    }

                    dTotalPurchaseCostOfAllParts += dTotalCost;
                }
                // set the customer price for materials
                oSLDocument.SetCellStyle(iNumOfRows, iGColumn, oCurrentMoneyStyle);
                oSLDocument.SetCellValue(iNumOfRows, iGColumn, dTotalPurchaseCostOfAllParts);

                //
                // sum all subcontract costs
                //
                List<SF1411Operation> oAllSubcontractsWithCost = oQuote.m_oSF1411Operations.Where(x => x.SubcontractCosts != 0).ToList();
                decimal dTotalSubcontractCosts = 0M;
                if (oAllSubcontractsWithCost.Count != 0)
                {
                    foreach (SF1411Operation oSubcontractWithCost in oAllSubcontractsWithCost)
                    {
                        dTotalSubcontractCosts += oSubcontractWithCost.ActualSubcontractCosts * oSubcontractWithCost.ActualQty;
                    }
                }
                // set the customer price for subcontracts
                oSLDocument.SetCellStyle(iNumOfRows, iHColumn, oCurrentMoneyStyle);
                oSLDocument.SetCellValue(iNumOfRows, iHColumn, dTotalSubcontractCosts);

                //
                // sum any additional costs
                //
                decimal dAdditionalCosts = oQuote.MiscCost;
                // set the customer price for additional costs
                oSLDocument.SetCellStyle(iNumOfRows, iIColumn, oCurrentMoneyStyle);
                oSLDocument.SetCellValue(iNumOfRows, iIColumn, dAdditionalCosts);

                //
                // compute total cost
                //
                decimal dTotalCosts = dTotalBurden + dTotalLabor + dTotalPurchaseCostOfAllParts + dTotalSubcontractCosts + dAdditionalCosts;
                // set the total price for the customer
                oSLDocument.SetCellStyle(iNumOfRows, iJColumn, oCurrentMoneyStyle);
                oSLDocument.SetCellValue(iNumOfRows, iJColumn, dTotalCosts);

                //
                // compute per unit price
                //
                decimal dPerUnitCosts = 0M;
                if (oQuote.SellingQty != 0)
                {
                    dPerUnitCosts = dTotalCosts / oQuote.SellingQty;
                }
                // set the customers per unit cost
                oSLDocument.SetCellStyle(iNumOfRows, iKColumn, oCurrentMoneyStyle);
                oSLDocument.SetCellValue(iNumOfRows, iKColumn, dPerUnitCosts);

                // increment to next row
                iNumOfRows++;
            }

            // save off the excel spreadsheet
            oSLDocument.SaveAs(sDestinationFileName);

            // email the document
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
                HSEmailHelper.SendEmail(oToAddresses, "Quote Cost Breakdown  " + iQuoteNum.ToString(), "Quote Cost Breakdown " + iQuoteNum.ToString(), oAttachments);
            }
        }

        public static void SetStyles(SLDocument oSLDocument)
        {
            // set up the style of cells
            g_oGood = oSLDocument.CreateStyle();
            g_oGood.ApplyNamedCellStyle(SLNamedCellStyleValues.Good);

            g_oNeutrual = oSLDocument.CreateStyle();
            g_oNeutrual.ApplyNamedCellStyle(SLNamedCellStyleValues.Neutral);

            g_oBad = oSLDocument.CreateStyle();
            g_oBad.ApplyNamedCellStyle(SLNamedCellStyleValues.Bad);

            g_oBold = oSLDocument.CreateStyle();
            g_oBold.SetFontBold(true);

            g_oUnderline = oSLDocument.CreateStyle();
            g_oUnderline.SetFontUnderline(UnderlineValues.Single);

            g_oCenterAlignment = new SLAlignment();
            g_oCenterAlignment.Horizontal = HorizontalAlignmentValues.Center;

            g_oRightAlignment = new SLAlignment();
            g_oRightAlignment.Horizontal = HorizontalAlignmentValues.Right;

            g_oRight = oSLDocument.CreateStyle();
            g_oRight.Alignment = g_oRightAlignment;

            g_oBoldRed = oSLDocument.CreateStyle();
            g_oBoldRed.SetFontBold(true);
            g_oBoldRed.SetFontColor(System.Drawing.Color.Red);

            g_oBoldUnderline = oSLDocument.CreateStyle();
            g_oBoldUnderline.SetFontBold(true);
            g_oBoldUnderline.SetFontUnderline(UnderlineValues.Single);

            g_oBoldUnderlineRight = oSLDocument.CreateStyle();
            g_oBoldUnderlineRight.SetFontBold(true);
            g_oBoldUnderlineRight.SetFontUnderline(UnderlineValues.Single);
            g_oBoldUnderlineRight.Alignment = g_oRightAlignment;

            g_oBoldCenter = oSLDocument.CreateStyle();
            g_oBoldCenter.SetFontBold(true);
            g_oBoldCenter.Alignment = g_oCenterAlignment;

            g_oBoldRight = oSLDocument.CreateStyle();
            g_oBoldRight.SetFontBold(true);
            g_oBoldRight.Alignment = g_oRightAlignment;

            g_oBoldCenterHeader = oSLDocument.CreateStyle();
            g_oBoldCenterHeader.SetFontBold(true);
            g_oBoldCenterHeader.SetFont(FontSchemeValues.Major, 14);
            g_oBoldCenterHeader.Alignment = g_oCenterAlignment;

            g_oCenter = oSLDocument.CreateStyle();
            g_oCenter.Alignment = g_oCenterAlignment;

            g_oUnderlineCenetered = oSLDocument.CreateStyle();
            g_oUnderlineCenetered.SetFontUnderline(UnderlineValues.Single);
            g_oUnderlineCenetered.Alignment = g_oCenterAlignment;

            g_oSLFill = new SLFill();
            g_oSLFill.SetPatternBackgroundColor(SLThemeColorIndexValues.Dark2Color);

            // create a box style
            g_oBoxStyleCentered = new SLStyle();
            g_oBoxStyleCentered.SetFontBold(true);
            g_oBoxStyleCentered.Alignment = g_oCenterAlignment;
            g_oBoxStyleCentered.SetTopBorder(BorderStyleValues.Thin, SLThemeColorIndexValues.Dark2Color);
            g_oBoxStyleCentered.SetBottomBorder(BorderStyleValues.Thin, SLThemeColorIndexValues.Dark2Color);
            g_oBoxStyleCentered.SetLeftBorder(BorderStyleValues.Thin, SLThemeColorIndexValues.Dark2Color);
            g_oBoxStyleCentered.SetRightBorder(BorderStyleValues.Thin, SLThemeColorIndexValues.Dark2Color);
            g_oBoxStyleCentered.SetFontColor(SLThemeColorIndexValues.Dark2Color);
            g_oBoxStyleCentered.FormatCode = "$#,##0.00";
            g_oBoxStyleCentered.Fill = g_oSLFill;

            g_oBoldCurrencyStyle = new SLStyle();
            g_oBoldCurrencyStyle.SetFontBold(true);
            g_oBoldCurrencyStyle.ApplyNamedCellStyle(SLNamedCellStyleValues.Currency);
            g_oBoldCurrencyStyle.Alignment = g_oRightAlignment;
            g_oBoldCurrencyStyle.SetTopBorder(BorderStyleValues.Thick, SLThemeColorIndexValues.Dark2Color);
            g_oBoldCurrencyStyle.SetBottomBorder(BorderStyleValues.Thick, SLThemeColorIndexValues.Dark2Color);
            g_oBoldCurrencyStyle.SetLeftBorder(BorderStyleValues.Thick, SLThemeColorIndexValues.Dark2Color);
            g_oBoldCurrencyStyle.SetRightBorder(BorderStyleValues.Thick, SLThemeColorIndexValues.Dark2Color);
            g_oBoldCurrencyStyle.FormatCode = "$#,##0.00";
            g_oBoldCurrencyStyle.Fill = g_oSLFill;

            g_oCurrencyStyleCentered = new SLStyle();
            g_oCurrencyStyleCentered.ApplyNamedCellStyle(SLNamedCellStyleValues.Currency);
            g_oCurrencyStyleCentered.Alignment = g_oCenterAlignment;
            g_oCurrencyStyleCentered.FormatCode = "$#,##0.00";

            g_oCurrencyStyleBoldCentered = new SLStyle();
            g_oCurrencyStyleBoldCentered.SetFontBold(true);
            g_oCurrencyStyleBoldCentered.ApplyNamedCellStyle(SLNamedCellStyleValues.Currency);
            g_oCurrencyStyleBoldCentered.Alignment = g_oCenterAlignment;
            g_oCurrencyStyleBoldCentered.FormatCode = "$#,##0.00";

            g_oDecimalBoxStyleBoldCentered = new SLStyle();
            g_oDecimalBoxStyleBoldCentered.SetFontBold(true);
            g_oDecimalBoxStyleBoldCentered.ApplyNamedCellStyle(SLNamedCellStyleValues.Currency);
            g_oDecimalBoxStyleBoldCentered.SetTopBorder(BorderStyleValues.Thin, SLThemeColorIndexValues.Dark2Color);
            g_oDecimalBoxStyleBoldCentered.SetBottomBorder(BorderStyleValues.Thin, SLThemeColorIndexValues.Dark2Color);
            g_oDecimalBoxStyleBoldCentered.SetLeftBorder(BorderStyleValues.Thin, SLThemeColorIndexValues.Dark2Color);
            g_oDecimalBoxStyleBoldCentered.SetRightBorder(BorderStyleValues.Thin, SLThemeColorIndexValues.Dark2Color);
            g_oDecimalBoxStyleBoldCentered.Alignment = g_oCenterAlignment;
            g_oDecimalBoxStyleBoldCentered.FormatCode = "###.00";

            g_oDecimalStyleCentered = new SLStyle();
            g_oDecimalStyleCentered.Alignment = g_oCenterAlignment;
            g_oDecimalStyleCentered.FormatCode = "###.00";

            g_oTextStyleHeaderCentered = new SLStyle();
            g_oTextStyleHeaderCentered.Alignment = g_oCenterAlignment;
            g_oTextStyleHeaderCentered.SetFontColor(System.Drawing.Color.White);
            g_oTextStyleHeaderCentered.SetPatternFill(PatternValues.Solid, System.Drawing.Color.Gray, System.Drawing.Color.Gray);

            g_oPercentStyleBoldCentered = new SLStyle();
            g_oPercentStyleBoldCentered.SetFontBold(true);
            g_oPercentStyleBoldCentered.ApplyNamedCellStyle(SLNamedCellStyleValues.Currency);
            g_oPercentStyleBoldCentered.Alignment = g_oCenterAlignment;
            g_oPercentStyleBoldCentered.FormatCode = "###.000%";

            g_oPercentStyleCentered = new SLStyle();
            g_oPercentStyleCentered.ApplyNamedCellStyle(SLNamedCellStyleValues.Currency);
            g_oPercentStyleCentered.Alignment = g_oCenterAlignment;
            g_oPercentStyleCentered.FormatCode = "###.000%";

            g_oLineTopStyle = new SLStyle();
            g_oLineTopStyle.Alignment = g_oCenterAlignment;
            g_oLineTopStyle.SetTopBorder(BorderStyleValues.Thin, SLThemeColorIndexValues.Dark2Color);

            g_oLineLeftStyle = new SLStyle();
            g_oLineLeftStyle.Alignment = g_oCenterAlignment;
            g_oLineLeftStyle.SetLeftBorder(BorderStyleValues.Thin, SLThemeColorIndexValues.Dark2Color);

            g_oLineRightStyle = new SLStyle();
            g_oLineRightStyle.Alignment = g_oCenterAlignment;
            g_oLineRightStyle.SetRightBorder(BorderStyleValues.Thin, SLThemeColorIndexValues.Dark2Color);

            g_oLineBottomStyle = new SLStyle();
            g_oLineBottomStyle.Alignment = g_oCenterAlignment;
            g_oLineBottomStyle.SetBottomBorder(BorderStyleValues.Thin, SLThemeColorIndexValues.Dark2Color);

            g_oRowStyle1 = new SLStyle();
            //g_oRowStyle1.Fill.SetPattern(PatternValues.Solid, SLThemeColorIndexValues.Light1Color, SLThemeColorIndexValues.Accent1Color);

            g_oRowStyle2 = new SLStyle();
            //g_oRowStyle2.Fill.SetPattern(PatternValues.Solid, SLThemeColorIndexValues.Light2Color, SLThemeColorIndexValues.Accent2Color);

            g_oCurrencyStyle1 = new SLStyle();
            //g_oCurrencyStyle1.Fill.SetPattern(PatternValues.Solid, SLThemeColorIndexValues.Light1Color, SLThemeColorIndexValues.Accent1Color);
            //g_oCurrencyStyle1.Alignment = g_oCenterAlignment;
            g_oCurrencyStyle1.FormatCode = "$#,##0.00";

            g_oCurrencyStyle2 = new SLStyle();
            //g_oCurrencyStyle2.Fill.SetPattern(PatternValues.Solid, SLThemeColorIndexValues.Light2Color, SLThemeColorIndexValues.Accent2Color);
            //g_oCurrencyStyle2.Alignment = g_oCenterAlignment;
            g_oCurrencyStyle2.FormatCode = "$#,##0.00";
        }

        public static void GenerateWorksheet(SLDocument oSLDocument, string sWorksheetName, HSQuote oQuote, bool bFirstWorksheet, out int iStartRowForParts, out int iEndRowForParts, out int iStartRowForSubcontracts, out int iEndRowForSubcontracts, out int iStartRowForLabor, out int iEndRowForLabor)
        {
            iStartRowForParts = 0;
            iEndRowForParts = 0;
            iStartRowForSubcontracts = 0;
            iEndRowForSubcontracts = 0;
            iStartRowForLabor = 0;
            iEndRowForLabor = 0;

            if (bFirstWorksheet == true)
            {
                oSLDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, sWorksheetName);
                bFirstWorksheet = false;
            }
            else
            {
                oSLDocument.AddWorksheet(sWorksheetName);
            }

            //set up column headers
            int iAColumn = 1;
            int iBColumn = 2;
            int iCColumn = 3;
            int iDColumn = 4;
            int iEColumn = 5;
            int iFColumn = 6;
            int iGColumn = 7;
            int iHColumn = 8;

            int iNumOfRows = 1;

            oSLDocument.SetCellValue(iNumOfRows, iAColumn, "QUOTE");
            oSLDocument.SetCellStyle(iNumOfRows, iAColumn, g_oBoldCenter);
            oSLDocument.SetCellValue(iNumOfRows, iBColumn, "LINE");
            oSLDocument.SetCellStyle(iNumOfRows, iBColumn, g_oBoldCenter);
            oSLDocument.SetCellValue(iNumOfRows, iCColumn, "QTY");
            oSLDocument.SetCellStyle(iNumOfRows, iCColumn, g_oBoldCenter);
            iNumOfRows++;
            oSLDocument.SetCellValue(iNumOfRows, iAColumn, oQuote.QuoteNum);
            oSLDocument.SetCellStyle(iNumOfRows, iAColumn, g_oCenter);
            oSLDocument.SetCellValue(iNumOfRows, iBColumn, oQuote.QuoteLine);
            oSLDocument.SetCellStyle(iNumOfRows, iBColumn, g_oCenter);
            oSLDocument.SetCellValue(iNumOfRows, iCColumn, oQuote.SellingQty);
            oSLDocument.SetCellStyle(iNumOfRows, iCColumn, g_oCenter);

            iNumOfRows++;
            // blank row
            iNumOfRows++;
            // blank row

            oSLDocument.SetCellValue(iNumOfRows, iAColumn, "SPINCRAFT-WI");
            oSLDocument.SetCellStyle(iNumOfRows, iAColumn, g_oBold);
            oSLDocument.SetCellValue(iNumOfRows, iDColumn, oQuote.CustomerName);

            iNumOfRows++;
            oSLDocument.SetCellValue(iNumOfRows, iAColumn, "2455 COMMERCE DRIVE");
            oSLDocument.SetCellValue(iNumOfRows, iDColumn, oQuote.PartDescription);
            oSLDocument.SetCellStyle(iNumOfRows, iDColumn, g_oBoldRed);

            iNumOfRows++;
            oSLDocument.SetCellValue(iNumOfRows, iAColumn, "NEW BERLIN, WI 53151");
            // get the date the quote was entered on
            oSLDocument.SetCellValue(iNumOfRows, iBColumn, oQuote.EntryDate.ToShortDateString());
            oSLDocument.SetCellStyle(iNumOfRows, iBColumn, g_oBoldCenter);
            oSLDocument.SetCellValue(iNumOfRows, iDColumn, "P/N " + oQuote.PartNum);

            iNumOfRows++;
            oSLDocument.SetCellValue(iNumOfRows, iAColumn, "SF 1411 REFERENCE SHEET");
            oSLDocument.SetCellStyle(iNumOfRows, iAColumn, g_oBold);

            iNumOfRows++;
            // blank row

            iNumOfRows++;
            oSLDocument.SetCellValue(iNumOfRows, iAColumn, "COST ELEMENTS:");
            oSLDocument.SetCellStyle(iNumOfRows, iAColumn, g_oBoldCenter);

            oSLDocument.SetCellValue(iNumOfRows, iBColumn, "UNIT COST");
            oSLDocument.SetCellStyle(iNumOfRows, iBColumn, g_oBoldCenter);

            iNumOfRows++;
            oSLDocument.SetCellValue(iNumOfRows, iAColumn, "RAW MATERIAL");

            // we will set the unit raw material costs after we know which rows we have to sum
            // at the end of the listing of material costs below
            int iRawMaterialTotalRow = iNumOfRows;
            int iRawMaterialTotalColumn = iBColumn;
            int iRawMaterialTotalStartRow = 0;
            int iRawMaterialTotalEndRow = 0;
            oSLDocument.SetCellStyle(iNumOfRows, iBColumn, g_oCurrencyStyleCentered);

            iNumOfRows++;
            oSLDocument.SetCellValue(iNumOfRows, iAColumn, "SUB-CONTRACTED ITEMS");

            // we will set the unit subcontracted costs after we know which rows we have to sum
            // at the end of the listing of subcontrac costs below
            int iSubcontractedItemsTotalRow = iNumOfRows;
            int iSubcontractedItemsTotalColumn = iBColumn;
            int iSubcontractedItemsTotalStartRow = 0;
            int iSubcontractedItemsTotalEndRow = 0;
            oSLDocument.SetCellStyle(iNumOfRows, iBColumn, g_oCurrencyStyleCentered);

            iNumOfRows++;
            oSLDocument.SetCellValue(iNumOfRows, iAColumn, "DIRECT LABOR");

            // we will set the unit direct labor costs after we know which rows we have to sum
            // at the end of the listing of labor costs below
            int iDirectLaborTotalRow = iNumOfRows;
            int iDirectLaborTotalColumn = iBColumn;
            oSLDocument.SetCellStyle(iNumOfRows, iBColumn, g_oCurrencyStyleCentered);

            iNumOfRows++;
            oSLDocument.SetCellValue(iNumOfRows, iAColumn, "MFG. OVERHEAD");

            // we will set the unit mfg burden costs after we know which rows we have to sum
            // at the end of the listing of mfg burden costs below
            int iMfgBurdenTotalRow = iNumOfRows;
            int iMfgBurdenTotalColumn = iBColumn;
            oSLDocument.SetCellStyle(iNumOfRows, iBColumn, g_oCurrencyStyleCentered);


            iNumOfRows++;
            oSLDocument.SetCellValue(iNumOfRows, iAColumn, "SUBTOTAL");
            oSLDocument.SetCellStyle(iNumOfRows, iAColumn, g_oBoldCenter);
            oSLDocument.SetCellValue(iNumOfRows, iBColumn, "=SUM(B10:B13)");
            oSLDocument.SetCellStyle(iNumOfRows, iBColumn, g_oCurrencyStyleBoldCentered);

            iNumOfRows++;
            oSLDocument.SetCellValue(iNumOfRows, iAColumn, "SALES, G&A EXPENSES");
            decimal dGAExpenses = oQuote.TotalCommission + oQuote.PriceMiscProfit;
            oSLDocument.SetCellValue(iNumOfRows, iBColumn, dGAExpenses / oQuote.SellingQty);
            oSLDocument.SetCellStyle(iNumOfRows, iBColumn, g_oCurrencyStyleCentered);

            iNumOfRows++;
            oSLDocument.SetCellValue(iNumOfRows, iAColumn, "GROSS MARGIN");
            // WE FORCE THE GROSS MARGIN TO 37%
            oSLDocument.SetCellValue(iNumOfRows, iBColumn, .37M);

            oSLDocument.SetCellStyle(iNumOfRows, iBColumn, g_oPercentStyleCentered);

            iNumOfRows++;
            oSLDocument.SetCellValue(iNumOfRows, iAColumn, "PROFIT OR FEE");
            oSLDocument.SetCellValue(iNumOfRows, iBColumn, "=B18-(B14 + B15)");
            oSLDocument.SetCellStyle(iNumOfRows, iBColumn, g_oCurrencyStyleBoldCentered);

            iNumOfRows++;
            oSLDocument.SetCellValue(iNumOfRows, iAColumn, "SELL PRICE");
            oSLDocument.SetCellStyle(iNumOfRows, iAColumn, g_oBoldCenter);
            oSLDocument.SetCellValue(iNumOfRows, iBColumn, "=(B14/(1-B16)) + B15");
            oSLDocument.SetCellStyle(iNumOfRows, iBColumn, g_oCurrencyStyleBoldCentered);

            iNumOfRows++;
            // blank row

            iNumOfRows++;
            oSLDocument.SetCellValue(iNumOfRows, iAColumn, "REFERENCES:");
            oSLDocument.SetCellStyle(iNumOfRows, iAColumn, g_oBold);

            iNumOfRows++;
            oSLDocument.SetCellValue(iNumOfRows, iAColumn, "1A.  RAW MATERIAL");
            oSLDocument.SetCellStyle(iNumOfRows, iAColumn, g_oBoldUnderline);

            //
            // we will list all materials that have a non-zero cost
            //
            // PART NUM, Description, Scrap, Vendor, Unit Cost
            List<SF1411Material> oAllMaterialsWithCost = oQuote.m_oSF1411Materials.Where(x => x.TotalCost != 0).ToList();
            if (oAllMaterialsWithCost.Count != 0)
            {
                // first create the header
                iNumOfRows++;
                oSLDocument.SetCellValue(iNumOfRows, iAColumn, "Part Number");
                oSLDocument.SetCellStyle(iNumOfRows, iAColumn, g_oBold);
                oSLDocument.SetCellValue(iNumOfRows, iBColumn, "Description");
                oSLDocument.SetCellStyle(iNumOfRows, iBColumn, g_oBold);
                oSLDocument.SetCellValue(iNumOfRows, iCColumn, "Scrap");
                oSLDocument.SetCellStyle(iNumOfRows, iCColumn, g_oBold);
                oSLDocument.SetCellValue(iNumOfRows, iDColumn, "Vendor");
                oSLDocument.SetCellStyle(iNumOfRows, iDColumn, g_oBold);
                oSLDocument.SetCellValue(iNumOfRows, iEColumn, "Unit Cost");
                oSLDocument.SetCellStyle(iNumOfRows, iEColumn, g_oBold);
                foreach (SF1411Material oMaterialWithCost in oAllMaterialsWithCost)
                {
                    iNumOfRows++;
                    // partnum
                    oSLDocument.SetCellValue(iNumOfRows, iAColumn, oMaterialWithCost.PartNum);

                    // description
                    oSLDocument.SetCellValue(iNumOfRows, iBColumn, oMaterialWithCost.PartDescription);

                    // check for fixed qty
                    decimal dBaseQty = 0M;
                    if (oMaterialWithCost.FixedQty == true)
                    {
                        dBaseQty = oMaterialWithCost.RequiredQty;
                    }
                    else
                    {
                        dBaseQty = oMaterialWithCost.RequiredQty * oQuote.SellingQty;
                    }
                    // variable that holds the total material quantity that will be required for this job -- affected by scrap
                    decimal dMaterialTotalQuantity = dBaseQty;

                    // consider scrap
                    if (oMaterialWithCost.ScrapType == "%")
                    {
                        oSLDocument.SetCellStyle(iNumOfRows, iCColumn, g_oPercentStyleBoldCentered);
                        oSLDocument.SetCellValue(iNumOfRows, iCColumn, oMaterialWithCost.EstScrap / 100.0M);
                        // if we have an estimated scrap percentage then we need to modify the total material required
                        if (oMaterialWithCost.EstScrap != 0)
                        {
                            dMaterialTotalQuantity += dMaterialTotalQuantity * (oMaterialWithCost.EstScrap / 100.0M);
                        }
                    }
                    else
                    {
                        oSLDocument.SetCellStyle(iNumOfRows, iCColumn, g_oDecimalBoxStyleBoldCentered);
                        oSLDocument.SetCellValue(iNumOfRows, iCColumn, oMaterialWithCost.EstScrap);
                        // if we have an estimated scrap quantity then we need to modify the total material required
                        if (oMaterialWithCost.EstScrap != 0)
                        {
                            dMaterialTotalQuantity += oMaterialWithCost.EstScrap;
                        }
                    }

                    // vendor
                    oSLDocument.SetCellValue(iNumOfRows, iDColumn, oMaterialWithCost.Vendor);

                    // we need to consider price breaks for all materials -- NEED TO EVALUATE WITH TOTAL REQUIRED QUANTITY
                    // check to see if a price break is required -- there are exactly 10 price breaks present
                    for (int iCurrentPosition = 0; iCurrentPosition < 10; iCurrentPosition++)
                    {
                        // if the quantity is zero then it is not in effect
                        if (oMaterialWithCost.PriceBreakQuantities[iCurrentPosition] != 0)
                        {
                            if (dMaterialTotalQuantity >= oMaterialWithCost.PriceBreakQuantities[iCurrentPosition])
                            {
                                oMaterialWithCost.EstMtlUnitCost = oMaterialWithCost.PriceBreakCosts[iCurrentPosition];
                                oMaterialWithCost.TotalCost = oMaterialWithCost.EstMtlUnitCost * dMaterialTotalQuantity;
                            }
                        }
                    }

                    decimal dTotalCost = oMaterialWithCost.EstMtlUnitCost * dMaterialTotalQuantity;
                    // we need to check if minimum price is in effect
                    if (dTotalCost < oMaterialWithCost.MinimumCost)
                    {
                        dTotalCost = oMaterialWithCost.MinimumCost;
                    }

                    // if the start row has not been set yet then we establish the start row
                    if (iRawMaterialTotalStartRow == 0)
                    {
                        iRawMaterialTotalStartRow = iNumOfRows;
                        iStartRowForParts = iNumOfRows;
                    }
                    // we always update the end row
                    iRawMaterialTotalEndRow = iNumOfRows;
                    iEndRowForParts = iNumOfRows;

                    // unit cost
                    oSLDocument.SetCellValue(iNumOfRows, iEColumn, dTotalCost / oQuote.SellingQty);
                    oSLDocument.SetCellStyle(iNumOfRows, iEColumn, g_oCurrencyStyleCentered);
                }
            }

            // add a row for the total parts
            iNumOfRows++;
            oSLDocument.SetCellValue(iNumOfRows, iAColumn, "Total Raw Materials");
            oSLDocument.SetCellStyle(iNumOfRows, iAColumn, g_oBoxStyleCentered);
            // total unit cost
            oSLDocument.SetCellStyle(iNumOfRows, iEColumn, g_oCurrencyStyleBoldCentered);

            // now we can set the total material costs if there were any
            if (iRawMaterialTotalStartRow != 0)
            {
                if (iRawMaterialTotalStartRow == iRawMaterialTotalEndRow)
                {
                    // there is only one row
                    oSLDocument.SetCellValue(iRawMaterialTotalRow, iRawMaterialTotalColumn, "=SUM(E" + iRawMaterialTotalStartRow.ToString() + ")");
                    oSLDocument.SetCellValue(iNumOfRows, iEColumn, "=SUM(E" + iRawMaterialTotalStartRow.ToString() + ")");
                }
                else
                {
                    // this means there are several rows
                    oSLDocument.SetCellValue(iRawMaterialTotalRow, iRawMaterialTotalColumn, "=SUM(E" + iRawMaterialTotalStartRow.ToString() + ":E" + iRawMaterialTotalEndRow.ToString() + ")");
                    oSLDocument.SetCellValue(iNumOfRows, iEColumn, "=SUM(E" + iRawMaterialTotalStartRow.ToString() + ":E" + iRawMaterialTotalEndRow.ToString() + ")");
                }
            }
            else
            {
                oSLDocument.SetCellValue(iNumOfRows, iFColumn, 0);
            }

            iNumOfRows++;
            // blank row

            iNumOfRows++;
            oSLDocument.SetCellValue(iNumOfRows, iAColumn, "2A.  SUB-CONTRACTED ITEMS");
            oSLDocument.SetCellStyle(iNumOfRows, iAColumn, g_oBoldUnderline);

            //
            // we will list all subcontracted items that have a non-zero cost
            //
            // Part Num, Descritpion, Base Qty, Scrap, Unit Price, Extended, Vendor, Cost
            List<SF1411Operation> oAllSubcontractsWithCost = oQuote.m_oSF1411Operations.Where(x => x.SubcontractCosts != 0).ToList();
            if (oAllSubcontractsWithCost.Count != 0)
            {
                // first create the header
                iNumOfRows++;
                oSLDocument.SetCellValue(iNumOfRows, iAColumn, "Part Number");
                oSLDocument.SetCellStyle(iNumOfRows, iAColumn, g_oBold);
                oSLDocument.SetCellValue(iNumOfRows, iBColumn, "Description");
                oSLDocument.SetCellStyle(iNumOfRows, iBColumn, g_oBold);
                oSLDocument.SetCellValue(iNumOfRows, iCColumn, "Scrap");
                oSLDocument.SetCellStyle(iNumOfRows, iCColumn, g_oBold);
                oSLDocument.SetCellValue(iNumOfRows, iDColumn, "Vendor");
                oSLDocument.SetCellStyle(iNumOfRows, iDColumn, g_oBold);
                oSLDocument.SetCellValue(iNumOfRows, iEColumn, "Unit Cost");
                oSLDocument.SetCellStyle(iNumOfRows, iEColumn, g_oBold);
                foreach (SF1411Operation oSubcontractWithCost in oAllSubcontractsWithCost)
                {
                    iNumOfRows++;
                    oSLDocument.SetCellValue(iNumOfRows, iAColumn, oSubcontractWithCost.PartNum);

                    oSLDocument.SetCellValue(iNumOfRows, iBColumn, oSubcontractWithCost.PartDescription);

                    if (string.Compare(oSubcontractWithCost.ScrapType, "%", true) == 0)
                    {
                        oSLDocument.SetCellStyle(iNumOfRows, iCColumn, g_oPercentStyleBoldCentered);
                        oSLDocument.SetCellValue(iNumOfRows, iCColumn, oSubcontractWithCost.EstScrap / 100.0M);
                    }
                    else
                    {
                        oSLDocument.SetCellStyle(iNumOfRows, iCColumn, g_oDecimalBoxStyleBoldCentered);
                        oSLDocument.SetCellValue(iNumOfRows, iCColumn, oSubcontractWithCost.EstScrap);
                    }

                    decimal dTotalSubcontractCostPerUnit = 0M;
                    if (oQuote.PercentType == "P")
                    {
                        // profit calculation
                        dTotalSubcontractCostPerUnit = ((oSubcontractWithCost.ActualSubcontractCosts * oSubcontractWithCost.ActualQty) + (oSubcontractWithCost.ActualSubcontractCosts * oSubcontractWithCost.ActualQty * oQuote.SubcontractMarkup / 100.0M));
                    }
                    else
                    {
                        // markup calculation
                        dTotalSubcontractCostPerUnit = (oSubcontractWithCost.ActualSubcontractCosts * oSubcontractWithCost.ActualQty) / ((100.0M - oQuote.SubcontractMarkup) / 100.0M);
                    }

                    oSLDocument.SetCellValue(iNumOfRows, iDColumn, oSubcontractWithCost.Vendor);

                    // do the total subcontract costs then divide by the numberof units to get the per unit subcontract cost
                    oSLDocument.SetCellValue(iNumOfRows, iEColumn, (oSubcontractWithCost.ActualSubcontractCosts * oSubcontractWithCost.ActualQty) / oQuote.SellingQty);
                    oSLDocument.SetCellStyle(iNumOfRows, iEColumn, g_oCurrencyStyleCentered);

                    // if the start row has not been set yet then we establish the start row
                    if (iSubcontractedItemsTotalStartRow == 0)
                    {
                        iSubcontractedItemsTotalStartRow = iNumOfRows;
                        iStartRowForSubcontracts = iNumOfRows;
                    }
                    // we always update the end row
                    iSubcontractedItemsTotalEndRow = iNumOfRows;
                    iEndRowForSubcontracts = iNumOfRows;
                }
                // add a row for the total unit subcontract costs
                iNumOfRows++;
                oSLDocument.SetCellValue(iNumOfRows, iAColumn, "Total Subcontracts");
                oSLDocument.SetCellStyle(iNumOfRows, iAColumn, g_oBoxStyleCentered);
            }

            // total unit cost
            // now we can set the total unit subcontract costs if there were any
            if (iSubcontractedItemsTotalStartRow != 0)
            {
                if (iSubcontractedItemsTotalStartRow == iSubcontractedItemsTotalEndRow)
                {
                    // there is only one row
                    oSLDocument.SetCellValue(iSubcontractedItemsTotalRow, iSubcontractedItemsTotalColumn, "=E" + iSubcontractedItemsTotalStartRow.ToString());
                    oSLDocument.SetCellValue(iNumOfRows, iEColumn, "=E" + iSubcontractedItemsTotalStartRow.ToString());
                }
                else
                {
                    // this means there are several rows
                    oSLDocument.SetCellValue(iSubcontractedItemsTotalRow, iSubcontractedItemsTotalColumn, "=SUM(E" + iSubcontractedItemsTotalStartRow.ToString() + ":E" + iSubcontractedItemsTotalEndRow.ToString() + ")");
                    oSLDocument.SetCellValue(iNumOfRows, iEColumn, "=SUM(E" + iSubcontractedItemsTotalStartRow.ToString() + ":E" + iSubcontractedItemsTotalEndRow.ToString() + ")");
                }
            }
            else
            {
                // no subcontract costs
                oSLDocument.SetCellValue(iNumOfRows, iEColumn, oQuote.SubcontractCost);
            }
            oSLDocument.SetCellStyle(iNumOfRows, iEColumn, g_oCurrencyStyleBoldCentered);



            iNumOfRows++;
            // blank row

            iNumOfRows++;
            oSLDocument.SetCellValue(iNumOfRows, iAColumn, "3A.  DIRECT LABOR");
            oSLDocument.SetCellStyle(iNumOfRows, iAColumn, g_oBoldUnderline);
            oSLDocument.SetCellValue(iNumOfRows, iFColumn, "4A.  MFG. OVERHEAD");
            oSLDocument.SetCellStyle(iNumOfRows, iFColumn, g_oBoldUnderline);

            iNumOfRows++;
            oSLDocument.SetCellValue(iNumOfRows, iAColumn, "COST CENTER");
            oSLDocument.SetCellStyle(iNumOfRows, iAColumn, g_oUnderline);
            oSLDocument.SetCellValue(iNumOfRows, iBColumn, "UNIT HOURS (without setup)");
            oSLDocument.SetCellStyle(iNumOfRows, iBColumn, g_oUnderlineCenetered);
            oSLDocument.SetCellValue(iNumOfRows, iCColumn, "TOTAL HOURS");
            oSLDocument.SetCellStyle(iNumOfRows, iCColumn, g_oUnderline);
            oSLDocument.SetCellValue(iNumOfRows, iDColumn, "LABOR RATE");
            oSLDocument.SetCellStyle(iNumOfRows, iDColumn, g_oUnderline);
            oSLDocument.SetCellValue(iNumOfRows, iEColumn, "LABOR COST");
            oSLDocument.SetCellStyle(iNumOfRows, iEColumn, g_oUnderline);

            oSLDocument.SetCellValue(iNumOfRows, iFColumn, "OVERHEAD RATE");
            oSLDocument.SetCellStyle(iNumOfRows, iFColumn, g_oUnderline);
            oSLDocument.SetCellValue(iNumOfRows, iGColumn, "OVERHEAD COST");
            oSLDocument.SetCellStyle(iNumOfRows, iGColumn, g_oUnderline);

            int iStartRowForOperations = 0;
            int iEndRowForOperations = 0;
            foreach (HSQuoteOperations oOperation in oQuote.AllOperations)
            {
                // we only include operations that are not subcontracts
                if (oOperation.SubcontractCosts == 0)
                {
                    decimal dTotalOpHours = oOperation.TotalProductionHours + oOperation.TotalSetupHours;

                    // list all operation names and details
                    iNumOfRows++;
                    if (iStartRowForOperations == 0)
                    {
                        iStartRowForOperations = iNumOfRows;
                        iStartRowForLabor = iNumOfRows;
                    }
                    iEndRowForOperations = iNumOfRows;
                    iEndRowForLabor = iNumOfRows;

                    // op code
                    oSLDocument.SetCellValue(iNumOfRows, iAColumn, oOperation.OpCode);

                    // per unit hours (except no setup)
                    // THIS WILL BE A FOMRULA DRIVEN BY THE GRID BELOW WHICH INCLUDES OP DETAILS
                    oSLDocument.SetCellStyle(iNumOfRows, iBColumn, g_oCenter);

                    // total hours
                    // THIS WILL BE A FOMRULA DRIVEN BY THE GRID BELOW WHICH INCLUDES OP DETAILS
                    oSLDocument.SetCellStyle(iNumOfRows, iCColumn, g_oCenter);

                    // labor rate
                    if (dTotalOpHours != 0)
                    {
                        oSLDocument.SetCellValue(iNumOfRows, iDColumn, oOperation.LaborCosts / (oOperation.TotalProductionHours + oOperation.TotalSetupHours));
                    }
                    else
                    {
                        oSLDocument.SetCellValue(iNumOfRows, iDColumn, 0);
                    }
                    oSLDocument.SetCellStyle(iNumOfRows, iDColumn, g_oCurrencyStyleCentered);

                    // labor costs
                    oSLDocument.SetCellValue(iNumOfRows, iEColumn, "=C" + iNumOfRows.ToString() + "*D" + iNumOfRows.ToString());
                    oSLDocument.SetCellStyle(iNumOfRows, iEColumn, g_oCurrencyStyleCentered);

                    // burden rate
                    if (dTotalOpHours != 0)
                    {
                        oSLDocument.SetCellValue(iNumOfRows, iFColumn, oOperation.BurdenCosts / (oOperation.TotalProductionHours + oOperation.TotalSetupHours));
                    }
                    else
                    {
                        oSLDocument.SetCellValue(iNumOfRows, iFColumn, 0);
                    }
                    oSLDocument.SetCellStyle(iNumOfRows, iFColumn, g_oCurrencyStyleCentered);

                    // burden costs
                    oSLDocument.SetCellValue(iNumOfRows, iGColumn, "=C" + iNumOfRows.ToString() + "*F" + iNumOfRows.ToString());
                    oSLDocument.SetCellStyle(iNumOfRows, iGColumn, g_oCurrencyStyleCentered);
                }
            }

            // now compute the totals
            iNumOfRows++;
            oSLDocument.SetCellValue(iNumOfRows, iAColumn, "TOTAL");

            oSLDocument.SetCellStyle(iNumOfRows, iBColumn, g_oDecimalBoxStyleBoldCentered);

            oSLDocument.SetCellStyle(iNumOfRows, iCColumn, g_oDecimalBoxStyleBoldCentered);

            // now we can set the total labor costs if there were any
            oSLDocument.SetCellStyle(iNumOfRows, iEColumn, g_oBoxStyleCentered);
            if (iStartRowForOperations != 0)
            {
                if (iStartRowForOperations == iEndRowForOperations)
                {
                    // only one row for labor

                    // set the total unit labor hours excluding setup
                    oSLDocument.SetCellValue(iNumOfRows, iBColumn, "=B" + iStartRowForOperations.ToString());

                    // set the total hours
                    oSLDocument.SetCellValue(iNumOfRows, iCColumn, "=C" + iStartRowForOperations.ToString());

                    // set the total labor cost
                    oSLDocument.SetCellValue(iNumOfRows, iEColumn, "=E" + iNumOfRows.ToString());

                    // set the labor cost element
                    oSLDocument.SetCellValue(iDirectLaborTotalRow, iDirectLaborTotalColumn, "=E" + iNumOfRows.ToString() + "/C2");
                }
                else
                {
                    // this means there are several rows

                    // set the total unit labor hours excluding setup
                    oSLDocument.SetCellValue(iNumOfRows, iBColumn, "=SUM(B" + iStartRowForOperations.ToString() + ":B" + iEndRowForOperations.ToString() + ")");

                    // set the total hours
                    oSLDocument.SetCellValue(iNumOfRows, iCColumn, "=SUM(C" + iStartRowForOperations.ToString() + ":C" + iEndRowForOperations.ToString() + ")");

                    // set the total labor costs
                    oSLDocument.SetCellValue(iNumOfRows, iEColumn, "=SUM(E" + iStartRowForOperations.ToString() + ":E" + iEndRowForOperations.ToString() + ")");

                    // set the labor cost element
                    oSLDocument.SetCellValue(iDirectLaborTotalRow, iDirectLaborTotalColumn, "=E" + iNumOfRows.ToString() + "/C2");
                }
            }
            else
            {
                // no direct labor costs
                oSLDocument.SetCellValue(iNumOfRows, iBColumn, 0);
                oSLDocument.SetCellValue(iNumOfRows, iCColumn, 0);
            }


            oSLDocument.SetCellStyle(iNumOfRows, iGColumn, g_oBoxStyleCentered);
            // now we can set the total burden costs if there were any
            if (iStartRowForOperations != 0)
            {
                if (iStartRowForOperations == iEndRowForOperations)
                {
                    // there is only one row

                    // set the total burden cost
                    oSLDocument.SetCellValue(iNumOfRows, iGColumn, "=G" + iStartRowForOperations.ToString());

                    // set the burden cost element
                    oSLDocument.SetCellValue(iMfgBurdenTotalRow, iMfgBurdenTotalColumn, "=G" + iNumOfRows.ToString() + "/C2");
                }
                else
                {
                    // this means there are several rows

                    // set the total burden cost
                    oSLDocument.SetCellValue(iNumOfRows, iGColumn, "=SUM(G" + iStartRowForOperations.ToString() + ":G" + iEndRowForOperations.ToString() + ")");

                    // set the burden cost element
                    oSLDocument.SetCellValue(iMfgBurdenTotalRow, iMfgBurdenTotalColumn, "=G" + iNumOfRows.ToString() + "/C2");
                }
            }
            else
            {
                // no direct labor costs
                oSLDocument.SetCellValue(iNumOfRows, iGColumn, 0);
            }

            iNumOfRows++;
            // blank row

            // we now add in all operations details
            iNumOfRows++;
            oSLDocument.SetCellValue(iNumOfRows, iAColumn, "ASSEMBLY SEQ");
            oSLDocument.SetCellStyle(iNumOfRows, iAColumn, g_oTextStyleHeaderCentered);
            oSLDocument.SetColumnWidth(iNumOfRows, iAColumn, 30);

            oSLDocument.SetCellValue(iNumOfRows, iBColumn, "OPERATION SEQ");
            oSLDocument.SetCellStyle(iNumOfRows, iBColumn, g_oTextStyleHeaderCentered);
            oSLDocument.SetColumnWidth(iNumOfRows, iBColumn, 60);

            oSLDocument.SetCellValue(iNumOfRows, iCColumn, "OPERATION DESCRIPTION");
            oSLDocument.SetCellStyle(iNumOfRows, iCColumn, g_oTextStyleHeaderCentered);
            oSLDocument.SetColumnWidth(iNumOfRows, iCColumn, 40);

            oSLDocument.SetCellValue(iNumOfRows, iDColumn, "OPERATION CODE");
            oSLDocument.SetCellStyle(iNumOfRows, iDColumn, g_oTextStyleHeaderCentered);
            oSLDocument.SetColumnWidth(iNumOfRows, iCColumn, 20);

            oSLDocument.SetCellValue(iNumOfRows, iEColumn, "SETUP HOURS");
            oSLDocument.SetCellStyle(iNumOfRows, iEColumn, g_oTextStyleHeaderCentered);
            oSLDocument.SetColumnWidth(iNumOfRows, iCColumn, 20);

            oSLDocument.SetCellValue(iNumOfRows, iFColumn, "PROD STD");
            oSLDocument.SetCellStyle(iNumOfRows, iFColumn, g_oTextStyleHeaderCentered);
            oSLDocument.SetColumnWidth(iNumOfRows, iCColumn, 20);

            oSLDocument.SetCellValue(iNumOfRows, iGColumn, "SUBCONTRACT");
            oSLDocument.SetCellStyle(iNumOfRows, iGColumn, g_oTextStyleHeaderCentered);
            oSLDocument.SetColumnWidth(iNumOfRows, iCColumn, 20);

            oSLDocument.SetCellValue(iNumOfRows, iHColumn, "TOTAL HOURS");
            oSLDocument.SetCellStyle(iNumOfRows, iHColumn, g_oTextStyleHeaderCentered);
            oSLDocument.SetColumnWidth(iNumOfRows, iCColumn, 20);

            // now walk through all of the operations
            int iStartRowForOpDetails = 0;
            int iEndRowForOpDetails = 0;
            foreach (SF1411Operation oOpDetail in oQuote.m_oSF1411Operations)
            {
                // list all operation names and details
                iNumOfRows++;
                if (iStartRowForOpDetails == 0)
                {
                    iStartRowForOpDetails = iNumOfRows;
                }
                iEndRowForOpDetails = iNumOfRows;

                // assembly seq
                oSLDocument.SetCellValue(iNumOfRows, iAColumn, oOpDetail.AssemblySeq);
                oSLDocument.SetCellStyle(iNumOfRows, iAColumn, g_oCenter);

                // op seq
                oSLDocument.SetCellValue(iNumOfRows, iBColumn, oOpDetail.OpSeq);
                oSLDocument.SetCellStyle(iNumOfRows, iBColumn, g_oCenter);

                // op description
                oSLDocument.SetCellValue(iNumOfRows, iCColumn, oOpDetail.OperationDesctiption);
                oSLDocument.SetCellStyle(iNumOfRows, iCColumn, g_oCenter);

                // op code
                oSLDocument.SetCellValue(iNumOfRows, iDColumn, oOpDetail.OpCode);
                oSLDocument.SetCellStyle(iNumOfRows, iDColumn, g_oCenter);

                // setup hours
                oSLDocument.SetCellValue(iNumOfRows, iEColumn, oOpDetail.TotalSetupHours);
                oSLDocument.SetCellStyle(iNumOfRows, iEColumn, g_oCenter);

                // production hours
                oSLDocument.SetCellValue(iNumOfRows, iFColumn, oOpDetail.TotalProductionHours / oQuote.SellingQty);
                oSLDocument.SetCellStyle(iNumOfRows, iFColumn, g_oCenter);

                // subcontract
                string sSubcontract = "FALSE";
                if ((oOpDetail.SubcontractCosts != 0) || (string.IsNullOrEmpty(oOpDetail.Vendor) == false))
                {
                    sSubcontract = "TRUE";
                }
                oSLDocument.SetCellValue(iNumOfRows, iGColumn, sSubcontract);
                oSLDocument.SetCellStyle(iNumOfRows, iGColumn, g_oCenter);

                // total hours
                oSLDocument.SetCellValue(iNumOfRows, iHColumn, "=E" + iNumOfRows.ToString() + "+(F" + iNumOfRows.ToString() + "*$C$2)");
                oSLDocument.SetCellStyle(iNumOfRows, iHColumn, g_oCenter);
            }

            // now that we have the op details filled in we can go back and set the total hours for the operations
            for (int iCurrentRow = iStartRowForOperations; iCurrentRow <= iEndRowForOperations; iCurrentRow++)
            {
                //
                //"=SUMIF($C$48:$C$66,A39,$E$48:$E$66)" -- UNIT HOURS
                //"=SUMIF($C$48:$C$66,A39,$G$48:$G$66)" -- TOTAL HOURS

                // fill in the formula for the unit hours
                oSLDocument.SetCellValue(iCurrentRow, iBColumn, "=SUMIF($D$" + iStartRowForOpDetails.ToString() + ":$D$" + iEndRowForOpDetails.ToString() + ",A" + iCurrentRow.ToString() + ",$F$" + iStartRowForOpDetails.ToString() + ":$F$" + iEndRowForOpDetails.ToString() + ")");

                // fill in the formula for the total hours
                oSLDocument.SetCellValue(iCurrentRow, iCColumn, "=SUMIF($D$" + iStartRowForOpDetails.ToString() + ":$D$" + iEndRowForOpDetails.ToString() + ",A" + iCurrentRow.ToString() + ",$H$" + iStartRowForOpDetails.ToString() + ":$H$" + iEndRowForOpDetails.ToString() + ")");
            }
        }

        public static void GenerateQuoteSummary(SLDocument oSLDocument, string sFirstWorksheetName, string sSecondWorksheetName, string sWorksheetName, HSQuote oQuote, bool bFirstWorksheet, int iStartRowForParts, int iEndRowForParts, int iStartRowForSubcontracts, int iEndRowForSubcontracts, int iStartRowForLabor, int iEndRowForLabor)
        {
            if (bFirstWorksheet == true)
            {
                oSLDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, sWorksheetName);
                bFirstWorksheet = false;
            }
            else
            {
                oSLDocument.AddWorksheet(sWorksheetName);
            }

            //set up column headers
            int iAColumn = 1;
            int iBColumn = 2;
            int iCColumn = 3;
            int iDColumn = 4;
            int iEColumn = 5;
            int iFColumn = 6;
            int iGColumn = 7;
            int iHColumn = 8;

            int iNumOfRows = 1;

            iNumOfRows++;
            // blank row

            oSLDocument.SetColumnWidth(iAColumn, 5);

            oSLDocument.SetCellValue(iNumOfRows, iBColumn, "Quote #");
            oSLDocument.SetColumnWidth(iBColumn, 35);

            oSLDocument.SetCellValue(iNumOfRows, iCColumn, oQuote.QuoteNum);
            oSLDocument.SetCellStyle(iNumOfRows, iCColumn, g_oBoldCenterHeader);
            oSLDocument.SetColumnWidth(iCColumn, 15);

            oSLDocument.SetColumnWidth(iDColumn, 5);

            oSLDocument.SetCellValue(iNumOfRows, iEColumn, "P/N " + oQuote.PartNum);
            oSLDocument.SetCellStyle(iNumOfRows, iEColumn, g_oBoldCenterHeader);
            oSLDocument.SetColumnWidth(iEColumn, 40);

            oSLDocument.SetColumnWidth(iFColumn, 20);

            oSLDocument.SetColumnWidth(iGColumn, 15);

            oSLDocument.SetColumnWidth(iHColumn, 5);

            iNumOfRows++;

            oSLDocument.SetCellValue(iNumOfRows, iBColumn, "Line");
            oSLDocument.SetCellValue(iNumOfRows, iCColumn, oQuote.QuoteLine);
            oSLDocument.SetCellStyle(iNumOfRows, iCColumn, g_oCenter);

            oSLDocument.SetCellValue(iNumOfRows, iEColumn, "Customer");
            oSLDocument.SetCellValue(iNumOfRows, iFColumn, oQuote.CustomerName);

            iNumOfRows++;

            oSLDocument.SetCellValue(iNumOfRows, iBColumn, "Qty");
            oSLDocument.SetCellValue(iNumOfRows, iCColumn, oQuote.SellingQty);
            oSLDocument.SetCellStyle(iNumOfRows, iCColumn, g_oCenter);

            oSLDocument.SetCellValue(iNumOfRows, iEColumn, "Description");
            oSLDocument.SetCellValue(iNumOfRows, iFColumn, oQuote.PartDescription);
            oSLDocument.SetCellStyle(iNumOfRows, iFColumn, g_oBoldRed);

            iNumOfRows++;
            // blank row
            iNumOfRows++;
            // blank row

            oSLDocument.SetCellValue(iNumOfRows, iCColumn, "Epicor");
            oSLDocument.SetCellStyle(iNumOfRows, iCColumn, g_oBoxStyleCentered);

            oSLDocument.SetCellValue(iNumOfRows, iEColumn, "What If");
            oSLDocument.SetCellStyle(iNumOfRows, iEColumn, g_oBoxStyleCentered);

            oSLDocument.SetCellValue(iNumOfRows, iGColumn, "Compare");
            oSLDocument.SetCellStyle(iNumOfRows, iGColumn, g_oBoxStyleCentered);

            iNumOfRows++;
            oSLDocument.SetCellValue(iNumOfRows, iBColumn, "COST ELEMENTS:");
            oSLDocument.SetCellStyle(iNumOfRows, iBColumn, g_oBoldRight);

            iNumOfRows++;

            // RAW MATERIALS COMPARISON
            oSLDocument.SetCellValue(iNumOfRows, iBColumn, "RAW MATERIAL");
            oSLDocument.SetCellStyle(iNumOfRows, iBColumn, g_oRight);

            oSLDocument.SetCellValue(iNumOfRows, iCColumn, "='" + sFirstWorksheetName + "'!B10");
            oSLDocument.SetCellStyle(iNumOfRows, iCColumn, g_oCurrencyStyleCentered);

            oSLDocument.SetCellValue(iNumOfRows, iEColumn, "='" + sSecondWorksheetName + "'!B10");
            oSLDocument.SetCellStyle(iNumOfRows, iEColumn, g_oCurrencyStyleCentered);

            oSLDocument.SetCellValue(iNumOfRows, iGColumn, "=E8-C8");
            oSLDocument.SetCellStyle(iNumOfRows, iGColumn, g_oCurrencyStyleCentered);

            iNumOfRows++;

            // SUBCONTRACTED ITEMS COMPARISON
            oSLDocument.SetCellValue(iNumOfRows, iBColumn, "SUB-CONTRACTED ITEMS");
            oSLDocument.SetCellStyle(iNumOfRows, iBColumn, g_oRight);

            oSLDocument.SetCellValue(iNumOfRows, iCColumn, "='" + sFirstWorksheetName + "'!B11");
            oSLDocument.SetCellStyle(iNumOfRows, iCColumn, g_oCurrencyStyleCentered);

            oSLDocument.SetCellValue(iNumOfRows, iEColumn, "='" + sSecondWorksheetName + "'!B11");
            oSLDocument.SetCellStyle(iNumOfRows, iEColumn, g_oCurrencyStyleCentered);

            oSLDocument.SetCellValue(iNumOfRows, iGColumn, "=E9-C9");
            oSLDocument.SetCellStyle(iNumOfRows, iGColumn, g_oCurrencyStyleCentered);

            iNumOfRows++;

            // DIRECT LABOR COMPARISON
            oSLDocument.SetCellValue(iNumOfRows, iBColumn, "DIRECT LABOR");
            oSLDocument.SetCellStyle(iNumOfRows, iBColumn, g_oRight);

            oSLDocument.SetCellValue(iNumOfRows, iCColumn, "='" + sFirstWorksheetName + "'!B12");
            oSLDocument.SetCellStyle(iNumOfRows, iCColumn, g_oCurrencyStyleCentered);

            oSLDocument.SetCellValue(iNumOfRows, iEColumn, "='" + sSecondWorksheetName + "'!B12");
            oSLDocument.SetCellStyle(iNumOfRows, iEColumn, g_oCurrencyStyleCentered);

            oSLDocument.SetCellValue(iNumOfRows, iGColumn, "=E10-C10");
            oSLDocument.SetCellStyle(iNumOfRows, iGColumn, g_oCurrencyStyleCentered);

            iNumOfRows++;

            // MFG OVERHEAD COMPARISON
            oSLDocument.SetCellValue(iNumOfRows, iBColumn, "MFG. OVERHEAD");
            oSLDocument.SetCellStyle(iNumOfRows, iBColumn, g_oRight);

            oSLDocument.SetCellValue(iNumOfRows, iCColumn, "='" + sFirstWorksheetName + "'!B13");
            oSLDocument.SetCellStyle(iNumOfRows, iCColumn, g_oCurrencyStyleCentered);

            oSLDocument.SetCellValue(iNumOfRows, iEColumn, "='" + sSecondWorksheetName + "'!B13");
            oSLDocument.SetCellStyle(iNumOfRows, iEColumn, g_oCurrencyStyleCentered);

            oSLDocument.SetCellValue(iNumOfRows, iGColumn, "=E11-C11");
            oSLDocument.SetCellStyle(iNumOfRows, iGColumn, g_oCurrencyStyleCentered);

            iNumOfRows++;

            // SALES AND G&A COMPARISON
            oSLDocument.SetCellValue(iNumOfRows, iBColumn, "SALES, G&A EXPENSES");
            oSLDocument.SetCellStyle(iNumOfRows, iBColumn, g_oRight);

            oSLDocument.SetCellValue(iNumOfRows, iCColumn, "='" + sFirstWorksheetName + "'!B14");
            oSLDocument.SetCellStyle(iNumOfRows, iCColumn, g_oCurrencyStyleCentered);

            oSLDocument.SetCellValue(iNumOfRows, iEColumn, "='" + sSecondWorksheetName + "'!B14");
            oSLDocument.SetCellStyle(iNumOfRows, iEColumn, g_oCurrencyStyleCentered);

            oSLDocument.SetCellValue(iNumOfRows, iGColumn, "=E12-C12");
            oSLDocument.SetCellStyle(iNumOfRows, iGColumn, g_oCurrencyStyleCentered);

            iNumOfRows++;

            // SUBTOTAL COMPARISON
            oSLDocument.SetCellValue(iNumOfRows, iBColumn, "SUBTOTAL");
            oSLDocument.SetCellStyle(iNumOfRows, iBColumn, g_oBoldRight);

            oSLDocument.SetCellValue(iNumOfRows, iCColumn, "='" + sFirstWorksheetName + "'!B15");
            oSLDocument.SetCellStyle(iNumOfRows, iCColumn, g_oCurrencyStyleCentered);

            oSLDocument.SetCellValue(iNumOfRows, iEColumn, "='" + sSecondWorksheetName + "'!B15");
            oSLDocument.SetCellStyle(iNumOfRows, iEColumn, g_oCurrencyStyleCentered);

            oSLDocument.SetCellValue(iNumOfRows, iGColumn, "=E13-C13");
            oSLDocument.SetCellStyle(iNumOfRows, iGColumn, g_oCurrencyStyleCentered);

            iNumOfRows++;

            // GROSS MARGIN COMPARISON
            oSLDocument.SetCellValue(iNumOfRows, iBColumn, "GROSS MARGIN");
            oSLDocument.SetCellStyle(iNumOfRows, iBColumn, g_oBoldRight);

            oSLDocument.SetCellValue(iNumOfRows, iCColumn, "='" + sFirstWorksheetName + "'!B16");
            oSLDocument.SetCellStyle(iNumOfRows, iCColumn, g_oPercentStyleCentered);

            oSLDocument.SetCellValue(iNumOfRows, iEColumn, "='" + sSecondWorksheetName + "'!B16");
            oSLDocument.SetCellStyle(iNumOfRows, iEColumn, g_oPercentStyleCentered);

            oSLDocument.SetCellValue(iNumOfRows, iGColumn, "=E14-C14");
            oSLDocument.SetCellStyle(iNumOfRows, iGColumn, g_oPercentStyleCentered);

            iNumOfRows++;

            // PROFIT OR FEE COMPARISON
            oSLDocument.SetCellValue(iNumOfRows, iBColumn, "PROFIT OR FEE");
            oSLDocument.SetCellStyle(iNumOfRows, iBColumn, g_oRight);

            oSLDocument.SetCellValue(iNumOfRows, iCColumn, "='" + sFirstWorksheetName + "'!B17");
            oSLDocument.SetCellStyle(iNumOfRows, iCColumn, g_oCurrencyStyleCentered);

            oSLDocument.SetCellValue(iNumOfRows, iEColumn, "='" + sSecondWorksheetName + "'!B17");
            oSLDocument.SetCellStyle(iNumOfRows, iEColumn, g_oCurrencyStyleCentered);

            oSLDocument.SetCellValue(iNumOfRows, iGColumn, "=E15-C15");
            oSLDocument.SetCellStyle(iNumOfRows, iGColumn, g_oCurrencyStyleCentered);


            iNumOfRows++;

            // SELL PRICE COMPARISON
            oSLDocument.SetCellValue(iNumOfRows, iBColumn, "SELL PRICE");
            oSLDocument.SetCellStyle(iNumOfRows, iBColumn, g_oBoldRight);

            oSLDocument.SetCellValue(iNumOfRows, iCColumn, "='" + sFirstWorksheetName + "'!B18");
            oSLDocument.SetCellStyle(iNumOfRows, iCColumn, g_oCurrencyStyleCentered);

            oSLDocument.SetCellValue(iNumOfRows, iEColumn, "='" + sSecondWorksheetName + "'!B18");
            oSLDocument.SetCellStyle(iNumOfRows, iEColumn, g_oCurrencyStyleCentered);

            oSLDocument.SetCellValue(iNumOfRows, iGColumn, "=E16-C16");
            oSLDocument.SetCellStyle(iNumOfRows, iGColumn, g_oCurrencyStyleCentered);

            iNumOfRows++;
            // blank row

            iNumOfRows++;

            // need to put a box around the raw material cells so we capture the start cell reference
            string sStartTopCellReference = "A" + iNumOfRows.ToString();
            string sEndTopCellReference = "H" + iNumOfRows.ToString();
            oSLDocument.SetCellStyle(sStartTopCellReference, sEndTopCellReference, g_oLineTopStyle);

            string sStartLeftCellReference = "A" + iNumOfRows.ToString();

            string sStartRightCellReference = "H" + iNumOfRows.ToString();

            iNumOfRows++;

            // RAW MATERIAL COMPARISON
            oSLDocument.SetCellValue(iNumOfRows, iBColumn, "1A.  RAW MATERIAL");
            oSLDocument.SetCellStyle(iNumOfRows, iBColumn, g_oBoldUnderlineRight);

            iNumOfRows++;

            oSLDocument.SetCellValue(iNumOfRows, iBColumn, "Part Number");
            oSLDocument.SetCellStyle(iNumOfRows, iBColumn, g_oBoldRight);

            for (int iCurrentRow = iStartRowForParts; iCurrentRow <= iEndRowForParts; iCurrentRow++)
            {
                iNumOfRows++;

                // part name from first worksheet
                oSLDocument.SetCellValue(iNumOfRows, iBColumn, "='" + sFirstWorksheetName + "'!A" + iCurrentRow.ToString());
                oSLDocument.SetCellStyle(iNumOfRows, iBColumn, g_oRight);

                // cost from first worksheet
                oSLDocument.SetCellValue(iNumOfRows, iCColumn, "='" + sFirstWorksheetName + "'!E" + iCurrentRow.ToString());
                oSLDocument.SetCellStyle(iNumOfRows, iCColumn, g_oCurrencyStyleCentered);

                // cost from second worksheet
                oSLDocument.SetCellValue(iNumOfRows, iEColumn, "='" + sSecondWorksheetName + "'!E" + iCurrentRow.ToString());
                oSLDocument.SetCellStyle(iNumOfRows, iEColumn, g_oCurrencyStyleCentered);

                // comparison
                oSLDocument.SetCellValue(iNumOfRows, iGColumn, "=E" + iNumOfRows.ToString() + "-C" + iNumOfRows.ToString());
                oSLDocument.SetCellStyle(iNumOfRows, iGColumn, g_oCurrencyStyleCentered);
            }

            iNumOfRows++;

            // add a left edge to box
            string sEndLeftCellReference = "A" + iNumOfRows.ToString();
            oSLDocument.SetCellStyle(sStartLeftCellReference, sEndLeftCellReference, g_oLineLeftStyle);

            // add a right edge to box
            string sEndRightCellReference = "H" + iNumOfRows.ToString();
            oSLDocument.SetCellStyle(sStartRightCellReference, sEndRightCellReference, g_oLineRightStyle);

            // add the bottom to the box
            string sStartBottomCellReference = "A" + iNumOfRows.ToString();
            string sEndBottomCellReference = "H" + iNumOfRows.ToString();
            oSLDocument.SetCellStyle(sStartBottomCellReference, sEndBottomCellReference, g_oLineBottomStyle);


            iNumOfRows++;
            // blank row

            iNumOfRows++;

            // need to put a box around the subcontract cells so we capture the start cell reference
            sStartTopCellReference = "A" + iNumOfRows.ToString();
            sEndTopCellReference = "H" + iNumOfRows.ToString();
            oSLDocument.SetCellStyle(sStartTopCellReference, sEndTopCellReference, g_oLineTopStyle);

            sStartLeftCellReference = "A" + iNumOfRows.ToString();

            sStartRightCellReference = "H" + iNumOfRows.ToString();

            iNumOfRows++;

            oSLDocument.SetCellValue(iNumOfRows, iBColumn, "2A.  SUB-CONTRACTED ITEMS");
            oSLDocument.SetCellStyle(iNumOfRows, iBColumn, g_oBoldUnderlineRight);

            iNumOfRows++;

            oSLDocument.SetCellValue(iNumOfRows, iBColumn, "Vendor");
            oSLDocument.SetCellStyle(iNumOfRows, iBColumn, g_oBoldRight);

            for (int iCurrentRow = iStartRowForSubcontracts; iCurrentRow <= iEndRowForSubcontracts; iCurrentRow++)
            {
                iNumOfRows++;

                // description from first worksheet
                oSLDocument.SetCellValue(iNumOfRows, iBColumn, "='" + sFirstWorksheetName + "'!B" + iCurrentRow.ToString());
                oSLDocument.SetCellStyle(iNumOfRows, iBColumn, g_oRight);

                // cost from first worksheet
                oSLDocument.SetCellValue(iNumOfRows, iCColumn, "='" + sFirstWorksheetName + "'!E" + iCurrentRow.ToString());
                oSLDocument.SetCellStyle(iNumOfRows, iCColumn, g_oCurrencyStyleCentered);

                // cost from second worksheet
                oSLDocument.SetCellValue(iNumOfRows, iEColumn, "='" + sSecondWorksheetName + "'!E" + iCurrentRow.ToString());
                oSLDocument.SetCellStyle(iNumOfRows, iEColumn, g_oCurrencyStyleCentered);

                // comparison
                oSLDocument.SetCellValue(iNumOfRows, iGColumn, "=E" + iNumOfRows.ToString() + "-C" + iNumOfRows.ToString());
                oSLDocument.SetCellStyle(iNumOfRows, iGColumn, g_oCurrencyStyleCentered);
            }

            iNumOfRows++;

            // add a left edge to box
            sEndLeftCellReference = "A" + iNumOfRows.ToString();
            oSLDocument.SetCellStyle(sStartLeftCellReference, sEndLeftCellReference, g_oLineLeftStyle);

            // add a right edge to box
            sEndRightCellReference = "H" + iNumOfRows.ToString();
            oSLDocument.SetCellStyle(sStartRightCellReference, sEndRightCellReference, g_oLineRightStyle);

            // add the bottom to the box
            sStartBottomCellReference = "A" + iNumOfRows.ToString();
            sEndBottomCellReference = "H" + iNumOfRows.ToString();
            oSLDocument.SetCellStyle(sStartBottomCellReference, sEndBottomCellReference, g_oLineBottomStyle);

            iNumOfRows++;
            // blank row

            iNumOfRows++;

            // need to put a box around the labor cells so we capture the start cell reference
            sStartTopCellReference = "A" + iNumOfRows.ToString();
            sEndTopCellReference = "H" + iNumOfRows.ToString();
            oSLDocument.SetCellStyle(sStartTopCellReference, sEndTopCellReference, g_oLineTopStyle);

            sStartLeftCellReference = "A" + iNumOfRows.ToString();

            sStartRightCellReference = "H" + iNumOfRows.ToString();

            oSLDocument.SetCellValue(iNumOfRows, iBColumn, "3A.  DIRECT LABOR");
            oSLDocument.SetCellStyle(iNumOfRows, iBColumn, g_oBoldUnderlineRight);

            for (int iCurrentRow = iStartRowForLabor; iCurrentRow <= iEndRowForLabor; iCurrentRow++)
            {
                iNumOfRows++;

                // op code from first worksheet
                oSLDocument.SetCellValue(iNumOfRows, iBColumn, "='" + sFirstWorksheetName + "'!A" + iCurrentRow.ToString());
                oSLDocument.SetCellStyle(iNumOfRows, iBColumn, g_oRight);

                // hours from first worksheet
                oSLDocument.SetCellValue(iNumOfRows, iCColumn, "='" + sFirstWorksheetName + "'!B" + iCurrentRow.ToString());
                oSLDocument.SetCellStyle(iNumOfRows, iCColumn, g_oCenter);

                // hours from second worksheet
                oSLDocument.SetCellValue(iNumOfRows, iEColumn, "='" + sSecondWorksheetName + "'!B" + iCurrentRow.ToString());
                oSLDocument.SetCellStyle(iNumOfRows, iEColumn, g_oCenter);

                // comparison
                oSLDocument.SetCellValue(iNumOfRows, iGColumn, "=E" + iNumOfRows.ToString() + "-C" + iNumOfRows.ToString());
                oSLDocument.SetCellStyle(iNumOfRows, iGColumn, g_oCenter);
            }

            iNumOfRows++;

            // add a row for totals

            oSLDocument.SetCellValue(iNumOfRows, iBColumn, "TOTAL");
            oSLDocument.SetCellStyle(iNumOfRows, iBColumn, g_oRight);

            // hours from first worksheet
            iEndRowForLabor += 1; // bump this to point to the totals row
            oSLDocument.SetCellValue(iNumOfRows, iCColumn, "='" + sFirstWorksheetName + "'!B" + iEndRowForLabor.ToString());
            oSLDocument.SetCellStyle(iNumOfRows, iCColumn, g_oDecimalBoxStyleBoldCentered);

            // hours from second worksheet
            oSLDocument.SetCellValue(iNumOfRows, iEColumn, "='" + sSecondWorksheetName + "'!B" + iEndRowForLabor.ToString());
            oSLDocument.SetCellStyle(iNumOfRows, iEColumn, g_oDecimalBoxStyleBoldCentered);

            // comparison
            oSLDocument.SetCellValue(iNumOfRows, iGColumn, "=E" + iNumOfRows.ToString() + "-C" + iNumOfRows.ToString());
            oSLDocument.SetCellStyle(iNumOfRows, iGColumn, g_oDecimalBoxStyleBoldCentered);

            iNumOfRows++;

            // add a left edge to box
            sEndLeftCellReference = "A" + iNumOfRows.ToString();
            oSLDocument.SetCellStyle(sStartLeftCellReference, sEndLeftCellReference, g_oLineLeftStyle);

            // add a right edge to box
            sEndRightCellReference = "H" + iNumOfRows.ToString();
            oSLDocument.SetCellStyle(sStartRightCellReference, sEndRightCellReference, g_oLineRightStyle);

            // add the bottom to the box
            sStartBottomCellReference = "A" + iNumOfRows.ToString();
            sEndBottomCellReference = "H" + iNumOfRows.ToString();
            oSLDocument.SetCellStyle(sStartBottomCellReference, sEndBottomCellReference, g_oLineBottomStyle);

        }

        #endregion

        #region Import Quotes
        static public bool ImportQuotes(Session oSession, string sOriginalFileName, HSUser oRequestingUser, string sTmpFileFolder)
        {
            bool bresult = false;

            // read in quote spreadsheet

            // delete any existing files with Quote in the name that are in the temp folder
            string[] oAllFiles = System.IO.Directory.GetFiles(sTmpFileFolder);
            foreach (string sFile in oAllFiles)
            {
                if (sFile.Contains("Quote") == true)
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

                List<string> oErrorMessages = new List<string>();

                // read in all data from this spreadsheet
                List<SLExcelData> oAllData = (new SLExcelReader()).ReadAllExcelSheets(sFileName);
                // now we will parse out each piece
                SLExcelData oQuoteHeaderData = oAllData.FirstOrDefault(x => string.Compare(x.SheetName, "QuoteHead", true) == 0);
                if (oQuoteHeaderData == null)
                {
                    oErrorMessages.Add("No quote header worksheet present!");
                }
                SLExcelData oQuoteLineData = oAllData.FirstOrDefault(x => string.Compare(x.SheetName, "QuoteLine", true) == 0);
                if (oQuoteLineData == null)
                {
                    oErrorMessages.Add("No quote line worksheet present!");
                }
                SLExcelData oOperationData = oAllData.FirstOrDefault(x => string.Compare(x.SheetName, "Operations", true) == 0);
                SLExcelData oMaterialData = oAllData.FirstOrDefault(x => string.Compare(x.SheetName, "Materials", true) == 0);
                SLExcelData oQuantityBreakData = oAllData.FirstOrDefault(x => string.Compare(x.SheetName, "QtyBreaks", true) == 0);

                List<HSQuoteHead> oQuoteHeaders = new List<HSQuoteHead>();
                List<string> oErrors = ParseQuoteHeadInformation(oQuoteHeaderData, oQuoteHeaders);
                oErrorMessages.AddRange(oErrors);

                List<HSQuoteLine> oQuoteLines = new List<HSQuoteLine>();
                oErrors = ParseQuoteLineInformation(oQuoteLineData, oQuoteLines);
                oErrorMessages.AddRange(oErrors);

                List<HSQuoteOperation> oQuoteOperations = new List<HSQuoteOperation>();
                oErrors = ParseQuoteOperationInformation(oOperationData, oQuoteOperations);
                oErrorMessages.AddRange(oErrors);

                List<HSQuoteMaterial> oQuoteMaterials = new List<HSQuoteMaterial>();
                oErrors = ParseQuoteMaterialInformation(oMaterialData, oQuoteMaterials);
                oErrorMessages.AddRange(oErrors);

                List<HSQuoteQuantityBreak> oQuoteQuantityBreaks = new List<HSQuoteQuantityBreak>();
                oErrors = ParseQuoteQuantityBreakInformation(oQuantityBreakData, oQuoteQuantityBreaks);
                oErrorMessages.AddRange(oErrors);

                // if we have no processing errors yet then try to import the data
                if (oErrorMessages.Count == 0)
                {
                    // now we will go through each quote and have them pull in their details
                    foreach (HSQuoteHead oQuoteHead in oQuoteHeaders)
                    {
                        oQuoteHead.LoadDetails(oQuoteLines, oQuoteOperations, oQuoteMaterials, oQuoteQuantityBreaks);
                    }

                    foreach (HSQuoteHead oQuote in oQuoteHeaders)
                    {
                        List<string> oQuoteErrors = oQuote.CreateQuoteHeader(oSession);
                        oErrorMessages.AddRange(oQuoteErrors);
                    }
                }

                // report out to user whether the quote was imported successfully or not
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
                    oStringBuilder.Append("The file named " + sOriginalFileName + " had erorrs while processing\n");
                    foreach (string sError in oErrorMessages)
                    {
                        oStringBuilder.Append(sError + "\n");
                    }
                }
                else
                {
                    oStringBuilder.Append("The file named " + sOriginalFileName + " was processed successfully\n");
                    // indicate there were no errors
                    bresult = true;
                }
                HSEmailHelper.SendEmail(oToAddresses, "Quote Imports Processed", oStringBuilder.ToString());

            }
            return bresult;
        }

        #region Read Quote Data From Spreadsheet
        private static List<string> ParseQuoteHeadInformation(SLExcelData oData, List<HSQuoteHead> oQuoteHeads)
        {
            List<string> oErrors = new List<string>();
            if (oData != null)
            {
                int iCounter = 0;
                foreach (List<string> sRowData in oData.DataRows)
                {
                    int iColumn = 0;

                    // if this is the first row then we need to first read in the quote header
                    if (iCounter == 0)
                    {
                        iColumn = 0;
                        foreach (string sCell in oData.Headers)
                        {
                            Console.WriteLine(sCell);
                            iColumn++;
                        }
                    }

                    HSQuoteHead oQuoteHead = new HSQuoteHead();
                    iCounter++;
                    // now we are reading in quote info
                    iColumn = 0;
                    int iTmp;
                    foreach (string sCell in sRowData)
                    {
                        if (string.IsNullOrEmpty(sCell) == false)
                        {
                            // the first column will be QuoteNum, then customer Id, then sales rep code
                            if (iColumn == 0)
                            {
                                if (int.TryParse(sCell.Trim(), out iTmp) == true)
                                {
                                    oQuoteHead.QuoteNum = iTmp;
                                }
                                else
                                {
                                    oErrors.Add("In the worksheet 'QuoteHead',On row: " + iCounter.ToString() + " - could not convert the value '" + sCell + "' to a quote number.");
                                }
                            }
                            else if (iColumn == 1)
                            {
                                oQuoteHead.CustomerID = sCell.Trim();
                            }
                            else if (iColumn == 2)
                            {
                                // this is optional
                                oQuoteHead.SalesRepCode = sCell.Trim();
                            }
                        }
                        iColumn++;
                    }
                    oQuoteHeads.Add(oQuoteHead);
                }
            }
            return oErrors;
        }

        public static List<string> ParseQuoteLineInformation(SLExcelData oData, List<HSQuoteLine> oQuoteLines)
        {
            List<string> oErrors = new List<string>();
            if (oData != null)
            {
                int iCounter = 0;
                foreach (List<string> sRowData in oData.DataRows)
                {
                    int iColumn = 0;

                    // if this is the first row then we need to first read in the quote header
                    if (iCounter == 0)
                    {
                        iColumn = 0;
                        foreach (string sCell in oData.Headers)
                        {
                            Console.WriteLine(sCell);
                            iColumn++;
                        }
                    }

                    HSQuoteLine oQuoteLine = new HSQuoteLine();
                    // now we are reading in quote info
                    iColumn = 0;
                    int iTmp;
                    decimal dTmp;
                    foreach (string sCell in sRowData)
                    {
                        if (string.IsNullOrEmpty(sCell) == false)
                        {
                            // the first column will be QuoteNum, then customer Id
                            if (iColumn == 0)
                            {
                                if (int.TryParse(sCell.Trim(), out iTmp) == true)
                                {
                                    oQuoteLine.QuoteNum = iTmp;
                                }
                                else
                                {
                                    oErrors.Add("In the worksheet 'QuoteLine', On row: " + iCounter.ToString() + " - could not convert the value '" + sCell + "' to a quote number.");
                                }
                            }
                            else if (iColumn == 1)
                            {
                                if (int.TryParse(sCell.Trim(), out iTmp) == true)
                                {
                                    oQuoteLine.QuoteLine = iTmp;
                                }
                                else
                                {
                                    oErrors.Add("In the worksheet 'QuoteLine', On row: " + iCounter.ToString() + " - could not convert the value '" + sCell + "' to a line number.");
                                }
                            }
                            else if (iColumn == 2)
                            {
                                oQuoteLine.PartNum = sCell.Trim();
                            }
                            else if (iColumn == 3)
                            {
                                oQuoteLine.Description = sCell.Trim();
                            }
                            else if (iColumn == 4)
                            {
                                if (decimal.TryParse(sCell.Trim(), out dTmp) == true)
                                {
                                    oQuoteLine.Quantity = dTmp;
                                }
                                else
                                {
                                    oErrors.Add("In the worksheet 'QuoteLine', On row: " + iCounter.ToString() + " - could not convert the value '" + sCell + "' to a quantity.");
                                }
                            }
                            else if (iColumn == 5)
                            {
                                if (decimal.TryParse(sCell.Trim(), out dTmp) == true)
                                {
                                    oQuoteLine.ExpectedUnitCost = dTmp;
                                }
                                else
                                {
                                    oErrors.Add("In the worksheet 'QuoteLine', On row: " + iCounter.ToString() + " - could not convert the value '" + sCell + "' to an expected unit price.");
                                }
                            }
                            else if (iColumn == 6)
                            {
                                oQuoteLine.ProductCode = sCell.Trim();
                            }
                        }
                        iColumn++;
                    }
                    oQuoteLines.Add(oQuoteLine);
                    iCounter++;
                }
            }
            return oErrors;
        }

        public static List<string> ParseQuoteOperationInformation(SLExcelData oData, List<HSQuoteOperation> oQuoteOperations)
        {
            List<string> oErrors = new List<string>();

            if (oData != null)
            {
                int iCounter = 0;
                foreach (List<string> sRowData in oData.DataRows)
                {
                    int iColumn = 0;

                    // if this is the first row then we need to first read in the operations header
                    if (iCounter == 0)
                    {
                        iColumn = 0;
                        foreach (string sCell in oData.Headers)
                        {
                            Console.WriteLine(sCell);
                            iColumn++;
                        }
                    }

                    HSQuoteOperation oQuoteOperation = new HSQuoteOperation();
                    // now we are reading in quote info
                    iColumn = 0;
                    int iTmp;
                    decimal dTmp;
                    bool bTmp;
                    foreach (string sCell in sRowData)
                    {
                        if (string.IsNullOrEmpty(sCell) == false)
                        {
                            // the first column will be QuoteNum, then customer Id
                            if (iColumn == 0)
                            {
                                if (int.TryParse(sCell.Trim(), out iTmp) == true)
                                {
                                    oQuoteOperation.QuoteNum = iTmp;
                                }
                                else
                                {
                                    oErrors.Add("In the worksheet 'Operations', On row: " + iCounter.ToString() + " - could not convert the value '" + sCell + "' to a quote number.");
                                }
                            }
                            else if (iColumn == 1)
                            {
                                if (int.TryParse(sCell.Trim(), out iTmp) == true)
                                {
                                    oQuoteOperation.QuoteLine = iTmp;
                                }
                                else
                                {
                                    oErrors.Add("In the worksheet 'Operations', On row: " + iCounter.ToString() + " - could not convert the value '" + sCell + "' to a line number.");
                                }
                            }
                            else if (iColumn == 2)
                            {
                                if (int.TryParse(sCell.Trim(), out iTmp) == true)
                                {
                                    oQuoteOperation.AssemblySequence = iTmp;
                                }
                                else
                                {
                                    oErrors.Add("In the worksheet 'Operations', On row: " + iCounter.ToString() + " - could not convert the value '" + sCell + "' to a assembly sequence.");
                                }
                            }
                            else if (iColumn == 3)
                            {
                                if (int.TryParse(sCell.Trim(), out iTmp) == true)
                                {
                                    oQuoteOperation.OperationSequence = iTmp;
                                }
                                else
                                {
                                    oErrors.Add("In the worksheet 'Operations', On row: " + iCounter.ToString() + " - could not convert the value '" + sCell + "' to a operation sequence.");
                                }
                            }
                            else if (iColumn == 4)
                            {
                                oQuoteOperation.OpCode = sCell.Trim();
                            }
                            else if (iColumn == 5)
                            {
                                oQuoteOperation.CommentText = sCell.Trim();
                            }
                            else if (iColumn == 6)
                            {
                                string sSubcontract = sCell.Trim();
                                if (string.Compare(sSubcontract, "0", true) == 0)
                                {
                                    oQuoteOperation.Subcontract = false;
                                }
                                else if (string.Compare(sSubcontract, "1", true) == 0)
                                {
                                    oQuoteOperation.Subcontract = true;
                                }
                                else
                                {
                                    if (bool.TryParse(sCell.Trim(), out bTmp) == true)
                                    {
                                        oQuoteOperation.Subcontract = bTmp;
                                    }
                                    else
                                    {
                                        oErrors.Add("In the worksheet 'Operations', On row: " + iCounter.ToString() + " - could not convert the value '" + sCell + "' to a subcontract boolean value.");
                                    }
                                }
                            }
                            else if (iColumn == 7)
                            {
                                if (decimal.TryParse(sCell.Trim(), out dTmp) == true)
                                {
                                    oQuoteOperation.ProductionStandard = dTmp;
                                }
                                else
                                {
                                    oErrors.Add("In the worksheet 'Operations', On row: " + iCounter.ToString() + " - could not convert the value '" + sCell + "' to a production standard number.");
                                }
                            }
                            else if (iColumn == 8)
                            {
                                if (decimal.TryParse(sCell.Trim(), out dTmp) == true)
                                {
                                    oQuoteOperation.ProdCrewSize = dTmp;
                                }
                                else
                                {
                                    oErrors.Add("In the worksheet 'Operations', On row: " + iCounter.ToString() + " - could not convert the value '" + sCell + "' to a production crew size number.");
                                }
                            }
                            else if (iColumn == 9)
                            {
                                oQuoteOperation.StandardFormat = sCell.Trim();
                            }
                            else if (iColumn == 10)
                            {
                                oQuoteOperation.StandardBasis = sCell.Trim();
                            }
                            else if (iColumn == 11)
                            {
                                if (decimal.TryParse(sCell.Trim(), out dTmp) == true)
                                {
                                    oQuoteOperation.QuantityPer = dTmp;
                                }
                                else
                                {
                                    oErrors.Add("In the worksheet 'Operations', On row: " + iCounter.ToString() + " - could not convert the value '" + sCell + "' to a quantity per number.");
                                }
                            }
                            else if (iColumn == 12)
                            {
                                if (decimal.TryParse(sCell.Trim(), out dTmp) == true)
                                {
                                    oQuoteOperation.SetupHours = dTmp;
                                }
                                else
                                {
                                    oErrors.Add("In the worksheet 'Operations', On row: " + iCounter.ToString() + " - could not convert the value '" + sCell + "' to a setup hours number.");
                                }
                            }
                            else if (iColumn == 13)
                            {
                                if (decimal.TryParse(sCell.Trim(), out dTmp) == true)
                                {
                                    oQuoteOperation.SetupCrewSize = dTmp;
                                }
                                else
                                {
                                    oErrors.Add("In the worksheet 'Operations', On row: " + iCounter.ToString() + " - could not convert the value '" + sCell + "' to a setup crew size number.");
                                }
                            }
                            else if (iColumn == 14)
                            {
                                oQuoteOperation.ResourceGroupID = sCell.Trim();
                            }
                            else if (iColumn == 15)
                            {
                                oQuoteOperation.ResourceGroupDescription = sCell.Trim();
                            }
                            else if (iColumn == 16)
                            {
                                oQuoteOperation.VendorId = sCell.Trim();
                            }
                            else if (iColumn == 17)
                            {
                                if (decimal.TryParse(sCell.Trim(), out dTmp) == true)
                                {
                                    oQuoteOperation.SubcontractCost = dTmp;
                                }
                                else
                                {
                                    oErrors.Add("In the worksheet 'Operations', On row: " + iCounter.ToString() + " - could not convert the value '" + sCell + "' to a subcontracting cost.");
                                }
                            }
                            else if (iColumn == 18)
                            {
                                if (decimal.TryParse(sCell.Trim(), out dTmp) == true)
                                {
                                    oQuoteOperation.DaysOut = dTmp;
                                }
                                else
                                {
                                    oErrors.Add("In the worksheet 'Operations', On row: " + iCounter.ToString() + " - could not convert the value '" + sCell + "' to the number of days required for the subcontract operation.");
                                }
                            }
                        }
                        iColumn++;
                    }
                    oQuoteOperations.Add(oQuoteOperation);
                    iCounter++;
                }
            }
            return oErrors;
        }

        public static List<string> ParseQuoteMaterialInformation(SLExcelData oData, List<HSQuoteMaterial> oQuoteMaterials)
        {
            List<string> oErrors = new List<string>();

            if (oData != null)
            {
                int iCounter = 0;
                foreach (List<string> sRowData in oData.DataRows)
                {
                    int iColumn = 0;

                    // if this is the first row then we need to first read in the quote header
                    if (iCounter == 0)
                    {
                        iColumn = 0;
                        foreach (string sCell in oData.Headers)
                        {
                            Console.WriteLine(sCell);
                            iColumn++;
                        }
                    }

                    HSQuoteMaterial oQuoteMaterial = new HSQuoteMaterial();
                    // now we are reading in quote info
                    iColumn = 0;
                    int iTmp;
                    decimal dTmp;
                    foreach (string sCell in sRowData)
                    {
                        if (string.IsNullOrEmpty(sCell) == false)
                        {
                            // the first column will be QuoteNum, then customer Id
                            if (iColumn == 0)
                            {
                                if (int.TryParse(sCell.Trim(), out iTmp) == true)
                                {
                                    oQuoteMaterial.QuoteNum = iTmp;
                                }
                                else
                                {
                                    oErrors.Add("In the worksheet 'Materials', On row: " + iCounter.ToString() + " - could not convert the value '" + sCell + "' to a quote number.");
                                }
                            }
                            else if (iColumn == 1)
                            {
                                if (int.TryParse(sCell.Trim(), out iTmp) == true)
                                {
                                    oQuoteMaterial.QuoteLine = iTmp;
                                }
                                else
                                {
                                    oErrors.Add("In the worksheet 'Materials', On row: " + iCounter.ToString() + " - could not convert the value '" + sCell + "' to a line number.");
                                }
                            }
                            else if (iColumn == 2)
                            {
                                if (int.TryParse(sCell.Trim(), out iTmp) == true)
                                {
                                    oQuoteMaterial.AssemblySequence = iTmp;
                                }
                                else
                                {
                                    oErrors.Add("In the worksheet 'Materials', On row: " + iCounter.ToString() + " - could not convert the value '" + sCell + "' to an assembly sequence number.");
                                }
                            }
                            else if (iColumn == 3)
                            {
                                if (int.TryParse(sCell.Trim(), out iTmp) == true)
                                {
                                    oQuoteMaterial.MaterialSequence = iTmp;
                                }
                                else
                                {
                                    oErrors.Add("In the worksheet 'Materials', On row: " + iCounter.ToString() + " - could not convert the value '" + sCell + "' to a material sequence number.");
                                }
                            }
                            else if (iColumn == 4)
                            {
                                oQuoteMaterial.PartNum = sCell.Trim();
                            }
                            else if (iColumn == 5)
                            {
                                if (decimal.TryParse(sCell.Trim(), out dTmp) == true)
                                {
                                    oQuoteMaterial.QuantityPer = dTmp;
                                }
                                else
                                {
                                    oErrors.Add("In the worksheet 'Materials', On row: " + iCounter.ToString() + " - could not convert the value '" + sCell + "' to a quantity per number.");
                                }
                            }
                            else if (iColumn == 6)
                            {
                                if (decimal.TryParse(sCell.Trim(), out dTmp) == true)
                                {
                                    oQuoteMaterial.UnitCost = dTmp;
                                }
                                else
                                {
                                    oErrors.Add("In the worksheet 'Materials', On row: " + iCounter.ToString() + " - could not convert the value '" + sCell + "' to a unit cost number.");
                                }
                            }
                            else if (iColumn == 7)
                            {
                                oQuoteMaterial.PartClass = sCell.Trim();
                            }
                            else if (iColumn == 8)
                            {
                                oQuoteMaterial.PartClassDescription = sCell.Trim();
                            }
                        }
                        iColumn++;
                    }
                    oQuoteMaterials.Add(oQuoteMaterial);
                    iCounter++;
                }
            }
            return oErrors;
        }

        public static List<string> ParseQuoteQuantityBreakInformation(SLExcelData oData, List<HSQuoteQuantityBreak> oQuoteQuantityBreaks)
        {
            List<string> oErrors = new List<string>();

            if (oData != null)
            {
                int iCounter = 0;
                foreach (List<string> sRowData in oData.DataRows)
                {
                    int iColumn = 0;

                    // if this is the first row then we need to first read in the quote header
                    if (iCounter == 0)
                    {
                        iColumn = 0;
                        foreach (string sCell in oData.Headers)
                        {
                            Console.WriteLine(sCell);
                            iColumn++;
                        }
                    }

                    HSQuoteQuantityBreak oQuoteQuantityBreak = new HSQuoteQuantityBreak();
                    // now we are reading in quote info
                    iColumn = 0;
                    int iTmp;
                    decimal dTmp;
                    foreach (string sCell in sRowData)
                    {
                        if (string.IsNullOrEmpty(sCell) == false)
                        {
                            // the first column will be QuoteNum, then customer Id
                            if (iColumn == 0)
                            {
                                if (int.TryParse(sCell.Trim(), out iTmp) == true)
                                {
                                    oQuoteQuantityBreak.QuoteNum = iTmp;
                                }
                                else
                                {
                                    oErrors.Add("In the worksheet 'QtyBreaks', On row: " + iCounter.ToString() + " - could not convert the value '" + sCell + "' to a quote number.");
                                }
                            }
                            else if (iColumn == 1)
                            {
                                if (int.TryParse(sCell.Trim(), out iTmp) == true)
                                {
                                    oQuoteQuantityBreak.QuoteLine = iTmp;
                                }
                                else
                                {
                                    oErrors.Add("In the worksheet 'QtyBreaks', On row: " + iCounter.ToString() + " - could not convert the value '" + sCell + "' to a line number.");
                                }
                            }
                            else if (iColumn == 2)
                            {
                                if (decimal.TryParse(sCell.Trim(), out dTmp) == true)
                                {
                                    oQuoteQuantityBreak.Quantity = dTmp;
                                }
                                else
                                {
                                    oErrors.Add("In the worksheet 'QtyBreaks', On row: " + iCounter.ToString() + " - could not convert the value '" + sCell + "' to a quantity number.");
                                }
                            }
                            else if (iColumn == 3)
                            {
                                oQuoteQuantityBreak.ProfitOrMarkup = sCell.Trim();
                            }
                            else if (iColumn == 4)
                            {
                                if (decimal.TryParse(sCell.Trim(), out dTmp) == true)
                                {
                                    oQuoteQuantityBreak.BurdenMarkup = dTmp;
                                }
                                else
                                {
                                    oErrors.Add("In the worksheet 'QtyBreaks', On row: " + iCounter.ToString() + " - could not convert the value '" + sCell + "' to a burden markup number.");
                                }
                            }
                            else if (iColumn == 5)
                            {
                                if (decimal.TryParse(sCell.Trim(), out dTmp) == true)
                                {
                                    oQuoteQuantityBreak.LaborMarkup = dTmp;
                                }
                                else
                                {
                                    oErrors.Add("In the worksheet 'QtyBreaks', On row: " + iCounter.ToString() + " - could not convert the value '" + sCell + "' to a labor markup number.");
                                }
                            }
                            else if (iColumn == 6)
                            {
                                if (decimal.TryParse(sCell.Trim(), out dTmp) == true)
                                {
                                    oQuoteQuantityBreak.MaterialMarkup = dTmp;
                                }
                                else
                                {
                                    oErrors.Add("In the worksheet 'QtyBreaks', On row: " + iCounter.ToString() + " - could not convert the value '" + sCell + "' to a material markup number.");
                                }
                            }
                            else if (iColumn == 7)
                            {
                                if (decimal.TryParse(sCell.Trim(), out dTmp) == true)
                                {
                                    oQuoteQuantityBreak.SubcontractMarkup = dTmp;
                                }
                                else
                                {
                                    oErrors.Add("In the worksheet 'QtyBreaks', On row: " + iCounter.ToString() + " - could not convert the value '" + sCell + "' to a subcontract markup number.");
                                }
                            }
                            else if (iColumn == 8)
                            {
                                if (decimal.TryParse(sCell.Trim(), out dTmp) == true)
                                {
                                    oQuoteQuantityBreak.MaterialBurdenMarkup = dTmp;
                                }
                                else
                                {
                                    oErrors.Add("In the worksheet 'QtyBreaks', On row: " + iCounter.ToString() + " - could not convert the value '" + sCell + "' to a material burden markup number.");
                                }
                            }
                            else if (iColumn == 9)
                            {
                                if (decimal.TryParse(sCell.Trim(), out dTmp) == true)
                                {
                                    oQuoteQuantityBreak.MiscellaneousCost = dTmp;
                                }
                                else
                                {
                                    oErrors.Add("In the worksheet 'QtyBreaks', On row: " + iCounter.ToString() + " - could not convert the value '" + sCell + "' to a miscellaneous cost number.");
                                }
                            }
                            else if (iColumn == 10)
                            {
                                oQuoteQuantityBreak.MiscellaneousDescription = sCell.Trim();
                            }
                            else if (iColumn == 11)
                            {
                                oQuoteQuantityBreak.MiscellaneousMarkup = decimal.Parse(sCell.Trim());
                            }
                            else if (iColumn == 12)
                            {
                                if (decimal.TryParse(sCell.Trim(), out dTmp) == true)
                                {
                                    oQuoteQuantityBreak.CommissionPercent = dTmp;
                                }
                                else
                                {
                                    oErrors.Add("In the worksheet 'QtyBreaks', On row: " + iCounter.ToString() + " - could not convert the value '" + sCell + "' to a commission percent number.");
                                }
                            }
                        }
                        iColumn++;
                    }
                    oQuoteQuantityBreaks.Add(oQuoteQuantityBreak);
                    iCounter++;
                }
            }
            return oErrors;
        }
        #endregion

        #endregion

        #region Properties

        public string Company
        {
            get { return m_sCompany; }
            set { m_sCompany = value; }
        }

        public int QuoteNum
        {
            get { return m_iQuoteNum; }
            set { m_iQuoteNum = value; }
        }

        public int QuoteLine
        {
            get { return m_iQuoteLine; }
            set { m_iQuoteLine = value; }
        }

        public int QuoteQtyNum
        {
            get { return m_iQuoteQtyNum; }
            set { m_iQuoteQtyNum = value; }
        }

        public string CustomerId
        {
            get { return m_sCustomerId; }
            set { m_sCustomerId = value; }
        }

        public string CustomerName
        {
            get { return m_sCustomerName; }
            set { m_sCustomerName = value; }
        }

        public DateTime EntryDate
        {
            get { return m_dtEntryDate; }
            set { m_dtEntryDate = value; }
        }
        public DateTime DateQuoted
        {
            get { return m_dtDateQuoted; }
            set { m_dtDateQuoted = value; }
        }
        public string QuoteComment
        {
            get { return m_sQuoteComment; }
            set { m_sQuoteComment = value; }
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
        public decimal SellingQty
        {
            get { return m_dSellingQty; }
            set { m_dSellingQty = value; }
        }

        public decimal BurdenCost
        {
            get { return m_dBurdenCost; }
            set { m_dBurdenCost = value; }
        }
        public decimal LaborCost
        {
            get { return m_dLaborCost; }
            set { m_dLaborCost = value; }
        }
        public decimal MaterialCost
        {
            get { return m_dMaterialCost; }
            set { m_dMaterialCost = value; }
        }
        public decimal SubcontractCost
        {
            get { return m_dSubcontractCost; }
            set { m_dSubcontractCost = value; }
        }
        public decimal MaterialBurdenCost
        {
            get { return m_dMaterialBurdenCost; }
            set { m_dMaterialBurdenCost = value; }
        }
        public string MiscCostDescription
        {
            get { return m_sMiscCostDescription; }
            set { m_sMiscCostDescription = value; }
        }
        public decimal MiscCost
        {
            get { return m_dMiscCost; }
            set { m_dMiscCost = value; }
        }
        public string PricePerCode
        {
            get { return m_sPricePerCode; }
            set { m_sPricePerCode = value; }
        }
        public string PercentType
        {
            get { return m_sPercentType; }
            set { m_sPercentType = value; }
        }
        public decimal BurdenMarkup
        {
            get { return m_dBurdenMarkup; }
            set { m_dBurdenMarkup = value; }
        }
        public decimal LaborMarkup
        {
            get { return m_dLaborMarkup; }
            set { m_dLaborMarkup = value; }
        }
        public decimal MaterialMarkup
        {
            get { return m_dMaterialMarkup; }
            set { m_dMaterialMarkup = value; }
        }
        public decimal SubcontractMarkup
        {
            get { return m_dSubcontractMarkup; }
            set { m_dSubcontractMarkup = value; }
        }
        public decimal MaterialBurdenMarkup
        {
            get { return m_dMaterialBurdenMarkup; }
            set { m_dMaterialBurdenMarkup = value; }
        }
        public decimal MiscCostMarkup
        {
            get { return m_dMiscCostMarkup; }
            set { m_dMiscCostMarkup = value; }
        }
        public decimal CommissionPercent
        {
            get { return m_dCommissionPercent; }
            set { m_dCommissionPercent = value; }
        }
        public decimal SellingFactor
        {
            get { return m_dSellingFactor; }
            set { m_dSellingFactor = value; }
        }
        public string SellingDirection
        {
            get { return m_sSellingDirection; }
            set { m_sSellingDirection = value; }
        }
        public string SalesUOM
        {
            get { return m_sSalesUOM; }
            set { m_sSalesUOM = value; }
        }
        public decimal PriceBurdenProfit
        {
            get { return m_dPriceBurdenProfit; }
            set { m_dPriceBurdenProfit = value; }
        }
        public decimal PriceLaborProfit
        {
            get { return m_dPriceLaborProfit; }
            set { m_dPriceLaborProfit = value; }
        }
        public decimal PriceMaterialProfit
        {
            get { return m_dPriceMaterialProfit; }
            set { m_dPriceMaterialProfit = value; }
        }
        public decimal PriceSubcontractProfit
        {
            get { return m_dPriceSubcontractProfit; }
            set { m_dPriceSubcontractProfit = value; }
        }
        public decimal PriceMaterialBurdenProfit
        {
            get { return m_dPriceMaterialBurdenProfit; }
            set { m_dPriceMaterialBurdenProfit = value; }
        }
        public decimal PriceMiscProfit
        {
            get { return m_dPriceMiscProfit; }
            set { m_dPriceMiscProfit = value; }
        }
        public decimal TotalCost
        {
            get { return m_dTotalCost; }
            set { m_dTotalCost = value; }
        }
        public decimal TotalProfit
        {
            get { return m_dTotalProfit; }
            set { m_dTotalProfit = value; }
        }
        public decimal PriceTotalMarkup
        {
            get { return m_dPriceTotalMarkup; }
            set { m_dPriceTotalMarkup = value; }
        }
        public decimal TotalCommission
        {
            get { return m_dTotalCommission; }
            set { m_dTotalCommission = value; }
        }
        public decimal PriceTotalCommissionMarkup
        {
            get { return m_dPriceTotalCommissionMarkup; }
            set { m_dPriceTotalCommissionMarkup = value; }
        }
        public decimal UnitCost
        {
            get { return m_dUnitCost; }
            set { m_dUnitCost = value; }
        }
        public decimal UnitPrice
        {
            get { return m_dUnitPrice; }
            set { m_dUnitPrice = value; }
        }
        public decimal UnitPriceWithCommission
        {
            get { return m_dUnitPriceWithCommission; }
            set { m_dUnitPriceWithCommission = value; }
        }

        public List<HSQuoteOperations> AllOperations
        {
            get { return m_oOperations; }
        }
        #endregion

        #region Data Members

        private string m_sCompany;
        private int m_iQuoteNum;
        private int m_iQuoteLine;
        private int m_iQuoteQtyNum;
        private string m_sCustomerId;
        private string m_sCustomerName;
        private DateTime m_dtEntryDate;
        private DateTime m_dtDateQuoted;
        private string m_sQuoteComment;
        private string m_sPartNum;
        private string m_sPartDescription;
        private decimal m_dSellingQty;
        private decimal m_dBurdenCost;
        private decimal m_dLaborCost;
        private decimal m_dMaterialCost;
        private decimal m_dSubcontractCost;
        private decimal m_dMaterialBurdenCost;
        private string m_sMiscCostDescription;
        private decimal m_dMiscCost;
        private string m_sPricePerCode;
        private string m_sPercentType;
        private decimal m_dBurdenMarkup;
        private decimal m_dLaborMarkup;
        private decimal m_dMaterialMarkup;
        private decimal m_dSubcontractMarkup;
        private decimal m_dMaterialBurdenMarkup;
        private decimal m_dMiscCostMarkup;
        private decimal m_dCommissionPercent;
        private decimal m_dSellingFactor;
        private string m_sSellingDirection;
        private string m_sSalesUOM;
        private decimal m_dPriceBurdenProfit;
        private decimal m_dPriceLaborProfit;
        private decimal m_dPriceMaterialProfit;
        private decimal m_dPriceSubcontractProfit;
        private decimal m_dPriceMaterialBurdenProfit;
        private decimal m_dPriceMiscProfit;
        private decimal m_dTotalCost;
        private decimal m_dTotalProfit;
        private decimal m_dPriceTotalMarkup;
        private decimal m_dTotalCommission;
        private decimal m_dPriceTotalCommissionMarkup;
        private decimal m_dUnitCost;
        private decimal m_dUnitPrice;
        private decimal m_dUnitPriceWithCommission;

        private List<HSQuoteOperations> m_oOperations = new List<HSQuoteOperations>();

        private List<SF1411Material> m_oSF1411Materials = new List<SF1411Material>();
        private List<SF1411Operation> m_oSF1411Operations = new List<SF1411Operation>();


        // set up the style of cells
        private static SLStyle g_oGood;
        private static SLStyle g_oNeutrual;
        private static SLStyle g_oBad;
        private static SLStyle g_oBold;
        private static SLStyle g_oBoldRed;
        private static SLStyle g_oBoldUnderline;
        private static SLStyle g_oBoldUnderlineRight;
        private static SLStyle g_oUnderline;
        private static SLStyle g_oBoldCenter;
        private static SLStyle g_oBoldRight;
        private static SLStyle g_oBoldCenterHeader;
        private static SLAlignment g_oCenterAlignment;
        private static SLStyle g_oCenter;
        private static SLStyle g_oUnderlineCenetered;
        private static SLStyle g_oRight;
        private static SLAlignment g_oRightAlignment;
        private static SLFill g_oSLFill;
        private static SLStyle g_oBoxStyleCentered;
        private static SLStyle g_oLineTopStyle;
        private static SLStyle g_oLineLeftStyle;
        private static SLStyle g_oLineRightStyle;
        private static SLStyle g_oLineBottomStyle;
        private static SLStyle g_oBoldCurrencyStyle;
        private static SLStyle g_oCurrencyStyleCentered;
        private static SLStyle g_oCurrencyStyleBoldCentered;
        private static SLStyle g_oDecimalBoxStyleBoldCentered;
        private static SLStyle g_oDecimalStyleCentered;
        private static SLStyle g_oTextStyleHeaderCentered;
        private static SLStyle g_oPercentStyleBoldCentered;
        private static SLStyle g_oPercentStyleCentered;
        // used to alternate row colors in spreadsheet
        private static bool g_bUsingStyle1;
        private static SLStyle g_oRowStyle1;
        private static SLStyle g_oRowStyle2;
        private static SLStyle g_oCurrencyStyle1;
        private static SLStyle g_oCurrencyStyle2;

        #endregion
    }

    public class HSQuoteHead
    {
        #region Constructors
        public HSQuoteHead()
        {
        }
        #endregion

        #region Methods
        public void LoadDetails(List<HSQuoteLine> oQuoteLines, List<HSQuoteOperation> oQuoteOperations, List<HSQuoteMaterial> oQuoteMaterials, List<HSQuoteQuantityBreak> oQuoteQuantityBreaks)
        {
            m_oLines = oQuoteLines.Where(x => x.QuoteNum == QuoteNum).ToList();
            // now we walk through each quote line and add the details to the line
            foreach(HSQuoteLine oQuoteLine in m_oLines)
            {
                oQuoteLine.LoadDetails(oQuoteOperations, oQuoteMaterials, oQuoteQuantityBreaks);
            }
        }

        #region Add Data To Epicor
        public List<string> CreateQuoteHeader(Session oSession)
        {
            List<string> oErrors = new List<string>();
            QuoteImpl oQuoteImpl = WCFServiceSupport.CreateImpl<QuoteImpl>(oSession, Erp.Proxy.BO.QuoteImpl.UriPath);
            Erp.BO.QuoteDataSet oQuoteDataSet = new QuoteDataSet();
            try
            {
                oQuoteImpl.GetNewQuoteHed(oQuoteDataSet);
                if ((oQuoteDataSet != null) && (oQuoteDataSet.QuoteHed != null))
                {
                    DataRow drQuoteHead = oQuoteDataSet.QuoteHed.Rows[0];
                    // indiate which customer info we need to load basic infor for
                    drQuoteHead["CustomerCustID"] = CustomerID;
                    // if the user has included the sales rep code then we will use it
                    if (string.IsNullOrEmpty(SalesRepCode) == false)
                    {
                        drQuoteHead["SalesRepCode"] = SalesRepCode;
                    }
                    oQuoteImpl.GetCustomerInfo(oQuoteDataSet);

                    // now we can edit any part of the quote header that needs to be updated
                    drQuoteHead.BeginEdit();

                    // indicate we are adding this quote
                    drQuoteHead["RowMod"] = "A";

                    // submit the quote header
                    drQuoteHead.EndEdit();
                    // create the quote head so we can add quote details
                    oQuoteImpl.Update(oQuoteDataSet);
                    // retain the quote number as this is needed by the quote details
                    QuoteNum = (int)drQuoteHead["QuoteNum"];
                    CustomerNum = (int)drQuoteHead["CustNum"];
                }
            }
            catch (Exception e)
            {
                oErrors.Add("Error creating Quote Head: " + QuoteNum.ToString() + ", Customer ID: " + CustomerID + ", ERROR: " + e.Message);
            }
            finally
            {
                // finally get rid of the object as we have completed adding the quote
                oQuoteImpl.Dispose();
            }

            // check to make sure there were no errors creating the quote header before we bother to create the lines
            if (oErrors.Count == 0)
            {
                // now create the lines that are part of this quote
                foreach (HSQuoteLine oLine in m_oLines)
                {
                    List<string> oLineErrors = oLine.CreateQuoteLine(oSession);
                    oErrors.AddRange(oLineErrors);
                }
            }
            return oErrors;
        }
        #endregion
        
        #endregion

        #region Properties

        public int QuoteNum
        {
            get { return m_iQuoteNum; }
            set
            {
                m_iQuoteNum = value;
                // now walk through all contained objects and set the quote number
                foreach (HSQuoteLine oLine in m_oLines)
                {
                    oLine.QuoteNum = m_iQuoteNum;
                }
            }
        }
        public string CustomerID
        {
            get { return m_sCustID; }
            set { m_sCustID = value; }
        }

        public string SalesRepCode
        {
            get { return m_sSalesRepCode; }
            set { m_sSalesRepCode = value; }
        }

        public int CustomerNum
        {
            get { return m_iCustNum; }
            set { m_iCustNum = value; }
        }

        public List<HSQuoteLine>    Lines
        {
            get { return m_oLines; }
        }
        #endregion

        #region Data Members

        private int m_iQuoteNum;
        private string m_sCustID;
        private int m_iCustNum;
        private string m_sSalesRepCode;

        private List<HSQuoteLine> m_oLines = new List<HSQuoteLine>();
        #endregion
    }

    public class HSQuoteLine
    {
        #region Constructors
        public HSQuoteLine()
        {
        }
        #endregion

        #region Methods
        public void LoadDetails(List<HSQuoteOperation> oQuoteOperations, List<HSQuoteMaterial> oQuoteMaterials, List<HSQuoteQuantityBreak> oQuoteQuantityBreaks)
        {
            m_oOperations = oQuoteOperations.Where(x => (x.QuoteNum == QuoteNum) && (x.QuoteLine == QuoteLine)).ToList();
            m_oMaterials = oQuoteMaterials.Where(x => (x.QuoteNum == QuoteNum) && (x.QuoteLine == QuoteLine)).ToList();
            m_oQuantityBreaks = oQuoteQuantityBreaks.Where(x => (x.QuoteNum == QuoteNum) && (x.QuoteLine == QuoteLine)).ToList();
        }

        public List<string> CreateQuoteLine(Session oSession)
        {
            List<string> oErrors = new List<string>();

            QuoteImpl oQuoteImpl = WCFServiceSupport.CreateImpl<QuoteImpl>(oSession, Erp.Proxy.BO.QuoteImpl.UriPath);
            try
            {
                Erp.BO.QuoteDataSet oQuoteDataSet = oQuoteImpl.GetByID(QuoteNum);
                if ((oQuoteDataSet != null) && (oQuoteDataSet.QuoteHed != null))
                {
                    // now we create a new quote line
                    oQuoteImpl.GetNewQuoteDtl(oQuoteDataSet, QuoteNum);
                    DataRow drQuoteDtl = oQuoteDataSet.QuoteDtl.Rows[oQuoteDataSet.QuoteDtl.Rows.Count - 1];
                    drQuoteDtl.BeginEdit();

                    drQuoteDtl["PartNum"] = PartNum;
                    drQuoteDtl["LineDesc"] = Description;
                    drQuoteDtl["OrderQty"] = Quantity;
                    drQuoteDtl["SellingExpectedQty"] = Quantity;
                    drQuoteDtl["DocExpUnitPrice"] = ExpectedUnitCost;
                    drQuoteDtl["ProdCode"] = ProductCode;
                    drQuoteDtl.EndEdit();
                    // create the quote line
                    oQuoteImpl.Update(oQuoteDataSet);

                    // need to make a second pass to get the price to stick
                    drQuoteDtl.BeginEdit();
                    drQuoteDtl["DocExpUnitPrice"] = ExpectedUnitCost;
                    drQuoteDtl.EndEdit();
                    // update the price
                    oQuoteImpl.Update(oQuoteDataSet);

                    // one last pass to set the discount if it applies
                    drQuoteDtl.BeginEdit();
                    drQuoteDtl["DocDspDiscount"] = 0;
                    drQuoteDtl.EndEdit();
                    // update the price
                    oQuoteImpl.Update(oQuoteDataSet);

                    // get the newly created quote line and propagate this to all materials, operations, and quantity breaks
                    QuoteLine = (int)drQuoteDtl["QuoteLine"];
                }
            }
            catch (Exception e)
            {
                oErrors.Add("Error creating Line: " + QuoteLine.ToString() + ", PartNum: " + PartNum + ", ERROR: " + e.Message);
            }
            finally
            {
                // finally get rid of the object as we have completed adding the quote
                oQuoteImpl.Dispose();
            }

            // check to make sure there were no errors creating the quote line before we proceed with operations, materials, etc.
            if (oErrors.Count == 0)
            {
                // now create any operations for this line
                foreach (HSQuoteOperation oOperation in m_oOperations)
                {
                    string sError = oOperation.CreateOperation(oSession);
                    if (string.IsNullOrEmpty(sError) == false)
                    {
                        oErrors.Add("Error creating operation, Line: " + QuoteLine.ToString() + ", Op Code: " + oOperation.OpCode + ", Op Sequence: " + oOperation.OperationSequence.ToString() + ", ERROR: " + sError);
                    }
                }

                // now create any materials for this line
                foreach (HSQuoteMaterial oMaterial in m_oMaterials)
                {
                    string sError = oMaterial.CreateMaterial(oSession);
                    if (string.IsNullOrEmpty(sError) == false)
                    {
                        oErrors.Add("Error creating material, Line: " + QuoteLine.ToString() + ", PartNum: " + oMaterial.PartNum + ", Mtl Sequence: " + oMaterial.MaterialSequence.ToString() + ", ERROR: " + sError);
                    }
                }

                // now create any quantity breaks for this line
                bool bFirstQuantityBreak = true;
                foreach (HSQuoteQuantityBreak oQuantityBreak in m_oQuantityBreaks)
                {
                    string sError = oQuantityBreak.CreateQuoteQuantity(oSession, bFirstQuantityBreak);
                    bFirstQuantityBreak = false;
                    if (string.IsNullOrEmpty(sError) == false)
                    {
                        oErrors.Add("Error creating qty break, Line: " + QuoteLine.ToString() + ", Qty: " + oQuantityBreak.Quantity.ToString() + ", ERROR: " + sError);
                    }
                }
            }

            return oErrors;
        }

        #endregion

        #region Properties
        public int QuoteNum
        {
            get { return m_iQuoteNum; }
            set
            { 
                m_iQuoteNum = value;
                // now walk through all contained objects and set the quote number
                foreach (HSQuoteOperation oOperation in m_oOperations)
                {
                    oOperation.QuoteNum = m_iQuoteNum;
                }
                foreach (HSQuoteMaterial oMaterial in m_oMaterials)
                {
                    oMaterial.QuoteNum = m_iQuoteNum;
                }
                foreach (HSQuoteQuantityBreak oQuantityBreak in m_oQuantityBreaks)
                {
                    oQuantityBreak.QuoteNum = m_iQuoteNum;
                }
            }
        }

        public int QuoteLine
        {
            get { return m_iQuoteLine; }
            set
            {
                m_iQuoteLine = value;
                // now walk through all contained objects and set the quote line
                foreach (HSQuoteOperation oOperation in m_oOperations)
                {
                    oOperation.QuoteLine = m_iQuoteLine;
                }
                foreach (HSQuoteMaterial oMaterial in m_oMaterials)
                {
                    oMaterial.QuoteLine = m_iQuoteLine;
                }
                foreach (HSQuoteQuantityBreak oQuantityBreak in m_oQuantityBreaks)
                {
                    oQuantityBreak.QuoteLine = m_iQuoteLine;
                }
            }
        }
        public string PartNum
        {
            get { return m_sPartNum; }
            set { m_sPartNum = value; }
        }

        public string Description
        {
            get { return m_sDescription; }
            set { m_sDescription = value; }
        }

        public decimal Quantity
        {
            get { return m_dQuantity; }
            set { m_dQuantity = value; }
        }

        public decimal ExpectedUnitCost
        {
            get { return m_dExpectedUnitCost; }
            set { m_dExpectedUnitCost = value; }
        }

        public string ProductCode
        {
            get { return m_sProductCode; }
            set { m_sProductCode = value; }
        }

        public List<HSQuoteOperation> Operations
        {
            get { return m_oOperations; }
        }

        public List<HSQuoteMaterial> Materials
        {
            get { return m_oMaterials; }
        }

        public List<HSQuoteQuantityBreak> QuantityBreaks
        {
            get { return m_oQuantityBreaks; }
        }
        #endregion

        #region Data Members

        private int m_iQuoteNum;
        private int m_iQuoteLine;
        private string m_sPartNum;
        private string m_sDescription;
        private decimal m_dQuantity;
        private decimal m_dExpectedUnitCost;
        private string m_sProductCode;

        private List<HSQuoteOperation> m_oOperations = new List<HSQuoteOperation>();
        private List<HSQuoteMaterial> m_oMaterials = new List<HSQuoteMaterial>();
        private List<HSQuoteQuantityBreak> m_oQuantityBreaks = new List<HSQuoteQuantityBreak>();

        #endregion
    }

    public class HSQuoteOperation
    {
        #region Constructors
        public HSQuoteOperation()
        {
        }
        #endregion

        #region Methods
        public string CreateOperation(Session oSession)
        {
            string sError = "";

            // see if we need to get the subcontract vendor num
            VendorImpl oVendorImpl = null;
            try
            {
                if (string.IsNullOrEmpty(VendorId) == false)
                {
                    // check to see if we need to look up a vendor
                    oVendorImpl = WCFServiceSupport.CreateImpl<VendorImpl>(oSession, Erp.Proxy.BO.VendorImpl.UriPath);
                    VendorDataSet oVendorData = oVendorImpl.GetByVendID(VendorId);
                    if (oVendorData.Vendor.Rows.Count == 0)
                    {
                        // we could not find the vendor so we report this as an error
                        throw new Exception("Could not find the vendor for the subcontract operation. The Subcontract Vendor Id provided was: " + VendorId);
                    }
                    // we have at least one vendor that matched so we get the vendor num
                    DataRow oVendorRow = oVendorData.Vendor.Rows[0];
                    VendorNum = (int)oVendorRow["VendorNum"];
                }
                else
                {
                    // force vendor num, days out, and costs to be zero -- means no subcontract vendor is set
                    VendorNum = 0;
                    DaysOut = 0;
                    SubcontractCost = 0;
                }
            }
            catch (Exception e)
            {
                sError = e.Message;
            }
            finally
            {
                // finally get rid of the object as we have completed adding the quote
                if (oVendorImpl != null)
                {
                    oVendorImpl.Dispose();
                }
            }

            if (string.IsNullOrEmpty(sError) == true)
            {
                // now we need to get the quote assembly object so we can add operations
                QuoteAsmImpl oQuoteAsmImpl = WCFServiceSupport.CreateImpl<QuoteAsmImpl>(oSession, Erp.Proxy.BO.QuoteAsmImpl.UriPath);
                try
                {
                    QuoteAsmDataSet oAssemblyDataSet = oQuoteAsmImpl.GetByID(QuoteNum, QuoteLine, AssemblySequence);
                    if (oAssemblyDataSet != null)
                    {
                        // now we add an operation to the quote detail line
                        oQuoteAsmImpl.GetNewOperation(oAssemblyDataSet, QuoteNum, QuoteLine, AssemblySequence, Subcontract);
                        DataRow drQuoteOperation = oAssemblyDataSet.QuoteOpr.Rows[oAssemblyDataSet.QuoteOpr.Rows.Count - 1];

                        // indicate we are now making changes to the operation data set
                        drQuoteOperation.BeginEdit();

                        string sRefreshMessage;
                        oQuoteAsmImpl.GetOprOpCodeInfo(OpCode, out sRefreshMessage, oAssemblyDataSet);

                        //int iPrimaryProductionOperationDetail = ???;
                        //oQuoteAsmImpl.CheckOperPrimaryProdOpDtl(oAssemblyDataSet, iPrimaryProductionOperationDetail);

                        drQuoteOperation["QtyPer"] = QuantityPer;
                        drQuoteOperation["ProdStandard"] = ProductionStandard;
                        drQuoteOperation["StdFormat"] = StandardFormat;
                        drQuoteOperation["StdBasis"] = StandardBasis;
                        drQuoteOperation["HoursPerMachine"] = SetupHours;
                        drQuoteOperation["PrimaryResourceGrpDesc"] = ResourceGroupDescription;
                        drQuoteOperation["PrimaryResourceGrpID"] = ResourceGroupID;
                        drQuoteOperation["SubContract"] = Subcontract;
                        drQuoteOperation["EstUnitCost"] = SubcontractCost;
                        drQuoteOperation["DaysOut"] = DaysOut;
                        drQuoteOperation["VendorNum"] = VendorNum;
                        if (string.IsNullOrEmpty(CommentText) == false)
                        {
                            drQuoteOperation["CommentText"] = CommentText;
                        }

                        //
                        // ADDITIONAL FIELDS WE MAY WANT TO SET
                        //
                        //"EstSetHours" [0]
                        //"Machines" [1]
                        //"EstScrapType" [%]
                        //"IUM" [EA]
                        //"WIHoursPerMachine" [8.00]
                        //"LaborEntryMethod"[T]
                        //"DspQtyIUM" [EA]
                        //"QuoteNumCurrencyCode" [USD]

                        // indicate we are adding the operation
                        drQuoteOperation["RowMod"] = "A";

                        // indicate we are now finished making changes to the operation data set
                        drQuoteOperation.EndEdit();

                        // save the oeprations
                        oQuoteAsmImpl.Update(oAssemblyDataSet);

                        // get the newly created operation sequence
                        OperationSequence = (int)drQuoteOperation["OprSeq"];

                        // now we need to reload this operation so that have the op detail record and we can set the crew sizes
                        oAssemblyDataSet = oQuoteAsmImpl.GetByID(QuoteNum, QuoteLine, AssemblySequence);
                        if (oAssemblyDataSet != null)
                        {
                            // now get the operation to add the crew size to
                            DataRow drQuoteOpDtl = oAssemblyDataSet.QuoteOpDtl.Rows[oAssemblyDataSet.QuoteOpDtl.Rows.Count - 1];

                            // indicate we are now making changes to the operation data set
                            drQuoteOpDtl.BeginEdit();

                            // need to set the crew sizes on the op detail table
                            drQuoteOpDtl["ProdCrewSize"] = ProdCrewSize;
                            drQuoteOpDtl["SetUpCrewSize"] = SetupCrewSize;

                            // indicate we are adding the operation
                            drQuoteOpDtl["RowMod"] = "U";

                            // indicate we are now finished making changes to the operation data set
                            drQuoteOpDtl.EndEdit();

                            // save the oeprations
                            oQuoteAsmImpl.Update(oAssemblyDataSet);
                        }
                    }
                }
                catch (Exception e)
                {
                    sError = e.Message;
                }
                finally
                {
                    // finally get rid of the object as we have completed adding the quote
                    oQuoteAsmImpl.Dispose();
                }
            }
            return sError;
        }
        public string SetCrewSize(Session oSession)
        {
            string sError = "";

            // now we need to get the quote assembly object so we can add operations
            QuoteAsmImpl oQuoteAsmImpl = WCFServiceSupport.CreateImpl<QuoteAsmImpl>(oSession, Erp.Proxy.BO.QuoteAsmImpl.UriPath);
            try
            {
                QuoteAsmDataSet oAssemblyDataSet = oQuoteAsmImpl.GetByID(QuoteNum, QuoteLine, AssemblySequence);
                if (oAssemblyDataSet != null)
                {
                    // now get the operation to add the crew size to
                    DataRow drQuoteOpDtl = oAssemblyDataSet.QuoteOpDtl.Rows[oAssemblyDataSet.QuoteOpDtl.Rows.Count - 1];

                    // indicate we are now making changes to the operation data set
                    drQuoteOpDtl.BeginEdit();

                    // need to set the crew sizes on the op detail table
                    drQuoteOpDtl["ProdCrewSize"] = ProdCrewSize;
                    drQuoteOpDtl["SetUpCrewSize"] = SetupCrewSize;

                    // indicate we are adding the operation
                    drQuoteOpDtl["RowMod"] = "U";

                    // indicate we are now finished making changes to the operation data set
                    drQuoteOpDtl.EndEdit();

                    // save the oeprations
                    oQuoteAsmImpl.Update(oAssemblyDataSet);
                }
            }
            catch (Exception e)
            {
                sError = e.Message;
            }
            finally
            {
                // finally get rid of the object as we have completed adding the quote
                oQuoteAsmImpl.Dispose();
            }

            return sError;
        }

        #endregion

        #region Properties

        public int QuoteNum
        {
            get { return m_iQuoteNum; }
            set { m_iQuoteNum = value; }
        }

        public int QuoteLine
        {
            get { return m_iQuoteLine; }
            set { m_iQuoteLine = value; }
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

        public string OpCode
        {
            get { return m_sOpCode; }
            set { m_sOpCode = value; }
        }

        public string CommentText
        {
            get { return m_sCommentText; }
            set { m_sCommentText = value; }
        }

        public bool Subcontract
        {
            get { return m_bSubcontract; }
            set { m_bSubcontract = value; }
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

        public decimal SubcontractCost
        {
            get { return m_dSubcontractCost; }
            set { m_dSubcontractCost = value; }
        }
        public decimal DaysOut
        {
            get { return m_dDaysOut; }
            set { m_dDaysOut = value; }
        }

        public decimal ProductionStandard
        {
            get { return m_dProductionStandard; }
            set { m_dProductionStandard = value; }
        }

        public decimal ProdCrewSize
        {
            get { return m_dProdCrewSize;
            }
            set { m_dProdCrewSize = value; }
        }

        public string StandardFormat
        {
            get { return m_sStandardFormat; }
            set { m_sStandardFormat = value; }
        }

        public string StandardBasis
        {
            get { return m_sStandardBasis; }
            set { m_sStandardBasis = value; }
        }

        public decimal QuantityPer
        {
            get { return m_dQuantityPer; }
            set { m_dQuantityPer = value; }
        }

        public decimal SetupHours
        {
            get { return m_dSetupHours; }
            set { m_dSetupHours = value; }
        }

        public decimal SetupCrewSize
        {
            get
            {
                return m_dSetupCrewSize;
            }
            set { m_dSetupCrewSize = value; }
        }

        public string ResourceGroupID
        {
            get { return m_sResourceGroupID; }
            set { m_sResourceGroupID = value; }
        }

        public string ResourceGroupDescription
        {
            get { return m_sResourceGroupDescription; }
            set { m_sResourceGroupDescription = value; }
        }
        #endregion

        #region Data Members

        private int m_iQuoteNum;
        private int m_iQuoteLine;
        private int m_iAssemblySequence;
        private int m_iOperationSequence;
        private string m_sOpCode;
        private string m_sCommentText;
        private bool m_bSubcontract;
        private string m_sVendorId;
        private int m_iVendorNum;
        private decimal m_dSubcontractCost;
        private decimal m_dDaysOut;
        private decimal m_dProductionStandard;
        private decimal m_dProdCrewSize;
        private string m_sStandardFormat;
        private string m_sStandardBasis;
        private decimal m_dQuantityPer;
        private decimal m_dSetupHours;
        private decimal m_dSetupCrewSize;
        private string m_sResourceGroupID;
        private string m_sResourceGroupDescription;

        #endregion
    }

    public class HSQuoteMaterial
    {
        #region Constructors
        public HSQuoteMaterial()
        {
        }
        #endregion

        #region Methods
        public string CreateMaterial(Session oSession)
        {
            string sError = "";

            string sClass = "";
            string sClassDescription = "";

            // see if the part exists
            PartImpl oPartImpl = WCFServiceSupport.CreateImpl<PartImpl>(oSession, Erp.Proxy.BO.PartImpl.UriPath);
            QuoteAsmImpl oQuoteAsmImpl = WCFServiceSupport.CreateImpl<QuoteAsmImpl>(oSession, Erp.Proxy.BO.QuoteAsmImpl.UriPath);
            try
            {
                try
                {
                    string sWhereClause = "PartNum=" + PartNum;
                    PartDataSet oPartDataSet = oPartImpl.GetByID(PartNum);
                    if ((oPartDataSet.Part.Rows != null) && (oPartDataSet.Part.Rows.Count == 1))
                    {
                        DataRow oRow = oPartDataSet.Part.Rows[0];
                        sClass = (string)oRow["ClassID"];
                        sClassDescription = (string)oRow["ClassDescription"];
                    }
                }
                catch (Exception)
                {
                    // we will ignore this exception
                    // it means we could not find this part
                }

                // now we need to get the quote assembly object so we can add operations and materials
                // load the assembly that was created when we created the quote head and line
                QuoteAsmDataSet oAssemblyDataSet = oQuoteAsmImpl.GetByID(QuoteNum, QuoteLine, AssemblySequence);
                if (oAssemblyDataSet != null)
                {
                    // now we will add a new quote material
                    oQuoteAsmImpl.GetNewQuoteMtl(oAssemblyDataSet, QuoteNum, QuoteLine, AssemblySequence);
                    DataRow drQuoteMaterial = oAssemblyDataSet.QuoteMtl.Rows[oAssemblyDataSet.QuoteMtl.Rows.Count - 1];

                    // indicate we are now making changes to the material data set
                    drQuoteMaterial.BeginEdit();

                    // we need to add the part material to the assembly
                    string vMessage;
                    bool vSubAvail;
                    string vMsgType;
                    string productConfiguratorMessage;
                    oQuoteAsmImpl.CheckPrePartInfo(ref m_sPartNum, "QuoteMtl", Guid.Empty, out vMessage, out vSubAvail, out vMsgType, out productConfiguratorMessage);

                    oQuoteAsmImpl.CheckQuoteMtlPartNum(oAssemblyDataSet, PartNum);

                    // the part description has now been filled in so we can extract this to use in the next step
                    // MIGHT NEED TO GET OTHER DATA AS WELL
                    string sDescription = m_sPartNum;
                    if (drQuoteMaterial["Description"] != null)
                    {
                        string sTmp = (string)drQuoteMaterial["Description"];
                        if (string.IsNullOrEmpty(sTmp) == false)
                        {
                            sDescription = sTmp;
                        }
                    }

                    // we need to set the part number, description, etc for this next call
                    drQuoteMaterial.BeginEdit();

                    drQuoteMaterial["PartNum"] = PartNum;
                    drQuoteMaterial["Description"] = sDescription;
                    //drQuoteMaterial["IUM"] = "EA";
                    //drQuoteMaterial["ScrapUOM"] = "%";
                    //drQuoteMaterial["BaseUOM"] = "EA";

                    oQuoteAsmImpl.GetMtlPartInfo(oAssemblyDataSet, PartNum);

                    // change the required quantities for this material
                    drQuoteMaterial["QtyPer"] = QuantityPer;
                    oQuoteAsmImpl.ChangeOpMtlReqQty(oAssemblyDataSet);

                    // change the estiamted unit cost for this material
                    oQuoteAsmImpl.ChangeQuoteMtlEstUnitCost(UnitCost, oAssemblyDataSet);

                    // change the material burden for this material
                    //oQuoteAsmImpl.ChangeQuoteMtlMtlBurRate(dMaterialBurdenRate, oAssemblyDataSet);

                    //
                    // ADDITONAL FIELDS WE MAY WANT TO SET
                    //
                    //IUM                       -> "EA"
                    //RelatedOperation          -> 10
                    //MfgComment                -> ""
                    //MinimumCost               -> 0
                    //EstUnitCost               -> 34.78
                    //FixedQty                  -> False
                    //BasePartNum               -> ""
                    //EstMtlUnitCost            -> 34.78
                    //MiscCharge                -> False
                    //BaseUOM                   -> "EA"
                    //EnableFixedQty            -> True
                    //PartNumSalesUM            -> "EA"
                    //PartNumIUM                -> "EA"
                    //PartNumPricePerCode       -> "E"
                    //QuoteNumCurrencyCode      -> "USD"

                    drQuoteMaterial["PartNum"] = PartNum;
                    drQuoteMaterial["Description"] = sDescription;
                    //drQuoteMaterial["IUM"] = "EA";
                    //drQuoteMaterial["ScrapUOM"] = "%";
                    //drQuoteMaterial["BaseUOM"] = "EA";
                    drQuoteMaterial["Class"] = sClass;
                    drQuoteMaterial["ClassDescription"] = sClassDescription;
                    // indicate we are adding this material
                    drQuoteMaterial["RowMod"] = "A";

                    // indicate we are now finished making changes to the material data set
                    drQuoteMaterial.EndEdit();

                    // save the materials to the quote
                    oQuoteAsmImpl.Update(oAssemblyDataSet);

                    MaterialSequence = (int)drQuoteMaterial["MtlSeq"];
                }
            }
            catch (Exception e)
            {
                sError = e.Message;
            }
            finally
            {
                // finally get rid of the object as we have completed adding the quote
                oPartImpl.Dispose();
                oQuoteAsmImpl.Dispose();
            }

            return sError;
        }

        #endregion

        #region Properties

        public int QuoteNum
        {
            get { return m_iQuoteNum; }
            set { m_iQuoteNum = value; }
        }

        public int QuoteLine
        {
            get { return m_iQuoteLine; }
            set { m_iQuoteLine = value; }
        }

        public int AssemblySequence
        {
            get { return m_iAssemblySequence; }
            set { m_iAssemblySequence = value; }
        }

        public int MaterialSequence
        {
            get { return m_iMaterialSequence; }
            set { m_iMaterialSequence = value; }
        }

        public string PartNum
        {
            get { return m_sPartNum; }
            set { m_sPartNum = value; }
        }

        public decimal QuantityPer
        {
            get { return m_dQuantityPer; }
            set { m_dQuantityPer = value; }
        }

        public decimal UnitCost
        {
            get { return m_dUnitCost; }
            set { m_dUnitCost = value; }
        }

        public string PartClass
        {
            get { return m_sPartClass; }
            set { m_sPartClass = value; }
        }

        public string PartClassDescription
        {
            get { return m_sPartClassDescription; }
            set { m_sPartClassDescription = value; }
        }
        #endregion

        #region Data Members

        private int m_iQuoteNum;
        private int m_iQuoteLine;
        private int m_iAssemblySequence;
        private int m_iMaterialSequence;
        private string m_sPartNum;
        private string m_sDescription;
        private decimal m_dQuantityPer;
        private decimal m_dUnitCost;
        private string m_sPartClass;
        private string m_sPartClassDescription;

        #endregion
    }

    public class HSQuoteQuantityBreak
    {
        #region Constructors
        public HSQuoteQuantityBreak()
        {
        }
        #endregion

        #region Methods
        public string CreateQuoteQuantity(Session oSession, bool bFirstQuantityBreak)
        {
            string sError = "";

            QuoteImpl oQuoteImpl = WCFServiceSupport.CreateImpl<QuoteImpl>(oSession, Erp.Proxy.BO.QuoteImpl.UriPath);
            QuoteAsmImpl oQuoteAsmImpl = WCFServiceSupport.CreateImpl<QuoteAsmImpl>(oSession, Erp.Proxy.BO.QuoteAsmImpl.UriPath);
            QuoteDataSet oQuoteDataSet = null;
            try
            {
                string sCurrencyBase;
                oQuoteImpl.GetCurrencyBase(out sCurrencyBase);

                bool bEnableSupplierPriceList;
                oQuoteAsmImpl.EnableSupplierPriceList(out bEnableSupplierPriceList);

                oQuoteImpl.prjWBSPhaseDefinitionIsAllowed();

                oQuoteImpl.GetExternalCRMIntegrationIsEnabled();

                oQuoteImpl.LaunchGlobalAlerts();

                string sResult = oQuoteImpl.GetCodeDescList("QuoteDtl", "ProcessMode");

                bool bUseThirdPartyScheduling;
                bool bEnableManifestRateShopping;
                string sManifestRateShoppingURL;
                oQuoteImpl.GetPlantConfCtrlValues("51504", "MfgSys", out bUseThirdPartyScheduling, out bEnableManifestRateShopping, out sManifestRateShoppingURL);

                oQuoteImpl.CheckQuoteSecurity(QuoteNum);

                oQuoteDataSet = oQuoteImpl.GetByID(QuoteNum);
                DataRow drQuoteQuantity = null;
                if (oQuoteDataSet != null)
                {
                    // if this is the first qty break it already exists and we just need to retrieve it
                    // if this is not the first qty break then we need to create it
                    if (bFirstQuantityBreak != true)
                    {
                        // we need to create this qty break
                        oQuoteImpl.GetNewQuoteQty(oQuoteDataSet, QuoteNum, QuoteLine);
                        drQuoteQuantity = oQuoteDataSet.QuoteQty.Rows[oQuoteDataSet.QuoteQty.Rows.Count - 1];

                        bool bHasPriceBreak;
                        oQuoteImpl.GetQtyPriceInfo(oQuoteDataSet, out bHasPriceBreak);
                    }
                    else
                    {
                        drQuoteQuantity = oQuoteDataSet.QuoteQty.Rows[oQuoteDataSet.QuoteQty.Rows.Count - 1];
                    }

                    // we set QuoteQty data set values
                    drQuoteQuantity["OurQuantity"] = Quantity;
                    drQuoteQuantity["SellingQuantity"] = Quantity;

                    if (string.Compare(ProfitOrMarkup, "P", true) == 0)
                    {
                        drQuoteQuantity["PercentType"] = "P";
                    }
                    else
                    {
                        drQuoteQuantity["PercentType"] = "M";
                    }

                    drQuoteQuantity["BurdenMarkUp"] = BurdenMarkup;
                    drQuoteQuantity["LaborMarkUp"] = LaborMarkup;
                    drQuoteQuantity["MaterialMarkupP"] = MaterialMarkup;
                    drQuoteQuantity["MaterialMarkupM"] = MaterialMarkup;
                    drQuoteQuantity["SubcontractMarkUp"] = SubcontractMarkup;
                    drQuoteQuantity["MtlBurMarkUp"] = MaterialBurdenMarkup;
                    if (string.IsNullOrEmpty(MiscellaneousDescription) == false)
                    {
                        drQuoteQuantity["MiscCostDesc"] = MiscellaneousDescription;
                    }
                    drQuoteQuantity["MiscCost"] = MiscellaneousCost;
                    drQuoteQuantity["MiscCostMarkUp"] = MiscellaneousMarkup;
                    drQuoteQuantity["CommissionPct"] = CommissionPercent;

                    oQuoteImpl.RecalcWorksheet(QuoteNum, QuoteLine, QuantityNum, oQuoteDataSet);

                    oQuoteImpl.CalcMaterialMarkup(QuoteNum, QuoteLine, QuantityNum, oQuoteDataSet);

                    oQuoteImpl.ValidateProfits(QuoteNum, QuoteLine, QuantityNum, oQuoteDataSet);

                    if (bFirstQuantityBreak != true)
                    {
                        drQuoteQuantity["RowMod"] = "A";
                    }
                    else
                    {
                        drQuoteQuantity["RowMod"] = "U";
                    }
                    oQuoteImpl.Update(oQuoteDataSet);

                    QuantityNum = (int)drQuoteQuantity["QtyNum"];
                }
            }
            catch (Exception e)
            {
                sError = e.Message;
            }
            finally
            {
                oQuoteImpl.Dispose();
                oQuoteAsmImpl.Dispose();
            }

            // begin set the quoted unit price
            if (string.IsNullOrEmpty(sError) == true)
            {
                oQuoteImpl = WCFServiceSupport.CreateImpl<QuoteImpl>(oSession, Erp.Proxy.BO.QuoteImpl.UriPath);
                oQuoteAsmImpl = WCFServiceSupport.CreateImpl<QuoteAsmImpl>(oSession, Erp.Proxy.BO.QuoteAsmImpl.UriPath);
                oQuoteDataSet = null;
                try
                {
                    string sCurrencyBase;
                    oQuoteImpl.GetCurrencyBase(out sCurrencyBase);

                    bool bEnableSupplierPriceList;
                    oQuoteAsmImpl.EnableSupplierPriceList(out bEnableSupplierPriceList);

                    oQuoteImpl.prjWBSPhaseDefinitionIsAllowed();

                    oQuoteImpl.GetExternalCRMIntegrationIsEnabled();

                    oQuoteImpl.LaunchGlobalAlerts();

                    string sResult = oQuoteImpl.GetCodeDescList("QuoteDtl", "ProcessMode");

                    bool bUseThirdPartyScheduling;
                    bool bEnableManifestRateShopping;
                    string sManifestRateShoppingURL;
                    oQuoteImpl.GetPlantConfCtrlValues("51504", "MfgSys", out bUseThirdPartyScheduling, out bEnableManifestRateShopping, out sManifestRateShoppingURL);

                    oQuoteImpl.CheckQuoteSecurity(QuoteNum);

                    oQuoteDataSet = oQuoteImpl.GetByID(QuoteNum);
                    DataRow drQuoteQuantity = null;
                    if (oQuoteDataSet != null)
                    {
                        foreach (DataRow dr in oQuoteDataSet.QuoteQty.Rows)
                        {
                            if (QuantityNum == (int)dr["QtyNum"])
                            {
                                drQuoteQuantity = dr;
                                break;
                            }
                        }

                        if (drQuoteQuantity != null)
                        {
                            // we found the quote quantity now we need to set the unit price
                            oQuoteImpl.RecalcWorksheet(QuoteNum, QuoteLine, QuantityNum, oQuoteDataSet);

                            oQuoteImpl.CalcMaterialMarkup(QuoteNum, QuoteLine, QuantityNum, oQuoteDataSet);

                            oQuoteImpl.ValidateProfits(QuoteNum, QuoteLine, QuantityNum, oQuoteDataSet);

                            decimal dQuotedUnitPrice = (decimal)drQuoteQuantity["CalcUPCommMarkup"];
                            decimal dTotalCost = (decimal)drQuoteQuantity["TotalCost"];

                            drQuoteQuantity.BeginEdit();
                            //CalcUPCommMarkup
                            drQuoteQuantity["WQUnitPrice"] = dQuotedUnitPrice;

                            oQuoteImpl.GetWSUnitPrice(oQuoteDataSet);

                            // QuotedMarkup =  (Price - UnitCost) / UnitCost * 100.0
                            decimal dQuotedMarkup = (dQuotedUnitPrice - dTotalCost) / dTotalCost * 100.0M;
                            // QuotedProfit = (Price - UnitCost) / Price * 100.0
                            decimal dQuotedProfit = (dQuotedUnitPrice - dTotalCost) / dQuotedUnitPrice * 100.0M;

                            drQuoteQuantity["UnitPrice"] = dQuotedUnitPrice;
                            drQuoteQuantity["DocUnitPrice"] = dQuotedUnitPrice;
                            drQuoteQuantity["CurrencySwitch"] = false;
                            drQuoteQuantity["QuotedMarkup"] = dQuotedMarkup;
                            drQuoteQuantity["QuotedProfit"] = dQuotedProfit;
                            drQuoteQuantity["TotalQuotedPrice"] = dQuotedUnitPrice;
                            drQuoteQuantity["WQUnitPrice"] = dQuotedUnitPrice;
                            drQuoteQuantity["RowMod"] = "U";

                            DataRow drTaxConnect = oQuoteDataSet.TaxConnectStatus.Rows[oQuoteDataSet.TaxConnectStatus.Rows.Count - 1];
                            drTaxConnect.BeginEdit();
                            drTaxConnect["ETCOffline"] = true;
                            drTaxConnect["RowMod"] = "U";
                            drTaxConnect.EndEdit();

                            drQuoteQuantity.EndEdit();
                            oQuoteImpl.Update(oQuoteDataSet);
                        }
                    }
                }
                catch (Exception e)
                {
                    sError = e.Message;
                }
                finally
                {
                    oQuoteImpl.Dispose();
                    oQuoteAsmImpl.Dispose();
                }
            }
            // end set the quoted unit price

            return sError;
        }

        #endregion

        #region Properties

        public int QuoteNum
        {
            get { return m_iQuoteNum; }
            set { m_iQuoteNum = value; }
        }

        public int QuoteLine
        {
            get { return m_iQuoteLine; }
            set { m_iQuoteLine = value; }
        }

        public int QuantityNum
        {
            get { return m_iQuantityNum; }
            set { m_iQuantityNum = value; }
        }

        public decimal Quantity
        {
            get { return m_dQuantity; }
            set { m_dQuantity = value; }
        }

        public string ProfitOrMarkup
        {
            get { return m_sProfitOrMarkup; }
            set { m_sProfitOrMarkup = value; }
        }

        public decimal BurdenMarkup
        {
            get { return m_dBurdenMarkup; }
            set { m_dBurdenMarkup = value; }
        }

        public decimal LaborMarkup
        {
            get { return m_dLaborMarkup; }
            set { m_dLaborMarkup = value; }
        }

        public decimal MaterialMarkup
        {
            get { return m_dMaterialMarkup; }
            set { m_dMaterialMarkup = value; }
        }

        public decimal SubcontractMarkup
        {
            get { return m_dSubcontractMarkup; }
            set { m_dSubcontractMarkup = value; }
        }

        public decimal MaterialBurdenMarkup
        {
            get { return m_dMaterialBurdenMarkup; }
            set { m_dMaterialBurdenMarkup = value; }
        }

        public decimal MiscellaneousCost
        {
            get { return m_dMiscCost; }
            set { m_dMiscCost = value; }
        }

        public string MiscellaneousDescription
        {
            get { return m_sMiscDescription; }
            set { m_sMiscDescription = value; }
        }

        public decimal MiscellaneousMarkup
        {
            get { return m_dMiscMarkup; }
            set { m_dMiscMarkup = value; }
        }

        public decimal CommissionPercent
        {
            get { return m_dCommissionPercent; }
            set { m_dCommissionPercent = value; }
        }

        #endregion

        #region Data Members

        private int m_iQuoteNum;
        private int m_iQuoteLine;
        private int m_iQuantityNum;
        private decimal m_dQuantity;
        private string m_sProfitOrMarkup;
        private decimal m_dBurdenMarkup;
        private decimal m_dLaborMarkup;
        private decimal m_dMaterialMarkup;
        private decimal m_dSubcontractMarkup;
        private decimal m_dMaterialBurdenMarkup;
        private decimal m_dMiscCost;
        private string m_sMiscDescription;
        private decimal m_dMiscMarkup;
        private decimal m_dCommissionPercent;

        #endregion
    }

}
