using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;

using Ice.Core;
using Erp.BO;
using Erp.Proxy.BO;
using Ice.Lib.Framework;
using Ice.Adapters;
using Ice.BO;
using HorizonScientific;
using System.IO;
using SpreadsheetLight;
using SpreadsheetLight.Charts;
using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Spreadsheet;
using Erp.Tablesets;
using Ice.Lib.Customization.Dialogs;
using Erp.Common.ContractInterfaces;


namespace HorizonScientific
{
    public class HSReceiptInTime
    {
        #region Constructors

        public HSReceiptInTime(DataRow oDataRow)
        {
            if (oDataRow["RcvDtl_Company"] != DBNull.Value)
            {
                m_sCompany = (string)oDataRow["RcvDtl_Company"];
            }
            if (oDataRow["RcvDtl_PartNum"] != DBNull.Value)
            {
                m_sPartNum = (string)oDataRow["RcvDtl_PartNum"];
            }
            if (oDataRow["Calculated_Description"] != DBNull.Value)
            {
                m_sDescription = (string)oDataRow["Calculated_Description"];
            }
            if (oDataRow["RcvDtl_PUM"] != DBNull.Value)
            {
                m_sPUM = (string)oDataRow["RcvDtl_PUM"];
            }
            if (oDataRow["RcvDtl_OurQty"] != DBNull.Value)
            {
                m_dOurQty = (decimal)oDataRow["RcvDtl_OurQty"];
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
            if (oDataRow["RcvDtl_PackSlip"] != DBNull.Value)
            {
                m_sPackSlip = (string)oDataRow["RcvDtl_PackSlip"];
            }
            if (oDataRow["RcvDtl_PackLine"] != DBNull.Value)
            {
                m_iPackLine = (int)oDataRow["RcvDtl_PackLine"];
            }
            if (oDataRow["Part_TrackLots"] != DBNull.Value)
            {
                m_bTrackLots = (bool)oDataRow["Part_TrackLots"];
            }
            if (oDataRow["RcvDtl_LotNum"] != DBNull.Value)
            {
                m_sLotNum = (string)oDataRow["RcvDtl_LotNum"];
            }
            if (oDataRow["RcvDtl_ReceiptType"] != DBNull.Value)
            {
                m_sReceiptType = (string)oDataRow["RcvDtl_ReceiptType"];
            }
            if (oDataRow["RcvDtl_ReceiptDate"] != DBNull.Value)
            {
                m_dtReceiptDate = (DateTime)oDataRow["RcvDtl_ReceiptDate"];
            }
        }

        #endregion

        #region Methods
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

        public string Description
        {
            get { return m_sDescription; }
        }

        public string PUM
        {
            get { return m_sPUM; }
        }

        public decimal OurQty
        {
            get { return m_dOurQty; }
        }
        public int PONum
        {
            get { return m_iPONum; }
        }
        public int POLine
        {
            get { return m_iPOLine; }
        }
        public int PORelNum
        {
            get { return m_iPORelNum; }
        }
        public string PackSlip
        {
            get { return m_sPackSlip; }
        }
        public int PackLine
        {
            get { return m_iPackLine; }
        }
        public bool TrackLots
        {
            get { return m_bTrackLots; }
        }
        public string LotNum
        {
            get { return m_sLotNum; }
        }

        public string ReceiptType
        {
            get { return m_sReceiptType; }
        }

        public DateTime ReceiptDate
        {
            get { return m_dtReceiptDate; }
        }

        #endregion

        #region Data Members

        private string m_sCompany;
        private string m_sPartNum;
        private string m_sDescription;
        private string m_sPUM;
        private decimal m_dOurQty;
        private int m_iPONum;
        private int m_iPOLine;
        private int m_iPORelNum;
        private string m_sPackSlip;
        private int m_iPackLine;
        private bool m_bTrackLots;
        private string m_sLotNum;
        private string m_sReceiptType;
        private DateTime m_dtReceiptDate;

        #endregion
    }
}
