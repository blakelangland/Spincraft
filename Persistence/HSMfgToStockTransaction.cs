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
    public class HSMfgToStockTransaction
    {
        #region Constructors

        public HSMfgToStockTransaction(DataRow oDataRow)
        {
            if (oDataRow["PartTran_Company"] != DBNull.Value)
            {
                m_sCompany = (string)oDataRow["PartTran_Company"];
            }
            if (oDataRow["PartTran_PartNum"] != DBNull.Value)
            {
                m_sPartNum = (string)oDataRow["PartTran_PartNum"];
            }
            if (oDataRow["PartTran_TranType"] != DBNull.Value)
            {
                m_sTranType = (string)oDataRow["PartTran_TranType"];
            }
            if (oDataRow["PartTran_TranQty"] != DBNull.Value)
            {
                m_dTranQty = (decimal)oDataRow["PartTran_TranQty"];
            }
            if (oDataRow["PartTran_TranDate"] != DBNull.Value)
            {
                m_dtTranDate = (DateTime)oDataRow["PartTran_TranDate"];
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

        public string TranType
        {
            get { return m_sTranType; }
        }

        public decimal TranQty
        {
            get { return m_dTranQty; }
        }

        public DateTime TranDate
        {
            get { return m_dtTranDate; }
        }

        #endregion

        #region Data Members

        private string m_sCompany;
        private string m_sPartNum;
        private string m_sTranType;
        private decimal m_dTranQty;
        private DateTime m_dtTranDate;

        #endregion
    }

}
