using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace HorizonScientific
{
    public class HSInventoryAdjustment
    {
        #region Constructors

        public HSInventoryAdjustment(HSInventoryAdjustment oOriginal)
        {
            this.m_dAddedToStock = oOriginal.m_dAddedToStock;
            this.m_dTakenFromStock = oOriginal.m_dTakenFromStock;
            this.m_dtTranDate = oOriginal.m_dtTranDate;
            this.m_sTranType = oOriginal.m_sTranType;
            this.m_sCompany = oOriginal.m_sCompany;
            // we always will force the part nums to be in upper case for comparisons
            if (string.IsNullOrEmpty(oOriginal.m_sPartNum) == false)
            {
                this.m_sPartNum = oOriginal.m_sPartNum.ToUpper();
            }
            this.m_dRunningTotal = oOriginal.m_dRunningTotal;
            this.m_dStdCost = oOriginal.m_dStdCost;
        }

        public HSInventoryAdjustment(DataRow oDataRow)
        {
            //PartTran_Company
            if ((oDataRow["PartTran_Company"] != DBNull.Value) && (string.IsNullOrEmpty((string)oDataRow["PartTran_Company"]) == false))
            {
                m_sCompany = (string)oDataRow["PartTran_Company"];
            }
            if ((oDataRow["PartTran_PartNum"] != DBNull.Value) && (string.IsNullOrEmpty((string)oDataRow["PartTran_PartNum"]) == false))
            {
                m_sPartNum = (string)oDataRow["PartTran_PartNum"];
                // we always will force the part nums to be in upper case for comparisons
                if (string.IsNullOrEmpty(m_sPartNum) == false)
                {
                    m_sPartNum = m_sPartNum.ToUpper();
                }
            }
            if (oDataRow["PartTran_TranDate"] != DBNull.Value)
            {
                m_dtTranDate = (DateTime)oDataRow["PartTran_TranDate"];
            }
            if ((oDataRow["PartTran_TranType"] != DBNull.Value) && (string.IsNullOrEmpty((string)oDataRow["PartTran_TranType"]) == false))
            {
                m_sTranType = (string)oDataRow["PartTran_TranType"];
            }
            m_dAddedToStock = (decimal)oDataRow["Calculated_AddedToStock"];
            m_dTakenFromStock = (decimal)oDataRow["Calculated_TakenFromStock"];
            m_dRunningTotal = 0;
            m_dStdCost = 0;
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

        public string PartNum
        {
            get { return m_sPartNum; }
            set { m_sPartNum = value; }
        }

        public DateTime TranDate
        {
            get { return m_dtTranDate; }
            set { m_dtTranDate = value; }
        }

        public string TranType
        {
            get { return m_sTranType; }
            set { m_sTranType = value; }
        }

        public decimal AddedToStock
        {
            get { return m_dAddedToStock; }
            set { m_dAddedToStock = value; }
        }

        public decimal TakenFromStock
        {
            get { return m_dTakenFromStock; }
            set { m_dTakenFromStock = value; }
        }

        public decimal RunningTotal
        {
            get { return m_dRunningTotal; }
            set { m_dRunningTotal = value; }
        }

        public decimal StdCost
        {
            get { return m_dStdCost; }
            set { m_dStdCost = value; }
        }

        public decimal TotalValue
        {
            get { return RunningTotal * StdCost; }
        }

        #endregion

        #region Data Members

        private string m_sCompany;
        private string m_sPartNum;
        private DateTime m_dtTranDate = DateTime.MinValue;
        private string m_sTranType;
        private decimal m_dAddedToStock;
        private decimal m_dTakenFromStock;
        private decimal m_dRunningTotal;
        private decimal m_dStdCost;

        #endregion
    }
}
