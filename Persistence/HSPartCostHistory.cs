using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace HorizonScientific
{
    public class HSPartCostHistory
    {
        #region Constructors

        public HSPartCostHistory(HSPartCostHistory oOriginal)
        {
            this.m_dBurdenUnitCost = oOriginal.m_dBurdenUnitCost;
            this.m_dLbrUnitCost = oOriginal.m_dLbrUnitCost;
            this.m_dMtlBurUnitCost = oOriginal.m_dMtlBurUnitCost;
            this.m_dMtlUnitCost = oOriginal.m_dMtlUnitCost;
            this.m_dStdCost = oOriginal.m_dStdCost;
            this.m_dSubUnitCost = oOriginal.m_dSubUnitCost;
            this.m_dtTranDate = oOriginal.m_dtTranDate;
            this.m_sCompany = oOriginal.m_sCompany;
            // we always will force the part num to be in upper case for comparisons
            if (string.IsNullOrEmpty(oOriginal.m_sPartNum) == false)
            {
                this.m_sPartNum = oOriginal.m_sPartNum.ToUpper();
            }
        }

        public HSPartCostHistory(DataRow oDataRow)
        {
            //PartTran_Company
            if ((oDataRow["PartTran_Company"] != DBNull.Value) && (string.IsNullOrEmpty((string)oDataRow["PartTran_Company"]) == false))
            {
                m_sCompany = (string)oDataRow["PartTran_Company"];
            }
            if ((oDataRow["PartTran_PartNum"] != DBNull.Value) && (string.IsNullOrEmpty((string)oDataRow["PartTran_PartNum"]) == false))
            {
                m_sPartNum = (string)oDataRow["PartTran_PartNum"];
                // we always will force the part num to be in upper case for comparisons
                if (string.IsNullOrEmpty(m_sPartNum) == false)
                {
                    m_sPartNum = m_sPartNum.ToUpper();
                }
            }
            if (oDataRow["PartTran_TranDate"] != DBNull.Value)
            {
                m_dtTranDate = (DateTime)oDataRow["PartTran_TranDate"];
            }
            m_dMtlUnitCost = (decimal)oDataRow["CostAdjTran_MtlUnitCost"];
            m_dLbrUnitCost = (decimal)oDataRow["CostAdjTran_LbrUnitCost"];
            m_dBurdenUnitCost = (decimal)oDataRow["CostAdjTran_BurUnitCost"];
            m_dSubUnitCost = (decimal)oDataRow["CostAdjTran_SubUnitCost"];
            m_dMtlBurUnitCost = (decimal)oDataRow["CostAdjTran_MtlBurUnitCost"];
            m_dStdCost = (decimal)oDataRow["Calculated_NewStdCost"];
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

        public decimal MtlUnitCost
        {
            get { return m_dMtlUnitCost; }
            set { m_dMtlUnitCost = value; }
        }

        public decimal LbrUnitCost
        {
            get { return m_dLbrUnitCost; }
            set { m_dLbrUnitCost = value; }
        }

        public decimal BurdenUnitCost
        {
            get { return m_dBurdenUnitCost; }
            set { m_dBurdenUnitCost = value; }
        }

        public decimal SubUnitCost
        {
            get { return m_dSubUnitCost; }
            set { m_dSubUnitCost = value; }
        }

        public decimal MtlBurUnitCost
        {
            get { return m_dMtlBurUnitCost; }
            set { m_dMtlBurUnitCost = value; }
        }

        public decimal StdCost
        {
            get { return m_dStdCost; }
            set { m_dStdCost = value; }
        }

        #endregion

        #region Data Members

        private string m_sCompany;
        private string m_sPartNum;
        private DateTime m_dtTranDate = DateTime.MinValue;
        private decimal m_dMtlUnitCost;
        private decimal m_dLbrUnitCost;
        private decimal m_dBurdenUnitCost;
        private decimal m_dSubUnitCost;
        private decimal m_dMtlBurUnitCost;
        private decimal m_dStdCost;

        #endregion
    }

}
