using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace HorizonScientific
{
    public class HSPartsOnHand
    {
        //PartsOnHand

        #region Constructors

        public HSPartsOnHand(DataRow oRow)
        {
            if ((oRow["PartBin_PartNum"] != DBNull.Value) && (string.IsNullOrEmpty((string)oRow["PartBin_PartNum"]) == false))
            {
                m_sPartNum = (string)oRow["PartBin_PartNum"];
            }
            if (oRow["Calculated_TotalOnHand"] != DBNull.Value)
            {
                m_dTotalOnHand = (decimal)oRow["Calculated_TotalOnHand"];
            }
        }

        #endregion

        #region Properties

        public string PartNum
        {
            get { return m_sPartNum; }
        }
        public decimal TotalOnHand
        {
            get { return m_dTotalOnHand; }
        }

        #endregion

        #region Data Members

        private string m_sPartNum;
        private decimal m_dTotalOnHand;

        #endregion
    }
}
