using HorizonScientific;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace HorizonScientific
{
    public class SalesOrderGroup
    {
        public SalesOrderGroup(List<SOBacklogForPartAnalysis> oOrdersShippingTogether)
        {
            m_oOrdersThatMustShipTogether = oOrdersShippingTogether;
        }

        public List<SOBacklogForPartAnalysis> OrdersThatMustShipTogether
        {
            get { return m_oOrdersThatMustShipTogether; }
            set { m_oOrdersThatMustShipTogether = value; }
        }

        public decimal TotalValue
        {
            get
            {
                decimal dTotalValue = 0M;
                foreach (SOBacklogForPartAnalysis oOrder in m_oOrdersThatMustShipTogether)
                {
                    dTotalValue += oOrder.ExtPrice;
                }
                return dTotalValue;
            }
        }

        public DateTime EffectiveShipDate
        {
            // this will return the latest effective ship date for the group
            get
            {
                DateTime dtEffectiveShipDate = DateTime.Now;
                foreach (SOBacklogForPartAnalysis oOrder in m_oOrdersThatMustShipTogether)
                {
                    if (oOrder.EffectiveShipDate > dtEffectiveShipDate)
                    {
                        dtEffectiveShipDate = oOrder.EffectiveShipDate;
                    }
                }
                return dtEffectiveShipDate;
            }
        }

        public bool PartShortage
        {
            get
            {
                bool bPartShortage = false;
                foreach (SOBacklogForPartAnalysis oSalesOrder in m_oOrdersThatMustShipTogether)
                {
                    bPartShortage = oSalesOrder.PartShortage;
                    if (bPartShortage == true)
                    {
                        break;
                    }
                }
                return bPartShortage;
            }
        }

        private List<SOBacklogForPartAnalysis> m_oOrdersThatMustShipTogether = new List<SOBacklogForPartAnalysis>();
    }
}
