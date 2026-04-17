using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace HorizonScientific
{
    public class HSDemandForPartsInTime
    {
        //DemandForPartsInTime

        #region Constructors

        public HSDemandForPartsInTime(DataRow oRow)
        {
            if ((oRow["PartDtl_PartNum"] != DBNull.Value) && (string.IsNullOrEmpty((string)oRow["PartDtl_PartNum"]) == false))
            {
                m_sPartNum = (string)oRow["PartDtl_PartNum"];
            }
            if (oRow["PartDtl_DueDate"] != DBNull.Value)
            {
                m_dtDueDate = (DateTime)oRow["PartDtl_DueDate"];
                m_bDueDateNotSet = false;
            }
            else
            {
                m_dtDueDate = DateTime.Now;
                m_bDueDateNotSet = true;
            }
            if ((oRow["Calculated_SourceFile"] != DBNull.Value) && (string.IsNullOrEmpty((string)oRow["Calculated_SourceFile"]) == false))
            {
                m_sSourceFile = (string)oRow["Calculated_SourceFile"];
            }
            if (oRow["Calculated_NetQuantity"] != DBNull.Value)
            {
                m_dNetQuantity = (decimal)oRow["Calculated_NetQuantity"];
            }
            if ((oRow["PartDtl_JobNum"] != DBNull.Value) && (string.IsNullOrEmpty((string)oRow["PartDtl_JobNum"]) == false))
            {
                m_sJobNum = (string)oRow["PartDtl_JobNum"];
            }
            if (oRow["PartDtl_OrderNum"] != DBNull.Value)
            {
                m_iOrderNum = (int)oRow["PartDtl_OrderNum"];
            }
            if (oRow["PartDtl_OrderLine"] != DBNull.Value)
            {
                m_iOrderLine = (int)oRow["PartDtl_OrderLine"];
            }
            if (oRow["PartDtl_OrderRelNum"] != DBNull.Value)
            {
                m_iOrderRelease = (int)oRow["PartDtl_OrderRelNum"];
            }
            if (oRow["PartDtl_PONum"] != DBNull.Value)
            {
                m_iPONum = (int)oRow["PartDtl_PONum"];
            }
            if (oRow["PartDtl_POLine"] != DBNull.Value)
            {
                m_iPOLine = (int)oRow["PartDtl_POLine"];
            }
            if (oRow["PartDtl_PORelNum"] != DBNull.Value)
            {
                m_iPORelease = (int)oRow["PartDtl_PORelNum"];
            }
        }

        #endregion

        #region Properties

        public string PartNum
        {
            get { return m_sPartNum; }
        }
        public DateTime DueDate
        {
            get { return m_dtDueDate; }
        }
        public string SourceFile
        {
            get { return m_sSourceFile; }
        }
        public decimal NetQuantity
        {
            get { return m_dNetQuantity; }
        }
        public string JobNum
        {
            get { return m_sJobNum; }
        }
        public int OrderNum
        {
            get { return m_iOrderNum; }
        }
        public int OrderLine
        {
            get { return m_iOrderLine; }
        }
        public int OrderRelease
        {
            get { return m_iOrderRelease; }
        }
        public int PONum
        {
            get { return m_iPONum; }
        }
        public int POLine
        {
            get { return m_iPOLine; }
        }
        public int PORelease
        {
            get { return m_iPORelease; }
        }

        public bool DueDateNotSet
        {
            get { return m_bDueDateNotSet; }
        }
        #endregion

        #region Data Members

        private string m_sPartNum;
        private DateTime m_dtDueDate;
        private string m_sSourceFile;
        private decimal m_dNetQuantity;
        private string m_sJobNum;
        private int m_iOrderNum;
        private int m_iOrderLine;
        private int m_iOrderRelease;
        private int m_iPONum;
        private int m_iPOLine;
        private int m_iPORelease;

        private bool m_bDueDateNotSet;
        #endregion
    }

}
