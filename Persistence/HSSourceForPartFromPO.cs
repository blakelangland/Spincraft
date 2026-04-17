using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace HorizonScientific
{
    public class HSSourceForPartFromPO
    {
        //SourceForPartsFromPOs

        #region Constructors

        public HSSourceForPartFromPO(DataRow oRow)
        {
            if ((oRow["PartDtl_PartNum"] != DBNull.Value) && (string.IsNullOrEmpty((string)oRow["PartDtl_PartNum"]) == false))
            {
                m_sPartNum = (string)oRow["PartDtl_PartNum"];
            }
            if ((oRow["PartDtl_RevisionNum"] != DBNull.Value) && (string.IsNullOrEmpty((string)oRow["PartDtl_RevisionNum"]) == false))
            {
                m_sRevisionNumber = (string)oRow["PartDtl_RevisionNum"];
            }
            if ((oRow["PODetail_LineDesc"] != DBNull.Value) && (string.IsNullOrEmpty((string)oRow["PODetail_LineDesc"]) == false))
            {
                m_sDescription = (string)oRow["PODetail_LineDesc"];
            }
            if (oRow["PartDtl_DueDate"] != DBNull.Value)
            {
                m_dtDueDate = (DateTime)oRow["PartDtl_DueDate"];
                m_bDueDateNotSet = false;
            }
            else
            {
                // if there is no due date set then we assume it is due now
                m_dtDueDate = DateTime.Now;
                m_bDueDateNotSet = true;
            }
            if (oRow["Calculated_NetDemandQuantity"] != DBNull.Value)
            {
                m_dNetDemandQuantity = (decimal)oRow["Calculated_NetDemandQuantity"];
            }
            if ((oRow["PartDtl_SourceFile"] != DBNull.Value) && (string.IsNullOrEmpty((string)oRow["PartDtl_SourceFile"]) == false))
            {
                m_sSourceFile = (string)oRow["PartDtl_SourceFile"];
            }
            if (oRow["PORel_PONum"] != DBNull.Value)
            {
                m_iPONumber = (int)oRow["PORel_PONum"];
            }
            if (oRow["PORel_POLine"] != DBNull.Value)
            {
                m_iLineNumber = (int)oRow["PORel_POLine"];
            }
            if (oRow["PORel_PORelNum"] != DBNull.Value)
            {
                m_iReleaseNumber = (int)oRow["PORel_PORelNum"];
            }
            if (oRow["PORel_XRelQty"] != DBNull.Value)
            {
                m_dRequiredQuantity = (decimal)oRow["PORel_XRelQty"];
            }
            if (oRow["PORel_ReceivedQty"] != DBNull.Value)
            {
                m_dReceivedQuantity = (decimal)oRow["PORel_ReceivedQty"];
            }
            if (oRow["PORel_DueDate"] != DBNull.Value)
            {
                m_dtPODueDate = (DateTime)oRow["PORel_DueDate"];
                m_bPODueDateNotSet = false;
            }
            else
            {
                // without a due date set we assume it must be due now
                m_dtPODueDate = DateTime.Now;
                m_bPODueDateNotSet = true;
            }
            if (oRow["PORel_PromiseDt"] != DBNull.Value)
            {
                m_dtPromiseDate = (DateTime)oRow["PORel_PromiseDt"];
                m_bPOPromiseDateNotSet = false;
            }
            else
            {
                // without a promise date set we assume it is the same as the due date
                m_dtPromiseDate = m_dtPODueDate;
                m_bPOPromiseDateNotSet = true;
            }
            if (oRow["POHeader_Approve"] != DBNull.Value)
            {
                m_bApproved = (bool)oRow["POHeader_Approve"];
            }
            if (oRow["POHeader_Confirmed"] != DBNull.Value)
            {
                m_bConfirmed = (bool)oRow["POHeader_Confirmed"];
            }
        }

        #endregion

        #region Properties

        public string PartNum
        {
            get { return m_sPartNum; }
        }
        public string RevisionNumber
        {
            get { return m_sRevisionNumber; }
        }
        public string Description
        {
            get { return m_sDescription; }
        }
        public DateTime DueDate
        {
            get { return m_dtDueDate; }
        }
        public decimal NetDemandQuantity
        {
            get { return m_dNetDemandQuantity; }
        }
        public string SourceFile
        {
            get { return m_sSourceFile; }
        }
        public int PONumber
        {
            get { return m_iPONumber; }
        }
        public int LineNumber
        {
            get { return m_iLineNumber; }
        }
        public int ReleaseNumber
        {
            get { return m_iReleaseNumber; }
        }
        public decimal RequiredQuantity
        {
            get { return m_dRequiredQuantity; }
        }
        public decimal ReceivedQuantity
        {
            get { return m_dReceivedQuantity; }
        }
        public DateTime PODueDate
        {
            get { return m_dtPODueDate; }
        }
        public DateTime PromiseDate
        {
            get { return m_dtPromiseDate; }
        }
        public bool Approved
        {
            get { return m_bApproved; }
        }
        public bool Confirmed
        {
            get { return m_bConfirmed; }
        }
  

        public bool DueDateNotSet
        {
            get { return m_bDueDateNotSet; }
        }
        public bool PODueDateNotSet
        {
            get { return m_bPODueDateNotSet; }
        }
        public bool POPromiseDateNotSet
        {
            get { return m_bPOPromiseDateNotSet; }
        }
        #endregion

        #region Data Members

        private string m_sPartNum;
        private string m_sRevisionNumber;
        private string m_sDescription;
        private DateTime m_dtDueDate;
        private decimal m_dNetDemandQuantity;
        private string m_sSourceFile;
        private int m_iPONumber;
        private int m_iLineNumber;
        private int m_iReleaseNumber;
        private decimal m_dRequiredQuantity;
        private decimal m_dReceivedQuantity;
        private DateTime m_dtPODueDate;
        private DateTime m_dtPromiseDate;
        private bool m_bApproved;
        private bool m_bConfirmed;

        private bool m_bDueDateNotSet;
        private bool m_bPODueDateNotSet;
        private bool m_bPOPromiseDateNotSet;
        #endregion
    }
}
