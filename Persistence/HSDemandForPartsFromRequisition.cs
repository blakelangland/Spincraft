using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace HorizonScientific
{
    public class HSDemandForPartsFromRequisition
    {
        //PartDemandFromReqs
        #region Constructors

        public HSDemandForPartsFromRequisition(DataRow oRow)
        {
            if (oRow["ReqHead_ReqNum"] != DBNull.Value)
            {
                m_iReqNumber = (int)oRow["ReqHead_ReqNum"];
            }
            if (oRow["ReqDetail_ReqLine"] != DBNull.Value)
            {
                m_iReqLine = (int)oRow["ReqDetail_ReqLine"];
            }
            if (oRow["ReqDetail_OpenLine"] != DBNull.Value)
            {
                m_bOpenLine = (bool)oRow["ReqDetail_OpenLine"];
            }
            if ((oRow["ReqDetail_PartNum"] != DBNull.Value) && (string.IsNullOrEmpty((string)oRow["ReqDetail_PartNum"]) == false))
            {
                m_sPartNum = (string)oRow["ReqDetail_PartNum"];
            }
            if (oRow["ReqDetail_DueDate"] != DBNull.Value)
            {
                m_dtDueDate = (DateTime)oRow["ReqDetail_DueDate"];
                m_bDueDateNotSet = false;
            }
            else
            {
                m_dtDueDate = DateTime.Now;
                m_bDueDateNotSet = true;
            }
            if (oRow["ReqDetail_XOrderQty"] != DBNull.Value)
            {
                m_dRequiredQuantity = (decimal)oRow["ReqDetail_XOrderQty"];
            }
            if ((oRow["ReqDetail_CommentText"] != DBNull.Value) && (string.IsNullOrEmpty((string)oRow["ReqDetail_CommentText"]) == false))
            {
                m_sComment = (string)oRow["ReqDetail_CommentText"];
            }
            if ((oRow["ReqHead_RequestorID"] != DBNull.Value) && (string.IsNullOrEmpty((string)oRow["ReqHead_RequestorID"]) == false))
            {
                m_sRequestorId = (string)oRow["ReqHead_RequestorID"];
            }
            if (oRow["ReqHead_RequestDate"] != DBNull.Value)
            {
                m_dtRequestDate = (DateTime)oRow["ReqHead_RequestDate"];
            }
            else
            {
                m_dtRequestDate = DateTime.Now;
            }
            if ((oRow["ReqHead_ReqActionID"] != DBNull.Value) && (string.IsNullOrEmpty((string)oRow["ReqHead_ReqActionID"]) == false))
            {
                m_sRequestAction = (string)oRow["ReqHead_ReqActionID"];
            }
            if ((oRow["ReqHead_CurrDispatcherID"] != DBNull.Value) && (string.IsNullOrEmpty((string)oRow["ReqHead_CurrDispatcherID"]) == false))
            {
                m_sDispatcherId = (string)oRow["ReqHead_CurrDispatcherID"];
            }
            if ((oRow["ReqHead_Note"] != DBNull.Value) && (string.IsNullOrEmpty((string)oRow["ReqHead_Note"]) == false))
            {
                m_sNote = (string)oRow["ReqHead_Note"];
            }
            if ((oRow["Calculated_ReqStatus"] != DBNull.Value) && (string.IsNullOrEmpty((string)oRow["Calculated_ReqStatus"]) == false))
            {
                m_sRequestStatus = (string)oRow["Calculated_ReqStatus"];
            }
        }

        #endregion

        #region Properties

        public int ReqNumber
        {
            get { return m_iReqNumber; }
        }
        public int LineNumber
        {
            get { return m_iReqLine; }
        }
        public bool OpenLine
        {
            get { return m_bOpenLine; }
        }
        public string PartNum
        {
            get { return m_sPartNum; }
        }
        public string Description
        {
            get { return m_sDescription; }
        }
        public DateTime DueDate
        {
            get { return m_dtDueDate; }
        }
        public decimal RequiredQuantity
        {
            get { return m_dRequiredQuantity; }
        }
        public string Comment
        {
            get { return m_sComment; }
        }
        public string RequestorId
        {
            get { return m_sRequestorId; }
        }
        public DateTime RequestDate
        {
            get { return m_dtRequestDate; }
        }
        public string RequestAction
        {
            get { return m_sRequestAction; }
        }
        public string DispatcherId
        {
            get { return m_sDispatcherId; }
        }
        public string Note
        {
            get { return m_sNote; }
        }
        public string RequestStatus
        {
            get { return m_sRequestStatus; }
        }
        public bool DueDateNotSet
        {
            get { return m_bDueDateNotSet; }
        }
        public bool RequestDateNotSet
        {
            get { return m_bRequestDateNotSet; }
        }
        #endregion

        #region Data Members

        private int m_iReqNumber;
        private int m_iReqLine;
        private bool m_bOpenLine;
        private string m_sPartNum;
        private string m_sDescription;
        private DateTime m_dtDueDate;
        private decimal m_dRequiredQuantity;
        private string m_sComment;
        private string m_sRequestorId;
        private DateTime m_dtRequestDate;
        private string m_sRequestAction;
        private string m_sDispatcherId;
        private string m_sNote;
        private string m_sRequestStatus;

        private bool m_bDueDateNotSet;
        private bool m_bRequestDateNotSet;
        #endregion
    }
}
