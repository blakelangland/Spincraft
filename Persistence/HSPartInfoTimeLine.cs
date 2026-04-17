using HorizonScientific;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace HorizonScientific
{
    public class HSPartInfoTimeLine
    {
        #region Constructors

        public HSPartInfoTimeLine()
        {
        }

        #endregion

        #region Properties

        public decimal CurrentInventoryLevel
        {
            get { return m_dCurrentInventoryLevel; }
            set { m_dCurrentInventoryLevel = value; }
        }
        public decimal NetChange
        {
            get { return m_dNetChange; }
            set { m_dNetChange = value; }
        }
        public DateTime DatePOIsIssued
        {
            get { return m_dtDatePOIsIssued; }
            set { m_dtDatePOIsIssued = value; }
        }
        public DateTime DateOfChange
        {
            get { return m_dtDateOfChange; }
            set { m_dtDateOfChange = value; }
        }
        public string ActionType
        {
            get { return m_sActionType; }
            set { m_sActionType = value; }
        }
        public string ReasonForChange
        {
            get { return m_sReasonForChange; }
            set { m_sReasonForChange = value; }
        }

        public int PONum
        {
            get { return m_iPONum; }
            set { m_iPONum = value; }
        }

        public int POLine
        {
            get { return m_iPOLine; }
            set { m_iPOLine = value; }
        }
        public int PORelease
        {
            get { return m_iPORelease; }
            set { m_iPORelease = value; }
        }
        public DateTime PromiseDate
        {
            get { return m_dtPromiseDate; }
            set { m_dtPromiseDate = value; }
        }

        public int OrderNum
        {
            get { return m_iOrderNum; }
            set { m_iOrderNum = value; }
        }

        public int OrderLine
        {
            get { return m_iOrderLine; }
            set { m_iOrderLine = value; }
        }

        public int OrderRelNum
        {
            get { return m_iOrderRelNum; }
            set { m_iOrderRelNum = value; }
        }

        public string JobNum
        {
            get { return m_sJobNum; }
            set { m_sJobNum = value; }
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

        public int RelatedOperation
        {
            get { return m_iRelatedOperation; }
            set { m_iRelatedOperation = value; }
        }

        public int RequisitionNumber
        {
            get { return m_iRequisitionNumber; }
            set { m_iRequisitionNumber = value; }
        }

        public string RequestedBy
        {
            get { return m_sRequestedBy; }
            set { m_sRequestedBy = value; }
        }

        public DateTime ExpectedArrivalDate
        {
            get { return m_dtExpectedArrivalDate; }
            set { m_dtExpectedArrivalDate = value; }
        }

        public string POComments
        {
            get { return m_sPOComments; }
            set { m_sPOComments = value; }
        }

        public PartTimeDomainInfo Parent
        {
            get { return m_oParent; }
            set { m_oParent = value; }
        }

        public bool PartShortage
        {
            get { return m_bPartShortage; }
            set { m_bPartShortage = value; }
        }
        #endregion

        #region Data Members

        private decimal m_dCurrentInventoryLevel;
        private decimal m_dNetChange;
        private DateTime m_dtDatePOIsIssued;
        private DateTime m_dtDateOfChange;
        private string m_sActionType;
        private string m_sReasonForChange;

        private int m_iPONum;
        private int m_iPOLine;
        private int m_iPORelease;
        private DateTime m_dtPromiseDate;
        private int m_iOrderNum;
        private int m_iOrderLine;
        private int m_iOrderRelNum;
        private string m_sJobNum;
        private int m_iAssemblySequence;
        private int m_iMaterialSequence;
        private int m_iRelatedOperation;
        private int m_iRequisitionNumber;
        private string m_sRequestedBy;

        private DateTime m_dtExpectedArrivalDate;
        private string m_sPOComments;

        private PartTimeDomainInfo m_oParent;
        private bool m_bPartShortage;

        #endregion
    }
}
