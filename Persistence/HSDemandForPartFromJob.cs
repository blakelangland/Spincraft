using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace HorizonScientific
{
    public class HSDemandForPartFromJob
    {
        //DemandForPartsFromJobs

        #region Constructors

        public HSDemandForPartFromJob(DataRow oRow)
        {
            if ((oRow["PartDtl_PartNum"] != DBNull.Value) && (string.IsNullOrEmpty((string)oRow["PartDtl_PartNum"]) == false))
            {
                m_sPartNum = (string)oRow["PartDtl_PartNum"];
            }
            if ((oRow["PartDtl_RevisionNum"] != DBNull.Value) && (string.IsNullOrEmpty((string)oRow["PartDtl_RevisionNum"]) == false))
            {
                m_sRevisionNumber = (string)oRow["PartDtl_RevisionNum"];
            }
            if ((oRow["JobMtl_Description"] != DBNull.Value) && (string.IsNullOrEmpty((string)oRow["JobMtl_Description"]) == false))
            {
                m_sDescription = (string)oRow["JobMtl_Description"];
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
            if (oRow["Calculated_NetDemandQuantity"] != DBNull.Value)
            {
                m_dNetDemandQuantity = (decimal)oRow["Calculated_NetDemandQuantity"];
            }
            if ((oRow["PartDtl_SourceFile"] != DBNull.Value) && (string.IsNullOrEmpty((string)oRow["PartDtl_SourceFile"]) == false))
            {
                m_sSourceFile = (string)oRow["PartDtl_SourceFile"];
            }
            if ((oRow["JobMtl_JobNum"] != DBNull.Value) && (string.IsNullOrEmpty((string)oRow["JobMtl_JobNum"]) == false))
            {
                m_sJobNumber = (string)oRow["JobMtl_JobNum"];
            }
            if (oRow["JobMtl_AssemblySeq"] != DBNull.Value)
            {
                m_iAssemblySequence = (int)oRow["JobMtl_AssemblySeq"];
            }
            if (oRow["JobMtl_MtlSeq"] != DBNull.Value)
            {
                m_iMaterialSequence = (int)oRow["JobMtl_MtlSeq"];
            }
            if (oRow["JobMtl_RelatedOperation"] != DBNull.Value)
            {
                m_iRelatedOperation = (int)oRow["JobMtl_RelatedOperation"];
            }
            if (oRow["JobMtl_RequiredQty"] != DBNull.Value)
            {
                m_dJobRequiredQuantity = (decimal)oRow["JobMtl_RequiredQty"];
            }
            if (oRow["JobMtl_IssuedQty"] != DBNull.Value)
            {
                m_dJobIssuedQuantity = (decimal)oRow["JobMtl_IssuedQty"];
            }
            if (oRow["JobMtl_ReqDate"] != DBNull.Value)
            {
                m_dtJobRequiredDate = (DateTime)oRow["JobMtl_ReqDate"];
                m_bRequiredDateNotSet = false;
            }
            else
            {
                m_dtJobRequiredDate = DateTime.Now;
                m_bRequiredDateNotSet = true;
            }
            if (oRow["JobHead_JobEngineered"] != DBNull.Value)
            {
                m_bEngineered = (bool)oRow["JobHead_JobEngineered"];
            }
            if (oRow["JobHead_JobReleased"] != DBNull.Value)
            {
                m_bReleased = (bool)oRow["JobHead_JobReleased"];
            }
            if (oRow["JobHead_JobFirm"] != DBNull.Value)
            {
                m_bFirm = (bool)oRow["JobHead_JobFirm"];
            }
            if (oRow["JobProd_OrderNum"] != DBNull.Value)
            {
                m_iOrderNum = (int)oRow["JobProd_OrderNum"];
            }
            if (oRow["JobProd_OrderLine"] != DBNull.Value)
            {
                m_iOrderLine = (int)oRow["JobProd_OrderLine"];
            }
            if (oRow["JobProd_OrderRelNum"] != DBNull.Value)
            {
                m_iOrderRelNum = (int)oRow["JobProd_OrderRelNum"];
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
        public string JobNumber
        {
            get { return m_sJobNumber; }
        }
        public int AssemblySequence
        {
            get { return m_iAssemblySequence; }
        }
        public int MaterialSequence
        {
            get { return m_iMaterialSequence; }
        }
        public int RelatedOperation
        {
            get { return m_iRelatedOperation; }
        }
        public decimal JobRequiredQuantity
        {
            get { return m_dJobRequiredQuantity; }
        }
        public decimal JobIssuedQuantity
        {
            get { return m_dJobIssuedQuantity; }
        }
        public DateTime JobRequiredDate
        {
            get { return m_dtJobRequiredDate; }
        }
        public bool Engineered
        {
            get { return m_bEngineered; }
        }
        public bool Released
        {
            get { return m_bReleased; }
        }
        public bool Firm
        {
            get { return m_bFirm; }
        }
        public int OrderNum
        {
            get { return m_iOrderNum; }
        }
        public int OrderLine
        {
            get { return m_iOrderLine; }
        }
        public int OrderRelNum
        {
            get { return m_iOrderRelNum; }
        }
        public bool DueDateNotSet
        {
            get { return m_bDueDateNotSet; }
        }
        public bool RequiredDateNotSet
        {
            get { return m_bRequiredDateNotSet; }
        }

        #endregion

        #region Data Members

        private string m_sPartNum;
        private string m_sRevisionNumber;
        private string m_sDescription;
        private DateTime m_dtDueDate;
        private decimal m_dNetDemandQuantity;
        private string m_sSourceFile;
        private string m_sJobNumber;
        private int m_iAssemblySequence;
        private int m_iMaterialSequence;
        private int m_iRelatedOperation;
        private decimal m_dJobRequiredQuantity;
        private decimal m_dJobIssuedQuantity;
        private DateTime m_dtJobRequiredDate;
        private bool m_bEngineered;
        private bool m_bReleased;
        private bool m_bFirm;
        private int m_iOrderNum;
        private int m_iOrderLine;
        private int m_iOrderRelNum;

        private bool m_bDueDateNotSet;
        private bool m_bRequiredDateNotSet;
        #endregion

    }
}
