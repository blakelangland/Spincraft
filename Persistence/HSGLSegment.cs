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


namespace HorizonScientific
{
    public class HSGLSegment
    {
        #region Constructors

        public HSGLSegment(DataRow oDataRow)
        {
            if ((oDataRow["COASegment_Company"] != DBNull.Value) && (string.IsNullOrEmpty((string)oDataRow["COASegment_Company"]) == false))
            {
                m_sCompany = (string)oDataRow["COASegment_Company"];
            }
            if ((oDataRow["COASegment_COACode"] != DBNull.Value) && (string.IsNullOrEmpty((string)oDataRow["COASegment_COACode"]) == false))
            {
                m_sCOACode = (string)oDataRow["COASegment_COACode"];
            }
            if (oDataRow["COASegment_SegmentNbr"] != DBNull.Value)
            {
                m_iSegmentNumber = (int)oDataRow["COASegment_SegmentNbr"];
            }
            if ((oDataRow["COASegment_SegmentName"] != DBNull.Value) && (string.IsNullOrEmpty((string)oDataRow["COASegment_SegmentName"]) == false))
            {
                SegmentName = (string)oDataRow["COASegment_SegmentName"];
            }
            if ((oDataRow["COASegment_Abbreviation"] != DBNull.Value) && (string.IsNullOrEmpty((string)oDataRow["COASegment_Abbreviation"]) == false))
            {
                m_sSegmentAbbreviation = (string)oDataRow["COASegment_Abbreviation"];
            }
            if (oDataRow["COASegment_MaxLength"] != DBNull.Value)
            {
                m_iMaximumLength = (int)oDataRow["COASegment_MaxLength"];
            }
            if (oDataRow["COASegment_MinLength"] != DBNull.Value)
            {
                m_iMinimumLength = (int)oDataRow["COASegment_MinLength"];
            }
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
        public string COACode
        {
            get { return m_sCOACode; }
            set { m_sCOACode = value; }
        }
        public int SegmentNumber
        {
            get { return m_iSegmentNumber; }
            set { m_iSegmentNumber = value; }
        }
        public string SegmentName
        {
            get { return m_sSegmentName; }
            set { m_sSegmentName = value; }
        }
        public string SegmentAbbreviation
        {
            get { return m_sSegmentAbbreviation; }
            set { m_sSegmentAbbreviation = value; }
        }
        public int MaximumLength
        {
            get { return m_iMaximumLength; }
            set { m_iMaximumLength = value; }
        }
        public int MinimumLength
        {
            get { return m_iMinimumLength; }
            set { m_iMinimumLength = value; }
        }
        #endregion

        #region Data Members
        private string m_sCompany;
        private string m_sCOACode;
        private int m_iSegmentNumber;
        private string m_sSegmentName;
        private string m_sSegmentAbbreviation;
        private int m_iMaximumLength;
        private int m_iMinimumLength;
        #endregion
    }
}
