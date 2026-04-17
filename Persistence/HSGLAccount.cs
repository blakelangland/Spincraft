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
    public class HSGLAccount
    {
        #region Constructors

        public HSGLAccount(DataRow oDataRow)
        {
            if ((oDataRow["GLAccount_Company"] != DBNull.Value) && (string.IsNullOrEmpty((string)oDataRow["GLAccount_Company"]) == false))
            {
                m_sCompany = (string)oDataRow["GLAccount_Company"];
            }
            if ((oDataRow["GLAccount_COACode"] != DBNull.Value) && (string.IsNullOrEmpty((string)oDataRow["GLAccount_COACode"]) == false))
            {
                m_sCOACode = (string)oDataRow["GLAccount_COACode"];
            }
            if (oDataRow["GLAccount_Active"] != DBNull.Value)
            {
                m_bActive = (bool)oDataRow["GLAccount_Active"];
            }
            if (oDataRow["GLAccount_EffFrom"] != DBNull.Value)
            {
                m_dtEffectiveFrom = (DateTime)oDataRow["GLAccount_EffFrom"];
            }
            if (oDataRow["GLAccount_EffTo"] != DBNull.Value)
            {
                m_dtEffectiveTo = (DateTime)oDataRow["GLAccount_EffTo"];
            }
            if ((oDataRow["GLAccount_GLAccount"] != DBNull.Value) && (string.IsNullOrEmpty((string)oDataRow["GLAccount_GLAccount"]) == false))
            {
                m_sGLAccount = (string)oDataRow["GLAccount_GLAccount"];
            }
            if ((oDataRow["GLAccount_AccountDesc"] != DBNull.Value) && (string.IsNullOrEmpty((string)oDataRow["GLAccount_AccountDesc"]) == false))
            {
                m_sAccountDescription = (string)oDataRow["GLAccount_AccountDesc"];
            }
            if ((oDataRow["GLAccount_SegValue1"] != DBNull.Value) && (string.IsNullOrEmpty((string)oDataRow["GLAccount_SegValue1"]) == false))
            {
                m_sSegmentValue1 = (string)oDataRow["GLAccount_SegValue1"];
            }
            if ((oDataRow["GLAccount_SegValue2"] != DBNull.Value) && (string.IsNullOrEmpty((string)oDataRow["GLAccount_SegValue2"]) == false))
            {
                m_sSegmentValue2 = (string)oDataRow["GLAccount_SegValue2"];
            }
            if ((oDataRow["GLAccount_SegValue3"] != DBNull.Value) && (string.IsNullOrEmpty((string)oDataRow["GLAccount_SegValue3"]) == false))
            {
                m_sSegmentValue3 = (string)oDataRow["GLAccount_SegValue3"];
            }
            if ((oDataRow["COAActCat_CategoryID"] != DBNull.Value) && (string.IsNullOrEmpty((string)oDataRow["COAActCat_CategoryID"]) == false))
            {
                m_sCategoryId = (string)oDataRow["COAActCat_CategoryID"];
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
        public string Description
        {
            get { return m_sAccountDescription; }
            set { m_sAccountDescription = value; }
        }
        public bool Active
        {
            get { return m_bActive; }
            set { m_bActive = value; }
        }
        public DateTime EffectiveFrom
        {
            get { return m_dtEffectiveFrom; }
            set { m_dtEffectiveFrom = value; }
        }
        public DateTime EffectiveTo
        {
            get { return m_dtEffectiveTo; }
            set { m_dtEffectiveTo = value; }
        }
        public string GLAccount
        {
            get { return m_sGLAccount; }
            set { m_sGLAccount = value; }
        }
        public string AccountDescription
        {
            get { return m_sAccountDescription; }
            set { m_sAccountDescription = value; }
        }
        public string SegmentValue1
        {
            get { return m_sSegmentValue1; }
            set { m_sSegmentValue1 = value; }
        }
        public string SegmentValue2
        {
            get { return m_sSegmentValue2; }
            set { m_sSegmentValue2 = value; }
        }
        public string SegmentValue3
        {
            get { return m_sSegmentValue3; }
            set { m_sSegmentValue3 = value; }
        }
        public string CategoryId
        {
            get { return m_sCategoryId; }
            set { m_sCategoryId = value; }
        }

        public HSGLAccountCategory Category
        {
            get { return m_oCategory; }
            set { m_oCategory = value; }
        }
        #endregion

        #region Data Members
        private string m_sCompany;
        private string m_sCOACode;
        private bool m_bActive;
        private DateTime m_dtEffectiveFrom;
        private DateTime m_dtEffectiveTo;
        private string m_sGLAccount;
        private string m_sAccountDescription;
        private string m_sSegmentValue1;
        private string m_sSegmentValue2;
        private string m_sSegmentValue3;
        private string m_sCategoryId;

        private HSGLAccountCategory m_oCategory;
        #endregion
    }
}
