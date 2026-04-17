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
    public class HSGLControl
    {
        #region Constructors

        public HSGLControl(DataRow oDataRow)
        {
            if ((oDataRow["GLCntrlAcct_Company"] != DBNull.Value) && (string.IsNullOrEmpty((string)oDataRow["GLCntrlAcct_Company"]) == false))
            {
                m_sCompany = (string)oDataRow["GLCntrlAcct_Company"];
            }
            if ((oDataRow["GLCntrlAcct_BookID"] != DBNull.Value) && (string.IsNullOrEmpty((string)oDataRow["GLCntrlAcct_BookID"]) == false))
            {
                m_sBookId = (string)oDataRow["GLCntrlAcct_BookID"];
            }
            if ((oDataRow["GLCntrlAcct_GLControlType"] != DBNull.Value) && (string.IsNullOrEmpty((string)oDataRow["GLCntrlAcct_GLControlType"]) == false))
            {
                m_sControlType = (string)oDataRow["GLCntrlAcct_GLControlType"];
            }
            if ((oDataRow["GLCntrlAcct_GLControlCode"] != DBNull.Value) && (string.IsNullOrEmpty((string)oDataRow["GLCntrlAcct_GLControlCode"]) == false))
            {
                m_sControlCode = (string)oDataRow["GLCntrlAcct_GLControlCode"];
            }
            if ((oDataRow["GLCntrlAcct_GLAcctContext"] != DBNull.Value) && (string.IsNullOrEmpty((string)oDataRow["GLCntrlAcct_GLAcctContext"]) == false))
            {
                m_sGLAccountContext = (string)oDataRow["GLCntrlAcct_GLAcctContext"];
            }
            if ((oDataRow["GLCntrlAcct_GLAccount"] != DBNull.Value) && (string.IsNullOrEmpty((string)oDataRow["GLCntrlAcct_GLAccount"]) == false))
            {
                m_sGLAccount = (string)oDataRow["GLCntrlAcct_GLAccount"];
            }
            if (oDataRow["GLCTAcctCntxt_Required"] != DBNull.Value)
            {
                m_bRequired = (bool)oDataRow["GLCTAcctCntxt_Required"];
            }
            if ((oDataRow["GLAcctDisp_GLAcctDisp"] != DBNull.Value) && (string.IsNullOrEmpty((string)oDataRow["GLAcctDisp_GLAcctDisp"]) == false))
            {
                m_sGLAccountDisplay = (string)oDataRow["GLAcctDisp_GLAcctDisp"];
            }
            if ((oDataRow["GLAcctDisp_AccountDesc"] != DBNull.Value) && (string.IsNullOrEmpty((string)oDataRow["GLAcctDisp_AccountDesc"]) == false))
            {
                m_sGLAccountDescription = (string)oDataRow["GLAcctDisp_AccountDesc"];
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
        public string BookId
        {
            get { return m_sBookId; }
            set { m_sBookId = value; }
        }
        public string ControlType
        {
            get { return m_sControlType; }
            set { m_sControlType = value; }
        }
        public string ControlCode
        {
            get { return m_sControlCode; }
            set { m_sControlCode = value; }
        }
        public string GLAccountContext
        {
            get { return m_sGLAccountContext; }
            set { m_sGLAccountContext = value; }
        }
        public string GLAccount
        {
            get { return m_sGLAccount; }
            set { m_sGLAccount = value; }
        }
        public bool Required
        {
            get { return m_bRequired; }
            set { m_bRequired = value; }
        }
        public string GLAccountDisplay
        {
            get { return m_sGLAccountDisplay; }
            set { m_sGLAccountDisplay = value; }
        }
        public string GLAccountDescription
        {
            get { return m_sGLAccountDescription; }
            set { m_sGLAccountDescription = value; }
        }
        #endregion

        #region Data Members
        private string m_sCompany;
        private string m_sBookId;
        private string m_sControlType;
        private string m_sControlCode;
        private string m_sGLAccountContext;
        private string m_sGLAccount;
        private bool m_bRequired;
        private string m_sGLAccountDisplay;
        private string m_sGLAccountDescription;
        #endregion
    }
}