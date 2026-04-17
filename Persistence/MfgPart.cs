using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Ice.Core;
using Erp.Adapters;
using Erp.BO;
using Erp.Proxy.BO;
using Ice.Lib.Framework;
using Ice.Adapters;
using Ice.BO;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Drawing;
//using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Drawing.Spreadsheet;
using SpreadsheetLight;
using SpreadsheetLight.Charts;
using Ice.Lib.Searches;
using System.IO;


namespace HorizonScientific
{
    public class MfgPart
    {
        #region Constructors

        public MfgPart(DataRow oRow, HSValidateParts oAllParts)
        {
            if (oRow["Part_Company"] != DBNull.Value)
            {
                m_sCompany = (string)oRow["Part_Company"];
            }
            if (oRow["Part_PartNum"] != DBNull.Value)
            {
                m_sPartNum = (string)oRow["Part_PartNum"];
            }
            if (oRow["Part_PartDescription"] != DBNull.Value)
            {
                m_sDescription = (string)oRow["Part_PartDescription"];
            }
            if (oRow["PartRev_RevisionNum"] != DBNull.Value)
            {
                m_sRevNum = (string)oRow["PartRev_RevisionNum"];
            }
            if (oRow["Part_InActive"] != DBNull.Value)
            {
                m_bInactive = (bool)oRow["Part_InActive"];
            }
            if (oRow["PartRev_Approved"] != DBNull.Value)
            {
                m_bRevApproved = (bool)oRow["PartRev_Approved"];
            }
            if (oRow["PartRev_ApprovedDate"] != DBNull.Value)
            {
                m_dtApprovedDate = (DateTime)oRow["PartRev_ApprovedDate"];
            }
            if (oRow["PartRev_EffectiveDate"] != DBNull.Value)
            {
                m_dtEffectiveDate = (DateTime)oRow["PartRev_EffectiveDate"];
            }
            if (oRow["Part_TypeCode"] != DBNull.Value)
            {
                m_sTypeCode = (string)oRow["Part_TypeCode"];
            }
            if (oRow["Part_PhantomBOM"] != DBNull.Value)
            {
                m_bPhantomBOM = (bool)oRow["Part_PhantomBOM"];
            }
            if (oRow["Part_QtyBearing"] != DBNull.Value)
            {
                m_bQtyBearing = (bool)oRow["Part_QtyBearing"];
            }
            if (oRow["Calculated_Cost"] != DBNull.Value)
            {
                m_dCostPerUnit = (decimal)oRow["Calculated_Cost"];
            }
            if (oRow["Part_IUM"] != DBNull.Value)
            {
                m_sUOM = (string)oRow["Part_IUM"];
            }
            // we default quantity per and ext quantity to 1
            m_dQuantityPer = 1;

            // find the reference to the part in the part master table
            m_oPartMaster = oAllParts.AllParts.FirstOrDefault(oItem => string.Compare(oItem.PartNum, m_sPartNum, true) == 0);
        }

        #endregion

        #region Methods

        public void AddPartMaterial(PartMaterial oChildPart)
        {
            m_oChildPartMaterials.Add(oChildPart);
        }

        public void AddPartOperations(List<PartOperation> oOperations)
        {
            // add these operations to this part ordered by the operation seq
            m_oPartOperations = oOperations.OrderBy(x => x.OprSeq).ToList();
        }

        public void SortParts()
        {
            // we will order the parts list such that purchased parts are first followed by manufactured parts
            // the parts will also be in alphabetical order
            List<PartMaterial> oPurchasedParts = m_oChildPartMaterials.Where(x => string.Compare(x.TypeCode, "P", true) == 0).ToList();
            oPurchasedParts = oPurchasedParts.OrderBy(x => x.MaterialPartNum).ToList();
            List<PartMaterial> oManufacturedParts = m_oChildPartMaterials.Where(x => string.Compare(x.TypeCode, "M", true) == 0).ToList();
            oManufacturedParts = oManufacturedParts.OrderBy(x => x.MaterialPartNum).ToList();
            // now clear the list and readd in proper order
            m_oChildPartMaterials.Clear();
            m_oChildPartMaterials.AddRange(oPurchasedParts);
            m_oChildPartMaterials.AddRange(oManufacturedParts);
        }

        public void PrintBOM(int iLevel, bool bApprovedBOM, DateTime dtEffectiveDate)
        {
            string sTabs = "";
            for (int iTabCounter = 0; iTabCounter < iLevel; iTabCounter++)
            {
                sTabs += "\t";
            }
            Console.WriteLine(sTabs + "MfgPart: " + PartNum + " Rev: " + RevNum + " EffectiveDate: " + EffectiveDate.ToShortDateString());
            if (MyPartMaterials.Count > 0)
            {
                iLevel++;
                sTabs = "";
                for (int iTabCounter = 0; iTabCounter < iLevel; iTabCounter++)
                {
                    sTabs += "\t";
                }
            }
            foreach (PartMaterial oTmp in MyPartMaterials)
            {
                if (string.Compare(oTmp.TypeCode, "M", true) == 0)
                {
                    // get the MfgPart for this material
                    // find this MfgPart in our list -- which one???
                    MfgPart oMyMfgPart = oTmp.GetMfgPart(bApprovedBOM, dtEffectiveDate);
                    if (oMyMfgPart == null)
                    {
                        Console.WriteLine("ERROR - Part: " + oTmp.MaterialPartNum + " IS SUPPOSED TO BE A MANUFACTURED PART BUT DOES NOT APPEAR IN THE LIST OF MFG PARTS");
                    }
                    else
                    {
                        oMyMfgPart.PrintBOM(iLevel, bApprovedBOM, dtEffectiveDate);
                    }
                }
                else
                {
                    Console.WriteLine(sTabs + "PartNum: " + oTmp.MaterialPartNum);
                }
            }
        }

        #endregion

        #region Properties
        public string Company
        {
            get { return m_sCompany; }
            set { m_sCompany = value; }
        }

        public string PartNum
        {
            get { return m_sPartNum; }
            set { m_sPartNum = value; }
        }

        public string Description
        {
            get { return m_sDescription; }
            set { m_sDescription = value; }
        }
        public string RevNum
        {
            get { return m_sRevNum; }
            set { m_sRevNum = value; }
        }
        private bool Inactive
        {
            get { return m_bInactive; }
            set { m_bInactive = value; }
        }
        public bool RevApproved
        {
            get { return m_bRevApproved; }
            set { m_bRevApproved = value; }
        }
        public DateTime ApprovedDate
        {
            get { return m_dtApprovedDate; }
            set { m_dtApprovedDate = value; }
        }
        public DateTime EffectiveDate
        {
            get { return m_dtEffectiveDate; }
            set { m_dtEffectiveDate = value; }
        }
        public string TypeCode
        {
            get { return m_sTypeCode; }
            set { m_sTypeCode = value; }
        }

        public bool PhantomBOM
        {
            get { return m_bPhantomBOM; }
            set { m_bPhantomBOM = value; }
        }
        public bool QtyBearing
        {
            get { return m_bQtyBearing; }
            set { m_bQtyBearing = value; }
        }
        public decimal QuantityPer
        {
            get { return m_dQuantityPer; }
            set { m_dQuantityPer = value; }
        }
        public decimal CostPerUnit
        {
            get { return m_dCostPerUnit; }
            set { m_dCostPerUnit = value; }
        }
        public string UOM
        {
            get { return m_sUOM; }
            set { m_sUOM = value; }
        }
        public List<PartMaterial> MyPartMaterials
        {
            get { return m_oChildPartMaterials; }
            set { m_oChildPartMaterials = value; }
        }

        public List<PartOperation> MyPartOperations
        {
            get { return m_oPartOperations; }
            set { m_oPartOperations = value; }
        }

        public HSPartData PartMaster
        {
            get { return m_oPartMaster; }
            set { m_oPartMaster = value; }
        }
        #endregion

        #region Data Members

        private string m_sCompany;
        private string m_sPartNum;
        private string m_sDescription;
        private string m_sRevNum;
        private bool m_bInactive;
        private bool m_bRevApproved;
        private DateTime m_dtApprovedDate;
        private DateTime m_dtEffectiveDate;
        private string m_sTypeCode;
        private bool m_bPhantomBOM;
        private bool m_bQtyBearing;
        private decimal m_dQuantityPer;
        private decimal m_dCostPerUnit;
        private string m_sUOM;

        private List<PartMaterial> m_oChildPartMaterials = new List<PartMaterial>();
        private List<PartOperation> m_oPartOperations = new List<PartOperation>();
        private HSPartData m_oPartMaster;
        #endregion
    }
}