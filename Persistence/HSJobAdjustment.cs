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
    public class HSJobAdjustment
    {
        #region constructors
        public HSJobAdjustment()
        {
        }
        #endregion

        #region Methods
        public void CloseOpenOperations(Session oSession, string sCompany, string sJobNum, int iAssemblySeq, int iOperationSeq, string sOpCode, string sResourceGroupId, string sLaborType,
			string sLaborEntryMethod, string sEmployeeNum, string sEmployeeName, string sJCDepartment)
        {
			JobAdjustmentAdapter oJobAdjustmentAdapter = null;
			try
			{
				// time the operation was completed
				DateTime dtCurrentDate = DateTime.Now;
				// we will always enable these so we can set the labor record
				bool bComplete = true;
				bool bOpComplete = true;
				bool bEnableLaborQty = true;
				bool bEnableScrapQty = true;
				bool bEnableDiscrepentQty = true;

				// fixed values
				int iLaborHeadSequence = 1; // always 1
				int iLaborDetailSequence = 1; // always 1
				decimal dQtyCompleted = 0; // always 0
				decimal dActProductionHours = 0; // always 0
				decimal dActSetupHours = 0; // always 0
				decimal dActBurCost = 0; // always 0
				decimal dActLaborCost = 0; // always 0
				decimal dLaborQty = 0; // always 0
				decimal dLaborHours = 0; // always 0
				decimal dLaborCost = 0; // always 0
				decimal dBurdenHours = 0; // always 0
				decimal dBurdenCost = 0; // always 0
				string sLaborRoleCode = "";
				bool bDisplayTimeTypeCode = false;
				bool bLaborDisplayProjectRoleCode = false;
				string sLaborTimeTypeCode = "";
				int iAttributeSetID = 0;

				ILauncher EpiLaunch = new ILauncher(oSession);
				oJobAdjustmentAdapter = new JobAdjustmentAdapter(EpiLaunch);
				oJobAdjustmentAdapter.BOConnect();

				JobAdjustmentDataSet oJobAdjustmentDataSet = new JobAdjustmentDataSet();
				JobAdjustmentDataSet.JobsDataTable oJobTable = oJobAdjustmentDataSet.Jobs;
				JobAdjustmentDataSet.JobsRow oJobRow = oJobTable.AddJobsRow(sCompany, sJobNum, Guid.NewGuid(), "A");

				// add the labor detail to the dataset
				JobAdjustmentDataSet.JALaborDtlDataTable oJALaborDetail = oJobAdjustmentDataSet.JALaborDtl;
				oJALaborDetail.AddJALaborDtlRow(dtCurrentDate, "Complete Opr", iAssemblySeq, iOperationSeq, dQtyCompleted, dActProductionHours, dActSetupHours, dActBurCost, dActLaborCost, 
					sEmployeeNum, sEmployeeName, dLaborQty, sLaborType, dLaborHours, dLaborCost, dBurdenHours, dBurdenCost, bComplete, bOpComplete, sCompany, iLaborHeadSequence, 
					iLaborDetailSequence, sOpCode, sResourceGroupId, sJCDepartment, sJobNum, sLaborEntryMethod, bEnableLaborQty, bEnableScrapQty, bEnableDiscrepentQty,
					sLaborRoleCode, bDisplayTimeTypeCode, bLaborDisplayProjectRoleCode, sLaborTimeTypeCode, iAttributeSetID, Guid.NewGuid(), "A");

				oJobAdjustmentAdapter.CommitLaborAdj(oJobAdjustmentDataSet);
			}
			catch (Exception e)
            {
				Console.WriteLine(e.Message);
				Console.WriteLine(e.InnerException);
            }
			finally
            {
				if (oJobAdjustmentAdapter != null)
                {
					oJobAdjustmentAdapter.Dispose();
					oJobAdjustmentAdapter = null;
				}

			}
		}
        #endregion

        #region Properties
        #endregion

        #region Data Memebers
        #endregion
    }
}
