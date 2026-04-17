using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing;
//using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Drawing.Spreadsheet;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Wordprocessing;

using Erp.Adapters;
using Erp.Adapters.Controls;
using Erp.BO;
using Erp.Proxy.BO;
using Ice.Adapters;
using Ice.BO;
using Ice.Core;
using Ice.Lib.Framework;
using Ice.Lib.Searches;
using SpreadsheetLight;
using SpreadsheetLight.Charts;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.ServiceModel.Channels;
using System.ServiceModel.Dispatcher;
using System.Text;
using System.Threading.Tasks;

using static Ice.Core.Session;



namespace HorizonScientific
{
    //
    // THESE CLASSES ARE USED TO DETERMINE WHICH OPERATIONS NEED TO BE COMPLETED BEFORE THE GIVEN OPERATION IS COMPLETED
    //
    public class EmployeeInfo
    {
        #region Constructors

        public EmployeeInfo(DataRow oRow)
        {
            if (oRow["EmpBasic_Company"] != DBNull.Value)
            {
                m_sCompany = (string)oRow["EmpBasic_Company"];
            }
            if (oRow["EmpBasic_EmpID"] != DBNull.Value)
            {
                m_sEmployeeId = (string)oRow["EmpBasic_EmpID"];
            }
            if (oRow["EmpBasic_FirstName"] != DBNull.Value)
            {
                m_sFirstName = (string)oRow["EmpBasic_FirstName"];
            }
            if (oRow["EmpBasic_LastName"] != DBNull.Value)
            {
                m_sLastName = (string)oRow["EmpBasic_LastName"];
            }
            if (oRow["EmpBasic_Name"] != DBNull.Value)
            {
                m_sName = (string)oRow["EmpBasic_Name"];
            }
            if (oRow["EmpBasic_Shift"] != DBNull.Value)
            {
                m_iShift = (int)oRow["EmpBasic_Shift"];
            }
            if (oRow["EmpBasic_LaborRate"] != DBNull.Value)
            {
                m_dLaborRate = (decimal)oRow["EmpBasic_LaborRate"];
            }
            if (oRow["EmpBasic_EmpStatus"] != DBNull.Value)
            {
                m_sEmployeeStatus = (string)oRow["EmpBasic_EmpStatus"];
            }
            if (oRow["EmpBasic_ExpenseCode"] != DBNull.Value)
            {
                m_sExpenseCode = (string)oRow["EmpBasic_ExpenseCode"];
            }
            if (oRow["EmpBasic_JCDept"] != DBNull.Value)
            {
                m_sJobCostDepartment = (string)oRow["EmpBasic_JCDept"];
            }
            if (oRow["EmpBasic_SupervisorID"] != DBNull.Value)
            {
                m_sSupervisor = (string)oRow["EmpBasic_SupervisorID"];
            }
            if (oRow["EmpBasic_DcdUserID"] != DBNull.Value)
            {
                m_sUserId = (string)oRow["EmpBasic_DcdUserID"];
            }
            if (oRow["EmpBasic_ProductionWorker"] != DBNull.Value)
            {
                m_bProductionWorker = (bool)oRow["EmpBasic_ProductionWorker"];
            }
            if (oRow["EmpBasic_MaterialHandler"] != DBNull.Value)
            {
                m_bMaterialHandler = (bool)oRow["EmpBasic_MaterialHandler"];
            }
            if (oRow["EmpBasic_ShipRecv"] != DBNull.Value)
            {
                m_bShipReceive = (bool)oRow["EmpBasic_ShipRecv"];
            }
            if (oRow["EmpBasic_ShopSupervisor"] != DBNull.Value)
            {
                m_bShopSupervisor = (bool)oRow["EmpBasic_ShopSupervisor"];
            }
            if (oRow["EmpBasic_WarehouseManager"] != DBNull.Value)
            {
                m_bWarehouseManager = (bool)oRow["EmpBasic_WarehouseManager"];
            }
            if (oRow["EmpBasic_CanReportQty"] != DBNull.Value)
            {
                m_bCanReportQty = (bool)oRow["EmpBasic_CanReportQty"];
            }
            if (oRow["EmpBasic_CanReportScrapQty"] != DBNull.Value)
            {
                m_bCanReportScrapQty = (bool)oRow["EmpBasic_CanReportScrapQty"];
            }
            if (oRow["EmpBasic_CanReportNCQty"] != DBNull.Value)
            {
                m_bCanReportNonConformantQty = (bool)oRow["EmpBasic_CanReportNCQty"];
            }
            if (oRow["EmpBasic_ResourceGrpID"] != DBNull.Value)
            {
                m_sResourceGroupId = (string)oRow["EmpBasic_ResourceGrpID"];
            }
            if (oRow["EmpBasic_ResourceID"] != DBNull.Value)
            {
                m_sResourceId = (string)oRow["EmpBasic_ResourceID"];
            }
            if (oRow["EmpBasic_CanEnterIndirectTime"] != DBNull.Value)
            {
                m_bCanEnterIndirectTime = (bool)oRow["EmpBasic_CanEnterIndirectTime"];
            }
            if (oRow["EmpBasic_CanEnterProductionTime"] != DBNull.Value)
            {
                m_bCanEnterProductionTime = (bool)oRow["EmpBasic_CanEnterProductionTime"];
            }
            if (oRow["EmpBasic_CanEnterSetupTime"] != DBNull.Value)
            {
                m_bCanEnterSetupTime = (bool)oRow["EmpBasic_CanEnterSetupTime"];
            }
            if (oRow["EmpBasic_DefaultIndirectCode"] != DBNull.Value)
            {
                m_sIndirectCode = (string)oRow["EmpBasic_DefaultIndirectCode"];
            }
            if (oRow["EmpBasic_DefaultExpenseCode"] != DBNull.Value)
            {
                m_sDefaultExpenseCode = (string)oRow["EmpBasic_DefaultExpenseCode"];
            }
            if (oRow["EmpBasic_DefaultResourceGrpID"] != DBNull.Value)
            {
                m_sDefaultResourceGroupId = (string)oRow["EmpBasic_DefaultResourceGrpID"];
            }
            if (oRow["EmpBasic_DefaultResourceID"] != DBNull.Value)
            {
                m_sDefaultResourceId = (string)oRow["EmpBasic_DefaultResourceID"];
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
        public string EmployeeId
        {
            get { return m_sEmployeeId; }
            set { m_sEmployeeId = value; }
        }
        public string FirstName
        {
            get { return m_sFirstName; }
            set { m_sFirstName = value; }
        }
        public string LastName
        {
            get { return m_sLastName; }
            set { m_sLastName = value; }
        }
        public string Name
        {
            get { return m_sName; }
            set { m_sName = value; }
        }
        public int Shift
        {
            get { return m_iShift; }
            set { m_iShift = value; }
        }
        public decimal LaborRate
        {
            get { return m_dLaborRate; }
            set { m_dLaborRate = value; }
        }

        public string EmployeeStatus
        {
            get { return m_sEmployeeStatus; }
            set { m_sEmployeeStatus = value; }
        }
        public string ExpenseCode
        {
            get { return m_sExpenseCode; }
            set { m_sExpenseCode = value; }
        }
        public string JobCostDepartment
        {
            get { return m_sJobCostDepartment; }
            set { m_sJobCostDepartment = value; }
        }
        public string Supervisor
        {
            get { return m_sSupervisor; }
            set { m_sSupervisor = value; }
        }
        public string UserId
        {
            get { return m_sUserId; }
            set { m_sUserId = value; }
        }
        public bool ProductionWorker
        {
            get { return m_bProductionWorker; }
            set { m_bProductionWorker = value; }
        }
        public bool MaterialHandler
        {
            get { return m_bMaterialHandler; }
            set { m_bMaterialHandler = value; }
        }
        public bool ShipReceive
        {
            get { return m_bShipReceive; }
            set { m_bShipReceive = value; }
        }
        public bool ShopSupervisor
        {
            get { return m_bShopSupervisor; }
            set { m_bShopSupervisor = value; }
        }
        public bool WarehouseManager
        {
            get { return m_bWarehouseManager; }
            set { m_bWarehouseManager = value; }
        }
        public bool CanReportQty
        {
            get { return m_bCanReportQty; }
            set { m_bCanReportQty = value; }
        }
        public bool CanReportScrapQty
        {
            get { return m_bCanReportScrapQty; }
            set { m_bCanReportScrapQty = value; }
        }
        public bool CanReportNonConformantQty
        {
            get { return m_bCanReportNonConformantQty; }
            set { m_bCanReportNonConformantQty = value; }
        }
        public string ResourceGroupId
        {
            get { return m_sResourceGroupId; }
            set { m_sResourceGroupId = value; }
        }
        public string ResourceId
        {
            get { return m_sResourceId; }
            set { m_sResourceId = value; }
        }
        public bool CanEnterIndirectTime
        {
            get { return m_bCanEnterIndirectTime; }
            set { m_bCanEnterIndirectTime = value; }
        }
        public bool CanEnterProductionTime
        {
            get { return m_bCanEnterProductionTime; }
            set { m_bCanEnterProductionTime = value; }
        }
        public bool CanEnterSetupTime
        {
            get { return m_bCanEnterSetupTime; }
            set { m_bCanEnterSetupTime = value; }
        }
        public string IndirectCode
        {
            get { return m_sIndirectCode; }
            set { m_sIndirectCode = value; }
        }
        public string DefaultExpenseCode
        {
            get { return m_sDefaultExpenseCode; }
            set { m_sDefaultExpenseCode = value; }
        }
        public string DefaultResourceGroupId
        {
            get { return m_sDefaultResourceGroupId; }
            set { m_sDefaultResourceGroupId = value; }
        }
        public string DefaultResourceId
        {
            get { return m_sDefaultResourceId; }
            set { m_sDefaultResourceId = value; }
        }
        #endregion

        #region Data Members
        private string m_sCompany;
        private string m_sEmployeeId;
        private string m_sFirstName;
        private string m_sLastName;
        private string m_sName;
        private int m_iShift;
        private decimal m_dLaborRate;
        private string m_sEmployeeStatus;
        private string m_sExpenseCode;
        private string m_sJobCostDepartment;
        private string m_sSupervisor;
        private string m_sUserId;
        private bool m_bProductionWorker;
        private bool m_bMaterialHandler;
        private bool m_bShipReceive;
        private bool m_bShopSupervisor;
        private bool m_bWarehouseManager;
        private bool m_bCanReportQty;
        private bool m_bCanReportScrapQty;
        private bool m_bCanReportNonConformantQty;
        private string m_sResourceGroupId;
        private string m_sResourceId;
        private bool m_bCanEnterIndirectTime;
        private bool m_bCanEnterProductionTime;
        private bool m_bCanEnterSetupTime;
        private string m_sIndirectCode;
        private string m_sDefaultExpenseCode;
        private string m_sDefaultResourceGroupId;
        private string m_sDefaultResourceId;

        #endregion
    }

    public class JobOperationHierarchy
    {
        #region Constructors

        public JobOperationHierarchy(DataRow oRow)
        {
            if (oRow["JobHead_JobNum"] != DBNull.Value)
            {
                m_sJobNum = (string)oRow["JobHead_JobNum"];
            }
            if (oRow["JobHead_PartNum"] != DBNull.Value)
            {
                m_sPartNum = (string)oRow["JobHead_PartNum"];
            }
            if (oRow["JobHead_RevisionNum"] != DBNull.Value)
            {
                m_sRevisionNum = (string)oRow["JobHead_RevisionNum"];
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
            if (oRow["Calculated_HasLabor"] != DBNull.Value)
            {
                m_bHasLabor = (bool)oRow["Calculated_HasLabor"];
            }
            if (oRow["JobOper_OpComplete"] != DBNull.Value)
            {
                m_bOperationComplete = (bool)oRow["JobOper_OpComplete"];
            }
            if (oRow["JobOper_SubContract"] != DBNull.Value)
            {
                m_bSubcontract = (bool)oRow["JobOper_SubContract"];
            }
            if (oRow["JobOper_AssemblySeq"] != DBNull.Value)
            {
                m_iAssemblySequence = (int)oRow["JobOper_AssemblySeq"];
            }
            if (oRow["JobOper_OprSeq"] != DBNull.Value)
            {
                m_iOperationSequence = (int)oRow["JobOper_OprSeq"];
            }
            if (oRow["JobOper_OpCode"] != DBNull.Value)
            {
                m_sOperationCode = (string)oRow["JobOper_OpCode"];
            }
            if (oRow["JobOper_EstSetHours"] != DBNull.Value)
            {
                m_dEstimatedSetupHours = (decimal)oRow["JobOper_EstSetHours"];
            }
            if (oRow["JobOper_EstProdHours"] != DBNull.Value)
            {
                m_dEstimatedProductionHours = (decimal)oRow["JobOper_EstProdHours"];
            }
            if (oRow["JobOper_ActSetupHours"] != DBNull.Value)
            {
                m_dActualSetupHours = (decimal)oRow["JobOper_ActSetupHours"];
            }
            if (oRow["JobOper_ActProdHours"] != DBNull.Value)
            {
                m_dActualProductionHours = (decimal)oRow["JobOper_ActProdHours"];
            }
            if (oRow["JobOper_ActSetupRwkHours"] != DBNull.Value)
            {
                m_dActualSetupReworkHours = (decimal)oRow["JobOper_ActSetupRwkHours"];
            }
            if (oRow["JobOper_ActProdRwkHours"] != DBNull.Value)
            {
                m_dActualProductionReworkHours = (decimal)oRow["JobOper_ActProdRwkHours"];
            }
            if (oRow["JobOper_RunQty"] != DBNull.Value)
            {
                m_dRunQuantity = (decimal)oRow["JobOper_RunQty"];
            }
            if (oRow["JobOper_QtyCompleted"] != DBNull.Value)
            {
                m_dQuantityCompleted = (decimal)oRow["JobOper_QtyCompleted"];
            }
            if (oRow["JobOper_LaborEntryMethod"] != DBNull.Value)
            {
                m_sLaborEntryMethod = (string)oRow["JobOper_LaborEntryMethod"];
            }
            if (oRow["JobOpDtl_ResourceGrpID"] != DBNull.Value)
            {
                m_sResourceGroupId = (string)oRow["JobOpDtl_ResourceGrpID"];
            }
            if (oRow["JobOpDtl_ResourceID"] != DBNull.Value)
            {
                m_sResourceId = (string)oRow["JobOpDtl_ResourceID"];
            }
            if (oRow["JobOper_SchedRelation"] != DBNull.Value)
            {
                m_sScheduleRelationship = (string)oRow["JobOper_SchedRelation"];
            }
            if (oRow["JobOper_StartDate"] != DBNull.Value)
            {
                m_dtStartDate = (DateTime)oRow["JobOper_StartDate"];
            }
            if (oRow["JobOper_StartHour"] != DBNull.Value)
            {
                m_dStartHour = (decimal)oRow["JobOper_StartHour"];
            }
            if (oRow["JobOper_DueDate"] != DBNull.Value)
            {
                m_dtDueDate = (DateTime)oRow["JobOper_DueDate"];
            }
            if (oRow["JobOper_DueHour"] != DBNull.Value)
            {
                m_dDueHour = (decimal)oRow["JobOper_DueHour"];
            }

            // check to see if we have anything remaining to do for this operation
            // if the operation is already set to complete we just skip it
            if (m_bOperationComplete == false)
            {
                // operation is not complete so we check the quantity
                m_dQuantityRemaining = m_dRunQuantity - m_dQuantityCompleted;
                if (m_dQuantityRemaining < 0)
                {
                    m_dQuantityRemaining = 0;
                }
                // see if there is any more setup time that should be logged
                m_dSetupTimeRemaining = m_dEstimatedSetupHours - m_dActualSetupHours;
                if (m_dSetupTimeRemaining < 0)
                {
                    m_dSetupTimeRemaining = 0;
                }
                // see if there is any more production time that should be logged
                m_dProductionTimeRemaining = m_dEstimatedProductionHours - m_dActualProductionHours;
                if (m_dProductionTimeRemaining < 0)
                {
                    m_dProductionTimeRemaining = 0;
                }
            }
        }

        #endregion

        #region Methods
        public void AddDependentAssembly(JobAssemblyHierarchy oAssembly)
        {
            m_oDependentAssemblies.Add(oAssembly);
        }

        public void AddMaterialToIssue(JobMaterialHierarchy oMaterial)
        {
            m_oMaterialsToIssue.Add(oMaterial);
        }
        #endregion

        #region Properties
        public string JobNum
        {
            get { return m_sJobNum; }
            set { m_sJobNum = value; }
        }
        public string PartNum
        {
            get { return m_sPartNum; }
            set { m_sPartNum = value; }
        }
        public string RevisionNum
        {
            get { return m_sRevisionNum; }
            set { m_sRevisionNum = value; }
        }
        public bool Engineered
        {
            get { return m_bEngineered; }
            set { m_bEngineered = value; }
        }
        public bool Released
        {
            get { return m_bReleased; }
            set { m_bReleased = value; }
        }
        public bool Firm
        {
            get { return m_bFirm; }
            set { m_bFirm = value; }
        }
        public bool HasLabor
        {
            get { return m_bHasLabor; }
            set { m_bHasLabor = value; }
        }
        public bool OperationComplete
        {
            get { return m_bOperationComplete; }
            set { m_bOperationComplete = value; }
        }
        public bool Subcontract
        {
            get { return m_bSubcontract; }
            set { m_bSubcontract = value; }
        }
        public int AssemblySequence
        {
            get { return m_iAssemblySequence; }
            set { m_iAssemblySequence = value; }
        }
        public int OperationSequence
        {
            get { return m_iOperationSequence; }
            set { m_iOperationSequence = value; }
        }
        public string OperationCode
        {
            get { return m_sOperationCode; }
            set { m_sOperationCode = value; }
        }
        public decimal EstimatedSetupHours
        {
            get { return m_dEstimatedSetupHours; }
            set { m_dEstimatedSetupHours = value; }
        }
        public decimal EstimatedProductionHours
        {
            get { return m_dEstimatedProductionHours; }
            set { m_dEstimatedProductionHours = value; }
        }
        public decimal ActualSetupHours
        {
            get { return m_dActualSetupHours; }
            set { m_dActualSetupHours = value; }
        }
        public decimal ActualProductionHours
        {
            get { return m_dActualProductionHours; }
            set { m_dActualProductionHours = value; }
        }
        public decimal ActualSetupReworkHours
        {
            get { return m_dActualSetupReworkHours; }
            set { m_dActualSetupReworkHours = value; }
        }
        public decimal ActualProductionReworkHours
        {
            get { return m_dActualProductionReworkHours; }
            set { m_dActualProductionReworkHours = value; }
        }
        public decimal RunQuantity
        {
            get { return m_dRunQuantity; }
            set { m_dRunQuantity = value; }
        }
        public decimal QuantityCompleted
        {
            get { return m_dQuantityCompleted; }
            set { m_dQuantityCompleted = value; }
        }
        public string LaborEntryMethod
        {
            get { return m_sLaborEntryMethod; }
            set { m_sLaborEntryMethod = value; }
        }
        public string ResourceGroupId
        {
            get { return m_sResourceGroupId; }
            set { m_sResourceGroupId = value; }
        }
        public string ResourceId
        {
            get { return m_sResourceId; }
            set { m_sResourceId = value; }
        }
        public string ScheduleRelationship
        {
            get { return m_sScheduleRelationship; }
            set { m_sScheduleRelationship = value; }
        }
        public DateTime StartDate
        {
            get { return m_dtStartDate; }
            set { m_dtStartDate = value; }
        }
        public decimal StartHour
        {
            get { return m_dStartHour; }
            set { m_dStartHour = value; }
        }
        public DateTime DueDate
        {
            get { return m_dtDueDate; }
            set { m_dtDueDate = value; }
        }
        public decimal DueHour
        {
            get { return m_dDueHour; }
            set { m_dDueHour = value; }
        }
        public decimal QuantityRemaining
        {
            get { return m_dQuantityRemaining; }
            set { m_dQuantityRemaining = value; }
        }
        public decimal SetupTimeRemaining
        {
            get { return m_dSetupTimeRemaining; }
            set { m_dSetupTimeRemaining = value; }
        }
        public decimal ProductionTimeRemaining
        {
            get { return m_dProductionTimeRemaining; }
            set { m_dProductionTimeRemaining = value; }
        }

        public List<JobAssemblyHierarchy> DependentAssemblies
        {
            get { return m_oDependentAssemblies; }
        }
        public List<JobMaterialHierarchy> MaterialsToIssue
        {
            get { return m_oMaterialsToIssue; }
        }
        #endregion

        #region Data Members
        private string m_sJobNum;
        private string m_sPartNum;
        private string m_sRevisionNum;
        private bool m_bEngineered;
        private bool m_bReleased;
        private bool m_bFirm;
        private bool m_bHasLabor;
        private bool m_bOperationComplete;
        private bool m_bSubcontract;
        private int m_iAssemblySequence;
        private int m_iOperationSequence;
        private string m_sOperationCode;
        private decimal m_dEstimatedSetupHours;
        private decimal m_dEstimatedProductionHours;
        private decimal m_dActualSetupHours;
        private decimal m_dActualProductionHours;
        private decimal m_dActualSetupReworkHours;
        private decimal m_dActualProductionReworkHours;
        private decimal m_dRunQuantity;
        private decimal m_dQuantityCompleted;
        private string m_sLaborEntryMethod;
        private string m_sResourceGroupId;
        private string m_sResourceId;
        private string m_sScheduleRelationship;
        private DateTime m_dtStartDate;
        private decimal m_dStartHour;
        private DateTime m_dtDueDate;
        private decimal m_dDueHour;

        // initialized to zero and only set if operation is not complete yet
        private decimal m_dQuantityRemaining;
        private decimal m_dSetupTimeRemaining;
        private decimal m_dProductionTimeRemaining;

        private List<JobAssemblyHierarchy> m_oDependentAssemblies = new List<JobAssemblyHierarchy>();
        private List<JobMaterialHierarchy> m_oMaterialsToIssue = new List<JobMaterialHierarchy>();

        #endregion
    }

    public class JobMaterialHierarchy
    {
        #region Constructors

        public JobMaterialHierarchy(DataRow oRow)
        {
            if (oRow["JobHead_JobNum"] != DBNull.Value)
            {
                m_sJobNum = (string)oRow["JobHead_JobNum"];
            }
            if (oRow["JobHead_PartNum"] != DBNull.Value)
            {
                m_sParentPartNum = (string)oRow["JobHead_PartNum"];
            }
            if (oRow["JobHead_RevisionNum"] != DBNull.Value)
            {
                m_sParentRevisionNum = (string)oRow["JobHead_RevisionNum"];
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
            if (oRow["JobMtl_IssuedComplete"] != DBNull.Value)
            {
                m_bIssuedComplete = (bool)oRow["JobMtl_IssuedComplete"];
            }
            if (oRow["JobMtl_AssemblySeq"] != DBNull.Value)
            {
                m_iAssemblySequence = (int)oRow["JobMtl_AssemblySeq"];
            }
            if (oRow["JobMtl_MtlSeq"] != DBNull.Value)
            {
                m_iMaterialSequence = (int)oRow["JobMtl_MtlSeq"];
            }
            if (oRow["JobMtl_PartNum"] != DBNull.Value)
            {
                m_sPartNum = (string)oRow["JobMtl_PartNum"];
            }
            if (oRow["JobMtl_RevisionNum"] != DBNull.Value)
            {
                m_sPartRevision = (string)oRow["JobMtl_RevisionNum"];
            }
            if (oRow["Part_TypeCode"] != DBNull.Value)
            {
                m_sTypeCode = (string)oRow["Part_TypeCode"];
            }
            if (oRow["JobMtl_Description"] != DBNull.Value)
            {
                m_sPartDescription = (string)oRow["JobMtl_Description"];
            }
            if (oRow["JobMtl_RelatedOperation"] != DBNull.Value)
            {
                m_iRelatedOperation = (int)oRow["JobMtl_RelatedOperation"];
            }
            if (oRow["JobMtl_RequiredQty"] != DBNull.Value)
            {
                m_dRequiredQuantity = (decimal)oRow["JobMtl_RequiredQty"];
            }
            if (oRow["JobMtl_IssuedQty"] != DBNull.Value)
            {
                m_dIssuedQuantity = (decimal)oRow["JobMtl_IssuedQty"];
            }
            if (oRow["JobMtl_WarehouseCode"] != DBNull.Value)
            {
                m_sFromWarehouseCode = (string)oRow["JobMtl_WarehouseCode"];
            }
            if (oRow["JobMtl_BuyIt"] != DBNull.Value)
            {
                m_bBuyIt = (bool)oRow["JobMtl_BuyIt"];
            }

            // compute what is remaining to be issued
            m_dRemainingQuantity = m_dRequiredQuantity - m_dIssuedQuantity;
            if (m_dRequiredQuantity < 0)
            {
                m_dRequiredQuantity = 0;
            }

            // determine which backlfull bin to use
            if (m_sParentPartNum.StartsWith("CS", StringComparison.InvariantCultureIgnoreCase) == true)
            {
                m_sFromWarehouseBin = "CS-BFLUSH";
            }
            else if (m_sParentPartNum.StartsWith("HM", StringComparison.InvariantCultureIgnoreCase) == true)
            {
                m_sFromWarehouseBin = "HM-BFLUSH";
            }
            else if (m_sParentPartNum.StartsWith("IM", StringComparison.InvariantCultureIgnoreCase) == true)
            {
                m_sFromWarehouseBin = "IM-BFLUSH";
            }
            else if (m_sParentPartNum.StartsWith("IT", StringComparison.InvariantCultureIgnoreCase) == true)
            {
                m_sFromWarehouseBin = "IT-BFLUSH";
            }
            else if (m_sParentPartNum.StartsWith("LP", StringComparison.InvariantCultureIgnoreCase) == true)
            {
                m_sFromWarehouseBin = "LP-BFLUSH";
            }
            else if (m_sParentPartNum.StartsWith("MD", StringComparison.InvariantCultureIgnoreCase) == true)
            {
                m_sFromWarehouseBin = "MD-BFLUSH";
            }
            else if (m_sParentPartNum.StartsWith("QD", StringComparison.InvariantCultureIgnoreCase) == true)
            {
                m_sFromWarehouseBin = "QD-BFLUSH";
            }
            else if (m_sParentPartNum.StartsWith("UC", StringComparison.InvariantCultureIgnoreCase) == true)
            {
                m_sFromWarehouseBin = "UC-BFLUSH";
            }
            else if (m_sParentPartNum.StartsWith("VF", StringComparison.InvariantCultureIgnoreCase) == true)
            {
                m_sFromWarehouseBin = "VF-BFLUSH";
            }
            else if (m_sParentPartNum.StartsWith("VH", StringComparison.InvariantCultureIgnoreCase) == true)
            {
                m_sFromWarehouseBin = "VH-BFLUSH";
            }
            else if (m_sParentPartNum.StartsWith("VL", StringComparison.InvariantCultureIgnoreCase) == true)
            {
                m_sFromWarehouseBin = "VL-BFLUSH";
            }
            else if (m_sParentPartNum.StartsWith("VR", StringComparison.InvariantCultureIgnoreCase) == true)
            {
                m_sFromWarehouseBin = "VR-BFLUSH";
            }
            else
            {
                // catch all
                m_sFromWarehouseBin = "IM-BFLUSH";
            }
            // to warehouse is the same as the from warehouse -- only one warehouse at Federal
            m_sToWarehouseCode = m_sFromWarehouseCode;
            // only one bin for WIP at Federal
            m_sToWarehouseBin = "WIP";
        }

        #endregion

        #region Methods
        #endregion

        #region Properties
        public string JobNum
        {
            get { return m_sJobNum; }
            set { m_sJobNum = value; }
        }
        public string ParentPartNum
        {
            get { return m_sParentPartNum; }
            set { m_sParentPartNum = value; }
        }
        public string ParentRevisionNum
        {
            get { return m_sParentRevisionNum; }
            set { m_sParentRevisionNum = value; }
        }
        public bool Engineered
        {
            get { return m_bEngineered; }
            set { m_bEngineered = value; }
        }
        public bool Released
        {
            get { return m_bReleased; }
            set { m_bReleased = value; }
        }
        public bool Firm
        {
            get { return m_bFirm; }
            set { m_bFirm = value; }
        }
        public bool IssuedComplete
        {
            get { return m_bIssuedComplete; }
            set { m_bIssuedComplete = value; }
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
        public string PartNum
        {
            get { return m_sPartNum; }
            set { m_sPartNum = value; }
        }
        public string PartRevision
        {
            get { return m_sPartRevision; }
            set { m_sPartRevision = value; }
        }
        public string PartDescription
        {
            get { return m_sPartDescription; }
            set { m_sPartDescription = value; }
        }
        public int RelatedOperation
        {
            get { return m_iRelatedOperation; }
            set { m_iRelatedOperation = value; }
        }
        public string TypeCode
        {
            get { return m_sTypeCode; }
            set { m_sTypeCode = value; }
        }
        public decimal RequiredQuantity
        {
            get { return m_dRequiredQuantity; }
            set { m_dRequiredQuantity = value; }
        }
        public decimal IssuedQuantity
        {
            get { return m_dIssuedQuantity; }
            set { m_dIssuedQuantity = value; }
        }
        public string FromWarehouseCode
        {
            get { return m_sFromWarehouseCode; }
            set { m_sFromWarehouseCode = value; }
        }
        public bool BuyIt
        {
            get { return m_bBuyIt; }
            set { m_bBuyIt = value; }
        }


        public bool ManuallyIssued
        {
            get { return m_bManuallyIssued; }
            set { m_bManuallyIssued = value; }
        }
        public decimal RemainingQuantity
        {
            get { return m_dRemainingQuantity; }
            set { m_dRemainingQuantity = value; }
        }
        public string FromWarehouseBin
        {
            get { return m_sFromWarehouseBin; }
            set { m_sFromWarehouseBin = value; }
        }
        public string ToWarehouseCode
        {
            get { return m_sToWarehouseCode; }
            set { m_sToWarehouseCode = value; }
        }
        public string ToWarehouseBin
        {
            get { return m_sToWarehouseBin; }
            set { m_sToWarehouseBin = value; }
        }
        #endregion

        #region Data Members
        private string m_sJobNum;
        private string m_sParentPartNum;
        private string m_sParentRevisionNum;
        private bool m_bEngineered;
        private bool m_bReleased;
        private bool m_bFirm;
        private bool m_bIssuedComplete;
        private int m_iAssemblySequence;
        private int m_iMaterialSequence;
        private string m_sPartNum;
        private string m_sPartRevision;
        private string m_sTypeCode;
        private string m_sPartDescription;
        private int m_iRelatedOperation;
        private decimal m_dRequiredQuantity;
        private decimal m_dIssuedQuantity;
        private string m_sFromWarehouseCode;
        private bool m_bBuyIt;

        private bool m_bManuallyIssued = false;
        private decimal m_dRemainingQuantity;
        private string m_sFromWarehouseBin;
        private string m_sToWarehouseCode;
        private string m_sToWarehouseBin;
        #endregion
    }

    public class JobAssemblyHierarchy
    {
        #region Constructors

        public JobAssemblyHierarchy(DataRow oRow)
        {
            if (oRow["JobHead_JobNum"] != DBNull.Value)
            {
                m_sJobNum = (string)oRow["JobHead_JobNum"];
            }
            if (oRow["JobHead_PartNum"] != DBNull.Value)
            {
                m_sParentPartNum = (string)oRow["JobHead_PartNum"];
            }
            if (oRow["JobHead_RevisionNum"] != DBNull.Value)
            {
                m_sParentRevisionNum = (string)oRow["JobHead_RevisionNum"];
            }
            if (oRow["JobAsmbl_AssemblySeq"] != DBNull.Value)
            {
                m_iAssemblySequence = (int)oRow["JobAsmbl_AssemblySeq"];
            }
            if (oRow["JobAsmbl_Parent"] != DBNull.Value)
            {
                m_iParentAssemblySequence = (int)oRow["JobAsmbl_Parent"];
            }
            if (oRow["JobAsmbl_RelatedOperation"] != DBNull.Value)
            {
                m_iParentAssemblyRelatedOperation = (int)oRow["JobAsmbl_RelatedOperation"];
            }
            if (oRow["JobAsmbl_PriorPeer"] != DBNull.Value)
            {
                m_iPriorPeer = (int)oRow["JobAsmbl_PriorPeer"];
            }
            if (oRow["JobAsmbl_NextPeer"] != DBNull.Value)
            {
                m_iNextPeer = (int)oRow["JobAsmbl_NextPeer"];
            }
            if (oRow["JobAsmbl_Child"] != DBNull.Value)
            {
                m_iFirstChild = (int)oRow["JobAsmbl_Child"];
            }
            if (oRow["Calculated_MinOperationSeq"] != DBNull.Value)
            {
                m_iMinimumOperationSequence = (int)oRow["Calculated_MinOperationSeq"];
            }
        }

        #endregion

        #region Methods

        public JobOperationHierarchy GetFirstOperation()
        {
            JobOperationHierarchy oFirstOperation = null;
            if (m_oOperations.Count == 0)
            {
                if (m_oParentAssembly != null)
                {
                    // we get the first operation in our parent assembly
                    oFirstOperation = m_oParentAssembly.GetFirstOperation();
                }
            }
            else
            {
                oFirstOperation = m_oOperations.OrderBy(p => p.OperationSequence).FirstOrDefault();
            }
            return oFirstOperation;
        }

        public JobOperationHierarchy GetFirstNonSubcontractOperation()
        {
            JobOperationHierarchy oFirstOperation = null;
            List<JobOperationHierarchy> oNonSubcontractOperations = m_oOperations.Where(p => p.Subcontract == false).OrderBy(p => p.OperationSequence).ToList();
            if (oNonSubcontractOperations.Count == 0)
            {
                if (m_oParentAssembly != null)
                {
                    // we get the first operation in our parent assembly
                    oFirstOperation = m_oParentAssembly.GetFirstNonSubcontractOperation();
                }
            }
            else
            {
                oFirstOperation = oNonSubcontractOperations.FirstOrDefault();
            }
            return oFirstOperation;
        }

        public JobOperationHierarchy GetOperation(int iOperationSequence)
        {
            JobOperationHierarchy oOperation = oOperation = m_oOperations.Where(p => p.OperationSequence == iOperationSequence).FirstOrDefault();
            if (oOperation == null)
            {
                // we need to get the first operation in this assembly
                if (m_oOperations.Count > 0)
                {
                    oOperation = m_oOperations[0];
                }

                if (oOperation == null)
                {
                    // we need to go up a level to the parent assembly
                    if (m_oParentAssembly != null)
                    {
                        oOperation = m_oParentAssembly.GetOperation(iOperationSequence);
                    }
                }
            }
            return oOperation;
        }

        public void ListOperationsInOrder(List<JobOperationHierarchy> oOrderedOperations)
        {
            // get our operations in order and add them to the list
            List<JobOperationHierarchy> oMyOperations = m_oOperations.OrderBy(p => p.OperationSequence).ToList();
            // need to handle the odd case where this assembly has no operations but may have child assemblies
            if (oMyOperations.Count == 0)
            {
                // if we have any child assemblies then we need to see review the child assembly operations
                List<JobAssemblyHierarchy> oOrderedChildAssemblies = m_oChildAssemblies.OrderBy(p => p.AssemblySequence).ToList();
                foreach (JobAssemblyHierarchy oDependentAssembly in oOrderedChildAssemblies)
                {
                    oDependentAssembly.ListOperationsInOrder(oOrderedOperations);
                }
            }
            else
            {
                foreach (JobOperationHierarchy oOperation in oMyOperations)
                {
                    // see if there is a dependency first
                    if (oOperation.DependentAssemblies.Count > 0)
                    {
                        // we have dependent assemblies so we need to walk through them and get their operations in order first
                        List<JobAssemblyHierarchy> oDependentAssemblies = oOperation.DependentAssemblies.OrderBy(p => p.AssemblySequence).ToList();
                        foreach (JobAssemblyHierarchy oDependentAssembly in oDependentAssemblies)
                        {
                            oDependentAssembly.ListOperationsInOrder(oOrderedOperations);
                        }
                    }
                    // now we add this operation and continue
                    oOrderedOperations.Add(oOperation);
                }
            }
            return;
        }

        public void AddChildAssembly(JobAssemblyHierarchy oChildAssembly)
        {
            // add child to our list
            m_oChildAssemblies.Add(oChildAssembly);
            // ensure child points to this assembly as its parent
            oChildAssembly.ParentAssembly = this;
        }

        public void AddOperaton(JobOperationHierarchy oOperation)
        {
            m_oOperations.Add(oOperation);
        }

        public void PrintAssemblyStructure(int iIndent)
        {
            // print our own assembly sequence
            for (int i = 0; i < iIndent; i++)
            {
                Console.Write("\t");
            }
            Console.Write("Assembly: " + m_iAssemblySequence.ToString() + " \n");
            foreach (JobAssemblyHierarchy oChilAssembly in m_oChildAssemblies)
            {
                oChilAssembly.PrintAssemblyStructure(iIndent + 1);
            }
        }

        #endregion

        #region Properties
        public string JobNum
        {
            get { return m_sJobNum; }
            set { m_sJobNum = value; }
        }
        public string ParentPartNum
        {
            get { return m_sParentPartNum; }
            set { m_sParentPartNum = value; }
        }
        public string ParentRevisionNum
        {
            get { return m_sParentRevisionNum; }
            set { m_sParentRevisionNum = value; }
        }
        public int AssemblySequence
        {
            get { return m_iAssemblySequence; }
            set { m_iAssemblySequence = value; }
        }
        public int ParentAssemblySequence
        {
            get { return m_iParentAssemblySequence; }
            set { m_iParentAssemblySequence = value; }
        }
        public int ParentAssemblyRelatedOperation
        {
            get { return m_iParentAssemblyRelatedOperation; }
            set { m_iParentAssemblyRelatedOperation = value; }
        }
        public int PriorPeer
        {
            get { return m_iPriorPeer; }
            set { m_iPriorPeer = value; }
        }
        public int NextPeer
        {
            get { return m_iNextPeer; }
            set { m_iNextPeer = value; }
        }
        public int FirstChild
        {
            get { return m_iFirstChild; }
            set { m_iFirstChild = value; }
        }

        public int MinimumOperationSequence
        {
            get { return m_iMinimumOperationSequence; }
            set { m_iMinimumOperationSequence = value; }
        }

        public List<JobAssemblyHierarchy> ChildAssemblies
        {
            get { return m_oChildAssemblies; }
        }
        public List<JobOperationHierarchy> Operations
        {
            get { return m_oOperations; }
        }
        public JobAssemblyHierarchy ParentAssembly
        {
            get { return m_oParentAssembly; }
            set { m_oParentAssembly = value; }
        }
        #endregion

        #region Data Members
        private string m_sJobNum;
        private string m_sParentPartNum;
        private string m_sParentRevisionNum;
        private int m_iAssemblySequence;
        private int m_iParentAssemblySequence;
        private int m_iParentAssemblyRelatedOperation;
        private int m_iPriorPeer;
        private int m_iNextPeer;
        private int m_iFirstChild;
        private int m_iMinimumOperationSequence;

        private List<JobAssemblyHierarchy> m_oChildAssemblies = new List<JobAssemblyHierarchy>();
        private List<JobOperationHierarchy> m_oOperations = new List<JobOperationHierarchy>();
        private JobAssemblyHierarchy m_oParentAssembly = null;

        #endregion
    }

    public class JobAnalysis
    {
        public List<JobOperationHierarchy> m_oOperationsToComplete = new List<JobOperationHierarchy>();
        public int TotalNumberOfOperations;
        public int OperationsCompleted;
    }

    public class AnalyzeOperationsToComplete
    {
        #region Constructors

        public AnalyzeOperationsToComplete() 
        {
            return;
        }

        public static List<string> GetJobListByPartNum(Session oSession, string sPartNum)
        {
            List<string> oJobNumbers = new List<string>();

            //
            // get the list of job numbers tied to this part num
            //
            Ice.Proxy.BO.DynamicQueryImpl oDynamicQuery = WCFServiceSupport.CreateImpl<Ice.Proxy.BO.DynamicQueryImpl>(oSession, Ice.Proxy.BO.DynamicQueryImpl.UriPath);
            QueryExecutionDataSet oQueryExecutionDataSet = oDynamicQuery.GetQueryExecutionParametersByID(BAQConstants.QUERY_GET_JOBS_TIED_TO_PART_NUM);
            foreach (DataRow oParameter in oQueryExecutionDataSet.ExecutionParameter)
            {
                if (string.Compare(oParameter["ParameterID"].ToString(), "PartNum") == 0)
                {
                    oParameter["ParameterValue"] = sPartNum;
                }
            }
            oQueryExecutionDataSet.AcceptChanges();
            DataSet oDataSet = oDynamicQuery.ExecuteByID(BAQConstants.QUERY_GET_JOBS_TIED_TO_PART_NUM, oQueryExecutionDataSet);
            oJobNumbers.Clear();
            foreach (DataRow oRow in oDataSet.Tables[0].Rows)
            {
                if (oRow["JobHead_JobNum"] != DBNull.Value)
                {
                    string sJonNum = (string)oRow["JobHead_JobNum"];
                    oJobNumbers.Add(sJonNum);
                }
            }
            return oJobNumbers;
        }

        public void LoadDataForJob(Session oSession, int iOrderNum, int iOrderLine, int iOrderRelNum, bool bLoadChildJobs)
        {
            string sJobNum = "";

            //
            // get the job number tied to this sales order, line, and release
            //
            Ice.Proxy.BO.DynamicQueryImpl oDynamicQuery = WCFServiceSupport.CreateImpl<Ice.Proxy.BO.DynamicQueryImpl>(oSession, Ice.Proxy.BO.DynamicQueryImpl.UriPath);
            QueryExecutionDataSet oQueryExecutionDataSet = oDynamicQuery.GetQueryExecutionParametersByID(BAQConstants.QUERY_GET_JOB_TIED_TO_ORDER_LINE_RELEASE);
            foreach (DataRow oParameter in oQueryExecutionDataSet.ExecutionParameter)
            {
                if (string.Compare(oParameter["ParameterID"].ToString(), "OrderNum") == 0)
                {
                    oParameter["ParameterValue"] = iOrderNum;
                }
                if (string.Compare(oParameter["ParameterID"].ToString(), "OrderLine") == 0)
                {
                    oParameter["ParameterValue"] = iOrderLine;
                }
                if (string.Compare(oParameter["ParameterID"].ToString(), "OrderRelNum") == 0)
                {
                    oParameter["ParameterValue"] = iOrderRelNum;
                }
            }
            oQueryExecutionDataSet.AcceptChanges();
            DataSet oDataSet = oDynamicQuery.ExecuteByID(BAQConstants.QUERY_GET_JOB_TIED_TO_ORDER_LINE_RELEASE, oQueryExecutionDataSet);

            foreach (DataRow oRow in oDataSet.Tables[0].Rows)
            {
                if (oRow["JobProd_JobNum"] != DBNull.Value)
                {
                    sJobNum = (string)oRow["JobProd_JobNum"];
                    break;
                }
            }

            LoadDataForJob(oSession, sJobNum, bLoadChildJobs);
        }

        public void LoadDataForJob(Session oSession, string sJobNum, bool bLoadChildJobs)
        {

            //
            // GET JOB OPERATIONS
            //
            Ice.Proxy.BO.DynamicQueryImpl oDynamicQuery = WCFServiceSupport.CreateImpl<Ice.Proxy.BO.DynamicQueryImpl>(oSession, Ice.Proxy.BO.DynamicQueryImpl.UriPath);
            QueryExecutionDataSet oQueryExecutionDataSet = oDynamicQuery.GetQueryExecutionParametersByID(BAQConstants.QUERY_GET_JOB_OPERATION_HIERARCHY);
            foreach (DataRow oParameter in oQueryExecutionDataSet.ExecutionParameter)
            {
                if (string.Compare(oParameter["ParameterID"].ToString(), "JobNum") == 0)
                {
                    oParameter["ParameterValue"] = sJobNum;
                }
            }
            oQueryExecutionDataSet.AcceptChanges();
            DataSet oDataSet = oDynamicQuery.ExecuteByID(BAQConstants.QUERY_GET_JOB_OPERATION_HIERARCHY, oQueryExecutionDataSet);
            m_oJobOperations.Clear();
            foreach (DataRow oRow in oDataSet.Tables[0].Rows)
            {
                JobOperationHierarchy oTmpJobOp = new JobOperationHierarchy(oRow);
                // its possible that we may have duplicate job operations due to the join to the
                // job opr detail table in order to extract resources so we will need to eliminate any duplicates
                JobOperationHierarchy oDupe = m_oJobOperations.FirstOrDefault(oItem => (oItem.AssemblySequence == oTmpJobOp.AssemblySequence) && (oItem.OperationSequence == oTmpJobOp.OperationSequence));
                if (oDupe != null)
                {
                    // we should use the one with the resource group set, and if it is not set then use the one with the resource set
                    if (string.IsNullOrEmpty(oDupe.ResourceGroupId) == false)
                    {
                        // dont add anything to the list as we have the one with the resource group set
                        oDupe.ResourceId = oTmpJobOp.ResourceId;
                    }
                    else
                    {
                        // we will set the resource group and resource to whatever the one we just created has
                        oDupe.ResourceGroupId = oTmpJobOp.ResourceGroupId;
                        oDupe.ResourceId = oTmpJobOp.ResourceId;
                    }
                }
                else
                {
                    // this has not been added yet so we can add it to the list
                    m_oJobOperations.Add(new JobOperationHierarchy(oRow));
                }
            }
            // order the operations by assembly and then sequence number
            m_oJobOperations = m_oJobOperations.OrderBy(oItem => oItem.AssemblySequence).ThenBy(x => x.OperationSequence).ToList();
            // build up ordered list of operations that have been completed
            m_oCompletedJobOperations = m_oJobOperations.Where(oItem => oItem.OperationComplete == true).ToList();


            //
            // GET JOB ASSEMBLIES
            //
            oQueryExecutionDataSet = oDynamicQuery.GetQueryExecutionParametersByID(BAQConstants.QUERY_GET_JOB_ASSEMBLY_HIERARCHY);
            foreach (DataRow oParameter in oQueryExecutionDataSet.ExecutionParameter)
            {
                if (string.Compare(oParameter["ParameterID"].ToString(), "JobNum") == 0)
                {
                    oParameter["ParameterValue"] = sJobNum;
                }
            }
            oQueryExecutionDataSet.AcceptChanges();
            oDataSet = oDynamicQuery.ExecuteByID(BAQConstants.QUERY_GET_JOB_ASSEMBLY_HIERARCHY, oQueryExecutionDataSet);
            m_oJobAssemblies.Clear();
            foreach (DataRow oRow in oDataSet.Tables[0].Rows)
            {
                m_oJobAssemblies.Add(new JobAssemblyHierarchy(oRow));
            }

            // establish which assembly is the root assembly -- this should be assembly 0 but we will look for the smallest assembly number
            if (m_oJobAssemblies.Count > 0)
            {
                m_oRootAssembly = m_oJobAssemblies.OrderBy(op => op.AssemblySequence).First();
            }

            // now we need to walk through the list of job assemblies and build the hierarchy of assemblies
            foreach (JobAssemblyHierarchy oAssembly in m_oJobAssemblies)
            {
                // see if this assembly has a parent assembly, if so we need to add this to the child list of the parent assembly
                // the root will never have a parent so we do not evaluate that one
                if (oAssembly != m_oRootAssembly)
                {
                    JobAssemblyHierarchy oParentAssembly = m_oJobAssemblies.FirstOrDefault(x => x.AssemblySequence == oAssembly.ParentAssemblySequence);
                    if (oParentAssembly != null)
                    {
                        oParentAssembly.AddChildAssembly(oAssembly);
                    }
                }
            }

            // ensure assemblies have thier list of operations
            foreach (JobAssemblyHierarchy oAssembly in m_oJobAssemblies)
            {
                List<JobOperationHierarchy> oOperationsInThisAssembly = m_oJobOperations.Where(x => x.AssemblySequence == oAssembly.AssemblySequence).ToList();
                foreach (JobOperationHierarchy oOperation in oOperationsInThisAssembly)
                {
                    oAssembly.AddOperaton(oOperation);
                }
            }

            //
            // now lets print out the assembly structure for debugging purposes
            // get the root assembly which is always assembly sequence 0
            //m_oRootAssembly.PrintAssemblyStructure(0);
            //

            //
            // now we need to indicate which operations have assemblies tied to them
            //
            foreach (JobAssemblyHierarchy oAssembly in m_oJobAssemblies)
            {
                // root assembly will never point to an operation so we skip it
                if (oAssembly.AssemblySequence != 0)
                {
                    // get the parent assembly sequence and the operation in that assembly that this assembly is tied to
                    JobAssemblyHierarchy oParentAssembly = m_oJobAssemblies.FirstOrDefault(x => x.AssemblySequence == oAssembly.ParentAssemblySequence);
                    if (oParentAssembly != null)
                    {
                        // now we need to get the operation in the parent assembly that this assembly is tied to
                        JobOperationHierarchy oOperation = oParentAssembly.GetOperation(oAssembly.ParentAssemblyRelatedOperation);
                        // now we need to establish that this operation has this assembly as a dependency
                        if (oOperation != null)
                        {
                            oOperation.AddDependentAssembly(oAssembly);
                        }
                    }
                }
            }

            //
            //List<JobOperationHierarchy> oOrderOperations = new List<JobOperationHierarchy>();
            //m_oRootAssembly.ListOperationsInOrder(oOrderOperations);
            //
            // now lets print out all operations in order and indicate which operations have assemblies tied to them
            // we start with the root assembly and walk through all operations in the root assembly in order
            // when we encounter an operation with dependedemblies we will walk through those assemblies and print out their operations as well
            //foreach (JobOperationHierarchy oOperation in oOrderOperations)
            //{
            //    Console.WriteLine("Operation: " + oOperation.AssemblySequence.ToString() + "-" + oOperation.OperationSequence.ToString());
            //}



            //
            // GET JOB MATERIALS
            //
            oQueryExecutionDataSet = oDynamicQuery.GetQueryExecutionParametersByID(BAQConstants.QUERY_GET_JOB_MATERIAL_HIERARCHY);
            foreach (DataRow oParameter in oQueryExecutionDataSet.ExecutionParameter)
            {
                if (string.Compare(oParameter["ParameterID"].ToString(), "JobNum") == 0)
                {
                    oParameter["ParameterValue"] = sJobNum;
                }
            }
            oQueryExecutionDataSet.AcceptChanges();
            oDataSet = oDynamicQuery.ExecuteByID(BAQConstants.QUERY_GET_JOB_MATERIAL_HIERARCHY, oQueryExecutionDataSet);
            m_oJobMaterials.Clear();
            foreach (DataRow oRow in oDataSet.Tables[0].Rows)
            {
                m_oJobMaterials.Add(new JobMaterialHierarchy(oRow));
            }
            // order the materials by assembly and then sequence number
            m_oJobMaterials = m_oJobMaterials.OrderBy(oItem => oItem.AssemblySequence).ThenBy(x => x.MaterialSequence).ToList();

            // now we need to find materials that are not tied to an operation and so will mot be backflushed automatically
            // we will have to manually issue these materials when we complete the operation
            foreach (JobMaterialHierarchy oMaterial in m_oJobMaterials)
            {
                if (oMaterial.RelatedOperation == 0)
                {
                    // this material is not tied to an operation so we will need to manually issue it when we complete the operation
                    oMaterial.ManuallyIssued = true;
                    // now we need to determine which operation we will issue this material against
                    JobAssemblyHierarchy oAssembly = m_oJobAssemblies.FirstOrDefault(x => x.AssemblySequence == oMaterial.AssemblySequence);
                    if (oAssembly != null)
                    {
                        // we will get the first operation in this assembly and use that operation sequence number
                        // if there is not an operation then this method will defer to the first operation on the parent, etc.
                        JobOperationHierarchy oOperation = oAssembly.GetFirstNonSubcontractOperation();
                        if (oOperation != null)
                        {
                            oOperation.AddMaterialToIssue(oMaterial);
                        }
                        else
                        {
                            // THIS MATERIAL WILL NEVER GET ISSUED AS WE CANT TIE IT TO ANY OPERATION
                            // REALLY SHOULD NEVER GET HERE!
                        }
                    }
                }
            }


            //
            // GET EMPLOYEE INFO FOR COMPLETEING OPERATIONS
            //
            oQueryExecutionDataSet = oDynamicQuery.GetQueryExecutionParametersByID(BAQConstants.QUERY_GET_EMPLOYEE_INFO);
            oQueryExecutionDataSet.ExecutionParameter.Clear();
            oDataSet = oDynamicQuery.ExecuteByID(BAQConstants.QUERY_GET_EMPLOYEE_INFO, oQueryExecutionDataSet);
            m_oEmployees.Clear();
            foreach (DataRow oRow in oDataSet.Tables[0].Rows)
            {
                m_oEmployees.Add(new EmployeeInfo(oRow));
            }

            m_oChildJobs.Clear();
            if (bLoadChildJobs == true)
            {
                // we need to check if there are any jobs tied to this parent job
                List<string> oJobNumbers = new List<string>();

                //
                // get the list of job numbers tied to this job num
                //
                oDynamicQuery = WCFServiceSupport.CreateImpl<Ice.Proxy.BO.DynamicQueryImpl>(oSession, Ice.Proxy.BO.DynamicQueryImpl.UriPath);
                oQueryExecutionDataSet = oDynamicQuery.GetQueryExecutionParametersByID(BAQConstants.QUERY_GET_JOBS_TIED_TO_JOB);
                foreach (DataRow oParameter in oQueryExecutionDataSet.ExecutionParameter)
                {
                    if (string.Compare(oParameter["ParameterID"].ToString(), "TargetJobNum") == 0)
                    {
                        oParameter["ParameterValue"] = sJobNum;
                    }
                }
                oQueryExecutionDataSet.AcceptChanges();
                oDataSet = oDynamicQuery.ExecuteByID(BAQConstants.QUERY_GET_JOBS_TIED_TO_JOB, oQueryExecutionDataSet);
                oJobNumbers.Clear();
                foreach (DataRow oRow in oDataSet.Tables[0].Rows)
                {
                    if (oRow["JobProd_JobNum"] != DBNull.Value)
                    {
                        string sJonNum = (string)oRow["JobProd_JobNum"];
                        oJobNumbers.Add(sJonNum);
                    }
                }

                foreach (string sChildJobNum in oJobNumbers)
                {
                    AnalyzeOperationsToComplete oChildJob = new AnalyzeOperationsToComplete();
                    oChildJob.LoadDataForJob(oSession, sChildJobNum, false);
                    m_oChildJobs.Add(oChildJob);
                }
            }
        }

        #endregion

        #region Methods
        public string CompleteOperations(Session oSession, List<JobOperationHierarchy> oOperationsToComplete, string sEmployeeId)
        {
            string sAllErrorMessages = string.Empty;

            EmployeeInfo oEmployee = m_oEmployees.Find(x => string.Compare(x.EmployeeId, sEmployeeId, true) == 0);
            if (oEmployee == null)
            {
                // we could not find the employee so we cant enter in labor for these operations
                sAllErrorMessages = "Could not complete the operations as the employee id: " + sEmployeeId + " was not found.";
            }
            else
            {
                LaborImpl oLaborImpl = WCFServiceSupport.CreateImpl<LaborImpl>(oSession, Erp.Proxy.BO.LaborImpl.UriPath);
                foreach (JobOperationHierarchy oOperation in oOperationsToComplete)
                {
                    // if this is a subcontract operation then we skip it 
                    // as it should be handled by a PO and receipt
                    if (oOperation.Subcontract == true)
                    {
                        continue;
                    }

                    string sMessage = string.Empty;

                    // CHECK LABOR TYPE -- Only correct when labor entry is either time and quantity, or quantity
                    if ((string.Compare(oOperation.LaborEntryMethod, "T", true) == 0) || (string.Compare(oOperation.LaborEntryMethod, "Q", true) == 0))
                    {
                        // see if we need to add setup time
                        if ((oOperation.SetupTimeRemaining > 0) && (string.Compare(oOperation.LaborEntryMethod, "T", true) == 0))
                        {
                            DateTime dtClockInDate = DateTime.Now;
                            LaborDataSet oLaborDataSet = new LaborDataSet();

                            // we will add a labor detail record to capture setup time
                            oLaborImpl.GetNewLaborDtlNoHdr(oLaborDataSet, sEmployeeId, false, dtClockInDate, 0, dtClockInDate, 0);
                            // get the index of the newly added labor detail record
                            int iAddedLaborIndex = oLaborDataSet.LaborDtl.Rows.Count - 1;
                            // load the job
                            oLaborImpl.DefaultJobNum(oLaborDataSet, oOperation.JobNum);
                            // get laobor rates based on employee and job
                            oLaborImpl.LaborRateCalc(oLaborDataSet);

                            // set the assembly and operation sequence
                            oLaborImpl.DefaultAssemblySeq(oLaborDataSet, oOperation.AssemblySequence);
                            sMessage = string.Empty;
                            oLaborImpl.DefaultOprSeq(oLaborDataSet, oOperation.OperationSequence, out sMessage);
                            if (string.IsNullOrEmpty(sMessage) == false)
                            {
                                sAllErrorMessages += sMessage + Environment.NewLine;
                            }
                            // this is for production labor
                            oLaborDataSet.LaborDtl[iAddedLaborIndex].LaborType = "S";
                            // calculate labor rates again now that we have the operation set and indicated this is for setup
                            oLaborImpl.LaborRateCalc(oLaborDataSet);

                            // if there are setup hours remaining then we enter that now
                            //if ((oOperation.SetupTimeRemaining > 0) && (string.Compare(oOperation.LaborEntryMethod, "T", true) == 0))
                            //{
                                // The following two lines are the normal way to record setup time, but ETG MA WANTS TO SET THE TIME TO ZERO SO AS NOT TO RECORD LABOR OR BURDEN
                                //oLaborDataSet.LaborDtl[iAddedLaborIndex].LaborHrs = oOperation.SetupTimeRemaining;
                                //oLaborDataSet.LaborDtl[iAddedLaborIndex].BurdenHrs = oOperation.SetupTimeRemaining;

                                // ETG MA WANTS TO SET THE TIME TO ZERO SO AS NOT TO RECORD LABOR OR BURDEN
                                oLaborDataSet.LaborDtl[iAddedLaborIndex].LaborHrs = 0.0M;
                                oLaborDataSet.LaborDtl[iAddedLaborIndex].BurdenHrs = 0.0M;
                            //}
                            // set this to entered
                            oLaborDataSet.LaborDtl[iAddedLaborIndex].TimeStatus = "E";

                            // check these entries for errors
                            sMessage = string.Empty;
                            oLaborImpl.CheckWarnings(oLaborDataSet, out sMessage);
                            if (string.IsNullOrEmpty(sMessage) == false)
                            {
                                sAllErrorMessages += sMessage + Environment.NewLine;
                            }

                            //
                            // I am calling this as I cannot get this to approve the time entry unless I first do this step
                            // likley I am missing the setting of some important fields but I am not sure which ones those would be
                            //
                            sMessage = string.Empty;
                            oLaborImpl.RecallFromApproval(oLaborDataSet, false, out sMessage);
                            if (string.IsNullOrEmpty(sMessage) == false)
                            {
                                sAllErrorMessages += sMessage + Environment.NewLine;
                            }

                            // update labor record
                            oLaborDataSet.LaborDtl[iAddedLaborIndex].RowMod = "U";
                            oLaborImpl.Update(oLaborDataSet);

                            // validate the charge rate for production
                            sMessage = string.Empty;
                            oLaborDataSet.LaborDtl[iAddedLaborIndex].RowMod = "U";
                            oLaborImpl.ValidateChargeRateForTimeType(oLaborDataSet, out sMessage);
                            if (string.IsNullOrEmpty(sMessage) == false)
                            {
                                sAllErrorMessages += sMessage + Environment.NewLine;
                            }
                            // submit this record for approval
                            sMessage = string.Empty;
                            oLaborImpl.SubmitForApproval(oLaborDataSet, false, out sMessage);
                            if (string.IsNullOrEmpty(sMessage) == false)
                            {
                                sAllErrorMessages += sMessage + Environment.NewLine;
                            }
                        }

                        // see if we need to add production time or quantity
                        if ((oOperation.QuantityRemaining > 0) || (oOperation.ProductionTimeRemaining > 0))
                        {
                            DateTime dtClockInDate = DateTime.Now;
                            LaborDataSet oLaborDataSet = new LaborDataSet();

                            // we will add a labor detail record to capture production time
                            oLaborImpl.GetNewLaborDtlNoHdr(oLaborDataSet, sEmployeeId, false, dtClockInDate, 0, dtClockInDate, 0);
                            // get the index of the newly added labor detail record
                            int iAddedLaborIndex = oLaborDataSet.LaborDtl.Rows.Count - 1;
                            // load the job
                            oLaborImpl.DefaultJobNum(oLaborDataSet, oOperation.JobNum);
                            // get laobor rates based on employee and job
                            oLaborImpl.LaborRateCalc(oLaborDataSet);
                            // set the assembly and operation sequence
                            oLaborImpl.DefaultAssemblySeq(oLaborDataSet, oOperation.AssemblySequence);
                            sMessage = string.Empty;
                            oLaborImpl.DefaultOprSeq(oLaborDataSet, oOperation.OperationSequence, out sMessage);
                            if (string.IsNullOrEmpty(sMessage) == false)
                            {
                                sAllErrorMessages += sMessage + Environment.NewLine;
                            }
                            // this is for production labor
                            oLaborDataSet.LaborDtl[iAddedLaborIndex].LaborType = "P";
                            // calculate labor rates again now that we have the operation set and indicated this is for production
                            oLaborImpl.LaborRateCalc(oLaborDataSet);
                            // if we have qty remaining we will enter that now
                            if (oOperation.QuantityRemaining > 0)
                            {
                                sMessage = string.Empty;
                                oLaborImpl.DefaultLaborQty(oLaborDataSet, oOperation.QuantityRemaining, out sMessage);
                                if (string.IsNullOrEmpty(sMessage) == false)
                                {
                                    sAllErrorMessages += sMessage + Environment.NewLine;
                                }
                            }
                            // if there are production hours remaining then we enter that now
                            //if ((oOperation.ProductionTimeRemaining > 0) && (string.Compare(oOperation.LaborEntryMethod, "T", true) == 0))
                            //{
                                // The following two lines are the normal way to record production time, but ETG MA WANTS TO SET THE TIME TO ZERO SO AS NOT TO RECORD LABOR OR BURDEN
                                //oLaborDataSet.LaborDtl[iAddedLaborIndex].LaborHrs = oOperation.ProductionTimeRemaining;
                                //oLaborDataSet.LaborDtl[iAddedLaborIndex].BurdenHrs = oOperation.ProductionTimeRemaining;
                                
                                // ETG MA WANTS TO SET THE TIME TO ZERO SO AS NOT TO RECORD LABOR OR BURDEN
                                oLaborDataSet.LaborDtl[iAddedLaborIndex].LaborHrs = 0.0M;
                                oLaborDataSet.LaborDtl[iAddedLaborIndex].BurdenHrs = 0.0M;
                            //}
                            // set this to entered
                            oLaborDataSet.LaborDtl[iAddedLaborIndex].TimeStatus = "E";

                            // check these entries for errors
                            sMessage = string.Empty;
                            oLaborImpl.CheckWarnings(oLaborDataSet, out sMessage);
                            if (string.IsNullOrEmpty(sMessage) == false)
                            {
                                sAllErrorMessages += sMessage + Environment.NewLine;
                            }

                            //
                            // I am calling this as I cannot get this to approve the time entry unless I first do this step
                            // likley I am missing the setting of some important fields but I am not sure which ones those would be
                            //
                            sMessage = string.Empty;
                            oLaborImpl.RecallFromApproval(oLaborDataSet, false, out sMessage);
                            if (string.IsNullOrEmpty(sMessage) == false)
                            {
                                sAllErrorMessages += sMessage + Environment.NewLine;
                            }

                            // update labor record
                            oLaborDataSet.LaborDtl[iAddedLaborIndex].RowMod = "U";
                            oLaborImpl.Update(oLaborDataSet);

                            // validate the charge rate for production
                            sMessage = string.Empty;
                            oLaborDataSet.LaborDtl[iAddedLaborIndex].RowMod = "U";
                            oLaborImpl.ValidateChargeRateForTimeType(oLaborDataSet, out sMessage);
                            if (string.IsNullOrEmpty(sMessage) == false)
                            {
                                sAllErrorMessages += sMessage + Environment.NewLine;
                            }
                            // submit this record for approval
                            sMessage = string.Empty;
                            oLaborImpl.SubmitForApproval(oLaborDataSet, false, out sMessage);
                            if (string.IsNullOrEmpty(sMessage) == false)
                            {
                                sAllErrorMessages += sMessage + Environment.NewLine;
                            }
                        }
                    }

                    // now we need to check if there are any materials tied to this operation that need to be manually issued
                    if (oOperation.MaterialsToIssue.Count > 0)
                    {
                        string sIssueMessage = this.IssueMaterials(oSession, oOperation.MaterialsToIssue, sEmployeeId);
                        if (string.IsNullOrEmpty(sIssueMessage) == false)
                        {
                            sAllErrorMessages += sIssueMessage + Environment.NewLine;
                        }
                    }
                }
            }

            if (string.IsNullOrEmpty(sAllErrorMessages) == false)
            {
                sAllErrorMessages = "The following errors were encountered when trying to complete the operations:" + Environment.NewLine + sAllErrorMessages;
            }

            return sAllErrorMessages;
        }

        public bool CanOperationBeStarted(int iAssemblySequence, int iOperationSequence, out string sOperationsThatMustBeCompletedFirst)
        {
            // this function will analyze if there are any operations that preceed this operation that have yet to be completed
            // if there are any operations that must be completed prior to this operation being started then those operations will
            // be returned in the string sOperationsThatMustBeCompletedFirst and false will be returned.
            // If there are no prior operations that need to be completed then the function will return true.
            sOperationsThatMustBeCompletedFirst = "";
            bool bCanOperationBeStarted = true;
            List<JobOperationHierarchy> oOperationsToBeCompleted = EvaluateOperationsToComplete(iAssemblySequence, iOperationSequence);
            // if there are any operations in another assembly then we cannot start this operation
            // however, if there are any operations in the current assembly then we need to review
            // the scheduling relation to determine if this operation can be started. If the scheduling
            // relation is set to Finish To Start then the prior operations must be completed first.
            if (oOperationsToBeCompleted.Count > 0)
            {
                foreach (JobOperationHierarchy oOperation in oOperationsToBeCompleted)
                {
                    // we only evaluate operations that are not yet complete
                    if (oOperation.OperationComplete == false)
                    {
                        // if this operation is in a different assembly then we need to include it in the list of operations that must first be completed
                        if (oOperation.AssemblySequence != iAssemblySequence)
                        {
                            bCanOperationBeStarted = false;
                            // add this operation to the list of operations that must be completed first
                            sOperationsThatMustBeCompletedFirst += "Asm: " + oOperation.AssemblySequence.ToString() + " Opr: " + oOperation.OperationSequence.ToString() + "\n";
                        }
                        else
                        {
                            // this operation is in the current assembly so we need to evaluate the scheduling relationship of all operations that
                            // proceed this operation and if any of those are set to "FS" then this operation needs to be completed first

                            // get a list of all operations past this operation that are not yet complete and where the scehduling relationship is "FS"
                            List<JobOperationHierarchy> oOperationsAfterThisOperation = oOperationsToBeCompleted.Where(oItem => (oItem.AssemblySequence == iAssemblySequence) &&
                                (oItem.OperationSequence > oOperation.OperationSequence) && (oItem.OperationComplete == false) && (string.Compare(oItem.ScheduleRelationship, "FS", true) == 0)).ToList();
                            if (oOperationsAfterThisOperation.Count > 0)
                            {
                                // this current operation must be complete before we can start the operation
                                bCanOperationBeStarted = false;
                                // add this operation to the list of operations that must be completed first
                                sOperationsThatMustBeCompletedFirst += "Asm: " + oOperation.AssemblySequence.ToString() + " Opr: " + oOperation.OperationSequence.ToString() + "\n";
                            }
                        }
                    }
                }
            }

            return bCanOperationBeStarted;
        }

        public string IssueMaterials(Session oSession, List<JobMaterialHierarchy> oMaterialsToIssue, string sEmployeeId)
        {
            string sAllErrorMessages = string.Empty;

            EmployeeInfo oEmployee = m_oEmployees.Find(x => string.Compare(x.EmployeeId, sEmployeeId, true) == 0);
            if (oEmployee == null)
            {
                // we could not find the employee so we cant enter in labor for these operations
                sAllErrorMessages = "Could not issue the material as the employee id: " + sEmployeeId + " was not found.";
            }
            else
            {
                foreach (JobMaterialHierarchy oMaterial in oMaterialsToIssue)
                {
                    // we will issue the full quantity remaining for this material
                    if (oMaterial.RemainingQuantity > 0)
                    {
                        this.IssueMaterial(oSession, oMaterial);
                    }
                }
            }

            if (string.IsNullOrEmpty(sAllErrorMessages) == false)
            {
                sAllErrorMessages = "The following errors were encountered when trying to issue material to the job:" + Environment.NewLine + sAllErrorMessages;
            }

            return sAllErrorMessages;
        }

        public void IssueMaterial(Session oSession, JobMaterialHierarchy oMaterial)
        {
            try
            {
                IssueReturnImpl oIssueReturnImpl = WCFServiceSupport.CreateImpl<IssueReturnImpl>(oSession, Erp.Proxy.BO.IssueReturnImpl.UriPath);

                string sMessage;
                string sCallProcess = "IssueMaterial";
                // Create service
                //using (var issueSvc = Ice.Assemblies.ServiceRenderer.GetService<IssueReturnSvcContract>(oSession))
                {
                    IssueReturnDataSet oIssueReturnDataSet = new IssueReturnDataSet();

                    // get a new IssueReturn record
                    oIssueReturnImpl.GetNewIssueReturn("STK-MTL", Guid.Empty, sCallProcess, oIssueReturnDataSet);
                    // get the index of the newly added isse material record
                    int iAddedRecordIndex = oIssueReturnDataSet.IssueReturn.Rows.Count - 1;
                    var row = oIssueReturnDataSet.IssueReturn[iAddedRecordIndex];

                    string sAvailTypes;
                    oIssueReturnImpl.GetAvailTranDocTypes(out sAvailTypes);

                    // set the core fields
                    row.TranType = "STK-MTL";
                    row.TranDate = DateTime.Now.Date;
                    row.ToJobNum = oMaterial.JobNum;
                    row.ToAssemblySeq = oMaterial.AssemblySequence;
                    row.ToJobSeq = oMaterial.MaterialSequence;
                    row.PartNum = oMaterial.PartNum;
                    row.FromWarehouseCode = oMaterial.FromWarehouseCode;
                    row.FromBinNum = oMaterial.FromWarehouseBin;
                    row.ToWarehouseCode = oMaterial.ToWarehouseCode;
                    row.ToBinNum = oMaterial.ToWarehouseBin;
                    row.QtyRequired = oMaterial.RemainingQuantity;
                    row.Company = oSession.CompanyID;

                    //oIssueReturnImpl.OnChangeTranType(ref oIssueReturnDataSet);

                    // change the job num
                    row.ToJobNum = oMaterial.JobNum;
                    oIssueReturnImpl.OnChangeToJobNum(oIssueReturnDataSet, sCallProcess, out sMessage);

                    // change the assembly seq
                    row.ToAssemblySeq = oMaterial.AssemblySequence;
                    oIssueReturnImpl.OnChangeToAssemblySeq(oIssueReturnDataSet, sCallProcess);

                    // change the material seq and part num
                    row.ToJobSeq = oMaterial.MaterialSequence;
                    row.PartNum = oMaterial.PartNum;
                    oIssueReturnImpl.OnChangeToJobSeq(oIssueReturnDataSet, sCallProcess, out sMessage);

                    // indicate the warehouse to issue the material from
                    row.FromWarehouseCode = oMaterial.FromWarehouseCode;
                    oIssueReturnImpl.OnChangeFromWarehouse(oIssueReturnDataSet, sCallProcess);

                    // indicate the bin to issue the material from
                    row.FromBinNum = oMaterial.FromWarehouseBin;
                    oIssueReturnImpl.OnChangeFromBinNum(oIssueReturnDataSet);

                    //indicate the amount to issue to the job
                    row.QtyRequired = oMaterial.RemainingQuantity;
                    oIssueReturnImpl.OnChangeTranQty(oMaterial.RemainingQuantity, oIssueReturnDataSet);

                    bool requiresUserInput = false;
                    bool bNegativeQuantityAction = false;
                    string sLegalNumberMessage;
                    string sPartTranPKs;

                    oIssueReturnImpl.PrePerformMaterialMovement(oIssueReturnDataSet, out requiresUserInput);
                    if (!requiresUserInput)
                    {
                        // set the transaction date
                        //row.TranDate = DateTime.Now.Date;
                        oIssueReturnImpl.PerformMaterialMovement(bNegativeQuantityAction, oIssueReturnDataSet, out sLegalNumberMessage, out sPartTranPKs);
                    }
                    else
                    {
                        Console.WriteLine("Manual intervention required (lot/serial/etc).");
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error issuing material: {ex.Message}");
            }
        }

        public List<JobOperationHierarchy> EvaluateOperationsToComplete(int iAssemblySequence, int iOperationSequence)
        {
            // get the ordered list of operations for this job
            List<JobOperationHierarchy> oOrderOperations = new List<JobOperationHierarchy>();
            // start at the root assembly and walk through all operations in order
            m_oRootAssembly.ListOperationsInOrder(oOrderOperations);

            List<JobOperationHierarchy> oOperationsToComplete = new List<JobOperationHierarchy>();
            // now we add these operations to the final list until we reach the one passed in
            foreach (JobOperationHierarchy oOperation in oOrderOperations)
            {
                oOperationsToComplete.Add(oOperation);
                if ((oOperation.AssemblySequence == iAssemblySequence) && (oOperation.OperationSequence == iOperationSequence))
                {
                    break;
                }
            }
            return oOperationsToComplete;
        }

        public JobOperationHierarchy NextOperationsToComplete()
        {
            // start at the root assembly and walk through all operations in order
            List<JobOperationHierarchy> oAllOperations = new List<JobOperationHierarchy>();
            m_oRootAssembly.ListOperationsInOrder(oAllOperations);
            JobOperationHierarchy oNextOperationToComplete = oAllOperations.FirstOrDefault(oItem => oItem.OperationComplete == false);
            return oNextOperationToComplete;
        }

        #endregion

        #region Properties

        public List<JobOperationHierarchy> JobOperations
        {
            get { return m_oJobOperations; }
            set { m_oJobOperations = value; }
        }

        public List<JobOperationHierarchy> CompletedJobOperations
        {
            get
            {
                return m_oCompletedJobOperations;
            }
        }

        public List<JobMaterialHierarchy> JobMaterials
        {
            get { return m_oJobMaterials; }
            set { m_oJobMaterials = value; }
        }

        public List<JobAssemblyHierarchy> JobAssemblies
        {
            get { return m_oJobAssemblies; }
            set { m_oJobAssemblies = value; }
        }

        public int TotalNumberOfOperations
        {
            get
            {
                return m_oJobOperations.Count;
            }
        }

        public decimal TotalEstimatedTime
        {
            get
            {
                decimal dTotalEstimatedTime = 0M;
                dTotalEstimatedTime = m_oJobOperations.Sum(oItem => oItem.EstimatedSetupHours + oItem.EstimatedProductionHours);
                return dTotalEstimatedTime;
            }
        }

        public decimal EstimatedTimeCompleted
        {
            get
            {
                decimal dEstimatedTimeCompleted = 0M;
                dEstimatedTimeCompleted = m_oCompletedJobOperations.Sum(oItem => oItem.EstimatedSetupHours + oItem.EstimatedProductionHours);
                return dEstimatedTimeCompleted;
            }
        }

        public int NumberOfOperationsCompleted
        {
            get
            {
                return m_oCompletedJobOperations.Count;
            }
        }

        public decimal PercentageCompleteByTimeEstiamtes
        {
            get
            {
                decimal dPercentageCompleteByTimeEstiamtes = 0M;
                decimal dTotalEstimatedTime = TotalEstimatedTime;
                decimal dEstimatedTimeCompleted = EstimatedTimeCompleted;
                if (dTotalEstimatedTime != 0)
                {
                    dPercentageCompleteByTimeEstiamtes = (dEstimatedTimeCompleted / dTotalEstimatedTime) * 100.0M;
                }
                return dPercentageCompleteByTimeEstiamtes;
            }
        }

        public decimal PercentageCompleteByOperationCount
        {
            get
            {
                decimal dPercentageCompleteByOperationCount = 0M;
                int iTotalOperationCount = TotalNumberOfOperations;
                int iNumberOfOperationsCompleted = NumberOfOperationsCompleted;
                if (iTotalOperationCount != 0)
                {
                    dPercentageCompleteByOperationCount = ((decimal)iNumberOfOperationsCompleted / (decimal)iTotalOperationCount) * 100.0M;
                }
                return dPercentageCompleteByOperationCount;
            }
        }
        #endregion

        #region Data Members
        private List<JobOperationHierarchy> m_oJobOperations = new List<JobOperationHierarchy>();
        private List<JobOperationHierarchy> m_oCompletedJobOperations = new List<JobOperationHierarchy>();
        private List<JobMaterialHierarchy> m_oJobMaterials = new List<JobMaterialHierarchy>();
        private List<JobAssemblyHierarchy> m_oJobAssemblies = new List<JobAssemblyHierarchy>();
        private List<EmployeeInfo> m_oEmployees = new List<EmployeeInfo>();
        private JobAssemblyHierarchy m_oRootAssembly;

        private List<AnalyzeOperationsToComplete> m_oChildJobs = new List<AnalyzeOperationsToComplete>();
        #endregion
    }
    //
    //
    //

}
