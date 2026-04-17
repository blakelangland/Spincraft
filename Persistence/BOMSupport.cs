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
    public class BOMSupport
    {
        #region constructors
        public BOMSupport(string sCompany)
        {
            m_sCompany = sCompany;

            if (string.Compare(m_sCompany, CompanyConfiguration.SPINCRAFT_WI_COMPANY_ID, true) == 0)
            {
                // Wisconsin ignores parts in the following classes
                m_oPartClassesToIgnore.Add("CATL");
                m_oPartClassesToIgnore.Add("COTL");
                m_oPartClassesToIgnore.Add("ENGD");
                m_oPartClassesToIgnore.Add("FA");
                m_oPartClassesToIgnore.Add("GOVT");
                m_oPartClassesToIgnore.Add("INSP");
                m_oPartClassesToIgnore.Add("LTAT");
                m_oPartClassesToIgnore.Add("MFG");
                m_oPartClassesToIgnore.Add("PAIN");
                m_oPartClassesToIgnore.Add("PUR");
                m_oPartClassesToIgnore.Add("SA");
                m_oPartClassesToIgnore.Add("SHIP");
                m_oPartClassesToIgnore.Add("SPNS");
                m_oPartClassesToIgnore.Add("SPTL");
                m_oPartClassesToIgnore.Add("SUPL");
                m_oPartClassesToIgnore.Add("WELD");

                // WI has no mfg parts thaht should be stocked
                m_oMfgPartsThatShouldBeStocked.Clear();
            }

            if (string.Compare(m_sCompany, CompanyConfiguration.SPINCRAFT_MA_COMPANY_ID, true) == 0)

            {
                m_oPartClassesToIgnore.Clear();

                // MA has no mfg parts that should be stocked
                m_oMfgPartsThatShouldBeStocked.Clear();

                m_sMfgPartsToInclude.Add("703");
                m_sMfgPartsToInclude.Add("787");
                m_sMfgPartsToInclude.Add("927");
                m_sMfgPartsToInclude.Add("1013");
                m_sMfgPartsToInclude.Add("1052");
                m_sMfgPartsToInclude.Add("1053");
                m_sMfgPartsToInclude.Add("1196");
                m_sMfgPartsToInclude.Add("2008");
                m_sMfgPartsToInclude.Add("2012");
                m_sMfgPartsToInclude.Add("2016");
                m_sMfgPartsToInclude.Add("2028");
                m_sMfgPartsToInclude.Add("2033");
                m_sMfgPartsToInclude.Add("2036");
                m_sMfgPartsToInclude.Add("2043");
                m_sMfgPartsToInclude.Add("2045");
                m_sMfgPartsToInclude.Add("2063");
                m_sMfgPartsToInclude.Add("2065");
                m_sMfgPartsToInclude.Add("2067");
                m_sMfgPartsToInclude.Add("2076");
                m_sMfgPartsToInclude.Add("2078");
                m_sMfgPartsToInclude.Add("2080");
                m_sMfgPartsToInclude.Add("2085");
                m_sMfgPartsToInclude.Add("2086");
                m_sMfgPartsToInclude.Add("2087");
                m_sMfgPartsToInclude.Add("2088");
                m_sMfgPartsToInclude.Add("2089"); 
                m_sMfgPartsToInclude.Add("2090"); 
                m_sMfgPartsToInclude.Add("9000");
            }
            
            if (string.Compare(m_sCompany, CompanyConfiguration.SPINCRAFT_UK_COMPANY_ID, true) == 0)
            {
                m_oPartClassesToIgnore.Clear();

                // mfg parts that should be stocked
                m_oMfgPartsThatShouldBeStocked.Clear();
                // medical parts
                m_oMfgPartsThatShouldBeStocked.Add("100322");
                m_oMfgPartsThatShouldBeStocked.Add("100297");
                m_oMfgPartsThatShouldBeStocked.Add("100299");
                m_oMfgPartsThatShouldBeStocked.Add("100302");
                m_oMfgPartsThatShouldBeStocked.Add("100323");
                m_oMfgPartsThatShouldBeStocked.Add("100310");
                m_oMfgPartsThatShouldBeStocked.Add("100313");
                m_oMfgPartsThatShouldBeStocked.Add("100303");
                m_oMfgPartsThatShouldBeStocked.Add("100305");
                m_oMfgPartsThatShouldBeStocked.Add("101183");
                m_oMfgPartsThatShouldBeStocked.Add("101182");
                // engineered parts
                m_oMfgPartsThatShouldBeStocked.Add("100261");
                m_oMfgPartsThatShouldBeStocked.Add("100264");
                m_oMfgPartsThatShouldBeStocked.Add("100267");
                // aviation
                m_oMfgPartsThatShouldBeStocked.Add("100334.");
                m_oMfgPartsThatShouldBeStocked.Add("100340");
                m_oMfgPartsThatShouldBeStocked.Add("100338");
                m_oMfgPartsThatShouldBeStocked.Add("100336");
                m_oMfgPartsThatShouldBeStocked.Add("100337");

                m_sMfgPartsToInclude.Clear();
            }
        }
        #endregion

        #region Methods
        public bool Initialize(Session oSession, HSValidateParts oValidateParts)
        {
            bool bSuccess = true;

            // loading all parts from the part master
            if (oValidateParts == null)
            {
                if (m_oValidateParts.Initialize(oSession) == false)
                {
                    Console.WriteLine("Failed to load the validate parts!");
                }
            }
            else
            {
                m_oValidateParts = oValidateParts;
            }

            // get a list of all manufactured parts in the system
            m_oMfgParts.Clear();
            Ice.Proxy.BO.DynamicQueryImpl oDynamicQuery = WCFServiceSupport.CreateImpl<Ice.Proxy.BO.DynamicQueryImpl>(oSession, Ice.Proxy.BO.DynamicQueryImpl.UriPath);
            QueryExecutionDataSet oQueryExecutionDataSet = oDynamicQuery.GetQueryExecutionParametersByID(BAQConstants.QUERY_LIST_ALL_MFG_PARTS_FOR_BOM_COMPARISON);
            DataSet oDataSet = oDynamicQuery.ExecuteByID(BAQConstants.QUERY_LIST_ALL_MFG_PARTS_FOR_BOM_COMPARISON, oQueryExecutionDataSet);
            foreach (DataRow oRow in oDataSet.Tables[0].Rows)
            {
                MfgPart oMfgPart = new MfgPart(oRow, m_oValidateParts);
                m_oMfgParts.Add(oMfgPart);
                List<MfgPart> oAllMfgParts = new List<MfgPart>();
                if (m_oParentMfgParts.ContainsKey(oMfgPart.PartNum) == true)
                {
                    oAllMfgParts = m_oParentMfgParts[oMfgPart.PartNum];
                }
                else
                {
                    m_oParentMfgParts[oMfgPart.PartNum] = oAllMfgParts;
                }
                oAllMfgParts.Add(oMfgPart);
            }

            // get all part materials in the system
            m_oMtlParts.Clear();
            oQueryExecutionDataSet = oDynamicQuery.GetQueryExecutionParametersByID(BAQConstants.LIST_ALL_PART_MTL);
            oDataSet = oDynamicQuery.ExecuteByID(BAQConstants.LIST_ALL_PART_MTL, oQueryExecutionDataSet);
            foreach (DataRow oRow in oDataSet.Tables[0].Rows)
            {
                PartMaterial oPartMaterial = new PartMaterial(oRow, m_oParentMfgParts, m_oValidateParts);
                List<PartMaterial> oAllFirstLevelParts = new List<PartMaterial>();
                string sKey = oPartMaterial.ParentPartNum + "-REV-" + oPartMaterial.ParentRevNum;
                if (m_oMtlParts.ContainsKey(sKey) == true)
                {
                    oAllFirstLevelParts = m_oMtlParts[sKey];
                }
                else
                {
                    m_oMtlParts[sKey] = oAllFirstLevelParts;
                }
                oAllFirstLevelParts.Add(oPartMaterial);
            }

            // get all operations in the system
            m_oOperationParts.Clear();
            oQueryExecutionDataSet = oDynamicQuery.GetQueryExecutionParametersByID(BAQConstants.QUERY_LIST_ALL_PART_OPERATIONS);
            oDataSet = oDynamicQuery.ExecuteByID(BAQConstants.QUERY_LIST_ALL_PART_OPERATIONS, oQueryExecutionDataSet);
            foreach (DataRow oRow in oDataSet.Tables[0].Rows)
            {
                PartOperation oPartOperation = new PartOperation(oRow);
                m_oOperationParts.Add(oPartOperation);
            }


            // now we will build up the BOMs and BOOs for the manufactured parts
            foreach (MfgPart oMfgPart in m_oMfgParts)
            {
                string sKey = oMfgPart.PartNum + "-REV-" + oMfgPart.RevNum;
                if (m_oMtlParts.ContainsKey(sKey) == true)
                {
                    // extract out all first level parts for this manufactured part
                    List<PartMaterial> oMaterials = m_oMtlParts[sKey];
                    foreach (PartMaterial oPartMtl in oMaterials)
                    {
                        oMfgPart.AddPartMaterial(oPartMtl);
                    }
                }
                // find all operations with the part part number and revision of this mfg part
                List<PartOperation> oMfgPartOperations = m_oOperationParts.Where(oItem => (string.Compare(oItem.PartNum, oMfgPart.PartNum, true) == 0) && (string.Compare(oItem.RevNum, oMfgPart.RevNum, true) == 0)).ToList();
                oMfgPart.AddPartOperations(oMfgPartOperations);
            }

            // now that we have built the BOMs we need to sort the material parts and operations
            foreach (MfgPart oMfgPart in m_oMfgParts)
            {
                oMfgPart.SortParts();
            }

            // we will filter out mfg parts that are not being used anymore if this is ETG MA
            if (string.Compare(m_sCompany, CompanyConfiguration.SPINCRAFT_MA_COMPANY_ID, true) == 0)
            {
                List<MfgPart> oFilteredList = new List<MfgPart>();
                foreach (MfgPart oMfgPart in m_oMfgParts)
                {
                    foreach (string sFilter in m_sMfgPartsToInclude)
                    {
                        if (oMfgPart.PartNum.StartsWith(sFilter) == true)
                        {
                            oFilteredList.Add(oMfgPart);
                            break;
                        }
                    }
                }
                m_oMfgParts = oFilteredList;
            }
            return bSuccess;
        }

        public MfgPart  GetPMfgPart(string sPartNum, string sPartRevNum)
        {
            MfgPart oFinishedGood = null;
            oFinishedGood = m_oMfgParts.FirstOrDefault(oItem => (string.Compare(oItem.PartNum, sPartNum, true) == 0) && (string.Compare(oItem.RevNum, sPartRevNum, true) == 0));
            return oFinishedGood;
        }

        public bool IsStockingPartWhiteListed(string sPartNum)
        {
            bool bIsWhiteListed = false;
            if (m_oMfgPartsThatShouldBeStocked.Contains(sPartNum) == true)
            {
                bIsWhiteListed = true;
            }
            return bIsWhiteListed;
        }

        public bool InitializeOperations(Session oSession)
        {
            bool bSuccess = true;

            // get a list of all operations in the system
            m_oOperations.Clear();
            Ice.Proxy.BO.DynamicQueryImpl oDynamicQuery = WCFServiceSupport.CreateImpl<Ice.Proxy.BO.DynamicQueryImpl>(oSession, Ice.Proxy.BO.DynamicQueryImpl.UriPath);
            QueryExecutionDataSet oQueryExecutionDataSet = oDynamicQuery.GetQueryExecutionParametersByID(BAQConstants.QUERY_LIST_ALL_OPERATIONS);
            DataSet oDataSet = oDynamicQuery.ExecuteByID(BAQConstants.QUERY_LIST_ALL_OPERATIONS, oQueryExecutionDataSet);
            foreach (DataRow oRow in oDataSet.Tables[0].Rows)
            {
                HSOperation oOperation = new HSOperation(oRow);
                m_oOperations.Add(oOperation);
            }

            return bSuccess;
        }

        public bool InitializeResourceGroups(Session oSession)
        {
            bool bSuccess = true;

            // get a list of all operations in the system
            m_oResourceGroups.Clear();
            Ice.Proxy.BO.DynamicQueryImpl oDynamicQuery = WCFServiceSupport.CreateImpl<Ice.Proxy.BO.DynamicQueryImpl>(oSession, Ice.Proxy.BO.DynamicQueryImpl.UriPath);
            QueryExecutionDataSet oQueryExecutionDataSet = oDynamicQuery.GetQueryExecutionParametersByID(BAQConstants.QUERY_LIST_ALL_RESOURCE_GROUPS);
            DataSet oDataSet = oDynamicQuery.ExecuteByID(BAQConstants.QUERY_LIST_ALL_RESOURCE_GROUPS, oQueryExecutionDataSet);
            foreach (DataRow oRow in oDataSet.Tables[0].Rows)
            {
                HSResourceGroup oResourceGroup = new HSResourceGroup(oRow);
                m_oResourceGroups.Add(oResourceGroup);
            }

            return bSuccess;
        }

        public bool InitializeResources(Session oSession)
        {
            bool bSuccess = true;

            // get a list of all operations in the system
            m_oResources.Clear();
            Ice.Proxy.BO.DynamicQueryImpl oDynamicQuery = WCFServiceSupport.CreateImpl<Ice.Proxy.BO.DynamicQueryImpl>(oSession, Ice.Proxy.BO.DynamicQueryImpl.UriPath);
            QueryExecutionDataSet oQueryExecutionDataSet = oDynamicQuery.GetQueryExecutionParametersByID(BAQConstants.QUERY_LIST_ALL_RESOURCES);
            DataSet oDataSet = oDynamicQuery.ExecuteByID(BAQConstants.QUERY_LIST_ALL_RESOURCES, oQueryExecutionDataSet);
            foreach (DataRow oRow in oDataSet.Tables[0].Rows)
            {
                HSResource oResource = new HSResource(oRow);
                m_oResources.Add(oResource);
            }

            return bSuccess;
        }

        public bool IgnorePartInClass(string sPartClass)
        {
            bool bIgnorePartClass = m_oPartClassesToIgnore.Contains(sPartClass);
            return bIgnorePartClass;
        }

        public bool IsResourceGroupActive(string sResourceGroupId)
        {
            bool bActive = false;
            HSResourceGroup oResourceGroup = m_oResourceGroups.FirstOrDefault(oItem => string.Compare(oItem.ResourceGroupId, sResourceGroupId, true) == 0);
            if ( (oResourceGroup != null) && (oResourceGroup.Inactive == false) )
            {
                bActive = true;
            }
            return bActive;
        }

        public bool IsResourceActive(string sResourceId)
        {
            bool bActive = false;
            HSResource oResource = m_oResources.FirstOrDefault(oItem => string.Compare(oItem.ResourceId, sResourceId, true) == 0);
            if ((oResource != null) && (oResource.Inactive == false))
            {
                bActive = true;
            }
            return bActive;
        }

        public void PerformValidation(string sCompany, string sTempFolder)
        {
            // we will walk through all mfg parts looking for issues

            // first we clear out all of our lists containing issues
            m_oMfgPartsWithoutMaterials.Clear();
            m_oMfgPartsWithInactiveMaterials.Clear();
            m_oMfgPartsWithMaterialsOnHold.Clear();
            m_oMfgPartsWithMaterialsOnRunOut.Clear();
            m_oMfgPartsWithDuplicatedMaterials.Clear();
            m_oMfgPartsWithMaterialsNotTiedToOperations.Clear();
            m_oMfgPartsWithoutOperations.Clear();
            m_oMfgPartsWithBadRevision.Clear();
            m_oMfgPartsWithNoApprovedRevision.Clear();
            m_oMfgPartsWithMultipleApprovedRevisions.Clear();
            m_oMfgPartsNotUsingPartRevs.Clear();
            m_oMfgPartsNotQuantityBearing.Clear();
            m_oMfgPartsWithIncosistentQuantityBearing.Clear();
            m_oMfgPartsWithIncorrectUOM.Clear();
            m_oMfgPartsWithInconsistentPhantomBOM.Clear();
            m_oMfgPartsWithInconsistentNonStock.Clear();
            m_oMfgPartsWithInconsistentTypeCode.Clear();
            m_oMfgPartsWithMaterialsWithZeroQty.Clear();
            m_oMfgPartsWithMaterialsWithZeroCost.Clear();
            m_oMfgPartsWithMRPTurnedOff.Clear();
            m_oMfgPartsWithBuyToOrderSet.Clear();
            m_oMfgPartsWithInconsistentBuyToOrder.Clear();
            m_oMfgPartsWithDropShipSet.Clear();
            m_oMfgPartsWithInconsistentDropShip.Clear();
            m_oMfgPartsWithoutPrimaryWarehouse.Clear();
            m_oMfgPartsWithoutGroupCode.Clear();
            m_oMfgPartsWithoutClassCode.Clear();
            m_oMfgPartsWithIncorrectCostMethod.Clear();
            m_oMfgPartsNeedingCostRoll.Clear();
            m_oMfgPartsWithInconsistentCostMethod.Clear();
            m_oMfgPartsSetToStocking.Clear();
            m_oMfgPartsStockedWithoutUnitPrice.Clear();
            m_oMfgPartsStockedWithoutMinOrSafetyLimits.Clear();

            m_oMfgPartsWithoutPortfolioCode.Clear();
            m_oMfgPartWithOperationsNotSetToEach.Clear();
            m_oMfgPartWithOperationsBadLaborEntry.Clear();
            m_oMfgPartWithOperationsBadStandardFormat.Clear();
            m_oMfgPartWithOperationsBadProductionStandard.Clear();
            m_oMfgPartWithOperationsBadOperationPerPartValue.Clear();
            m_oMfgPartWithNonZeroOperationsPerPartValue.Clear();
            m_oMfgPartWithOperationsWithAdditionalSetupQty.Clear();
            m_oMfgPartWithOperationsWithAdditionalSetupHours.Clear();
            m_oMfgPartWithOperationsWithResourceSpecified.Clear();
            m_oMfgPartSubcontractOperationMissingQtyPer.Clear();
            m_oMfgPartSubcontractOperationMissingDaysOut.Clear();
            m_oMfgPartSubcontractOperationMissingUnitCost.Clear();
            m_oMfgPartMaterialsWithFixedQty.Clear();
            m_oMfgPartMaterialsNeedViewAsAssemblyTurnedOn.Clear();
            m_oMfgPartMaterialsNeedViewAsAssemblyTurnedOff.Clear();
            m_oMfgPartMaterialsHasPlanAsAssemblyTurnedOn.Clear();

            // get the list of all active mfg parts
            List<MfgPart> oAllActiveMfgParts = m_oMfgParts.Where(oItem => (oItem.PartMaster != null) && (oItem.PartMaster.Inactive == false)).ToList();
            foreach (MfgPart oMfgPart in oAllActiveMfgParts)
            {
                //
                // no materials
                //
                if (oMfgPart.MyPartMaterials.Count == 0)
                {
                    m_oMfgPartsWithoutMaterials.Add(oMfgPart);
                }

                //
                // material parts inactive
                //
                List<PartMaterial> oInactiveParts = oMfgPart.MyPartMaterials.Where(oItem => (oItem.PartMaster != null) && (oItem.PartMaster.Inactive == true)).ToList();
                if (oInactiveParts.Count > 0)
                {
                    m_oMfgPartsWithInactiveMaterials.Add(oMfgPart);
                }

                //
                // materials on hold
                //
                List<PartMaterial> oPartsOnHold = oMfgPart.MyPartMaterials.Where(oItem => (oItem.PartMaster != null) && (oItem.PartMaster.PartOnHold == true)).ToList();
                if (oPartsOnHold.Count > 0)
                {
                    m_oMfgPartsWithMaterialsOnHold.Add(oMfgPart);
                }

                //
                // materials on run out
                //
                List<PartMaterial> oPartsOnRunOut = oMfgPart.MyPartMaterials.Where(oItem => (oItem.PartMaster != null) && (oItem.PartMaster.PartRunOut == true)).ToList();
                if (oPartsOnRunOut.Count > 0)
                {
                    m_oMfgPartsWithMaterialsOnRunOut.Add(oMfgPart);
                }

                //
                // duplicated materials
                //
                foreach (PartMaterial oPartMaterial in oMfgPart.MyPartMaterials)
                {
                    // we see if this part appears more than once in the list of materials
                    List<PartMaterial> oDuplicates = oMfgPart.MyPartMaterials.Where(oItem => string.Compare(oItem.MaterialPartNum, oPartMaterial.MaterialPartNum, true) == 0).ToList();
                    if (oDuplicates.Count > 1)
                    {
                        // this part appears more than once at this BOM level
                        m_oMfgPartsWithDuplicatedMaterials.Add(oMfgPart);
                        break;
                    }
                }

                //
                // materials not tied to operation
                //
                foreach (PartMaterial oPartMaterial in oMfgPart.MyPartMaterials)
                {
                    // we see if this part is tied to an operation 
                    List<PartMaterial> oNoOperation = oMfgPart.MyPartMaterials.Where(oItem => (oItem.RelatedOperation == 0) && (oItem.PartMaster != null) && (m_oPartClassesToIgnore.Contains(oItem.PartMaster.ClassID) == false)).ToList();
                    if (oNoOperation.Count > 1)
                    {
                        // this part is not tied to an operation
                        m_oMfgPartsWithMaterialsNotTiedToOperations.Add(oMfgPart);
                        break;
                    }
                }

                //
                // Mfg Parts without operations
                //
                if (oMfgPart.MyPartOperations.Count == 0)
                {
                    m_oMfgPartsWithoutOperations.Add(oMfgPart);
                }

                //
                // MfgPart has bad revision
                //
                if (string.IsNullOrEmpty(oMfgPart.RevNum) == true)
                {
                    // an Mfg Part with this name has not been added yet so we add it
                    m_oMfgPartsWithBadRevision.Add(oMfgPart);
                }

                //
                // MfgPart has no approved revision
                //
                // if this is an approved rev then we need to check if there are more approved revs for this part num
                List<MfgPart> oApprovedRevisions = oAllActiveMfgParts.Where(oItem => (string.Compare(oItem.PartNum, oMfgPart.PartNum, true) == 0) && (oItem.RevApproved == true)).ToList();
                if (oApprovedRevisions.Count == 0)
                {
                    // there are no approved revs for this part
                    MfgPart oTmp = m_oMfgPartsWithNoApprovedRevision.FirstOrDefault(oItem => string.Compare(oItem.PartNum, oMfgPart.PartNum, true) == 0);
                    if (oTmp == null)
                    {
                        // an Mfg Part with this name has not been added yet so we add it
                        m_oMfgPartsWithNoApprovedRevision.Add(oMfgPart);
                    }
                }

                //
                // MfgPart has multiple approved revisions
                //
                if (oMfgPart.RevApproved == true)
                {
                    // if this is an approved rev then we need to check if there are more approved revs for this part num
                    List<MfgPart> oMultipleApprovedRevisions = oAllActiveMfgParts.Where(oItem => (string.Compare(oItem.PartNum, oMfgPart.PartNum, true) == 0) && (oItem.RevApproved == true)).ToList();
                    if (oMultipleApprovedRevisions.Count > 1)
                    {
                        // mutiple approved revs for part
                        m_oMfgPartsWithMultipleApprovedRevisions.Add(oMfgPart);
                    }
                }

                //
                // MfgPart does not have Use Part Rev set
                //
                if ((oMfgPart.PartMaster != null) && (oMfgPart.PartMaster.UsePartRev == false))
                {
                    MfgPart oTmp = m_oMfgPartsNotUsingPartRevs.FirstOrDefault(oItem => string.Compare(oItem.PartNum, oMfgPart.PartNum, true) == 0);
                    if (oTmp == null)
                    {
                        // an Mfg Part with this name has not been added yet so we add it
                        m_oMfgPartsNotUsingPartRevs.Add(oMfgPart);
                    }
                }

                //
                // MfgPart does not have Qty Bearing set
                //
                if ((oMfgPart.PartMaster != null) && (oMfgPart.PartMaster.PartQtyBearing == false))
                {
                    MfgPart oTmp = m_oMfgPartsNotQuantityBearing.FirstOrDefault(oItem => string.Compare(oItem.PartNum, oMfgPart.PartNum, true) == 0);
                    if (oTmp == null)
                    {
                        // an Mfg Part with this name has not been added yet so we add it
                        m_oMfgPartsNotQuantityBearing.Add(oMfgPart);
                    }
                }

                //
                // MfgPart has inconsistent qty bearing set
                //
                if ((oMfgPart.PartMaster != null) && (oMfgPart.PartMaster.PartQtyBearing != oMfgPart.PartMaster.PlantQtyBearing))
                {
                    MfgPart oTmp = m_oMfgPartsWithIncosistentQuantityBearing.FirstOrDefault(oItem => string.Compare(oItem.PartNum, oMfgPart.PartNum, true) == 0);
                    if (oTmp == null)
                    {
                        // an Mfg Part with this name has not been added yet so we add it
                        m_oMfgPartsWithIncosistentQuantityBearing.Add(oMfgPart);
                    }
                }

                //
                // MfgPart has the incorrect UOM set
                //
                if ((oMfgPart.PartMaster != null) && (string.Compare("EA", oMfgPart.PartMaster.IUOM, true) != 0))
                {
                    MfgPart oTmp = m_oMfgPartsWithIncorrectUOM.FirstOrDefault(oItem => string.Compare(oItem.PartNum, oMfgPart.PartNum, true) == 0);
                    if (oTmp == null)
                    {
                        // an Mfg Part with this name has not been added yet so we add it
                        m_oMfgPartsWithIncorrectUOM.Add(oMfgPart);
                    }
                }

                //
                // should non-stock be on or should non-stock be off?????
                //

                //
                // inconsistent phantom BOM setting
                //
                if ((oMfgPart.PartMaster != null) && (oMfgPart.PartMaster.PartPhantomBOM != oMfgPart.PartMaster.PlantPhantomBOM))
                {
                    MfgPart oTmp = m_oMfgPartsWithInconsistentPhantomBOM.FirstOrDefault(oItem => string.Compare(oItem.PartNum, oMfgPart.PartNum, true) == 0);
                    if (oTmp == null)
                    {
                        // an Mfg Part with this name has not been added yet so we add it
                        m_oMfgPartsWithInconsistentPhantomBOM.Add(oMfgPart);
                    }
                }

                //
                // inconsistent non-stock setting
                //
                if ((oMfgPart.PartMaster != null) && (oMfgPart.PartMaster.PartNonStock != oMfgPart.PartMaster.PlantNonStockItem))
                {
                    MfgPart oTmp = m_oMfgPartsWithInconsistentNonStock.FirstOrDefault(oItem => string.Compare(oItem.PartNum, oMfgPart.PartNum, true) == 0);
                    if (oTmp == null)
                    {
                        // an Mfg Part with this name has not been added yet so we add it
                        m_oMfgPartsWithInconsistentNonStock.Add(oMfgPart);
                    }
                }

                //
                // inconsistent source type setting
                //
                if ((oMfgPart.PartMaster != null) && (string.Compare(oMfgPart.PartMaster.PartTypeCode, oMfgPart.PartMaster.PlantSourceType, true) != 0))
                {
                    MfgPart oTmp = m_oMfgPartsWithInconsistentTypeCode.FirstOrDefault(oItem => string.Compare(oItem.PartNum, oMfgPart.PartNum, true) == 0);
                    if (oTmp == null)
                    {
                        // an Mfg Part with this name has not been added yet so we add it
                        m_oMfgPartsWithInconsistentTypeCode.Add(oMfgPart);
                    }
                }

                //
                // part materials have zero qty
                //
                List<PartMaterial> oPartsWithZeroQuantity = oMfgPart.MyPartMaterials.Where(oItem => (oItem.QuantityPer == 0) && (oItem.PartMaster != null) && (m_oPartClassesToIgnore.Contains(oItem.PartMaster.ClassID) == false)).ToList();
                if (oPartsWithZeroQuantity.Count > 0)
                {
                    m_oMfgPartsWithMaterialsWithZeroQty.Add(oMfgPart);
                }

                //
                // part materials have zero cost
                //
                List<PartMaterial> oPartsWithZeroCost = oMfgPart.MyPartMaterials.Where(oItem => (oItem.PartMaster.Cost == 0) && (oItem.PartMaster != null) && (m_oPartClassesToIgnore.Contains(oItem.PartMaster.ClassID) == false)).ToList();
                if (oPartsWithZeroCost.Count > 0)
                {
                    m_oMfgPartsWithMaterialsWithZeroCost.Add(oMfgPart);
                }

                //
                // MRP turned off
                //
                if ((oMfgPart.PartMaster != null) && (oMfgPart.PartMaster.PlantProcessMRP == false))
                {
                    MfgPart oTmp = m_oMfgPartsWithMRPTurnedOff.FirstOrDefault(oItem => string.Compare(oItem.PartNum, oMfgPart.PartNum, true) == 0);
                    if (oTmp == null)
                    {
                        // an Mfg Part with this name has not been added yet so we add it
                        m_oMfgPartsWithMRPTurnedOff.Add(oMfgPart);
                    }
                }

                //
                // Buy To Order set
                //
                if ((oMfgPart.PartMaster != null) && (oMfgPart.PartMaster.PartBuyToOrder == true))
                {
                    MfgPart oTmp = m_oMfgPartsWithBuyToOrderSet.FirstOrDefault(oItem => string.Compare(oItem.PartNum, oMfgPart.PartNum, true) == 0);
                    if (oTmp == null)
                    {
                        // an Mfg Part with this name has not been added yet so we add it
                        m_oMfgPartsWithBuyToOrderSet.Add(oMfgPart);
                    }
                }

                //
                // inconsistent buy to order
                //
                if ((oMfgPart.PartMaster != null) && (oMfgPart.PartMaster.PartBuyToOrder != oMfgPart.PartMaster.PlantBuyToOrder))
                {
                    MfgPart oTmp = m_oMfgPartsWithInconsistentBuyToOrder.FirstOrDefault(oItem => string.Compare(oItem.PartNum, oMfgPart.PartNum, true) == 0);
                    if (oTmp == null)
                    {
                        // an Mfg Part with this name has not been added yet so we add it
                        m_oMfgPartsWithInconsistentBuyToOrder.Add(oMfgPart);
                    }
                }

                //
                // drop ship set
                //
                if ((oMfgPart.PartMaster != null) && (oMfgPart.PartMaster.PartDropShip == true))
                {
                    MfgPart oTmp = m_oMfgPartsWithDropShipSet.FirstOrDefault(oItem => string.Compare(oItem.PartNum, oMfgPart.PartNum, true) == 0);
                    if (oTmp == null)
                    {
                        // an Mfg Part with this name has not been added yet so we add it
                        m_oMfgPartsWithDropShipSet.Add(oMfgPart);
                    }
                }

                //
                // inconsistent drop ship set
                //
                if (oMfgPart.PartMaster.PartDropShip != oMfgPart.PartMaster.PlantDropShip)
                {
                    MfgPart oTmp = m_oMfgPartsWithInconsistentDropShip.FirstOrDefault(oItem => string.Compare(oItem.PartNum, oMfgPart.PartNum, true) == 0);
                    if (oTmp == null)
                    {
                        // an Mfg Part with this name has not been added yet so we add it
                        m_oMfgPartsWithInconsistentDropShip.Add(oMfgPart);
                    }
                }

                //
                // no primary warehouse set
                //
                if ((oMfgPart.PartMaster != null) && (string.IsNullOrEmpty(oMfgPart.PartMaster.PrimWhse) == true))
                {
                    MfgPart oTmp = m_oMfgPartsWithoutPrimaryWarehouse.FirstOrDefault(oItem => string.Compare(oItem.PartNum, oMfgPart.PartNum, true) == 0);
                    if (oTmp == null)
                    {
                        // an Mfg Part with this name has not been added yet so we add it
                        m_oMfgPartsWithoutPrimaryWarehouse.Add(oMfgPart);
                    }
                }

                //
                // no primary bin set
                //

                //
                // no group code set
                //
                if ((oMfgPart.PartMaster != null) && (string.IsNullOrEmpty(oMfgPart.PartMaster.Group) == true))
                {
                    MfgPart oTmp = m_oMfgPartsWithoutGroupCode.FirstOrDefault(oItem => string.Compare(oItem.PartNum, oMfgPart.PartNum, true) == 0);
                    if (oTmp == null)
                    {
                        // an Mfg Part with this name has not been added yet so we add it
                        m_oMfgPartsWithoutGroupCode.Add(oMfgPart);
                    }
                }

                //
                // no class code set
                //
                if ((oMfgPart.PartMaster != null) && (string.IsNullOrEmpty(oMfgPart.PartMaster.ClassID) == true))
                {
                    MfgPart oTmp = m_oMfgPartsWithoutClassCode.FirstOrDefault(oItem => string.Compare(oItem.PartNum, oMfgPart.PartNum, true) == 0);
                    if (oTmp == null)
                    {
                        // an Mfg Part with this name has not been added yet so we add it
                        m_oMfgPartsWithoutClassCode.Add(oMfgPart);
                    }
                }

                //
                // check costing method -- could be different for each site
                //
                if (string.Compare(m_sCompany, CompanyConfiguration.SPINCRAFT_MA_COMPANY_ID, true) == 0)
                {
                    // ETG MA SHOULD BE LAST
                    if ((oMfgPart.PartMaster != null) && (string.Compare(oMfgPart.PartMaster.PartCostingMethod, "L", true) != 0))
                    {
                        MfgPart oTmp = m_oMfgPartsWithIncorrectCostMethod.FirstOrDefault(oItem => string.Compare(oItem.PartNum, oMfgPart.PartNum, true) == 0);
                        if (oTmp == null)
                        {
                            // an Mfg Part with this name has not been added yet so we add it
                            m_oMfgPartsWithIncorrectCostMethod.Add(oMfgPart);
                        }
                    }
                }
                else if (string.Compare(m_sCompany, CompanyConfiguration.SPINCRAFT_WI_COMPANY_ID, true) == 0)
                {
                    // ETG WI SHOULD BE AVERAGE OR STANDARD
                    if ((oMfgPart.PartMaster != null) && (string.Compare(oMfgPart.PartMaster.PartCostingMethod, "A", true) != 0) && (string.Compare(oMfgPart.PartMaster.PartCostingMethod, "S", true) != 0))
                    {
                        MfgPart oTmp = m_oMfgPartsWithIncorrectCostMethod.FirstOrDefault(oItem => string.Compare(oItem.PartNum, oMfgPart.PartNum, true) == 0);
                        if (oTmp == null)
                        {
                            // an Mfg Part with this name has not been added yet so we add it
                            m_oMfgPartsWithIncorrectCostMethod.Add(oMfgPart);
                        }
                    }
                }
                else if (string.Compare(m_sCompany, CompanyConfiguration.SPINCRAFT_UK_COMPANY_ID, true) == 0)
                {
                    // ETG UK SHOULD BE AVERAGE OR LAST
                    if ((oMfgPart.PartMaster != null) && (string.Compare(oMfgPart.PartMaster.PartCostingMethod, "F", true) != 0))
                    {
                        MfgPart oTmp = m_oMfgPartsWithIncorrectCostMethod.FirstOrDefault(oItem => string.Compare(oItem.PartNum, oMfgPart.PartNum, true) == 0);
                        if (oTmp == null)
                        {
                            // an Mfg Part with this name has not been added yet so we add it
                            m_oMfgPartsWithIncorrectCostMethod.Add(oMfgPart);
                        }
                    }
                }

                //
                // check for std cost method where std cost is zero -- no cost roll done
                //
                if ((oMfgPart.PartMaster != null) && (string.Compare(oMfgPart.PartMaster.PartCostingMethod, "S", true) == 0) && (oMfgPart.CostPerUnit == 0))
                {
                    MfgPart oTmp = m_oMfgPartsNeedingCostRoll.FirstOrDefault(oItem => string.Compare(oItem.PartNum, oMfgPart.PartNum, true) == 0);
                    if (oTmp == null)
                    {
                        // an Mfg Part with this name has not been added yet so we add it
                        m_oMfgPartsNeedingCostRoll.Add(oMfgPart);
                    }
                }

                //
                // inconsistent costing method
                //
                if ((oMfgPart.PartMaster != null) && (string.Compare(oMfgPart.PartMaster.PartCostingMethod, oMfgPart.PartMaster.PlantCostingMethod, true) != 0))
                {
                    MfgPart oTmp = m_oMfgPartsWithInconsistentCostMethod.FirstOrDefault(oItem => string.Compare(oItem.PartNum, oMfgPart.PartNum, true) == 0);
                    if (oTmp == null)
                    {
                        // an Mfg Part with this name has not been added yet so we add it
                        m_oMfgPartsWithInconsistentCostMethod.Add(oMfgPart);
                    }
                }

                //
                // mfg parts set to be stocking -- all parts should be make direct unless they are white listed
                //
                if ((oMfgPart.PartMaster != null) && (oMfgPart.PartMaster.PartNonStock == false))
                {
                    // check if part is white listed
                    if (IsStockingPartWhiteListed(oMfgPart.PartNum) == false)
                    {
                        MfgPart oTmp = m_oMfgPartsSetToStocking.FirstOrDefault(oItem => string.Compare(oItem.PartNum, oMfgPart.PartNum, true) == 0);
                        if (oTmp == null)
                        {
                            // an Mfg Part with this name has not been added yet so we add it
                            m_oMfgPartsSetToStocking.Add(oMfgPart);
                        }
                    }
                }

                //
                // mfg part that is set to be stocked but that has no price (margin for POC cant be computed)
                //
                if ((oMfgPart.PartMaster != null) && (oMfgPart.PartMaster.PartNonStock == false) && (oMfgPart.PartMaster.UnitPrice == 0))
                {
                    MfgPart oTmp = m_oMfgPartsStockedWithoutUnitPrice.FirstOrDefault(oItem => string.Compare(oItem.PartNum, oMfgPart.PartNum, true) == 0);
                    if (oTmp == null)
                    {
                        // an Mfg Part with this name has not been added yet so we add it
                        m_oMfgPartsStockedWithoutUnitPrice.Add(oMfgPart);
                    }
                }

                //
                // mfg part that is set to be stocked but that has no min or safety limits set
                //
                if ((oMfgPart.PartMaster != null) && (oMfgPart.PartMaster.PartNonStock == false) && (oMfgPart.PartMaster.Minimum == 0) && (oMfgPart.PartMaster.Safety == 0))
                {
                    MfgPart oTmp = m_oMfgPartsStockedWithoutMinOrSafetyLimits.FirstOrDefault(oItem => string.Compare(oItem.PartNum, oMfgPart.PartNum, true) == 0);
                    if (oTmp == null)
                    {
                        // an Mfg Part with this name has not been added yet so we add it
                        m_oMfgPartsStockedWithoutMinOrSafetyLimits.Add(oMfgPart);
                    }
                }

                //
                // no portfolio code set
                //
                if ((oMfgPart.PartMaster != null) && (string.IsNullOrEmpty(oMfgPart.PartMaster.ProductPortfolio) == true))
                {
                    MfgPart oTmp = m_oMfgPartsWithoutPortfolioCode.FirstOrDefault(oItem => string.Compare(oItem.PartNum, oMfgPart.PartNum, true) == 0);
                    if (oTmp == null)
                    {
                        // an Mfg Part with this name has not been added yet so we add it
                        m_oMfgPartsWithoutPortfolioCode.Add(oMfgPart);
                    }
                }

                //
                // OPERATION ADDITIONAL CHECKS
                //

                //
                // operation std basis shouls be "E" unless subcontract in which case it is blank
                //
                List<PartOperation> oOperationsNotSetToEach = oMfgPart.MyPartOperations.Where(oItem => (string.Compare(oItem.StandardBasis, "E", true) != 0) && (oItem.Subcontract == false)).ToList();
                if (oOperationsNotSetToEach.Count > 0)
                {
                    m_oMfgPartWithOperationsNotSetToEach.Add(oMfgPart);
                }

                // labor entry should be Time and Quantity -- "T"
                List<PartOperation> oOperationsWithBadLaborEntry = oMfgPart.MyPartOperations.Where(oItem => (string.Compare(oItem.LaborEntryMethod, "T", true) != 0)).ToList();
                if (oOperationsWithBadLaborEntry.Count > 0)
                {
                    m_oMfgPartWithOperationsBadLaborEntry.Add(oMfgPart);
                }

                // standard format should be "HP" hours / piece
                List<PartOperation> oOperationsWithBadStandardFormat = new List<PartOperation>();
                if (string.Compare(m_sCompany, CompanyConfiguration.SPINCRAFT_MA_COMPANY_ID, true) == 0)
                {
                    // for MA it should always be HP
                    oOperationsWithBadStandardFormat = oMfgPart.MyPartOperations.Where(oItem => (string.Compare(oItem.StandardFormat, "HP", true) != 0)).ToList();
                    if (oOperationsWithBadStandardFormat.Count > 0)
                    {
                        m_oMfgPartWithOperationsBadStandardFormat.Add(oMfgPart);
                    }
                }
                else if (string.Compare(m_sCompany, CompanyConfiguration.SPINCRAFT_WI_COMPANY_ID, true) == 0)
                {
                    // for WI it should always be HP
                    oOperationsWithBadStandardFormat = oMfgPart.MyPartOperations.Where(oItem => (string.Compare(oItem.StandardFormat, "HP", true) != 0)).ToList();
                    if (oOperationsWithBadStandardFormat.Count > 0)
                    {
                        m_oMfgPartWithOperationsBadStandardFormat.Add(oMfgPart);
                    }
                }
                else if (string.Compare(m_sCompany, CompanyConfiguration.SPINCRAFT_UK_COMPANY_ID, true) == 0)
                {
                    // for UK it should be either HP or MP
                    oOperationsWithBadStandardFormat = oMfgPart.MyPartOperations.Where(oItem => (string.Compare(oItem.StandardFormat, "HP", true) != 0) && (string.Compare(oItem.StandardFormat, "MP", true) != 0)).ToList();
                    if (oOperationsWithBadStandardFormat.Count > 0)
                    {
                        m_oMfgPartWithOperationsBadStandardFormat.Add(oMfgPart);
                    }
                }


                // the production standard should be non-zero
                List<PartOperation> oOperationsWithBadProductionStandard = oMfgPart.MyPartOperations.Where(oItem => (oItem.ProductionStandard == 0) && (oItem.Subcontract == false)).ToList();
                if (oOperationsWithBadProductionStandard.Count > 0)
                {
                    m_oMfgPartWithOperationsBadProductionStandard.Add(oMfgPart);
                }

                // if std format is "OM" or "OH" then the OpsPerPart field must be > 0
                List<PartOperation> oOperationsMissingOperationsPerPartValue = oMfgPart.MyPartOperations.Where(oItem => ((string.Compare(oItem.StandardFormat, "OM", true) == 0) || (string.Compare(oItem.StandardFormat, "OH", true) == 0)) && (oItem.OperationsPerPart == 0)).ToList();
                if (oOperationsMissingOperationsPerPartValue.Count > 0)
                {
                    m_oMfgPartWithOperationsBadOperationPerPartValue.Add(oMfgPart);
                }

                // if std format is anything other than "OM" or "OH" then the OpsPerPart field should be zero
                List<PartOperation> oOperationsNonZeroOperationsPerPartValue = oMfgPart.MyPartOperations.Where(oItem => (string.Compare(oItem.StandardFormat, "OM", true) != 0) && (string.Compare(oItem.StandardFormat, "OH", true) != 0) && (oItem.OperationsPerPart != 0)).ToList();
                if (oOperationsNonZeroOperationsPerPartValue.Count > 0)
                {
                    m_oMfgPartWithNonZeroOperationsPerPartValue.Add(oMfgPart);
                }

                // additional setup qty should be zero
                List<PartOperation> oOperationsWithAdditionalSetupQuantity = oMfgPart.MyPartOperations.Where(oItem => oItem.AdditionalSetQty != 0).ToList();
                if (oOperationsWithAdditionalSetupQuantity.Count > 0)
                {
                    m_oMfgPartWithOperationsWithAdditionalSetupQty.Add(oMfgPart);
                }

                // additional setup hours should be zero
                List<PartOperation> oOperationsWithAdditionalSetupHours = oMfgPart.MyPartOperations.Where(oItem => oItem.AdditonalSetupHours != 0).ToList();
                if (oOperationsWithAdditionalSetupHours.Count > 0)
                {
                    m_oMfgPartWithOperationsWithAdditionalSetupHours.Add(oMfgPart);
                }

                // we should not specify a resource id on the operation -- too specific
                List<PartOperation> oOperationsWithResourceSet = oMfgPart.MyPartOperations.Where(oItem => string.IsNullOrEmpty(oItem.ResourceId) == false).ToList();
                if (oOperationsWithResourceSet.Count > 0)
                {
                    m_oMfgPartWithOperationsWithResourceSpecified.Add(oMfgPart);
                }

                // auto-recieve into inventory should be false

                // if this is a subcontract we should include the qty/per
                List<PartOperation> oSubcontractOperationMissingQtyPer = oMfgPart.MyPartOperations.Where(oItem => (oItem.Subcontract == true) && (oItem.QtyPerParent == 0)).ToList();
                if (oSubcontractOperationMissingQtyPer.Count > 0)
                {
                    m_oMfgPartSubcontractOperationMissingQtyPer.Add(oMfgPart);
                }

                // if this is a subcontract we should include the days out
                List<PartOperation> oSubcontractOperationMissingDaysOut = oMfgPart.MyPartOperations.Where(oItem => (oItem.Subcontract == true) && (oItem.DaysOut == 0)).ToList();
                if (oSubcontractOperationMissingDaysOut.Count > 0)
                {
                    m_oMfgPartSubcontractOperationMissingDaysOut.Add(oMfgPart);
                }

                // if this is a subcontract we should include the unit cost
                List<PartOperation> oSubcontractOperationMissingUnitCost = oMfgPart.MyPartOperations.Where(oItem => (oItem.Subcontract == true) && (oItem.EstUnitCost == 0)).ToList();
                if (oSubcontractOperationMissingUnitCost.Count > 0)
                {
                    m_oMfgPartSubcontractOperationMissingUnitCost.Add(oMfgPart);
                }

                // should we check to see where we are using price breaks?
                // should we check to see where we are using production standard breaks?

                // should we verify that where the operation is set to scrap, the materials on the operation have the corresponding scrap factor in place?


                //
                // MATERIAL ADDITIONAL CHECKS
                //

                // report when fixed qty is set
                List<PartMaterial> oFixQtyPartMaterials = oMfgPart.MyPartMaterials.Where(oItem => (oItem.FixedQty == true) && (oItem.PartMaster != null) && (m_oPartClassesToIgnore.Contains(oItem.PartMaster.ClassID) == false)).ToList();
                if (oFixQtyPartMaterials.Count > 0)
                {
                    m_oMfgPartMaterialsWithFixedQty.Add(oMfgPart);
                }

                // check if scrap is consistent with the operation scrap?

                // if material is an M then View as Assembly should be turned on
                List<PartMaterial> oMaterialIsManufacturedButNotSetToViewAsAssembly = oMfgPart.MyPartMaterials.Where(oItem => (string.Compare(oItem.TypeCode, "M", true) == 0) && (oItem.ViewAsAsm == false)).ToList();
                if (oMaterialIsManufacturedButNotSetToViewAsAssembly.Count > 0)
                {
                    m_oMfgPartMaterialsNeedViewAsAssemblyTurnedOn.Add(oMfgPart);
                }

                // if material is an P then View as Assembly should be turned off
                List<PartMaterial> oMaterialShouldNotSetToViewAsAssembly = oMfgPart.MyPartMaterials.Where(oItem => (string.Compare(oItem.TypeCode, "P", true) == 0) && (oItem.ViewAsAsm == true)).ToList();
                if (oMaterialShouldNotSetToViewAsAssembly.Count > 0)
                {
                    m_oMfgPartMaterialsNeedViewAsAssemblyTurnedOff.Add(oMfgPart);
                }

                // report if we use plan as assembly
                List<PartMaterial> oMaterialIsSetToPlanAsAssembly = oMfgPart.MyPartMaterials.Where(oItem => oItem.PlanAsAsm == true).ToList();
                if (oMaterialIsSetToPlanAsAssembly.Count > 0)
                {
                    m_oMfgPartMaterialsHasPlanAsAssemblyTurnedOn.Add(oMfgPart);
                }
            }
            CreateReport(sCompany, sTempFolder);
        }

        public void CreateReport(string sCompany, string sTmpFileDirectory)
        {
            // get the file name
            DateTime dtToday = DateTime.Now;
            string sDate = dtToday.Month.ToString() + "-" + dtToday.Day.ToString() + "-" + dtToday.Year.ToString();
            string sDestinationFileName = sTmpFileDirectory + "\\" + sCompany + "-MfgPartBOMReport-" + sDate + ".xlsx";
            int iNumberOfRows = 1;
            int iNumberOfColumns = 1;

            HSUser oServiceAccount = HSUser.GetUserById(HSUser.SPINCRAFT_SERVICE_ACCOUNT_ID);
            List<string> oToAddresses = new List<string>();

            SLFill oSLFill = new SLFill();
            oSLFill.SetPatternBackgroundColor(SLThemeColorIndexValues.Accent1Color);
            oSLFill.SetPattern(PatternValues.Solid, SLThemeColorIndexValues.Light2Color, SLThemeColorIndexValues.Light2Color);

            SLStyle oBoldStyle = new SLStyle();
            oBoldStyle.SetFontBold(true);
            oBoldStyle.SetTopBorder(BorderStyleValues.Thin, SLThemeColorIndexValues.Dark2Color);
            oBoldStyle.SetBottomBorder(BorderStyleValues.Thin, SLThemeColorIndexValues.Dark2Color);
            oBoldStyle.SetLeftBorder(BorderStyleValues.Thin, SLThemeColorIndexValues.Dark2Color);
            oBoldStyle.SetRightBorder(BorderStyleValues.Thin, SLThemeColorIndexValues.Dark2Color);
            oBoldStyle.SetFontColor(SLThemeColorIndexValues.Dark2Color);

            SLStyle oBoldStyle2 = new SLStyle();
            oBoldStyle2.SetFontBold(true);
            oBoldStyle2.SetTopBorder(BorderStyleValues.Thick, SLThemeColorIndexValues.Dark2Color);
            oBoldStyle2.SetBottomBorder(BorderStyleValues.Thick, SLThemeColorIndexValues.Dark2Color);
            oBoldStyle2.SetLeftBorder(BorderStyleValues.Thick, SLThemeColorIndexValues.Dark2Color);
            oBoldStyle.SetRightBorder(BorderStyleValues.Thick, SLThemeColorIndexValues.Dark2Color);
            oBoldStyle2.SetFontColor(System.Drawing.Color.IndianRed);
            oBoldStyle2.Fill = oSLFill;

            System.Drawing.Color oHighlightYellow = System.Drawing.Color.FromArgb(255, 255, 0);

            SLStyle oHighlightHeaderStyle = new SLStyle();
            oHighlightHeaderStyle.SetFontBold(true);
            oHighlightHeaderStyle.SetFont(FontSchemeValues.Major, 12);
            oHighlightHeaderStyle.SetHorizontalAlignment(HorizontalAlignmentValues.Center);
            oHighlightHeaderStyle.SetPatternFill(PatternValues.Solid, oHighlightYellow, oHighlightYellow);
            oHighlightHeaderStyle.SetFontColor(SLThemeColorIndexValues.Dark2Color);

            #region BOM Issues
            //
            // BOM management issues
            //
            oToAddresses = HSUser.GetEmailsForUsersInGroup(HSUser.REPORT_ON_MANUFACTURED_PART_ISSUES);
            oToAddresses.Add(HSUser.SPINCRAFT_ROOT_USER);
            bool bDataInReport = false;
            bool bFirstWorksheet = true;

            // if the file is already there then get rid of it as we are reprocessing for some reason
            if (File.Exists(sDestinationFileName) == true)
            {
                try
                {
                    File.Delete(sDestinationFileName);
                }
                catch (Exception)
                {
                    // we will ingore this if we cannot delete the file
                }
            }
            SLDocument oSLBOMDocument = new SLDocument();

            // missing materials
            iNumberOfRows = 1;
            iNumberOfColumns = 1;
            if (m_oMfgPartsWithoutMaterials.Count != 0)
            {
                if (bFirstWorksheet == true)
                {
                    oSLBOMDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "Missing Materials");
                    bFirstWorksheet = false;
                }
                else
                {
                    oSLBOMDocument.AddWorksheet("Missing Materials");
                }
                //set column header
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "PartNum");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 30);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Revision");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 30);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Description");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 60);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "The Mfg Part Has No Materials In the BOM");
                oSLBOMDocument.SetCellStyle(iNumberOfRows++, iNumberOfColumns, oHighlightHeaderStyle);
                oSLBOMDocument.SetColumnWidth(iNumberOfColumns++, 90);

                foreach (MfgPart oMfgPart in m_oMfgPartsWithoutMaterials)
                {
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 1, oMfgPart.PartNum);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 2, oMfgPart.RevNum);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 3, StringExt.Truncate(oMfgPart.Description, 50));

                    iNumberOfRows++;
                    bDataInReport = true;
                }
            }

            // inactive materials
            iNumberOfRows = 1;
            iNumberOfColumns = 1;
            if (m_oMfgPartsWithInactiveMaterials.Count != 0)
            {
                if (bFirstWorksheet == true)
                {
                    oSLBOMDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "Inactive Materials");
                    bFirstWorksheet = false;
                }
                else
                {
                    oSLBOMDocument.AddWorksheet("Inactive Materials");
                }
                //set column header
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Parent Part");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 30);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Revision");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 30);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Inactive Part");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 30);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "The Mfg Part Has Materials In The BOM That Are Inactive");
                oSLBOMDocument.SetCellStyle(iNumberOfRows++, iNumberOfColumns, oHighlightHeaderStyle);
                oSLBOMDocument.SetColumnWidth(iNumberOfColumns++, 90);

                foreach (MfgPart oMfgPart in m_oMfgPartsWithInactiveMaterials)
                {
                    List<PartMaterial> oInactiveParts = oMfgPart.MyPartMaterials.Where(oItem => oItem.PartMaster.Inactive == true).ToList();
                    foreach (PartMaterial oPartMaterial in oInactiveParts)
                    {
                        oSLBOMDocument.SetCellValue(iNumberOfRows, 1, oMfgPart.PartNum);
                        oSLBOMDocument.SetCellValue(iNumberOfRows, 2, oMfgPart.RevNum);
                        oSLBOMDocument.SetCellValue(iNumberOfRows, 3, oPartMaterial.PartMaster.PartNum);

                        iNumberOfRows++;
                        bDataInReport = true;
                    }
                }
            }

            // materials on hold
            iNumberOfRows = 1;
            iNumberOfColumns = 1;
            if (m_oMfgPartsWithMaterialsOnHold.Count != 0)
            {
                if (bFirstWorksheet == true)
                {
                    oSLBOMDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "Mtl On Hold");
                    bFirstWorksheet = false;
                }
                else
                {
                    oSLBOMDocument.AddWorksheet("Mtl On Hold");
                }
                //set column header
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Parent Part");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 30);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Revision");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 30);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Part On Hold");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 30);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "The Mfg Part Has Materials In The BOM That Are On Hold");
                oSLBOMDocument.SetCellStyle(iNumberOfRows++, iNumberOfColumns, oHighlightHeaderStyle);
                oSLBOMDocument.SetColumnWidth(iNumberOfColumns++, 90);

                foreach (MfgPart oMfgPart in m_oMfgPartsWithMaterialsOnHold)
                {
                    List<PartMaterial> oPartsOnHold = oMfgPart.MyPartMaterials.Where(oItem => oItem.PartMaster.PartOnHold == true).ToList();
                    foreach (PartMaterial oPartMaterial in oPartsOnHold)
                    {
                        oSLBOMDocument.SetCellValue(iNumberOfRows, 1, oMfgPart.PartNum);
                        oSLBOMDocument.SetCellValue(iNumberOfRows, 2, oMfgPart.RevNum);
                        oSLBOMDocument.SetCellValue(iNumberOfRows, 3, oPartMaterial.PartMaster.PartNum);

                        iNumberOfRows++;
                        bDataInReport = true;
                    }
                }
            }

            // materials on run out
            iNumberOfRows = 1;
            iNumberOfColumns = 1;
            if (m_oMfgPartsWithMaterialsOnRunOut.Count != 0)
            {
                if (bFirstWorksheet == true)
                {
                    oSLBOMDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "Mtl On Run Out");
                    bFirstWorksheet = false;
                }
                else
                {
                    oSLBOMDocument.AddWorksheet("Mtl On Run Out");
                }
                //set column header
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Parent Part");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 30);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Revision");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 30);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Part On Run Out");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 30);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "The Mfg Part Has Materials In The BOM That Are On Run Out");
                oSLBOMDocument.SetCellStyle(iNumberOfRows++, iNumberOfColumns, oHighlightHeaderStyle);
                oSLBOMDocument.SetColumnWidth(iNumberOfColumns++, 90);

                foreach (MfgPart oMfgPart in m_oMfgPartsWithMaterialsOnRunOut)
                {
                    List<PartMaterial> oPartsOnRunOut = oMfgPart.MyPartMaterials.Where(oItem => oItem.PartMaster.PartRunOut == true).ToList();
                    foreach (PartMaterial oPartMaterial in oPartsOnRunOut)
                    {
                        oSLBOMDocument.SetCellValue(iNumberOfRows, 1, oMfgPart.PartNum);
                        oSLBOMDocument.SetCellValue(iNumberOfRows, 2, oMfgPart.RevNum);
                        oSLBOMDocument.SetCellValue(iNumberOfRows, 3, oPartMaterial.PartMaster.PartNum);

                        iNumberOfRows++;
                        bDataInReport = true;
                    }
                }
            }

            iNumberOfRows = 1;
            iNumberOfColumns = 1;
            if (m_oMfgPartsWithDuplicatedMaterials.Count != 0)
            {
                if (bFirstWorksheet == true)
                {
                    oSLBOMDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "Duplicated Mtl");
                    bFirstWorksheet = false;
                }
                else
                {
                    oSLBOMDocument.AddWorksheet("Duplicated Mtl");
                }
                //set column header
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Parent Part");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 30);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Revision");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 30);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Duplicated Part");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 30);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "The Mfg Part Has Materials In The BOM That Are Duplicated");
                oSLBOMDocument.SetCellStyle(iNumberOfRows++, iNumberOfColumns, oHighlightHeaderStyle);
                oSLBOMDocument.SetColumnWidth(iNumberOfColumns++, 90);

                foreach (MfgPart oMfgPart in m_oMfgPartsWithDuplicatedMaterials)
                {
                    List<string> oDupParts = new List<string>();
                    foreach (PartMaterial oMaterial in oMfgPart.MyPartMaterials)
                    {
                        // we see if this mfg part has materials that appear more than once in the list
                        List<PartMaterial> oDuplicates = oMfgPart.MyPartMaterials.Where(oItem => string.Compare(oItem.MaterialPartNum, oMaterial.MaterialPartNum, true) == 0).ToList();
                        if (oDuplicates.Count > 1)
                        {
                            if (oDupParts.Contains(oMaterial.MaterialPartNum) == false)
                            {
                                // add it to our list so we only report this once
                                oDupParts.Add(oMaterial.MaterialPartNum);
                            }
                        }
                    }
                    //now we write out the duplicated parts for this BOM
                    foreach (string sDuplcatedPart in oDupParts)
                    {
                        // put this in the spreadsheet
                        oSLBOMDocument.SetCellValue(iNumberOfRows, 1, oMfgPart.PartNum);
                        oSLBOMDocument.SetCellValue(iNumberOfRows, 2, oMfgPart.RevNum);
                        oSLBOMDocument.SetCellValue(iNumberOfRows, 3, sDuplcatedPart);

                        iNumberOfRows++;
                        bDataInReport = true;
                    }
                }
            }

            // materials not tied to an operation
            iNumberOfRows = 1;
            iNumberOfColumns = 1;
            if (m_oMfgPartsWithMaterialsNotTiedToOperations.Count != 0)
            {
                if (bFirstWorksheet == true)
                {
                    oSLBOMDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "Mtl No Operation");
                    bFirstWorksheet = false;
                }
                else
                {
                    oSLBOMDocument.AddWorksheet("Mtl No Operation");
                }
                //set column header
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Parent Part");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 30);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Revision");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 30);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Mtl No Op");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 30);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "The Mfg Part Has Materials In The BOM Not Assigned To An Operation");
                oSLBOMDocument.SetCellStyle(iNumberOfRows++, iNumberOfColumns, oHighlightHeaderStyle);
                oSLBOMDocument.SetColumnWidth(iNumberOfColumns++, 90);

                foreach (MfgPart oMfgPart in m_oMfgPartsWithMaterialsNotTiedToOperations)
                {
                    List<PartMaterial> oPartsWithoutOperation = oMfgPart.MyPartMaterials.Where(oItem => oItem.RelatedOperation == 0).ToList();
                    foreach (PartMaterial oPartMaterial in oPartsWithoutOperation)
                    {
                        oSLBOMDocument.SetCellValue(iNumberOfRows, 1, oMfgPart.PartNum);
                        oSLBOMDocument.SetCellValue(iNumberOfRows, 2, oMfgPart.RevNum);
                        oSLBOMDocument.SetCellValue(iNumberOfRows, 3, oPartMaterial.PartMaster.PartNum);

                        iNumberOfRows++;
                        bDataInReport = true;
                    }
                }
            }

            // mfg part without operations
            iNumberOfRows = 1;
            iNumberOfColumns = 1;
            if (m_oMfgPartsWithoutOperations.Count != 0)
            {
                if (bFirstWorksheet == true)
                {
                    oSLBOMDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "Mfg Part No Operations");
                    bFirstWorksheet = false;
                }
                else
                {
                    oSLBOMDocument.AddWorksheet("Mfg Part No Operations");
                }
                //set column header
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Part");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 30);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Revision");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 30);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Description");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 60);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "The Mfg Part Has No Operations In The BOM");
                oSLBOMDocument.SetCellStyle(iNumberOfRows++, iNumberOfColumns, oHighlightHeaderStyle);
                oSLBOMDocument.SetColumnWidth(iNumberOfColumns++, 90);


                foreach (MfgPart oMfgPart in m_oMfgPartsWithoutOperations)
                {
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 1, oMfgPart.PartNum);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 2, oMfgPart.RevNum);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 3, StringExt.Truncate(oMfgPart.Description, 50));

                    iNumberOfRows++;
                    bDataInReport = true;
                }
            }

            // mfg part has no revision
            iNumberOfRows = 1;
            iNumberOfColumns = 1;
            if (m_oMfgPartsWithBadRevision.Count != 0)
            {
                if (bFirstWorksheet == true)
                {
                    oSLBOMDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "Mfg Part Bad Rev");
                    bFirstWorksheet = false;
                }
                else
                {
                    oSLBOMDocument.AddWorksheet("Mfg Part Bad Rev");
                }
                //set column header
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Part");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 30);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Description");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 60);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "The Mfg Part Has A Bad Revision Number");
                oSLBOMDocument.SetCellStyle(iNumberOfRows++, iNumberOfColumns, oHighlightHeaderStyle);
                oSLBOMDocument.SetColumnWidth(iNumberOfColumns++, 90);

                foreach (MfgPart oMfgPart in m_oMfgPartsWithBadRevision)
                {
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 1, oMfgPart.PartNum);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 2, StringExt.Truncate(oMfgPart.Description, 50));

                    iNumberOfRows++;
                    bDataInReport = true;
                }
            }

            // no approved revisions
            iNumberOfRows = 1;
            iNumberOfColumns = 1;
            if (m_oMfgPartsWithNoApprovedRevision.Count != 0)
            {
                if (bFirstWorksheet == true)
                {
                    oSLBOMDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "Mfg Part No Approved Rev");
                    bFirstWorksheet = false;
                }
                else
                {
                    oSLBOMDocument.AddWorksheet("Mfg Part No Approved Rev");
                }
                //set column header
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Part");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 30);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Description");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 60);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "The Mfg Part Has A Bad Or Missing Revision Number");
                oSLBOMDocument.SetCellStyle(iNumberOfRows++, iNumberOfColumns, oHighlightHeaderStyle);
                oSLBOMDocument.SetColumnWidth(iNumberOfColumns++, 90);

                foreach (MfgPart oMfgPart in m_oMfgPartsWithNoApprovedRevision)
                {
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 1, oMfgPart.PartNum);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 2, StringExt.Truncate(oMfgPart.Description, 50));

                    iNumberOfRows++;
                    bDataInReport = true;
                }
            }

            // multiple approved revisions
            iNumberOfRows = 1;
            iNumberOfColumns = 1;
            if (m_oMfgPartsWithMultipleApprovedRevisions.Count != 0)
            {
                if (bFirstWorksheet == true)
                {
                    oSLBOMDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "Multiple Approved Revs");
                    bFirstWorksheet = false;
                }
                else
                {
                    oSLBOMDocument.AddWorksheet("Multiple Approved Revs");
                }
                //set column header
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Part");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 30);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Revision");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 30);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Description");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 60);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "The Mfg Part Should Only Have One Approved Revision Number");
                oSLBOMDocument.SetCellStyle(iNumberOfRows++, iNumberOfColumns, oHighlightHeaderStyle);
                oSLBOMDocument.SetColumnWidth(iNumberOfColumns++, 90);

                foreach (MfgPart oMfgPart in m_oMfgPartsWithMultipleApprovedRevisions)
                {
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 1, oMfgPart.PartNum);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 2, oMfgPart.RevNum);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 3, StringExt.Truncate(oMfgPart.Description, 50));

                    iNumberOfRows++;
                    bDataInReport = true;
                }
            }

            // not using part rev
            iNumberOfRows = 1;
            iNumberOfColumns = 1;
            if (m_oMfgPartsNotUsingPartRevs.Count != 0)
            {
                if (bFirstWorksheet == true)
                {
                    oSLBOMDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "Use Part Rev Not Set");
                    bFirstWorksheet = false;
                }
                else
                {
                    oSLBOMDocument.AddWorksheet("Use Part Rev Not Set");
                }
                //set column header
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Part");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 30);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Description");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 60);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "The Mfg Part Should Have Use Part Rev Turned On");
                oSLBOMDocument.SetCellStyle(iNumberOfRows++, iNumberOfColumns, oHighlightHeaderStyle);
                oSLBOMDocument.SetColumnWidth(iNumberOfColumns++, 90);

                foreach (MfgPart oMfgPart in m_oMfgPartsNotUsingPartRevs)
                {
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 1, oMfgPart.PartNum);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 2, StringExt.Truncate(oMfgPart.Description, 50));

                    iNumberOfRows++;
                    bDataInReport = true;
                }
            }

            // not qty bearing
            iNumberOfRows = 1;
            iNumberOfColumns = 1;
            if (m_oMfgPartsNotQuantityBearing.Count != 0)
            {
                if (bFirstWorksheet == true)
                {
                    oSLBOMDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "Qty Bearing Not Set");
                    bFirstWorksheet = false;
                }
                else
                {
                    oSLBOMDocument.AddWorksheet("Qty Bearing Not Set");
                }
                //set column header
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Part");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 30);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Description");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 60);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "The Mfg Part Should Be Quantity Bearing");
                oSLBOMDocument.SetCellStyle(iNumberOfRows++, iNumberOfColumns, oHighlightHeaderStyle);
                oSLBOMDocument.SetColumnWidth(iNumberOfColumns++, 90);

                foreach (MfgPart oMfgPart in m_oMfgPartsNotQuantityBearing)
                {
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 1, oMfgPart.PartNum);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 2, StringExt.Truncate(oMfgPart.Description, 50));

                    iNumberOfRows++;
                    bDataInReport = true;
                }
            }

            // inconsistent qty bearing
            iNumberOfRows = 1;
            iNumberOfColumns = 1;
            if (m_oMfgPartsWithIncosistentQuantityBearing.Count != 0)
            {
                if (bFirstWorksheet == true)
                {
                    oSLBOMDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "Inconsistent Qty Bearing");
                    bFirstWorksheet = false;
                }
                else
                {
                    oSLBOMDocument.AddWorksheet("Inconsistent Qty Bearing");
                }
                //set column header
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Part");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 30);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Description");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 60);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "The Mfg Part Needs To Have Qty Bearing Checked For The Part Master And Plant");
                oSLBOMDocument.SetCellStyle(iNumberOfRows++, iNumberOfColumns, oHighlightHeaderStyle);
                oSLBOMDocument.SetColumnWidth(iNumberOfColumns++, 90);

                foreach (MfgPart oMfgPart in m_oMfgPartsWithIncosistentQuantityBearing)
                {
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 1, oMfgPart.PartNum);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 2, StringExt.Truncate(oMfgPart.Description, 50));

                    iNumberOfRows++;
                    bDataInReport = true;
                }
            }

            // incorrect UOM
            iNumberOfRows = 1;
            iNumberOfColumns = 1;
            if (m_oMfgPartsWithIncorrectUOM.Count != 0)
            {
                if (bFirstWorksheet == true)
                {
                    oSLBOMDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "Incorrect UOM");
                    bFirstWorksheet = false;
                }
                else
                {
                    oSLBOMDocument.AddWorksheet("Incorrect UOM");
                }
                //set column header
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Part");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 30);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Description");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 60);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "UOM");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 30);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "The Mfg Part Should Have EACH for the Unit Of Measure");
                oSLBOMDocument.SetCellStyle(iNumberOfRows++, iNumberOfColumns, oHighlightHeaderStyle);
                oSLBOMDocument.SetColumnWidth(iNumberOfColumns++, 90);

                foreach (MfgPart oMfgPart in m_oMfgPartsWithIncorrectUOM)
                {
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 1, oMfgPart.PartNum);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 2, StringExt.Truncate(oMfgPart.Description, 50));
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 3, oMfgPart.UOM);

                    iNumberOfRows++;
                    bDataInReport = true;
                }
            }

            // inconsisten Phantom BOM
            iNumberOfRows = 1;
            iNumberOfColumns = 1;
            if (m_oMfgPartsWithInconsistentPhantomBOM.Count != 0)
            {
                if (bFirstWorksheet == true)
                {
                    oSLBOMDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "Inconsistent Phantom Setting");
                    bFirstWorksheet = false;
                }
                else
                {
                    oSLBOMDocument.AddWorksheet("Inconsistent Phantom Setting");
                }
                //set column header
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Part");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 30);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Description");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 60);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "The Mfg Part Has Inconsistent Phantom BOM Settings On Part Master And Plant");
                oSLBOMDocument.SetCellStyle(iNumberOfRows++, iNumberOfColumns, oHighlightHeaderStyle);
                oSLBOMDocument.SetColumnWidth(iNumberOfColumns++, 90);

                foreach (MfgPart oMfgPart in m_oMfgPartsWithInconsistentPhantomBOM)
                {
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 1, oMfgPart.PartNum);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 2, StringExt.Truncate(oMfgPart.Description, 50));

                    iNumberOfRows++;
                    bDataInReport = true;
                }
            }

            // inconsistent non-stock
            iNumberOfRows = 1;
            iNumberOfColumns = 1;
            if (m_oMfgPartsWithInconsistentNonStock.Count != 0)
            {
                if (bFirstWorksheet == true)
                {
                    oSLBOMDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "Inconsistent Non-Stock");
                    bFirstWorksheet = false;
                }
                else
                {
                    oSLBOMDocument.AddWorksheet("Inconsistent Non-Stock");
                }
                //set column header
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Part");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 30);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Description");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 60);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "The Mfg Part Has Inconsistent Non-Stock Setting On The Part Master And Plant");
                oSLBOMDocument.SetCellStyle(iNumberOfRows++, iNumberOfColumns, oHighlightHeaderStyle);
                oSLBOMDocument.SetColumnWidth(iNumberOfColumns++, 90);

                foreach (MfgPart oMfgPart in m_oMfgPartsWithInconsistentNonStock)
                {
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 1, oMfgPart.PartNum);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 2, StringExt.Truncate(oMfgPart.Description, 50));

                    iNumberOfRows++;
                    bDataInReport = true;
                }
            }

            // inconsistent type code
            iNumberOfRows = 1;
            iNumberOfColumns = 1;
            if (m_oMfgPartsWithInconsistentTypeCode.Count != 0)
            {
                if (bFirstWorksheet == true)
                {
                    oSLBOMDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "Inconsistent Type");
                    bFirstWorksheet = false;
                }
                else
                {
                    oSLBOMDocument.AddWorksheet("Inconsistent Type");
                }
                //set column header
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Part");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 30);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Description");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 60);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "The Mfg Part Has Inconsistent Type Code For The Part Master And Plant");
                oSLBOMDocument.SetCellStyle(iNumberOfRows++, iNumberOfColumns, oHighlightHeaderStyle);
                oSLBOMDocument.SetColumnWidth(iNumberOfColumns++, 90);

                foreach (MfgPart oMfgPart in m_oMfgPartsWithInconsistentTypeCode)
                {
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 1, oMfgPart.PartNum);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 2, StringExt.Truncate(oMfgPart.Description, 50));

                    iNumberOfRows++;
                    bDataInReport = true;
                }
            }

            // mrp turned off
            iNumberOfRows = 1;
            iNumberOfColumns = 1;
            if (m_oMfgPartsWithMRPTurnedOff.Count != 0)
            {
                if (bFirstWorksheet == true)
                {
                    oSLBOMDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "MRP Off");
                    bFirstWorksheet = false;
                }
                else
                {
                    oSLBOMDocument.AddWorksheet("MRP Off");
                }
                //set column header
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Part");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 30);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Description");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 60);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "The Mfg Part Has MRP Turned Off");
                oSLBOMDocument.SetCellStyle(iNumberOfRows++, iNumberOfColumns, oHighlightHeaderStyle);
                oSLBOMDocument.SetColumnWidth(iNumberOfColumns++, 90);

                foreach (MfgPart oMfgPart in m_oMfgPartsWithMRPTurnedOff)
                {
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 1, oMfgPart.PartNum);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 2, StringExt.Truncate(oMfgPart.Description, 50));

                    iNumberOfRows++;
                    bDataInReport = true;
                }
            }

            // buy to order turned on
            iNumberOfRows = 1;
            iNumberOfColumns = 1;
            if (m_oMfgPartsWithBuyToOrderSet.Count != 0)
            {
                if (bFirstWorksheet == true)
                {
                    oSLBOMDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "Buy To Order Set");
                    bFirstWorksheet = false;
                }
                else
                {
                    oSLBOMDocument.AddWorksheet("Buy To Order Set");
                }
                //set column header
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Part");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 30);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Description");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 60);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "The Mfg Part Should Not Have Buy To Order Turned On");
                oSLBOMDocument.SetCellStyle(iNumberOfRows++, iNumberOfColumns, oHighlightHeaderStyle);
                oSLBOMDocument.SetColumnWidth(iNumberOfColumns++, 90);

                foreach (MfgPart oMfgPart in m_oMfgPartsWithBuyToOrderSet)
                {
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 1, oMfgPart.PartNum);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 2, StringExt.Truncate(oMfgPart.Description, 50));

                    iNumberOfRows++;
                    bDataInReport = true;
                }
            }

            // inconsistent buy to order
            iNumberOfRows = 1;
            iNumberOfColumns = 1;
            if (m_oMfgPartsWithInconsistentBuyToOrder.Count != 0)
            {
                if (bFirstWorksheet == true)
                {
                    oSLBOMDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "Inconsistent Buy To Order");
                    bFirstWorksheet = false;
                }
                else
                {
                    oSLBOMDocument.AddWorksheet("Inconsistent Buy To Order");
                }
                //set column header
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Part");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 30);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Description");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 60);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "The Mfg Part Has An Inconsistent Buy To Order Setting On The Part Master And Plant");
                oSLBOMDocument.SetCellStyle(iNumberOfRows++, iNumberOfColumns, oHighlightHeaderStyle);
                oSLBOMDocument.SetColumnWidth(iNumberOfColumns++, 90);

                foreach (MfgPart oMfgPart in m_oMfgPartsWithInconsistentBuyToOrder)
                {
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 1, oMfgPart.PartNum);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 2, StringExt.Truncate(oMfgPart.Description, 50));

                    iNumberOfRows++;
                    bDataInReport = true;
                }
            }

            // drop ship set
            iNumberOfRows = 1;
            iNumberOfColumns = 1;
            if (m_oMfgPartsWithDropShipSet.Count != 0)
            {
                if (bFirstWorksheet == true)
                {
                    oSLBOMDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "Drop Ship Set");
                    bFirstWorksheet = false;
                }
                else
                {
                    oSLBOMDocument.AddWorksheet("Drop Ship Set");
                }
                //set column header
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Part");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 30);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Description");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 60);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "The Mfg Part Should not Have Drop Set Turned on");
                oSLBOMDocument.SetCellStyle(iNumberOfRows++, iNumberOfColumns, oHighlightHeaderStyle);
                oSLBOMDocument.SetColumnWidth(iNumberOfColumns++, 90);

                foreach (MfgPart oMfgPart in m_oMfgPartsWithDropShipSet)
                {
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 1, oMfgPart.PartNum);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 2, StringExt.Truncate(oMfgPart.Description, 50));

                    iNumberOfRows++;
                    bDataInReport = true;
                }
            }

            // inconsistent drop ship
            iNumberOfRows = 1;
            iNumberOfColumns = 1;
            if (m_oMfgPartsWithInconsistentDropShip.Count != 0)
            {
                if (bFirstWorksheet == true)
                {
                    oSLBOMDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "Inconsistent Drop Ship");
                    bFirstWorksheet = false;
                }
                else
                {
                    oSLBOMDocument.AddWorksheet("Inconsistent Drop Ship");
                }
                //set column header
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Part");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 30);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Description");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 60);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "The Mfg Part Has An Inconsistent Drop Ship Setting On The Part Master And Plant");
                oSLBOMDocument.SetCellStyle(iNumberOfRows++, iNumberOfColumns, oHighlightHeaderStyle);
                oSLBOMDocument.SetColumnWidth(iNumberOfColumns++, 90);

                foreach (MfgPart oMfgPart in m_oMfgPartsWithInconsistentDropShip)
                {
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 1, oMfgPart.PartNum);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 2, StringExt.Truncate(oMfgPart.Description, 50));

                    iNumberOfRows++;
                    bDataInReport = true;
                }
            }

            // no primary warehouse
            iNumberOfRows = 1;
            iNumberOfColumns = 1;
            if (m_oMfgPartsWithoutPrimaryWarehouse.Count != 0)
            {
                if (bFirstWorksheet == true)
                {
                    oSLBOMDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "No Primary Warehouse");
                    bFirstWorksheet = false;
                }
                else
                {
                    oSLBOMDocument.AddWorksheet("No Primary Warehouse");
                }
                //set column header
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Part");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 30);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Description");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 60);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "The Mfg Part Doe Not Have A Primary Warehouse Set");
                oSLBOMDocument.SetCellStyle(iNumberOfRows++, iNumberOfColumns, oHighlightHeaderStyle);
                oSLBOMDocument.SetColumnWidth(iNumberOfColumns++, 90);

                foreach (MfgPart oMfgPart in m_oMfgPartsWithoutPrimaryWarehouse)
                {
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 1, oMfgPart.PartNum);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 2, StringExt.Truncate(oMfgPart.Description, 50));

                    iNumberOfRows++;
                    bDataInReport = true;
                }
            }

            // no group code
            iNumberOfRows = 1;
            iNumberOfColumns = 1;
            if (m_oMfgPartsWithoutGroupCode.Count != 0)
            {
                if (bFirstWorksheet == true)
                {
                    oSLBOMDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "No Product Group");
                    bFirstWorksheet = false;
                }
                else
                {
                    oSLBOMDocument.AddWorksheet("No Product Group");
                }
                //set column header
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Part");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 30);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Description");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 30);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "The Mfg Part Doe Not Have A Product Group Set");
                oSLBOMDocument.SetCellStyle(iNumberOfRows++, iNumberOfColumns, oHighlightHeaderStyle);
                oSLBOMDocument.SetColumnWidth(iNumberOfColumns++, 90);

                foreach (MfgPart oMfgPart in m_oMfgPartsWithoutGroupCode)
                {
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 1, oMfgPart.PartNum);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 2, oMfgPart.Description);

                    iNumberOfRows++;
                    bDataInReport = true;
                }
            }

            // no class code
            iNumberOfRows = 1;
            iNumberOfColumns = 1;
            if (m_oMfgPartsWithoutClassCode.Count != 0)
            {
                if (bFirstWorksheet == true)
                {
                    oSLBOMDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "No Class Code");
                    bFirstWorksheet = false;
                }
                else
                {
                    oSLBOMDocument.AddWorksheet("No Class Code");
                }
                //set column header
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Part");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 30);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Description");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 30);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "The Mfg Part Does Not have A Class Code Set");
                oSLBOMDocument.SetCellStyle(iNumberOfRows++, iNumberOfColumns, oHighlightHeaderStyle);
                oSLBOMDocument.SetColumnWidth(iNumberOfColumns++, 90);

                foreach (MfgPart oMfgPart in m_oMfgPartsWithoutClassCode)
                {
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 1, oMfgPart.PartNum);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 2, oMfgPart.Description);

                    iNumberOfRows++;
                    bDataInReport = true;
                }
            }

            // incorrect cost method
            iNumberOfRows = 1;
            iNumberOfColumns = 1;
            if (m_oMfgPartsWithIncorrectCostMethod.Count != 0)
            {
                if (bFirstWorksheet == true)
                {
                    oSLBOMDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "Bad Cost Method");
                    bFirstWorksheet = false;
                }
                else
                {
                    oSLBOMDocument.AddWorksheet("Bad Cost Method");
                }
                //set column header
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Part");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 30);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Description");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 60);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Cost Method");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 30);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "The Mfg Part Has A Bad Cost Method");
                oSLBOMDocument.SetCellStyle(iNumberOfRows++, iNumberOfColumns, oHighlightHeaderStyle);
                oSLBOMDocument.SetColumnWidth(iNumberOfColumns++, 90);

                foreach (MfgPart oMfgPart in m_oMfgPartsWithIncorrectCostMethod)
                {
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 1, oMfgPart.PartNum);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 2, StringExt.Truncate(oMfgPart.Description, 50));
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 3, oMfgPart.PartMaster.PartCostingMethod);

                    iNumberOfRows++;
                    bDataInReport = true;
                }
            }

            // mfg parts with std cost needing cost roll
            if (m_oMfgPartsNeedingCostRoll.Count != 0)
            {
                if (bFirstWorksheet == true)
                {
                    oSLBOMDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "Need Cost Roll");
                    bFirstWorksheet = false;
                }
                else
                {
                    oSLBOMDocument.AddWorksheet("Need Cost Roll");
                }
                //set column header
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Part");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 30);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Description");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 60);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Cost Method");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 30);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Cost");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 30);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "The Mfg Part Is Set To Standard And Requires A Cost Roll");
                oSLBOMDocument.SetCellStyle(iNumberOfRows++, iNumberOfColumns, oHighlightHeaderStyle);
                oSLBOMDocument.SetColumnWidth(iNumberOfColumns++, 90);

                foreach (MfgPart oMfgPart in m_oMfgPartsNeedingCostRoll)
                {
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 1, oMfgPart.PartNum);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 2, StringExt.Truncate(oMfgPart.Description, 50));
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 3, oMfgPart.PartMaster.PartCostingMethod);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 4, oMfgPart.PartMaster.Cost);

                    iNumberOfRows++;
                    bDataInReport = true;
                }
            }

            // inconsistent cost method
            iNumberOfRows = 1;
            iNumberOfColumns = 1;
            if (m_oMfgPartsWithInconsistentCostMethod.Count != 0)
            {
                if (bFirstWorksheet == true)
                {
                    oSLBOMDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "Inconsistent Cost Method");
                    bFirstWorksheet = false;
                }
                else
                {
                    oSLBOMDocument.AddWorksheet("Inconsistent Cost Method");
                }
                //set column header
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Part");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 30);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Description");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 60);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "The Mfg Part Cost Method Is Inconsistent On The Part Master And Plant");
                oSLBOMDocument.SetCellStyle(iNumberOfRows++, iNumberOfColumns, oHighlightHeaderStyle);
                oSLBOMDocument.SetColumnWidth(iNumberOfColumns++, 90);

                foreach (MfgPart oMfgPart in m_oMfgPartsWithInconsistentCostMethod)
                {
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 1, oMfgPart.PartNum);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 2, StringExt.Truncate(oMfgPart.Description, 50));

                    iNumberOfRows++;
                    bDataInReport = true;
                }
            }

            // mfg parts that are stocked
            if (m_oMfgPartsSetToStocking.Count != 0)
            {
                if (bFirstWorksheet == true)
                {
                    oSLBOMDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "Stocked Mfg Parts");
                    bFirstWorksheet = false;
                }
                else
                {
                    oSLBOMDocument.AddWorksheet("Stocked Mfg Parts");
                }
                //set column header
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Part");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 30);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Description");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 60);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "The Mfg Part Should Be Set To Non-Stock (Make Direct)");
                oSLBOMDocument.SetCellStyle(iNumberOfRows++, iNumberOfColumns, oHighlightHeaderStyle);
                oSLBOMDocument.SetColumnWidth(iNumberOfColumns++, 90);

                foreach (MfgPart oMfgPart in m_oMfgPartsSetToStocking)
                {
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 1, oMfgPart.PartNum);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 2, StringExt.Truncate(oMfgPart.Description, 50));

                    iNumberOfRows++;
                    bDataInReport = true;
                }
            }

            // no unit price set
            iNumberOfRows = 1;
            iNumberOfColumns = 1;
            if (m_oMfgPartsStockedWithoutUnitPrice.Count != 0)
            {
                if (bFirstWorksheet == true)
                {
                    oSLBOMDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "Stock Mfg No Price");
                    bFirstWorksheet = false;
                }
                else
                {
                    oSLBOMDocument.AddWorksheet("Stock Mfg No Price");
                }
                //set column header
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Part");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 30);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Description");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 60);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "The Stocked Mfg Part Does Not Have The Price Set");
                oSLBOMDocument.SetCellStyle(iNumberOfRows++, iNumberOfColumns, oHighlightHeaderStyle);
                oSLBOMDocument.SetColumnWidth(iNumberOfColumns++, 90);

                foreach (MfgPart oMfgPart in m_oMfgPartsStockedWithoutUnitPrice)
                {
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 1, oMfgPart.PartNum);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 2, StringExt.Truncate(oMfgPart.Description, 50));

                    iNumberOfRows++;
                    bDataInReport = true;
                }
            }

            // stocking but no min or safety
            iNumberOfRows = 1;
            iNumberOfColumns = 1;
            if (m_oMfgPartsStockedWithoutMinOrSafetyLimits.Count != 0)
            {
                if (bFirstWorksheet == true)
                {
                    oSLBOMDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "Missing Limits");
                    bFirstWorksheet = false;
                }
                else
                {
                    oSLBOMDocument.AddWorksheet("Missing Limits");
                }
                //set column header
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Part");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 30);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Description");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 60);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "The Mfg Part Is Set To Stocking But Has No Min Or Safety Limit Set");
                oSLBOMDocument.SetCellStyle(iNumberOfRows++, iNumberOfColumns, oHighlightHeaderStyle);
                oSLBOMDocument.SetColumnWidth(iNumberOfColumns++, 90);

                foreach (MfgPart oMfgPart in m_oMfgPartsStockedWithoutMinOrSafetyLimits)
                {
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 1, oMfgPart.PartNum);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 2, StringExt.Truncate(oMfgPart.Description, 50));

                    iNumberOfRows++;
                    bDataInReport = true;
                }
            }

            // no portfolio code set
            iNumberOfRows = 1;
            iNumberOfColumns = 1;
            if (m_oMfgPartsWithoutPortfolioCode.Count != 0)
            {
                if (bFirstWorksheet == true)
                {
                    oSLBOMDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "No Portfolio Code");
                    bFirstWorksheet = false;
                }
                else
                {
                    oSLBOMDocument.AddWorksheet("No Portfolio Code");
                }
                //set column header
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Part");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 30);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Description");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 60);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "The Mfg Part Is Missing The Product Portfolio Code");
                oSLBOMDocument.SetCellStyle(iNumberOfRows++, iNumberOfColumns, oHighlightHeaderStyle);
                oSLBOMDocument.SetColumnWidth(iNumberOfColumns++, 90);

                foreach (MfgPart oMfgPart in m_oMfgPartsWithoutPortfolioCode)
                {
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 1, oMfgPart.PartNum);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 2, StringExt.Truncate(oMfgPart.Description, 50));

                    iNumberOfRows++;
                    bDataInReport = true;
                }
            }

            // operations not set to each
            iNumberOfRows = 1;
            iNumberOfColumns = 1;
            if (m_oMfgPartWithOperationsNotSetToEach.Count != 0)
            {
                if (bFirstWorksheet == true)
                {
                    oSLBOMDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "Op Not Set To Each");
                    bFirstWorksheet = false;
                }
                else
                {
                    oSLBOMDocument.AddWorksheet("Op Not Set To Each");
                }
                //set column header
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Part");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 30);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Revision");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 30);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Op Code");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 30);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Op Seq");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 30);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Std Basis");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 30);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "The Operation On The Mfg Part Is Not Set To Each");
                oSLBOMDocument.SetCellStyle(iNumberOfRows++, iNumberOfColumns, oHighlightHeaderStyle);
                oSLBOMDocument.SetColumnWidth(iNumberOfColumns++, 90);

                foreach (MfgPart oMfgPart in m_oMfgPartWithOperationsNotSetToEach)
                {
                    // find the list of operations not set to each for this part
                    List<PartOperation> oOperationsNotSetToEach = oMfgPart.MyPartOperations.Where(oItem => string.Compare(oItem.StandardBasis, "E", true) != 0).ToList();
                    foreach (PartOperation oPartOperation in oOperationsNotSetToEach)
                    {
                        oSLBOMDocument.SetCellValue(iNumberOfRows, 1, oMfgPart.PartNum);
                        oSLBOMDocument.SetCellValue(iNumberOfRows, 2, oMfgPart.RevNum);
                        oSLBOMDocument.SetCellValue(iNumberOfRows, 3, oPartOperation.OpCode);
                        oSLBOMDocument.SetCellValue(iNumberOfRows, 4, oPartOperation.OprSeq);
                        oSLBOMDocument.SetCellValue(iNumberOfRows, 5, oPartOperation.StandardBasis);

                        iNumberOfRows++;
                        bDataInReport = true;
                    }

                    bDataInReport = true;
                }
            }

            // operations not using time and quantity for labor entry
            iNumberOfRows = 1;
            iNumberOfColumns = 1;
            if (m_oMfgPartWithOperationsBadLaborEntry.Count != 0)
            {
                if (bFirstWorksheet == true)
                {
                    oSLBOMDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "Op Has Bad Labor Entry");
                    bFirstWorksheet = false;
                }
                else
                {
                    oSLBOMDocument.AddWorksheet("Op Has Bad Labor Entry");
                }
                //set column header
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Part");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 30);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Revision");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 30);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Op Code");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 30);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Op Seq");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 30);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Labor Entry");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 30);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "The Operation Labor Entry Should Be Set To Time And Qty");
                oSLBOMDocument.SetCellStyle(iNumberOfRows++, iNumberOfColumns, oHighlightHeaderStyle);
                oSLBOMDocument.SetColumnWidth(iNumberOfColumns++, 90);

                foreach (MfgPart oMfgPart in m_oMfgPartWithOperationsBadLaborEntry)
                {
                    // find the list of operations not set to time and quantity
                    List<PartOperation> oOperationsWithBadLaborEntry = oMfgPart.MyPartOperations.Where(oItem => (string.Compare(oItem.LaborEntryMethod, "T", true) != 0)).ToList();
                    foreach (PartOperation oPartOperation in oOperationsWithBadLaborEntry)
                    {
                        oSLBOMDocument.SetCellValue(iNumberOfRows, 1, oMfgPart.PartNum);
                        oSLBOMDocument.SetCellValue(iNumberOfRows, 2, oMfgPart.RevNum);
                        oSLBOMDocument.SetCellValue(iNumberOfRows, 3, oPartOperation.OpCode);
                        oSLBOMDocument.SetCellValue(iNumberOfRows, 4, oPartOperation.OprSeq);
                        oSLBOMDocument.SetCellValue(iNumberOfRows, 5, oPartOperation.LaborEntryMethod);

                        iNumberOfRows++;
                        bDataInReport = true;
                    }

                    bDataInReport = true;
                }
            }

            // operations not using "HP" for standard
            iNumberOfRows = 1;
            iNumberOfColumns = 1;
            if (m_oMfgPartWithOperationsBadStandardFormat.Count != 0)
            {
                if (bFirstWorksheet == true)
                {
                    oSLBOMDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "Op With Bad Std Format");
                    bFirstWorksheet = false;
                }
                else
                {
                    oSLBOMDocument.AddWorksheet("Op With Bad Std Format");
                }
                //set column header
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Part");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 30);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Revision");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 30);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Op Code");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 30);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Op Seq");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 30);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Std Format");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 30);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "The Operation Time On The Mfg Part Should Be Set To Hours Per Piece");
                oSLBOMDocument.SetCellStyle(iNumberOfRows++, iNumberOfColumns, oHighlightHeaderStyle);
                oSLBOMDocument.SetColumnWidth(iNumberOfColumns++, 90);

                foreach (MfgPart oMfgPart in m_oMfgPartWithOperationsBadStandardFormat)
                {
                    // find the list of operations not set to "HP" for this part
                    List<PartOperation> oOperationsWithBadStandardFormat = oMfgPart.MyPartOperations.Where(oItem => (string.Compare(oItem.StandardFormat, "HP", true) != 0)).ToList();
                    foreach (PartOperation oPartOperation in oOperationsWithBadStandardFormat)
                    {
                        oSLBOMDocument.SetCellValue(iNumberOfRows, 1, oMfgPart.PartNum);
                        oSLBOMDocument.SetCellValue(iNumberOfRows, 2, oMfgPart.RevNum);
                        oSLBOMDocument.SetCellValue(iNumberOfRows, 3, oPartOperation.OpCode);
                        oSLBOMDocument.SetCellValue(iNumberOfRows, 4, oPartOperation.OprSeq);
                        oSLBOMDocument.SetCellValue(iNumberOfRows, 5, oPartOperation.StandardFormat);

                        iNumberOfRows++;
                        bDataInReport = true;
                    }

                    bDataInReport = true;
                }
            }

            // all operations where the production standard is zero
            iNumberOfRows = 1;
            iNumberOfColumns = 1;
            if (m_oMfgPartWithOperationsBadProductionStandard.Count != 0)
            {
                if (bFirstWorksheet == true)
                {
                    oSLBOMDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "Op With Prod Std Set To Zero");
                    bFirstWorksheet = false;
                }
                else
                {
                    oSLBOMDocument.AddWorksheet("Op With Prod Std Set To Zero");
                }
                //set column header
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Part");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 30);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Revision");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 30);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Op Code");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 30);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Op Seq");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 30);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Production Std");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 30);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "The Operation Time On The Mfg Part Is Set To Zero");
                oSLBOMDocument.SetCellStyle(iNumberOfRows++, iNumberOfColumns, oHighlightHeaderStyle);
                oSLBOMDocument.SetColumnWidth(iNumberOfColumns++, 90);

                foreach (MfgPart oMfgPart in m_oMfgPartWithOperationsBadProductionStandard)
                {
                    // find the list of operations where the production standard is zero
                    List<PartOperation> oOperationsWithBadProductionStandard = oMfgPart.MyPartOperations.Where(oItem => (oItem.ProductionStandard == 0) && (oItem.Subcontract == false)).ToList();
                    foreach (PartOperation oPartOperation in oOperationsWithBadProductionStandard)
                    {
                        oSLBOMDocument.SetCellValue(iNumberOfRows, 1, oMfgPart.PartNum);
                        oSLBOMDocument.SetCellValue(iNumberOfRows, 2, oMfgPart.RevNum);
                        oSLBOMDocument.SetCellValue(iNumberOfRows, 3, oPartOperation.OpCode);
                        oSLBOMDocument.SetCellValue(iNumberOfRows, 4, oPartOperation.OprSeq);
                        oSLBOMDocument.SetCellValue(iNumberOfRows, 5, oPartOperation.ProductionStandard);

                        iNumberOfRows++;
                        bDataInReport = true;
                    }

                    bDataInReport = true;
                }
            }

            // all operations where the std format is OM or OH and the ops per part is zero
            iNumberOfRows = 1;
            iNumberOfColumns = 1;
            if (m_oMfgPartWithOperationsBadOperationPerPartValue.Count != 0)
            {
                if (bFirstWorksheet == true)
                {
                    oSLBOMDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "Op With Bad Ops Per Part Value");
                    bFirstWorksheet = false;
                }
                else
                {
                    oSLBOMDocument.AddWorksheet("Op With Bad Ops Per Part Value");
                }
                //set column header
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Part");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 30);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Revision");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 30);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Op Code");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 30);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Op Seq");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 30);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Std Format");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 30);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Ops Per Part");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 30);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "The Operations Per Part Should Be Set To Zero");
                oSLBOMDocument.SetCellStyle(iNumberOfRows++, iNumberOfColumns, oHighlightHeaderStyle);
                oSLBOMDocument.SetColumnWidth(iNumberOfColumns++, 90);

                foreach (MfgPart oMfgPart in m_oMfgPartWithOperationsBadOperationPerPartValue)
                {
                    // if std format is "OM" or "OH" then the OpsPerPart field must be > 0
                    List<PartOperation> oOperationsMissingOperationsPerPartValue = oMfgPart.MyPartOperations.Where(oItem => ((string.Compare(oItem.StandardFormat, "OM", true) == 0) || (string.Compare(oItem.StandardFormat, "OH", true) == 0)) && (oItem.OperationsPerPart == 0)).ToList();

                    foreach (PartOperation oPartOperation in oOperationsMissingOperationsPerPartValue)
                    {
                        oSLBOMDocument.SetCellValue(iNumberOfRows, 1, oMfgPart.PartNum);
                        oSLBOMDocument.SetCellValue(iNumberOfRows, 2, oMfgPart.RevNum);
                        oSLBOMDocument.SetCellValue(iNumberOfRows, 3, oPartOperation.OpCode);
                        oSLBOMDocument.SetCellValue(iNumberOfRows, 4, oPartOperation.OprSeq);
                        oSLBOMDocument.SetCellValue(iNumberOfRows, 5, oPartOperation.StandardFormat);
                        oSLBOMDocument.SetCellValue(iNumberOfRows, 6, oPartOperation.OperationsPerPart);

                        iNumberOfRows++;
                        bDataInReport = true;
                    }

                    bDataInReport = true;
                }
            }

            // all operations where the std format is NOT OM or OH and the ops per part is NOT zero
            iNumberOfRows = 1;
            iNumberOfColumns = 1;
            if (m_oMfgPartWithNonZeroOperationsPerPartValue.Count != 0)
            {
                if (bFirstWorksheet == true)
                {
                    oSLBOMDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "Op Std With Bad Value");
                    bFirstWorksheet = false;
                }
                else
                {
                    oSLBOMDocument.AddWorksheet("Op Std With Bad Value");
                }
                //set column header
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Part");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 30);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Revision");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 30);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Op Code");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 30);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Op Seq");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 30);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Std Format");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 30);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Ops Per Part");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 30);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "The Operation Per Part Should Be Set When Using Operations Per Hour Or Minute");
                oSLBOMDocument.SetCellStyle(iNumberOfRows++, iNumberOfColumns, oHighlightHeaderStyle);
                oSLBOMDocument.SetColumnWidth(iNumberOfColumns++, 90);

                foreach (MfgPart oMfgPart in m_oMfgPartWithNonZeroOperationsPerPartValue)
                {
                    // std format is NOT "OM" or "OH" but the OpsPerPart field is set to non-zero value
                    List<PartOperation> oOperationsNonZeroOperationsPerPartValue = oMfgPart.MyPartOperations.Where(oItem => (string.Compare(oItem.StandardFormat, "OM", true) != 0) && (string.Compare(oItem.StandardFormat, "OH", true) != 0) && (oItem.OperationsPerPart != 0)).ToList();

                    foreach (PartOperation oPartOperation in oOperationsNonZeroOperationsPerPartValue)
                    {
                        oSLBOMDocument.SetCellValue(iNumberOfRows, 1, oMfgPart.PartNum);
                        oSLBOMDocument.SetCellValue(iNumberOfRows, 2, oMfgPart.RevNum);
                        oSLBOMDocument.SetCellValue(iNumberOfRows, 3, oPartOperation.OpCode);
                        oSLBOMDocument.SetCellValue(iNumberOfRows, 4, oPartOperation.OprSeq);
                        oSLBOMDocument.SetCellValue(iNumberOfRows, 5, oPartOperation.StandardFormat);
                        oSLBOMDocument.SetCellValue(iNumberOfRows, 6, oPartOperation.OperationsPerPart);

                        iNumberOfRows++;
                        bDataInReport = true;
                    }

                    bDataInReport = true;
                }
            }

            // all operations where there is an additional setup qty
            iNumberOfRows = 1;
            iNumberOfColumns = 1;
            if (m_oMfgPartWithOperationsWithAdditionalSetupQty.Count != 0)
            {
                if (bFirstWorksheet == true)
                {
                    oSLBOMDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "Op With Additional Setup Qty");
                    bFirstWorksheet = false;
                }
                else
                {
                    oSLBOMDocument.AddWorksheet("Op With Additional Setup Qty");
                }
                //set column header
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Part");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 30);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Revision");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 30);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Op Code");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 30);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Op Seq");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 30);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Additional Qty");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 30);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "The Operation Is Using The Additional Setup Quantity");
                oSLBOMDocument.SetCellStyle(iNumberOfRows++, iNumberOfColumns, oHighlightHeaderStyle);
                oSLBOMDocument.SetColumnWidth(iNumberOfColumns++, 90);

                foreach (MfgPart oMfgPart in m_oMfgPartWithOperationsWithAdditionalSetupQty)
                {
                    // operation where additional setup qty is specified
                    List<PartOperation> oOperationsWithAdditionalSetupQuantity = oMfgPart.MyPartOperations.Where(oItem => oItem.AdditionalSetQty != 0).ToList();
                    foreach (PartOperation oPartOperation in oOperationsWithAdditionalSetupQuantity)
                    {
                        oSLBOMDocument.SetCellValue(iNumberOfRows, 1, oMfgPart.PartNum);
                        oSLBOMDocument.SetCellValue(iNumberOfRows, 2, oMfgPart.RevNum);
                        oSLBOMDocument.SetCellValue(iNumberOfRows, 3, oPartOperation.OpCode);
                        oSLBOMDocument.SetCellValue(iNumberOfRows, 4, oPartOperation.OprSeq);
                        oSLBOMDocument.SetCellValue(iNumberOfRows, 5, oPartOperation.AdditionalSetQty);

                        iNumberOfRows++;
                        bDataInReport = true;
                    }

                    bDataInReport = true;
                }
            }

            // all operations where there is an additional setup hours
            iNumberOfRows = 1;
            iNumberOfColumns = 1;
            if (m_oMfgPartWithOperationsWithAdditionalSetupHours.Count != 0)
            {
                if (bFirstWorksheet == true)
                {
                    oSLBOMDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "Op With Additional Setup Hours");
                    bFirstWorksheet = false;
                }
                else
                {
                    oSLBOMDocument.AddWorksheet("Op With Additional Setup Hours");
                }
                //set column header
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Part");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 30);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Revision");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 30);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Op Code");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 30);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Op Seq");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 30);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Additional Setup Hours");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 30);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "The Operation Is Using Additional Setup Hours");
                oSLBOMDocument.SetCellStyle(iNumberOfRows++, iNumberOfColumns, oHighlightHeaderStyle);
                oSLBOMDocument.SetColumnWidth(iNumberOfColumns++, 90);

                foreach (MfgPart oMfgPart in m_oMfgPartWithOperationsWithAdditionalSetupHours)
                {
                    // op where additional setup hours is specified
                    List<PartOperation> oOperationsWithAdditionalSetupHours = oMfgPart.MyPartOperations.Where(oItem => oItem.AdditonalSetupHours != 0).ToList();
                    foreach (PartOperation oPartOperation in oOperationsWithAdditionalSetupHours)
                    {
                        oSLBOMDocument.SetCellValue(iNumberOfRows, 1, oMfgPart.PartNum);
                        oSLBOMDocument.SetCellValue(iNumberOfRows, 2, oMfgPart.RevNum);
                        oSLBOMDocument.SetCellValue(iNumberOfRows, 3, oPartOperation.OpCode);
                        oSLBOMDocument.SetCellValue(iNumberOfRows, 4, oPartOperation.OprSeq);
                        oSLBOMDocument.SetCellValue(iNumberOfRows, 5, oPartOperation.AdditonalSetupHours);

                        iNumberOfRows++;
                        bDataInReport = true;
                    }

                    bDataInReport = true;
                }
            }

            // all operations where there is a resource specified
            iNumberOfRows = 1;
            iNumberOfColumns = 1;
            if (m_oMfgPartWithOperationsWithResourceSpecified.Count != 0)
            {
                if (bFirstWorksheet == true)
                {
                    oSLBOMDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "Op With Resource Specified");
                    bFirstWorksheet = false;
                }
                else
                {
                    oSLBOMDocument.AddWorksheet("Op With Resource Specified");
                }
                //set column header
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Part");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 30);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Revision");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 30);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Op Code");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 30);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Op Seq");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 30);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Resource");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 30);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "The Operation Should Not Specify A Resource Just The Resource Group");
                oSLBOMDocument.SetCellStyle(iNumberOfRows++, iNumberOfColumns, oHighlightHeaderStyle);
                oSLBOMDocument.SetColumnWidth(iNumberOfColumns++, 90);

                foreach (MfgPart oMfgPart in m_oMfgPartWithOperationsWithResourceSpecified)
                {
                    // operation where resource is specified
                    List<PartOperation> oOperationsWithResourceSet = oMfgPart.MyPartOperations.Where(oItem => string.IsNullOrEmpty(oItem.ResourceId) == false).ToList();
                    foreach (PartOperation oPartOperation in oOperationsWithResourceSet)
                    {
                        oSLBOMDocument.SetCellValue(iNumberOfRows, 1, oMfgPart.PartNum);
                        oSLBOMDocument.SetCellValue(iNumberOfRows, 2, oMfgPart.RevNum);
                        oSLBOMDocument.SetCellValue(iNumberOfRows, 3, oPartOperation.OpCode);
                        oSLBOMDocument.SetCellValue(iNumberOfRows, 4, oPartOperation.OprSeq);
                        oSLBOMDocument.SetCellValue(iNumberOfRows, 5, oPartOperation.ResourceId);

                        iNumberOfRows++;
                        bDataInReport = true;
                    }

                    bDataInReport = true;
                }
            }

            // all subcontract operations where Qty Per is not set
            iNumberOfRows = 1;
            iNumberOfColumns = 1;
            if (m_oMfgPartSubcontractOperationMissingQtyPer.Count != 0)
            {
                if (bFirstWorksheet == true)
                {
                    oSLBOMDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "Subcontract No QtyPer");
                    bFirstWorksheet = false;
                }
                else
                {
                    oSLBOMDocument.AddWorksheet("Subcontract No QtyPer");
                }
                //set column header
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Part");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 30);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Revision");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 30);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Op Code");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 30);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Op Seq");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 30);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Qty Per");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 30);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "The Subcontract Operation Should Indicate The Quantity Per");
                oSLBOMDocument.SetCellStyle(iNumberOfRows++, iNumberOfColumns, oHighlightHeaderStyle);
                oSLBOMDocument.SetColumnWidth(iNumberOfColumns++, 90);

                foreach (MfgPart oMfgPart in m_oMfgPartSubcontractOperationMissingQtyPer)
                {
                    // subcontract where qty per is not set
                    List<PartOperation> oSubcontractOperationMissingQtyPer = oMfgPart.MyPartOperations.Where(oItem => (oItem.Subcontract == true) && (oItem.QtyPerParent == 0)).ToList();
                    foreach (PartOperation oPartOperation in oSubcontractOperationMissingQtyPer)
                    {
                        oSLBOMDocument.SetCellValue(iNumberOfRows, 1, oMfgPart.PartNum);
                        oSLBOMDocument.SetCellValue(iNumberOfRows, 2, oMfgPart.RevNum);
                        oSLBOMDocument.SetCellValue(iNumberOfRows, 3, oPartOperation.OpCode);
                        oSLBOMDocument.SetCellValue(iNumberOfRows, 4, oPartOperation.OprSeq);
                        oSLBOMDocument.SetCellValue(iNumberOfRows, 5, oPartOperation.QtyPerParent);

                        iNumberOfRows++;
                        bDataInReport = true;
                    }

                    bDataInReport = true;
                }
            }

            // all subcontract operations where Days Out is not set
            iNumberOfRows = 1;
            iNumberOfColumns = 1;
            if (m_oMfgPartSubcontractOperationMissingDaysOut.Count != 0)
            {
                if (bFirstWorksheet == true)
                {
                    oSLBOMDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "Subcontract No Days Out");
                    bFirstWorksheet = false;
                }
                else
                {
                    oSLBOMDocument.AddWorksheet("Subcontract No Days Out");
                }
                //set column header
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Part");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 30);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Revision");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 30);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Op Code");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 30);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Op Seq");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 30);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Days Out");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 30);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "The Subcontract Operation Should Specify The Time Required");
                oSLBOMDocument.SetCellStyle(iNumberOfRows++, iNumberOfColumns, oHighlightHeaderStyle);
                oSLBOMDocument.SetColumnWidth(iNumberOfColumns++, 90);

                foreach (MfgPart oMfgPart in m_oMfgPartSubcontractOperationMissingDaysOut)
                {
                    // subcontract where days out not set
                    List<PartOperation> oSubcontractOperationMissingDaysOut = oMfgPart.MyPartOperations.Where(oItem => (oItem.Subcontract == true) && (oItem.DaysOut == 0)).ToList();
                    foreach (PartOperation oPartOperation in oSubcontractOperationMissingDaysOut)
                    {
                        oSLBOMDocument.SetCellValue(iNumberOfRows, 1, oMfgPart.PartNum);
                        oSLBOMDocument.SetCellValue(iNumberOfRows, 2, oMfgPart.RevNum);
                        oSLBOMDocument.SetCellValue(iNumberOfRows, 3, oPartOperation.OpCode);
                        oSLBOMDocument.SetCellValue(iNumberOfRows, 4, oPartOperation.OprSeq);
                        oSLBOMDocument.SetCellValue(iNumberOfRows, 5, oPartOperation.DaysOut);

                        iNumberOfRows++;
                        bDataInReport = true;
                    }

                    bDataInReport = true;
                }
            }

            // all subcontract operations where unit cost is not set
            iNumberOfRows = 1;
            iNumberOfColumns = 1;
            if (m_oMfgPartSubcontractOperationMissingUnitCost.Count != 0)
            {
                if (bFirstWorksheet == true)
                {
                    oSLBOMDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "Subcontract No Unit Cost");
                    bFirstWorksheet = false;
                }
                else
                {
                    oSLBOMDocument.AddWorksheet("Subcontract No Unit Cost");
                }
                //set column header
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Part");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 30);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Revision");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 30);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Op Code");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 30);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Op Seq");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 30);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Unit Cost");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 30);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "The Subcontract Operation Should Indicate The Cost Per Unit");
                oSLBOMDocument.SetCellStyle(iNumberOfRows++, iNumberOfColumns, oHighlightHeaderStyle);
                oSLBOMDocument.SetColumnWidth(iNumberOfColumns++, 90);

                foreach (MfgPart oMfgPart in m_oMfgPartSubcontractOperationMissingUnitCost)
                {
                    // subcontract where days out not set
                    List<PartOperation> oSubcontractOperationMissingUnitCost = oMfgPart.MyPartOperations.Where(oItem => (oItem.Subcontract == true) && (oItem.EstUnitCost == 0)).ToList();
                    foreach (PartOperation oPartOperation in oSubcontractOperationMissingUnitCost)
                    {
                        oSLBOMDocument.SetCellValue(iNumberOfRows, 1, oMfgPart.PartNum);
                        oSLBOMDocument.SetCellValue(iNumberOfRows, 2, oMfgPart.RevNum);
                        oSLBOMDocument.SetCellValue(iNumberOfRows, 3, oPartOperation.OpCode);
                        oSLBOMDocument.SetCellValue(iNumberOfRows, 4, oPartOperation.OprSeq);
                        oSLBOMDocument.SetCellValue(iNumberOfRows, 5, oPartOperation.EstUnitCost);

                        iNumberOfRows++;
                        bDataInReport = true;
                    }

                    bDataInReport = true;
                }
            }

            // materials with fixed qty
            iNumberOfRows = 1;
            iNumberOfColumns = 1;
            if (m_oMfgPartMaterialsWithFixedQty.Count != 0)
            {
                if (bFirstWorksheet == true)
                {
                    oSLBOMDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "Mtl Fixed Qty");
                    bFirstWorksheet = false;
                }
                else
                {
                    oSLBOMDocument.AddWorksheet("Mtl Fixed Qty");
                }
                //set column header
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Parent Part");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 30);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Revision");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 30);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Mtl");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 30);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Mtl Seq");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 30);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "A Material  On The Mfg Part Is Set To Fixed Quantity");
                oSLBOMDocument.SetCellStyle(iNumberOfRows++, iNumberOfColumns, oHighlightHeaderStyle);
                oSLBOMDocument.SetColumnWidth(iNumberOfColumns++, 90);

                foreach (MfgPart oMfgPart in m_oMfgPartMaterialsWithFixedQty)
                {
                    // find all materials that have fixed qty
                    List<PartMaterial> oPartsWithFixedQty = oMfgPart.MyPartMaterials.Where(oItem => oItem.FixedQty == true).ToList();
                    foreach (PartMaterial oPartMaterial in oPartsWithFixedQty)
                    {
                        oSLBOMDocument.SetCellValue(iNumberOfRows, 1, oMfgPart.PartNum);
                        oSLBOMDocument.SetCellValue(iNumberOfRows, 2, oMfgPart.RevNum);
                        oSLBOMDocument.SetCellValue(iNumberOfRows, 3, oPartMaterial.PartMaster.PartNum);
                        oSLBOMDocument.SetCellValue(iNumberOfRows, 4, oPartMaterial.PartMaster.PartNum);

                        iNumberOfRows++;
                        bDataInReport = true;
                    }
                }
            }

            // materials with qty set to zero
            iNumberOfRows = 1;
            iNumberOfColumns = 1;
            if (m_oMfgPartsWithMaterialsWithZeroQty.Count != 0)
            {
                if (bFirstWorksheet == true)
                {
                    oSLBOMDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "Mtl Qty Zero");
                    bFirstWorksheet = false;
                }
                else
                {
                    oSLBOMDocument.AddWorksheet("Mtl Qty Zero");
                }
                //set column header
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Parent Part");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 30);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Revision");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 30);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Mtl");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 30);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Mtl Seq");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 30);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "The Material On The Mfg Part Has The Quantity Set To Zero");
                oSLBOMDocument.SetCellStyle(iNumberOfRows++, iNumberOfColumns, oHighlightHeaderStyle);
                oSLBOMDocument.SetColumnWidth(iNumberOfColumns++, 90);

                foreach (MfgPart oMfgPart in m_oMfgPartsWithMaterialsWithZeroQty)
                {
                    // find all materials that have qty set to zero
                    List<PartMaterial> oPartsWithZeroQty = oMfgPart.MyPartMaterials.Where(oItem => oItem.QuantityPer == 0).ToList();
                    foreach (PartMaterial oPartMaterial in oPartsWithZeroQty)
                    {
                        oSLBOMDocument.SetCellValue(iNumberOfRows, 1, oMfgPart.PartNum);
                        oSLBOMDocument.SetCellValue(iNumberOfRows, 2, oMfgPart.RevNum);
                        oSLBOMDocument.SetCellValue(iNumberOfRows, 3, oPartMaterial.PartMaster.PartNum);
                        oSLBOMDocument.SetCellValue(iNumberOfRows, 4, oPartMaterial.MtlSeq);

                        iNumberOfRows++;
                        bDataInReport = true;
                    }
                }
            }

            // materials with zero cost
            iNumberOfRows = 1;
            iNumberOfColumns = 1;
            if (m_oMfgPartsWithMaterialsWithZeroCost.Count != 0)
            {
                if (bFirstWorksheet == true)
                {
                    oSLBOMDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "Mtl Zero Cost");
                    bFirstWorksheet = false;
                }
                else
                {
                    oSLBOMDocument.AddWorksheet("Mtl Zero Cost");
                }
                //set column header
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Parent Part");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 30);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Revision");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 30);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Mtl");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 30);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Mtl Seq");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 30);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "The Material On The Mfg Part Has The Cost Set To Zero");
                oSLBOMDocument.SetCellStyle(iNumberOfRows++, iNumberOfColumns, oHighlightHeaderStyle);
                oSLBOMDocument.SetColumnWidth(iNumberOfColumns++, 90);

                foreach (MfgPart oMfgPart in m_oMfgPartsWithMaterialsWithZeroCost)
                {
                    // find all materials that have zero cost
                    List<PartMaterial> oPartsWithZeroCost = oMfgPart.MyPartMaterials.Where(oItem => oItem.PartMaster.Cost == 0).ToList();
                    foreach (PartMaterial oPartMaterial in oPartsWithZeroCost)
                    {
                        oSLBOMDocument.SetCellValue(iNumberOfRows, 1, oMfgPart.PartNum);
                        oSLBOMDocument.SetCellValue(iNumberOfRows, 2, oMfgPart.RevNum);
                        oSLBOMDocument.SetCellValue(iNumberOfRows, 3, oPartMaterial.PartMaster.PartNum);
                        oSLBOMDocument.SetCellValue(iNumberOfRows, 4, oPartMaterial.MtlSeq);

                        iNumberOfRows++;
                        bDataInReport = true;
                    }
                }
            }

            // materials that are manufactured but have view as assembly turned off
            iNumberOfRows = 1;
            iNumberOfColumns = 1;
            if (m_oMfgPartMaterialsNeedViewAsAssemblyTurnedOn.Count != 0)
            {
                if (bFirstWorksheet == true)
                {
                    oSLBOMDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "Mfg Mtl Needs View Assembly On");
                    bFirstWorksheet = false;
                }
                else
                {
                    oSLBOMDocument.AddWorksheet("Mfg Mtl Needs View Assembly On");
                }
                //set column header
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Parent Part");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 30);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Revision");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 30);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Mtl");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 30);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "The Material On The Mfg Part Is A Subassembly But View As Assembly Is Off");
                oSLBOMDocument.SetCellStyle(iNumberOfRows++, iNumberOfColumns, oHighlightHeaderStyle);
                oSLBOMDocument.SetColumnWidth(iNumberOfColumns++, 90);

                foreach (MfgPart oMfgPart in m_oMfgPartMaterialsNeedViewAsAssemblyTurnedOn)
                {
                    // find all materials that are Mfg but have view as assembly turned off
                    List<PartMaterial> oMfgMtlWithViewAsAssemblyOff = oMfgPart.MyPartMaterials.Where(oItem => (string.Compare(oItem.TypeCode, "M", true) == 0) && (oItem.ViewAsAsm == false)).ToList();
                    foreach (PartMaterial oPartMaterial in oMfgMtlWithViewAsAssemblyOff)
                    {
                        oSLBOMDocument.SetCellValue(iNumberOfRows, 1, oMfgPart.PartNum);
                        oSLBOMDocument.SetCellValue(iNumberOfRows, 2, oMfgPart.RevNum);
                        oSLBOMDocument.SetCellValue(iNumberOfRows, 3, oPartMaterial.PartMaster.PartNum);

                        iNumberOfRows++;
                        bDataInReport = true;
                    }
                }
            }

            // purchased materials need view as assembly turned off
            iNumberOfRows = 1;
            iNumberOfColumns = 1;
            if (m_oMfgPartMaterialsNeedViewAsAssemblyTurnedOff.Count != 0)
            {
                if (bFirstWorksheet == true)
                {
                    oSLBOMDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "Pur Mtl Needs View Assembly Off");
                    bFirstWorksheet = false;
                }
                else
                {
                    oSLBOMDocument.AddWorksheet("Pur Mtl Needs View Assembly Off");
                }
                //set column header
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Parent Part");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 30);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Revision");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 30);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Mtl");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 30);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "The Material On The Mfg Part Is Purchased But View As Assembly Is On");
                oSLBOMDocument.SetCellStyle(iNumberOfRows++, iNumberOfColumns, oHighlightHeaderStyle);
                oSLBOMDocument.SetColumnWidth(iNumberOfColumns++, 90);

                foreach (MfgPart oMfgPart in m_oMfgPartMaterialsNeedViewAsAssemblyTurnedOff)
                {
                    // find all materials that are purchased parts that have view as assembly turned on
                    List<PartMaterial> oPartsWithPlanAsAssemblyTurnedOn = oMfgPart.MyPartMaterials.Where(oItem => (string.Compare(oItem.TypeCode, "P", true) == 0) && (oItem.ViewAsAsm == true)).ToList();
                    foreach (PartMaterial oPartMaterial in oPartsWithPlanAsAssemblyTurnedOn)
                    {
                        oSLBOMDocument.SetCellValue(iNumberOfRows, 1, oMfgPart.PartNum);
                        oSLBOMDocument.SetCellValue(iNumberOfRows, 2, oMfgPart.RevNum);
                        oSLBOMDocument.SetCellValue(iNumberOfRows, 3, oPartMaterial.PartMaster.PartNum);

                        iNumberOfRows++;
                        bDataInReport = true;
                    }
                }
            }

            // materials with plan as assembly turned on
            iNumberOfRows = 1;
            iNumberOfColumns = 1;
            if (m_oMfgPartMaterialsHasPlanAsAssemblyTurnedOn.Count != 0)
            {
                if (bFirstWorksheet == true)
                {
                    oSLBOMDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "Plan Assembly On");
                    bFirstWorksheet = false;
                }
                else
                {
                    oSLBOMDocument.AddWorksheet("Plan Assembly On");
                }
                //set column header
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Parent Part");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 30);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Revision");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 30);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Mtl");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 30);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "The Material On The Mfg Part Is Has Plan As Assembly Turned On");
                oSLBOMDocument.SetCellStyle(iNumberOfRows++, iNumberOfColumns, oHighlightHeaderStyle);
                oSLBOMDocument.SetColumnWidth(iNumberOfColumns++, 90);

                foreach (MfgPart oMfgPart in m_oMfgPartMaterialsHasPlanAsAssemblyTurnedOn)
                {
                    // find all materials on this part that are using Plan As Assembly
                    List<PartMaterial> oPartsWithPlanAsAssemblyTurnedOn = oMfgPart.MyPartMaterials.Where(oItem => oItem.PlanAsAsm == true).ToList();
                    foreach (PartMaterial oPartMaterial in oPartsWithPlanAsAssemblyTurnedOn)
                    {
                        oSLBOMDocument.SetCellValue(iNumberOfRows, 1, oMfgPart.PartNum);
                        oSLBOMDocument.SetCellValue(iNumberOfRows, 2, oMfgPart.RevNum);
                        oSLBOMDocument.SetCellValue(iNumberOfRows, 3, oPartMaterial.PartMaster.PartNum);

                        iNumberOfRows++;
                        bDataInReport = true;
                    }
                }
            }


            if (bDataInReport == true)
            {
                oSLBOMDocument.SaveAs(sDestinationFileName);
                // Check to see if we created a file and if so email it
                if (File.Exists(sDestinationFileName) == true)
                {
                    List<string> oAttachments = new List<string>();
                    oAttachments.Add(sDestinationFileName);

                    HSEmailHelper.SendEmail(oToAddresses, "Mfg Parts BOM Report", "Mfg Parts BOM Report for " + sDate, oAttachments);
                }
            }
            #endregion

        }

        public void PerformOperationValidation(string sCompany, string sTmpFileDirectory)
        {
            // get the file name
            DateTime dtToday = DateTime.Now;
            string sDate = dtToday.Month.ToString() + "-" + dtToday.Day.ToString() + "-" + dtToday.Year.ToString();
            string sDestinationFileName = sTmpFileDirectory + "\\" + sCompany + "-OperationsReport-" + sDate + ".xlsx";
            int iNumberOfRows = 1;
            int iNumberOfColumns = 1;

            HSUser oServiceAccount = HSUser.GetUserById(HSUser.SPINCRAFT_SERVICE_ACCOUNT_ID);
            List<string> oToAddresses = new List<string>();
            oToAddresses = HSUser.GetEmailsForUsersInGroup(HSUser.REPORT_ON_MANUFACTURED_PART_ISSUES);
            oToAddresses.Add(HSUser.SPINCRAFT_ROOT_USER);

            SLFill oSLFill = new SLFill();
            oSLFill.SetPatternBackgroundColor(SLThemeColorIndexValues.Accent1Color);
            oSLFill.SetPattern(PatternValues.Solid, SLThemeColorIndexValues.Light2Color, SLThemeColorIndexValues.Light2Color);

            SLStyle oBoldStyle = new SLStyle();
            oBoldStyle.SetFontBold(true);
            oBoldStyle.SetTopBorder(BorderStyleValues.Thin, SLThemeColorIndexValues.Dark2Color);
            oBoldStyle.SetBottomBorder(BorderStyleValues.Thin, SLThemeColorIndexValues.Dark2Color);
            oBoldStyle.SetLeftBorder(BorderStyleValues.Thin, SLThemeColorIndexValues.Dark2Color);
            oBoldStyle.SetRightBorder(BorderStyleValues.Thin, SLThemeColorIndexValues.Dark2Color);
            oBoldStyle.SetFontColor(SLThemeColorIndexValues.Dark2Color);

            SLStyle oBoldStyle2 = new SLStyle();
            oBoldStyle2.SetFontBold(true);
            oBoldStyle2.SetTopBorder(BorderStyleValues.Thick, SLThemeColorIndexValues.Dark2Color);
            oBoldStyle2.SetBottomBorder(BorderStyleValues.Thick, SLThemeColorIndexValues.Dark2Color);
            oBoldStyle2.SetLeftBorder(BorderStyleValues.Thick, SLThemeColorIndexValues.Dark2Color);
            oBoldStyle.SetRightBorder(BorderStyleValues.Thick, SLThemeColorIndexValues.Dark2Color);
            oBoldStyle2.SetFontColor(System.Drawing.Color.IndianRed);

            System.Drawing.Color oHighlightYellow = System.Drawing.Color.FromArgb(255, 255, 0);

            SLStyle oHighlightHeaderStyle = new SLStyle();
            oHighlightHeaderStyle.SetFontBold(true);
            oHighlightHeaderStyle.SetFont(FontSchemeValues.Major, 12);
            oHighlightHeaderStyle.SetHorizontalAlignment(HorizontalAlignmentValues.Center);
            oHighlightHeaderStyle.SetPatternFill(PatternValues.Solid, oHighlightYellow, oHighlightYellow);
            oHighlightHeaderStyle.SetFontColor(SLThemeColorIndexValues.Dark2Color);

            //
            // operation issues
            //
            oToAddresses = HSUser.GetEmailsForUsersInGroup(HSUser.REPORT_ON_MANUFACTURED_PART_ISSUES);
            oToAddresses.Add(HSUser.SPINCRAFT_ROOT_USER);
            bool bDataInReport = false;
            bool bFirstWorksheet = true;

            // if the file is already there then get rid of it as we are reprocessing for some reason
            if (File.Exists(sDestinationFileName) == true)
            {
                try
                {
                    File.Delete(sDestinationFileName);
                }
                catch (Exception)
                {
                    // we will ingore this if we cannot delete the file
                }
            }
            SLDocument oSLBOMDocument = new SLDocument();

            // no op code description
            List<HSOperation> oOperationsMissingOpCode = m_oOperations.Where(oItem => string.IsNullOrEmpty(oItem.OpCode) == true).ToList();
            if (oOperationsMissingOpCode.Count > 0)
            {
                iNumberOfRows = 1;
                iNumberOfColumns = 1;
                if (bFirstWorksheet == true)
                {
                    oSLBOMDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "Missing Description");
                    bFirstWorksheet = false;
                }
                else
                {
                    oSLBOMDocument.AddWorksheet("Missing Description");
                }
                //set column header
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Operation");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 30);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "The Operation Has No Description");
                oSLBOMDocument.SetCellStyle(iNumberOfRows++, iNumberOfColumns, oHighlightHeaderStyle);
                oSLBOMDocument.SetColumnWidth(iNumberOfColumns++, 90);

                foreach (HSOperation oOperation in oOperationsMissingOpCode)
                {
                    oSLBOMDocument.SetCellValue(iNumberOfRows++, 1, oOperation.OpCode);
                    bDataInReport = true;
                }
            }

            // labor rate set
            List<HSOperation> oOperationsWithBillLaborRateSet = m_oOperations.Where(oItem => oItem.BilLLaborRate != 0).ToList();
            if (oOperationsWithBillLaborRateSet.Count > 0)
            {
                iNumberOfRows = 1;
                iNumberOfColumns = 1;
                if (bFirstWorksheet == true)
                {
                    oSLBOMDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "Labor Rate Set");
                    bFirstWorksheet = false;
                }
                else
                {
                    oSLBOMDocument.AddWorksheet("Labor Rate Set");
                }
                //set column header
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Operation");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 30);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Description");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 30);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Bill Rate");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 30);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "The Operation Should Not Set The Bill Labor Rate");
                oSLBOMDocument.SetCellStyle(iNumberOfRows++, iNumberOfColumns, oHighlightHeaderStyle);
                oSLBOMDocument.SetColumnWidth(iNumberOfColumns++, 90);

                foreach (HSOperation oOperation in oOperationsWithBillLaborRateSet)
                {
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 1, oOperation.OpCode);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 2, oOperation.Description);
                    oSLBOMDocument.SetCellValue(iNumberOfRows++, 3, oOperation.BilLLaborRate);
                    bDataInReport = true;
                }
            }

            // labor hours set
            List<HSOperation> oOperationsWithLaborHoursSet = m_oOperations.Where(oItem => oItem.EstLaborHours != 0).ToList();
            if (oOperationsWithLaborHoursSet.Count > 0)
            {
                iNumberOfRows = 1;
                iNumberOfColumns = 1;
                if (bFirstWorksheet == true)
                {
                    oSLBOMDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "Labor Hours Set");
                    bFirstWorksheet = false;
                }
                else
                {
                    oSLBOMDocument.AddWorksheet("Labor Hours Set");
                }
                //set column header
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Operation");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 30);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Description");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 30);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Labor Hours");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 30);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "The Operation Should Not Set The Labor Hours");
                oSLBOMDocument.SetCellStyle(iNumberOfRows++, iNumberOfColumns, oHighlightHeaderStyle);
                oSLBOMDocument.SetColumnWidth(iNumberOfColumns++, 90);

                foreach (HSOperation oOperation in oOperationsWithLaborHoursSet)
                {
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 1, oOperation.OpCode);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 2, oOperation.Description);
                    oSLBOMDocument.SetCellValue(iNumberOfRows++, 3, oOperation.EstLaborHours);
                    bDataInReport = true;
                }
            }

            // primary supplier should only be set on non-subcontract
            List<HSOperation> oNonSubcontractWithSupplierSet = m_oOperations.Where(oItem => (oItem.VendorNum != 0) && (oItem.Subcontract == false)).ToList();
            if (oNonSubcontractWithSupplierSet.Count > 0)
            {
                iNumberOfRows = 1;
                iNumberOfColumns = 1;
                if (bFirstWorksheet == true)
                {
                    oSLBOMDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "Non-Subcontract With Vendor");
                    bFirstWorksheet = false;
                }
                else
                {
                    oSLBOMDocument.AddWorksheet("Non-Subcontract With Vendor");
                }
                //set column header
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Operation");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 30);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Description");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 30);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Vendor");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 30);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "The Vendor Is Set On An Operation That Is Not A Subcontract");
                oSLBOMDocument.SetCellStyle(iNumberOfRows++, iNumberOfColumns, oHighlightHeaderStyle);
                oSLBOMDocument.SetColumnWidth(iNumberOfColumns++, 90);

                foreach (HSOperation oOperation in oNonSubcontractWithSupplierSet)
                {
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 1, oOperation.OpCode);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 2, oOperation.Description);
                    oSLBOMDocument.SetCellValue(iNumberOfRows++, 3, oOperation.VendorNum);
                    bDataInReport = true;
                }
            }

            // send ahead not set to Hours
            List<HSOperation> oIncorrectSendAheadType = m_oOperations.Where(oItem => string.Compare(oItem.SendAheadType, "HOURS", true) != 0).ToList();
            if (oIncorrectSendAheadType.Count > 0)
            {
                iNumberOfRows = 1;
                iNumberOfColumns = 1;
                if (bFirstWorksheet == true)
                {
                    oSLBOMDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "Incorrect Send Ahead Type");
                    bFirstWorksheet = false;
                }
                else
                {
                    oSLBOMDocument.AddWorksheet("Incorrect Send Ahead Type");
                }
                //set column header
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Operation");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 30);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Description");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 30);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Send Ahead Type");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 30);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "The The Send Ahead Type Should Be Set To Hours");
                oSLBOMDocument.SetCellStyle(iNumberOfRows++, iNumberOfColumns, oHighlightHeaderStyle);
                oSLBOMDocument.SetColumnWidth(iNumberOfColumns++, 90);

                foreach (HSOperation oOperation in oIncorrectSendAheadType)
                {
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 1, oOperation.OpCode);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 2, oOperation.Description);
                    oSLBOMDocument.SetCellValue(iNumberOfRows++, 3, oOperation.SendAheadType);
                    bDataInReport = true;
                }
            }

            // send ahead offset not ZERO
            List<HSOperation> oSendAheadOffsetNotZero = m_oOperations.Where(oItem => oItem.SendAheadOffset != 0).ToList();
            if (oSendAheadOffsetNotZero.Count > 0)
            {
                iNumberOfRows = 1;
                iNumberOfColumns = 1;
                if (bFirstWorksheet == true)
                {
                    oSLBOMDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "Send Ahead Offset Not Zero");
                    bFirstWorksheet = false;
                }
                else
                {
                    oSLBOMDocument.AddWorksheet("Send Ahead Offset Not Zero");
                }
                //set column header
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Operation");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 30);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Description");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 30);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Send Ahead Offset");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 30);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "The Send Ahead Offset Should Be Set To Zero");
                oSLBOMDocument.SetCellStyle(iNumberOfRows++, iNumberOfColumns, oHighlightHeaderStyle);
                oSLBOMDocument.SetColumnWidth(iNumberOfColumns++, 90);

                foreach (HSOperation oOperation in oSendAheadOffsetNotZero)
                {
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 1, oOperation.OpCode);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 2, oOperation.Description);
                    oSLBOMDocument.SetCellValue(iNumberOfRows++, 3, oOperation.SendAheadOffset);
                    bDataInReport = true;
                }
            }

            // resource id should not be set
            List<HSOperation> oResoruceSetOnOperation = m_oOperations.Where(oItem => string.IsNullOrEmpty(oItem.ResourceId) != true).ToList();
            if (oResoruceSetOnOperation.Count > 0)
            {
                iNumberOfRows = 1;
                iNumberOfColumns = 1;
                if (bFirstWorksheet == true)
                {
                    oSLBOMDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "Resource Set On Operation");
                    bFirstWorksheet = false;
                }
                else
                {
                    oSLBOMDocument.AddWorksheet("Resource Set On Operation");
                }
                //set column header
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Operation");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 30);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Description");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 30);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Resource");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 30);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "The Operation Should Only Set The Resource Group Not The Resource");
                oSLBOMDocument.SetCellStyle(iNumberOfRows++, iNumberOfColumns, oHighlightHeaderStyle);
                oSLBOMDocument.SetColumnWidth(iNumberOfColumns++, 90);

                foreach (HSOperation oOperation in oResoruceSetOnOperation)
                {
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 1, oOperation.OpCode);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 2, oOperation.Description);
                    oSLBOMDocument.SetCellValue(iNumberOfRows++, 3, oOperation.ResourceId);
                    bDataInReport = true;
                }
            }

            // if resoruce group id set then primary production and primary setup should be true
            List<HSOperation> oResourceGroupNotPrimary = m_oOperations.Where(oItem => (string.IsNullOrEmpty(oItem.ResourceGroupId) != true) && ((oItem.PrimaryProduction == false) || (oItem.PrimarySetup == false))).ToList();
            if (oResourceGroupNotPrimary.Count > 0)
            {
                iNumberOfRows = 1;
                iNumberOfColumns = 1;
                if (bFirstWorksheet == true)
                {
                    oSLBOMDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "Resource Group Not Primary");
                    bFirstWorksheet = false;
                }
                else
                {
                    oSLBOMDocument.AddWorksheet("Resource Group Not Primary");
                }
                //set column header
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Operation");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 30);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Description");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 30);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Resource Group");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 30);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "The Resource Group For The Operation Should Be set As Primary");
                oSLBOMDocument.SetCellStyle(iNumberOfRows++, iNumberOfColumns, oHighlightHeaderStyle);
                oSLBOMDocument.SetColumnWidth(iNumberOfColumns++, 90);

                foreach (HSOperation oOperation in oResourceGroupNotPrimary)
                {
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 1, oOperation.OpCode);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 2, oOperation.Description);
                    oSLBOMDocument.SetCellValue(iNumberOfRows++, 3, oOperation.ResourceGroupId);
                    bDataInReport = true;
                }
            }

            // setup hours not set to ZERO
            List<HSOperation> oSetupHoursNotZero = m_oOperations.Where(oItem => oItem.SetupHours != 0).ToList();
            if (oSetupHoursNotZero.Count > 0)
            {
                iNumberOfRows = 1;
                iNumberOfColumns = 1;
                if (bFirstWorksheet == true)
                {
                    oSLBOMDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "Setup Hours Not Zero");
                    bFirstWorksheet = false;
                }
                else
                {
                    oSLBOMDocument.AddWorksheet("Setup Hours Not Zero");
                }
                //set column header
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Operation");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 30);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Description");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 30);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Setup Hours");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 30);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "The Operation Setup Hours Should Be Zero");
                oSLBOMDocument.SetCellStyle(iNumberOfRows++, iNumberOfColumns, oHighlightHeaderStyle);
                oSLBOMDocument.SetColumnWidth(iNumberOfColumns++, 90);

                foreach (HSOperation oOperation in oSetupHoursNotZero)
                {
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 1, oOperation.OpCode);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 2, oOperation.Description);
                    oSLBOMDocument.SetCellValue(iNumberOfRows++, 3, oOperation.SetupHours);
                    bDataInReport = true;
                }
            }

            // production hours not set to ZERO
            List<HSOperation> oPrpductionHoursNotZero = m_oOperations.Where(oItem => oItem.ProductionHours != 0).ToList();
            if (oPrpductionHoursNotZero.Count > 0)
            {
                iNumberOfRows = 1;
                iNumberOfColumns = 1;
                if (bFirstWorksheet == true)
                {
                    oSLBOMDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "Production Hours Not Zero");
                    bFirstWorksheet = false;
                }
                else
                {
                    oSLBOMDocument.AddWorksheet("Production Hours Not Zero");
                }
                //set column header
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Operation");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 30);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Description");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 30);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Production Hours");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 30);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "The Production Hours Should Be Set to Zero");
                oSLBOMDocument.SetCellStyle(iNumberOfRows++, iNumberOfColumns, oHighlightHeaderStyle);
                oSLBOMDocument.SetColumnWidth(iNumberOfColumns++, 90);

                foreach (HSOperation oOperation in oPrpductionHoursNotZero)
                {
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 1, oOperation.OpCode);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 2, oOperation.Description);
                    oSLBOMDocument.SetCellValue(iNumberOfRows++, 3, oOperation.ProductionHours);
                    bDataInReport = true;
                }
            }

            // scheduling blocks not set to 1
            List<HSOperation> oSchedulingBlocksNotSetToOne = m_oOperations.Where(oItem => oItem.SchedulingBlocks != 1).ToList();
            if (oSchedulingBlocksNotSetToOne.Count > 0)
            {
                iNumberOfRows = 1;
                iNumberOfColumns = 1;
                if (bFirstWorksheet == true)
                {
                    oSLBOMDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "Scheduling Blocks Not One");
                    bFirstWorksheet = false;
                }
                else
                {
                    oSLBOMDocument.AddWorksheet("Scheduling Blocks Not One");
                }
                //set column header
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Operation");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 30);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Description");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 30);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Scheduling Blocks");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 30);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "The Scheduling Blocks Should Be Set To One");
                oSLBOMDocument.SetCellStyle(iNumberOfRows++, iNumberOfColumns, oHighlightHeaderStyle);
                oSLBOMDocument.SetColumnWidth(iNumberOfColumns++, 90);

                foreach (HSOperation oOperation in oSchedulingBlocksNotSetToOne)
                {
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 1, oOperation.OpCode);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 2, oOperation.Description);
                    oSLBOMDocument.SetCellValue(iNumberOfRows++, 3, oOperation.SchedulingBlocks);
                    bDataInReport = true;
                }
            }

            // concurrent capacity not set to ZERO
            List<HSOperation> oConcurrentCapacityNotSetToZero = m_oOperations.Where(oItem => oItem.ConcurrentCapacity != 0).ToList();
            if (oConcurrentCapacityNotSetToZero.Count > 0)
            {
                iNumberOfRows = 1;
                iNumberOfColumns = 1;
                if (bFirstWorksheet == true)
                {
                    oSLBOMDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "Concurrent Capacity Not Zero");
                    bFirstWorksheet = false;
                }
                else
                {
                    oSLBOMDocument.AddWorksheet("Concurrent Capacity Not Zero");
                }
                //set column header
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Operation");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 30);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Description");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 30);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Concurrent Capacity");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 30);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "The Concurrent Capacity Should Be Zero");
                oSLBOMDocument.SetCellStyle(iNumberOfRows++, iNumberOfColumns, oHighlightHeaderStyle);
                oSLBOMDocument.SetColumnWidth(iNumberOfColumns++, 90);

                foreach (HSOperation oOperation in oConcurrentCapacityNotSetToZero)
                {
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 1, oOperation.OpCode);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 2, oOperation.Description);
                    oSLBOMDocument.SetCellValue(iNumberOfRows++, 3, oOperation.ConcurrentCapacity);
                    bDataInReport = true;
                }
            }

            // daily production rate not set to ZERO
            List<HSOperation> oDailyProductionRateNotSetToZero = m_oOperations.Where(oItem => oItem.DailyProductionRate != 0).ToList();
            if (oDailyProductionRateNotSetToZero.Count > 0)
            {
                iNumberOfRows = 1;
                iNumberOfColumns = 1;
                if (bFirstWorksheet == true)
                {
                    oSLBOMDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "Daily Production Rate Not Zero");
                    bFirstWorksheet = false;
                }
                else
                {
                    oSLBOMDocument.AddWorksheet("Daily Production Rate Not Zero");
                }
                //set column header
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Operation");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 30);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Description");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 30);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Daily Production Rate");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 30);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "The Daily Production Rate Should Not Be Zero");
                oSLBOMDocument.SetCellStyle(iNumberOfRows++, iNumberOfColumns, oHighlightHeaderStyle);
                oSLBOMDocument.SetColumnWidth(iNumberOfColumns++, 90);

                foreach (HSOperation oOperation in oDailyProductionRateNotSetToZero)
                {
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 1, oOperation.OpCode);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 2, oOperation.Description);
                    oSLBOMDocument.SetCellValue(iNumberOfRows++, 3, oOperation.DailyProductionRate);
                    bDataInReport = true;
                }
            }

            if (bDataInReport == true)
            {
                oSLBOMDocument.SaveAs(sDestinationFileName);
                // Check to see if we created a file and if so email it
                if (File.Exists(sDestinationFileName) == true)
                {
                    List<string> oAttachments = new List<string>();
                    oAttachments.Add(sDestinationFileName);

                    HSEmailHelper.SendEmail(oToAddresses, "Operations Report", "Operations BOM Report for " + sDate, oAttachments);
                }
            }
        }

        public void PerformResourceGroupValidation(string sCompany, string sTmpFileDirectory)
        {
            // get the file name
            DateTime dtToday = DateTime.Now;
            string sDate = dtToday.Month.ToString() + "-" + dtToday.Day.ToString() + "-" + dtToday.Year.ToString();
            string sDestinationFileName = sTmpFileDirectory + "\\" + sCompany + "-ResourceGroupReport-" + sDate + ".xlsx";
            int iNumberOfRows = 1;
            int iNumberOfColumns = 1;

            HSUser oServiceAccount = HSUser.GetUserById(HSUser.SPINCRAFT_SERVICE_ACCOUNT_ID);
            List<string> oToAddresses = new List<string>();
            oToAddresses = HSUser.GetEmailsForUsersInGroup(HSUser.REPORT_ON_MANUFACTURED_PART_ISSUES);
            oToAddresses.Add(HSUser.SPINCRAFT_ROOT_USER);

            SLFill oSLFill = new SLFill();
            oSLFill.SetPatternBackgroundColor(SLThemeColorIndexValues.Accent1Color);
            oSLFill.SetPattern(PatternValues.Solid, SLThemeColorIndexValues.Light2Color, SLThemeColorIndexValues.Light2Color);

            SLStyle oBoldStyle = new SLStyle();
            oBoldStyle.SetFontBold(true);
            oBoldStyle.SetTopBorder(BorderStyleValues.Thin, SLThemeColorIndexValues.Dark2Color);
            oBoldStyle.SetBottomBorder(BorderStyleValues.Thin, SLThemeColorIndexValues.Dark2Color);
            oBoldStyle.SetLeftBorder(BorderStyleValues.Thin, SLThemeColorIndexValues.Dark2Color);
            oBoldStyle.SetRightBorder(BorderStyleValues.Thin, SLThemeColorIndexValues.Dark2Color);
            oBoldStyle.SetFontColor(SLThemeColorIndexValues.Dark2Color);

            SLStyle oBoldStyle2 = new SLStyle();
            oBoldStyle2.SetFontBold(true);
            oBoldStyle2.SetTopBorder(BorderStyleValues.Thick, SLThemeColorIndexValues.Dark2Color);
            oBoldStyle2.SetBottomBorder(BorderStyleValues.Thick, SLThemeColorIndexValues.Dark2Color);
            oBoldStyle2.SetLeftBorder(BorderStyleValues.Thick, SLThemeColorIndexValues.Dark2Color);
            oBoldStyle.SetRightBorder(BorderStyleValues.Thick, SLThemeColorIndexValues.Dark2Color);
            oBoldStyle2.SetFontColor(System.Drawing.Color.IndianRed);

            System.Drawing.Color oHighlightYellow = System.Drawing.Color.FromArgb(255, 255, 0);

            SLStyle oHighlightHeaderStyle = new SLStyle();
            oHighlightHeaderStyle.SetFontBold(true);
            oHighlightHeaderStyle.SetFont(FontSchemeValues.Major, 12);
            oHighlightHeaderStyle.SetHorizontalAlignment(HorizontalAlignmentValues.Center);
            oHighlightHeaderStyle.SetPatternFill(PatternValues.Solid, oHighlightYellow, oHighlightYellow);
            oHighlightHeaderStyle.SetFontColor(SLThemeColorIndexValues.Dark2Color);

            oToAddresses = HSUser.GetEmailsForUsersInGroup(HSUser.REPORT_ON_MANUFACTURED_PART_ISSUES);
            oToAddresses.Add(HSUser.SPINCRAFT_ROOT_USER);
            bool bDataInReport = false;
            bool bFirstWorksheet = true;

            // if the file is already there then get rid of it as we are reprocessing for some reason
            if (File.Exists(sDestinationFileName) == true)
            {
                try
                {
                    File.Delete(sDestinationFileName);
                }
                catch (Exception)
                {
                    // we will ingore this if we cannot delete the file
                }
            }
            SLDocument oSLBOMDocument = new SLDocument();

            List<HSResourceGroup> oActiveResourceGroups = m_oResourceGroups.Where(oItem => oItem.Inactive == false).ToList();

            // no description
            List<HSResourceGroup> oNoResourceDescription = oActiveResourceGroups.Where(oItem => string.IsNullOrEmpty(oItem.Description) == true).ToList();
            if (oNoResourceDescription.Count > 0)
            {
                iNumberOfRows = 1;
                iNumberOfColumns = 1;
                if (bFirstWorksheet == true)
                {
                    oSLBOMDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "No Description");
                    bFirstWorksheet = false;
                }
                else
                {
                    oSLBOMDocument.AddWorksheet("No Description");
                }
                //set column header
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Resource Group");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 30);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "The Resource Group Has No Description");
                oSLBOMDocument.SetCellStyle(iNumberOfRows++, iNumberOfColumns, oHighlightHeaderStyle);
                oSLBOMDocument.SetColumnWidth(iNumberOfColumns++, 90);

                foreach (HSResourceGroup oResourceGroup in oNoResourceDescription)
                {
                    oSLBOMDocument.SetCellValue(iNumberOfRows++, 1, oResourceGroup.ResourceGroupId);
                    bDataInReport = true;
                }
            }

            // not a location
            List<HSResourceGroup> oNotALocation = oActiveResourceGroups.Where(oItem => oItem.Location == false).ToList();
            if (oNotALocation.Count > 0)
            {
                iNumberOfRows = 1;
                iNumberOfColumns = 1;
                if (bFirstWorksheet == true)
                {
                    oSLBOMDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "Not Location");
                    bFirstWorksheet = false;
                }
                else
                {
                    oSLBOMDocument.AddWorksheet("Not Location");
                }
                //set column header
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Resource Group");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 30);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Description");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 30);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "The Resource Group Has Not Been Set As A Location");
                oSLBOMDocument.SetCellStyle(iNumberOfRows++, iNumberOfColumns, oHighlightHeaderStyle);
                oSLBOMDocument.SetColumnWidth(iNumberOfColumns++, 90);

                foreach (HSResourceGroup oResourceGroup in oNotALocation)
                {
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 1, oResourceGroup.ResourceGroupId);
                    oSLBOMDocument.SetCellValue(iNumberOfRows++, 2, oResourceGroup.Description);
                    bDataInReport = true;
                }
            }

            //// no op code set
            //List<HSResourceGroup> oNoOpCode = oActiveResourceGroups.Where(oItem => string.IsNullOrEmpty(oItem.OpCode) == true).ToList();
            //if (oNoOpCode.Count > 0)
            //{
            //    iNumberOfRows = 1;
            //    iNumberOfColumns = 1;
            //    if (bFirstWorksheet == true)
            //    {
            //        oSLBOMDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "No Op Code");
            //        bFirstWorksheet = false;
            //    }
            //    else
            //    {
            //        oSLBOMDocument.AddWorksheet("No Op Code");
            //    }
            //    //set column header
            //    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Resource Group");
            //    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 30);
            //    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Description");
            //    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 30);
            //    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "The Op Code Has Not Been Set");
            //    oSLBOMDocument.SetCellStyle(iNumberOfRows++, iNumberOfColumns, oHighlightHeaderStyle);
            //    oSLBOMDocument.SetColumnWidth(iNumberOfColumns++, 90);
            //    foreach (HSResourceGroup oResourceGroup in oNoOpCode)
            //    {
            //        oSLBOMDocument.SetCellValue(iNumberOfRows, 1, oResourceGroup.ResourceGroupId);
            //        oSLBOMDocument.SetCellValue(iNumberOfRows++, 2, oResourceGroup.Description);
            //        bDataInReport = true;
            //    }
            //}

            // no calendar set
            List<HSResourceGroup> oNoCalendar = oActiveResourceGroups.Where(oItem => string.IsNullOrEmpty(oItem.CalendarId) == true).ToList();
            if (oNoCalendar.Count > 0)
            {
                iNumberOfRows = 1;
                iNumberOfColumns = 1;
                if (bFirstWorksheet == true)
                {
                    oSLBOMDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "No Calendar Set");
                    bFirstWorksheet = false;
                }
                else
                {
                    oSLBOMDocument.AddWorksheet("No Calendar Set");
                }
                //set column header
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Resource Group");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 30);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Description");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 30);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "The Resource Group Is Not Using A Calendar");
                oSLBOMDocument.SetCellStyle(iNumberOfRows++, iNumberOfColumns, oHighlightHeaderStyle);
                oSLBOMDocument.SetColumnWidth(iNumberOfColumns++, 90);

                foreach (HSResourceGroup oResourceGroup in oNoCalendar)
                {
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 1, oResourceGroup.ResourceGroupId);
                    oSLBOMDocument.SetCellValue(iNumberOfRows++, 2, oResourceGroup.Description);
                    bDataInReport = true;
                }
            }

            // no department set
            List<HSResourceGroup> oNoDepartment = oActiveResourceGroups.Where(oItem => string.IsNullOrEmpty(oItem.JCDepartment) == true).ToList();
            if (oNoDepartment.Count > 0)
            {
                iNumberOfRows = 1;
                iNumberOfColumns = 1;
                if (bFirstWorksheet == true)
                {
                    oSLBOMDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "No Department Set");
                    bFirstWorksheet = false;
                }
                else
                {
                    oSLBOMDocument.AddWorksheet("No Department Set");
                }
                //set column header
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Resource Group");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 30);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Description");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 30);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "The Resource Group Does Not Have A Job Department Set");
                oSLBOMDocument.SetCellStyle(iNumberOfRows++, iNumberOfColumns, oHighlightHeaderStyle);
                oSLBOMDocument.SetColumnWidth(iNumberOfColumns++, 90);

                foreach (HSResourceGroup oResourceGroup in oNoDepartment)
                {
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 1, oResourceGroup.ResourceGroupId);
                    oSLBOMDocument.SetCellValue(iNumberOfRows++, 2, oResourceGroup.Description);
                    bDataInReport = true;
                }
            }

            // subcontract with labor rates
            List<HSResourceGroup> oSubcontractWithRates = oActiveResourceGroups.Where(oItem => (oItem.Subcontract == true) &&
                ((oItem.SetupLaborRate != 0) || (oItem.SetupBurdenRate != 0) || (oItem.ProductionLaborRate != 0) || (oItem.ProductionBurdenRate != 0))).ToList();
            if (oSubcontractWithRates.Count > 0)
            {
                iNumberOfRows = 1;
                iNumberOfColumns = 1;
                if (bFirstWorksheet == true)
                {
                    oSLBOMDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "Subcontract With Rates");
                    bFirstWorksheet = false;
                }
                else
                {
                    oSLBOMDocument.AddWorksheet("Subcontract With Rates");
                }
                //set column header
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Resource Group");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 30);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Description");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 30);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Setup Labor");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 30);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Setup Burden");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 30);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Production Labor");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 30);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Production Burden");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 30);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "The Resource Group Is A Subcontract But Has The Labor Or Burden Rate Set");
                oSLBOMDocument.SetCellStyle(iNumberOfRows++, iNumberOfColumns, oHighlightHeaderStyle);
                oSLBOMDocument.SetColumnWidth(iNumberOfColumns++, 90);

                foreach (HSResourceGroup oResourceGroup in oSubcontractWithRates)
                {
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 1, oResourceGroup.ResourceGroupId);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 2, oResourceGroup.Description);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 3, oResourceGroup.SetupLaborRate);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 4, oResourceGroup.SetupBurdenRate);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 5, oResourceGroup.ProductionLaborRate);
                    oSLBOMDocument.SetCellValue(iNumberOfRows++, 6, oResourceGroup.ProductionBurdenRate);
                    bDataInReport = true;
                }
            }

            // subcontract with quote labor rates
            List<HSResourceGroup> oSubcontractWithQuoteRates = oActiveResourceGroups.Where(oItem => (oItem.Subcontract == true) &&
                ((oItem.QuoteSetupLaborRate != 0) || (oItem.QuoteSetupBurdenRate != 0) || (oItem.QuoteProductionLaborRate != 0) || (oItem.QuoteProductionBurdenRate != 0))).ToList();
            if (oSubcontractWithQuoteRates.Count > 0)
            {
                iNumberOfRows = 1;
                iNumberOfColumns = 1;
                if (bFirstWorksheet == true)
                {
                    oSLBOMDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "Subcontract With Quote Rates");
                    bFirstWorksheet = false;
                }
                else
                {
                    oSLBOMDocument.AddWorksheet("Subcontract With Quote Rates");
                }
                //set column header
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Resource Group");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 30);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Description");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 30);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Quote Setup Labor");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 30);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Quote Setup Burden");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 30);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Quote Production Labor");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 30);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Quote Production Burden");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 30);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "The Resource Group Is A Subcontract But Has The Quote Rates Set");
                oSLBOMDocument.SetCellStyle(iNumberOfRows++, iNumberOfColumns, oHighlightHeaderStyle);
                oSLBOMDocument.SetColumnWidth(iNumberOfColumns++, 90);

                foreach (HSResourceGroup oResourceGroup in oSubcontractWithQuoteRates)
                {
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 1, oResourceGroup.ResourceGroupId);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 2, oResourceGroup.Description);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 3, oResourceGroup.QuoteSetupLaborRate);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 4, oResourceGroup.QuoteSetupBurdenRate);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 5, oResourceGroup.QuoteProductionLaborRate);
                    oSLBOMDocument.SetCellValue(iNumberOfRows++, 6, oResourceGroup.QuoteProductionBurdenRate);
                    bDataInReport = true;
                }
            }

            // no labor rates
            List<HSResourceGroup> oNoRates = oActiveResourceGroups.Where(oItem => (oItem.Subcontract == false) &&
                ((oItem.SetupLaborRate == 0) || (oItem.SetupBurdenRate == 0) || (oItem.ProductionLaborRate == 0) || (oItem.ProductionBurdenRate == 0))).ToList();
            if (oNoRates.Count > 0)
            {
                iNumberOfRows = 1;
                iNumberOfColumns = 1;
                if (bFirstWorksheet == true)
                {
                    oSLBOMDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "Prod No Rates");
                    bFirstWorksheet = false;
                }
                else
                {
                    oSLBOMDocument.AddWorksheet("Prod No Rates");
                }
                //set column header
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Resource Group");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 30);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Description");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 30);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Setup Labor");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 30);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Setup Burden");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 30);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Production Labor");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 30);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Production Burden");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 30);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "The Resource Group Does Not Have The Labor or Production Rates Set");
                oSLBOMDocument.SetCellStyle(iNumberOfRows++, iNumberOfColumns, oHighlightHeaderStyle);
                oSLBOMDocument.SetColumnWidth(iNumberOfColumns++, 90);

                foreach (HSResourceGroup oResourceGroup in oNoRates)
                {
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 1, oResourceGroup.ResourceGroupId);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 2, oResourceGroup.Description);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 3, oResourceGroup.SetupLaborRate);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 4, oResourceGroup.SetupBurdenRate);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 5, oResourceGroup.ProductionLaborRate);
                    oSLBOMDocument.SetCellValue(iNumberOfRows++, 6, oResourceGroup.ProductionBurdenRate);
                    bDataInReport = true;
                }
            }

            // quote no labor rates
            List<HSResourceGroup> oQuoteWithoutRates = oActiveResourceGroups.Where(oItem => (oItem.Subcontract == false) &&
                ((oItem.QuoteSetupLaborRate == 0) || (oItem.QuoteSetupBurdenRate == 0) || (oItem.QuoteProductionLaborRate == 0) || (oItem.QuoteProductionBurdenRate == 0))).ToList();
            if (oQuoteWithoutRates.Count > 0)
            {
                iNumberOfRows = 1;
                iNumberOfColumns = 1;
                if (bFirstWorksheet == true)
                {
                    oSLBOMDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "Quote No Rates");
                    bFirstWorksheet = false;
                }
                else
                {
                    oSLBOMDocument.AddWorksheet("Quote No Rates");
                }
                //set column header
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Resource Group");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 30);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Description");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 30);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Quote Setup Labor");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 30);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Quote Setup Burden");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 30);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Quote Production Labor");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 30);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Quote Production Burden");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 30);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "The Resource Group Does Not Have The Quote Rates Set");
                oSLBOMDocument.SetCellStyle(iNumberOfRows++, iNumberOfColumns, oHighlightHeaderStyle);
                oSLBOMDocument.SetColumnWidth(iNumberOfColumns++, 90);

                foreach (HSResourceGroup oResourceGroup in oQuoteWithoutRates)
                {
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 1, oResourceGroup.ResourceGroupId);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 2, oResourceGroup.Description);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 3, oResourceGroup.QuoteSetupLaborRate);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 4, oResourceGroup.QuoteSetupBurdenRate);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 5, oResourceGroup.QuoteProductionLaborRate);
                    oSLBOMDocument.SetCellValue(iNumberOfRows++, 6, oResourceGroup.QuoteProductionBurdenRate);
                    bDataInReport = true;
                }
            }

            // use estimates set
            List<HSResourceGroup> oUseEstimatesSet = oActiveResourceGroups.Where(oItem => oItem.UseEstimates == true).ToList();
            if (oUseEstimatesSet.Count > 0)
            {
                iNumberOfRows = 1;
                iNumberOfColumns = 1;
                if (bFirstWorksheet == true)
                {
                    oSLBOMDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "Use Estimates Set");
                    bFirstWorksheet = false;
                }
                else
                {
                    oSLBOMDocument.AddWorksheet("Use Estimates Set");
                }
                //set column header
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Resource Group");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 30);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Description");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 30);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "The Resource Group Is Set To Use Estimates For Rates");
                oSLBOMDocument.SetCellStyle(iNumberOfRows++, iNumberOfColumns, oHighlightHeaderStyle);
                oSLBOMDocument.SetColumnWidth(iNumberOfColumns++, 90);

                foreach (HSResourceGroup oResourceGroup in oUseEstimatesSet)
                {
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 1, oResourceGroup.ResourceGroupId);
                    oSLBOMDocument.SetCellValue(iNumberOfRows++, 2, oResourceGroup.Description);

                    bDataInReport = true;
                }
            }

            // split burden set
            List<HSResourceGroup> oSplitBurdenSet = oActiveResourceGroups.Where(oItem => oItem.SplitBurden == true).ToList();
            if (oSplitBurdenSet.Count > 0)
            {
                iNumberOfRows = 1;
                iNumberOfColumns = 1;
                if (bFirstWorksheet == true)
                {
                    oSLBOMDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "Split Burden Set");
                    bFirstWorksheet = false;
                }
                else
                {
                    oSLBOMDocument.AddWorksheet("Split Burden Set");
                }
                //set column header
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Resource Group");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 30);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Description");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 30);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "The Resource Group Is Set To Split Burden");
                oSLBOMDocument.SetCellStyle(iNumberOfRows++, iNumberOfColumns, oHighlightHeaderStyle);
                oSLBOMDocument.SetColumnWidth(iNumberOfColumns++, 90);

                foreach (HSResourceGroup oResourceGroup in oSplitBurdenSet)
                {
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 1, oResourceGroup.ResourceGroupId);
                    oSLBOMDocument.SetCellValue(iNumberOfRows++, 2, oResourceGroup.Description);

                    bDataInReport = true;
                }
            }

            // burden equals labor not set
            List<HSResourceGroup> oBurdenEqualsLaborNotSet = oActiveResourceGroups.Where(oItem => oItem.BurdenEqualsLabor == false).ToList();
            if (oBurdenEqualsLaborNotSet.Count > 0)
            {
                iNumberOfRows = 1;
                iNumberOfColumns = 1;
                if (bFirstWorksheet == true)
                {
                    oSLBOMDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "Burden Equals Labor Not Set");
                    bFirstWorksheet = false;
                }
                else
                {
                    oSLBOMDocument.AddWorksheet("Burden Equals Labor Not Set");
                }
                //set column header
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Resource Group");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 30);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Description");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 30);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "The Resource Group Does Not Have Burden Hours Set To Equal Labor Hours");
                oSLBOMDocument.SetCellStyle(iNumberOfRows++, iNumberOfColumns, oHighlightHeaderStyle);
                oSLBOMDocument.SetColumnWidth(iNumberOfColumns++, 90);

                foreach (HSResourceGroup oResourceGroup in oBurdenEqualsLaborNotSet)
                {
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 1, oResourceGroup.ResourceGroupId);
                    oSLBOMDocument.SetCellValue(iNumberOfRows++, 2, oResourceGroup.Description);

                    bDataInReport = true;
                }
            }

            // burden type not flat
            List<HSResourceGroup> oProductionBurdenTypeNotFlat = oActiveResourceGroups.Where(oItem => string.Compare(oItem.BurdenType, "F", true) != 0).ToList();
            if (oProductionBurdenTypeNotFlat.Count > 0)
            {
                iNumberOfRows = 1;
                iNumberOfColumns = 1;
                if (bFirstWorksheet == true)
                {
                    oSLBOMDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "Prod Burden Type Not Flat");
                    bFirstWorksheet = false;
                }
                else
                {
                    oSLBOMDocument.AddWorksheet("Prod Burden Type Not Flat");
                }
                //set column header
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Resource Group");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 30);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Description");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 30);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Burden Type");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 30);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "The Resource Group Does Not Have The Burden Type Set To A Flat Rate");
                oSLBOMDocument.SetCellStyle(iNumberOfRows++, iNumberOfColumns, oHighlightHeaderStyle);
                oSLBOMDocument.SetColumnWidth(iNumberOfColumns++, 90);

                foreach (HSResourceGroup oResourceGroup in oProductionBurdenTypeNotFlat)
                {
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 1, oResourceGroup.ResourceGroupId);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 2, oResourceGroup.Description);
                    oSLBOMDocument.SetCellValue(iNumberOfRows++, 3, oResourceGroup.BurdenType);

                    bDataInReport = true;
                }
            }

            // quote burden type not flat
            List<HSResourceGroup> oQuoteBurdenTypeNotFlat = oActiveResourceGroups.Where(oItem => string.Compare(oItem.QuoteBurdenType, "F", true) != 0).ToList();
            if (oQuoteBurdenTypeNotFlat.Count > 0)
            {
                iNumberOfRows = 1;
                iNumberOfColumns = 1;
                if (bFirstWorksheet == true)
                {
                    oSLBOMDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "Quote Burden Type Not Flat");
                    bFirstWorksheet = false;
                }
                else
                {
                    oSLBOMDocument.AddWorksheet("Quote Burden Type Not Flat");
                }
                //set column header
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Resource Group");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 30);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Description");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 30);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Burden Type");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 30);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "The Resource Group Does Not Have THe Quote Burden Set To A Flat Rate");
                oSLBOMDocument.SetCellStyle(iNumberOfRows++, iNumberOfColumns, oHighlightHeaderStyle);
                oSLBOMDocument.SetColumnWidth(iNumberOfColumns++, 90);

                foreach (HSResourceGroup oResourceGroup in oQuoteBurdenTypeNotFlat)
                {
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 1, oResourceGroup.ResourceGroupId);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 2, oResourceGroup.Description);
                    oSLBOMDocument.SetCellValue(iNumberOfRows++, 3, oResourceGroup.BurdenType);

                    bDataInReport = true;
                }
            }

            // split operatiopns set
            List<HSResourceGroup> oSplitOperationsSet = oActiveResourceGroups.Where(oItem => oItem.SplitOperations == true).ToList();
            if (oSplitOperationsSet.Count > 0)
            {
                iNumberOfRows = 1;
                iNumberOfColumns = 1;
                if (bFirstWorksheet == true)
                {
                    oSLBOMDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "Split Operations Set");
                    bFirstWorksheet = false;
                }
                else
                {
                    oSLBOMDocument.AddWorksheet("Split Operations Set");
                }
                //set column header
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Resource Group");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 30);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Description");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 30);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "The Resource Group Has The Operations Set To Split");
                oSLBOMDocument.SetCellStyle(iNumberOfRows++, iNumberOfColumns, oHighlightHeaderStyle);
                oSLBOMDocument.SetColumnWidth(iNumberOfColumns++, 90);

                foreach (HSResourceGroup oResourceGroup in oSplitOperationsSet)
                {
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 1, oResourceGroup.ResourceGroupId);
                    oSLBOMDocument.SetCellValue(iNumberOfRows++, 2, oResourceGroup.Description);

                    bDataInReport = true;
                }
            }

            // inform overload not set
            List<HSResourceGroup> oInformOverloadNotSet = oActiveResourceGroups.Where(oItem => oItem.InformOverload == false).ToList();
            if (oInformOverloadNotSet.Count > 0)
            {
                iNumberOfRows = 1;
                iNumberOfColumns = 1;
                if (bFirstWorksheet == true)
                {
                    oSLBOMDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "Inform Overload Not Set");
                    bFirstWorksheet = false;
                }
                else
                {
                    oSLBOMDocument.AddWorksheet("Inform Overload Not Set");
                }
                //set column header
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Resource Group");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 30);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Description");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 30);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "The Resource Group Does not Have Inform Of Overloads Turned On");
                oSLBOMDocument.SetCellStyle(iNumberOfRows++, iNumberOfColumns, oHighlightHeaderStyle);
                oSLBOMDocument.SetColumnWidth(iNumberOfColumns++, 90);

                foreach (HSResourceGroup oResourceGroup in oInformOverloadNotSet)
                {
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 1, oResourceGroup.ResourceGroupId);
                    oSLBOMDocument.SetCellValue(iNumberOfRows++, 2, oResourceGroup.Description);

                    bDataInReport = true;
                }
            }

            // move hours set
            List<HSResourceGroup> oMoveHoursSet = oActiveResourceGroups.Where(oItem => oItem.MoveHours != 0).ToList();
            if (oMoveHoursSet.Count > 0)
            {
                iNumberOfRows = 1;
                iNumberOfColumns = 1;
                if (bFirstWorksheet == true)
                {
                    oSLBOMDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "Move Hours Set");
                    bFirstWorksheet = false;
                }
                else
                {
                    oSLBOMDocument.AddWorksheet("Move Hours Set");
                }
                //set column header
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Resource Group");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 30);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Description");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 30);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Move Hours");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 30);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "The Resource Group Has The Move Hours Set");
                oSLBOMDocument.SetCellStyle(iNumberOfRows++, iNumberOfColumns, oHighlightHeaderStyle);
                oSLBOMDocument.SetColumnWidth(iNumberOfColumns++, 90);

                foreach (HSResourceGroup oResourceGroup in oMoveHoursSet)
                {
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 1, oResourceGroup.ResourceGroupId);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 2, oResourceGroup.Description);
                    oSLBOMDocument.SetCellValue(iNumberOfRows++, 3, oResourceGroup.MoveHours);

                    bDataInReport = true;
                }
            }

            // queue hours set
            List<HSResourceGroup> oQueueHoursSet = oActiveResourceGroups.Where(oItem => oItem.QueueHours != 0).ToList();
            if (oQueueHoursSet.Count > 0)
            {
                iNumberOfRows = 1;
                iNumberOfColumns = 1;
                if (bFirstWorksheet == true)
                {
                    oSLBOMDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "Queue Hours Set");
                    bFirstWorksheet = false;
                }
                else
                {
                    oSLBOMDocument.AddWorksheet("Queue Hours Set");
                }
                //set column header
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Resource Group");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 30);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Description");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 30);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Queue Hours");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 30);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "The Resource Group Has The Queue Hours Set");
                oSLBOMDocument.SetCellStyle(iNumberOfRows++, iNumberOfColumns, oHighlightHeaderStyle);
                oSLBOMDocument.SetColumnWidth(iNumberOfColumns++, 90);

                foreach (HSResourceGroup oResourceGroup in oQueueHoursSet)
                {
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 1, oResourceGroup.ResourceGroupId);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 2, oResourceGroup.Description);
                    oSLBOMDocument.SetCellValue(iNumberOfRows++, 3, oResourceGroup.QueueHours);

                    bDataInReport = true;
                }
            }

            if (bDataInReport == true)
            {
                oSLBOMDocument.SaveAs(sDestinationFileName);
                // Check to see if we created a file and if so email it
                if (File.Exists(sDestinationFileName) == true)
                {
                    List<string> oAttachments = new List<string>();
                    oAttachments.Add(sDestinationFileName);

                    HSEmailHelper.SendEmail(oToAddresses, "Resource Group Report", "Resource Group Report for " + sDate, oAttachments);
                }
            }

        }

        public void PerformResourceValidation(string sCompany, string sTmpFileDirectory)
        {
            // get the file name
            DateTime dtToday = DateTime.Now;
            string sDate = dtToday.Month.ToString() + "-" + dtToday.Day.ToString() + "-" + dtToday.Year.ToString();
            string sDestinationFileName = sTmpFileDirectory + "\\" + sCompany + "-ResourceReport-" + sDate + ".xlsx";
            int iNumberOfRows = 1;
            int iNumberOfColumns = 1;

            HSUser oServiceAccount = HSUser.GetUserById(HSUser.SPINCRAFT_SERVICE_ACCOUNT_ID);
            List<string> oToAddresses = new List<string>();
            oToAddresses = HSUser.GetEmailsForUsersInGroup(HSUser.REPORT_ON_MANUFACTURED_PART_ISSUES);
            oToAddresses.Add(HSUser.SPINCRAFT_ROOT_USER);

            SLFill oSLFill = new SLFill();
            oSLFill.SetPatternBackgroundColor(SLThemeColorIndexValues.Accent1Color);
            oSLFill.SetPattern(PatternValues.Solid, SLThemeColorIndexValues.Light2Color, SLThemeColorIndexValues.Light2Color);

            SLStyle oBoldStyle = new SLStyle();
            oBoldStyle.SetFontBold(true);
            oBoldStyle.SetTopBorder(BorderStyleValues.Thin, SLThemeColorIndexValues.Dark2Color);
            oBoldStyle.SetBottomBorder(BorderStyleValues.Thin, SLThemeColorIndexValues.Dark2Color);
            oBoldStyle.SetLeftBorder(BorderStyleValues.Thin, SLThemeColorIndexValues.Dark2Color);
            oBoldStyle.SetRightBorder(BorderStyleValues.Thin, SLThemeColorIndexValues.Dark2Color);
            oBoldStyle.SetFontColor(SLThemeColorIndexValues.Dark2Color);

            SLStyle oBoldStyle2 = new SLStyle();
            oBoldStyle2.SetFontBold(true);
            oBoldStyle2.SetTopBorder(BorderStyleValues.Thick, SLThemeColorIndexValues.Dark2Color);
            oBoldStyle2.SetBottomBorder(BorderStyleValues.Thick, SLThemeColorIndexValues.Dark2Color);
            oBoldStyle2.SetLeftBorder(BorderStyleValues.Thick, SLThemeColorIndexValues.Dark2Color);
            oBoldStyle.SetRightBorder(BorderStyleValues.Thick, SLThemeColorIndexValues.Dark2Color);
            oBoldStyle2.SetFontColor(System.Drawing.Color.IndianRed);

            System.Drawing.Color oHighlightYellow = System.Drawing.Color.FromArgb(255, 255, 0);

            SLStyle oHighlightHeaderStyle = new SLStyle();
            oHighlightHeaderStyle.SetFontBold(true);
            oHighlightHeaderStyle.SetFont(FontSchemeValues.Major, 12);
            oHighlightHeaderStyle.SetHorizontalAlignment(HorizontalAlignmentValues.Center);
            oHighlightHeaderStyle.SetPatternFill(PatternValues.Solid, oHighlightYellow, oHighlightYellow);
            oHighlightHeaderStyle.SetFontColor(SLThemeColorIndexValues.Dark2Color);

            oToAddresses = HSUser.GetEmailsForUsersInGroup(HSUser.REPORT_ON_MANUFACTURED_PART_ISSUES);
            oToAddresses.Add(HSUser.SPINCRAFT_ROOT_USER);
            bool bDataInReport = false;
            bool bFirstWorksheet = true;

            // if the file is already there then get rid of it as we are reprocessing for some reason
            if (File.Exists(sDestinationFileName) == true)
            {
                try
                {
                    File.Delete(sDestinationFileName);
                }
                catch (Exception)
                {
                    // we will ingore this if we cannot delete the file
                }
            }
            SLDocument oSLBOMDocument = new SLDocument();

            List<HSResource> oActiveResources = m_oResources.Where(oItem => oItem.Inactive == false).ToList();

            // description not set
            List<HSResource> oNoDescription = oActiveResources.Where(oItem => string.IsNullOrEmpty(oItem.Description)).ToList();
            if (oNoDescription.Count > 0)
            {
                iNumberOfRows = 1;
                iNumberOfColumns = 1;
                if (bFirstWorksheet == true)
                {
                    oSLBOMDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "No Description");
                    bFirstWorksheet = false;
                }
                else
                {
                    oSLBOMDocument.AddWorksheet("No Description");
                }
                //set column header
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Resource");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows++, iNumberOfColumns++, 30);

                foreach (HSResource oResource in oNoDescription)
                {
                    oSLBOMDocument.SetCellValue(iNumberOfRows++, 1, oResource.ResourceId);

                    bDataInReport = true;
                }
            }

            // resource type not set
            List<HSResource> oNoResourceType = oActiveResources.Where(oItem => string.IsNullOrEmpty(oItem.ResourceType)).ToList();
            if (oNoResourceType.Count > 0)
            {
                iNumberOfRows = 1;
                iNumberOfColumns = 1;
                if (bFirstWorksheet == true)
                {
                    oSLBOMDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "No Resource Type");
                    bFirstWorksheet = false;
                }
                else
                {
                    oSLBOMDocument.AddWorksheet("No Resource Type");
                }
                //set column header
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Resource");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 30);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Description");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows++, iNumberOfColumns++, 30);

                foreach (HSResource oResource in oNoResourceType)
                {
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 1, oResource.ResourceId);
                    oSLBOMDocument.SetCellValue(iNumberOfRows++, 2, oResource.Description);

                    bDataInReport = true;
                }
            }

            // not a location
            List<HSResource> oNotLocation = oActiveResources.Where(oItem => oItem.ResourceLocation == false).ToList();
            if (oNotLocation.Count > 0)
            {
                iNumberOfRows = 1;
                iNumberOfColumns = 1;
                if (bFirstWorksheet == true)
                {
                    oSLBOMDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "No Location");
                    bFirstWorksheet = false;
                }
                else
                {
                    oSLBOMDocument.AddWorksheet("No Location");
                }
                //set column header
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Resource");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 30);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Description");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows++, iNumberOfColumns++, 30);

                foreach (HSResource oResource in oNotLocation)
                {
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 1, oResource.ResourceId);
                    oSLBOMDocument.SetCellValue(iNumberOfRows++, 2, oResource.Description);

                    bDataInReport = true;
                }
            }

            // machine not finite
            List<HSResource> oMachineNotFinite = oActiveResources.Where(oItem => (string.Compare(oItem.ResourceType, "M", true) == 0) && (oItem.ResourceFinite == false)).ToList();
            if (oMachineNotFinite.Count > 0)
            {
                iNumberOfRows = 1;
                iNumberOfColumns = 1;
                if (bFirstWorksheet == true)
                {
                    oSLBOMDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "Machine Not Finite");
                    bFirstWorksheet = false;
                }
                else
                {
                    oSLBOMDocument.AddWorksheet("Machine Not Finite");
                }
                //set column header
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Resource");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 30);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Description");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 30);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Resource Type");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows++, iNumberOfColumns++, 30);

                foreach (HSResource oResource in oMachineNotFinite)
                {
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 1, oResource.ResourceId);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 2, oResource.Description);
                    oSLBOMDocument.SetCellValue(iNumberOfRows++, 3, oResource.ResourceType);

                    bDataInReport = true;
                }
            }

            // tool not finite
            List<HSResource> oToolNotFinite = oActiveResources.Where(oItem => (string.Compare(oItem.ResourceType, "T", true) == 0) && (oItem.ResourceFinite == false)).ToList();
            if (oToolNotFinite.Count > 0)
            {
                iNumberOfRows = 1;
                iNumberOfColumns = 1;
                if (bFirstWorksheet == true)
                {
                    oSLBOMDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "Tool Not Finite");
                    bFirstWorksheet = false;
                }
                else
                {
                    oSLBOMDocument.AddWorksheet("Tool Not Finite");
                }
                //set column header
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Resource");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 30);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Description");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 30);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Resource Type");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows++, iNumberOfColumns++, 30);

                foreach (HSResource oResource in oToolNotFinite)
                {
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 1, oResource.ResourceId);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 2, oResource.Description);
                    oSLBOMDocument.SetCellValue(iNumberOfRows++, 3, oResource.ResourceType);

                    bDataInReport = true;
                }
            }

            // no operation code
            //List<HSResource> oNoOperationCode = oActiveResources.Where(oItem => string.IsNullOrEmpty(oItem.OpCode) == true).ToList();
            //if (oNoOperationCode.Count > 0)
            //{
            //    iNumberOfRows = 1;
            //    iNumberOfColumns = 1;
            //    if (bFirstWorksheet == true)
            //    {
            //        oSLBOMDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "No Op Code");
            //        bFirstWorksheet = false;
            //    }
            //    else
            //    {
            //        oSLBOMDocument.AddWorksheet("No Op Code");
            //    }
            //    //set column header
            //    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Resource");
            //    oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 30);
            //    oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Description");
            //    oSLBOMDocument.SetColumnWidth(iNumberOfRows++, iNumberOfColumns++, 30);

            //    foreach (HSResource oResource in oNoOperationCode)
            //    {
            //        oSLBOMDocument.SetCellValue(iNumberOfRows, 1, oResource.ResourceId);
            //        oSLBOMDocument.SetCellValue(iNumberOfRows++, 2, oResource.Description);

            //        bDataInReport = true;
            //    }
            //}

            // labor not from group
            List<HSResource> oLaborNotFromGroup = oActiveResources.Where(oItem => oItem.GetDefaultLaborFromGroup == false).ToList();
            if (oLaborNotFromGroup.Count > 0)
            {
                iNumberOfRows = 1;
                iNumberOfColumns = 1;
                if (bFirstWorksheet == true)
                {
                    oSLBOMDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "Labor Not From Group");
                    bFirstWorksheet = false;
                }
                else
                {
                    oSLBOMDocument.AddWorksheet("Labor Not From Group");
                }
                //set column header
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Resource");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 30);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Description");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows++, iNumberOfColumns++, 30);

                foreach (HSResource oResource in oLaborNotFromGroup)
                {
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 1, oResource.ResourceId);
                    oSLBOMDocument.SetCellValue(iNumberOfRows++, 2, oResource.Description);

                    bDataInReport = true;
                }
            }

            // burden not from group
            List<HSResource> oBurdenNotFromGroup = oActiveResources.Where(oItem => oItem.GetDefaultBurdenFromGroup == false).ToList();
            if (oBurdenNotFromGroup.Count > 0)
            {
                iNumberOfRows = 1;
                iNumberOfColumns = 1;
                if (bFirstWorksheet == true)
                {
                    oSLBOMDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "Burden Not From Group");
                    bFirstWorksheet = false;
                }
                else
                {
                    oSLBOMDocument.AddWorksheet("Burden Not From Group");
                }
                //set column header
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Resource");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 30);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Description");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows++, iNumberOfColumns++, 30);

                foreach (HSResource oResource in oBurdenNotFromGroup)
                {
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 1, oResource.ResourceId);
                    oSLBOMDocument.SetCellValue(iNumberOfRows++, 2, oResource.Description);

                    bDataInReport = true;
                }
            }

            // burden type not flat
            List<HSResource> oBurdenTypeNotFlat = oActiveResources.Where(oItem => string.Compare(oItem.BurdenType, "F", true) != 0).ToList();
            if (oBurdenTypeNotFlat.Count > 0)
            {
                iNumberOfRows = 1;
                iNumberOfColumns = 1;
                if (bFirstWorksheet == true)
                {
                    oSLBOMDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "Burden Type Not Flat");
                    bFirstWorksheet = false;
                }
                else
                {
                    oSLBOMDocument.AddWorksheet("Burden Type Not Flat");
                }
                //set column header
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Resource");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 30);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Description");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 30);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Burden Type");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows++, iNumberOfColumns++, 30);

                foreach (HSResource oResource in oBurdenTypeNotFlat)
                {
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 1, oResource.ResourceId);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 2, oResource.Description);
                    oSLBOMDocument.SetCellValue(iNumberOfRows++, 3, oResource.BurdenType);

                    bDataInReport = true;
                }
            }

            // burden type not flat
            List<HSResource> oQuoteBurdenTypeNotFlat = oActiveResources.Where(oItem => string.Compare(oItem.QuoteBurdenType, "F", true) != 0).ToList();
            if (oQuoteBurdenTypeNotFlat.Count > 0)
            {
                iNumberOfRows = 1;
                iNumberOfColumns = 1;
                if (bFirstWorksheet == true)
                {
                    oSLBOMDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "Quote Burden Type Not Flat");
                    bFirstWorksheet = false;
                }
                else
                {
                    oSLBOMDocument.AddWorksheet("Quote Burden Type Not Flat");
                }
                //set column header
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Resource");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 30);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Description");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 30);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Burden Type");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows++, iNumberOfColumns++, 30);

                foreach (HSResource oResource in oQuoteBurdenTypeNotFlat)
                {
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 1, oResource.ResourceId);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 2, oResource.Description);
                    oSLBOMDocument.SetCellValue(iNumberOfRows++, 3, oResource.QuoteBurdenType);

                    bDataInReport = true;
                }
            }

            // rates set on resource
            List<HSResource> oRatesSet = oActiveResources.Where(oItem => (oItem.SetupLaborRate != 0) || (oItem.SetupBurdenRate != 0) || (oItem.ProductionLaborRate != 0) || (oItem.ProductionBurdenRate != 0)).ToList();
            if (oRatesSet.Count > 0)
            {
                iNumberOfRows = 1;
                iNumberOfColumns = 1;
                if (bFirstWorksheet == true)
                {
                    oSLBOMDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "Rates Set");
                    bFirstWorksheet = false;
                }
                else
                {
                    oSLBOMDocument.AddWorksheet("Rates Set");
                }
                //set column header
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Resource");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 30);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Description");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 30);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Setup Labor");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 30);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Setup Burden");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 30);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Production Labor");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 30);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Production Burden");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows++, iNumberOfColumns++, 30);

                foreach (HSResource oResource in oRatesSet)
                {
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 1, oResource.ResourceId);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 2, oResource.Description);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 3, oResource.SetupLaborRate);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 4, oResource.SetupBurdenRate);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 5, oResource.ProductionLaborRate);
                    oSLBOMDocument.SetCellValue(iNumberOfRows++, 6, oResource.ProductionBurdenRate);
                    bDataInReport = true;
                }
            }

            // quote rates set on resource
            List<HSResource> oQuoteRatesSet = oActiveResources.Where(oItem => (oItem.QuoteSetupLaborRate != 0) || (oItem.QuoteSetupBurdenRate != 0) || (oItem.QuoteProductionLaborRate != 0) || (oItem.QuoteProductionBurdenRate != 0)).ToList();
            if (oQuoteRatesSet.Count > 0)
            {
                iNumberOfRows = 1;
                iNumberOfColumns = 1;
                if (bFirstWorksheet == true)
                {
                    oSLBOMDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "Quote Rates Set");
                    bFirstWorksheet = false;
                }
                else
                {
                    oSLBOMDocument.AddWorksheet("Quote Rates Set");
                }
                //set column header
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Resource");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 30);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Description");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 30);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Quote Setup Labor");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 30);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Quote Setup Burden");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 30);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Quote Production Labor");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 30);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Quote Production Burden");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows++, iNumberOfColumns++, 30);

                foreach (HSResource oResource in oQuoteRatesSet)
                {
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 1, oResource.ResourceId);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 2, oResource.Description);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 3, oResource.QuoteSetupLaborRate);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 4, oResource.QuoteSetupBurdenRate);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 5, oResource.QuoteProductionLaborRate);
                    oSLBOMDocument.SetCellValue(iNumberOfRows++, 6, oResource.QuoteProductionBurdenRate);
                    bDataInReport = true;
                }
            }

            // quote rates set on resource
            List<HSResource> oNotUsingResourceGroupSchedulingValues = oActiveResources.Where(oItem => oItem.GetDefaultMoveQueueFromGroup == false).ToList();
            if (oNotUsingResourceGroupSchedulingValues.Count > 0)
            {
                iNumberOfRows = 1;
                iNumberOfColumns = 1;
                if (bFirstWorksheet == true)
                {
                    oSLBOMDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "Not Using Resource Group Scheduling");
                    bFirstWorksheet = false;
                }
                else
                {
                    oSLBOMDocument.AddWorksheet("Not Using Resource Group Scheduling");
                }
                //set column header
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Resource");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 30);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Description");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows++, iNumberOfColumns++, 30);

                foreach (HSResource oResource in oNotUsingResourceGroupSchedulingValues)
                {
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 1, oResource.ResourceId);
                    oSLBOMDocument.SetCellValue(iNumberOfRows++, 2, oResource.Description);
                    bDataInReport = true;
                }
            }


            // split operations set
            List<HSResource> oSplitOperationsSet = oActiveResources.Where(oItem => oItem.SplitOperations == true).ToList();
            if (oSplitOperationsSet.Count > 0)
            {
                iNumberOfRows = 1;
                iNumberOfColumns = 1;
                if (bFirstWorksheet == true)
                {
                    oSLBOMDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "Split Operations Set");
                    bFirstWorksheet = false;
                }
                else
                {
                    oSLBOMDocument.AddWorksheet("Split Operations Set");
                }
                //set column header
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Resource");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 30);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Description");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows++, iNumberOfColumns++, 30);

                foreach (HSResource oResource in oSplitOperationsSet)
                {
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 1, oResource.ResourceId);
                    oSLBOMDocument.SetCellValue(iNumberOfRows++, 2, oResource.Description);
                    bDataInReport = true;
                }
            }

            // queue hours set
            List<HSResource> oQueueHoursSet = oActiveResources.Where(oItem => oItem.QueueHours != 0).ToList();
            if (oQueueHoursSet.Count > 0)
            {
                iNumberOfRows = 1;
                iNumberOfColumns = 1;
                if (bFirstWorksheet == true)
                {
                    oSLBOMDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "Queue Hours Set");
                    bFirstWorksheet = false;
                }
                else
                {
                    oSLBOMDocument.AddWorksheet("Queue Hours Set");
                }
                //set column header
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Resource");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 30);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Description");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 30);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Queue Hours");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows++, iNumberOfColumns++, 30);

                foreach (HSResource oResource in oQueueHoursSet)
                {
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 1, oResource.ResourceId);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 2, oResource.Description);
                    oSLBOMDocument.SetCellValue(iNumberOfRows++, 3, oResource.QueueHours);
                    bDataInReport = true;
                }
            }

            // move hours set
            List<HSResource> oMoveHoursSet = oActiveResources.Where(oItem => oItem.MoveHours != 0).ToList();
            if (oMoveHoursSet.Count > 0)
            {
                iNumberOfRows = 1;
                iNumberOfColumns = 1;
                if (bFirstWorksheet == true)
                {
                    oSLBOMDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "Move Hours Set");
                    bFirstWorksheet = false;
                }
                else
                {
                    oSLBOMDocument.AddWorksheet("Move Hours Set");
                }
                //set column header
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Resource");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 30);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Description");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 30);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Move Hours");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows++, iNumberOfColumns++, 30);

                foreach (HSResource oResource in oMoveHoursSet)
                {
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 1, oResource.ResourceId);
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 2, oResource.Description);
                    oSLBOMDocument.SetCellValue(iNumberOfRows++, 3, oResource.MoveHours);
                    bDataInReport = true;
                }
            }

            // calendar not set
            List<HSResource> oCalendarNotSet = oActiveResources.Where(oItem => string.IsNullOrEmpty(oItem.CalendarId) == true).ToList();
            if (oCalendarNotSet.Count > 0)
            {
                iNumberOfRows = 1;
                iNumberOfColumns = 1;
                if (bFirstWorksheet == true)
                {
                    oSLBOMDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "No Calendar");
                    bFirstWorksheet = false;
                }
                else
                {
                    oSLBOMDocument.AddWorksheet("No Calendar");
                }
                //set column header
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Resource");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 30);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Description");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows++, iNumberOfColumns++, 30);

                foreach (HSResource oResource in oCalendarNotSet)
                {
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 1, oResource.ResourceId);
                    oSLBOMDocument.SetCellValue(iNumberOfRows++, 2, oResource.Description);
                    bDataInReport = true;
                }
            }

            // inform overload not set
            List<HSResource> oInformOverloadNotSet = oActiveResources.Where(oItem => oItem.InformOverload != true).ToList();
            if (oInformOverloadNotSet.Count > 0)
            {
                iNumberOfRows = 1;
                iNumberOfColumns = 1;
                if (bFirstWorksheet == true)
                {
                    oSLBOMDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "Inform Overload Not Set");
                    bFirstWorksheet = false;
                }
                else
                {
                    oSLBOMDocument.AddWorksheet("Inform Overload Not Set");
                }
                //set column header
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Resource");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows, iNumberOfColumns++, 30);
                oSLBOMDocument.SetCellValue(iNumberOfRows, iNumberOfColumns, "Description");
                oSLBOMDocument.SetColumnWidth(iNumberOfRows++, iNumberOfColumns++, 30);

                foreach (HSResource oResource in oInformOverloadNotSet)
                {
                    oSLBOMDocument.SetCellValue(iNumberOfRows, 1, oResource.ResourceId);
                    oSLBOMDocument.SetCellValue(iNumberOfRows++, 2, oResource.Description);
                    bDataInReport = true;
                }
            }

            if (bDataInReport == true)
            {
                oSLBOMDocument.SaveAs(sDestinationFileName);
                // Check to see if we created a file and if so email it
                if (File.Exists(sDestinationFileName) == true)
                {
                    List<string> oAttachments = new List<string>();
                    oAttachments.Add(sDestinationFileName);

                    HSEmailHelper.SendEmail(oToAddresses, "Resource Report", "Resource Report for " + sDate, oAttachments);
                }
            }
        }
        #endregion

        #region Data Members
        private string m_sCompany;

        // list of exceptions for checks
        private List<string> m_oPartClassesToIgnore = new List<string>();
        private List<string> m_sMfgPartsToInclude = new List<string>();
        private List<string> m_oMfgPartsThatShouldBeStocked = new List<string>();

        private Dictionary<string, List<PartMaterial>> m_oMtlParts = new Dictionary<string, List<PartMaterial>>();
        private List<PartOperation> m_oOperationParts = new List<PartOperation>();
        private List<MfgPart> m_oMfgParts = new List<MfgPart>();
        private Dictionary<string, List<MfgPart>> m_oParentMfgParts = new Dictionary<string, List<MfgPart>>();
        private HSValidateParts m_oValidateParts = new HSValidateParts();

        private List<HSOperation> m_oOperations = new List<HSOperation>();
        private List<HSResourceGroup> m_oResourceGroups = new List<HSResourceGroup>();
        private List<HSResource> m_oResources = new List<HSResource>();

        private List<MfgPart> m_oMfgPartsWithoutMaterials = new List<MfgPart>();
        private List<MfgPart> m_oMfgPartsWithInactiveMaterials = new List<MfgPart>();
        private List<MfgPart> m_oMfgPartsWithMaterialsOnHold = new List<MfgPart>();
        private List<MfgPart> m_oMfgPartsWithMaterialsOnRunOut = new List<MfgPart>();
        private List<MfgPart> m_oMfgPartsWithDuplicatedMaterials = new List<MfgPart>();
        private List<MfgPart> m_oMfgPartsWithMaterialsNotTiedToOperations = new List<MfgPart>();
        private List<MfgPart> m_oMfgPartsWithoutOperations = new List<MfgPart>();
        private List<MfgPart> m_oMfgPartsWithBadRevision = new List<MfgPart>();
        private List<MfgPart> m_oMfgPartsWithNoApprovedRevision = new List<MfgPart>();
        private List<MfgPart> m_oMfgPartsWithMultipleApprovedRevisions = new List<MfgPart>();
        private List<MfgPart> m_oMfgPartsNotUsingPartRevs = new List<MfgPart>();
        private List<MfgPart> m_oMfgPartsNotQuantityBearing = new List<MfgPart>();
        private List<MfgPart> m_oMfgPartsWithIncosistentQuantityBearing = new List<MfgPart>();
        private List<MfgPart> m_oMfgPartsWithIncorrectUOM = new List<MfgPart>();
        private List<MfgPart> m_oMfgPartsWithInconsistentPhantomBOM = new List<MfgPart>();
        private List<MfgPart> m_oMfgPartsWithInconsistentNonStock = new List<MfgPart>();
        private List<MfgPart> m_oMfgPartsWithInconsistentTypeCode = new List<MfgPart>();
        private List<MfgPart> m_oMfgPartsWithMaterialsWithZeroQty = new List<MfgPart>();
        private List<MfgPart> m_oMfgPartsWithMaterialsWithZeroCost = new List<MfgPart>();
        private List<MfgPart> m_oMfgPartsWithMRPTurnedOff = new List<MfgPart>();
        private List<MfgPart> m_oMfgPartsWithBuyToOrderSet = new List<MfgPart>();
        private List<MfgPart> m_oMfgPartsWithInconsistentBuyToOrder = new List<MfgPart>();
        private List<MfgPart> m_oMfgPartsWithDropShipSet = new List<MfgPart>();
        private List<MfgPart> m_oMfgPartsWithInconsistentDropShip = new List<MfgPart>();
        private List<MfgPart> m_oMfgPartsWithoutPrimaryWarehouse = new List<MfgPart>();
        private List<MfgPart> m_oMfgPartsWithoutGroupCode = new List<MfgPart>();
        private List<MfgPart> m_oMfgPartsWithoutClassCode = new List<MfgPart>();
        private List<MfgPart> m_oMfgPartsWithIncorrectCostMethod = new List<MfgPart>();
        private List<MfgPart> m_oMfgPartsNeedingCostRoll = new List<MfgPart>();
        private List<MfgPart> m_oMfgPartsWithInconsistentCostMethod = new List<MfgPart>();
        private List<MfgPart> m_oMfgPartsSetToStocking = new List<MfgPart>();
        private List<MfgPart> m_oMfgPartsStockedWithoutUnitPrice = new List<MfgPart>();
        private List<MfgPart> m_oMfgPartsStockedWithoutMinOrSafetyLimits = new List<MfgPart>();
        private List<MfgPart> m_oMfgPartsWithoutPortfolioCode = new List<MfgPart>();


        private List<MfgPart> m_oMfgPartWithOperationsNotSetToEach = new List<MfgPart>();
        private List<MfgPart> m_oMfgPartWithOperationsBadLaborEntry = new List<MfgPart>();
        private List<MfgPart> m_oMfgPartWithOperationsBadStandardFormat = new List<MfgPart>();
        private List<MfgPart> m_oMfgPartWithOperationsBadProductionStandard = new List<MfgPart>();
        private List<MfgPart> m_oMfgPartWithOperationsBadOperationPerPartValue = new List<MfgPart>();
        private List<MfgPart> m_oMfgPartWithNonZeroOperationsPerPartValue = new List<MfgPart>();
        private List<MfgPart> m_oMfgPartWithOperationsWithAdditionalSetupQty = new List<MfgPart>();
        private List<MfgPart> m_oMfgPartWithOperationsWithAdditionalSetupHours = new List<MfgPart>();
        private List<MfgPart> m_oMfgPartWithOperationsWithResourceSpecified = new List<MfgPart>();
        private List<MfgPart> m_oMfgPartSubcontractOperationMissingQtyPer = new List<MfgPart>();
        private List<MfgPart> m_oMfgPartSubcontractOperationMissingDaysOut = new List<MfgPart>();
        private List<MfgPart> m_oMfgPartSubcontractOperationMissingUnitCost = new List<MfgPart>();
        private List<MfgPart> m_oMfgPartMaterialsWithFixedQty = new List<MfgPart>();
        private List<MfgPart> m_oMfgPartMaterialsNeedViewAsAssemblyTurnedOn = new List<MfgPart>();
        private List<MfgPart> m_oMfgPartMaterialsNeedViewAsAssemblyTurnedOff = new List<MfgPart>();
        private List<MfgPart> m_oMfgPartMaterialsHasPlanAsAssemblyTurnedOn = new List<MfgPart>();

        #endregion

    }
}
