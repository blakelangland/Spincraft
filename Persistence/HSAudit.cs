using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using System.IO;

using Ice.Core;
using Erp.BO;
using Erp.Proxy.BO;
using Ice.Lib.Framework;
using Ice.Adapters;
using Ice.BO;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Drawing.Spreadsheet;
using SpreadsheetLight;
using HSPersistence;
using SpreadsheetLight.Charts;
using Ice.Lib.Searches;
using System.ServiceModel.Configuration;
using System.Net;

namespace HorizonScientific
{
	public class HSAudit
	{
		static public bool Initialize(Session oSession)
		{
			bool bSuccess = true;

			// now we will clear out all lists
			g_oMenuSecurityItems.Clear();
			g_oAllUsers.Clear();
			g_oActiveUserGroups.Clear();

            // then we need to query each sales order to determine its status
            Ice.Proxy.BO.DynamicQueryImpl oDynamicQuery = WCFServiceSupport.CreateImpl<Ice.Proxy.BO.DynamicQueryImpl>(oSession, Ice.Proxy.BO.DynamicQueryImpl.UriPath);
            QueryExecutionDataSet oQueryExecutionDataSet = oDynamicQuery.GetQueryExecutionParametersByID(BAQConstants.QUERY_ACTIVE_GROUPS);
			oQueryExecutionDataSet.Clear();
			DataSet oDataSet = oDynamicQuery.ExecuteByID(BAQConstants.QUERY_ACTIVE_GROUPS, oQueryExecutionDataSet);
			foreach (DataRow oRow in oDataSet.Tables[0].Rows)
			{
				g_oActiveUserGroups.Add(new ActiveUserGroup(oRow));
			}
			g_oActiveUserGroups = g_oActiveUserGroups.OrderBy(x => x.Description).ToList();

			oQueryExecutionDataSet = oDynamicQuery.GetQueryExecutionParametersByID(BAQConstants.QUERY_ALL_USERS);
			oQueryExecutionDataSet.Clear();
			oDataSet = oDynamicQuery.ExecuteByID(BAQConstants.QUERY_ALL_USERS, oQueryExecutionDataSet);
			foreach (DataRow oRow in oDataSet.Tables[0].Rows)
			{
				g_oAllUsers.Add(new UserSecurityInfo(oRow, g_oActiveUserGroups));
			}
			// sort by login Id
			g_oAllUsers = g_oAllUsers.OrderBy(x => x.UserID).ToList();

			oQueryExecutionDataSet = oDynamicQuery.GetQueryExecutionParametersByID(BAQConstants.QUERY_AUDIT_MENU_SECURITY);
			oQueryExecutionDataSet.Clear();
			oDataSet = oDynamicQuery.ExecuteByID(BAQConstants.QUERY_AUDIT_MENU_SECURITY, oQueryExecutionDataSet);
			foreach (DataRow oRow in oDataSet.Tables[0].Rows)
			{
				MenuSecurity oMenu = new MenuSecurity(oRow, g_oAllUsers, g_oActiveUserGroups);
				g_oMenuSecurityItems.Add(oMenu);
			}

			return bSuccess;
		}

		static public void GenerateReport(string sTmpFileDirectory, HSUser oRequestingUser)
		{
			DateTime dtToday = DateTime.Now;
			string sDestinationFileName = sTmpFileDirectory + "Audit Report -" + dtToday.Month.ToString() + "-" + dtToday.Day.ToString() + "-" + dtToday.Year.ToString() + ".xlsx";

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

			HSUser oServiceAccount = HSUser.GetUserById(HSUser.SPINCRAFT_SERVICE_ACCOUNT_ID);
			List<string> oToAddresses = new List<string>();
			oToAddresses.Add(HSUser.SPINCRAFT_ROOT_USER);
			if (oRequestingUser != null)
			{
				// this report was requested by someone directly from Epicor
				oToAddresses.Add(oRequestingUser.Email);
			}
			bool bFirstWorksheet = true;

			SLDocument oSLAuditDocument = new SLDocument();
			// set up the style of cells
			g_oGood = oSLAuditDocument.CreateStyle();
			g_oGood.ApplyNamedCellStyle(SLNamedCellStyleValues.Good);
			g_oNeutral = oSLAuditDocument.CreateStyle();
			g_oNeutral.ApplyNamedCellStyle(SLNamedCellStyleValues.Neutral);
			g_oBad = oSLAuditDocument.CreateStyle();
			g_oBad.ApplyNamedCellStyle(SLNamedCellStyleValues.Bad);
			g_oBoldText = oSLAuditDocument.CreateStyle();
			g_oBoldText.SetFontBold(true);

			if (g_oAllUsers.Count != 0)
			{
				#region User Info
				if (bFirstWorksheet == true)
				{
					oSLAuditDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "User Info");
					bFirstWorksheet = false;
				}
				else
				{
					oSLAuditDocument.AddWorksheet("User Info");
				}

				//set up row headers
				int iNumOfRows = 1;
				int iNumOfColmns = 0;

				// BASIC USER INFO
				iNumOfColmns++;
				oSLAuditDocument.SetCellValue(iNumOfRows, iNumOfColmns, "Company ID");
				oSLAuditDocument.SetCellStyle(iNumOfRows, iNumOfColmns, g_oBoldText);
				oSLAuditDocument.SetColumnWidth(iNumOfColmns, 12);
				iNumOfColmns++;
				oSLAuditDocument.SetCellValue(iNumOfRows, iNumOfColmns, "Account Disabled");
				oSLAuditDocument.SetCellStyle(iNumOfRows, iNumOfColmns, g_oBoldText);
				oSLAuditDocument.SetColumnWidth(iNumOfColmns, 20);
				iNumOfColmns++;
				oSLAuditDocument.SetCellValue(iNumOfRows, iNumOfColmns, "Security Manager");
				oSLAuditDocument.SetCellStyle(iNumOfRows, iNumOfColmns, g_oBoldText);
				oSLAuditDocument.SetColumnWidth(iNumOfColmns, 20);
				iNumOfColmns++;
				oSLAuditDocument.SetCellValue(iNumOfRows, iNumOfColmns, "User Id");
				oSLAuditDocument.SetCellStyle(iNumOfRows, iNumOfColmns, g_oBoldText);
				oSLAuditDocument.SetColumnWidth(iNumOfColmns, 15);
				iNumOfColmns++;
				oSLAuditDocument.SetCellValue(iNumOfRows, iNumOfColmns, "OS User Id");
				oSLAuditDocument.SetCellStyle(iNumOfRows, iNumOfColmns, g_oBoldText);
				oSLAuditDocument.SetColumnWidth(iNumOfColmns, 15);
				iNumOfColmns++;
				oSLAuditDocument.SetCellValue(iNumOfRows, iNumOfColmns, "Name");
				oSLAuditDocument.SetCellStyle(iNumOfRows, iNumOfColmns, g_oBoldText);
				oSLAuditDocument.SetColumnWidth(iNumOfColmns, 15);
				iNumOfColmns++;
				oSLAuditDocument.SetCellValue(iNumOfRows, iNumOfColmns, "Create Date");
				oSLAuditDocument.SetCellStyle(iNumOfRows, iNumOfColmns, g_oBoldText);
				oSLAuditDocument.SetColumnWidth(iNumOfColmns, 15);
				iNumOfColmns++;
				oSLAuditDocument.SetCellValue(iNumOfRows, iNumOfColmns, "Last Login");
				oSLAuditDocument.SetCellStyle(iNumOfRows, iNumOfColmns, g_oBoldText);
				oSLAuditDocument.SetColumnWidth(iNumOfColmns, 15);
				iNumOfColmns++;
				oSLAuditDocument.SetCellValue(iNumOfRows, iNumOfColmns, "Password Last Changed");
				oSLAuditDocument.SetCellStyle(iNumOfRows, iNumOfColmns, g_oBoldText);
				oSLAuditDocument.SetColumnWidth(iNumOfColmns, 25);
				iNumOfColmns++;
				oSLAuditDocument.SetCellValue(iNumOfRows, iNumOfColmns, "Password Expires");
				oSLAuditDocument.SetCellStyle(iNumOfRows, iNumOfColmns, g_oBoldText);
				oSLAuditDocument.SetColumnWidth(iNumOfColmns, 20);
				iNumOfColmns++;
				oSLAuditDocument.SetCellValue(iNumOfRows, iNumOfColmns, "Password Expires Days");
				oSLAuditDocument.SetCellStyle(iNumOfRows, iNumOfColmns, g_oBoldText);
				oSLAuditDocument.SetColumnWidth(iNumOfColmns, 25);

				iNumOfRows = 2;
				iNumOfColmns = 1;
				foreach (UserSecurityInfo oUser in g_oAllUsers)
				{
					// basic user info
					oSLAuditDocument.SetCellValue(iNumOfRows, iNumOfColmns, oUser.CompanyID);
					iNumOfColmns++;

					oSLAuditDocument.SetCellValue(iNumOfRows, iNumOfColmns, oUser.Disabled);
					if (oUser.Disabled == true)
					{
						oSLAuditDocument.SetCellStyle(iNumOfRows, iNumOfColmns, g_oBad);
					}
					else
					{
						oSLAuditDocument.SetCellStyle(iNumOfRows, iNumOfColmns, g_oGood);
					}
					iNumOfColmns++;

					oSLAuditDocument.SetCellValue(iNumOfRows, iNumOfColmns, oUser.SecurityMgr);
					if (oUser.SecurityMgr == true)
					{
						oSLAuditDocument.SetCellStyle(iNumOfRows, iNumOfColmns, g_oGood);
					}
					else
					{
						oSLAuditDocument.SetCellStyle(iNumOfRows, iNumOfColmns, g_oBad);
					}
					iNumOfColmns++;

					oSLAuditDocument.SetCellValue(iNumOfRows, iNumOfColmns++, oUser.UserID);
					oSLAuditDocument.SetCellValue(iNumOfRows, iNumOfColmns++, oUser.OSUserID);
					oSLAuditDocument.SetCellValue(iNumOfRows, iNumOfColmns++, oUser.Name);
					oSLAuditDocument.SetCellValue(iNumOfRows, iNumOfColmns++, oUser.CreateDate.ToShortDateString());
					oSLAuditDocument.SetCellValue(iNumOfRows, iNumOfColmns++, oUser.LastLogOnAttempt.ToShortDateString());
					oSLAuditDocument.SetCellValue(iNumOfRows, iNumOfColmns++, oUser.PwdLastChanged.ToShortDateString());
					oSLAuditDocument.SetCellValue(iNumOfRows, iNumOfColmns++, oUser.PwdExpires.ToShortDateString());
					oSLAuditDocument.SetCellValue(iNumOfRows, iNumOfColmns++, oUser.PwdExpiresDays.ToString());
					// special permissions


					// move to next row for user and reset column number
					iNumOfRows++;
					iNumOfColmns = 1;
				}
				#endregion

				#region Special Roles
				if (bFirstWorksheet == true)
				{
					oSLAuditDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "Special Roles");
					bFirstWorksheet = false;
				}
				else
				{
					oSLAuditDocument.AddWorksheet("Special Roles");
				}

				//set up row headers
				iNumOfRows = 0;
				iNumOfColmns = 1;

				// BASIC USER INFO
				iNumOfRows++;
				oSLAuditDocument.SetCellValue(iNumOfRows, iNumOfColmns, "Company ID");
				oSLAuditDocument.SetCellStyle(iNumOfRows, iNumOfColmns, g_oBoldText);
				oSLAuditDocument.SetColumnWidth(iNumOfColmns, 12);
				iNumOfColmns++;
				oSLAuditDocument.SetCellValue(iNumOfRows, iNumOfColmns, "Account Disabled");
				oSLAuditDocument.SetCellStyle(iNumOfRows, iNumOfColmns, g_oBoldText);
				oSLAuditDocument.SetColumnWidth(iNumOfColmns, 20);
				iNumOfColmns++;
				oSLAuditDocument.SetCellValue(iNumOfRows, iNumOfColmns, "Security Manager");
				oSLAuditDocument.SetCellStyle(iNumOfRows, iNumOfColmns, g_oBoldText);
				oSLAuditDocument.SetColumnWidth(iNumOfColmns, 20);
				iNumOfColmns++;
				oSLAuditDocument.SetCellValue(iNumOfRows, iNumOfColmns, "User Id");
				oSLAuditDocument.SetCellStyle(iNumOfRows, iNumOfColmns, g_oBoldText);
				oSLAuditDocument.SetColumnWidth(iNumOfColmns, 15);


				// SPECIAL PERMISSIONS
				iNumOfColmns++;
				oSLAuditDocument.SetCellValue(iNumOfRows, iNumOfColmns, "Can Impersonate");
				oSLAuditDocument.SetCellStyle(iNumOfRows, iNumOfColmns, g_oBoldText);
				oSLAuditDocument.SetColumnWidth(iNumOfColmns, 20);
				iNumOfColmns++;
				oSLAuditDocument.SetCellValue(iNumOfRows, iNumOfColmns, "Can Customize");
				oSLAuditDocument.SetCellStyle(iNumOfRows, iNumOfColmns, g_oBoldText);
				oSLAuditDocument.SetColumnWidth(iNumOfColmns, 20);
				iNumOfColmns++;
				oSLAuditDocument.SetCellValue(iNumOfRows, iNumOfColmns, "Create Solutions");
				oSLAuditDocument.SetCellStyle(iNumOfRows, iNumOfColmns, g_oBoldText);
				oSLAuditDocument.SetColumnWidth(iNumOfColmns, 20);
				iNumOfColmns++;
				oSLAuditDocument.SetCellValue(iNumOfRows, iNumOfColmns, "Install Solutions");
				oSLAuditDocument.SetCellStyle(iNumOfRows, iNumOfColmns, g_oBoldText);
				oSLAuditDocument.SetColumnWidth(iNumOfColmns, 20);
				iNumOfColmns++;
				oSLAuditDocument.SetCellValue(iNumOfRows, iNumOfColmns, "Dashboard Developer");
				oSLAuditDocument.SetCellStyle(iNumOfRows, iNumOfColmns, g_oBoldText);
				oSLAuditDocument.SetColumnWidth(iNumOfColmns, 20);
				iNumOfColmns++;
				oSLAuditDocument.SetCellValue(iNumOfRows, iNumOfColmns, "BPM Developer");
				oSLAuditDocument.SetCellStyle(iNumOfRows, iNumOfColmns, g_oBoldText);
				oSLAuditDocument.SetColumnWidth(iNumOfColmns, 20);
				iNumOfColmns++;
				oSLAuditDocument.SetCellValue(iNumOfRows, iNumOfColmns, "Advanced BAQ Rights");
				oSLAuditDocument.SetCellStyle(iNumOfRows, iNumOfColmns, g_oBoldText);
				oSLAuditDocument.SetColumnWidth(iNumOfColmns, 25);
				iNumOfColmns++;
				oSLAuditDocument.SetCellValue(iNumOfRows, iNumOfColmns, "SSRS Designer");
				oSLAuditDocument.SetCellStyle(iNumOfRows, iNumOfColmns, g_oBoldText);
				oSLAuditDocument.SetColumnWidth(iNumOfColmns, 20);
				iNumOfColmns++;
				oSLAuditDocument.SetCellValue(iNumOfRows, iNumOfColmns, "DMT User");
				oSLAuditDocument.SetCellStyle(iNumOfRows, iNumOfColmns, g_oBoldText);
				oSLAuditDocument.SetColumnWidth(iNumOfColmns, 20);

				iNumOfRows = 2;
				iNumOfColmns = 1;
				foreach (UserSecurityInfo oUser in g_oAllUsers)
				{
					// basic user info
					oSLAuditDocument.SetCellValue(iNumOfRows, iNumOfColmns, oUser.CompanyID);
					iNumOfColmns++;

					oSLAuditDocument.SetCellValue(iNumOfRows, iNumOfColmns, oUser.Disabled);
					if (oUser.Disabled == true)
					{
						oSLAuditDocument.SetCellStyle(iNumOfRows, iNumOfColmns, g_oBad);
					}
					else
					{
						oSLAuditDocument.SetCellStyle(iNumOfRows, iNumOfColmns, g_oGood);
					}

					iNumOfColmns++;
					oSLAuditDocument.SetCellValue(iNumOfRows, iNumOfColmns, oUser.SecurityMgr);
					if (oUser.SecurityMgr == true)
					{
						oSLAuditDocument.SetCellStyle(iNumOfRows, iNumOfColmns, g_oGood);
					}
					else
					{
						oSLAuditDocument.SetCellStyle(iNumOfRows, iNumOfColmns, g_oBad);
					}

					iNumOfColmns++;
					oSLAuditDocument.SetCellValue(iNumOfRows, iNumOfColmns, oUser.UserID);

					iNumOfColmns++;
					oSLAuditDocument.SetCellValue(iNumOfRows, iNumOfColmns, oUser.CanImpersonate);
					if (oUser.CanImpersonate == true)
					{
						oSLAuditDocument.SetCellStyle(iNumOfRows, iNumOfColmns, g_oGood);
					}
					else
					{
						oSLAuditDocument.SetCellStyle(iNumOfRows, iNumOfColmns, g_oBad);
					}

					iNumOfColmns++;
					oSLAuditDocument.SetCellValue(iNumOfRows, iNumOfColmns, oUser.CanCustomize);
					if (oUser.CanCustomize == true)
					{
						oSLAuditDocument.SetCellStyle(iNumOfRows, iNumOfColmns, g_oGood);
					}
					else
					{
						oSLAuditDocument.SetCellStyle(iNumOfRows, iNumOfColmns, g_oBad);
					}

					iNumOfColmns++;
					oSLAuditDocument.SetCellValue(iNumOfRows, iNumOfColmns, oUser.SolutionManagerCreate);
					if (oUser.SolutionManagerCreate == true)
					{
						oSLAuditDocument.SetCellStyle(iNumOfRows, iNumOfColmns, g_oGood);
					}
					else
					{
						oSLAuditDocument.SetCellStyle(iNumOfRows, iNumOfColmns, g_oBad);
					}

					iNumOfColmns++;
					oSLAuditDocument.SetCellValue(iNumOfRows, iNumOfColmns, oUser.SolutionManagerInstall);
					if (oUser.SolutionManagerInstall == true)
					{
						oSLAuditDocument.SetCellStyle(iNumOfRows, iNumOfColmns, g_oGood);
					}
					else
					{
						oSLAuditDocument.SetCellStyle(iNumOfRows, iNumOfColmns, g_oBad);
					}

					iNumOfColmns++;
					oSLAuditDocument.SetCellValue(iNumOfRows, iNumOfColmns, oUser.DashboardDeveloper);
					if (oUser.DashboardDeveloper == true)
					{
						oSLAuditDocument.SetCellStyle(iNumOfRows, iNumOfColmns, g_oGood);
					}
					else
					{
						oSLAuditDocument.SetCellStyle(iNumOfRows, iNumOfColmns, g_oBad);
					}

					iNumOfColmns++;
					oSLAuditDocument.SetCellValue(iNumOfRows, iNumOfColmns, oUser.BPMAdvancedDeveloper);
					if (oUser.BPMAdvancedDeveloper == true)
					{
						oSLAuditDocument.SetCellStyle(iNumOfRows, iNumOfColmns, g_oGood);
					}
					else
					{
						oSLAuditDocument.SetCellStyle(iNumOfRows, iNumOfColmns, g_oBad);
					}

					iNumOfColmns++;
					oSLAuditDocument.SetCellValue(iNumOfRows, iNumOfColmns, oUser.BPMAdvancedBAQRights);
					if (oUser.BPMAdvancedBAQRights == true)
					{
						oSLAuditDocument.SetCellStyle(iNumOfRows, iNumOfColmns, g_oGood);
					}
					else
					{
						oSLAuditDocument.SetCellStyle(iNumOfRows, iNumOfColmns, g_oBad);
					}

					iNumOfColmns++;
					oSLAuditDocument.SetCellValue(iNumOfRows, iNumOfColmns, oUser.SSRSReportDesigner);
					if (oUser.SSRSReportDesigner == true)
					{
						oSLAuditDocument.SetCellStyle(iNumOfRows, iNumOfColmns, g_oGood);
					}
					else
					{
						oSLAuditDocument.SetCellStyle(iNumOfRows, iNumOfColmns, g_oBad);
					}

					iNumOfColmns++;
					oSLAuditDocument.SetCellValue(iNumOfRows, iNumOfColmns, oUser.DMTUser);
					if (oUser.DMTUser == true)
					{
						oSLAuditDocument.SetCellStyle(iNumOfRows, iNumOfColmns, g_oGood);
					}
					else
					{
						oSLAuditDocument.SetCellStyle(iNumOfRows, iNumOfColmns, g_oBad);
					}

					// move to next row for user and reset column number
					iNumOfColmns = 1;
					iNumOfRows++;
				}
				#endregion

				#region SOD
				if (bFirstWorksheet == true)
				{
					oSLAuditDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "SOD Report");
					bFirstWorksheet = false;
				}
				else
				{
					oSLAuditDocument.AddWorksheet("SOD Report");
				}

				//set up row headers
				iNumOfRows = 0;
				iNumOfColmns = 1;

				// BASIC USER INFO
				iNumOfRows++;
				oSLAuditDocument.SetCellValue(iNumOfRows, iNumOfColmns, "Company ID");
				oSLAuditDocument.SetCellStyle(iNumOfRows, iNumOfColmns, g_oBoldText);
				oSLAuditDocument.SetColumnWidth(iNumOfColmns, 12);
				iNumOfColmns++;
				oSLAuditDocument.SetCellValue(iNumOfRows, iNumOfColmns, "Account Disabled");
				oSLAuditDocument.SetCellStyle(iNumOfRows, iNumOfColmns, g_oBoldText);
				oSLAuditDocument.SetColumnWidth(iNumOfColmns, 17);
				iNumOfColmns++;
				oSLAuditDocument.SetCellValue(iNumOfRows, iNumOfColmns, "User Id");
				oSLAuditDocument.SetCellStyle(iNumOfRows, iNumOfColmns, g_oBoldText);
				oSLAuditDocument.SetColumnWidth(iNumOfColmns, 15);

				// SPECIAL ROLES
				iNumOfColmns++;
				oSLAuditDocument.SetCellValue(iNumOfRows, iNumOfColmns, "Security Manager");
				oSLAuditDocument.SetCellStyle(iNumOfRows, iNumOfColmns, g_oBoldText);
				oSLAuditDocument.SetColumnWidth(iNumOfColmns, 17);

				// SOD PERMISSIONS
				iNumOfColmns++;
				oSLAuditDocument.SetCellValue(iNumOfRows, iNumOfColmns, "Adjustment");
				oSLAuditDocument.SetCellStyle(iNumOfRows, iNumOfColmns, g_oBoldText);
				oSLAuditDocument.SetColumnWidth(iNumOfColmns, 12);
				iNumOfColmns++;
				oSLAuditDocument.SetCellValue(iNumOfRows, iNumOfColmns, "Payment Entry");
				oSLAuditDocument.SetCellStyle(iNumOfRows, iNumOfColmns, g_oBoldText);
				oSLAuditDocument.SetColumnWidth(iNumOfColmns, 15);
				iNumOfColmns++;
				oSLAuditDocument.SetCellValue(iNumOfRows, iNumOfColmns, "AP Invoice Entry");
				oSLAuditDocument.SetCellStyle(iNumOfRows, iNumOfColmns, g_oBoldText);
				oSLAuditDocument.SetColumnWidth(iNumOfColmns, 15);
				iNumOfColmns++;
				oSLAuditDocument.SetCellValue(iNumOfRows, iNumOfColmns, "AR Invoice Entry");
				oSLAuditDocument.SetCellStyle(iNumOfRows, iNumOfColmns, g_oBoldText);
				oSLAuditDocument.SetColumnWidth(iNumOfColmns, 15);
				iNumOfColmns++;
				oSLAuditDocument.SetCellValue(iNumOfRows, iNumOfColmns, "Cash Receipt");
				oSLAuditDocument.SetCellStyle(iNumOfRows, iNumOfColmns, g_oBoldText);
				oSLAuditDocument.SetColumnWidth(iNumOfColmns, 15);
				iNumOfColmns++;
				oSLAuditDocument.SetCellValue(iNumOfRows, iNumOfColmns, "Apply Credit Memo");
				oSLAuditDocument.SetCellStyle(iNumOfRows, iNumOfColmns, g_oBoldText);
				oSLAuditDocument.SetColumnWidth(iNumOfColmns, 20);
				iNumOfColmns++;
				oSLAuditDocument.SetCellValue(iNumOfRows, iNumOfColmns, "Write Off And Adjustment");
				oSLAuditDocument.SetCellStyle(iNumOfRows, iNumOfColmns, g_oBoldText);
				oSLAuditDocument.SetColumnWidth(iNumOfColmns, 25);
				iNumOfColmns++;
				oSLAuditDocument.SetCellValue(iNumOfRows, iNumOfColmns, "Customer Credit Manager");
				oSLAuditDocument.SetCellStyle(iNumOfRows, iNumOfColmns, g_oBoldText);
				oSLAuditDocument.SetColumnWidth(iNumOfColmns, 25);
				iNumOfColmns++;
				oSLAuditDocument.SetCellValue(iNumOfRows, iNumOfColmns, "Journal Entry");
				oSLAuditDocument.SetCellStyle(iNumOfRows, iNumOfColmns, g_oBoldText);
				oSLAuditDocument.SetColumnWidth(iNumOfColmns, 15);
				iNumOfColmns++;
				oSLAuditDocument.SetCellValue(iNumOfRows, iNumOfColmns, "Cost Adjustment");
				oSLAuditDocument.SetCellStyle(iNumOfRows, iNumOfColmns, g_oBoldText);
				oSLAuditDocument.SetColumnWidth(iNumOfColmns, 15);
				iNumOfColmns++;
				oSLAuditDocument.SetCellValue(iNumOfRows, iNumOfColmns, "RMA Processing");
				oSLAuditDocument.SetCellStyle(iNumOfRows, iNumOfColmns, g_oBoldText);
				oSLAuditDocument.SetColumnWidth(iNumOfColmns, 15);
				iNumOfColmns++;
				oSLAuditDocument.SetCellValue(iNumOfRows, iNumOfColmns, "Customer");
				oSLAuditDocument.SetCellStyle(iNumOfRows, iNumOfColmns, g_oBoldText);
				oSLAuditDocument.SetColumnWidth(iNumOfColmns, 10);
				iNumOfColmns++;
				oSLAuditDocument.SetCellValue(iNumOfRows, iNumOfColmns, "Order Entry");
				oSLAuditDocument.SetCellStyle(iNumOfRows, iNumOfColmns, g_oBoldText);
				oSLAuditDocument.SetColumnWidth(iNumOfColmns, 12);
				iNumOfColmns++;
				oSLAuditDocument.SetCellValue(iNumOfRows, iNumOfColmns, "Purchase Order Entry");
				oSLAuditDocument.SetCellStyle(iNumOfRows, iNumOfColmns, g_oBoldText);
				oSLAuditDocument.SetColumnWidth(iNumOfColmns, 20);
				iNumOfColmns++;
				oSLAuditDocument.SetCellValue(iNumOfRows, iNumOfColmns, "PO Approval");
				oSLAuditDocument.SetCellStyle(iNumOfRows, iNumOfColmns, g_oBoldText);
				oSLAuditDocument.SetColumnWidth(iNumOfColmns, 12);
				iNumOfColmns++;
				oSLAuditDocument.SetCellValue(iNumOfRows, iNumOfColmns, "Supplier");
				oSLAuditDocument.SetCellStyle(iNumOfRows, iNumOfColmns, g_oBoldText);
				oSLAuditDocument.SetColumnWidth(iNumOfColmns, 10);
				iNumOfColmns++;
				oSLAuditDocument.SetCellValue(iNumOfRows, iNumOfColmns, "Customer Shipment Entry");
				oSLAuditDocument.SetCellStyle(iNumOfRows, iNumOfColmns, g_oBoldText);
				oSLAuditDocument.SetColumnWidth(iNumOfColmns, 25);
				iNumOfColmns++;
				oSLAuditDocument.SetCellValue(iNumOfRows, iNumOfColmns, "Receipt Entry");
				oSLAuditDocument.SetCellStyle(iNumOfRows, iNumOfColmns, g_oBoldText);
				oSLAuditDocument.SetColumnWidth(iNumOfColmns, 15);

				iNumOfRows = 2;
				iNumOfColmns = 1;
				foreach (UserSecurityInfo oUser in g_oAllUsers)
				{
					// basic user info
					oSLAuditDocument.SetCellValue(iNumOfRows, iNumOfColmns, oUser.CompanyID);
					iNumOfColmns++;

					oSLAuditDocument.SetCellValue(iNumOfRows, iNumOfColmns, oUser.Disabled);
					if (oUser.Disabled == true)
					{
						oSLAuditDocument.SetCellStyle(iNumOfRows, iNumOfColmns, g_oBad);
					}
					else
					{
						oSLAuditDocument.SetCellStyle(iNumOfRows, iNumOfColmns, g_oGood);
					}
					iNumOfColmns++;
					oSLAuditDocument.SetCellValue(iNumOfRows, iNumOfColmns, oUser.UserID);

					iNumOfColmns++;
					oSLAuditDocument.SetCellValue(iNumOfRows, iNumOfColmns, oUser.SecurityMgr);
					if (oUser.SecurityMgr == true)
					{
						oSLAuditDocument.SetCellStyle(iNumOfRows, iNumOfColmns, g_oGood);
					}
					else
					{
						oSLAuditDocument.SetCellStyle(iNumOfRows, iNumOfColmns, g_oBad);
					}

					// SOD

					// Adjustment
					iNumOfColmns++;
					bool bUserHasAccess = false;
					MenuSecurity oMenuSecurity = g_oMenuSecurityItems.FirstOrDefault(x => string.Compare(x.MenuDesc, "Adjustment", true) == 0);
					if (oMenuSecurity != null)
					{
						bUserHasAccess = oMenuSecurity.CanUserAccessMenuItem(oUser);
					}
					oSLAuditDocument.SetCellValue(iNumOfRows, iNumOfColmns, bUserHasAccess);
					if (bUserHasAccess == true)
					{
						oSLAuditDocument.SetCellStyle(iNumOfRows, iNumOfColmns, g_oGood);
					}
					else
					{
						oSLAuditDocument.SetCellStyle(iNumOfRows, iNumOfColmns, g_oBad);
					}

					// payment entry
					iNumOfColmns++;
					bUserHasAccess = false;
					oMenuSecurity = g_oMenuSecurityItems.FirstOrDefault(x => string.Compare(x.MenuDesc, "Payment Entry", true) == 0);
					if (oMenuSecurity != null)
					{
						bUserHasAccess = oMenuSecurity.CanUserAccessMenuItem(oUser);
					}
					oSLAuditDocument.SetCellValue(iNumOfRows, iNumOfColmns, bUserHasAccess);
					if (bUserHasAccess == true)
					{
						oSLAuditDocument.SetCellStyle(iNumOfRows, iNumOfColmns, g_oGood);
					}
					else
					{
						oSLAuditDocument.SetCellStyle(iNumOfRows, iNumOfColmns, g_oBad);
					}

					// AP Invoice Entry
					iNumOfColmns++;
					bUserHasAccess = false;
					oMenuSecurity = g_oMenuSecurityItems.FirstOrDefault(x => string.Compare(x.MenuDesc, "AP Invoice Entry", true) == 0);
					if (oMenuSecurity != null)
					{
						bUserHasAccess = oMenuSecurity.CanUserAccessMenuItem(oUser);
					}
					oSLAuditDocument.SetCellValue(iNumOfRows, iNumOfColmns, bUserHasAccess);
					if (bUserHasAccess == true)
					{
						oSLAuditDocument.SetCellStyle(iNumOfRows, iNumOfColmns, g_oGood);
					}
					else
					{
						oSLAuditDocument.SetCellStyle(iNumOfRows, iNumOfColmns, g_oBad);
					}

					// AR Invoice Entry
					iNumOfColmns++;
					bUserHasAccess = false;
					oMenuSecurity = g_oMenuSecurityItems.FirstOrDefault(x => string.Compare(x.MenuDesc, "AR Invoice Entry", true) == 0);
					if (oMenuSecurity != null)
					{
						bUserHasAccess = oMenuSecurity.CanUserAccessMenuItem(oUser);
					}
					oSLAuditDocument.SetCellValue(iNumOfRows, iNumOfColmns, bUserHasAccess);
					if (bUserHasAccess == true)
					{
						oSLAuditDocument.SetCellStyle(iNumOfRows, iNumOfColmns, g_oGood);
					}
					else
					{
						oSLAuditDocument.SetCellStyle(iNumOfRows, iNumOfColmns, g_oBad);
					}

					// Cash Receipt
					iNumOfColmns++;
					bUserHasAccess = false;
					oMenuSecurity = g_oMenuSecurityItems.FirstOrDefault(x => string.Compare(x.MenuDesc, "Cash Receipt Entry", true) == 0);
					if (oMenuSecurity != null)
					{
						bUserHasAccess = oMenuSecurity.CanUserAccessMenuItem(oUser);
					}
					oSLAuditDocument.SetCellValue(iNumOfRows, iNumOfColmns, bUserHasAccess);
					if (bUserHasAccess == true)
					{
						oSLAuditDocument.SetCellStyle(iNumOfRows, iNumOfColmns, g_oGood);
					}
					else
					{
						oSLAuditDocument.SetCellStyle(iNumOfRows, iNumOfColmns, g_oBad);
					}

					// Apply Credit Memo
					iNumOfColmns++;
					bUserHasAccess = false;
					oMenuSecurity = g_oMenuSecurityItems.FirstOrDefault(x => string.Compare(x.MenuDesc, "Apply Credit Memo", true) == 0);
					if (oMenuSecurity != null)
					{
						bUserHasAccess = oMenuSecurity.CanUserAccessMenuItem(oUser);
					}
					oSLAuditDocument.SetCellValue(iNumOfRows, iNumOfColmns, bUserHasAccess);
					if (bUserHasAccess == true)
					{
						oSLAuditDocument.SetCellStyle(iNumOfRows, iNumOfColmns, g_oGood);
					}
					else
					{
						oSLAuditDocument.SetCellStyle(iNumOfRows, iNumOfColmns, g_oBad);
					}

					// Write Off And Adjustment
					iNumOfColmns++;
					bUserHasAccess = false;
					oMenuSecurity = g_oMenuSecurityItems.FirstOrDefault(x => string.Compare(x.MenuDesc, "Write Off And Adjustment", true) == 0);
					if (oMenuSecurity != null)
					{
						bUserHasAccess = oMenuSecurity.CanUserAccessMenuItem(oUser);
					}
					oSLAuditDocument.SetCellValue(iNumOfRows, iNumOfColmns, bUserHasAccess);
					if (bUserHasAccess == true)
					{
						oSLAuditDocument.SetCellStyle(iNumOfRows, iNumOfColmns, g_oGood);
					}
					else
					{
						oSLAuditDocument.SetCellStyle(iNumOfRows, iNumOfColmns, g_oBad);
					}

					// Customer Credit Manager
					iNumOfColmns++;
					bUserHasAccess = false;
					oMenuSecurity = g_oMenuSecurityItems.FirstOrDefault(x => string.Compare(x.MenuDesc, "Customer Credit Manager", true) == 0);
					if (oMenuSecurity != null)
					{
						bUserHasAccess = oMenuSecurity.CanUserAccessMenuItem(oUser);
					}
					oSLAuditDocument.SetCellValue(iNumOfRows, iNumOfColmns, bUserHasAccess);
					if (bUserHasAccess == true)
					{
						oSLAuditDocument.SetCellStyle(iNumOfRows, iNumOfColmns, g_oGood);
					}
					else
					{
						oSLAuditDocument.SetCellStyle(iNumOfRows, iNumOfColmns, g_oBad);
					}

					// Journal Entry
					iNumOfColmns++;
					bUserHasAccess = false;
					oMenuSecurity = g_oMenuSecurityItems.FirstOrDefault(x => string.Compare(x.MenuDesc, "Journal Entry", true) == 0);
					if (oMenuSecurity != null)
					{
						bUserHasAccess = oMenuSecurity.CanUserAccessMenuItem(oUser);
					}
					oSLAuditDocument.SetCellValue(iNumOfRows, iNumOfColmns, bUserHasAccess);
					if (bUserHasAccess == true)
					{
						oSLAuditDocument.SetCellStyle(iNumOfRows, iNumOfColmns, g_oGood);
					}
					else
					{
						oSLAuditDocument.SetCellStyle(iNumOfRows, iNumOfColmns, g_oBad);
					}

					// Cost Adjustment
					iNumOfColmns++;
					bUserHasAccess = false;
					oMenuSecurity = g_oMenuSecurityItems.FirstOrDefault(x => string.Compare(x.MenuDesc, "Cost Adjustment", true) == 0);
					if (oMenuSecurity != null)
					{
						bUserHasAccess = oMenuSecurity.CanUserAccessMenuItem(oUser);
					}
					oSLAuditDocument.SetCellValue(iNumOfRows, iNumOfColmns, bUserHasAccess);
					if (bUserHasAccess == true)
					{
						oSLAuditDocument.SetCellStyle(iNumOfRows, iNumOfColmns, g_oGood);
					}
					else
					{
						oSLAuditDocument.SetCellStyle(iNumOfRows, iNumOfColmns, g_oBad);
					}

					// RMA Processing
					iNumOfColmns++;
					bUserHasAccess = false;
					oMenuSecurity = g_oMenuSecurityItems.FirstOrDefault(x => string.Compare(x.MenuDesc, "RMA Processing", true) == 0);
					if (oMenuSecurity != null)
					{
						bUserHasAccess = oMenuSecurity.CanUserAccessMenuItem(oUser);
					}
					oSLAuditDocument.SetCellValue(iNumOfRows, iNumOfColmns, bUserHasAccess);
					if (bUserHasAccess == true)
					{
						oSLAuditDocument.SetCellStyle(iNumOfRows, iNumOfColmns, g_oGood);
					}
					else
					{
						oSLAuditDocument.SetCellStyle(iNumOfRows, iNumOfColmns, g_oBad);
					}

					// Customer
					iNumOfColmns++;
					bUserHasAccess = false;
					oMenuSecurity = g_oMenuSecurityItems.FirstOrDefault(x => string.Compare(x.MenuDesc, "Customer", true) == 0);
					if (oMenuSecurity != null)
					{
						bUserHasAccess = oMenuSecurity.CanUserAccessMenuItem(oUser);
					}
					oSLAuditDocument.SetCellValue(iNumOfRows, iNumOfColmns, bUserHasAccess);
					if (bUserHasAccess == true)
					{
						oSLAuditDocument.SetCellStyle(iNumOfRows, iNumOfColmns, g_oGood);
					}
					else
					{
						oSLAuditDocument.SetCellStyle(iNumOfRows, iNumOfColmns, g_oBad);
					}

					// Order Entry
					iNumOfColmns++;
					bUserHasAccess = false;
					oMenuSecurity = g_oMenuSecurityItems.FirstOrDefault(x => string.Compare(x.MenuDesc, "Order Entry", true) == 0);
					if (oMenuSecurity != null)
					{
						bUserHasAccess = oMenuSecurity.CanUserAccessMenuItem(oUser);
					}
					oSLAuditDocument.SetCellValue(iNumOfRows, iNumOfColmns, bUserHasAccess);
					if (bUserHasAccess == true)
					{
						oSLAuditDocument.SetCellStyle(iNumOfRows, iNumOfColmns, g_oGood);
					}
					else
					{
						oSLAuditDocument.SetCellStyle(iNumOfRows, iNumOfColmns, g_oBad);
					}

					// Purchase Order Entry
					iNumOfColmns++;
					bUserHasAccess = false;
					oMenuSecurity = g_oMenuSecurityItems.FirstOrDefault(x => string.Compare(x.MenuDesc, "Purchase Order Entry", true) == 0);
					if (oMenuSecurity != null)
					{
						bUserHasAccess = oMenuSecurity.CanUserAccessMenuItem(oUser);
					}
					oSLAuditDocument.SetCellValue(iNumOfRows, iNumOfColmns, bUserHasAccess);
					if (bUserHasAccess == true)
					{
						oSLAuditDocument.SetCellStyle(iNumOfRows, iNumOfColmns, g_oGood);
					}
					else
					{
						oSLAuditDocument.SetCellStyle(iNumOfRows, iNumOfColmns, g_oBad);
					}

					// PO Approval
					iNumOfColmns++;
					bUserHasAccess = false;
					oMenuSecurity = g_oMenuSecurityItems.FirstOrDefault(x => string.Compare(x.MenuDesc, "PO Approval", true) == 0);
					if (oMenuSecurity != null)
					{
						bUserHasAccess = oMenuSecurity.CanUserAccessMenuItem(oUser);
					}
					oSLAuditDocument.SetCellValue(iNumOfRows, iNumOfColmns, bUserHasAccess);
					if (bUserHasAccess == true)
					{
						oSLAuditDocument.SetCellStyle(iNumOfRows, iNumOfColmns, g_oGood);
					}
					else
					{
						oSLAuditDocument.SetCellStyle(iNumOfRows, iNumOfColmns, g_oBad);
					}

					// Supplier
					iNumOfColmns++;
					bUserHasAccess = false;
					oMenuSecurity = g_oMenuSecurityItems.FirstOrDefault(x => string.Compare(x.MenuDesc, "Supplier", true) == 0);
					if (oMenuSecurity != null)
					{
						bUserHasAccess = oMenuSecurity.CanUserAccessMenuItem(oUser);
					}
					oSLAuditDocument.SetCellValue(iNumOfRows, iNumOfColmns, bUserHasAccess);
					if (bUserHasAccess == true)
					{
						oSLAuditDocument.SetCellStyle(iNumOfRows, iNumOfColmns, g_oGood);
					}
					else
					{
						oSLAuditDocument.SetCellStyle(iNumOfRows, iNumOfColmns, g_oBad);
					}

					// Customer Shipment Entry
					iNumOfColmns++;
					bUserHasAccess = false;
					oMenuSecurity = g_oMenuSecurityItems.FirstOrDefault(x => string.Compare(x.MenuDesc, "Customer Shipment Entry", true) == 0);
					if (oMenuSecurity != null)
					{
						bUserHasAccess = oMenuSecurity.CanUserAccessMenuItem(oUser);
					}
					oSLAuditDocument.SetCellValue(iNumOfRows, iNumOfColmns, bUserHasAccess);
					if (bUserHasAccess == true)
					{
						oSLAuditDocument.SetCellStyle(iNumOfRows, iNumOfColmns, g_oGood);
					}
					else
					{
						oSLAuditDocument.SetCellStyle(iNumOfRows, iNumOfColmns, g_oBad);
					}

					// Receipt Entry
					iNumOfColmns++;
					bUserHasAccess = false;
					oMenuSecurity = g_oMenuSecurityItems.FirstOrDefault(x => string.Compare(x.MenuDesc, "Receipt Entry", true) == 0);
					if (oMenuSecurity != null)
					{
						bUserHasAccess = oMenuSecurity.CanUserAccessMenuItem(oUser);
					}
					oSLAuditDocument.SetCellValue(iNumOfRows, iNumOfColmns, bUserHasAccess);
					if (bUserHasAccess == true)
					{
						oSLAuditDocument.SetCellStyle(iNumOfRows, iNumOfColmns, g_oGood);
					}
					else
					{
						oSLAuditDocument.SetCellStyle(iNumOfRows, iNumOfColmns, g_oBad);
					}

					// move to next row for user and reset column number
					iNumOfRows++;
					iNumOfColmns = 1;
				}
				#endregion

				#region User Groups
				if (bFirstWorksheet == true)
				{
					oSLAuditDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "User Groups Report");
					bFirstWorksheet = false;
				}
				else
				{
					oSLAuditDocument.AddWorksheet("User Groups Report");
				}

				//set up column headers
				iNumOfRows = 0;
				iNumOfColmns = 1;

				// BASIC USER INFO
				iNumOfRows++;
				oSLAuditDocument.SetCellValue(iNumOfRows, iNumOfColmns, "Company ID");
				oSLAuditDocument.SetCellStyle(iNumOfRows, iNumOfColmns, g_oBoldText);
				oSLAuditDocument.SetColumnWidth(iNumOfColmns, 12);
				iNumOfColmns++;
				oSLAuditDocument.SetCellValue(iNumOfRows, iNumOfColmns, "Account Disabled");
				oSLAuditDocument.SetColumnWidth(iNumOfColmns, 20);
				oSLAuditDocument.SetCellStyle(iNumOfRows, iNumOfColmns, g_oBoldText);
				iNumOfColmns++;
				oSLAuditDocument.SetCellValue(iNumOfRows, iNumOfColmns, "User Id");
				oSLAuditDocument.SetColumnWidth(iNumOfColmns, 20);
				oSLAuditDocument.SetCellStyle(iNumOfRows, iNumOfColmns, g_oBoldText);
				// USER GROUPS
				iNumOfColmns++;
				oSLAuditDocument.SetCellValue(iNumOfRows, iNumOfColmns, "Groups");
				oSLAuditDocument.SetColumnWidth(iNumOfColmns, 250);
				oSLAuditDocument.SetCellStyle(iNumOfRows, iNumOfColmns, g_oBoldText);

				iNumOfRows = 2;
				iNumOfColmns = 1;
				foreach (UserSecurityInfo oUser in g_oAllUsers)
				{
					// basic user info
					oSLAuditDocument.SetCellValue(iNumOfRows, iNumOfColmns, oUser.CompanyID);
					iNumOfColmns++;

					oSLAuditDocument.SetCellValue(iNumOfRows, iNumOfColmns, oUser.Disabled);
					if (oUser.Disabled == true)
					{
						oSLAuditDocument.SetCellStyle(iNumOfRows, iNumOfColmns, g_oBad);
					}
					else
					{
						oSLAuditDocument.SetCellStyle(iNumOfRows, iNumOfColmns, g_oGood);
					}

					iNumOfColmns++;
					oSLAuditDocument.SetCellValue(iNumOfRows, iNumOfColmns, oUser.UserID);

					// LIST GROUPS FOR USER 
					iNumOfColmns++;
					string sGroups = "";
					foreach (ActiveUserGroup oGroup in oUser.Groups)
					{
						sGroups += oGroup.Description + ":";
					}
					oSLAuditDocument.SetCellValue(iNumOfRows, iNumOfColmns, sGroups);

					// move to next row for user and reset column number
					iNumOfColmns = 1;
					iNumOfRows++;
				}
				#endregion
			}


			// send the email to the shipping team
			if (bFirstWorksheet == false)
			{
				// we create a worksheet so we have some data in the spreadsheet so we email the spreadsheet
				oSLAuditDocument.SaveAs(sDestinationFileName);
				List<string> oAttachments = new List<string>();
				oAttachments.Add(sDestinationFileName);
				HSEmailHelper.SendEmail(oToAddresses, "Audit Report", "Audit Report", oAttachments);
			}
		}


		#region Data Members
		private static SLStyle g_oGood;
		private static SLStyle g_oNeutral;
		private static SLStyle g_oBad;
		private static SLStyle g_oBoldText;

		private static List<MenuSecurity> g_oMenuSecurityItems = new List<MenuSecurity>();
		private static List<UserSecurityInfo> g_oAllUsers = new List<UserSecurityInfo>();
		private static List<ActiveUserGroup> g_oActiveUserGroups = new List<ActiveUserGroup>();
		#endregion
	}

	public class MenuSecurity
	{
		#region Constructors
		public MenuSecurity(DataRow oDataRow, List<UserSecurityInfo> oUsers, List<ActiveUserGroup> oGroups)
		{
			if ((oDataRow["Menu_MenuID"] != DBNull.Value) && (string.IsNullOrEmpty((string)oDataRow["Menu_MenuID"]) == false))
			{
				m_sMenuID = (string)oDataRow["Menu_MenuID"];
			}
			if ((oDataRow["Calculated_MenuDesc"] != DBNull.Value) && (string.IsNullOrEmpty((string)oDataRow["Calculated_MenuDesc"]) == false))
			{
				m_sMenuDesc = (string)oDataRow["Calculated_MenuDesc"];
			}
			if ((oDataRow["Calculated_ParentMenuInfo"] != DBNull.Value) && (string.IsNullOrEmpty((string)oDataRow["Calculated_ParentMenuInfo"]) == false))
			{
				m_sParentInfo = (string)oDataRow["Calculated_ParentMenuInfo"];
			}
			if ((oDataRow["Menu_MenuType"] != DBNull.Value) && (string.IsNullOrEmpty((string)oDataRow["Menu_MenuType"]) == false))
			{
				m_sMenuType = (string)oDataRow["Menu_MenuType"];
			}
			if ((oDataRow["Security_SecCode"] != DBNull.Value) && (string.IsNullOrEmpty((string)oDataRow["Security_SecCode"]) == false))
			{
				m_sSecurityCode = (string)oDataRow["Security_SecCode"];
			}
			if ((oDataRow["Security_SecurityMgr"] != DBNull.Value))
			{
				m_bSecurityMgr = (bool)oDataRow["Security_SecurityMgr"];
			}
			if ((oDataRow["Security_NoEntryList"] != DBNull.Value) && (string.IsNullOrEmpty((string)oDataRow["Security_NoEntryList"]) == false))
			{
				string sNoEntry = (string)oDataRow["Security_NoEntryList"];
				List<string> oNoEntryList = sNoEntry.Split(',').ToList();
				foreach (string sTmpNoEntry in oNoEntryList)
				{
					// if the list contains an * then no one is allowed access
					if (string.Compare(sTmpNoEntry, "*", true) == 0)
					{
						m_bDisallowEveryone = true;
						break;
					}

					// get the user from the user list or get the group from the group list
					UserSecurityInfo oUser = oUsers.FirstOrDefault(oItem => string.Compare(oItem.UserID, sTmpNoEntry, true) == 0);
					if (oUser != null)
					{
						// we have found the user so put the user in the no entry list
						if (m_oNoEntryList.Contains(oUser) == false)
						{
							m_oNoEntryList.Add(oUser);
						}
					}
					else
					{
						// check to see if this is a group instead
						ActiveUserGroup oGroup = oGroups.FirstOrDefault(oItem => string.Compare(oItem.Code, sTmpNoEntry, true) == 0);
						if (oGroup != null)
						{
							// we have found the group so we will extract all users from this group and add them to the no entry list
							foreach (UserSecurityInfo oTmpUser in oGroup.Users)
							{
								if (m_oNoEntryList.Contains(oTmpUser) == false)
								{
									m_oNoEntryList.Add(oTmpUser);
								}
							}
						}
					}
				}
			}
			if ((oDataRow["Security_EntryList"] != DBNull.Value) && (string.IsNullOrEmpty((string)oDataRow["Security_EntryList"]) == false))
			{
				string sEntry = (string)oDataRow["Security_EntryList"];
				List<string> oEntryList = sEntry.Split(',').ToList();
				foreach (string sTmpEntry in oEntryList)
				{
					// if the list contains an * then everyone is allowed access
					if (string.Compare(sTmpEntry, "*", true) == 0)
					{
						m_bAllowEveryone = true;
						break;
					}

					// get the user from the user list or get the group from the group list
					UserSecurityInfo oUser = oUsers.FirstOrDefault(oItem => string.Compare(oItem.UserID, sTmpEntry, true) == 0);
					if (oUser != null)
					{
						// we have found the user so put the user in the entry list
						if (m_oEntryList.Contains(oUser) == false)
						{
							m_oEntryList.Add(oUser);
						}
					}
					else
					{
						// check to see if this is a group instead
						ActiveUserGroup oGroup = oGroups.FirstOrDefault(oItem => string.Compare(oItem.Code, sTmpEntry, true) == 0);
						if (oGroup != null)
						{
							// we have found the group so we will extract all users from this group and add them to the no entry list
							foreach (UserSecurityInfo oTmpUser in oGroup.Users)
							{
								if (m_oEntryList.Contains(oTmpUser) == false)
								{
									m_oEntryList.Add(oTmpUser);
								}
							}
						}
					}
				}
			}
		}
		#endregion

		#region Methods
		public bool CanUserAccessMenuItem(UserSecurityInfo oUser)
		{
			bool bAccess = false;
			if (oUser.SecurityMgr == true)
			{
				bAccess = true;
			}
			else if (m_bSecurityMgr == true)
			{
				// only a security manage has access
				bAccess = oUser.SecurityMgr;
			}
			else if (m_bDisallowEveryone == true)
			{
				// no one is allowed to access this menu item
				bAccess = false;
			}
			else
			{
				// if user is in the disallow list we do not continue as the user cannot access the menu item
				if (m_oNoEntryList.Contains(oUser) == true)
				{
					// this user has explicitly been denied access to this menu item
					bAccess = false;
				}
				else
				{
					if (m_oEntryList.Contains(oUser) == true)
					{
						// this user has explicitly been granted access to this menu item
						bAccess = true;
					}
					else if (m_bAllowEveryone == true)
					{
						// if we get to this point and the user has not been
						// explicitly denied, then everyone is allowed to access this menu item
						bAccess = true;
					}
				}
			}

			return bAccess;
		}
		#endregion

		#region Properties

		public string MenuID
		{
			get { return m_sMenuID; }
		}
		public string MenuType
		{
			get { return m_sMenuType; }
		}
		public string MenuDesc
		{
			get { return m_sMenuDesc; }
		}
		public string ParentInfo
		{
			get { return m_sParentInfo; }
		}
		public string SecurityCode
		{
			get { return m_sSecurityCode; }
		}
		public bool SecurityMgr
		{
			get { return m_bSecurityMgr; }
		}
		public bool AllowAccessToEveryone
		{
			get { return m_bAllowEveryone; }
		}
		public bool DisallowAccessToEveryone
		{
			get { return m_bDisallowEveryone; }
		}
		#endregion

		#region Data Members

		private string m_sMenuID;
		private string m_sMenuDesc;
		private string m_sParentInfo;
		private string m_sMenuType;
		private string m_sSecurityCode;
		private bool m_bSecurityMgr;
		private bool m_bAllowEveryone;
		private bool m_bDisallowEveryone;

		private List<UserSecurityInfo> m_oNoEntryList = new List<UserSecurityInfo>();
		private List<UserSecurityInfo> m_oEntryList = new List<UserSecurityInfo>();

		#endregion
	}

	public class UserSecurityInfo
	{
		#region Constructors
		public UserSecurityInfo(DataRow oDataRow, List<ActiveUserGroup> oGroups)
		{
			m_sCompanyID = (string)oDataRow["UserComp_Company"];
			if ((oDataRow["UserFile_UserDisabled"] != DBNull.Value))
			{
				m_bDisabled = (bool)oDataRow["UserFile_UserDisabled"];
			}
			m_sUserID = (string)oDataRow["UserFile_DcdUserID"];
			if ((oDataRow["UserFile_Name"] != DBNull.Value) && (string.IsNullOrEmpty((string)oDataRow["UserFile_Name"]) == false))
			{
				m_sName = (string)oDataRow["UserFile_Name"];
			}
			if ((oDataRow["UserFile_OSUserID"] != DBNull.Value) && (string.IsNullOrEmpty((string)oDataRow["UserFile_OSUserID"]) == false))
			{
				m_sOSUserId = (string)oDataRow["UserFile_OSUserID"];
			}
			if ((oDataRow["UserFile_SecurityMgr"] != DBNull.Value))
			{
				m_bSecurityMgr = (bool)oDataRow["UserFile_SecurityMgr"];
			}
			if (oDataRow["UserFile_PwdLastChanged"] != DBNull.Value)
			{
				m_dtPwdLastChanged = (DateTime)oDataRow["UserFile_PwdLastChanged"];
			}
			if (oDataRow["UserFile_PwdExpiresDays"] != DBNull.Value)
			{
				m_iPwdExpiresDays = (int)oDataRow["UserFile_PwdExpiresDays"];
			}
			if (oDataRow["UserFile_PwdExpires"] != DBNull.Value)
			{
				m_dtPwdExpires = (DateTime)oDataRow["UserFile_PwdExpires"];
			}
			if ((oDataRow["UserFile_RequireSso"] != DBNull.Value))
			{
				m_bRequireSSO = (bool)oDataRow["UserFile_RequireSso"];
			}
			if ((oDataRow["UserFile_DomainName"] != DBNull.Value) && (string.IsNullOrEmpty((string)oDataRow["UserFile_DomainName"]) == false))
			{
				m_sDomainName = (string)oDataRow["UserFile_DomainName"];
			}
			if (oDataRow["SysUserFile_CreateDate_c"] != DBNull.Value)
			{
				m_dtCreateDate = (DateTime)oDataRow["SysUserFile_CreateDate_c"];
			}
			if (oDataRow["SysUserFile_LastLogOnAttempt"] != DBNull.Value)
			{
				m_dtLastLogOnAttempt = (DateTime)oDataRow["SysUserFile_LastLogOnAttempt"];
			}
			if ((oDataRow["UserFile_GroupList"] != DBNull.Value) && (string.IsNullOrEmpty((string)oDataRow["UserFile_GroupList"]) == false))
			{
				string sTemp = (string)oDataRow["UserFile_GroupList"];
				List<string> oGroupCodes = sTemp.Split('~').ToList();
				m_oGroups.Clear();
				foreach (string sGroupCode in oGroupCodes)
				{
					ActiveUserGroup oGroup = oGroups.FirstOrDefault(oItem => string.Compare(oItem.Code, sGroupCode, true) == 0);
					if (oGroup != null)
					{
						m_oGroups.Add(oGroup);
						oGroup.AddUser(this);
					}
				}
			}
			if ((oDataRow["UserFile_CanCustomize"] != DBNull.Value))
			{
				m_bCanCustomize = (bool)oDataRow["UserFile_CanCustomize"];
			}
			if ((oDataRow["UserFile_DashboardDeveloper"] != DBNull.Value))
			{
				m_dDashboardDeveloper = (bool)oDataRow["UserFile_DashboardDeveloper"];
			}
			if ((oDataRow["UserFile_BPMAdvancedUser"] != DBNull.Value))
			{
				m_bBPMAdvancedDeveloper = (bool)oDataRow["UserFile_BPMAdvancedUser"];
			}
			if ((oDataRow["UserFile_AdvBAQRights"] != DBNull.Value))
			{
				m_bBPMAdvancedBAQRights = (bool)oDataRow["UserFile_AdvBAQRights"];
			}
			if ((oDataRow["UserFile_SolutionMgrCreate"] != DBNull.Value))
			{
				m_bSolutionManagerCreate = (bool)oDataRow["UserFile_SolutionMgrCreate"];
			}
			if ((oDataRow["UserFile_SolutionMgrInstall"] != DBNull.Value))
			{
				m_bSolutionManagerInstall = (bool)oDataRow["UserFile_SolutionMgrInstall"];
			}
			if (m_bSecurityMgr == true)
			{
				// these flags do not exist in this version
				m_bDMTUser = true;
				m_bSSRSReportDesigner = true;
			}
			//if ((oDataRow["SysUserFile_SSRSReportDesigner"] != DBNull.Value))
			//{
			//	m_bSSRSReportDesigner = (bool)oDataRow["SysUserFile_SSRSReportDesigner"];
			//}
			//if ((oDataRow["SysUserFile_DMTUser"] != DBNull.Value))
			//{
			//	m_bDMTUser = (bool)oDataRow["SysUserFile_DMTUser"];
			//}
			if ((oDataRow["SysUserFile_CanImpersonate"] != DBNull.Value))
			{
				m_bCanImpersonate = (bool)oDataRow["SysUserFile_CanImpersonate"];
			}
		}
		#endregion

		#region Properties

		public string CompanyID
		{
			get { return m_sCompanyID; }
		}
		public bool Disabled
		{
			get { return m_bDisabled; }
		}
		public string UserID
		{
			get { return m_sUserID; }
		}
		public string OSUserID
		{
			get { return m_sOSUserId; }
		}
		public string Name
		{
			get { return m_sName; }
		}
		public bool SecurityMgr
		{
			get { return m_bSecurityMgr; }
		}
		public DateTime PwdLastChanged
		{
			get { return m_dtPwdLastChanged; }
		}
		public int PwdExpiresDays
		{
			get { return m_iPwdExpiresDays; }
		}
		public DateTime PwdExpires
		{
			get { return m_dtPwdExpires; }
		}
		public bool RequireSSO
		{
			get { return m_bRequireSSO; }
		}
		public string DomainName
		{
			get { return m_sDomainName; }
		}
		public DateTime CreateDate
		{
			get { return m_dtCreateDate; }
		}
		public DateTime LastLogOnAttempt
		{
			get { return m_dtLastLogOnAttempt; }
		}
		public List<ActiveUserGroup> Groups
		{
			get { return m_oGroups; }
		}
		public bool CanCustomize
		{
			get { return m_bCanCustomize; }
		}
		public bool DashboardDeveloper
		{
			get { return m_dDashboardDeveloper; }
		}
		public bool BPMAdvancedDeveloper
		{
			get { return m_bBPMAdvancedDeveloper; }
		}
		public bool BPMAdvancedBAQRights
		{
			get { return m_bBPMAdvancedBAQRights; }
		}
		public bool SolutionManagerCreate
		{
			get { return m_bSolutionManagerCreate; }
		}
		public bool SolutionManagerInstall
		{
			get { return m_bSolutionManagerInstall; }
		}
		public bool SSRSReportDesigner
		{
			get { return m_bSSRSReportDesigner; }
		}
		public bool DMTUser
		{
			get { return m_bDMTUser; }
		}
		public bool CanImpersonate
		{
			get { return m_bCanImpersonate; }
		}
		#endregion

		#region Data Members

		private string m_sCompanyID;
		private bool m_bDisabled;
		private string m_sUserID;
		private string m_sOSUserId;
		private string m_sName;
		private bool m_bSecurityMgr;
		private DateTime m_dtPwdLastChanged;
		private int m_iPwdExpiresDays;
		private DateTime m_dtPwdExpires;
		private bool m_bRequireSSO;
		private string m_sDomainName;
		private DateTime m_dtCreateDate;
		private DateTime m_dtLastLogOnAttempt;
		private bool m_bCanCustomize;
		private bool m_dDashboardDeveloper;
		private bool m_bBPMAdvancedDeveloper;
		private bool m_bBPMAdvancedBAQRights;
		private bool m_bSolutionManagerCreate;
		private bool m_bSolutionManagerInstall;
		private bool m_bSSRSReportDesigner;
		private bool m_bDMTUser;
		private bool m_bCanImpersonate;
		private List<ActiveUserGroup> m_oGroups = new List<ActiveUserGroup>();

		#endregion
	}

	public class ActiveUserGroup
	{
		#region Constructors
		public ActiveUserGroup(DataRow oDataRow)
		{
			if ((oDataRow["SecGroup_SecGroupCode"] != DBNull.Value) && (string.IsNullOrEmpty((string)oDataRow["SecGroup_SecGroupCode"]) == false))
			{
				m_sCode = (string)oDataRow["SecGroup_SecGroupCode"];
			}
			if ((oDataRow["SecGroup_SecGroupDesc"] != DBNull.Value) && (string.IsNullOrEmpty((string)oDataRow["SecGroup_SecGroupDesc"]) == false))
			{
				m_sDescription = (string)oDataRow["SecGroup_SecGroupDesc"];
			}
		}
		#endregion

		#region Methods

		public void AddUser(UserSecurityInfo oUser)
		{
			m_oUsers.Add(oUser);
		}

		#endregion

		#region Properties

		public string Code
		{
			get { return m_sCode; }
		}
		public string Description
		{
			get { return m_sDescription; }
		}
		public List<UserSecurityInfo> Users
		{
			get { return m_oUsers; }
		}
		#endregion

		#region Data Members

		private string m_sCode;
		private string m_sDescription;

		private List<UserSecurityInfo> m_oUsers = new List<UserSecurityInfo>();

		#endregion
	}
}
