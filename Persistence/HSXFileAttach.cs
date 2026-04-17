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
using Erp.Adapters;
using Ice.Tablesets;

namespace HorizonScientific
{
    public class HSXFileAttach
    {
        public HSXFileAttach()
        {

        }

        public void Initialize(Session oSession, string sFileName, string sFullPathFileName, string sPartDescription, string sRev, string sDestinationFolder)
        {
            string sErrorMessage = string.Empty;

            PartImpl oPartImpl = WCFServiceSupport.CreateImpl<PartImpl>(oSession, Erp.Proxy.BO.PartImpl.UriPath);

            // read in 100 parts at a time
            int iPageSize = 100;
            // start with the first page
            int iCurrentPage = 1;
            // see if there are any more pages to load
            bool bMorePages = true;

            string sWhereClause = "CONTAINS(PartDescription,'" + sPartDescription + "') ";
            string sPartRevClause = sRev; 
            string sPartNum = string.Empty;
            PartDataSet oTmpList = oPartImpl.GetRows(sWhereClause, "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", iPageSize, iCurrentPage++, out bMorePages);
            if ((oTmpList.Tables[0].Rows == null) || (oTmpList.Tables[0].Rows.Count == 0))
            {
                sErrorMessage = "Could not find a part with the description '" + sPartDescription + "'";
                throw new Exception(sErrorMessage);
            }
            else if (oTmpList.Tables[0].Rows.Count > 1)
            {
                 sErrorMessage = "Found more than one part with the description '" + sPartDescription + "'";
                throw new Exception(sErrorMessage);
            }
            else
            {
                // we have exactly one part with this description
                // we will now get the part num so we can get the correct dataset
                DataRow oDataRow = oTmpList.Tables[0].Rows[0];
                sPartNum = (string)oDataRow["PartNum"];
                // get the data set from the part num andthe rev
                PartDataSet oPartDataSet = oPartImpl.GetByID(sPartNum);

                DataRow oPartRevDataSet = null;
                sErrorMessage = GetPartRevFromPart(oPartDataSet, sPartNum, sRev, out oPartRevDataSet);
                // check to see if we could find the rev number
                if (sErrorMessage.Contains(" has no revision ") == true)
                {
                    // the rev does not exist so we must first create the rev
                    string sRevisionNum = "";
                    string sAltMethod = "";
                    oPartImpl.GetNewPartRev(oPartDataSet, sPartNum, sRevisionNum, sAltMethod);
                    DataRow drPartRev = oPartDataSet.PartRev.Rows[oPartDataSet.PartRev.Rows.Count - 1];
                    drPartRev.BeginEdit();
                    drPartRev["RevisionNum"] = sRev;
                    drPartRev["RevShortDesc"] = "REV " + sRev;
                    drPartRev["RowMod"] = "A";
                    drPartRev.EndEdit();
                    oPartImpl.Update(oPartDataSet);
                    // reload the part data set
                    oPartDataSet = oPartImpl.GetByID(sPartNum);
                    // now we get the part rev
                    GetPartRevFromPart(oPartDataSet, sPartNum, sRev, out oPartRevDataSet);
                }
                else if (string.IsNullOrEmpty(sErrorMessage) == false)
                {
                    // an error occurred so we report this to the user
                    throw new Exception(sErrorMessage);
                }

                // we now have the part rev we need to add the attachment to
                if (oPartRevDataSet != null)
                {
                    // we have the rev
                    bool bApprovedRev = (bool)oPartRevDataSet["Approved"];
                    if (bApprovedRev == true)
                    {
                        //bool bPromptForPassword = false;
                        //oPartImpl.PromptForPassword(out bPromptForPassword);

                        // first we unapprove the rev and save this
                        oPartRevDataSet.BeginEdit();
                        oPartRevDataSet["RowMod"] = "U";
                        oPartRevDataSet.EndEdit();
                        oPartImpl.ChangePartRevApproved(false, true, oPartDataSet);
                        // save the part rev attachment changes
                        oPartImpl.Update(oPartDataSet);

                        // now we reload the part data set
                        oPartDataSet = oPartImpl.GetByID(sPartNum);
                        oPartRevDataSet = null;
                        GetPartRevFromPart(oPartDataSet, sPartNum, sRev, out oPartRevDataSet);
                    }

                    //////////// test to see if the part attachment aleady exists
                    //////////// if it exists then we report this as an error to the user
                    //////////bool bPartAttachmentExists = false;
                    //////////foreach (DataRow oPartAttachmentRow in oPartDataSet.PartAttch.DataSet.Tables[0].Rows)
                    //////////{
                    //////////    // check for the name
                    //////////    string sAttachmentName = (string)oPartAttachmentRow["DrawDesc"];
                    //////////    if (string.Compare(sAttachmentName, sFileName, true) == 0)
                    //////////    {
                    //////////        bPartAttachmentExists = true;
                    //////////        break;
                    //////////    }
                    //////////}
                    //////////if (bPartAttachmentExists == true)
                    //////////{
                    //////////    throw new Exception("This part rev drawing attachment already exists!");
                    //////////}

                    // move the file to its permanent location
                    string sDestinationFilenName = sDestinationFolder + sFileName;
                    if (File.Exists(sDestinationFilenName) == true)
                    {
                        // file already exists so lets delete it and readd in case there needs to be an update
                        File.Delete(sDestinationFilenName);
                    }
                    // now move the file
                    File.Move(sFullPathFileName, sDestinationFilenName);

                    // create the new part attachment
                    string sAltMethod = "";
                    string sProcessMfgID = "";
                    oPartImpl.GetNewPartRevAttch(oPartDataSet, sPartNum, sRev, sAltMethod, sProcessMfgID);
                    DataRow oPartAttachment = oPartDataSet.PartRevAttch.Rows[oPartDataSet.PartRevAttch.Rows.Count - 1];
                    oPartAttachment.BeginEdit();
                    oPartAttachment["RowMod"] = "A";
                    oPartAttachment["DrawDesc"] = sFileName;
                    oPartAttachment["FileName"] = sDestinationFilenName;
                    oPartAttachment["DocTypeID"] = "DRAW";
                    // indicate we are now finished making changes to the part attachment
                    oPartAttachment.EndEdit();
                    // save the part rev attachment changes
                    oPartImpl.Update(oPartDataSet);


                    // last step is to approve the rev regardless if it was approved previously or not
                    // we reload the part data set
                    oPartDataSet = oPartImpl.GetByID(sPartNum);
                    oPartRevDataSet = null;
                    GetPartRevFromPart(oPartDataSet, sPartNum, sRev, out oPartRevDataSet);

                    // we need to approve the rev
                    oPartRevDataSet.BeginEdit();
                    oPartRevDataSet["RowMod"] = "U";
                    oPartRevDataSet.EndEdit();
                    oPartImpl.ChangePartRevApproved(true, true, oPartDataSet);

                    // save the part rev attachment changes
                    oPartImpl.Update(oPartDataSet);

                    // we should unapprove all part revs with the exception of the lates rev that has been updated
                    oPartDataSet = oPartImpl.GetByID(sPartNum);
                    foreach (DataRow oTmpPartRev in oPartDataSet.PartRev.Rows)
                    {
                        string sTmpRev = (string)oTmpPartRev["RevisionNum"];
                        oPartRevDataSet = null;
                        if (string.Compare(sTmpRev, sRev, true) != 0)
                        {
                            GetPartRevFromPart(oPartDataSet, sPartNum, sTmpRev, out oPartRevDataSet);
                            if (oPartRevDataSet != null)
                            {
                                bApprovedRev = (bool)oPartRevDataSet["Approved"];
                                if (bApprovedRev == true)
                                {
                                    // first we unapprove the rev and save this
                                    oPartRevDataSet.BeginEdit();
                                    oPartRevDataSet["RowMod"] = "U";
                                    oPartRevDataSet.EndEdit();
                                    oPartImpl.ChangePartRevApproved(false, true, oPartDataSet);
                                    // save the part rev changes
                                    oPartImpl.Update(oPartDataSet);
                                }
                            }
                        }
                    }
                }
            }

            return;
        }

        public string GetPartRevFromPart(PartDataSet oPartDataSet, string sPartNum, string sRev, out DataRow oPartRev)
        {
            //
            // if the rev doesnt exist then create it
            // we need to unapprove all other revs
            //

            oPartRev = null;
            string sErrorMessage = string.Empty;

            if (oPartDataSet.PartRev.Rows == null)
            {
                sErrorMessage = "There are no part revs for part number '" + sPartNum + "'";
            }
            else
            {
                foreach (DataRow oTmpPartRev in oPartDataSet.PartRev.Rows)
                {
                    string sTmpRev = (string)oTmpPartRev["RevisionNum"];
                    if (string.Compare(sTmpRev, sRev, true) == 0)
                    {
                        oPartRev = oTmpPartRev;
                        break;
                    }
                }

                if (oPartRev == null)
                {
                    sErrorMessage = "The part number '" + sPartNum + "' has no revision '" + sRev + "'";
                }
            }
            return sErrorMessage;
        }

        public string GetPartRevAttachmentFromPart(PartDataSet oPartDataSet, string sPartNum, string sRev, out DataRow oPartRev)
        {
            //
            // if the rev doesnt exist then create it
            // we need to unapprove all other revs
            //
            oPartRev = null;
            string sErrorMessage = string.Empty;

            if (oPartDataSet.PartRev.Rows == null)
            {
                sErrorMessage = "There are no part revs for part number '" + sPartNum + "'";
            }
            else
            {
                foreach (DataRow oTmpPartRev in oPartDataSet.PartRev.Rows)
                {
                    string sTmpRev = (string)oTmpPartRev["RevisionNum"];
                    if (string.Compare(sTmpRev, sRev, true) == 0)
                    {
                        oPartRev = oTmpPartRev;
                        break;
                    }
                }

                if (oPartRev == null)
                {
                    sErrorMessage = "The part number '" + sPartNum + "' has no revision '" + sRev + "'";
                }
            }
            return sErrorMessage;
        }
    }
}
