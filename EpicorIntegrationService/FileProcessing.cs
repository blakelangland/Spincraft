using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Ice.Core;
using HorizonScientific;
using System.IO;

// MUST TARGET .NET 4.8
// We should reference the adapter and contract classes to both read and write to Epicor
// We start by connecting to Epicor through a session object
// The Epicor business objects (CreateImpl) are used to invoke the methods

//
// When installing the service make sure you run InstallUtil from C:\Windows\Microsoft.NET\Framework64\v4.0.30319
// You need to run the commnand console with administrator privileges
//

namespace EpicorIntegrationService
{
    public class FileProcessing
    {
        public static void ProcessFiles(string sFileName)
        {
            bool bProcessedFiles = false;
            string sFile = "";
            try
            {
                g_oSession = new Session(HSUser.SPINCRAFT_SERVICE_ACCOUNT_ID, HSUser.SPINCRAFT_SERVICE_PASSWORD, Session.LicenseType.Default, CONFIGURATION_FILE);

                if (g_oSession != null)
                {
                    if (HSUser.Initialize(AppSession) == false)
                    {
                        Service1.WriteToFile("Failed to load all users!");
                        return;
                    }

                    // we need to parse the file name to get the part number and the rev
                    string[] oFileNameParts = sFileName.Split('~');
                    sFile = System.IO.Path.GetFileName(sFileName);
                    // ensure we have all required information
                    if (oFileNameParts.Length >= 5)
                    {
                        // the part description will always be the first index
                        string sPartDescription = oFileNameParts[0];
                        // strip off the path
                        sPartDescription = System.IO.Path.GetFileName(sPartDescription);

                        // the rev number will always be the fifth index
                        string sRevNum = oFileNameParts[4];
                        // remove the extension
                        sRevNum = System.IO.Path.GetFileNameWithoutExtension(sRevNum);

                        // We will rename the file to remove the ~ and extra spaces
                        string sNewFileName = sFileName.Replace("~", " ");
                        sNewFileName = sNewFileName.Replace("  ", " ");
                        // check to make sure the file does not already exist
                        // if the file exists then we will delete the file from the upload folder
                        if (File.Exists(sNewFileName) == true)
                        {
                            File.Delete(sNewFileName);
                        }
                        // we basically rename the file to be more appropriate
                        File.Move(sFileName, sNewFileName);
                        // get the correct file name now that it has been proeprly edited
                        sFile = System.IO.Path.GetFileName(sNewFileName);

                        HSXFileAttach oFileAttach = new HSXFileAttach();
                        oFileAttach.Initialize(g_oSession, sFile, sNewFileName, sPartDescription, sRevNum, REV_ATTACHMENT_FOLDER);
                        bProcessedFiles = true;
                    }
                }
            }
            catch (Exception error)
            {
                // send an email out with error message
                HSUser oServiceAccount = HSUser.GetUserById(HSUser.SPINCRAFT_SERVICE_ACCOUNT_ID);
                List<string> oToAddresses = new List<string>();
                oToAddresses.Add(HSUser.SPINCRAFT_ROOT_USER);

                StringBuilder oStringBuilder = new StringBuilder();
                oStringBuilder.Clear();
                oStringBuilder.Append("An error occured while importing engineering drawing file: " + sFile + "\n");

                oStringBuilder.Append("Exception Message:\n");
                oStringBuilder.Append(error.Message);
                oStringBuilder.Append("\n\n");
                oStringBuilder.Append("Exception Inner Message:\n");
                oStringBuilder.Append(error.InnerException);
                oStringBuilder.Append("\n\n");
                HSEmailHelper.SendEmail(oToAddresses, "Error Executing Import Engineering Drawing " + sFile, oStringBuilder.ToString());
            }
            finally
            {
                if (bProcessedFiles == true)
                {
                    HSUser oServiceAccount = HSUser.GetUserById(HSUser.SPINCRAFT_SERVICE_ACCOUNT_ID);
                    List<string> oToAddresses = new List<string>();
                    oToAddresses.Add(HSUser.SPINCRAFT_ROOT_USER);

                    StringBuilder oStringBuilder = new StringBuilder();
                    oStringBuilder.Clear();
                    oStringBuilder.Append("The engineering drawing " + sFile + " was imported.");

                    HSEmailHelper.SendEmail(oToAddresses, "The Import Engineering Drawing Process Ran", oStringBuilder.ToString());
                }

                if (g_oSession != null)
                {
                    g_oSession.Dispose();
                    g_oSession = null;
                }
            }
        }


        public static Session AppSession
        {
            get
            {
                return g_oSession;
            }
        }

        private static Session g_oSession = null;
        //private static string CONFIGURATION_FILE = @"E:\Epicor\ERP11\LocalClients\SpincraftWIPilot\config\SpincraftWIPilot.sysconfig";
        private static string CONFIGURATION_FILE = @"E:\Epicor\ERP11\LocalClients\SpincraftWILive\config\SpincraftWILive.sysconfig";

        private static string REV_ATTACHMENT_FOLDER = @"\\10.51.0.26\Data\Epicor\Live\51515\PartRev\";

        public static string g_sConfigLocation = "";
        //public static string g_sSpincraftMAConfiguration = @"E:\Epicor\ERP11\LocalClients\SpincraftMALive\config\SpincraftMAPilotSSO.sysconfig";
        //public static string g_sSpincraftWIConfiguration = @"E:\Epicor\ERP11\LocalClients\SpincraftWILive\config\SpincraftWIPilotSSO.sysconfig";
        //public static string g_sSpincraftMAConfiguration = @"E:\Epicor\ERP11\LocalClients\SpincraftCALive\config\SpincraftMATestSSO.sysconfig";
        public static string g_sSpincraftMAConfiguration = @"E:\Epicor\ERP11\LocalClients\SpincraftMALive\config\SpincraftMALiveSSO.sysconfig";
        public static string g_sSpincraftWIConfiguration = @"E:\Epicor\ERP11\LocalClients\SpincraftWILive\config\SpincraftWILiveSSO.sysconfig";
        public static string g_sSpincraftCAConfiguration = @"E:\Epicor\ERP11\LocalClients\SpincraftCALive\config\SpincraftCALiveSSO.sysconfig";

        private static string TEMP_FILE_DIRECTORY = @"E:\Epicor\ERP11\Spincraft\TempFiles\";
        private static string UPLOAD_FILE_DIRECTORY = @"\\ETUS25AN-AP0001\UploadedFiles\";
        private static string ARCHIVE_FILE_DIRECTORY = @"\\ETUS25AN-AP0001\UploadedFiles\Archived\";
        private static string TEMPLATES_FILE_DIRECTORY = @"E:\Epicor\ERP11\Spincraft\Templates\";
    }
}
