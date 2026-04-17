using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Linq;
using System.ServiceProcess;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Timers;
using System.Threading;


namespace EpicorIntegrationService
{
    public partial class Service1 : ServiceBase
    {
        public Service1()
        {
            InitializeComponent();
        }

        public void OnDebug()
        {
            OnStart(null);
        }

        protected override void OnStart(string[] args)
        {
            WriteToFile("Service was started");

            // create a new instance of the file watcher
            DirectoryInfo oDirectory = new DirectoryInfo(ENGINEERING_FILE_DIRECTORY);
            // Checks whether the folder is enabled and
            // also the directory is a valid location
            if (oDirectory.Exists)
            {
                // Sets the folder location
                m_oFileSystemWatcher.Path = oDirectory.FullName;

                m_oFileSystemWatcher.Created += m_oFileSystemWatcher_Created;
                //m_oFileSystemWatcher.Renamed += m_oFileSystemWatcher_Renamed;
                //m_oFileSystemWatcher.Changed += m_oFileSystemWatcher_Changed;
                m_oFileSystemWatcher.IncludeSubdirectories = false;

                // Begin watching
                m_oFileSystemWatcher.EnableRaisingEvents = true;

                // Record a log entry into Windows Event Log
                WriteToFile(String.Format("Starting to monitor files with extension ({0}) in the folder ({1})", m_oFileSystemWatcher.Filter, m_oFileSystemWatcher.Path));
            }
        }

        protected override void OnStop()
        {
            WriteToFile("Service is stopped at " + DateTime.Now);
            if (m_oFileSystemWatcher != null)
            {
                m_oFileSystemWatcher.EnableRaisingEvents = false;
                m_oFileSystemWatcher.Created -= m_oFileSystemWatcher_Created;
                //m_oFileSystemWatcher.Renamed -= m_oFileSystemWatcher_Renamed;
                //m_oFileSystemWatcher.Changed -= m_oFileSystemWatcher_Changed;

                m_oFileSystemWatcher.Dispose();
            }
        }

        void m_oFileSystemWatcher_Changed(object sender, FileSystemEventArgs e)
        {
            Thread.Sleep(5000);
            WriteToFile("File " + e.FullPath + " was reanmed and will be processed.");
            FileProcessing.ProcessFiles(e.FullPath);
        }

        void m_oFileSystemWatcher_Renamed(object sender, RenamedEventArgs e)
        {
            Thread.Sleep(5000);
            WriteToFile("File " + e.FullPath + " was reanmed and will be processed.");
            FileProcessing.ProcessFiles(e.FullPath);
        }

        void m_oFileSystemWatcher_Created(object sender, FileSystemEventArgs e)
        {
            Thread.Sleep(5000);
            WriteToFile("File " + e.FullPath + " was created and will be processed.");
            FileProcessing.ProcessFiles(e.FullPath);
        }

        public static void WriteToFile(string Message)
        {
            string path = AppDomain.CurrentDomain.BaseDirectory + "\\Logs";
            if (!Directory.Exists(path))
            {
                Directory.CreateDirectory(path);
            }
            string filepath = AppDomain.CurrentDomain.BaseDirectory + "\\Logs\\ServiceLog_" + DateTime.Now.Date.ToShortDateString().Replace('/', '_') + ".txt";
            if (!File.Exists(filepath))
            {
                // Create a file to write to.   
                using (StreamWriter sw = File.CreateText(filepath))
                {
                    sw.WriteLine(DateTime.Now.ToString() + ":" + Message);
                }
            }
            else
            {
                using (StreamWriter sw = File.AppendText(filepath))
                {
                    sw.WriteLine(DateTime.Now.ToString() + ":" + Message);
                }
            }
        }

        private static string ENGINEERING_FILE_DIRECTORY = @"C:\Epicor\SpincraftUK\UploadedFiles\";
        private FileSystemWatcher m_oFileSystemWatcher = new FileSystemWatcher();
    }
}
