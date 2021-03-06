using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.ServiceProcess;
using System.IO;
using System.Data.SqlClient;
using System.Runtime.InteropServices;
using System.Timers;

namespace SysFileWatcher_new
{
    public partial class FileWatcherService : ServiceBase
    {
        private static string updateFullPath = "UPDATE FolderLogs SET latestversion = '0' WHERE folderpath = @fullPath";
        private static string updateOldPath = "UPDATE FolderLogs SET latestversion = '0' WHERE folderpath = @oldFullPath";
        private static string dbTablesRename = "folderpath, oldfolderpath, folderchange, datetime, latestversion";
        private static string dbTables = "folderpath, folderchange, datetime, latestversion, revitversion";
        private int eventId = 1;

        [DllImport("advapi32.dll", SetLastError = true)]
        private static extern bool SetServiceStatus(IntPtr handle, ref ServiceStatus serviceStatus);

        public FileWatcherService()
        {
            InitializeComponent();
            eventLog1 = new EventLog();
            if (!EventLog.SourceExists("LogSource"))
            {
                EventLog.CreateEventSource(
                    "LogSource", "FileWatcherLog");
            }
            eventLog1.Source = "LogSource";
            eventLog1.Log = "FileWatcherLog";
        }

        public void OnTimer(object sender, ElapsedEventArgs args)
        {
            // TODO: Insert monitoring activities here.
            eventLog1.WriteEntry("Monitoring the System", EventLogEntryType.Information, eventId++);
        }

        public enum ServiceState
        {
            SERVICE_STOPPED = 0x00000001,
            SERVICE_START_PENDING = 0x00000002,
            SERVICE_STOP_PENDING = 0x00000003,
            SERVICE_RUNNING = 0x00000004,
            SERVICE_CONTINUE_PENDING = 0x00000005,
            SERVICE_PAUSE_PENDING = 0x00000006,
            SERVICE_PAUSED = 0x00000007,
        }

        [StructLayout(LayoutKind.Sequential)]
        public struct ServiceStatus
        {
            public int dwServiceType;
            public ServiceState dwCurrentState;
            public int dwControlsAccepted;
            public int dwWin32ExitCode;
            public int dwServiceSpecificExitCode;
            public int dwCheckPoint;
            public int dwWaitHint;
        };

        private void OnCreated(object sender, FileSystemEventArgs e)
        {
            eventLog1.WriteEntry("revit file created, before");
            RevitFile revitFile = new RevitFile($"{e.FullPath}", eventLog1);
            string revitVersion = revitFile.GetFormat();
            MakeDataBaseQuery($"INSERT INTO FolderLogs ({dbTables}) VALUES (@fullPath, 1, @datetime, 1, {revitVersion})", e);
            eventLog1.WriteEntry("revit file created, after");
        }

        private void OnChanged(object sender, FileSystemEventArgs e)
        {
            eventLog1.WriteEntry("revit file changed, before");
            RevitFile revitFile = new RevitFile($"{e.FullPath}", eventLog1);
            string revitVersion = revitFile.GetFormat();
            MakeDataBaseQuery($"{updateFullPath} INSERT INTO FolderLogs ({dbTables}) VALUES (@fullPath, 2, @datetime, 1, {revitVersion})", e);
            eventLog1.WriteEntry("revit file changed, after");
        }

        private static void OnRenamed(object sender, RenamedEventArgs e)
        {
            var q = $"{updateOldPath} INSERT INTO FolderLogs ({dbTablesRename}) VALUES (@fullPath, @oldFullPath, 3, @datetime, 1)";
            using (SqlConnection sqlCon = new SqlConnection("Server=hfb-sql02;Integrated security=SSPI;database=SystemFileWatcher"))
            {
                SqlCommand sqlda = new SqlCommand(q, sqlCon);
                sqlda.Parameters.AddWithValue("@oldFullPath", e.OldFullPath);
                sqlda.Parameters.AddWithValue("@fullPath", e.FullPath);
                sqlda.Parameters.AddWithValue("@datetime", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff"));
                try
                {
                    sqlCon.Open();
                    sqlda.ExecuteNonQuery();
                }
                catch (Exception ex)
                {
                    sqlCon.Close();
                    string exStr = ex.ToString();
                    string errorQ = $"INSERT INTO ErrorLog (ErrorText, DateTime) VALUES ('{exStr}', @datetime)";
                    SqlCommand errorSqlda = new SqlCommand(errorQ, sqlCon);
                    sqlCon.Open();
                    errorSqlda.ExecuteNonQuery();

                }
            }
        }

        private void OnDeleted(object sender, FileSystemEventArgs e)
        {
            MakeDataBaseQuery($"{updateFullPath} INSERT INTO FolderLogs ({dbTables}) VALUES (@fullPath, 4, @datetime, 0, NULL)", e);
        }

        private void MakeDataBaseQuery(string q, FileSystemEventArgs e)
        {
            eventLog1.WriteEntry("in makedatabasequery");
            using (SqlConnection sqlCon = new SqlConnection("Server=hfb-sql02;Integrated security=SSPI;database=SystemFileWatcher"))
            {
                SqlCommand sqlda = new SqlCommand (q, sqlCon);
                sqlda.Parameters.AddWithValue("@fullPath", e.FullPath);
                sqlda.Parameters.AddWithValue("@datetime", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff"));
                eventLog1.WriteEntry("just before try.");
                try
                {
                    sqlCon.Open();
                    sqlda.ExecuteNonQuery();
                }
                catch (Exception ex)
                {
                    eventLog1.WriteEntry(ex.Message);
                    try
                    {
                        sqlCon.Close();
                    }
                    catch (Exception exe)
                    {
                        eventLog1.WriteEntry(exe.Message);
                        throw exe;
                    }
                    eventLog1.WriteEntry("after close");
                    string exStr = ex.ToString();
                    string errorQ = $"INSERT INTO ErrorLog (ErrorText, DateTime) VALUES ('{exStr}', @datetime)";
                    SqlCommand errorSqlda = new SqlCommand(errorQ, sqlCon);
                    sqlCon.Open();
                    errorSqlda.ExecuteNonQuery();
                }
            }
        }

        protected override void OnStart(string[] args)
        {
            eventLog1.WriteEntry("In OnStart.");

            // Update the service state to Start Pending.
            ServiceStatus serviceStatus = new ServiceStatus();
            serviceStatus.dwCurrentState = ServiceState.SERVICE_START_PENDING;
            serviceStatus.dwWaitHint = 100000;
            SetServiceStatus(this.ServiceHandle, ref serviceStatus);


            // Set up a timer.
            Timer timer = new Timer();
            timer.Interval = 300000; // 5 minutes
            timer.Elapsed += new ElapsedEventHandler(this.OnTimer);
            timer.Start();
            #region Code starts here

            var directories = Directory.GetDirectories(@"\\hfb-fs01\projecten");
            foreach (string projectfolder in directories)
            {
                try
                {
                    FileSystemWatcher watcher = new FileSystemWatcher (projectfolder + "\\99 HFB Labs");

                    watcher.NotifyFilter = NotifyFilters.Attributes
                                         | NotifyFilters.CreationTime
                                         | NotifyFilters.DirectoryName
                                         | NotifyFilters.FileName
                                         | NotifyFilters.LastAccess
                                         | NotifyFilters.LastWrite
                                         | NotifyFilters.Security
                                         | NotifyFilters.Size;

                    watcher.Changed += OnChanged;
                    watcher.Created += OnCreated;
                    watcher.Deleted += OnDeleted;
                    watcher.Renamed += OnRenamed;

                    watcher.Filter = "*.rvt";
                    watcher.IncludeSubdirectories = true;
                    watcher.EnableRaisingEvents = true;

                }
                catch (ArgumentException)
                {
                    // For folders that don't have the 99 HFB Labs subfolder
                    // Empty catch to skip those folders
                }
            }

            #endregion Code ends here

            // Update the service state to Running.
            serviceStatus.dwCurrentState = ServiceState.SERVICE_RUNNING;
            SetServiceStatus(this.ServiceHandle, ref serviceStatus);
        }

        protected override void OnStop()
        {
            eventLog1.WriteEntry("In OnStop.");
        }

        private void eventLog1_EntryWritten(object sender, EntryWrittenEventArgs e)
        {

        }
    }
}
