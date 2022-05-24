using System;
using System.Collections.Generic;
using System.Linq;
using OpenMcdf;
using System.Diagnostics;

namespace SysFileWatcher_new
{
    public class RevitFile
    {
        private string textdata;
        private Dictionary<string, string> FileInfo;
        private EventLog mylog;
        public RevitFile(string filepath, EventLog eventlog)
        {
            mylog = eventlog;
            // Extract data from the BasicFileInfo stream
            CompoundFile cf = new CompoundFile(filepath);
            CFStream foundStream = cf.RootStorage.GetStream("BasicFileInfo");
            byte[] data = foundStream.GetData();
            cf.Close();
            // Convert binary data to readable format
            textdata = System.Text.Encoding.UTF8.GetString(data);
            textdata = textdata.Replace("\0", string.Empty);
            // Parse the text data and store it in a dictionary
            string[] separators = { Environment.NewLine };
            string[] lines = textdata.Split(separators, StringSplitOptions.RemoveEmptyEntries);
            FileInfo = new Dictionary<string, string>();
            foreach (string line in lines.Skip(1))  // First line contains the same data as the other lines.
            {
                string[] parts = line.Split(':');
                if (parts.Length == 2)
                    FileInfo.Add(parts[0], parts[1].Trim());
            }
        }
        // Detailed version info
        public string GetRevitBuild()
        {
            return FileInfo["Build"];
        }
        // Might be the same as GetRevitVersionMaj
        public string GetFormat()
        {
            return FileInfo["Format"];
        }
        public string GetCentralModelPath()
        {
            return FileInfo["Central Model Path"];
        }
    }
}