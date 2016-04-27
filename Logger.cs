using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Configuration;

namespace Utility.Log
{
    class Logger
    {
        public enum LOGLEVEL {INFO, ERROR, WARN};
        private const int LOG_RETAINER_LIMIT = 30;

        private static string fullFilePath = String.Empty;

        public static void ConstructLogFile(string dir, string filename)
        {
            fullFilePath = dir + filename + string.Format("-{0:yyyy-MM-dd_hh-mm-ss-tt}.txt", DateTime.Now);

            if (!File.Exists(fullFilePath))
            {
                File.Create(fullFilePath).Close();
            }

            CleanseDirectory(dir);   
        }

        public static void Log(string logEntry, int loglvl = 0)
        {
            string logRecord = string.Format("{0:MM/dd/yyyy hh:mm:ss tt}", DateTime.Now) + "\t" +
                               Enum.GetName(typeof(LOGLEVEL), loglvl) + "\t" +
                               logEntry + "\n";
            // Write log entry to console
            Console.WriteLine(logRecord);
            // Write log entry to log file
            File.AppendAllText(fullFilePath, logRecord);
        }

        public static void CleanseDirectory(string logDirectory) 
        {
            foreach(string logfile in Directory.GetFiles(Path.GetFullPath(logDirectory))) 
            {
                if ((DateTime.Now - Directory.GetLastWriteTime(logfile)).Days >= LOG_RETAINER_LIMIT)
                {
                    File.Delete(logfile);
                }
            }

        }

        public static string GetFullFilePath() { return fullFilePath; }
    }
}
