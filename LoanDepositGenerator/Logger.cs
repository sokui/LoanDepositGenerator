using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;

namespace LoanDepositGenerator
{
    class Logger
    {
        private static StreamWriter logFile;
        private static string path;
        private static RichTextBox processRBox;

        public static void setLogFile(string folderPath, RichTextBox processRBox)
        {
            Logger.processRBox = processRBox;
            closeLogFile();
            if (Directory.Exists(folderPath))
            {
                string logFileFolderPath = Path.Combine(folderPath, PortfolioFile.folder);
                path = Path.Combine(logFileFolderPath, "log-" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".txt");
                if (!Directory.Exists(logFileFolderPath))
                {
                    Directory.CreateDirectory(logFileFolderPath);
                }
                logFile = new System.IO.StreamWriter(@path);
                printLog("Successfully initialized the log file " + logFileFolderPath);
            }
        }

        public static void closeLogFile()
        {
            if (logFile != null)
            {
                logFile.Close();
                logFile = null;
            }
        }

        public static void printLog(string info)
        {
            processRBox.Invoke(new MethodInvoker(delegate { processRBox.AppendText(info + "\n"); }));
            //processRBox.AppendText(info + "\n");
            if (logFile != null)
            {
                logFile.WriteLine(info);
                logFile.Flush();
            }
        }
    }
}
