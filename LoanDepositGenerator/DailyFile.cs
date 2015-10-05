using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;

namespace LoanDepositGenerator
{
    abstract class DailyFile : IComparable<DailyFile>
    {
        public DailyFile(string path)
        {
            if (Utility.validateName(System.IO.Path.GetFileName(path), getFileNamePattern()))
            {
                Path = path;
                DTime = getDateTime(path);
            }
            else
            {
                throw new InvalidFileNameException();
            }
        }
        public string Path { get; set; }

        public DateTime DTime { get; set; }

        public int CompareTo(DailyFile file2)
        {
            return this.DTime.CompareTo(file2.DTime);
        }

        abstract public DateTime getDateTime(string path);

        abstract public string getFileNamePattern();

        public static bool isValidDailyFolder(string path)
        {
            string dailyLoanAndDepositFolderPath = System.IO.Path.Combine(path, DailyLoanAndDepositFile.dailyLoanAndDepositFolderName);
            string dailyRateFolderPath = System.IO.Path.Combine(path, DailyRateFile.dailyRateFolderName);
            string baseFileFolderPath = System.IO.Path.Combine(path, PortfolioFile.baseFileFolder);
            if (!Directory.Exists(dailyLoanAndDepositFolderPath))
            {
                Logger.printLog(dailyLoanAndDepositFolderPath + " does not exist.");
                return false;
            }

            if (!Directory.Exists(dailyRateFolderPath))
            {
                Logger.printLog(dailyRateFolderPath + " does not exist.");
                return false;
            }

            if (!Directory.Exists(baseFileFolderPath))
            {
                Logger.printLog(baseFileFolderPath + " does not exist.");
                return false;
            }

            string[] baseFiles = Directory.GetFiles(baseFileFolderPath);
            if (baseFiles.Length > 2)
            {
                Logger.printLog(baseFileFolderPath + " can only contains at most two base files. One is deposit base and the other is loan base.");
                return false;
            }
            return true;
        }
    }
}
