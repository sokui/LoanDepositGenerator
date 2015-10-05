using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Globalization;
using System.IO;
using System.Text.RegularExpressions;

namespace LoanDepositGenerator
{
    class DailyLoanAndDepositFile : DailyFile
    {
        private static readonly string[] formats = { "yyyyMMdd", "yyyy-M-d" };
        private static readonly char[] delimiter = { ' ', '.' };
        public static readonly string dailyLoanAndDepositFolderName = "daily_deposit_loan";
        //public static string fileNamePattern = @"[1-9][0-9][0-9][0-9]-[0-1]?[0-9]-[0-3]?[0-9] Customer Deposit & Loan Breakdown\.xlsx?";

        public DailyLoanAndDepositFile(string path) : base(path) { }

        public override DateTime getDateTime(string path)
        {
            string[] fields = System.IO.Path.GetFileName(path).Split(delimiter);
            DateTime temp;
            if (DateTime.TryParseExact(fields[0], formats, new CultureInfo("en-US"), DateTimeStyles.None, out temp))
            {
                return temp;
            }
            return DateTime.MinValue;
        }

        public static List<DailyLoanAndDepositFile> getSortedFiles(string folderPath)
        {
            List<DailyLoanAndDepositFile> dailyFiles = new List<DailyLoanAndDepositFile>();
            string[] files = Directory.GetFiles(folderPath);
            foreach (string f in files)
            {
                try
                {
                    dailyFiles.Add(new DailyLoanAndDepositFile(f));
                }
                catch (InvalidFileNameException)
                {
                    continue;
                }
            }
            dailyFiles.Sort();
            return dailyFiles;
        }

        public override string getFileNamePattern()
        {
            return @"^[1-9][0-9][0-9][0-9]-[0-1]?[0-9]-[0-3]?[0-9] Customer Deposit & Loan Breakdown\.xlsx?";
        }
    }
}
