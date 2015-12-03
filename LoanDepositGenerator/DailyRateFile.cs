using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Globalization;
using System.IO;
using System.Text.RegularExpressions;
using Excel = Microsoft.Office.Interop.Excel;

namespace LoanDepositGenerator
{
    class DailyRateFile : DailyFile
    {
        private static readonly string[] formats = { "yyyyMMdd" };
        private static readonly char[] delimiter = { ' ', '.' };
        public static readonly string dailyRateFolderName = "daily_rate";
        public static readonly string[] effectiveCurrencies = { "USD", "HKD", "CNY", "EUR" };
        private static readonly int[] effectiveCurrencyColumns = { 3, 7, 11, 16 };
        private static readonly int rateBeginRowNumber = 7;
        private static readonly double[] basicSpans = { 1, 7, 14, 30, 60, 90, 120, 150, 365 / 2, 365 * 3 / 4, 365, 365 * 2, 365 * 3, 365 * 4, 365 * 5 };
        //public static string fileNamePattern = @"Daily rate table[1-9][0-9][0-9][0-9][0-1]?[0-9][0-3]?[0-9]\.xlsx?";


        public DailyRateFile(string path) : base(path) { }

        public override DateTime getDateTime(string path)
        {
            string[] fields = System.IO.Path.GetFileName(path).Split(delimiter);
            DateTime temp;
            if (DateTime.TryParseExact(fields[fields.Length - 2].Substring(5), formats, new CultureInfo("en-US"), DateTimeStyles.None, out temp))
            {
                return temp;
            }
            return DateTime.MinValue;
        }

        public static List<DailyRateFile> getSortedFiles(string folderPath)
        {
            List<DailyRateFile> dailyFiles = new List<DailyRateFile>();
            string[] files = Directory.GetFiles(folderPath);
            foreach (string f in files)
            {
                try
                {
                    dailyFiles.Add(new DailyRateFile(f));
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
            return @"^Daily rate table[1-9][0-9][0-9][0-9][0-1]?[0-9][0-3]?[0-9]\.xlsx?";
        }

        public static Dictionary<string, double[,]> getRateSpans(Excel.Worksheet dailyRateWorkSheet)
        {
            Dictionary<string, double[,]> rates = new Dictionary<string, double[,]>();
            for (int i = 0; i < effectiveCurrencies.Length; i++)
            {
                int count = 0;
                for (int j = 0; j < basicSpans.Length; j++)
                {
                    Excel.Range valueField = dailyRateWorkSheet.Cells[rateBeginRowNumber + j, effectiveCurrencyColumns[i]];
                    Excel.Range periodField = dailyRateWorkSheet.Cells[rateBeginRowNumber + j, effectiveCurrencyColumns[i] - 2];
                    //Excel.Range valueFieldTmp = dailyRateWorkSheet.Cells[rateBeginRowNumber + j, effectiveCurrencyColumns[i]+1];
                    //Excel.Range periodFieldTmp = dailyRateWorkSheet.Cells[rateBeginRowNumber + j, effectiveCurrencyColumns[i] - 1];
                    if (valueField.Value2 != null && periodField.Value2 != null)
                    {
                        count++;
                    }
                }

                if (count <= 0)
                {
                    continue;
                }
                rates.Add(effectiveCurrencies[i], new double[2, count]);
                int index = 0;
                for (int j = 0; j < basicSpans.Length && index < count; j++)
                {
                    Excel.Range cell = dailyRateWorkSheet.Cells[rateBeginRowNumber + j, effectiveCurrencyColumns[i]];
                    if (cell.Value2 == null)
                    {
                        continue;
                    }
                    rates[effectiveCurrencies[i]][0, index] = basicSpans[j];
                    rates[effectiveCurrencies[i]][1, index++] = (double)cell.Value2;
                }
            }
            return rates;
        }
    }
}
