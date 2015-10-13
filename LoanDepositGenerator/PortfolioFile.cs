using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;
using System.Windows.Forms;
using System.IO;
using System.Globalization;
using System.Text.RegularExpressions;

namespace LoanDepositGenerator
{
    class PortfolioFile
    {
        private static readonly string[] columnNames = { "帐号", "序号", "帐户名", "客户号", "产品码", "货币", "余额或本金", "余额或本金的等值港币", "余额或本金的科目号", "AO Name", "Branch", "System Date", "交易日", "起息日", "到期日", "当前利率%", "Interest Income per day", "FTP %", "FTP Amount per day", "计息基准天数", "Interest Income per dY", "FTP Amount in HKD Equ", "", "Interest Income in HKD Equ" };
        private static readonly string[] loanProducts = { "BLFTDS", "BLFTDD", "BLRVDS", "LNGBDBDD", "LNGBDBDM", "SYFTDD", "OWDBS", "LCPPNDD", "ADEBDA", "EX-REFI", "ADTRIW", "IM-REFI", "TT-IMPDD", "BLRVDD" };
        private static readonly string[] depositProducts = { "DDPCUR01", "DDPSAV01", "MMFCUD", "PDP" };
        private static Dictionary<string, double> allSpecialRates = initializeSpecialRates();
        private static Dictionary<string, double> EURSpecialRates = initializeEURSpecialRates();
        private static readonly string[] rate0001Products = { "DDPSAV01" };
        private static readonly int effectiveColumnNumber = 16;
        private static readonly string[] transactionDateFormats = { "yyyy/M/d", "yyyy/MM/dd" };
        public static readonly string folder = "result";
        public static readonly string baseFileFolder = "base_file";
        public static string depositFileNamePattern = @"FTP - Deposits [1-2][0-9][0-9][0-9][0-1][0-9][0-3][0-9]\.xlsx?";
        public static string loanFileNamePattern = @"loan portfolio [1-2][0-9][0-9][0-9][0-1][0-9][0-3][0-9]\.xlsx?";
        public PortfolioType Type { get; set; }
        public string Name { get; set; }
        public DateTime DTime { get; set; }
        private string path;

        private static Dictionary<string, double> initializeSpecialRates()
        { 
            Dictionary<string, double> rates = new Dictionary<string,double>();
            rates = new Dictionary<string,double>();
            rates.Add("DDPCUR01",0);
            rates.Add("PDP",0);
            return rates;
        }

        private static Dictionary<string, double> initializeEURSpecialRates()
        {
            Dictionary<string, double> rates = new Dictionary<string, double>();
            rates = new Dictionary<string, double>();
            rates.Add("DDPSAV01", 0.0001);
            return rates;
        }

        private PortfolioFile(string path)
        {
            this.path = path;
        }
        public PortfolioFile(PortfolioType type, DateTime dt, string folderPath)
        {
            if (!System.IO.Directory.Exists(folderPath))
            {
                throw new DirectoryNotFoundException();
            }
            Type = type;
            if (type == PortfolioType.LoanPortfolio)
            {
                Name = "loan portfolio " + dt.ToString("yyyyMMdd") + ".xlsx";
            }
            else
            {
                Name = "FTP - Deposits " + dt.ToString("yyyyMMdd") + ".xlsx";
            }
            DTime = dt;
            string folderName = Path.Combine(folderPath, folder);
            path = System.IO.Path.Combine(folderName, Name);
        }

        /// <summary>
        /// Validate the file name for a portfolio file.
        /// </summary>
        /// <param name="name">portfolio file name. E.g. loan portfolio 20150901.xlsx</param>
        /// <returns>
        /// 1: deposit portfolio file
        /// -1: loan portfolio file
        /// 0: invalid file name
        /// </returns>
        public static int validatePortfolioFileName(string name)
        {
            if (Utility.validateName(name, depositFileNamePattern))
            {
                return 1;
            }
            else if (Utility.validateName(name, loanFileNamePattern))
            {
                return -1;
            }

            return 0;
        }

        public static DateTime getDateTimeFromPortfolioFileName(string name)
        {
            if (string.IsNullOrWhiteSpace(name))
            {
                throw new InvalidFileNameException();
            }
            string[] formats = { "yyyyMMdd" };
            char[] separators = { ' ', '.' };
            string[] fields = name.Split(separators);
            DateTime dtime;
            DateTime.TryParseExact(fields[fields.Length - 2], formats, new CultureInfo("en-US"), DateTimeStyles.None, out dtime);
            if (dtime == DateTime.MinValue)
            {
                throw new InvalidFileNameException();
            }
            return dtime;
        }

        public static PortfolioFile loadPortfolioFile(string path)
        {
            if (!File.Exists(path))
            {
                throw new FileNotFoundException();
            }

            PortfolioFile pf = new PortfolioFile(path);
            pf.Name = Path.GetFileName(path);
            int portfolioType = validatePortfolioFileName(pf.Name);
            switch (portfolioType)
            {
                case 1:
                    pf.Type = PortfolioType.DepositPortfolio;
                    break;
                case -1:
                    pf.Type = PortfolioType.LoanPortfolio;
                    break;
                default:
                    throw new InvalidFileNameException();
            }

            pf.DTime = getDateTimeFromPortfolioFileName(pf.Name);
            return pf;
        }

        public string getPath()
        {
            return path;
        }

        public void createPortfolioXlFile(string folderPath)
        {
            if (!System.IO.Directory.Exists(folderPath))
            {
                throw new DirectoryNotFoundException();
            }

            Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();

            if (xlApp == null)
            {
                MessageBox.Show("Excel is not properly installed!!");
                return;
            }
            object misValue = System.Reflection.Missing.Value;

            Excel.Workbook xlWorkBook = xlApp.Workbooks.Add(misValue);
            Excel.Worksheet xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
            int count = 1;
            foreach (string colName in columnNames)
            {
                xlWorkSheet.Cells[1, count++] = colName;
            }
            Excel.Range row1 = xlWorkSheet.UsedRange;
            row1.Columns.AutoFit();
            try
            {
                xlWorkBook.SaveAs(path);
                xlWorkBook.Close(false, misValue, misValue);
                xlApp.Quit();
            }
            finally
            {
                Utility.releaseObject(xlWorkSheet);
                Utility.releaseObject(xlWorkBook);
                Utility.releaseObject(xlApp);
            }
        }
        public void addDailyData(DailyLoanAndDepositFile dailyLoanAndDepositFile, DailyRateFile dailyRateFile)
        {
            Excel.Application xlApp = new Excel.Application();

            Excel.Workbooks workbooksTmp = xlApp.Workbooks;
            Excel.Workbook dailyLoanAndDepositWorkBook = workbooksTmp.Open(dailyLoanAndDepositFile.Path);
            Excel.Workbook dailyRateWorkBook = workbooksTmp.Open(dailyRateFile.Path);
            Excel.Workbook portfolioWorkBook = workbooksTmp.Open(path);

            Excel.Sheets dailyLoanAndDepositWorkSheetTmp = dailyLoanAndDepositWorkBook.Worksheets;
            Excel.Sheets dailyRateWorkSheetsTmp = dailyRateWorkBook.Worksheets;
            Excel.Sheets portfolioSheetsTmp = portfolioWorkBook.Worksheets;
            Excel.Worksheet dailyLoanAndDepositWorkSheet = (Excel.Worksheet)dailyLoanAndDepositWorkSheetTmp.get_Item(1);
            Excel.Worksheet dailyRateWorkSheet = Type == PortfolioType.DepositPortfolio ? (Excel.Worksheet)dailyRateWorkSheetsTmp.get_Item(1) : (Excel.Worksheet)dailyRateWorkSheetsTmp.get_Item(2);
            Excel.Worksheet portfolioWorkSheet = (Excel.Worksheet)portfolioSheetsTmp.get_Item(1);

            Excel.Range dailyLoanAndDepositRange = dailyLoanAndDepositWorkSheet.UsedRange;
            Excel.Range dailyRateRange = dailyRateWorkSheet.UsedRange;
            Excel.Range portfolioRange = portfolioWorkSheet.UsedRange;
            try
            {
                Dictionary<string, double[,]> rates = DailyRateFile.getRateSpans(dailyRateWorkSheet);
                int portfolioRowCounter = portfolioRange.Rows.Count + 1;
                for (int i = 2; i < dailyLoanAndDepositRange.Rows.Count && generatorForm.running; i++)
                {
                    Logger.printLog("Processing Row " + i + " in Sheet " + dailyLoanAndDepositWorkSheet.Name + " of File " + dailyLoanAndDepositWorkBook.FullName);
                    Excel.Range currentRow = dailyLoanAndDepositWorkSheet.get_Range(i + ":" + i);
                    if (!isValidRow(currentRow))
                    {
                        Logger.printLog("Row " + i + " in Sheet " + dailyLoanAndDepositWorkSheet.Name + " of File " + dailyLoanAndDepositWorkBook.FullName + " is not valid");
                        continue;
                    }
                    if (getTransactionType(currentRow) != Type)
                    {
                        continue;
                    }

                    Excel.Range portfolioRow = portfolioWorkSheet.get_Range(portfolioRowCounter + ":" + portfolioRowCounter);
                    copyRow(currentRow, portfolioRow);
                    writeFTPRate(rates, portfolioWorkSheet, portfolioRow, portfolioRowCounter, dailyRateFile.DTime);
                    portfolioRowCounter++;
                }
            }
            finally
            {
                portfolioWorkBook.Save();
                portfolioWorkBook.Close(false, null, null);
                dailyLoanAndDepositWorkBook.Close(false, null, null);
                dailyRateWorkBook.Close(false, null, null);

                xlApp.Quit();
                Utility.releaseObject(dailyLoanAndDepositWorkSheet);
                Utility.releaseObject(dailyRateWorkSheet);
                Utility.releaseObject(portfolioWorkSheet);
                Utility.releaseObject(dailyLoanAndDepositWorkSheetTmp);
                Utility.releaseObject(dailyRateWorkSheetsTmp);
                Utility.releaseObject(portfolioSheetsTmp);
                Utility.releaseObject(dailyLoanAndDepositWorkBook);
                Utility.releaseObject(dailyRateWorkBook);
                Utility.releaseObject(portfolioWorkBook);
                Utility.releaseObject(workbooksTmp);
                Utility.releaseObject(xlApp);
            }
        }

        public void filterDailyFiles(List<DailyLoanAndDepositFile> sortedDailyLoanAndDepositFiles, List<DailyRateFile> sortedDailyRateFiles)
        {
            for (int i = 0; i < sortedDailyLoanAndDepositFiles.Count; i++)
            {
                if (sortedDailyLoanAndDepositFiles[i].DTime > DTime)
                {
                    break;
                }
                sortedDailyLoanAndDepositFiles.RemoveAt(i--);
            }

            for (int i = 0; i < sortedDailyRateFiles.Count; i++)
            {
                if (sortedDailyRateFiles[i].DTime > DTime)
                {
                    break;
                }
                sortedDailyRateFiles.RemoveAt(i--);
            }

            for (int i = 0, j = 0; i < sortedDailyLoanAndDepositFiles.Count && j < sortedDailyRateFiles.Count; i++, j++)
            {
                if (sortedDailyLoanAndDepositFiles[i].DTime < sortedDailyRateFiles[j].DTime)
                {
                    sortedDailyLoanAndDepositFiles.RemoveAt(i--);
                    j--;
                }
                else if (sortedDailyLoanAndDepositFiles[i].DTime > sortedDailyRateFiles[j].DTime)
                {
                    sortedDailyRateFiles.RemoveAt(j--);
                    i--;
                }
            }
        }

        private bool isValidRow(Excel.Range row)
        {
            for (int i = 1; i < effectiveColumnNumber; i++)
            {
                Excel.Range cell = (row.Cells[1, i] as Excel.Range);
                if (cell == null || cell.Value2 == null || string.IsNullOrWhiteSpace(cell.Value2.ToString()))
                {
                    return false;
                }
            }

            DateTime systemDate = Utility.parseDate((row.Cells[1, 12] as Excel.Range), transactionDateFormats);
            if (systemDate == DateTime.MinValue)
            {
                return false;
            }

            if (!loanProducts.Contains<string>((string)(row.Cells[1, 5] as Excel.Range).Value2)
                && !depositProducts.Contains<string>((string)(row.Cells[1, 5] as Excel.Range).Value2))
            {
                return false;
            }
            return true;
        }

        private void copyRow(Excel.Range sourceRow, Excel.Range destRow)
        {
            for (int i = 1; i <= effectiveColumnNumber; i++)
            {
                destRow.Cells[1, i] = sourceRow.Cells[1, i];
                (destRow.Cells[1, i] as Excel.Range).NumberFormat = (sourceRow.Cells[1, i] as Excel.Range).NumberFormat;
            }
        }

        private void writeFTPRate(Dictionary<string, double[,]> rates, Excel.Worksheet portfolioWorkSheet, Excel.Range row, int rowNumber, DateTime dailyFileDateTime)
        {
            // handle special products rate
            string productName = (string)(row.Cells[1, 5] as Excel.Range).Value2;
            if (allSpecialRates.ContainsKey(productName))
            {
                (row.Cells[1, 18] as Excel.Range).Value2 = allSpecialRates[productName];
                return;
            }
            string currencyType = (string)(row.Cells[1, 6] as Excel.Range).Value2;
            if (currencyType.Equals(DailyRateFile.effectiveCurrencies[3]) && EURSpecialRates.ContainsKey(productName))
            {
                (row.Cells[1, 18] as Excel.Range).Value2 = EURSpecialRates[productName];
                return;
            }

            double[,] rateTable = null;
            if (rates.ContainsKey(currencyType))
            {
                rateTable = rates[currencyType];
            }
            DateTime systemDate, transactionDate, startDate, endDate;
            systemDate = Utility.parseDate((row.Cells[1, 12] as Excel.Range), transactionDateFormats);
            transactionDate = Utility.parseDate((row.Cells[1, 13] as Excel.Range), transactionDateFormats);
            startDate = Utility.parseDate((row.Cells[1, 14] as Excel.Range), transactionDateFormats);
            endDate = Utility.parseDate((row.Cells[1, 15] as Excel.Range), transactionDateFormats);

            if (startDate == DateTime.MinValue || endDate == DateTime.MinValue)
            {
                if ((transactionDate == DateTime.MinValue && systemDate == DateTime.MinValue) || (transactionDate == DateTime.MinValue && systemDate > dailyFileDateTime) || transactionDate > dailyFileDateTime)
                {
                    return;
                }
                else if (transactionDate == dailyFileDateTime || (transactionDate == DateTime.MinValue && systemDate == dailyFileDateTime))
                {
                    if (rateTable != null && !currencyType.Equals(DailyRateFile.effectiveCurrencies[3]))
                    {
                        (row.Cells[1, 18] as Excel.Range).Value2 = rateTable[1, 0];
                    }
                    return;
                }
                else
                {
                    Excel.Range previousRow = findRow(portfolioWorkSheet, row, rowNumber);
                    if (previousRow != null)
                    {
                        (row.Cells[1, 18] as Excel.Range).Value2 = (previousRow.Cells[1, 18] as Excel.Range).Value2;
                    }
                    return;
                }
            }

            if (rateTable == null)
            {
                return;
            }
            if (startDate < dailyFileDateTime)
            {
                Excel.Range previousRow = findRow(portfolioWorkSheet, row, rowNumber);
                if (previousRow != null)
                {
                    (row.Cells[1, 18] as Excel.Range).Value2 = (previousRow.Cells[1, 18] as Excel.Range).Value2;
                }
                return;
            }
            if (startDate > dailyFileDateTime)
            {
                return;
            }
            double daySpan = (endDate - startDate).TotalDays;

            if (daySpan < rateTable[0, 0])
            {
                if (!currencyType.Equals(DailyRateFile.effectiveCurrencies[3]))
                {
                    (row.Cells[1, 18] as Excel.Range).Value2 = rateTable[1, 0];
                }
                return;
            }
            for (int i = 0; i < rateTable.GetLength(1); i++)
            {
                if (daySpan == rateTable[0, i] || (daySpan > rateTable[0, i] && i == rateTable.GetLength(1) - 1))
                {
                    (row.Cells[1, 18] as Excel.Range).Value2 = rateTable[1, i];
                    return;
                }
                if (daySpan > rateTable[0, i] && daySpan < rateTable[0, i + 1])
                {
                    (row.Cells[1, 18] as Excel.Range).Value2 = daySpan >= (rateTable[0, i] + rateTable[0, i + 1]) / 2 ? rateTable[1, i + 1] : rateTable[1, i];
                    return;
                }
            }
            return;
        }

        private Excel.Range findRow(Excel.Worksheet portfolioWorkSheet, Excel.Range row, int rowNumber)
        {
            Excel.Range currentFind = null;
            Excel.Range firstFind = null;
            Excel.Range previousColumn1 = portfolioWorkSheet.get_Range("A1:A" + (rowNumber - 1));
            currentFind = previousColumn1.Find(row.Cells[1, 1]);
            while (currentFind != null)
            {
                Excel.Range currentRow = currentFind.EntireRow;
                if (equalRows(currentRow, row))
                {
                    return currentRow;
                }
                if (firstFind == null)
                {
                    firstFind = currentFind;
                }
                else if (currentFind.get_Address(Excel.XlReferenceStyle.xlA1)
                      == firstFind.get_Address(Excel.XlReferenceStyle.xlA1))
                {
                    break;
                }
                currentFind = previousColumn1.FindNext(currentFind); 
            }
            return null;
            //for (int i = portfolioWorkSheet.UsedRange.Rows.Count - 1; i > 1; i--)
            //{
            //    Excel.Range currentRow = portfolioWorkSheet.Rows[i];
            //    //Excel.Range currentRow = (portfolioWorkSheet.get_Range(portfolioWorkSheet.Cells[i, 1], portfolioWorkSheet.Cells[i, baseEffectiveColumnNumber]) as Excel.Range);
            //    if (equalRows(currentRow, row))
            //    {
            //        return currentRow;
            //    }
            //}
            //return null;
        }

        private bool equalRows(Excel.Range row1, Excel.Range row2)
        {
            int[] equalityIndex = { 1, 3, 4, 5, 6, 13, 14, 15 };
            foreach (int i in equalityIndex)
            {
                if ((row1.Cells[1, i] as Excel.Range).Value2 != (row2.Cells[1, i] as Excel.Range).Value2)
                {
                    return false;
                }
            }
            return true;
        }

        private PortfolioType getTransactionType(Excel.Range row)
        {
            if (loanProducts.Contains<string>((string)(row.Cells[1, 5] as Excel.Range).Value2))
            {
                return PortfolioType.LoanPortfolio;
            }
            return PortfolioType.DepositPortfolio;
        }
    }
}
