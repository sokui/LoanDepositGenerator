using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
using System.Globalization;
using Microsoft.Office.Interop.Excel;
using System.Text.RegularExpressions;
using System.Threading;

namespace LoanDepositGenerator
{
    public partial class generatorForm : Form
    {
        public static bool running = true;
        Thread generatorThread;
        public generatorForm()
        {
            InitializeComponent();
        }

        private void inputBtn_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog fbd = new FolderBrowserDialog();
            DialogResult result = fbd.ShowDialog();
            inputPathTxtBox.Text = fbd.SelectedPath;
        }

        private void startBtn_Click(object sender, EventArgs e)
        {
            generatorThread = new Thread(new ThreadStart(run));
            generatorThread.Start();
        }

        private void run()
        {
            processRBox.Invoke(new MethodInvoker(delegate { processRBox.Clear(); }));
            //processRBox.Clear();
            running = true;
            startBtn.Invoke(new MethodInvoker(delegate { startBtn.Enabled = false; }));
            //startBtn.Enabled = false;
            //string dailyFileFolderPath = inputPathTxtBox.Text;
            string dailyFileFolderPath = null;
            inputPathTxtBox.Invoke(new MethodInvoker(delegate { dailyFileFolderPath = inputPathTxtBox.Text; }));
            try
            {
                if (!System.IO.Directory.Exists(dailyFileFolderPath))
                {
                    MessageBox.Show("The input daily file folder does not exist.");
                    return;
                }

                Logger.setLogFile(null, processRBox);
                Logger.printLog("Validating folder structure of " + dailyFileFolderPath + "......");
                if (!DailyFile.isValidDailyFolder(dailyFileFolderPath))
                {
                    Logger.printLog(dailyFileFolderPath + " does not have a good folder structure.");
                    return;
                }
                Logger.setLogFile(dailyFileFolderPath, processRBox);

                Logger.printLog("Loading dailyLoanAndDepositFiles and dailyRateFiles......");
                List<DailyLoanAndDepositFile> dailyLoanAndDepositFiles = DailyLoanAndDepositFile.getSortedFiles(Path.Combine(dailyFileFolderPath, DailyLoanAndDepositFile.dailyLoanAndDepositFolderName));
                List<DailyRateFile> dailyRateFiles = DailyRateFile.getSortedFiles(Path.Combine(dailyFileFolderPath, DailyRateFile.dailyRateFolderName));
                DateTime newBaseDate = Utility.findLastCommonDate(dailyLoanAndDepositFiles, dailyRateFiles);
                if (newBaseDate == DateTime.MinValue)
                {
                    Logger.printLog("The input daily files does not match on the time.");
                    return;
                }
                Logger.printLog("The input daily files include information up to " + newBaseDate.ToString("yyyy-MM-dd"));

                string baseFilePath = Path.Combine(dailyFileFolderPath, PortfolioFile.baseFileFolder);
                Logger.printLog("Loading the base files under " + baseFilePath + " ......");
                string[] baseFiles = Directory.GetFiles(baseFilePath);
                List<PortfolioFile> portfolioFiles = new List<PortfolioFile>();
                foreach (string bf in baseFiles)
                {
                    PortfolioFile basePortfolioFile = PortfolioFile.loadPortfolioFile(bf);
                    if (basePortfolioFile.DTime > newBaseDate)
                    {
                        Logger.printLog("The base portfolio file " + bf + " already includes all information of the input daily files. Skip!");
                        continue;
                    }
                    PortfolioFile newPortfolioFile = new PortfolioFile(basePortfolioFile.Type, newBaseDate, dailyFileFolderPath);
                    string newFileFolderPath = Path.GetDirectoryName(newPortfolioFile.getPath());
                    Directory.CreateDirectory(newFileFolderPath);
                    File.Copy(basePortfolioFile.getPath(), newPortfolioFile.getPath(), true);
                    Logger.printLog("Create the new portfolio file " + newPortfolioFile.getPath() + " based on the file " + basePortfolioFile.getPath());

                    Logger.printLog("Filtering dailyLoanAndDepositFiles and dailyRateFiles");
                    List<DailyLoanAndDepositFile> clonedDailyLoanAndDepositFiles = Utility.cloneList<DailyLoanAndDepositFile>(dailyLoanAndDepositFiles);
                    List<DailyRateFile> clonedDailyRateFiles = Utility.cloneList<DailyRateFile>(dailyRateFiles);
                    basePortfolioFile.filterDailyFiles(clonedDailyLoanAndDepositFiles, clonedDailyRateFiles);
                    for (int i = 0; i < clonedDailyLoanAndDepositFiles.Count && running; i++)
                    {
                        Logger.printLog("\n==========Adding " + clonedDailyLoanAndDepositFiles[i].Path + " to " + newPortfolioFile.getPath() + "==========");
                        newPortfolioFile.addDailyData(clonedDailyLoanAndDepositFiles[i], clonedDailyRateFiles[i]);
                        if (running)
                        {
                            Logger.printLog("Successfully added " + clonedDailyLoanAndDepositFiles[i].Path + " to " + newPortfolioFile.getPath());
                        }
                    }
                }
                if (running)
                {
                    Logger.printLog("\n=========Congratulations! All files are processed!!=========");
                    Logger.printLog("Please check all outputs and logs under folder " + Path.Combine(dailyFileFolderPath, PortfolioFile.folder));
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + ":" + System.Environment.NewLine + ex.StackTrace);
                return;
            }
            finally
            {
                startBtn.Invoke(new MethodInvoker(delegate { startBtn.Enabled = true; }));
                //startBtn.Enabled = true;
                generatorThread.Abort();
            }
        }

        private void stopBtn_Click(object sender, EventArgs e)
        {
            Logger.printLog("\n=========Stopping running=========");
            running = false;
        }
    }
}
