using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.Excel;
using System.Windows.Forms;
using System.Globalization;
using System.Text.RegularExpressions;

namespace LoanDepositGenerator
{
    class Utility
    {
        public static bool validateName(string name, string pat)
        {
            Regex r = new Regex(pat);
            Match m = r.Match(name);
            if (m.Success && m.Groups.Count == 1)
            {
                return true;
            }
            return false;
        }

        public static DateTime parseDate(Range cell, string[] formats)
        {
            if (cell == null || cell.Count != 1)
            {
                throw new ArgumentException();
            }
            var type = cell.Value2.GetType();
            if (type.ToString().Equals("System.Double"))
            {
                return DateTime.FromOADate(cell.Value2);
            }
            else
            {
                DateTime dt;
                DateTime.TryParseExact(cell.Value2.ToString(), formats, new CultureInfo("en-US"), DateTimeStyles.None, out dt);
                return dt;
            }
        }

        public static void releaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
                MessageBox.Show("Unable to release the Object " + ex.ToString());
            }
            finally
            {
                GC.Collect();
            }
        }

        public static List<T> cloneList<T>(List<T> l)
        {
            if (l == null)
            {
                return null;
            }
            List<T> clonedList = new List<T>();
            foreach (T e in l)
            {
                clonedList.Add(e);
            }
            return clonedList;
        }

        public static DateTime findLastCommonDate<T, K>(List<T> sortedFileList1, List<K> sortedFileList2)
            where T : DailyFile
            where K : DailyFile
        {
            if (sortedFileList1 == null || sortedFileList2 == null)
            {
                return DateTime.MinValue;
            }
            for (int i = sortedFileList1.Count - 1, j = sortedFileList2.Count - 1; i >= 0 && j >= 0; i--, j--)
            {
                if (sortedFileList1[i].DTime == sortedFileList2[j].DTime)
                {
                    return sortedFileList1[i].DTime;
                }
                if (sortedFileList1[i].DTime > sortedFileList2[j].DTime)
                {
                    j++;
                }
                else
                {
                    i++;
                }
            }
            return DateTime.MinValue;
        }
    }
}
