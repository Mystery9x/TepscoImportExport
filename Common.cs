using System.Windows.Forms;
using MessageBox = System.Windows.Forms.MessageBox;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace TepscoImportExport
{
    internal class Common
    {
        /// ================================================================================
        /// <summary>
        /// Show information to user
        /// </summary>
        /// <param name="message"></param>
        /// <param name="title"></param>
        /// ================================================================================
        public static void ShowInfor(string message, string title = "インフォメーション")
        {
            MessageBox.Show(message, title, MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        /// ================================================================================
        /// <summary>
        /// Show warning to user
        /// </summary>
        /// <param name="message"></param>
        /// <param name="title"></param>
        /// ================================================================================
        public static void ShowWarning(string message, string title = "警告")
        {
            MessageBox.Show(message, title, MessageBoxButtons.OK, MessageBoxIcon.Warning);
        }

        /// ================================================================================
        /// <summary>Write new excel</summary>
        /// <param name="_listItem"></param>
        /// <param name="_listItemLinks"></param>
        /// <param name="path"></param>
        /// <returns></returns>
        /// ================================================================================
        public bool WriteNewExcel(string path, System.Windows.Forms.ProgressBar progressBar, List<List<string>> lstTab, List<List<string>> lstLink, bool check)
        {
            try
            {
                Excel.Application xlApp = null;
                Excel.Workbook xlWorkbook = null;
                Excel._Worksheet xlWorksheet = null;
                Excel._Worksheet xlWorksheet1 = null;
                Excel.Range xlRange = null;
                Excel.Range xlRange1 = null;

                try
                {
                    progressBar.Maximum = lstTab.Count + lstLink.Count;

                    xlApp = new Microsoft.Office.Interop.Excel.Application();
                    xlWorkbook = xlApp.Workbooks.Open(path);
                    xlWorksheet = xlWorkbook.Sheets[1];
                    xlWorksheet1 = xlWorkbook.Sheets[2];

                    Excel.Range c1 = (Excel.Range)xlWorksheet.Cells[2, 1];
                    Excel.Range c2 = (Excel.Range)xlWorksheet.Cells[lstTab.Count + 1, 65];

                    Excel.Range c3 = (Excel.Range)xlWorksheet1.Cells[2, 1];
                    Excel.Range c4 = (Excel.Range)xlWorksheet1.Cells[lstLink.Count + 1, 34];

                    xlRange = xlWorksheet.get_Range(c1, c2);
                    xlRange1 = xlWorksheet1.get_Range(c3, c4);

                    int rowCount1 = lstLink.Count;
                    int rowCount = lstTab.Count;

                    object[,] arr = new object[rowCount, 65];
                    object[,] arr1 = new object[rowCount1, 34];

                    int count = 1;
                    if (rowCount != 0)
                    {
                        List<string> lst = new List<string>();
                        for (int r = 0; r < rowCount; r++)
                        {
                            int colCount = lstTab[r].Count;
                            progressBar.PerformStep();
                            if (lstTab[r].Count != 0)
                            {
                                if (check == true/* && r > 0*/)
                                {
                                    //if (lstTab[r][1] == arr[r - 1, 1])
                                    //    continue;
                                    if (lst.Contains(lstTab[r][1]) == true)
                                        continue;
                                }

                                for (int c = 0; c < colCount; c++)
                                {
                                    if (c == 65)
                                        break;
                                    if (c == 0)
                                    {
                                        arr[r, c] = count;
                                    }
                                    arr[r, c] = lstTab[r][c];
                                }

                                lst.Add(lstTab[r][1]);
                                count++;
                            }
                        }
                    }
                    count = 1;
                    if (rowCount1 != 0)
                    {
                        List<string> lst = new List<string>();
                        for (int r = 0; r < rowCount1; r++)
                        {
                            int colCount1 = lstLink[r].Count;
                            progressBar.PerformStep();
                            if (lstLink[r].Count != 0)
                            {
                                if (check == true)
                                {
                                    if (check == true)
                                    {
                                        //if (lstLink[r][1] == arr[r - 1, 1])
                                        //    continue;
                                        if (lst.Contains(lstLink[r][1]) == true)
                                            continue;
                                    }
                                }

                                for (int c = 0; c < colCount1; c++)
                                {
                                    if (c == 34)
                                        break;
                                    if (c == 0)
                                    {
                                        arr[r, c] = count;
                                    }
                                    arr1[r, c] = lstLink[r][c];
                                }
                                count++;
                            }
                        }
                    }

                    xlRange.Value = arr;
                    xlRange1.Value = arr1;

                    return true;
                }
                catch (Exception ex)
                {
                    string mess = ex.Message;
                    return false;
                }
                finally
                {
                    GC.Collect();
                    GC.WaitForPendingFinalizers();

                    if (xlRange != null)
                        Marshal.ReleaseComObject(xlRange);

                    if (xlWorksheet != null)
                        Marshal.ReleaseComObject(xlWorksheet);

                    //close and release
                    if (xlWorkbook != null)
                    {
                        xlWorkbook.Save();

                        xlWorkbook.Close();
                        Marshal.ReleaseComObject(xlWorkbook);
                    }

                    if (xlApp != null)
                    {
                        //quit and release
                        xlApp.Quit();
                        Marshal.ReleaseComObject(xlApp);
                    }
                }
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message.ToString(), "Err Message", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
                return false;
            }
        }

        /// ================================================================================
        /// <summary>Get Value </summary>
        /// <param name="values"></param>
        /// <param name="row"></param>
        /// <param name="col"></param>
        /// <returns></returns>
        /// ================================================================================
        private object GetValue(object[,] values, int row, int col)
        {
            int length_row = values.GetLength(0);
            int length_col = values.GetLength(1);

            try
            {
                if (row > length_row || col > length_col)
                    return null;

                var value = values[row, col];
                if (value != null)
                    return value;

                return null;
            }
            catch (System.Exception)
            {
                return null;
            }
        }

        /// ================================================================================
        /// <summary>Read File Excel</summary>
        /// <param name="path"></param>
        /// <returns></returns>
        /// ================================================================================
        public bool ReadFile(string path)
        {
            return false;
        }
    }
}