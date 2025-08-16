using System;
using System.Text.RegularExpressions;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;

namespace FundFSAddIn
{
    public partial class ThisAddIn
    {
        

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        public void OpenExcelAndActivateSheet(string filePath, string nameOrSheet)
        {
            if (!System.IO.File.Exists(filePath))
                throw new Exception("找不到來源 Excel 檔案：" + filePath);

            Excel.Application excel = null;
            Excel.Workbook wb = null;
            bool newExcelInstance = false;

            try
            {
                // 嘗試取得現有的 Excel 實例
                try
                {
                    excel = (Excel.Application)Marshal.GetActiveObject("Excel.Application");
                }
                catch
                {
                    // 如果沒有現有實例，創建新的
                    excel = new Excel.Application { Visible = true };
                    newExcelInstance = true;
                }

                excel.Visible = true; // 確保可見

                // 檢查檔案是否已經開啟
                foreach (Excel.Workbook openWb in excel.Workbooks)
                {
                    if (string.Equals(openWb.FullName, filePath, StringComparison.OrdinalIgnoreCase))
                    {
                        wb = openWb;
                        break;
                    }
                }

                // 如果檔案尚未開啟，則開啟它
                if (wb == null)
                {
                    wb = excel.Workbooks.Open(filePath, ReadOnly: false);
                }

                // 啟動對應的工作表或名稱
                if (nameOrSheet.StartsWith("表格_"))
                {
                    // 處理表格工作表
                    Excel.Worksheet ws = null;
                    foreach (Excel.Worksheet s in wb.Sheets)
                    {
                        if (string.Equals(s.Name, nameOrSheet, StringComparison.OrdinalIgnoreCase))
                        {
                            ws = s;
                            break;
                        }
                    }
                    if (ws == null)
                        throw new Exception("找不到指定工作表：" + nameOrSheet);
                    ws.Activate();
                }
                else if (nameOrSheet.StartsWith("文字_"))
                {
                    // 處理文字定義名稱
                    bool found = false;
                    foreach (Excel.Name n in wb.Names)
                    {
                        if (n.Name == nameOrSheet)
                        {
                            var refersTo = n.RefersTo;
                            // 例如 =Sheet1!$A$1
                            string sheetName = null;
                            string cellAddress = null;
                            var match = Regex.Match(refersTo, @"=([^!]+)!([$A-Z0-9:]+)");
                            if (match.Success)
                            {
                                sheetName = match.Groups[1].Value.Replace("'", ""); // 去除單引號
                                cellAddress = match.Groups[2].Value;
                            }
                            if (!string.IsNullOrEmpty(sheetName))
                            {
                                Excel.Worksheet ws = null;
                                foreach (Excel.Worksheet s in wb.Sheets)
                                {
                                    if (string.Equals(s.Name, sheetName, StringComparison.OrdinalIgnoreCase))
                                    {
                                        ws = s;
                                        break;
                                    }
                                }
                                if (ws != null)
                                {
                                    ws.Activate();
                                    ws.Application.Goto(ws.Range[cellAddress]);
                                    found = true;
                                    break;
                                }
                            }
                        }
                    }
                    if (!found)
                        throw new Exception("找不到對應的 Excel 名稱或範圍：" + nameOrSheet);
                }
                else
                {
                    throw new Exception("不支援的前綴，需為 表格_ 或 文字_：" + nameOrSheet);
                }

                // 將 Excel 視窗帶到前景
                excel.WindowState = Excel.XlWindowState.xlNormal;
                // excel.Activate(); // Application 無此方法，移除
                IntPtr hwnd = (IntPtr)excel.Hwnd;
                if (hwnd != IntPtr.Zero)
                {
                    NativeMethods.SetForegroundWindow(hwnd);
                }
            }
            catch (Exception ex)
            {
                // 如果是新創建的實例且出錯，關閉它
                if (newExcelInstance && excel != null)
                {
                    try
                    {
                        if (wb != null) wb.Close(false);
                        excel.Quit();
                    }
                    catch { }
                }
                throw new Exception("Excel 操作失敗：" + ex.Message);
            }
            // 注意：不釋放 COM 物件，保持 Excel 開啟狀態
        }


        #region VSTO 產生的程式碼
        /// <summary>
        /// 此為設計工具支援所需的方法 - 請勿使用程式碼編輯器修改
        /// 這個方法的內容。
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        #endregion
    }

    internal static class NativeMethods
    {
        [DllImport("user32.dll")]
        internal static extern bool SetForegroundWindow(IntPtr hWnd);
    }
}
