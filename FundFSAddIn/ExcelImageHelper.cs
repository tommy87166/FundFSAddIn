using System;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;
using System.Diagnostics;

namespace FundFSAddIn
{
    public static class ExcelImageHelper
    {
        // 將工作表的Print Area以EMF格式複製到剪貼簿
        public static void CopyPrintAreaToClipboard(string workbookPath, string sheetName)
        {
            Excel.Application excel = null;
            Excel.Workbook wb = null;
            Excel.Worksheet ws = null;
            Excel.Range rng = null;
            Excel.Window win = null;
            Excel.XlWindowView? originalView = null;
            try
            {
                excel = new Excel.Application { Visible = false, DisplayAlerts = false };
                wb = excel.Workbooks.Open(workbookPath, ReadOnly: true);
                ws = ResolveSheet(wb, sheetName);

                // 記錄原本檢視模式 (可能需要 ActiveWindow)
                win = ws.Application.ActiveWindow;
                if (win != null)
                {
                    originalView = (Excel.XlWindowView)win.View;
                    if (originalView != Excel.XlWindowView.xlNormalView)
                        win.View = Excel.XlWindowView.xlNormalView;
                }
                rng = GetPrintArea(ws);
                rng.CopyPicture(Excel.XlPictureAppearance.xlScreen, Excel.XlCopyPictureFormat.xlPicture);

                // 還原檢視模式
                if (originalView != null && win != null && win.View != originalView)
                    win.View = originalView.Value;
            }
            catch (Exception ex)
            {
                Debug.WriteLine("CopyPrintAreaToClipboard 失敗: " + ex);
                throw; // 讓呼叫端決定如何顯示
            }
            finally
            {
                ReleaseCom(rng);
                ReleaseCom(win);
                if (wb != null) wb.Close(false);
                if (excel != null) excel.Quit();
                ReleaseCom(ws);
                ReleaseCom(wb);
                ReleaseCom(excel);
            }
        }

        // ---------- helpers ----------
        private static Excel.Worksheet ResolveSheet(Excel.Workbook wb, string sheetName)
        {
            if (wb == null) throw new ArgumentNullException(nameof(wb));
            if (string.IsNullOrWhiteSpace(sheetName))
                throw new ArgumentException("必須提供工作表名稱", nameof(sheetName));
            try
            {
                return (Excel.Worksheet)wb.Sheets[sheetName];
            }
            catch (Exception ex)
            {
                throw new Exception("找不到指定工作表：" + sheetName, ex);
            }
        }

        // 取得 Print Area（可多區域），若未設定則拋出例外
        private static Excel.Range GetPrintArea(Excel.Worksheet ws)
        {
            string printArea = ws.PageSetup.PrintArea; // 可能為空或 "A1:D20,A30:D40" 等
            if (string.IsNullOrWhiteSpace(printArea))
                throw new Exception("該工作表未設定列印範圍");

            string[] areas = printArea.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries);
            Excel.Range union = ws.Range[areas[0]];
            for (int i = 1; i < areas.Length; i++)
            {
                Excel.Range next = ws.Range[areas[i]];
                try
                {
                    union = ws.Application.Union(union, next);
                }
                finally
                {
                    ReleaseCom(next);
                }
            }
            return union;
        }

        private static void ReleaseCom(object o)
        {
            try
            {
                if (o != null && Marshal.IsComObject(o))
                    Marshal.ReleaseComObject(o);
            }
            catch (Exception ex)
            {
                Debug.WriteLine("ReleaseCom 失敗: " + ex.Message);
            }
        }
    }
}
