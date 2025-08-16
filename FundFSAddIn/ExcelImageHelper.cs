using System;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;

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
            Excel.XlWindowView? originalView = null;
            try
            {
                excel = new Excel.Application { Visible = false, DisplayAlerts = false };
                wb = excel.Workbooks.Open(workbookPath, ReadOnly: true);
                ws = ResolveSheet(wb, sheetName);

                // 記錄原本檢視模式
                originalView = (Excel.XlWindowView)ws.Application.ActiveWindow.View;
                // 若不是標準檢視，切換到標準檢視
                if (originalView != Excel.XlWindowView.xlNormalView)
                    ws.Application.ActiveWindow.View = Excel.XlWindowView.xlNormalView;

                Excel.Range rng = GetPrintArea(ws);
                rng.CopyPicture(Excel.XlPictureAppearance.xlScreen, Excel.XlCopyPictureFormat.xlPicture);

                // 複製完畢後還原檢視模式
                if (originalView != null && ws.Application.ActiveWindow.View != originalView)
                    ws.Application.ActiveWindow.View = originalView.Value;
            }
            finally
            {
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
            catch
            {
                throw new Exception("找不到指定工作表：" + sheetName);
            }
        }

        // 取得 Print Area（可多區域），若未設定則拋出例外
        private static Excel.Range GetPrintArea(Excel.Worksheet ws)
        {
            string printArea = ws.PageSetup.PrintArea; // 可能為空或 "A1:D20,A30:D40" 等
            if (string.IsNullOrWhiteSpace(printArea))
                throw new Exception("該工作表未設定列印範圍");

            // 多區域以逗號分隔；需逐一 union
            string[] areas = printArea.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries);
            Excel.Range union = ws.Range[areas[0]];
            for (int i = 1; i < areas.Length; i++)
            {
                Excel.Range next = ws.Range[areas[i]];
                union = ws.Application.Union(union, next);
            }
            return union;
        }

        private static void ReleaseCom(object o)
        {
            if (o != null && Marshal.IsComObject(o))
                Marshal.ReleaseComObject(o);
        }
    }
}
