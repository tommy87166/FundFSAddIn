using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using Word = Microsoft.Office.Interop.Word;

namespace FundFSAddIn
{
    public static class ExcelImageHelper
    {
        // 將 Excel Print Area 以增強型中繼圖格式貼到 Word 內容控制項
        public static void PasteExcelPrintAreaAsMetafileToContentControl(Word.Document doc, string workbookPath, string sheetNameOrIndex, string tag)
        {
            Excel.Application excel = null;
            Excel.Workbook wb = null;
            Excel.Worksheet ws = null;
            try
            {
                excel = new Excel.Application { Visible = false, DisplayAlerts = false };
                wb = excel.Workbooks.Open(workbookPath, ReadOnly: true);
                ws = ResolveSheet(wb, sheetNameOrIndex);
                Excel.Range rng = GetPrintAreaOrUsedRange(ws);
                rng.CopyPicture(Excel.XlPictureAppearance.xlScreen, Excel.XlCopyPictureFormat.xlPicture);

                Word.Range wrange = doc.Application.Selection?.Range ?? doc.Content;
                wrange.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                Word.ContentControl cc = doc.ContentControls.Add(Word.WdContentControlType.wdContentControlRichText, wrange);
                cc.Tag = tag;
                cc.Title = tag;
                cc.Range.PasteSpecial(Word.WdPasteDataType.wdPasteEnhancedMetafile);
                cc.LockContents = true; // 貼上後再鎖定內容，避免使用者修改
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

        // 將 Excel Print Area 以 EMF 格式複製到剪貼簿，供內容控制項更新時使用
        public static void CopyPrintAreaToClipboardAsMetafile(string workbookPath, string sheetNameOrIndex)
        {
            Excel.Application excel = null;
            Excel.Workbook wb = null;
            Excel.Worksheet ws = null;
            try
            {
                excel = new Excel.Application { Visible = false, DisplayAlerts = false };
                wb = excel.Workbooks.Open(workbookPath, ReadOnly: true);
                ws = ResolveSheet(wb, sheetNameOrIndex);
                Excel.Range rng = GetPrintAreaOrUsedRange(ws);
                rng.CopyPicture(Excel.XlPictureAppearance.xlScreen, Excel.XlCopyPictureFormat.xlPicture);
                // 圖片已在剪貼簿
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
        private static Excel.Worksheet ResolveSheet(Excel.Workbook wb, string sheetNameOrIndex)
        {
            if (string.IsNullOrWhiteSpace(sheetNameOrIndex))
                return (Excel.Worksheet)wb.Sheets[1];

            if (int.TryParse(sheetNameOrIndex, out int idx))
                return (Excel.Worksheet)wb.Sheets[idx];

            return (Excel.Worksheet)wb.Sheets[sheetNameOrIndex];
        }

        // 取得 Print Area（可多區域），若未設定則回傳 UsedRange
        private static Excel.Range GetPrintAreaOrUsedRange(Excel.Worksheet ws)
        {
            string printArea = ws.PageSetup.PrintArea; // 可能為空或 "A1:D20,A30:D40" 等
            if (string.IsNullOrWhiteSpace(printArea))
                return ws.UsedRange;

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
