using System;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;
using Excel = Microsoft.Office.Interop.Excel;

namespace FundFSAddIn
{
    public partial class ThisAddIn
    {
        private string _lastExcelFilePath = null;

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            this.Application.WindowBeforeDoubleClick += Application_WindowBeforeDoubleClick;
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        private void Application_WindowBeforeDoubleClick(Word.Selection Sel, ref bool Cancel)
        {
            if (Sel == null || Sel.Range == null) return;
            foreach (Word.ContentControl cc in Sel.Range.ContentControls)
            {
                // 只處理插入的圖片內容控制項（Tag 為工作表名稱）
                if (!string.IsNullOrEmpty(cc.Tag))
                {
                    string sheet = cc.Tag;
                    string file = _lastExcelFilePath;
                    if (string.IsNullOrEmpty(file))
                    {
                        MessageBox.Show("無法取得來源 Excel 檔案路徑，請先插入一次內容控制項。", "錯誤");
                        return;
                    }
                    try
                    {
                        OpenExcelAndActivateSheet(file, sheet);
                        Cancel = true;
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("無法開啟 Excel 或工作表：\r\n" + ex.Message);
                    }
                    break;
                }
            }
        }

        public void OpenExcelAndActivateSheet(string filePath, string sheetName)
        {
            if (!System.IO.File.Exists(filePath))
                throw new Exception("找不到來源 Excel 檔案：" + filePath);
            var excel = new Excel.Application { Visible = true };
            var wb = excel.Workbooks.Open(filePath, ReadOnly: false);
            Excel.Worksheet ws = null;
            try
            {
                foreach (Excel.Worksheet s in wb.Sheets)
                {
                    if (string.Equals(s.Name, sheetName, StringComparison.OrdinalIgnoreCase))
                    {
                        ws = s;
                        break;
                    }
                }
                if (ws == null)
                    throw new Exception("找不到指定工作表：" + sheetName);
                ws.Activate();
            }
            catch
            {
                wb.Close(false);
                throw;
            }
        }

        public string GetLastExcelFilePath()
        {
            return _lastExcelFilePath;
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
}
