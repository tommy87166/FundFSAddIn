using System;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;
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
