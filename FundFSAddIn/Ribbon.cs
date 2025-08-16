using Word = Microsoft.Office.Interop.Word;
using Microsoft.Office.Tools.Ribbon;
using System;
using System.Windows.Forms;
using Microsoft.VisualBasic;
using System.Collections.Generic;
using Excel = Microsoft.Office.Interop.Excel;

namespace FundFSAddIn
{
    public partial class Ribbon
    {
        private string _excelFilePath = null;

        private void Ribbon_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void btnInsert_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                if (string.IsNullOrEmpty(_excelFilePath) || !System.IO.File.Exists(_excelFilePath))
                {
                    var ofd = new OpenFileDialog
                    {
                        Title = "選擇 Excel 檔案",
                        Filter = "Excel 檔案|*.xlsx;*.xlsm;*.xls"
                    };
                    if (ofd.ShowDialog() != DialogResult.OK) return;
                    _excelFilePath = ofd.FileName;
                }

                var sheets = GetExcelSheetNames(_excelFilePath);
                if (sheets == null || sheets.Count == 0)
                {
                    MessageBox.Show("找不到任何工作表。", "錯誤");
                    return;
                }

                string sheet = ShowSheetSelectDialog(sheets);
                if (string.IsNullOrWhiteSpace(sheet)) return;

                // 直接用工作表名稱作為內容控制項名稱
                string tag = sheet;

                // 更新 ThisAddIn 的 _lastExcelFilePath 供雙擊時使用
                var thisAddIn = Globals.ThisAddIn as FundFSAddIn.ThisAddIn;
                if (thisAddIn != null)
                {
                    var field = typeof(FundFSAddIn.ThisAddIn).GetField("_lastExcelFilePath", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance);
                    if (field != null) field.SetValue(thisAddIn, _excelFilePath);
                }

                Word.Document doc = Globals.ThisAddIn.Application.ActiveDocument;
                ExcelImageHelper.PasteExcelPrintAreaAsMetafileToContentControl(doc, _excelFilePath, sheet, tag);

                MessageBox.Show($"已插入 Excel 原生圖片（EMF）並放入內容控制項，Tag = {tag}");
            }
            catch (Exception ex)
            {
                MessageBox.Show("發生錯誤：\r\n" + ex.Message);
            }
        }

        private void btnUpdate_Click(object sender, RibbonControlEventArgs e)
        {
            MessageBox.Show("目前僅支援插入 EMF 圖片，不支援更新內容控制項圖片。請先刪除舊內容控制項再插入新圖片。", "提示");
        }

        private string Prompt(string text, string defaultValue)
        {
            return Microsoft.VisualBasic.Interaction.InputBox(text, "輸入", defaultValue ?? "");
        }

        // 取得 Excel 檔案的所有工作表名稱
        private List<string> GetExcelSheetNames(string filePath)
        {
            var list = new List<string>();
            Excel.Application excel = null;
            Excel.Workbook wb = null;
            try
            {
                excel = new Excel.Application { Visible = false, DisplayAlerts = false };
                wb = excel.Workbooks.Open(filePath, ReadOnly: true);
                foreach (Excel.Worksheet ws in wb.Sheets)
                {
                    list.Add(ws.Name);
                }
            }
            catch { }
            finally
            {
                if (wb != null) wb.Close(false);
                if (excel != null) excel.Quit();
                if (wb != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(wb);
                if (excel != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(excel);
            }
            return list;
        }

        // 顯示選擇工作表的對話框
        private string ShowSheetSelectDialog(List<string> sheets)
        {
            using (var form = new Form())
            {
                form.Text = "選擇工作表";
                form.Width = 300;
                form.Height = 180;
                var listBox = new ListBox { Dock = DockStyle.Fill };
                listBox.Items.AddRange(sheets.ToArray());
                form.Controls.Add(listBox);
                var btnOK = new Button { Text = "確定", Dock = DockStyle.Bottom, DialogResult = DialogResult.OK };
                form.Controls.Add(btnOK);
                form.AcceptButton = btnOK;
                if (form.ShowDialog() == DialogResult.OK && listBox.SelectedItem != null)
                    return listBox.SelectedItem.ToString();
            }
            return null;
        }

        private void btnGoToExcel_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                var app = Globals.ThisAddIn.Application;
                var sel = app.Selection;
                if (sel == null || sel.Range == null)
                {
                    MessageBox.Show("請先選取一個內容控制項。", "提示");
                    return;
                }
                foreach (Word.ContentControl cc in sel.Range.ContentControls)
                {
                    if (!string.IsNullOrEmpty(cc.Tag))
                    {
                        string sheet = cc.Tag;
                        var thisAddIn = Globals.ThisAddIn as FundFSAddIn.ThisAddIn;
                        var file = thisAddIn?.GetLastExcelFilePath();
                        if (string.IsNullOrEmpty(file))
                        {
                            MessageBox.Show("無法取得來源 Excel 檔案路徑，請先插入一次內容控制項。", "錯誤");
                            return;
                        }
                        thisAddIn.OpenExcelAndActivateSheet(file, sheet);
                        return;
                    }
                }
                MessageBox.Show("請先選取一個內容控制項。", "提示");
            }
            catch (Exception ex)
            {
                MessageBox.Show("發生錯誤：\r\n" + ex.Message);
            }
        }

        private void btnUpdatePicture_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                var app = Globals.ThisAddIn.Application;
                var sel = app.Selection;
                if (sel == null || sel.Range == null)
                {
                    MessageBox.Show("請先選取一個內容控制項。", "提示");
                    return;
                }
                foreach (Word.ContentControl cc in sel.Range.ContentControls)
                {
                    if (!string.IsNullOrEmpty(cc.Tag))
                    {
                        string sheet = cc.Tag;
                        var thisAddIn = Globals.ThisAddIn as FundFSAddIn.ThisAddIn;
                        var file = thisAddIn?.GetLastExcelFilePath();
                        if (string.IsNullOrEmpty(file))
                        {
                            MessageBox.Show("無法取得來源 Excel 檔案路徑，請先插入一次內容控制項。", "錯誤");
                            return;
                        }
                        // 取得最新圖片到剪貼簿
                        ExcelImageHelper.CopyPrintAreaToClipboardAsMetafile(file, sheet);
                        cc.LockContents = false;
                        cc.Range.Delete();
                        cc.Range.PasteSpecial(Word.WdPasteDataType.wdPasteEnhancedMetafile);
                        cc.LockContents = true;
                        MessageBox.Show("已更新圖片。", "成功");
                        return;
                    }
                }
                MessageBox.Show("請先選取一個內容控制項。", "提示");
            }
            catch (Exception ex)
            {
                MessageBox.Show("發生錯誤：\r\n" + ex.Message);
            }
        }
    }
}
