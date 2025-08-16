using Word = Microsoft.Office.Interop.Word;
using Microsoft.Office.Tools.Ribbon;
using System;
using System.Windows.Forms;
using Microsoft.VisualBasic;
using System.Collections.Generic;
using Excel = Microsoft.Office.Interop.Excel;
using System.Diagnostics;
using System.Runtime.InteropServices;

namespace FundFSAddIn
{
    public partial class Ribbon
    {
        private string _excelFilePath = null;
        private const string TablePrefix = "表格_";
        private const string TextPrefix = "文字_";

        private void Ribbon_Load(object sender, RibbonUIEventArgs e)
        {
            UpdateExcelFileNameLabel();
        }

        // 按下按鈕後，開啟檔案對話框選擇 Excel 附註檔
        private void btnSetExcelFilePath_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                var ofd = new OpenFileDialog
                {
                    Title = "選擇 Excel 檔案",
                    Filter = "Excel 檔案|*.xlsx;*.xlsm;*.xls",
                    CheckFileExists = true
                };
                if (ofd.ShowDialog() == DialogResult.OK)
                {
                    _excelFilePath = ofd.FileName;
                    UpdateExcelFileNameLabel();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("發生錯誤：\r\n" + ex.Message);
                Debug.WriteLine(ex);
            }
        }

        // 更新顯示的 Excel 檔案名稱標籤
        private void UpdateExcelFileNameLabel()
        {
            if (string.IsNullOrEmpty(_excelFilePath) || !System.IO.File.Exists(_excelFilePath))
            {
                lblExcelFileName.Label = "尚未指定附註檔";
                btnInsertTable.Enabled = false;
                btnInsertText.Enabled = false;
                btnGoToExcel.Enabled = false;
                btnUpdateOne.Enabled = false;
                btnUpdateAll.Enabled = false;
                btnDeleteCC.Enabled = false;
            }
            else
            {
                var fileName = System.IO.Path.GetFileNameWithoutExtension(_excelFilePath);
                lblExcelFileName.Label = "附註檔:" + fileName;
                btnInsertTable.Enabled = true;
                btnInsertText.Enabled = true;
                btnGoToExcel.Enabled = true;
                btnUpdateOne.Enabled = true;
                btnUpdateAll.Enabled = true;
                btnDeleteCC.Enabled = true;
            }
        }

        // 按下按鈕後，插入 Excel 附註檔中的表格圖片到 Word 文件
        private void btnInsertTable_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                ValidateExcelPath();

                var sheets = GetExcelSheetNames(_excelFilePath);
                if (sheets == null || sheets.Count == 0)
                {
                    MessageBox.Show("找不到任何工作表。", "錯誤");
                    return;
                }
                string sheet = ShowSheetSelectDialog(sheets);
                if (string.IsNullOrWhiteSpace(sheet)) return;
                string tag = sheet; // 直接用工作表名稱作為內容控制項名稱

                Word.Document doc = Globals.ThisAddIn.Application.ActiveDocument;
                ExcelImageHelper.CopyPrintAreaToClipboard(_excelFilePath, sheet);
                Word.Range wrange = doc.Application.Selection?.Range ?? doc.Content;
                wrange.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                Word.ContentControl cc = doc.ContentControls.Add(Word.WdContentControlType.wdContentControlRichText, wrange);
                cc.Tag = tag;
                cc.Title = tag;
                cc.Range.PasteSpecial(Word.WdPasteDataType.wdPasteEnhancedMetafile);
                cc.LockContents = true; // 貼上後再鎖定內容，避免使用者修改
                cc.LockContentControl = true; // 鎖定控制項本身不可刪除或移動
            }
            catch (Exception ex)
            {
                MessageBox.Show("發生錯誤：\r\n" + ex.Message);
                Debug.WriteLine(ex);
            }
        }

        // 取得帶有 "表格_" 前綴的工作表名稱列表
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
                    try
                    {
                        if (ws.Name.StartsWith(TablePrefix, StringComparison.Ordinal))
                        {
                            list.Add(ws.Name);
                        }
                    }
                    finally
                    {
                        ReleaseCom(ws);
                    }
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex);
            }
            finally
            {
                if (wb != null) wb.Close(false);
                if (excel != null) excel.Quit();
                ReleaseCom(wb);
                ReleaseCom(excel);
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
                var listBox = new ListBox { Dock = DockStyle.Fill, IntegralHeight = false };
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

        // 按下按鈕後，開啟 Excel 並跳轉到對應工作表
        private void btnGoToExcel_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                var app = Globals.ThisAddIn.Application;
                var sel = app.Selection;
                if (sel == null || sel.Range == null)
                {
                    MessageBox.Show("請先選取一個附註。", "提示");
                    return;
                }
                foreach (Word.ContentControl cc in sel.Range.ContentControls)
                {
                    if (!string.IsNullOrEmpty(cc.Tag))
                    {
                        string sheet = cc.Tag;
                        if (string.IsNullOrEmpty(_excelFilePath))
                        {
                            MessageBox.Show("無法取得附註檔Excel路徑。", "錯誤");
                            return;
                        }
                        var thisAddIn = Globals.ThisAddIn as FundFSAddIn.ThisAddIn;
                        thisAddIn.OpenExcelAndActivateSheet(_excelFilePath, sheet);
                        return;
                    }
                }
                MessageBox.Show("請先選取一個附註。", "提示");
            }
            catch (Exception ex)
            {
                MessageBox.Show("發生錯誤：\r\n" + ex.Message);
                Debug.WriteLine(ex);
            }
        }

        // 按下按鈕後，更新選取的內容控制項圖片
        private void btnUpdateOne_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                ValidateExcelPath();
                var app = Globals.ThisAddIn.Application;
                var sel = app.Selection;
                if (sel == null || sel.Range == null)
                {
                    MessageBox.Show("請先選取一個附註。", "提示");
                    return;
                }
                foreach (Word.ContentControl cc in sel.Range.ContentControls)
                {
                    if (!string.IsNullOrEmpty(cc.Tag))
                    {
                        if (cc.Tag.StartsWith(TablePrefix, StringComparison.Ordinal))
                        {
                            UpdateTableContentControl(cc, cc.Tag);
                            return;
                        }
                        if (cc.Tag.StartsWith(TextPrefix, StringComparison.Ordinal))
                        {
                            UpdateTextContentControl(cc, cc.Tag);
                            return;
                        }
                    }
                }
                MessageBox.Show("請先選取一個附註 (表格_* 或 文字_*)。", "提示");
            }
            catch (Exception ex)
            {
                MessageBox.Show("發生錯誤：\r\n" + ex.Message);
                Debug.WriteLine(ex);
            }
        }

        // 新增：更新全部附註 (文字 + 表格)
        private void btnUpdateAll_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                ValidateExcelPath();
                Word.Document doc = Globals.ThisAddIn.Application.ActiveDocument;
                int total = 0;
                foreach (Word.ContentControl ccAll in doc.ContentControls)
                {
                    if (!string.IsNullOrEmpty(ccAll.Tag) && (ccAll.Tag.StartsWith(TablePrefix, StringComparison.Ordinal) || ccAll.Tag.StartsWith(TextPrefix, StringComparison.Ordinal)))
                        total++;
                }
                int updated = 0;
                foreach (Word.ContentControl cc in doc.ContentControls)
                {
                    try
                    {
                        if (string.IsNullOrEmpty(cc.Tag)) continue;
                        if (cc.Tag.StartsWith(TablePrefix, StringComparison.Ordinal))
                        {
                            UpdateTableContentControl(cc, cc.Tag);
                            updated++;
                        }
                        else if (cc.Tag.StartsWith(TextPrefix, StringComparison.Ordinal))
                        {
                            UpdateTextContentControl(cc, cc.Tag);
                            updated++;
                        }
                    }
                    catch (Exception exOne)
                    {
                        Debug.WriteLine("更新控制項失敗: " + cc.Tag + " => " + exOne.Message);
                    }
                }
                MessageBox.Show("已更新附註數量:" + updated + " 全部附註數量:" + total, "更新完成");
            }
            catch (Exception ex)
            {
                MessageBox.Show("發生錯誤：\r\n" + ex.Message);
                Debug.WriteLine(ex);
            }
        }

        // 抽取：更新表格內容控制項
        private void UpdateTableContentControl(Word.ContentControl cc, string sheetName)
        {
            if (string.IsNullOrEmpty(sheetName)) throw new ArgumentException("sheetName 空白", nameof(sheetName));
            ExcelImageHelper.CopyPrintAreaToClipboard(_excelFilePath, sheetName);
            cc.LockContents = false;
            Word.Range r = cc.Range.Duplicate;
            r.Text = string.Empty;
            r.PasteSpecial(Word.WdPasteDataType.wdPasteEnhancedMetafile);
            cc.LockContents = true;
            cc.LockContentControl = true;
        }

        // 抽取：更新文字內容控制項
        private void UpdateTextContentControl(Word.ContentControl cc, string definedName)
        {
            if (string.IsNullOrEmpty(definedName)) throw new ArgumentException("definedName 空白", nameof(definedName));
            string val = GetExcelDefinedNameValue(_excelFilePath, definedName);
            if (val == null)
            {
                throw new Exception("找不到名稱值: " + definedName);
            }
            cc.LockContents = false;
            cc.Range.Text = val;
            cc.LockContents = true;
            cc.LockContentControl = true;
        }

        // 取得帶有 "文字_" 前綴的已定義名稱列表
        private List<string> GetExcelTextDefinedNames(string filePath)
        {
            var list = new List<string>();
            Excel.Application excel = null;
            Excel.Workbook wb = null;
            try
            {
                excel = new Excel.Application { Visible = false, DisplayAlerts = false };
                wb = excel.Workbooks.Open(filePath, ReadOnly: true);
                foreach (Excel.Name name in wb.Names)
                {
                    try
                    {
                        if (name != null && name.Name.StartsWith(TextPrefix, StringComparison.Ordinal))
                        {
                            list.Add(name.Name);
                        }
                    }
                    finally
                    {
                        ReleaseCom(name);
                    }
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex);
            }
            finally
            {
                if (wb != null) wb.Close(false);
                if (excel != null) excel.Quit();
                ReleaseCom(wb);
                ReleaseCom(excel);
            }
            return list;
        }

        // 讓使用者選擇一個已定義名稱，並將其值插入 Word 內容控制項
        private void btnInsertText_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                ValidateExcelPath();
                var names = GetExcelTextDefinedNames(_excelFilePath);
                if (names == null || names.Count == 0)
                {
                    MessageBox.Show("找不到任何文字定義名稱。", "錯誤");
                    return;
                }
                string name = ShowTextNameSelectDialog(names);
                if (string.IsNullOrWhiteSpace(name)) return;

                string value = GetExcelDefinedNameValue(_excelFilePath, name);
                if (value == null)
                {
                    MessageBox.Show("無法取得名稱值。", "錯誤");
                    return;
                }
                Word.Document doc = Globals.ThisAddIn.Application.ActiveDocument;
                Word.Range wrange = doc.Application.Selection?.Range ?? doc.Content;
                wrange.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                Word.ContentControl cc = doc.ContentControls.Add(Word.WdContentControlType.wdContentControlRichText, wrange);
                cc.Tag = name;
                cc.Title = name;
                cc.Range.Text = value;
                cc.LockContents = true;
                cc.LockContentControl = true;
            }
            catch (Exception ex)
            {
                MessageBox.Show("發生錯誤：\r\n" + ex.Message);
                Debug.WriteLine(ex);
            }
        }

        // 取得指定名稱的值
        private string GetExcelDefinedNameValue(string filePath, string name)
        {
            Excel.Application excel = null;
            Excel.Workbook wb = null;
            Excel.Range range = null;
            try
            {
                excel = new Excel.Application { Visible = false, DisplayAlerts = false };
                wb = excel.Workbooks.Open(filePath, ReadOnly: true);
                foreach (Excel.Name n in wb.Names)
                {
                    try
                    {
                        if (n != null && n.Name == name)
                        {
                            range = n.RefersToRange; // 可能為 null
                            if (range != null)
                            {
                                object val = range.Value2;
                                return val?.ToString();
                            }
                        }
                    }
                    finally
                    {
                        ReleaseCom(n);
                    }
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex);
            }
            finally
            {
                if (range != null) ReleaseCom(range);
                if (wb != null) wb.Close(false);
                if (excel != null) excel.Quit();
                ReleaseCom(wb);
                ReleaseCom(excel);
            }
            return null;
        }

        // 顯示選擇名稱的對話框
        private string ShowTextNameSelectDialog(List<string> names)
        {
            using (var form = new Form())
            {
                form.Text = "選擇文字名稱";
                form.Width = 300;
                form.Height = 180;
                var listBox = new ListBox { Dock = DockStyle.Fill, IntegralHeight = false };
                listBox.Items.AddRange(names.ToArray());
                form.Controls.Add(listBox);
                var btnOK = new Button { Text = "確定", Dock = DockStyle.Bottom, DialogResult = DialogResult.OK };
                form.Controls.Add(btnOK);
                form.AcceptButton = btnOK;
                if (form.ShowDialog() == DialogResult.OK && listBox.SelectedItem != null)
                    return listBox.SelectedItem.ToString();
            }
            return null;
        }

        private void ValidateExcelPath()
        {
            if (string.IsNullOrEmpty(_excelFilePath) || !System.IO.File.Exists(_excelFilePath))
            {
                throw new Exception("未指定附註檔");
            }
        }

        private static void ReleaseCom(object o)
        {
            try
            {
                if (o != null && Marshal.IsComObject(o))
                {
                    Marshal.ReleaseComObject(o);
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine("ReleaseCom 失敗: " + ex.Message);
            }
        }

        // 按下按鈕後，刪除選取的內容控制項 (補回遺失的方法)
        private void btnDeleteCC_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                var app = Globals.ThisAddIn.Application;
                var sel = app.Selection;
                if (sel == null || sel.Range == null)
                {
                    MessageBox.Show("請先選取一個附註。", "提示");
                    return;
                }
                bool deleted = false;
                foreach (Word.ContentControl cc in sel.Range.ContentControls)
                {
                    cc.LockContentControl = false;
                    cc.LockContents = false;
                    cc.Delete(true);
                    deleted = true;
                }
                if (!deleted)
                {
                    MessageBox.Show("請先選取一個附註。", "提示");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("發生錯誤：\r\n" + ex.Message);
                Debug.WriteLine(ex);
            }
        }
    }
}
