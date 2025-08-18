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

        // 共用的 Excel Application / Workbook
        private Excel.Application _sharedExcelApp;
        private Excel.Workbook _sharedWorkbook;

        //建立或取得共用 Workbook
        private void EnsureWorkbook()
        {
            if (string.IsNullOrEmpty(_excelFilePath))
                throw new InvalidOperationException("尚未指定 Excel 路徑。");

            // 若已存在且活著就直接用
            if (_sharedWorkbook != null && _sharedExcelApp != null)
            {
                try
                {
                    var _ = _sharedWorkbook.Name;
                    return;
                }
                catch
                {
                    ReleaseExcelResources();
                }
            }

            // 只用 new Excel.Application，不要用 Marshal.GetActiveObject
            _sharedExcelApp = new Excel.Application
            {
                Visible = false,
                DisplayAlerts = false
            };
            _sharedWorkbook = _sharedExcelApp.Workbooks.Open(_excelFilePath, ReadOnly: true);
        }

        // 釋放共用資源
        private void ReleaseExcelResources()
        {
            try
            {
                if (_sharedWorkbook != null)
                {
                    try
                    {
                        _sharedWorkbook.Close(false);
                    }
                    catch { }
                    finally
                    {
                        Marshal.FinalReleaseComObject(_sharedWorkbook);
                        _sharedWorkbook = null;
                    }
                }
            }
            catch { }

            try
            {
                if (_sharedExcelApp != null)
                {
                    try
                    {
                        _sharedExcelApp.Quit();
                    }
                    catch { }
                    finally
                    {
                        Marshal.FinalReleaseComObject(_sharedExcelApp);
                        _sharedExcelApp = null;
                    }
                }
            }
            catch { }

            // 強制垃圾回收，確保釋放所有 RCW
            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();
            GC.WaitForPendingFinalizers();
        }

        private void Ribbon_Load(object sender, RibbonUIEventArgs e)
        {
            UpdateExcelFileNameLabel();
        }

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
                    if (!string.Equals(_excelFilePath, ofd.FileName, StringComparison.OrdinalIgnoreCase))
                    {
                        // 路徑改變時釋放舊的共用資源
                        ReleaseExcelResources();
                    }
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
                string tag = sheet;

                EnsureWorkbook(); // 使用共用 Workbook
                Excel.Worksheet ws = null;
                try
                {
                    ws = (Excel.Worksheet)_sharedWorkbook.Sheets[sheet];
                    string printArea = ws.PageSetup.PrintArea;
                    if (string.IsNullOrWhiteSpace(printArea))
                        throw new Exception("該工作表未設定列印範圍");
                    Excel.Range rng = ws.Range[printArea];
                    rng.Copy();
                    Word.Document doc = Globals.ThisAddIn.Application.ActiveDocument;
                    Word.Range wrange = doc.Application.Selection?.Range ?? doc.Content;
                    wrange.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                    Word.ContentControl cc = doc.ContentControls.Add(Word.WdContentControlType.wdContentControlRichText, wrange);
                    cc.Tag = tag;
                    cc.Title = tag;
                    cc.Range.PasteSpecial(
                        DataType: Word.WdPasteDataType.wdPasteOLEObject,
                        Link: true,
                        Placement: Word.WdOLEPlacement.wdInLine,
                        DisplayAsIcon: false);
                    cc.Range.Font.Reset();
                    cc.Range.Bold = 0;
                    cc.Range.Italic = 0;
                    cc.Range.Underline = Word.WdUnderline.wdUnderlineNone;
                    cc.Range.HighlightColorIndex = Word.WdColorIndex.wdNoHighlight;
                    cc.Range.Font.ColorIndex = Word.WdColorIndex.wdAuto;
                    cc.Range.ParagraphFormat.SpaceBefore = 0f;
                    cc.Range.ParagraphFormat.SpaceAfter = 0f;
                    cc.Range.ParagraphFormat.LineSpacingRule = Word.WdLineSpacing.wdLineSpaceSingle;
                    cc.LockContents = true;
                    cc.LockContentControl = true;
                }
                finally
                {
                    ReleaseCom(ws);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("發生錯誤：\r\n" + ex.Message);
                Debug.WriteLine(ex);
            }
        }

        private List<string> GetExcelSheetNames(string filePath)
        {
            var list = new List<string>();
            try
            {
                EnsureWorkbook();
                foreach (Excel.Worksheet ws in _sharedWorkbook.Sheets)
                {
                    try
                    {
                        if (ws.Name.StartsWith(TablePrefix, StringComparison.Ordinal))
                            list.Add(ws.Name);
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
            return list;
        }

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

        private void btnUpdateAll_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                ValidateExcelPath();
                Word.Document doc = Globals.ThisAddIn.Application.ActiveDocument;
                int total = 0;
                foreach (Word.ContentControl ccAll in doc.ContentControls)
                {
                    if (!string.IsNullOrEmpty(ccAll.Tag) &&
                        (ccAll.Tag.StartsWith(TablePrefix, StringComparison.Ordinal) ||
                         ccAll.Tag.StartsWith(TextPrefix, StringComparison.Ordinal)))
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

        private void UpdateTableContentControl(Word.ContentControl cc, string sheetName)
        {
            if (string.IsNullOrEmpty(sheetName)) throw new ArgumentException("sheetName 空白", nameof(sheetName));
            // 注意：ExcelImageHelper 內部若再開 Excel 仍會產生額外實例，需考慮也改造成共用。
            ExcelImageHelper.CopyPrintAreaToClipboard(_excelFilePath, sheetName);
            cc.LockContents = false;
            Word.Range r = cc.Range.Duplicate;
            r.Text = string.Empty;
            r.PasteSpecial(Word.WdPasteDataType.wdPasteEnhancedMetafile);
            cc.LockContents = true;
            cc.LockContentControl = true;
        }

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

        private List<string> GetExcelTextDefinedNames(string filePath)
        {
            var list = new List<string>();
            try
            {
                EnsureWorkbook();
                foreach (Excel.Name name in _sharedWorkbook.Names)
                {
                    try
                    {
                        if (name != null && name.Name.StartsWith(TextPrefix, StringComparison.Ordinal))
                            list.Add(name.Name);
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
            return list;
        }

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

        private string GetExcelDefinedNameValue(string filePath, string name)
        {
            try
            {
                if (!string.Equals(filePath, _excelFilePath, StringComparison.OrdinalIgnoreCase))
                    throw new InvalidOperationException("傳入的檔案路徑與目前已載入的路徑不一致。");
                EnsureWorkbook();
                foreach (Excel.Name n in _sharedWorkbook.Names)
                {
                    try
                    {
                        if (n != null && n.Name == name)
                        {
                            Excel.Range range = null;
                            try
                            {
                                range = n.RefersToRange;
                                if (range != null)
                                {
                                    object val = range.Value2;
                                    return val?.ToString();
                                }
                            }
                            finally
                            {
                                ReleaseCom(range);
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
            return null;
        }

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
                throw new Exception("未指定附註檔");
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
