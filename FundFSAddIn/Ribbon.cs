using Microsoft.Office.Interop.Word;
using Microsoft.Office.Tools.Ribbon;
using Microsoft.VisualBasic;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using Word = Microsoft.Office.Interop.Word;

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
            _sharedWorkbook = _sharedExcelApp.Workbooks.Open(_excelFilePath, ReadOnly: false);
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

        //載入Ribbon
        private void Ribbon_Load(object sender, RibbonUIEventArgs e)
        {
            UpdateExcelFileNameLabel();
        }
        private void UpdateExcelFileNameLabel()
        {
            if (string.IsNullOrEmpty(_excelFilePath) || !System.IO.File.Exists(_excelFilePath))
            {
                lblExcelFileName.Label = " ❌  尚未指定附註檔";
                btnInsertTable.Enabled = false;
                btnInsertText.Enabled = false;
                btnGoToExcel.Enabled = false;
                btnUpdateOne.Enabled = false;
                btnUpdateAll.Enabled = false;
                btnDeleteCC.Enabled = false;
                btnRemapLinks.Enabled = false;
            }
            else
            {
                var fileName = System.IO.Path.GetFileNameWithoutExtension(_excelFilePath);
                lblExcelFileName.Label = " ✔️  已開啟附註檔(" + fileName+")";
                btnInsertTable.Enabled = true;
                btnInsertText .Enabled = true;
                btnGoToExcel.Enabled = true;
                btnUpdateOne.Enabled = true;
                btnUpdateAll.Enabled = true;
                btnDeleteCC.Enabled = true;
                btnRemapLinks.Enabled = true;
            }
        }

        private void ValidateExcelPath()
        {
            if (string.IsNullOrEmpty(_excelFilePath) || !System.IO.File.Exists(_excelFilePath))
                throw new Exception("未指定附註檔");
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
                    foreach (Word.InlineShape shape in cc.Range.InlineShapes)
                    {
                        if (shape.Type == Word.WdInlineShapeType.wdInlineShapeLinkedOLEObject && shape.LinkFormat != null)
                        {
                            shape.LinkFormat.AutoUpdate = false;
                        }
                    }
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
                MessageBox.Show("發生錯誤：\r\n" + ex.Message);
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
                        string nameOrSheet = cc.Tag;
                        if (string.IsNullOrEmpty(_excelFilePath))
                        {
                            MessageBox.Show("無法取得附註檔Excel路徑。", "錯誤");
                            return;
                        }

                        EnsureWorkbook();
                        _sharedExcelApp.Visible = true;

                        if (nameOrSheet.StartsWith(TablePrefix, StringComparison.OrdinalIgnoreCase))
                        {
                            // 表格_：直接啟動對應工作表
                            Excel.Worksheet ws = null;
                            foreach (Excel.Worksheet s in _sharedWorkbook.Sheets)
                            {
                                if (string.Equals(s.Name, nameOrSheet, StringComparison.OrdinalIgnoreCase))
                                {
                                    ws = s;
                                    break;
                                }
                            }
                            if (ws == null)
                            {
                                MessageBox.Show("找不到指定工作表：" + nameOrSheet, "錯誤");
                                return;
                            }
                            ws.Activate();
                        }
                        else if (nameOrSheet.StartsWith(TextPrefix, StringComparison.OrdinalIgnoreCase))
                        {
                            // 文字_：啟動對應名稱的儲存格
                            bool found = false;
                            foreach (Excel.Name n in _sharedWorkbook.Names)
                            {
                                if (n.Name == nameOrSheet)
                                {
                                    var refersTo = n.RefersTo;
                                    string sheetName = null;
                                    string cellAddress = null;
                                    var match = System.Text.RegularExpressions.Regex.Match(refersTo, @"=([^!]+)!([$A-Z0-9:]+)");
                                    if (match.Success)
                                    {
                                        sheetName = match.Groups[1].Value.Replace("'", "");
                                        cellAddress = match.Groups[2].Value;
                                    }
                                    if (!string.IsNullOrEmpty(sheetName))
                                    {
                                        Excel.Worksheet ws = null;
                                        foreach (Excel.Worksheet s in _sharedWorkbook.Sheets)
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
                            {
                                MessageBox.Show("找不到對應的 Excel 名稱或範圍：" + nameOrSheet, "錯誤");
                                return;
                            }
                        }
                        else
                        {
                            MessageBox.Show("不支援的前綴，需為 表格_ 或 文字_：" + nameOrSheet, "錯誤");
                            return;
                        }

                        // 將 Excel 視窗帶到前景
                        _sharedExcelApp.WindowState = Excel.XlWindowState.xlNormal;
                        IntPtr hwnd = (IntPtr)_sharedExcelApp.Hwnd;
                        if (hwnd != IntPtr.Zero)
                        {
                            NativeMethods.SetForegroundWindow(hwnd);
                        }
                        return;
                    }
                }
                MessageBox.Show("請先選取一個附註。", "提示");
            }
            catch (Exception ex)
            {
                MessageBox.Show("發生錯誤：\r\n" + ex.Message);
            }
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
                MessageBox.Show("發生錯誤：\r\n" + ex.Message);
            }
            return list;
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
            }
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

                // 取得 Excel 範圍
                EnsureWorkbook();
                Excel.Name excelName = null;
                Excel.Range range = null;
                try
                {
                    foreach (Excel.Name n in _sharedWorkbook.Names)
                    {
                        if (n != null && n.Name == name)
                        {
                            excelName = n;
                            range = n.RefersToRange;
                            break;
                        }
                    }
                    if (range == null)
                    {
                        MessageBox.Show("無法取得名稱對應的範圍。", "錯誤");
                        return;
                    }
                    range.Copy();
                }
                finally
                {
                    ReleaseCom(range);
                    ReleaseCom(excelName);
                }

                // 插入 Word 內容控制項並貼上 OLE 物件
                Word.Document doc = Globals.ThisAddIn.Application.ActiveDocument;
                Word.Range wrange = doc.Application.Selection?.Range ?? doc.Content;
                wrange.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                Word.ContentControl cc = doc.ContentControls.Add(Word.WdContentControlType.wdContentControlRichText, wrange);
                cc.Tag = name;
                cc.Title = name;
                // 貼上 OLE 物件並合併格式
                cc.Range.PasteSpecial(
                    DataType: Word.WdPasteDataType.wdPasteRTF,
                    Link: true,
                    Placement: Word.WdOLEPlacement.wdInLine,
                    DisplayAsIcon: false);
                foreach (Word.Field field in cc.Range.Fields)
                {
                    if (field.LinkFormat != null)
                    {
                        field.LinkFormat.AutoUpdate = false;
                    }
                }
                cc.LockContents = true;
                cc.LockContentControl = true;
            }
            catch (Exception ex)
            {
                MessageBox.Show("發生錯誤：\r\n" + ex.Message);
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
                int updatedCount = 0;
                foreach (Word.ContentControl cc in sel.Range.ContentControls)
                {
                    foreach (Word.Field field in cc.Range.Fields)
                    {
                        if (!string.IsNullOrEmpty(cc.Tag) && (cc.Tag.StartsWith(TablePrefix, StringComparison.Ordinal) || cc.Tag.StartsWith(TextPrefix, StringComparison.Ordinal)))
                        {
                            field.Update();
                            updatedCount++;
                        }
                    }
                }
                MessageBox.Show($"已更新{updatedCount}個附註。", "完成");
            }
            catch (Exception ex)
            {
                MessageBox.Show("發生錯誤：\r\n" + ex.Message);
            }
        }

        private void btnUpdateAll_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                ValidateExcelPath();
                Word.Document doc = Globals.ThisAddIn.Application.ActiveDocument;
                int updatedCount = 0;
                foreach (Word.ContentControl cc in doc.ContentControls)
                {
                    foreach (Word.Field field in cc.Range.Fields)
                        {   
                            if (!string.IsNullOrEmpty(cc.Tag) && (cc.Tag.StartsWith(TablePrefix, StringComparison.Ordinal) || cc.Tag.StartsWith(TextPrefix, StringComparison.Ordinal))) {
                                field.Update();
                                updatedCount++;
                            }
                        }
                }
                MessageBox.Show($"已更新{updatedCount}個附註。", "完成");
            }
            catch (Exception ex)
            {
                MessageBox.Show("發生錯誤：\r\n" + ex.Message);
            }
        }

        private void btnRemapLinks_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                ValidateExcelPath();
                Word.Document doc = Globals.ThisAddIn.Application.ActiveDocument;
                int updatedField = 0;
                string escapedPath = EscapePathForLinkField(_excelFilePath); // 轉義路徑（雙反斜線）

                foreach (Word.ContentControl cc in doc.ContentControls)
                {
                    cc.LockContents = false;
                    foreach (Word.Field field in cc.Range.Fields)
                    {

                            if (field.Type == Word.WdFieldType.wdFieldLink)
                            {
                                string code = field.Code.Text;
                                // 找出第一個被引號包住且副檔名為 xls/xlsx/xlsm 的路徑
                                var match = System.Text.RegularExpressions.Regex.Match(
                                    code,
                                    "\"([^\"]+\\.(?:xls|xlsx|xlsm))\"",
                                    System.Text.RegularExpressions.RegexOptions.IgnoreCase);

                                if (match.Success)
                                {
                                    string oldPathRaw = match.Groups[1].Value; // 可能是已經含有雙反斜線的版本
                                                                               // 若舊路徑與目前路徑（未轉義原始值比較）不同才替換
                                    if (!string.Equals(NormalizePathForCompare(oldPathRaw), _excelFilePath, StringComparison.OrdinalIgnoreCase))
                                    {
                                        // 用 Regex.Replace 指定第一個符合的引號包住的路徑整段替換
                                        string newCode = System.Text.RegularExpressions.Regex.Replace(
                                            code,
                                            "\"([^\"]+\\.(?:xls|xlsx|xlsm))\"",
                                            "\"" + escapedPath + "\"",
                                            System.Text.RegularExpressions.RegexOptions.IgnoreCase
                                        );

                                        field.Code.Text = newCode;
                                        field.Update();
                                        updatedField++;
                                    }
                                }
                            }

                    }
                    cc.LockContents = true;
                }

                MessageBox.Show($"已重新連結之附註數量: {updatedField}", "完成");
            }
            catch (Exception ex)
            {
                MessageBox.Show("發生錯誤：\r\n" + ex.Message);
            }
        }

        // 將實際檔案路徑轉為 LINK 欄位所需格式（雙反斜線）
        private static string EscapePathForLinkField(string path)
        {
            if (string.IsNullOrEmpty(path)) return path;
            // Word LINK 欄位中路徑通常顯示為 C:\\Folder\\File.xlsx
            return path.Replace("\\", "\\\\");
        }

        // 將欄位中可能是已雙反斜線顯示的路徑還原為單反斜線，以便比較
        private static string NormalizePathForCompare(string fieldPath)
        {
            if (string.IsNullOrEmpty(fieldPath)) return fieldPath;
            return fieldPath.Replace("\\\\", "\\");
        }









    }
}
