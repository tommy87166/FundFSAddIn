using Microsoft.Office.Tools.Ribbon;
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
        //常數
        private string _excelFilePath = null;
        private const string TablePrefix = "表格_";
        private const string TextPrefix = "文字_";

        //共用的 Excel Application / Workbook
        private Excel.Application _sharedExcelApp;
        private Excel.Workbook _sharedWorkbook;

        //載入Ribbon
        private void Ribbon_Load(object sender, RibbonUIEventArgs e)
        {
            UpdateExcelFileNameLabel();
        }

        //更新工作列狀態
        private void UpdateExcelFileNameLabel()
        {
            if (string.IsNullOrEmpty(_excelFilePath) || !System.IO.File.Exists(_excelFilePath))
            {
                group_setting.Label = " ❌  尚未指定附註檔";
                btnInsertTable.Enabled = false;
                btnInsertText.Enabled = false;
                btnGoToExcel.Enabled = false;
                btnUpdateOne.Enabled = false;
                btnUpdateAll.Enabled = false;
                btnDeleteCC.Enabled = false;
                btnRemapLinks.Enabled = false;
                btnHideExcel.Enabled = false;
            }
            else
            {
                var fileName = System.IO.Path.GetFileNameWithoutExtension(_excelFilePath);
                group_setting.Label = " ✔️  已開啟附註檔(" + fileName+")";
                btnInsertTable.Enabled = true;
                btnInsertText .Enabled = true;
                btnGoToExcel.Enabled = true;
                btnUpdateOne.Enabled = true;
                btnUpdateAll.Enabled = true;
                btnDeleteCC.Enabled = true;
                btnRemapLinks.Enabled = true;
                btnHideExcel.Enabled = true;
            }
        }

        //功能-插入表格段
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
                copyTableFromExcel(sheet);
                Word.Document doc = Globals.ThisAddIn.Application.ActiveDocument;
                Word.Range wrange = doc.Application.Selection?.Range ?? doc.Content;
                wrange.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                Word.ContentControl cc = doc.ContentControls.Add(Word.WdContentControlType.wdContentControlRichText, wrange);
                cc.Tag = tag;
                cc.Title = tag;
                pasteTableIntoCC(cc);
                cc.LockContents = true;
                cc.LockContentControl = true;
            }
            catch (Exception ex)
            {
                MessageBox.Show("發生錯誤：\r\n" + ex.Message);
            }
        }

        //功能-插入文字段
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
                copyTextFromExcel(name);
                // 插入 Word 內容控制項並貼上 OLE 物件
                Word.Document doc = Globals.ThisAddIn.Application.ActiveDocument;
                Word.Range wrange = doc.Application.Selection?.Range ?? doc.Content;
                wrange.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                Word.ContentControl cc = doc.ContentControls.Add(Word.WdContentControlType.wdContentControlRichText, wrange);
                cc.Tag = name;
                cc.Title = name;
                pasteTextIntoCC(cc);
                cc.LockContents = true;
                cc.LockContentControl = true;
            }
            catch (Exception ex)
            {
                MessageBox.Show("發生錯誤：\r\n" + ex.Message);
            }
        }

        //功能-開啟附註來源
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

        //功能-更新單一附註
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
                    if (!string.IsNullOrEmpty(cc.Tag) && (cc.Tag.StartsWith(TablePrefix, StringComparison.Ordinal) || cc.Tag.StartsWith(TextPrefix, StringComparison.Ordinal))) {
                        foreach (Word.Field field in cc.Range.Fields) { field.Update(); updatedCount++; }
                    }
                }
                MessageBox.Show($"已更新{updatedCount}個附註。", "完成");
            }
            catch (Exception ex)
            {
                MessageBox.Show("發生錯誤：\r\n" + ex.Message);
            }
        }

        //功能-更新所有附註
        private void btnUpdateAll_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                ValidateExcelPath();
                Word.Document doc = Globals.ThisAddIn.Application.ActiveDocument;
                int updatedCount = 0;
                foreach (Word.ContentControl cc in doc.ContentControls)
                {
                    if (!string.IsNullOrEmpty(cc.Tag) && (cc.Tag.StartsWith(TablePrefix, StringComparison.Ordinal) || cc.Tag.StartsWith(TextPrefix, StringComparison.Ordinal))) {
                        foreach (Word.Field field in cc.Range.Fields) { field.Update(); updatedCount++; }
                    }
                }
                MessageBox.Show($"已更新{updatedCount}個附註。", "完成");
            }
            catch (Exception ex)
            {
                MessageBox.Show("發生錯誤：\r\n" + ex.Message);
            }
        }

        //功能-刪除附註
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

        //功能-鎖定附註
        private void btnLock_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                Word.Document doc = Globals.ThisAddIn.Application.ActiveDocument;
                int locked = 0;
                foreach (Word.ContentControl cc in doc.ContentControls)
                {
                    if (!string.IsNullOrEmpty(cc.Tag) &&
                        (cc.Tag.StartsWith(TablePrefix, StringComparison.Ordinal) ||
                         cc.Tag.StartsWith(TextPrefix, StringComparison.Ordinal)))
                    {
                        cc.LockContents = false;
                        foreach (Word.Field field in cc.Range.Fields) {
                            if (field.LinkFormat != null){
                                field.LinkFormat.AutoUpdate = false;
                            }
                        }
                        cc.LockContentControl = true;
                        cc.LockContents = true;
                        locked++;
                    }
                }
                MessageBox.Show($"已鎖定 {locked} 個附註，並已將其設定為手動更新。", "完成");
            }
            catch (Exception ex)
            {
                MessageBox.Show("發生錯誤：\r\n" + ex.Message, "錯誤");
            }
        }

        //功能-解鎖附註
        private void btnUnlock_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                Word.Document doc = Globals.ThisAddIn.Application.ActiveDocument;
                int unlocked = 0;
                foreach (Word.ContentControl cc in doc.ContentControls)
                {
                    if (!string.IsNullOrEmpty(cc.Tag) &&
                        (cc.Tag.StartsWith(TablePrefix, StringComparison.Ordinal) ||
                         cc.Tag.StartsWith(TextPrefix, StringComparison.Ordinal)))
                    {
                        cc.LockContents = false;
                        foreach (Word.Field field in cc.Range.Fields) {
                            if (field.LinkFormat != null)
                            {
                                field.LinkFormat.AutoUpdate = true;
                            }
                        }
                        cc.LockContentControl = true;
                        unlocked++;
                    }
                }
                MessageBox.Show($"已解鎖 {unlocked} 個附註，並已將其設定為自動更新。", "完成");
            }
            catch (Exception ex)
            {
                MessageBox.Show("發生錯誤：\r\n" + ex.Message, "錯誤");
            }
        }


        //功能-設定來源附註檔
        private void btnSetExcelFilePath_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                var ofd = new OpenFileDialog
                {
                    Title = "選擇 Excel 檔案",
                    Filter = "Excel 檔案|*.xlsx",
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


        //功能-重新映射連結
        private void btnRemapLinks_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                Word.Document doc = Globals.ThisAddIn.Application.ActiveDocument;
                //Excel Related
                ValidateExcelPath();
                var sheets = GetExcelSheetNames(_excelFilePath); //取得全部工作表名稱
                var names = GetExcelTextDefinedNames(_excelFilePath); //取得全部文字定義名稱
                EnsureWorkbook(); // 使用共用 Workbook
                //Counts
                int updatedTable = 0;
                int updatedText = 0;
                foreach (Word.ContentControl cc in doc.ContentControls)
                {
                    try {
                        cc.LockContents = false; //先解鎖
                        //表格段處理
                        if (cc.Tag.StartsWith(TablePrefix, StringComparison.Ordinal)){
                            if (!sheets.Contains(cc.Tag)) { throw new ArgumentException($"找不到工作表-{cc.Tag}"); }
                            //複製表格段
                            copyTableFromExcel(cc.Tag);
                            pasteTableIntoCC(cc);
                            updatedTable++;
                        }
                        else if (cc.Tag.StartsWith(TextPrefix, StringComparison.Ordinal))
                        {
                            if (!names.Contains(cc.Tag)) { throw new ArgumentException($"找不到文字段-{cc.Tag}"); }
                            copyTextFromExcel(cc.Tag);
                            pasteTextIntoCC(cc);
                            updatedText++;
                        }
                    }
                    catch (Exception ex){
                        cc.Range.Delete();
                        cc.Range.Text = "發生錯誤: " + ex.Message;
                    }
                    finally {
                        cc.LockContents = true; //先解鎖
                        cc.LockContentControl = true; //鎖定內容控制項
                    }
                }

                MessageBox.Show($"已重新連結之表格數量: {updatedTable} 已重新連結之文字數量: {updatedText}", "完成");
            }
            catch (Exception ex)
            {
                MessageBox.Show("發生錯誤：\r\n" + ex.Message);
            }
        }
        
        //Helper-自Excel中複製表格
        private void copyTableFromExcel (String sheet)
        {
            EnsureWorkbook(); // 使用共用 Workbook
            Excel.Worksheet ws = null;
            ws = (Excel.Worksheet) _sharedWorkbook.Sheets[sheet];
            string printArea = ws.PageSetup.PrintArea;
            if (string.IsNullOrWhiteSpace(printArea))
                throw new Exception("該工作表未設定列印範圍");
            Excel.Range rng = ws.Range[printArea];
            rng.Copy();
            ReleaseCom(ws);
        }

        //Helper-自Excel中複製文字
        private void copyTextFromExcel(String name)
        {
            // 取得 Excel 範圍
            EnsureWorkbook();
            Excel.Name excelName = null;
            Excel.Range range = null;
            foreach (Excel.Name n in _sharedWorkbook.Names)
            {
                if (n != null && n.Name == name)
                {
                    excelName = n;
                    range = n.RefersToRange;
                    break;
                }
            }
            range.Copy();
            ReleaseCom(range);
            ReleaseCom(excelName);
        }

        //Helper-將表格段貼入CC中
        private void pasteTableIntoCC(Word.ContentControl cc)
        {
            cc.Range.Delete();
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
        }

        //Helper-將文字段貼入CC中
        private void pasteTextIntoCC(Word.ContentControl cc) {
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
        }

        //Helper-建立或取得共用 Workbook
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

        //釋放共用資源
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

        //Helper-取得帶有"表格_"前綴的工作表名稱
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
        
        //Helper-取得帶有"文字_"前綴的範圍名稱
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

        //Helper-顯示表格段選擇對話框
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

        //Helper-顯示文字段選擇對話框
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

        //Helper-驗證 Excel 路徑是否有效
        private void ValidateExcelPath()
        {
            if (string.IsNullOrEmpty(_excelFilePath) || !System.IO.File.Exists(_excelFilePath))
                throw new Exception("未指定附註檔");
        }

        private void btnHideExcel_Click(object sender, RibbonControlEventArgs e)
        {
            if (_sharedExcelApp != null)
            {
                try
                {
                    _sharedExcelApp.Visible = false;
                }
                catch (Exception ex)
                {
                    MessageBox.Show("發生錯誤：\r\n" + ex.Message);
                }
            }
        }
    }
}
