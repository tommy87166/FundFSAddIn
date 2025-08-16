namespace FundFSAddIn
{
    partial class Ribbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// 設計工具所需的變數。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public Ribbon()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

        /// <summary> 
        /// 清除任何使用中的資源。
        /// </summary>
        /// <param name="disposing">如果應該處置 Managed 資源則為 true，否則為 false。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region 元件設計工具產生的程式碼

        /// <summary>
        /// 此為設計工具支援所需的方法 - 請勿使用程式碼編輯器修改
        /// 這個方法的內容。
        /// </summary>
        private void InitializeComponent()
        {
            this.tab1 = this.Factory.CreateRibbonTab();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.tab = this.Factory.CreateRibbonTab();
            this.group_inserttable = this.Factory.CreateRibbonGroup();
            this.group_opensource = this.Factory.CreateRibbonGroup();
            this.group_edit = this.Factory.CreateRibbonGroup();
            this.group_setting = this.Factory.CreateRibbonGroup();
            this.group_display = this.Factory.CreateRibbonGroup();
            this.lblExcelFileName = this.Factory.CreateRibbonLabel();
            this.btnInsertTable = this.Factory.CreateRibbonButton();
            this.btnInsertText = this.Factory.CreateRibbonButton();
            this.btnGoToExcel = this.Factory.CreateRibbonButton();
            this.btnUpdateOne = this.Factory.CreateRibbonButton();
            this.btnDeleteCC = this.Factory.CreateRibbonButton();
            this.btnSetExcelFilePath = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.tab.SuspendLayout();
            this.group_inserttable.SuspendLayout();
            this.group_opensource.SuspendLayout();
            this.group_edit.SuspendLayout();
            this.group_setting.SuspendLayout();
            this.group_display.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Groups.Add(this.group1);
            this.tab1.Label = "TabAddIns";
            this.tab1.Name = "tab1";
            // 
            // group1
            // 
            this.group1.Label = "group1";
            this.group1.Name = "group1";
            // 
            // tab
            // 
            this.tab.Groups.Add(this.group_inserttable);
            this.tab.Groups.Add(this.group_opensource);
            this.tab.Groups.Add(this.group_edit);
            this.tab.Groups.Add(this.group_setting);
            this.tab.Groups.Add(this.group_display);
            this.tab.Label = "基金財報附註";
            this.tab.Name = "tab";
            // 
            // group_inserttable
            // 
            this.group_inserttable.Items.Add(this.btnInsertTable);
            this.group_inserttable.Items.Add(this.btnInsertText);
            this.group_inserttable.Name = "group_inserttable";
            // 
            // group_opensource
            // 
            this.group_opensource.Items.Add(this.btnGoToExcel);
            this.group_opensource.Name = "group_opensource";
            // 
            // group_edit
            // 
            this.group_edit.Items.Add(this.btnUpdateOne);
            this.group_edit.Items.Add(this.btnDeleteCC);
            this.group_edit.Name = "group_edit";
            // 
            // group_setting
            // 
            this.group_setting.Items.Add(this.btnSetExcelFilePath);
            this.group_setting.Name = "group_setting";
            // 
            // group_display
            // 
            this.group_display.Items.Add(this.lblExcelFileName);
            this.group_display.Name = "group_display";
            // 
            // lblExcelFileName
            // 
            this.lblExcelFileName.Label = "尚未指定附註檔";
            this.lblExcelFileName.Name = "lblExcelFileName";
            // 
            // btnInsertTable
            // 
            this.btnInsertTable.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnInsertTable.Image = global::FundFSAddIn.Properties.Resources.icon_table;
            this.btnInsertTable.Label = "插入表格段";
            this.btnInsertTable.Name = "btnInsertTable";
            this.btnInsertTable.ShowImage = true;
            this.btnInsertTable.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnInsertTable_Click);
            // 
            // btnInsertText
            // 
            this.btnInsertText.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnInsertText.Image = global::FundFSAddIn.Properties.Resources.icon_text;
            this.btnInsertText.Label = "插入文字段";
            this.btnInsertText.Name = "btnInsertText";
            this.btnInsertText.ShowImage = true;
            this.btnInsertText.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnInsertText_Click);
            // 
            // btnGoToExcel
            // 
            this.btnGoToExcel.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnGoToExcel.Image = global::FundFSAddIn.Properties.Resources.icon_send;
            this.btnGoToExcel.Label = "開啟來源附註";
            this.btnGoToExcel.Name = "btnGoToExcel";
            this.btnGoToExcel.ShowImage = true;
            this.btnGoToExcel.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnGoToExcel_Click);
            // 
            // btnUpdateOne
            // 
            this.btnUpdateOne.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnUpdateOne.Image = global::FundFSAddIn.Properties.Resources.icon_refresh_1;
            this.btnUpdateOne.Label = "更新單一附註";
            this.btnUpdateOne.Name = "btnUpdateOne";
            this.btnUpdateOne.ShowImage = true;
            this.btnUpdateOne.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnUpdateOne_Click);
            // 
            // btnDeleteCC
            // 
            this.btnDeleteCC.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnDeleteCC.Image = global::FundFSAddIn.Properties.Resources.icon_trash;
            this.btnDeleteCC.Label = "刪除附註";
            this.btnDeleteCC.Name = "btnDeleteCC";
            this.btnDeleteCC.ShowImage = true;
            this.btnDeleteCC.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnDeleteCC_Click);
            // 
            // btnSetExcelFilePath
            // 
            this.btnSetExcelFilePath.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnSetExcelFilePath.Image = global::FundFSAddIn.Properties.Resources.icon_sheet;
            this.btnSetExcelFilePath.Label = "設定附註檔";
            this.btnSetExcelFilePath.Name = "btnSetExcelFilePath";
            this.btnSetExcelFilePath.ShowImage = true;
            this.btnSetExcelFilePath.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnSetExcelFilePath_Click);
            // 
            // Ribbon
            // 
            this.Name = "Ribbon";
            this.RibbonType = "Microsoft.Word.Document";
            this.Tabs.Add(this.tab1);
            this.Tabs.Add(this.tab);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.tab.ResumeLayout(false);
            this.tab.PerformLayout();
            this.group_inserttable.ResumeLayout(false);
            this.group_inserttable.PerformLayout();
            this.group_opensource.ResumeLayout(false);
            this.group_opensource.PerformLayout();
            this.group_edit.ResumeLayout(false);
            this.group_edit.PerformLayout();
            this.group_setting.ResumeLayout(false);
            this.group_setting.PerformLayout();
            this.group_display.ResumeLayout(false);
            this.group_display.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group_inserttable;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnInsertTable;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnSetExcelFilePath;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnGoToExcel;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnUpdateOne;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnDeleteCC;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group_edit;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group_setting;
        internal Microsoft.Office.Tools.Ribbon.RibbonLabel lblExcelFileName;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group_display;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnInsertText;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group_opensource;
    }

    partial class ThisRibbonCollection
    {
        internal Ribbon Ribbon
        {
            get { return this.GetRibbon<Ribbon>(); }
        }
    }
}
