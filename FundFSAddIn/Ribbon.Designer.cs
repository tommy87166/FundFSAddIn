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
            this.group_edit = this.Factory.CreateRibbonGroup();
            this.group2 = this.Factory.CreateRibbonGroup();
            this.group_lock = this.Factory.CreateRibbonGroup();
            this.group_setting = this.Factory.CreateRibbonGroup();
            this.group_version = this.Factory.CreateRibbonGroup();
            this.lbVersion = this.Factory.CreateRibbonLabel();
            this.label1 = this.Factory.CreateRibbonLabel();
            this.btnInsertTable = this.Factory.CreateRibbonButton();
            this.btnInsertText = this.Factory.CreateRibbonButton();
            this.btnUpdateOne = this.Factory.CreateRibbonButton();
            this.btnUpdateAll = this.Factory.CreateRibbonButton();
            this.btnDeleteCC = this.Factory.CreateRibbonButton();
            this.btnGoToExcel = this.Factory.CreateRibbonButton();
            this.btnHideExcel = this.Factory.CreateRibbonButton();
            this.btnLock = this.Factory.CreateRibbonButton();
            this.btnUnlock = this.Factory.CreateRibbonButton();
            this.btnSetExcelFilePath = this.Factory.CreateRibbonButton();
            this.btnRemapOneLink = this.Factory.CreateRibbonButton();
            this.btnRemapLinks = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.tab.SuspendLayout();
            this.group_inserttable.SuspendLayout();
            this.group_edit.SuspendLayout();
            this.group2.SuspendLayout();
            this.group_lock.SuspendLayout();
            this.group_setting.SuspendLayout();
            this.group_version.SuspendLayout();
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
            this.tab.Groups.Add(this.group_edit);
            this.tab.Groups.Add(this.group2);
            this.tab.Groups.Add(this.group_lock);
            this.tab.Groups.Add(this.group_setting);
            this.tab.Groups.Add(this.group_version);
            this.tab.Label = "財報附註工具";
            this.tab.Name = "tab";
            // 
            // group_inserttable
            // 
            this.group_inserttable.Items.Add(this.btnInsertTable);
            this.group_inserttable.Items.Add(this.btnInsertText);
            this.group_inserttable.Label = "插入附註";
            this.group_inserttable.Name = "group_inserttable";
            // 
            // group_edit
            // 
            this.group_edit.Items.Add(this.btnUpdateOne);
            this.group_edit.Items.Add(this.btnUpdateAll);
            this.group_edit.Items.Add(this.btnDeleteCC);
            this.group_edit.Label = "更新附註";
            this.group_edit.Name = "group_edit";
            // 
            // group2
            // 
            this.group2.Items.Add(this.btnGoToExcel);
            this.group2.Items.Add(this.btnHideExcel);
            this.group2.Label = "Excel操作";
            this.group2.Name = "group2";
            // 
            // group_lock
            // 
            this.group_lock.Items.Add(this.btnLock);
            this.group_lock.Items.Add(this.btnUnlock);
            this.group_lock.Label = "鎖定/解鎖";
            this.group_lock.Name = "group_lock";
            // 
            // group_setting
            // 
            this.group_setting.Items.Add(this.btnSetExcelFilePath);
            this.group_setting.Items.Add(this.btnRemapOneLink);
            this.group_setting.Items.Add(this.btnRemapLinks);
            this.group_setting.Label = "指定附註檔";
            this.group_setting.Name = "group_setting";
            // 
            // group_version
            // 
            this.group_version.Items.Add(this.lbVersion);
            this.group_version.Items.Add(this.label1);
            this.group_version.Name = "group_version";
            // 
            // lbVersion
            // 
            this.lbVersion.Label = "版本";
            this.lbVersion.Name = "lbVersion";
            // 
            // label1
            // 
            this.label1.Label = "Made by Tom Pai";
            this.label1.Name = "label1";
            // 
            // btnInsertTable
            // 
            this.btnInsertTable.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnInsertTable.Image = global::FundFSAddIn.Properties.Resources.icon_table;
            this.btnInsertTable.Label = "插入表格";
            this.btnInsertTable.Name = "btnInsertTable";
            this.btnInsertTable.ShowImage = true;
            this.btnInsertTable.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnInsertTable_Click);
            // 
            // btnInsertText
            // 
            this.btnInsertText.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnInsertText.Image = global::FundFSAddIn.Properties.Resources.icon_text;
            this.btnInsertText.Label = "插入文字";
            this.btnInsertText.Name = "btnInsertText";
            this.btnInsertText.ShowImage = true;
            this.btnInsertText.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnInsertText_Click);
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
            // btnUpdateAll
            // 
            this.btnUpdateAll.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnUpdateAll.Image = global::FundFSAddIn.Properties.Resources.icon_refresh;
            this.btnUpdateAll.Label = "更新全部附註";
            this.btnUpdateAll.Name = "btnUpdateAll";
            this.btnUpdateAll.ShowImage = true;
            this.btnUpdateAll.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnUpdateAll_Click);
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
            // btnGoToExcel
            // 
            this.btnGoToExcel.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnGoToExcel.Image = global::FundFSAddIn.Properties.Resources.icon_send;
            this.btnGoToExcel.Label = "開啟附註來源";
            this.btnGoToExcel.Name = "btnGoToExcel";
            this.btnGoToExcel.ShowImage = true;
            this.btnGoToExcel.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnGoToExcel_Click);
            // 
            // btnHideExcel
            // 
            this.btnHideExcel.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnHideExcel.Image = global::FundFSAddIn.Properties.Resources.icon_hide;
            this.btnHideExcel.Label = "隱藏Excel";
            this.btnHideExcel.Name = "btnHideExcel";
            this.btnHideExcel.ShowImage = true;
            this.btnHideExcel.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnHideExcel_Click);
            // 
            // btnLock
            // 
            this.btnLock.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnLock.Image = global::FundFSAddIn.Properties.Resources.icon_lock;
            this.btnLock.Label = "鎖定附註";
            this.btnLock.Name = "btnLock";
            this.btnLock.ShowImage = true;
            this.btnLock.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnLock_Click);
            // 
            // btnUnlock
            // 
            this.btnUnlock.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnUnlock.Image = global::FundFSAddIn.Properties.Resources.icon_unlock;
            this.btnUnlock.Label = "解鎖附註";
            this.btnUnlock.Name = "btnUnlock";
            this.btnUnlock.ShowImage = true;
            this.btnUnlock.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnUnlock_Click);
            // 
            // btnSetExcelFilePath
            // 
            this.btnSetExcelFilePath.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnSetExcelFilePath.Image = global::FundFSAddIn.Properties.Resources.icon_excel;
            this.btnSetExcelFilePath.Label = "設定附註檔";
            this.btnSetExcelFilePath.Name = "btnSetExcelFilePath";
            this.btnSetExcelFilePath.ShowImage = true;
            this.btnSetExcelFilePath.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnSetExcelFilePath_Click);
            // 
            // btnRemapOneLink
            // 
            this.btnRemapOneLink.Image = global::FundFSAddIn.Properties.Resources.icon_link;
            this.btnRemapOneLink.Label = "修復單一附註連結";
            this.btnRemapOneLink.Name = "btnRemapOneLink";
            this.btnRemapOneLink.ShowImage = true;
            this.btnRemapOneLink.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnRemapOneLink_Click);
            // 
            // btnRemapLinks
            // 
            this.btnRemapLinks.Image = global::FundFSAddIn.Properties.Resources.icon_link;
            this.btnRemapLinks.Label = "修復所有附註連結";
            this.btnRemapLinks.Name = "btnRemapLinks";
            this.btnRemapLinks.ShowImage = true;
            this.btnRemapLinks.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnRemapLinks_Click);
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
            this.group_edit.ResumeLayout(false);
            this.group_edit.PerformLayout();
            this.group2.ResumeLayout(false);
            this.group2.PerformLayout();
            this.group_lock.ResumeLayout(false);
            this.group_lock.PerformLayout();
            this.group_setting.ResumeLayout(false);
            this.group_setting.PerformLayout();
            this.group_version.ResumeLayout(false);
            this.group_version.PerformLayout();
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
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnUpdateAll;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnDeleteCC;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group_edit;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group_setting;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnInsertText;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnRemapLinks;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group_lock;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnLock;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnUnlock;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnHideExcel;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group2;
        internal Microsoft.Office.Tools.Ribbon.RibbonLabel lbVersion;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group_version;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnRemapOneLink;
        internal Microsoft.Office.Tools.Ribbon.RibbonLabel label1;
    }

    partial class ThisRibbonCollection
    {
        internal Ribbon Ribbon
        {
            get { return this.GetRibbon<Ribbon>(); }
        }
    }
}
