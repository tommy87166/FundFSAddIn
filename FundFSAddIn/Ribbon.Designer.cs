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
            this.group2 = this.Factory.CreateRibbonGroup();
            this.btnInsert = this.Factory.CreateRibbonButton();
            this.group3 = this.Factory.CreateRibbonGroup();
            this.btnGoToExcel = this.Factory.CreateRibbonButton();
            this.btnUpdatePicture = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.tab.SuspendLayout();
            this.group2.SuspendLayout();
            this.group3.SuspendLayout();
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
            this.tab.Groups.Add(this.group2);
            this.tab.Groups.Add(this.group3);
            this.tab.Label = "基金財報附註自動化";
            this.tab.Name = "tab";
            // 
            // group2
            // 
            this.group2.Items.Add(this.btnInsert);
            this.group2.Label = "插入";
            this.group2.Name = "group2";
            // 
            // btnInsert
            // 
            this.btnInsert.Label = "插入表格段";
            this.btnInsert.Name = "btnInsert";
            this.btnInsert.ShowImage = true;
            this.btnInsert.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnInsert_Click);
            // 
            // group3
            // 
            this.group3.Items.Add(this.btnGoToExcel);
            this.group3.Items.Add(this.btnUpdatePicture);
            this.group3.Label = "group3";
            this.group3.Name = "group3";
            // 
            // btnGoToExcel
            // 
            this.btnGoToExcel.Label = "開啟來源 Excel";
            this.btnGoToExcel.Name = "btnGoToExcel";
            this.btnGoToExcel.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnGoToExcel_Click);
            // 
            // btnUpdatePicture
            // 
            this.btnUpdatePicture.Label = "更新圖片";
            this.btnUpdatePicture.Name = "btnUpdatePicture";
            this.btnUpdatePicture.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnUpdatePicture_Click);
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
            this.group2.ResumeLayout(false);
            this.group2.PerformLayout();
            this.group3.ResumeLayout(false);
            this.group3.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group2;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnInsert;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnGoToExcel;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnUpdatePicture;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group3;
    }

    partial class ThisRibbonCollection
    {
        internal Ribbon Ribbon
        {
            get { return this.GetRibbon<Ribbon>(); }
        }
    }
}
