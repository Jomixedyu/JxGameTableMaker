namespace TableMaker
{
    partial class RibbonTab : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// 必需的设计器变量。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public RibbonTab()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

        /// <summary> 
        /// 清理所有正在使用的资源。
        /// </summary>
        /// <param name="disposing">如果应释放托管资源，为 true；否则为 false。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region 组件设计器生成的代码

        /// <summary>
        /// 设计器支持所需的方法 - 不要修改
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InitializeComponent()
        {
            this.tab1 = this.Factory.CreateRibbonTab();
            this.group3 = this.Factory.CreateRibbonGroup();
            this.IsLoadCheck = this.Factory.CreateRibbonCheckBox();
            this.group4 = this.Factory.CreateRibbonGroup();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.group2 = this.Factory.CreateRibbonGroup();
            this.WindowGroup = this.Factory.CreateRibbonGroup();
            this.NewTableBtn = this.Factory.CreateRibbonButton();
            this.NewSheetBtn = this.Factory.CreateRibbonButton();
            this.NewExmTableBtn = this.Factory.CreateRibbonButton();
            this.ErrorCheckBtn = this.Factory.CreateRibbonButton();
            this.ErrorCheckAllBtn = this.Factory.CreateRibbonButton();
            this.ExportBtn = this.Factory.CreateRibbonButton();
            this.QuickSqliteBtn = this.Factory.CreateRibbonButton();
            this.QuickCSVBtn = this.Factory.CreateRibbonButton();
            this.button3 = this.Factory.CreateRibbonButton();
            this.group5 = this.Factory.CreateRibbonGroup();
            this.UndoBtn = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.group3.SuspendLayout();
            this.group4.SuspendLayout();
            this.group1.SuspendLayout();
            this.group2.SuspendLayout();
            this.WindowGroup.SuspendLayout();
            this.group5.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Groups.Add(this.group3);
            this.tab1.Groups.Add(this.group4);
            this.tab1.Groups.Add(this.group5);
            this.tab1.Groups.Add(this.group1);
            this.tab1.Groups.Add(this.group2);
            this.tab1.Groups.Add(this.WindowGroup);
            this.tab1.Label = "TableMaker";
            this.tab1.Name = "tab1";
            // 
            // group3
            // 
            this.group3.Items.Add(this.IsLoadCheck);
            this.group3.Label = "Option";
            this.group3.Name = "group3";
            // 
            // IsLoadCheck
            // 
            this.IsLoadCheck.Checked = true;
            this.IsLoadCheck.Label = "加载";
            this.IsLoadCheck.Name = "IsLoadCheck";
            // 
            // group4
            // 
            this.group4.Items.Add(this.NewTableBtn);
            this.group4.Items.Add(this.NewSheetBtn);
            this.group4.Items.Add(this.NewExmTableBtn);
            this.group4.Label = "新建";
            this.group4.Name = "group4";
            // 
            // group1
            // 
            this.group1.Items.Add(this.ErrorCheckBtn);
            this.group1.Items.Add(this.ErrorCheckAllBtn);
            this.group1.Label = "检查";
            this.group1.Name = "group1";
            // 
            // group2
            // 
            this.group2.Items.Add(this.ExportBtn);
            this.group2.Items.Add(this.QuickSqliteBtn);
            this.group2.Items.Add(this.QuickCSVBtn);
            this.group2.Label = "导出";
            this.group2.Name = "group2";
            // 
            // WindowGroup
            // 
            this.WindowGroup.Items.Add(this.button3);
            this.WindowGroup.Label = "窗口";
            this.WindowGroup.Name = "WindowGroup";
            // 
            // NewTableBtn
            // 
            this.NewTableBtn.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.NewTableBtn.Label = "在当前表新建";
            this.NewTableBtn.Name = "NewTableBtn";
            this.NewTableBtn.OfficeImageId = "AdpNewTable";
            this.NewTableBtn.ShowImage = true;
            this.NewTableBtn.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.NewTableBtn_Click);
            // 
            // NewSheetBtn
            // 
            this.NewSheetBtn.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.NewSheetBtn.Label = "新建工作表";
            this.NewSheetBtn.Name = "NewSheetBtn";
            this.NewSheetBtn.OfficeImageId = "AdpNewTable";
            this.NewSheetBtn.ShowImage = true;
            this.NewSheetBtn.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.NewSheetBtn_Click);
            // 
            // NewExmTableBtn
            // 
            this.NewExmTableBtn.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.NewExmTableBtn.Label = "新建样例工作表";
            this.NewExmTableBtn.Name = "NewExmTableBtn";
            this.NewExmTableBtn.OfficeImageId = "AdpNewTable";
            this.NewExmTableBtn.ShowImage = true;
            this.NewExmTableBtn.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.NewExmTableBtn_Click);
            // 
            // ErrorCheckBtn
            // 
            this.ErrorCheckBtn.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.ErrorCheckBtn.Label = "检查当前页错误";
            this.ErrorCheckBtn.Name = "ErrorCheckBtn";
            this.ErrorCheckBtn.OfficeImageId = "Refresh";
            this.ErrorCheckBtn.ShowImage = true;
            this.ErrorCheckBtn.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ErrorCheckBtn_Click);
            // 
            // ErrorCheckAllBtn
            // 
            this.ErrorCheckAllBtn.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.ErrorCheckAllBtn.Label = "检查所有错误";
            this.ErrorCheckAllBtn.Name = "ErrorCheckAllBtn";
            this.ErrorCheckAllBtn.OfficeImageId = "RefreshAll";
            this.ErrorCheckAllBtn.ShowImage = true;
            this.ErrorCheckAllBtn.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ErrorCheckAllBtn_Click);
            // 
            // ExportBtn
            // 
            this.ExportBtn.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.ExportBtn.Label = "导出";
            this.ExportBtn.Name = "ExportBtn";
            this.ExportBtn.OfficeImageId = "Export";
            this.ExportBtn.ShowImage = true;
            this.ExportBtn.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ExportBtn_Click);
            // 
            // QuickSqliteBtn
            // 
            this.QuickSqliteBtn.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.QuickSqliteBtn.Label = "快速导出Sqlite";
            this.QuickSqliteBtn.Name = "QuickSqliteBtn";
            this.QuickSqliteBtn.OfficeImageId = "AdpStoredProcedureEditSql";
            this.QuickSqliteBtn.ShowImage = true;
            this.QuickSqliteBtn.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.QuickSqliteBtn_Click);
            // 
            // QuickCSVBtn
            // 
            this.QuickCSVBtn.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.QuickCSVBtn.Label = "快速导出CSV";
            this.QuickCSVBtn.Name = "QuickCSVBtn";
            this.QuickCSVBtn.OfficeImageId = "AddCalendarMenu";
            this.QuickCSVBtn.ShowImage = true;
            this.QuickCSVBtn.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.QuickCSVBtn_Click);
            // 
            // button3
            // 
            this.button3.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button3.Label = "检视板";
            this.button3.Name = "button3";
            this.button3.OfficeImageId = "AppendOnly";
            this.button3.ShowImage = true;
            // 
            // group5
            // 
            this.group5.Items.Add(this.UndoBtn);
            this.group5.Label = "操作";
            this.group5.Name = "group5";
            // 
            // UndoBtn
            // 
            this.UndoBtn.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.UndoBtn.Label = "撤回";
            this.UndoBtn.Name = "UndoBtn";
            this.UndoBtn.OfficeImageId = "AdpNewTable";
            this.UndoBtn.ShowImage = true;
            this.UndoBtn.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.UndoBtn_Click);
            // 
            // RibbonTab
            // 
            this.Name = "RibbonTab";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.RibbonTab_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.group3.ResumeLayout(false);
            this.group3.PerformLayout();
            this.group4.ResumeLayout(false);
            this.group4.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.group2.ResumeLayout(false);
            this.group2.PerformLayout();
            this.WindowGroup.ResumeLayout(false);
            this.WindowGroup.PerformLayout();
            this.group5.ResumeLayout(false);
            this.group5.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton ErrorCheckBtn;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton ErrorCheckAllBtn;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group2;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton ExportBtn;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group3;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox IsLoadCheck;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group4;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton NewTableBtn;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton NewSheetBtn;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton NewExmTableBtn;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton QuickSqliteBtn;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton QuickCSVBtn;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup WindowGroup;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button3;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group5;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton UndoBtn;
    }

    partial class ThisRibbonCollection
    {
        internal RibbonTab RibbonTab
        {
            get { return this.GetRibbon<RibbonTab>(); }
        }
    }
}
