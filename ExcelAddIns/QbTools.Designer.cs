namespace ExcelAddIns
{
    partial class QbTools : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// 必需的设计器变量。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public QbTools()
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(QbTools));
            this.QbToolBar = this.Factory.CreateRibbonTab();
            this.猪头开发的工具 = this.Factory.CreateRibbonGroup();
            this.Search搜索工作表 = this.Factory.CreateRibbonButton();
            this.QbToolBar.SuspendLayout();
            this.猪头开发的工具.SuspendLayout();
            this.SuspendLayout();
            // 
            // QbToolBar
            // 
            this.QbToolBar.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.QbToolBar.ControlId.OfficeId = "TabHome";
            this.QbToolBar.Groups.Add(this.猪头开发的工具);
            this.QbToolBar.Label = "TabHome";
            this.QbToolBar.Name = "QbToolBar";
            // 
            // 猪头开发的工具
            // 
            this.猪头开发的工具.Items.Add(this.Search搜索工作表);
            this.猪头开发的工具.Label = "猪头开发的工具";
            this.猪头开发的工具.Name = "猪头开发的工具";
            // 
            // Search搜索工作表
            // 
            this.Search搜索工作表.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.Search搜索工作表.Image = ((System.Drawing.Image)(resources.GetObject("Search搜索工作表.Image")));
            this.Search搜索工作表.Label = "Search搜索工作表";
            this.Search搜索工作表.Name = "Search搜索工作表";
            this.Search搜索工作表.ScreenTip = "搜索下面的一大堆工作表";
            this.Search搜索工作表.ShowImage = true;
            this.Search搜索工作表.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.Search搜索工作表_Click);
            // 
            // QbTools
            // 
            this.Name = "QbTools";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.QbToolBar);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.QbTools_Load);
            this.QbToolBar.ResumeLayout(false);
            this.QbToolBar.PerformLayout();
            this.猪头开发的工具.ResumeLayout(false);
            this.猪头开发的工具.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab QbToolBar;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup 猪头开发的工具;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton Search搜索工作表;
    }

    partial class ThisRibbonCollection
    {
        internal QbTools QbTools
        {
            get { return this.GetRibbon<QbTools>(); }
        }
    }
}
