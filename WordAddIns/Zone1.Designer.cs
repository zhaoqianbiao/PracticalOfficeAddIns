namespace WordAddIns
{
    partial class Zone1 : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// 必需的设计器变量。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public Zone1()
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Zone1));
            this.QbTools = this.Factory.CreateRibbonTab();
            this.Group1 = this.Factory.CreateRibbonGroup();
            this.OpenInFolder = this.Factory.CreateRibbonButton();
            this.QbTools.SuspendLayout();
            this.Group1.SuspendLayout();
            this.SuspendLayout();
            // 
            // QbTools
            // 
            this.QbTools.Groups.Add(this.Group1);
            this.QbTools.Label = "QbTools";
            this.QbTools.Name = "QbTools";
            // 
            // Group1
            // 
            this.Group1.Items.Add(this.OpenInFolder);
            this.Group1.Label = "Group1";
            this.Group1.Name = "Group1";
            // 
            // OpenInFolder
            // 
            this.OpenInFolder.Image = ((System.Drawing.Image)(resources.GetObject("OpenInFolder.Image")));
            this.OpenInFolder.Label = "OpenInFolder";
            this.OpenInFolder.Name = "OpenInFolder";
            this.OpenInFolder.ShowImage = true;
            this.OpenInFolder.SuperTip = "OpenDocxInFolder在文件夹中显示";
            this.OpenInFolder.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.OpenInFolder_Click);
            // 
            // Zone1
            // 
            this.Name = "Zone1";
            this.RibbonType = "Microsoft.Word.Document";
            this.Tabs.Add(this.QbTools);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Zone1_Load);
            this.QbTools.ResumeLayout(false);
            this.QbTools.PerformLayout();
            this.Group1.ResumeLayout(false);
            this.Group1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab QbTools;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup Group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton OpenInFolder;
    }

    partial class ThisRibbonCollection
    {
        internal Zone1 Zone1
        {
            get { return this.GetRibbon<Zone1>(); }
        }
    }
}
