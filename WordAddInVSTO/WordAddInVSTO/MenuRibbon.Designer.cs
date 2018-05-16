namespace WordAddInVSTO
{
    partial class MenuRibbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// 必需的设计器变量。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public MenuRibbon()
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(MenuRibbon));
            this.tab1 = this.Factory.CreateRibbonTab();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.sendBtn = this.Factory.CreateRibbonButton();
            this.group2 = this.Factory.CreateRibbonGroup();
            this.detectTBtn = this.Factory.CreateRibbonToggleButton();
            this.tab1.SuspendLayout();
            this.group1.SuspendLayout();
            this.group2.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.Groups.Add(this.group1);
            this.tab1.Groups.Add(this.group2);
            this.tab1.Label = "TextChangeDemo";
            this.tab1.Name = "tab1";
            // 
            // group1
            // 
            this.group1.Items.Add(this.sendBtn);
            this.group1.Label = "功能按钮区域";
            this.group1.Name = "group1";
            // 
            // sendBtn
            // 
            this.sendBtn.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.sendBtn.Image = ((System.Drawing.Image)(resources.GetObject("sendBtn.Image")));
            this.sendBtn.Label = "远程保存";
            this.sendBtn.Name = "sendBtn";
            this.sendBtn.ShowImage = true;
            this.sendBtn.SuperTip = "点击按钮，将当前整个文档内容以文本形式发送到服务端。";
            this.sendBtn.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.sendBtn_Click);
            // 
            // group2
            // 
            this.group2.Items.Add(this.detectTBtn);
            this.group2.Label = "功能开关区域";
            this.group2.Name = "group2";
            // 
            // detectTBtn
            // 
            this.detectTBtn.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.detectTBtn.Image = ((System.Drawing.Image)(resources.GetObject("detectTBtn.Image")));
            this.detectTBtn.Label = "更新开关";
            this.detectTBtn.Name = "detectTBtn";
            this.detectTBtn.ShowImage = true;
            this.detectTBtn.SuperTip = "打开开关，实时监听文章内容的更新和修改，并且将整个句子发送到服务端。";
            this.detectTBtn.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.detectTBtn_Click);
            // 
            // MenuRibbon
            // 
            this.Name = "MenuRibbon";
            this.RibbonType = "Microsoft.Word.Document";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.MenuRibbon_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.group2.ResumeLayout(false);
            this.group2.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group2;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton sendBtn;
        internal Microsoft.Office.Tools.Ribbon.RibbonToggleButton detectTBtn;
    }

    partial class ThisRibbonCollection
    {
        internal MenuRibbon MenuRibbon
        {
            get { return this.GetRibbon<MenuRibbon>(); }
        }
    }
}
