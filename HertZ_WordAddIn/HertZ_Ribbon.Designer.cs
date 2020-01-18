namespace HertZ_WordAddIn
{
    partial class HertZ_Ribbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// 必需的设计器变量。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public HertZ_Ribbon()
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
            this.HertZTab = this.Factory.CreateRibbonTab();
            this.Format = this.Factory.CreateRibbonGroup();
            this.TextFormat = this.Factory.CreateRibbonButton();
            this.TableFormat = this.Factory.CreateRibbonButton();
            this.SettingFormat = this.Factory.CreateRibbonButton();
            this.HertZTab.SuspendLayout();
            this.Format.SuspendLayout();
            this.SuspendLayout();
            // 
            // HertZTab
            // 
            this.HertZTab.Groups.Add(this.Format);
            this.HertZTab.Label = "HertZ";
            this.HertZTab.Name = "HertZTab";
            this.HertZTab.Position = this.Factory.RibbonPosition.AfterOfficeId("TabDeveloper");
            // 
            // Format
            // 
            this.Format.Items.Add(this.TextFormat);
            this.Format.Items.Add(this.TableFormat);
            this.Format.Items.Add(this.SettingFormat);
            this.Format.Label = "格式调整";
            this.Format.Name = "Format";
            // 
            // TextFormat
            // 
            this.TextFormat.Label = "调整文本格式";
            this.TextFormat.Name = "TextFormat";
            this.TextFormat.OfficeImageId = "PivotTableLayoutShowInCompactForm";
            this.TextFormat.ScreenTip = "按照设置好的格式调整当前文档";
            this.TextFormat.ShowImage = true;
            this.TextFormat.SuperTip = "请在使用前先设置好默认样式";
            this.TextFormat.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.TextFormat_Click);
            // 
            // TableFormat
            // 
            this.TableFormat.Label = "调整表格格式";
            this.TableFormat.Name = "TableFormat";
            this.TableFormat.OfficeImageId = "CreateQueryFromWizard";
            this.TableFormat.ScreenTip = "按照设置好的格式调整当前表格格式";
            this.TableFormat.ShowImage = true;
            this.TableFormat.SuperTip = "请在使用前先设置好默认样式";
            // 
            // SettingFormat
            // 
            this.SettingFormat.Label = "设置默认样式";
            this.SettingFormat.Name = "SettingFormat";
            this.SettingFormat.OfficeImageId = "AddInManager";
            this.SettingFormat.ScreenTip = "点击设置默认样式";
            this.SettingFormat.ShowImage = true;
            this.SettingFormat.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.SettingFormat_Click);
            // 
            // HertZ_Ribbon
            // 
            this.Name = "HertZ_Ribbon";
            this.RibbonType = "Microsoft.Word.Document";
            this.Tabs.Add(this.HertZTab);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.HertZ_Ribbon_Load);
            this.HertZTab.ResumeLayout(false);
            this.HertZTab.PerformLayout();
            this.Format.ResumeLayout(false);
            this.Format.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab HertZTab;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup Format;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton TextFormat;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton TableFormat;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton SettingFormat;
    }

    partial class ThisRibbonCollection
    {
        internal HertZ_Ribbon HertZ_Ribbon
        {
            get { return this.GetRibbon<HertZ_Ribbon>(); }
        }
    }
}
