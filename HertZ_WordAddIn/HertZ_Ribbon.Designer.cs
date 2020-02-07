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
            this.JiuQi = this.Factory.CreateRibbonGroup();
            this.ChangeExcelPath = this.Factory.CreateRibbonButton();
            this.UpdateLink = this.Factory.CreateRibbonButton();
            this.DiscAllLink = this.Factory.CreateRibbonSplitButton();
            this.DiscLink = this.Factory.CreateRibbonButton();
            this.VerGroup = this.Factory.CreateRibbonGroup();
            this.VerInfo = this.Factory.CreateRibbonButton();
            this.HertZTab.SuspendLayout();
            this.Format.SuspendLayout();
            this.JiuQi.SuspendLayout();
            this.VerGroup.SuspendLayout();
            this.SuspendLayout();
            // 
            // HertZTab
            // 
            this.HertZTab.Groups.Add(this.Format);
            this.HertZTab.Groups.Add(this.JiuQi);
            this.HertZTab.Groups.Add(this.VerGroup);
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
            this.TableFormat.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.TableFormat_Click);
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
            // JiuQi
            // 
            this.JiuQi.Items.Add(this.ChangeExcelPath);
            this.JiuQi.Items.Add(this.UpdateLink);
            this.JiuQi.Items.Add(this.DiscAllLink);
            this.JiuQi.Label = "久其";
            this.JiuQi.Name = "JiuQi";
            // 
            // ChangeExcelPath
            // 
            this.ChangeExcelPath.Label = "修改Excel路径";
            this.ChangeExcelPath.Name = "ChangeExcelPath";
            this.ChangeExcelPath.OfficeImageId = "AccessRelinkLists";
            this.ChangeExcelPath.ScreenTip = "修改附注链接的Excel文件路径";
            this.ChangeExcelPath.ShowImage = true;
            this.ChangeExcelPath.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ChangeExcelPath_Click);
            // 
            // UpdateLink
            // 
            this.UpdateLink.Label = "更新全部域";
            this.UpdateLink.Name = "UpdateLink";
            this.UpdateLink.OfficeImageId = "Refresh";
            this.UpdateLink.ScreenTip = "点击刷新所有表格数据";
            this.UpdateLink.ShowImage = true;
            this.UpdateLink.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.UpdateLink_Click);
            // 
            // DiscAllLink
            // 
            this.DiscAllLink.Items.Add(this.DiscLink);
            this.DiscAllLink.Label = "断开全部链接";
            this.DiscAllLink.Name = "DiscAllLink";
            this.DiscAllLink.OfficeImageId = "RecordsDeleteRecord";
            this.DiscAllLink.ScreenTip = "点击断开全部表格的域链接";
            this.DiscAllLink.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.DiscAllLink_Click);
            // 
            // DiscLink
            // 
            this.DiscLink.Label = "断开当前链接";
            this.DiscLink.Name = "DiscLink";
            this.DiscLink.OfficeImageId = "SlideDelete";
            this.DiscLink.ScreenTip = "点击断开当前表格的域链接";
            this.DiscLink.ShowImage = true;
            this.DiscLink.Visible = false;
            this.DiscLink.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.DiscLink_Click);
            // 
            // VerGroup
            // 
            this.VerGroup.Items.Add(this.VerInfo);
            this.VerGroup.Name = "VerGroup";
            // 
            // VerInfo
            // 
            this.VerInfo.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.VerInfo.Image = global::HertZ_WordAddIn.Properties.Resources.HertZ_Logo;
            this.VerInfo.Label = "版本信息";
            this.VerInfo.Name = "VerInfo";
            this.VerInfo.ScreenTip = "点击查看版本信息及使用说明";
            this.VerInfo.ShowImage = true;
            this.VerInfo.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.VerInfo_Click);
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
            this.JiuQi.ResumeLayout(false);
            this.JiuQi.PerformLayout();
            this.VerGroup.ResumeLayout(false);
            this.VerGroup.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab HertZTab;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup Format;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton TextFormat;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton TableFormat;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton SettingFormat;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup VerGroup;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton VerInfo;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup JiuQi;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton ChangeExcelPath;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton UpdateLink;
        internal Microsoft.Office.Tools.Ribbon.RibbonSplitButton DiscAllLink;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton DiscLink;
    }

    partial class ThisRibbonCollection
    {
        internal HertZ_Ribbon HertZ_Ribbon
        {
            get { return this.GetRibbon<HertZ_Ribbon>(); }
        }
    }
}
