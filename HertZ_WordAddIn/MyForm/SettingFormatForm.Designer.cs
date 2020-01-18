namespace HertZ_WordAddIn
{
    partial class SettingFormatForm
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(SettingFormatForm));
            this.SettingFormat = new System.Windows.Forms.TabControl();
            this.PageSetting = new System.Windows.Forms.TabPage();
            this.PageFormat = new System.Windows.Forms.GroupBox();
            this.FooterMargin = new System.Windows.Forms.NumericUpDown();
            this.FooterMarginL = new System.Windows.Forms.Label();
            this.HeadMargin = new System.Windows.Forms.NumericUpDown();
            this.PageFormatCheck = new System.Windows.Forms.CheckBox();
            this.HeadMarginL = new System.Windows.Forms.Label();
            this.PageMargin = new System.Windows.Forms.GroupBox();
            this.PageMarginCheck = new System.Windows.Forms.CheckBox();
            this.RightMargin = new System.Windows.Forms.NumericUpDown();
            this.RightMarginL = new System.Windows.Forms.Label();
            this.LeftMargin = new System.Windows.Forms.NumericUpDown();
            this.LeftMarginL = new System.Windows.Forms.Label();
            this.BottomMargin = new System.Windows.Forms.NumericUpDown();
            this.BottomMarginL = new System.Windows.Forms.Label();
            this.TopMargin = new System.Windows.Forms.NumericUpDown();
            this.TopMarginL = new System.Windows.Forms.Label();
            this.TextFormat = new System.Windows.Forms.TabPage();
            this.TitleGroup = new System.Windows.Forms.GroupBox();
            this.RowSpace = new System.Windows.Forms.ComboBox();
            this.RowSpaceL = new System.Windows.Forms.Label();
            this.AfterMainBody = new System.Windows.Forms.NumericUpDown();
            this.AfterMainBodyL = new System.Windows.Forms.Label();
            this.BeforeMainBody = new System.Windows.Forms.NumericUpDown();
            this.BeforeMainBodyL = new System.Windows.Forms.Label();
            this.FirstTitleCheck = new System.Windows.Forms.CheckBox();
            this.TitleIndentCkeck = new System.Windows.Forms.CheckBox();
            this.TitleGroupCheck = new System.Windows.Forms.CheckBox();
            this.FontGroup = new System.Windows.Forms.GroupBox();
            this.FontSize = new System.Windows.Forms.ComboBox();
            this.FontSizeL = new System.Windows.Forms.Label();
            this.NumFontL = new System.Windows.Forms.Label();
            this.NumFont = new System.Windows.Forms.ComboBox();
            this.CnFontL = new System.Windows.Forms.Label();
            this.CnFont = new System.Windows.Forms.ComboBox();
            this.FontGroupCheck = new System.Windows.Forms.CheckBox();
            this.TableFormat = new System.Windows.Forms.TabPage();
            this.comboBox1 = new System.Windows.Forms.ComboBox();
            this.label1 = new System.Windows.Forms.Label();
            this.checkBox2 = new System.Windows.Forms.CheckBox();
            this.checkBox1 = new System.Windows.Forms.CheckBox();
            this.SaveButton = new System.Windows.Forms.Button();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.checkBox3 = new System.Windows.Forms.CheckBox();
            this.groupBox3 = new System.Windows.Forms.GroupBox();
            this.checkBox4 = new System.Windows.Forms.CheckBox();
            this.SettingFormat.SuspendLayout();
            this.PageSetting.SuspendLayout();
            this.PageFormat.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.FooterMargin)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.HeadMargin)).BeginInit();
            this.PageMargin.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.RightMargin)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.LeftMargin)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.BottomMargin)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.TopMargin)).BeginInit();
            this.TextFormat.SuspendLayout();
            this.TitleGroup.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.AfterMainBody)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.BeforeMainBody)).BeginInit();
            this.FontGroup.SuspendLayout();
            this.TableFormat.SuspendLayout();
            this.groupBox1.SuspendLayout();
            this.groupBox3.SuspendLayout();
            this.SuspendLayout();
            // 
            // SettingFormat
            // 
            this.SettingFormat.Controls.Add(this.PageSetting);
            this.SettingFormat.Controls.Add(this.TextFormat);
            this.SettingFormat.Controls.Add(this.TableFormat);
            this.SettingFormat.Dock = System.Windows.Forms.DockStyle.Top;
            this.SettingFormat.Font = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.SettingFormat.Location = new System.Drawing.Point(0, 0);
            this.SettingFormat.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.SettingFormat.Name = "SettingFormat";
            this.SettingFormat.SelectedIndex = 0;
            this.SettingFormat.Size = new System.Drawing.Size(674, 535);
            this.SettingFormat.TabIndex = 1;
            // 
            // PageSetting
            // 
            this.PageSetting.Controls.Add(this.PageFormat);
            this.PageSetting.Controls.Add(this.PageMargin);
            this.PageSetting.Location = new System.Drawing.Point(8, 45);
            this.PageSetting.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.PageSetting.Name = "PageSetting";
            this.PageSetting.Padding = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.PageSetting.Size = new System.Drawing.Size(658, 482);
            this.PageSetting.TabIndex = 0;
            this.PageSetting.Text = "页面设置";
            this.PageSetting.UseVisualStyleBackColor = true;
            // 
            // PageFormat
            // 
            this.PageFormat.Controls.Add(this.FooterMargin);
            this.PageFormat.Controls.Add(this.FooterMarginL);
            this.PageFormat.Controls.Add(this.HeadMargin);
            this.PageFormat.Controls.Add(this.PageFormatCheck);
            this.PageFormat.Controls.Add(this.HeadMarginL);
            this.PageFormat.Dock = System.Windows.Forms.DockStyle.Top;
            this.PageFormat.Location = new System.Drawing.Point(3, 205);
            this.PageFormat.Name = "PageFormat";
            this.PageFormat.Size = new System.Drawing.Size(652, 170);
            this.PageFormat.TabIndex = 1;
            this.PageFormat.TabStop = false;
            // 
            // FooterMargin
            // 
            this.FooterMargin.DecimalPlaces = 2;
            this.FooterMargin.Font = new System.Drawing.Font("微软雅黑", 7.875F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.FooterMargin.Increment = new decimal(new int[] {
            1,
            0,
            0,
            65536});
            this.FooterMargin.Location = new System.Drawing.Point(445, 71);
            this.FooterMargin.Maximum = new decimal(new int[] {
            20,
            0,
            0,
            0});
            this.FooterMargin.Name = "FooterMargin";
            this.FooterMargin.Size = new System.Drawing.Size(100, 35);
            this.FooterMargin.TabIndex = 4;
            // 
            // FooterMarginL
            // 
            this.FooterMarginL.AutoSize = true;
            this.FooterMarginL.Location = new System.Drawing.Point(352, 73);
            this.FooterMarginL.Name = "FooterMarginL";
            this.FooterMarginL.Size = new System.Drawing.Size(86, 31);
            this.FooterMarginL.TabIndex = 3;
            this.FooterMarginL.Text = "页脚：";
            // 
            // HeadMargin
            // 
            this.HeadMargin.DecimalPlaces = 2;
            this.HeadMargin.Font = new System.Drawing.Font("微软雅黑", 7.875F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.HeadMargin.Increment = new decimal(new int[] {
            1,
            0,
            0,
            65536});
            this.HeadMargin.Location = new System.Drawing.Point(175, 71);
            this.HeadMargin.Maximum = new decimal(new int[] {
            20,
            0,
            0,
            0});
            this.HeadMargin.Name = "HeadMargin";
            this.HeadMargin.Size = new System.Drawing.Size(100, 35);
            this.HeadMargin.TabIndex = 2;
            // 
            // PageFormatCheck
            // 
            this.PageFormatCheck.AutoSize = true;
            this.PageFormatCheck.BackColor = System.Drawing.Color.White;
            this.PageFormatCheck.Location = new System.Drawing.Point(3, 0);
            this.PageFormatCheck.Name = "PageFormatCheck";
            this.PageFormatCheck.Size = new System.Drawing.Size(94, 35);
            this.PageFormatCheck.TabIndex = 1;
            this.PageFormatCheck.Text = "版式";
            this.PageFormatCheck.UseVisualStyleBackColor = false;
            // 
            // HeadMarginL
            // 
            this.HeadMarginL.AutoSize = true;
            this.HeadMarginL.Location = new System.Drawing.Point(82, 73);
            this.HeadMarginL.Name = "HeadMarginL";
            this.HeadMarginL.Size = new System.Drawing.Size(86, 31);
            this.HeadMarginL.TabIndex = 0;
            this.HeadMarginL.Text = "页眉：";
            // 
            // PageMargin
            // 
            this.PageMargin.Controls.Add(this.PageMarginCheck);
            this.PageMargin.Controls.Add(this.RightMargin);
            this.PageMargin.Controls.Add(this.RightMarginL);
            this.PageMargin.Controls.Add(this.LeftMargin);
            this.PageMargin.Controls.Add(this.LeftMarginL);
            this.PageMargin.Controls.Add(this.BottomMargin);
            this.PageMargin.Controls.Add(this.BottomMarginL);
            this.PageMargin.Controls.Add(this.TopMargin);
            this.PageMargin.Controls.Add(this.TopMarginL);
            this.PageMargin.Dock = System.Windows.Forms.DockStyle.Top;
            this.PageMargin.Location = new System.Drawing.Point(3, 4);
            this.PageMargin.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.PageMargin.Name = "PageMargin";
            this.PageMargin.Padding = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.PageMargin.Size = new System.Drawing.Size(652, 201);
            this.PageMargin.TabIndex = 0;
            this.PageMargin.TabStop = false;
            this.PageMargin.UseWaitCursor = true;
            // 
            // PageMarginCheck
            // 
            this.PageMarginCheck.AutoSize = true;
            this.PageMarginCheck.BackColor = System.Drawing.Color.White;
            this.PageMarginCheck.Location = new System.Drawing.Point(3, 3);
            this.PageMarginCheck.Name = "PageMarginCheck";
            this.PageMarginCheck.Size = new System.Drawing.Size(118, 35);
            this.PageMarginCheck.TabIndex = 8;
            this.PageMarginCheck.Text = "页边距";
            this.PageMarginCheck.UseVisualStyleBackColor = false;
            this.PageMarginCheck.UseWaitCursor = true;
            // 
            // RightMargin
            // 
            this.RightMargin.DecimalPlaces = 2;
            this.RightMargin.Font = new System.Drawing.Font("微软雅黑", 7.875F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.RightMargin.Increment = new decimal(new int[] {
            1,
            0,
            0,
            65536});
            this.RightMargin.Location = new System.Drawing.Point(444, 119);
            this.RightMargin.Maximum = new decimal(new int[] {
            20,
            0,
            0,
            0});
            this.RightMargin.Name = "RightMargin";
            this.RightMargin.Size = new System.Drawing.Size(100, 35);
            this.RightMargin.TabIndex = 7;
            this.RightMargin.UseWaitCursor = true;
            // 
            // RightMarginL
            // 
            this.RightMarginL.AutoSize = true;
            this.RightMarginL.Location = new System.Drawing.Point(376, 121);
            this.RightMarginL.Name = "RightMarginL";
            this.RightMarginL.Size = new System.Drawing.Size(62, 31);
            this.RightMarginL.TabIndex = 6;
            this.RightMarginL.Text = "右：";
            this.RightMarginL.UseWaitCursor = true;
            // 
            // LeftMargin
            // 
            this.LeftMargin.DecimalPlaces = 2;
            this.LeftMargin.Font = new System.Drawing.Font("微软雅黑", 7.875F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.LeftMargin.Increment = new decimal(new int[] {
            1,
            0,
            0,
            65536});
            this.LeftMargin.Location = new System.Drawing.Point(174, 119);
            this.LeftMargin.Maximum = new decimal(new int[] {
            20,
            0,
            0,
            0});
            this.LeftMargin.Name = "LeftMargin";
            this.LeftMargin.Size = new System.Drawing.Size(100, 35);
            this.LeftMargin.TabIndex = 5;
            this.LeftMargin.UseWaitCursor = true;
            // 
            // LeftMarginL
            // 
            this.LeftMarginL.AutoSize = true;
            this.LeftMarginL.Location = new System.Drawing.Point(106, 121);
            this.LeftMarginL.Name = "LeftMarginL";
            this.LeftMarginL.Size = new System.Drawing.Size(62, 31);
            this.LeftMarginL.TabIndex = 4;
            this.LeftMarginL.Text = "左：";
            this.LeftMarginL.UseWaitCursor = true;
            // 
            // BottomMargin
            // 
            this.BottomMargin.DecimalPlaces = 2;
            this.BottomMargin.Font = new System.Drawing.Font("微软雅黑", 7.875F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.BottomMargin.Increment = new decimal(new int[] {
            1,
            0,
            0,
            65536});
            this.BottomMargin.Location = new System.Drawing.Point(444, 52);
            this.BottomMargin.Maximum = new decimal(new int[] {
            20,
            0,
            0,
            0});
            this.BottomMargin.Name = "BottomMargin";
            this.BottomMargin.Size = new System.Drawing.Size(100, 35);
            this.BottomMargin.TabIndex = 3;
            this.BottomMargin.UseWaitCursor = true;
            // 
            // BottomMarginL
            // 
            this.BottomMarginL.AutoSize = true;
            this.BottomMarginL.Location = new System.Drawing.Point(376, 54);
            this.BottomMarginL.Name = "BottomMarginL";
            this.BottomMarginL.Size = new System.Drawing.Size(62, 31);
            this.BottomMarginL.TabIndex = 2;
            this.BottomMarginL.Text = "下：";
            this.BottomMarginL.UseWaitCursor = true;
            // 
            // TopMargin
            // 
            this.TopMargin.DecimalPlaces = 2;
            this.TopMargin.Font = new System.Drawing.Font("微软雅黑", 7.875F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.TopMargin.Increment = new decimal(new int[] {
            1,
            0,
            0,
            65536});
            this.TopMargin.Location = new System.Drawing.Point(174, 52);
            this.TopMargin.Maximum = new decimal(new int[] {
            20,
            0,
            0,
            0});
            this.TopMargin.Name = "TopMargin";
            this.TopMargin.Size = new System.Drawing.Size(100, 35);
            this.TopMargin.TabIndex = 1;
            this.TopMargin.UseWaitCursor = true;
            // 
            // TopMarginL
            // 
            this.TopMarginL.AutoSize = true;
            this.TopMarginL.Location = new System.Drawing.Point(106, 54);
            this.TopMarginL.Name = "TopMarginL";
            this.TopMarginL.Size = new System.Drawing.Size(62, 31);
            this.TopMarginL.TabIndex = 0;
            this.TopMarginL.Text = "上：";
            this.TopMarginL.UseWaitCursor = true;
            // 
            // TextFormat
            // 
            this.TextFormat.Controls.Add(this.TitleGroup);
            this.TextFormat.Controls.Add(this.FontGroup);
            this.TextFormat.Location = new System.Drawing.Point(8, 45);
            this.TextFormat.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.TextFormat.Name = "TextFormat";
            this.TextFormat.Padding = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.TextFormat.Size = new System.Drawing.Size(658, 482);
            this.TextFormat.TabIndex = 1;
            this.TextFormat.Text = "文字排版";
            this.TextFormat.UseVisualStyleBackColor = true;
            // 
            // TitleGroup
            // 
            this.TitleGroup.Controls.Add(this.RowSpace);
            this.TitleGroup.Controls.Add(this.RowSpaceL);
            this.TitleGroup.Controls.Add(this.AfterMainBody);
            this.TitleGroup.Controls.Add(this.AfterMainBodyL);
            this.TitleGroup.Controls.Add(this.BeforeMainBody);
            this.TitleGroup.Controls.Add(this.BeforeMainBodyL);
            this.TitleGroup.Controls.Add(this.FirstTitleCheck);
            this.TitleGroup.Controls.Add(this.TitleIndentCkeck);
            this.TitleGroup.Controls.Add(this.TitleGroupCheck);
            this.TitleGroup.Dock = System.Windows.Forms.DockStyle.Top;
            this.TitleGroup.Location = new System.Drawing.Point(3, 194);
            this.TitleGroup.Name = "TitleGroup";
            this.TitleGroup.Size = new System.Drawing.Size(652, 283);
            this.TitleGroup.TabIndex = 1;
            this.TitleGroup.TabStop = false;
            // 
            // RowSpace
            // 
            this.RowSpace.Font = new System.Drawing.Font("微软雅黑", 7.875F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.RowSpace.FormattingEnabled = true;
            this.RowSpace.Location = new System.Drawing.Point(376, 218);
            this.RowSpace.Name = "RowSpace";
            this.RowSpace.Size = new System.Drawing.Size(180, 36);
            this.RowSpace.TabIndex = 8;
            // 
            // RowSpaceL
            // 
            this.RowSpaceL.AutoSize = true;
            this.RowSpaceL.Location = new System.Drawing.Point(370, 172);
            this.RowSpaceL.Name = "RowSpaceL";
            this.RowSpaceL.Size = new System.Drawing.Size(110, 31);
            this.RowSpaceL.TabIndex = 7;
            this.RowSpaceL.Text = "行间距：";
            // 
            // AfterMainBody
            // 
            this.AfterMainBody.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.AfterMainBody.DecimalPlaces = 1;
            this.AfterMainBody.Font = new System.Drawing.Font("微软雅黑", 7.875F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.AfterMainBody.Increment = new decimal(new int[] {
            5,
            0,
            0,
            65536});
            this.AfterMainBody.Location = new System.Drawing.Point(190, 218);
            this.AfterMainBody.Maximum = new decimal(new int[] {
            50,
            0,
            0,
            0});
            this.AfterMainBody.Name = "AfterMainBody";
            this.AfterMainBody.Size = new System.Drawing.Size(90, 35);
            this.AfterMainBody.TabIndex = 6;
            // 
            // AfterMainBodyL
            // 
            this.AfterMainBodyL.AutoSize = true;
            this.AfterMainBodyL.Location = new System.Drawing.Point(68, 220);
            this.AfterMainBodyL.Name = "AfterMainBodyL";
            this.AfterMainBodyL.Size = new System.Drawing.Size(126, 31);
            this.AfterMainBodyL.TabIndex = 5;
            this.AfterMainBodyL.Text = "段后(行)：";
            // 
            // BeforeMainBody
            // 
            this.BeforeMainBody.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.BeforeMainBody.DecimalPlaces = 1;
            this.BeforeMainBody.Font = new System.Drawing.Font("微软雅黑", 7.875F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.BeforeMainBody.Increment = new decimal(new int[] {
            5,
            0,
            0,
            65536});
            this.BeforeMainBody.Location = new System.Drawing.Point(190, 170);
            this.BeforeMainBody.Maximum = new decimal(new int[] {
            50,
            0,
            0,
            0});
            this.BeforeMainBody.Name = "BeforeMainBody";
            this.BeforeMainBody.Size = new System.Drawing.Size(90, 35);
            this.BeforeMainBody.TabIndex = 4;
            // 
            // BeforeMainBodyL
            // 
            this.BeforeMainBodyL.AutoSize = true;
            this.BeforeMainBodyL.Location = new System.Drawing.Point(68, 172);
            this.BeforeMainBodyL.Name = "BeforeMainBodyL";
            this.BeforeMainBodyL.Size = new System.Drawing.Size(126, 31);
            this.BeforeMainBodyL.TabIndex = 3;
            this.BeforeMainBodyL.Text = "段前(行)：";
            // 
            // FirstTitleCheck
            // 
            this.FirstTitleCheck.AutoSize = true;
            this.FirstTitleCheck.Location = new System.Drawing.Point(74, 102);
            this.FirstTitleCheck.Name = "FirstTitleCheck";
            this.FirstTitleCheck.Size = new System.Drawing.Size(447, 35);
            this.FirstTitleCheck.TabIndex = 2;
            this.FirstTitleCheck.Text = "一级标题、表格下一行 段前0.5倍行距";
            this.FirstTitleCheck.UseVisualStyleBackColor = true;
            // 
            // TitleIndentCkeck
            // 
            this.TitleIndentCkeck.AutoSize = true;
            this.TitleIndentCkeck.Location = new System.Drawing.Point(74, 60);
            this.TitleIndentCkeck.Name = "TitleIndentCkeck";
            this.TitleIndentCkeck.Size = new System.Drawing.Size(406, 35);
            this.TitleIndentCkeck.TabIndex = 1;
            this.TitleIndentCkeck.Text = "标题行靠左对齐，并左缩进-2字符";
            this.TitleIndentCkeck.UseVisualStyleBackColor = true;
            // 
            // TitleGroupCheck
            // 
            this.TitleGroupCheck.AutoSize = true;
            this.TitleGroupCheck.BackColor = System.Drawing.Color.White;
            this.TitleGroupCheck.Location = new System.Drawing.Point(3, 0);
            this.TitleGroupCheck.Name = "TitleGroupCheck";
            this.TitleGroupCheck.Size = new System.Drawing.Size(142, 35);
            this.TitleGroupCheck.TabIndex = 0;
            this.TitleGroupCheck.Text = "标题段落";
            this.TitleGroupCheck.UseVisualStyleBackColor = false;
            // 
            // FontGroup
            // 
            this.FontGroup.Controls.Add(this.FontSize);
            this.FontGroup.Controls.Add(this.FontSizeL);
            this.FontGroup.Controls.Add(this.NumFontL);
            this.FontGroup.Controls.Add(this.NumFont);
            this.FontGroup.Controls.Add(this.CnFontL);
            this.FontGroup.Controls.Add(this.CnFont);
            this.FontGroup.Controls.Add(this.FontGroupCheck);
            this.FontGroup.Dock = System.Windows.Forms.DockStyle.Top;
            this.FontGroup.Location = new System.Drawing.Point(3, 4);
            this.FontGroup.Name = "FontGroup";
            this.FontGroup.Size = new System.Drawing.Size(652, 190);
            this.FontGroup.TabIndex = 0;
            this.FontGroup.TabStop = false;
            // 
            // FontSize
            // 
            this.FontSize.Font = new System.Drawing.Font("微软雅黑", 7.875F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.FontSize.FormattingEnabled = true;
            this.FontSize.Location = new System.Drawing.Point(456, 65);
            this.FontSize.Name = "FontSize";
            this.FontSize.Size = new System.Drawing.Size(100, 36);
            this.FontSize.TabIndex = 6;
            // 
            // FontSizeL
            // 
            this.FontSizeL.AutoSize = true;
            this.FontSizeL.Location = new System.Drawing.Point(368, 68);
            this.FontSizeL.Name = "FontSizeL";
            this.FontSizeL.Size = new System.Drawing.Size(86, 31);
            this.FontSizeL.TabIndex = 5;
            this.FontSizeL.Text = "字号：";
            // 
            // NumFontL
            // 
            this.NumFontL.AutoSize = true;
            this.NumFontL.Location = new System.Drawing.Point(68, 122);
            this.NumFontL.Name = "NumFontL";
            this.NumFontL.Size = new System.Drawing.Size(86, 31);
            this.NumFontL.TabIndex = 4;
            this.NumFontL.Text = "数字：";
            // 
            // NumFont
            // 
            this.NumFont.Font = new System.Drawing.Font("微软雅黑", 7.875F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.NumFont.FormattingEnabled = true;
            this.NumFont.Location = new System.Drawing.Point(156, 120);
            this.NumFont.Name = "NumFont";
            this.NumFont.Size = new System.Drawing.Size(170, 36);
            this.NumFont.TabIndex = 3;
            // 
            // CnFontL
            // 
            this.CnFontL.AutoSize = true;
            this.CnFontL.Location = new System.Drawing.Point(68, 68);
            this.CnFontL.Name = "CnFontL";
            this.CnFontL.Size = new System.Drawing.Size(86, 31);
            this.CnFontL.TabIndex = 2;
            this.CnFontL.Text = "中文：";
            // 
            // CnFont
            // 
            this.CnFont.Font = new System.Drawing.Font("微软雅黑", 7.875F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.CnFont.FormattingEnabled = true;
            this.CnFont.Location = new System.Drawing.Point(156, 66);
            this.CnFont.Name = "CnFont";
            this.CnFont.Size = new System.Drawing.Size(170, 36);
            this.CnFont.TabIndex = 1;
            // 
            // FontGroupCheck
            // 
            this.FontGroupCheck.AutoSize = true;
            this.FontGroupCheck.BackColor = System.Drawing.Color.White;
            this.FontGroupCheck.Location = new System.Drawing.Point(3, 0);
            this.FontGroupCheck.Name = "FontGroupCheck";
            this.FontGroupCheck.Size = new System.Drawing.Size(142, 35);
            this.FontGroupCheck.TabIndex = 0;
            this.FontGroupCheck.Text = "字体字号";
            this.FontGroupCheck.UseVisualStyleBackColor = false;
            // 
            // TableFormat
            // 
            this.TableFormat.Controls.Add(this.groupBox3);
            this.TableFormat.Controls.Add(this.groupBox2);
            this.TableFormat.Controls.Add(this.groupBox1);
            this.TableFormat.Location = new System.Drawing.Point(8, 45);
            this.TableFormat.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.TableFormat.Name = "TableFormat";
            this.TableFormat.Size = new System.Drawing.Size(658, 482);
            this.TableFormat.TabIndex = 2;
            this.TableFormat.Text = "表格排版";
            this.TableFormat.UseVisualStyleBackColor = true;
            // 
            // comboBox1
            // 
            this.comboBox1.Font = new System.Drawing.Font("微软雅黑", 7.875F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.comboBox1.FormattingEnabled = true;
            this.comboBox1.Location = new System.Drawing.Point(114, 134);
            this.comboBox1.Name = "comboBox1";
            this.comboBox1.Size = new System.Drawing.Size(167, 36);
            this.comboBox1.TabIndex = 3;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(6, 137);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(110, 31);
            this.label1.TabIndex = 2;
            this.label1.Text = "行间距：";
            // 
            // checkBox2
            // 
            this.checkBox2.AutoSize = true;
            this.checkBox2.Location = new System.Drawing.Point(12, 89);
            this.checkBox2.Name = "checkBox2";
            this.checkBox2.Size = new System.Drawing.Size(242, 35);
            this.checkBox2.TabIndex = 1;
            this.checkBox2.Text = "段前段后 0 行间距";
            this.checkBox2.UseVisualStyleBackColor = true;
            // 
            // checkBox1
            // 
            this.checkBox1.AutoSize = true;
            this.checkBox1.Location = new System.Drawing.Point(12, 48);
            this.checkBox1.Name = "checkBox1";
            this.checkBox1.Size = new System.Drawing.Size(269, 35);
            this.checkBox1.TabIndex = 0;
            this.checkBox1.Text = "左、右、首行 不缩进";
            this.checkBox1.UseVisualStyleBackColor = true;
            // 
            // SaveButton
            // 
            this.SaveButton.Location = new System.Drawing.Point(260, 560);
            this.SaveButton.Name = "SaveButton";
            this.SaveButton.Size = new System.Drawing.Size(150, 50);
            this.SaveButton.TabIndex = 2;
            this.SaveButton.Text = "保  存";
            this.SaveButton.UseVisualStyleBackColor = true;
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.checkBox1);
            this.groupBox1.Controls.Add(this.comboBox1);
            this.groupBox1.Controls.Add(this.checkBox2);
            this.groupBox1.Controls.Add(this.label1);
            this.groupBox1.Dock = System.Windows.Forms.DockStyle.Top;
            this.groupBox1.Location = new System.Drawing.Point(0, 0);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(658, 186);
            this.groupBox1.TabIndex = 4;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "间距";
            // 
            // groupBox2
            // 
            this.groupBox2.Dock = System.Windows.Forms.DockStyle.Top;
            this.groupBox2.Location = new System.Drawing.Point(0, 186);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(658, 190);
            this.groupBox2.TabIndex = 5;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "边框";
            // 
            // checkBox3
            // 
            this.checkBox3.AutoSize = true;
            this.checkBox3.Location = new System.Drawing.Point(6, 48);
            this.checkBox3.Name = "checkBox3";
            this.checkBox3.Size = new System.Drawing.Size(142, 35);
            this.checkBox3.TabIndex = 0;
            this.checkBox3.Text = "垂直居中";
            this.checkBox3.UseVisualStyleBackColor = true;
            // 
            // groupBox3
            // 
            this.groupBox3.Controls.Add(this.checkBox4);
            this.groupBox3.Controls.Add(this.checkBox3);
            this.groupBox3.Dock = System.Windows.Forms.DockStyle.Top;
            this.groupBox3.Location = new System.Drawing.Point(0, 376);
            this.groupBox3.Name = "groupBox3";
            this.groupBox3.Size = new System.Drawing.Size(658, 100);
            this.groupBox3.TabIndex = 6;
            this.groupBox3.TabStop = false;
            this.groupBox3.Text = "其它";
            // 
            // checkBox4
            // 
            this.checkBox4.AutoSize = true;
            this.checkBox4.Location = new System.Drawing.Point(319, 48);
            this.checkBox4.Name = "checkBox4";
            this.checkBox4.Size = new System.Drawing.Size(238, 35);
            this.checkBox4.TabIndex = 1;
            this.checkBox4.Text = "标题及合计行加粗";
            this.checkBox4.UseVisualStyleBackColor = true;
            // 
            // SettingFormatForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(14F, 31F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(674, 629);
            this.Controls.Add(this.SaveButton);
            this.Controls.Add(this.SettingFormat);
            this.Font = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.Name = "SettingFormatForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "设置默认样式";
            this.Load += new System.EventHandler(this.SettingFormatForm_Load);
            this.SettingFormat.ResumeLayout(false);
            this.PageSetting.ResumeLayout(false);
            this.PageFormat.ResumeLayout(false);
            this.PageFormat.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.FooterMargin)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.HeadMargin)).EndInit();
            this.PageMargin.ResumeLayout(false);
            this.PageMargin.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.RightMargin)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.LeftMargin)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.BottomMargin)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.TopMargin)).EndInit();
            this.TextFormat.ResumeLayout(false);
            this.TitleGroup.ResumeLayout(false);
            this.TitleGroup.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.AfterMainBody)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.BeforeMainBody)).EndInit();
            this.FontGroup.ResumeLayout(false);
            this.FontGroup.PerformLayout();
            this.TableFormat.ResumeLayout(false);
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.groupBox3.ResumeLayout(false);
            this.groupBox3.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.TabControl SettingFormat;
        private System.Windows.Forms.TabPage PageSetting;
        private System.Windows.Forms.TabPage TextFormat;
        private System.Windows.Forms.GroupBox PageFormat;
        private System.Windows.Forms.GroupBox PageMargin;
        private System.Windows.Forms.CheckBox PageMarginCheck;
        private System.Windows.Forms.NumericUpDown RightMargin;
        private System.Windows.Forms.Label RightMarginL;
        private System.Windows.Forms.NumericUpDown LeftMargin;
        private System.Windows.Forms.Label LeftMarginL;
        private System.Windows.Forms.NumericUpDown BottomMargin;
        private System.Windows.Forms.Label BottomMarginL;
        private System.Windows.Forms.NumericUpDown TopMargin;
        private System.Windows.Forms.Label TopMarginL;
        private System.Windows.Forms.TabPage TableFormat;
        private System.Windows.Forms.NumericUpDown FooterMargin;
        private System.Windows.Forms.Label FooterMarginL;
        private System.Windows.Forms.NumericUpDown HeadMargin;
        private System.Windows.Forms.CheckBox PageFormatCheck;
        private System.Windows.Forms.Label HeadMarginL;
        private System.Windows.Forms.GroupBox FontGroup;
        private System.Windows.Forms.Label CnFontL;
        private System.Windows.Forms.ComboBox CnFont;
        private System.Windows.Forms.CheckBox FontGroupCheck;
        private System.Windows.Forms.Button SaveButton;
        private System.Windows.Forms.Label NumFontL;
        private System.Windows.Forms.ComboBox NumFont;
        private System.Windows.Forms.ComboBox FontSize;
        private System.Windows.Forms.Label FontSizeL;
        private System.Windows.Forms.GroupBox TitleGroup;
        private System.Windows.Forms.NumericUpDown AfterMainBody;
        private System.Windows.Forms.Label AfterMainBodyL;
        private System.Windows.Forms.NumericUpDown BeforeMainBody;
        private System.Windows.Forms.Label BeforeMainBodyL;
        private System.Windows.Forms.CheckBox FirstTitleCheck;
        private System.Windows.Forms.CheckBox TitleIndentCkeck;
        private System.Windows.Forms.CheckBox TitleGroupCheck;
        private System.Windows.Forms.ComboBox RowSpace;
        private System.Windows.Forms.Label RowSpaceL;
        private System.Windows.Forms.ComboBox comboBox1;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.CheckBox checkBox2;
        private System.Windows.Forms.CheckBox checkBox1;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.GroupBox groupBox3;
        private System.Windows.Forms.CheckBox checkBox4;
        private System.Windows.Forms.CheckBox checkBox3;
        private System.Windows.Forms.GroupBox groupBox2;
    }
}