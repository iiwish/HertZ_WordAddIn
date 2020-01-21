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
            this.TextFormat = new System.Windows.Forms.TabPage();
            this.SkipPgs = new System.Windows.Forms.NumericUpDown();
            this.SkipPgsL = new System.Windows.Forms.Label();
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
            this.OtherGroup = new System.Windows.Forms.GroupBox();
            this.OtherGroupCheck = new System.Windows.Forms.CheckBox();
            this.TableFontSize = new System.Windows.Forms.ComboBox();
            this.TableFontSizeL = new System.Windows.Forms.Label();
            this.TableTitleWiderCheck = new System.Windows.Forms.CheckBox();
            this.TableVerticalCenterCheck = new System.Windows.Forms.CheckBox();
            this.BorderGroup = new System.Windows.Forms.GroupBox();
            this.BorderGroupCheck = new System.Windows.Forms.CheckBox();
            this.TableLRBorderCheck = new System.Windows.Forms.CheckBox();
            this.TableTBBorderCheck = new System.Windows.Forms.CheckBox();
            this.SpaceGroup = new System.Windows.Forms.GroupBox();
            this.TableAfterMainBody = new System.Windows.Forms.NumericUpDown();
            this.TableBeforeMainBody = new System.Windows.Forms.NumericUpDown();
            this.TableAfterMainBodyL = new System.Windows.Forms.Label();
            this.TableBeforeMainBodyL = new System.Windows.Forms.Label();
            this.SpaceGroupCheck = new System.Windows.Forms.CheckBox();
            this.TableIndentCheck = new System.Windows.Forms.CheckBox();
            this.TableRowSpace = new System.Windows.Forms.ComboBox();
            this.TableRowSpaceL = new System.Windows.Forms.Label();
            this.SaveButton = new System.Windows.Forms.Button();
            this.SettingFormat.SuspendLayout();
            this.TextFormat.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.SkipPgs)).BeginInit();
            this.TitleGroup.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.AfterMainBody)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.BeforeMainBody)).BeginInit();
            this.FontGroup.SuspendLayout();
            this.TableFormat.SuspendLayout();
            this.OtherGroup.SuspendLayout();
            this.BorderGroup.SuspendLayout();
            this.SpaceGroup.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.TableAfterMainBody)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.TableBeforeMainBody)).BeginInit();
            this.SuspendLayout();
            // 
            // SettingFormat
            // 
            this.SettingFormat.Controls.Add(this.TextFormat);
            this.SettingFormat.Controls.Add(this.TableFormat);
            this.SettingFormat.Dock = System.Windows.Forms.DockStyle.Top;
            this.SettingFormat.Font = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.SettingFormat.Location = new System.Drawing.Point(0, 0);
            this.SettingFormat.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.SettingFormat.Name = "SettingFormat";
            this.SettingFormat.SelectedIndex = 0;
            this.SettingFormat.Size = new System.Drawing.Size(674, 580);
            this.SettingFormat.TabIndex = 1;
            // 
            // TextFormat
            // 
            this.TextFormat.Controls.Add(this.SkipPgs);
            this.TextFormat.Controls.Add(this.SkipPgsL);
            this.TextFormat.Controls.Add(this.TitleGroup);
            this.TextFormat.Controls.Add(this.FontGroup);
            this.TextFormat.Location = new System.Drawing.Point(8, 45);
            this.TextFormat.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.TextFormat.Name = "TextFormat";
            this.TextFormat.Padding = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.TextFormat.Size = new System.Drawing.Size(658, 527);
            this.TextFormat.TabIndex = 1;
            this.TextFormat.Text = "文字排版";
            this.TextFormat.UseVisualStyleBackColor = true;
            // 
            // SkipPgs
            // 
            this.SkipPgs.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.SkipPgs.Font = new System.Drawing.Font("微软雅黑", 7.875F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.SkipPgs.Location = new System.Drawing.Point(294, 490);
            this.SkipPgs.Name = "SkipPgs";
            this.SkipPgs.Size = new System.Drawing.Size(80, 35);
            this.SkipPgs.TabIndex = 3;
            // 
            // SkipPgsL
            // 
            this.SkipPgsL.AutoSize = true;
            this.SkipPgsL.Location = new System.Drawing.Point(71, 492);
            this.SkipPgsL.Name = "SkipPgsL";
            this.SkipPgsL.Size = new System.Drawing.Size(238, 31);
            this.SkipPgsL.TabIndex = 2;
            this.SkipPgsL.Text = "格式调整跳过(段）：";
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
            this.TableFormat.Controls.Add(this.OtherGroup);
            this.TableFormat.Controls.Add(this.BorderGroup);
            this.TableFormat.Controls.Add(this.SpaceGroup);
            this.TableFormat.Location = new System.Drawing.Point(8, 45);
            this.TableFormat.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.TableFormat.Name = "TableFormat";
            this.TableFormat.Size = new System.Drawing.Size(658, 527);
            this.TableFormat.TabIndex = 2;
            this.TableFormat.Text = "表格排版";
            this.TableFormat.UseVisualStyleBackColor = true;
            // 
            // OtherGroup
            // 
            this.OtherGroup.Controls.Add(this.OtherGroupCheck);
            this.OtherGroup.Controls.Add(this.TableFontSize);
            this.OtherGroup.Controls.Add(this.TableFontSizeL);
            this.OtherGroup.Controls.Add(this.TableTitleWiderCheck);
            this.OtherGroup.Controls.Add(this.TableVerticalCenterCheck);
            this.OtherGroup.Dock = System.Windows.Forms.DockStyle.Top;
            this.OtherGroup.Location = new System.Drawing.Point(0, 328);
            this.OtherGroup.Name = "OtherGroup";
            this.OtherGroup.Size = new System.Drawing.Size(658, 151);
            this.OtherGroup.TabIndex = 6;
            this.OtherGroup.TabStop = false;
            // 
            // OtherGroupCheck
            // 
            this.OtherGroupCheck.AutoSize = true;
            this.OtherGroupCheck.BackColor = System.Drawing.Color.White;
            this.OtherGroupCheck.Location = new System.Drawing.Point(6, 0);
            this.OtherGroupCheck.Name = "OtherGroupCheck";
            this.OtherGroupCheck.Size = new System.Drawing.Size(94, 35);
            this.OtherGroupCheck.TabIndex = 4;
            this.OtherGroupCheck.Text = "其他";
            this.OtherGroupCheck.UseVisualStyleBackColor = false;
            // 
            // TableFontSize
            // 
            this.TableFontSize.Font = new System.Drawing.Font("微软雅黑", 7.875F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.TableFontSize.FormattingEnabled = true;
            this.TableFontSize.Location = new System.Drawing.Point(121, 98);
            this.TableFontSize.Name = "TableFontSize";
            this.TableFontSize.Size = new System.Drawing.Size(100, 36);
            this.TableFontSize.TabIndex = 3;
            // 
            // TableFontSizeL
            // 
            this.TableFontSizeL.AutoSize = true;
            this.TableFontSizeL.Location = new System.Drawing.Point(44, 100);
            this.TableFontSizeL.Name = "TableFontSizeL";
            this.TableFontSizeL.Size = new System.Drawing.Size(86, 31);
            this.TableFontSizeL.TabIndex = 2;
            this.TableFontSizeL.Text = "字号：";
            // 
            // TableTitleWiderCheck
            // 
            this.TableTitleWiderCheck.AutoSize = true;
            this.TableTitleWiderCheck.Location = new System.Drawing.Point(319, 48);
            this.TableTitleWiderCheck.Name = "TableTitleWiderCheck";
            this.TableTitleWiderCheck.Size = new System.Drawing.Size(238, 35);
            this.TableTitleWiderCheck.TabIndex = 1;
            this.TableTitleWiderCheck.Text = "标题及合计行加粗";
            this.TableTitleWiderCheck.UseVisualStyleBackColor = true;
            this.TableTitleWiderCheck.Visible = false;
            // 
            // TableVerticalCenterCheck
            // 
            this.TableVerticalCenterCheck.AutoSize = true;
            this.TableVerticalCenterCheck.Location = new System.Drawing.Point(50, 50);
            this.TableVerticalCenterCheck.Name = "TableVerticalCenterCheck";
            this.TableVerticalCenterCheck.Size = new System.Drawing.Size(142, 35);
            this.TableVerticalCenterCheck.TabIndex = 0;
            this.TableVerticalCenterCheck.Text = "垂直居中";
            this.TableVerticalCenterCheck.UseVisualStyleBackColor = true;
            // 
            // BorderGroup
            // 
            this.BorderGroup.Controls.Add(this.BorderGroupCheck);
            this.BorderGroup.Controls.Add(this.TableLRBorderCheck);
            this.BorderGroup.Controls.Add(this.TableTBBorderCheck);
            this.BorderGroup.Dock = System.Windows.Forms.DockStyle.Top;
            this.BorderGroup.Location = new System.Drawing.Point(0, 230);
            this.BorderGroup.Name = "BorderGroup";
            this.BorderGroup.Size = new System.Drawing.Size(658, 98);
            this.BorderGroup.TabIndex = 5;
            this.BorderGroup.TabStop = false;
            // 
            // BorderGroupCheck
            // 
            this.BorderGroupCheck.AutoSize = true;
            this.BorderGroupCheck.BackColor = System.Drawing.Color.White;
            this.BorderGroupCheck.Location = new System.Drawing.Point(6, 0);
            this.BorderGroupCheck.Name = "BorderGroupCheck";
            this.BorderGroupCheck.Size = new System.Drawing.Size(94, 35);
            this.BorderGroupCheck.TabIndex = 2;
            this.BorderGroupCheck.Text = "边框";
            this.BorderGroupCheck.UseVisualStyleBackColor = false;
            // 
            // TableLRBorderCheck
            // 
            this.TableLRBorderCheck.AutoSize = true;
            this.TableLRBorderCheck.Location = new System.Drawing.Point(319, 50);
            this.TableLRBorderCheck.Name = "TableLRBorderCheck";
            this.TableLRBorderCheck.Size = new System.Drawing.Size(166, 35);
            this.TableLRBorderCheck.TabIndex = 1;
            this.TableLRBorderCheck.Text = "左右无边框";
            this.TableLRBorderCheck.UseVisualStyleBackColor = true;
            // 
            // TableTBBorderCheck
            // 
            this.TableTBBorderCheck.AutoSize = true;
            this.TableTBBorderCheck.Location = new System.Drawing.Point(50, 50);
            this.TableTBBorderCheck.Name = "TableTBBorderCheck";
            this.TableTBBorderCheck.Size = new System.Drawing.Size(228, 35);
            this.TableTBBorderCheck.TabIndex = 0;
            this.TableTBBorderCheck.Text = "上下边框宽度1磅";
            this.TableTBBorderCheck.UseVisualStyleBackColor = true;
            // 
            // SpaceGroup
            // 
            this.SpaceGroup.Controls.Add(this.TableAfterMainBody);
            this.SpaceGroup.Controls.Add(this.TableBeforeMainBody);
            this.SpaceGroup.Controls.Add(this.TableAfterMainBodyL);
            this.SpaceGroup.Controls.Add(this.TableBeforeMainBodyL);
            this.SpaceGroup.Controls.Add(this.SpaceGroupCheck);
            this.SpaceGroup.Controls.Add(this.TableIndentCheck);
            this.SpaceGroup.Controls.Add(this.TableRowSpace);
            this.SpaceGroup.Controls.Add(this.TableRowSpaceL);
            this.SpaceGroup.Dock = System.Windows.Forms.DockStyle.Top;
            this.SpaceGroup.Location = new System.Drawing.Point(0, 0);
            this.SpaceGroup.Name = "SpaceGroup";
            this.SpaceGroup.Size = new System.Drawing.Size(658, 230);
            this.SpaceGroup.TabIndex = 4;
            this.SpaceGroup.TabStop = false;
            // 
            // TableAfterMainBody
            // 
            this.TableAfterMainBody.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.TableAfterMainBody.DecimalPlaces = 1;
            this.TableAfterMainBody.Font = new System.Drawing.Font("微软雅黑", 7.875F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.TableAfterMainBody.Increment = new decimal(new int[] {
            5,
            0,
            0,
            65536});
            this.TableAfterMainBody.Location = new System.Drawing.Point(164, 172);
            this.TableAfterMainBody.Maximum = new decimal(new int[] {
            99,
            0,
            0,
            0});
            this.TableAfterMainBody.Name = "TableAfterMainBody";
            this.TableAfterMainBody.Size = new System.Drawing.Size(90, 35);
            this.TableAfterMainBody.TabIndex = 8;
            // 
            // TableBeforeMainBody
            // 
            this.TableBeforeMainBody.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.TableBeforeMainBody.DecimalPlaces = 1;
            this.TableBeforeMainBody.Font = new System.Drawing.Font("微软雅黑", 7.875F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.TableBeforeMainBody.Increment = new decimal(new int[] {
            5,
            0,
            0,
            65536});
            this.TableBeforeMainBody.Location = new System.Drawing.Point(164, 118);
            this.TableBeforeMainBody.Maximum = new decimal(new int[] {
            99,
            0,
            0,
            0});
            this.TableBeforeMainBody.Name = "TableBeforeMainBody";
            this.TableBeforeMainBody.Size = new System.Drawing.Size(90, 35);
            this.TableBeforeMainBody.TabIndex = 7;
            // 
            // TableAfterMainBodyL
            // 
            this.TableAfterMainBodyL.AutoSize = true;
            this.TableAfterMainBodyL.Location = new System.Drawing.Point(50, 173);
            this.TableAfterMainBodyL.Name = "TableAfterMainBodyL";
            this.TableAfterMainBodyL.Size = new System.Drawing.Size(126, 31);
            this.TableAfterMainBodyL.TabIndex = 6;
            this.TableAfterMainBodyL.Text = "段后(行)：";
            // 
            // TableBeforeMainBodyL
            // 
            this.TableBeforeMainBodyL.AutoSize = true;
            this.TableBeforeMainBodyL.Location = new System.Drawing.Point(50, 120);
            this.TableBeforeMainBodyL.Name = "TableBeforeMainBodyL";
            this.TableBeforeMainBodyL.Size = new System.Drawing.Size(126, 31);
            this.TableBeforeMainBodyL.TabIndex = 5;
            this.TableBeforeMainBodyL.Text = "段前(行)：";
            // 
            // SpaceGroupCheck
            // 
            this.SpaceGroupCheck.AutoSize = true;
            this.SpaceGroupCheck.BackColor = System.Drawing.Color.White;
            this.SpaceGroupCheck.Location = new System.Drawing.Point(6, 3);
            this.SpaceGroupCheck.Name = "SpaceGroupCheck";
            this.SpaceGroupCheck.Size = new System.Drawing.Size(94, 35);
            this.SpaceGroupCheck.TabIndex = 4;
            this.SpaceGroupCheck.Text = "间距";
            this.SpaceGroupCheck.UseVisualStyleBackColor = false;
            // 
            // TableIndentCheck
            // 
            this.TableIndentCheck.AutoSize = true;
            this.TableIndentCheck.Location = new System.Drawing.Point(50, 50);
            this.TableIndentCheck.Name = "TableIndentCheck";
            this.TableIndentCheck.Size = new System.Drawing.Size(269, 35);
            this.TableIndentCheck.TabIndex = 0;
            this.TableIndentCheck.Text = "左、右、首行 不缩进";
            this.TableIndentCheck.UseVisualStyleBackColor = true;
            // 
            // TableRowSpace
            // 
            this.TableRowSpace.Font = new System.Drawing.Font("微软雅黑", 7.875F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.TableRowSpace.FormattingEnabled = true;
            this.TableRowSpace.Location = new System.Drawing.Point(342, 168);
            this.TableRowSpace.Name = "TableRowSpace";
            this.TableRowSpace.Size = new System.Drawing.Size(180, 36);
            this.TableRowSpace.TabIndex = 3;
            // 
            // TableRowSpaceL
            // 
            this.TableRowSpaceL.AutoSize = true;
            this.TableRowSpaceL.Location = new System.Drawing.Point(336, 119);
            this.TableRowSpaceL.Name = "TableRowSpaceL";
            this.TableRowSpaceL.Size = new System.Drawing.Size(110, 31);
            this.TableRowSpaceL.TabIndex = 2;
            this.TableRowSpaceL.Text = "行间距：";
            // 
            // SaveButton
            // 
            this.SaveButton.Location = new System.Drawing.Point(255, 594);
            this.SaveButton.Name = "SaveButton";
            this.SaveButton.Size = new System.Drawing.Size(150, 50);
            this.SaveButton.TabIndex = 2;
            this.SaveButton.Text = "保 存";
            this.SaveButton.UseVisualStyleBackColor = true;
            this.SaveButton.Click += new System.EventHandler(this.SaveButton_Click);
            // 
            // SettingFormatForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(14F, 31F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(674, 665);
            this.Controls.Add(this.SaveButton);
            this.Controls.Add(this.SettingFormat);
            this.Font = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.Name = "SettingFormatForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "设置默认样式";
            this.TopMost = true;
            this.Load += new System.EventHandler(this.SettingFormatForm_Load);
            this.SettingFormat.ResumeLayout(false);
            this.TextFormat.ResumeLayout(false);
            this.TextFormat.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.SkipPgs)).EndInit();
            this.TitleGroup.ResumeLayout(false);
            this.TitleGroup.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.AfterMainBody)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.BeforeMainBody)).EndInit();
            this.FontGroup.ResumeLayout(false);
            this.FontGroup.PerformLayout();
            this.TableFormat.ResumeLayout(false);
            this.OtherGroup.ResumeLayout(false);
            this.OtherGroup.PerformLayout();
            this.BorderGroup.ResumeLayout(false);
            this.BorderGroup.PerformLayout();
            this.SpaceGroup.ResumeLayout(false);
            this.SpaceGroup.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.TableAfterMainBody)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.TableBeforeMainBody)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.TabControl SettingFormat;
        private System.Windows.Forms.TabPage TextFormat;
        private System.Windows.Forms.TabPage TableFormat;
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
        private System.Windows.Forms.ComboBox TableRowSpace;
        private System.Windows.Forms.Label TableRowSpaceL;
        private System.Windows.Forms.CheckBox TableIndentCheck;
        private System.Windows.Forms.GroupBox SpaceGroup;
        private System.Windows.Forms.GroupBox OtherGroup;
        private System.Windows.Forms.CheckBox TableTitleWiderCheck;
        private System.Windows.Forms.CheckBox TableVerticalCenterCheck;
        private System.Windows.Forms.GroupBox BorderGroup;
        private System.Windows.Forms.CheckBox TableLRBorderCheck;
        private System.Windows.Forms.CheckBox TableTBBorderCheck;
        private System.Windows.Forms.CheckBox BorderGroupCheck;
        private System.Windows.Forms.CheckBox SpaceGroupCheck;
        private System.Windows.Forms.NumericUpDown TableBeforeMainBody;
        private System.Windows.Forms.Label TableAfterMainBodyL;
        private System.Windows.Forms.Label TableBeforeMainBodyL;
        private System.Windows.Forms.NumericUpDown TableAfterMainBody;
        private System.Windows.Forms.ComboBox TableFontSize;
        private System.Windows.Forms.Label TableFontSizeL;
        private System.Windows.Forms.CheckBox OtherGroupCheck;
        private System.Windows.Forms.NumericUpDown SkipPgs;
        private System.Windows.Forms.Label SkipPgsL;
    }
}