namespace HertZ_WordAddIn
{
    partial class MyProcessBar
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(MyProcessBar));
            this.ProcessBar = new System.Windows.Forms.ProgressBar();
            this.CancelBtn = new System.Windows.Forms.Button();
            this.ProcessLabel = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // ProcessBar
            // 
            this.ProcessBar.Location = new System.Drawing.Point(35, 50);
            this.ProcessBar.Name = "ProcessBar";
            this.ProcessBar.Size = new System.Drawing.Size(859, 35);
            this.ProcessBar.Step = 1;
            this.ProcessBar.TabIndex = 1;
            // 
            // CancelBtn
            // 
            this.CancelBtn.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.CancelBtn.Location = new System.Drawing.Point(794, 141);
            this.CancelBtn.Name = "CancelBtn";
            this.CancelBtn.Size = new System.Drawing.Size(100, 50);
            this.CancelBtn.TabIndex = 0;
            this.CancelBtn.Text = "取 消";
            this.CancelBtn.UseVisualStyleBackColor = true;
            this.CancelBtn.Click += new System.EventHandler(this.CancelBtn_Click);
            // 
            // ProcessLabel
            // 
            this.ProcessLabel.AutoSize = true;
            this.ProcessLabel.Location = new System.Drawing.Point(29, 141);
            this.ProcessLabel.Name = "ProcessLabel";
            this.ProcessLabel.Size = new System.Drawing.Size(182, 31);
            this.ProcessLabel.TabIndex = 2;
            this.ProcessLabel.Text = "加载中，请稍候";
            // 
            // MyProcessBar
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(14F, 31F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.CancelButton = this.CancelBtn;
            this.ClientSize = new System.Drawing.Size(933, 229);
            this.ControlBox = false;
            this.Controls.Add(this.ProcessLabel);
            this.Controls.Add(this.CancelBtn);
            this.Controls.Add(this.ProcessBar);
            this.Font = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.Name = "MyProcessBar";
            this.Text = "修改中，请稍候。。。";
            this.TopMost = true;
            this.Load += new System.EventHandler(this.MyProcessBar_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.ProgressBar ProcessBar;
        private System.Windows.Forms.Button CancelBtn;
        private System.Windows.Forms.Label ProcessLabel;
    }
}