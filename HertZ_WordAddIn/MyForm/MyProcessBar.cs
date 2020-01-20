using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;

namespace HertZ_WordAddIn
{
    public partial class MyProcessBar : Form
    {
        //引入调整格式用的函数模块
        private readonly FormatFunC FormatFunC = new FormatFunC();

        public MyProcessBar()
        {
            InitializeComponent();
        }

        
        private void MyProcessBar_Load(object sender, EventArgs e)
        {
            
            //定义最大值
            ProcessBar.Maximum = 100;

        }

        public void Increase(int nValue)
        {
            float percent = (float)nValue * 100 / (float)ProcessBar.Maximum;

            if (percent > 0)
            {
                ProcessLabel.Text = "当前进度：(" + percent.ToString(".") + "%)";
            }
            else
            {
                ProcessLabel.Text = "当前进度：(0%)";
            }
            
            if (nValue < ProcessBar.Maximum)
            {
                ProcessBar.Value = nValue;
            }
            else
            {
                ProcessBar.Value = ProcessBar.Maximum;
                Exited = true;
                this.Close();
            }
        }

        bool Exited = false;
        private void CancelBtn_Click(object sender, EventArgs e)
        {
            if (!Exited)
            {
                if (MessageBox.Show("确认取消当前任务？", "系统提示", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                {
                    Exited = true;
                    this.Close();
                }
                else
                {
                    Exited = false;
                }
            }
        }

    }
}
