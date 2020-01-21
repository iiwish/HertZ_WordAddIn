using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Deployment.Application;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace HertZ_WordAddIn
{
    public partial class VerInfo : Form
    {
        public VerInfo()
        {
            InitializeComponent();
        }

        private void VerInfo_Load(object sender, EventArgs e)
        {
            try
            {
                label1.Text = "当前版本：" + ApplicationDeployment.CurrentDeployment.CurrentVersion.ToString();
            }
            catch
            {
                label1.Text = "版本号获取异常";
            }
        }

        private void Manual_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start("https://www.jianshu.com/nb/42169573");
        }

        private void OnlineVideo_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start("");
        }

    }
}
