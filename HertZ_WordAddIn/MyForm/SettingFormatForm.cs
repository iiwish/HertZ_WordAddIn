using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace HertZ_WordAddIn
{
    public partial class SettingFormatForm : Form
    {
        private System.Drawing.Text.InstalledFontCollection objFont = new System.Drawing.Text.InstalledFontCollection();
        public SettingFormatForm()
        {
            InitializeComponent();
        }

        private void SettingFormatForm_Load(object sender, EventArgs e)
        {
            //遍历系统字体并添加到字体选择框
            foreach (System.Drawing.FontFamily i in objFont.Families)
            {
                CnFont.Items.Add(i.Name.ToString());
                NumFont.Items.Add(i.Name.ToString());
            }
            //CnFont.SelectedIndex = 0;

            string[] FontSizeList = 
                new[]{ "初号", "小初", "一号","小一","二号","小二","三号","小三","四号","小四","五号",
                   "小五", "六号", "小六","七号","八号","5","5.5","6.5","7.5","8","9","10","10.5","11",
                   "12","14","16","18","20","22","24","26","28","36","48","72" };
            FontSize.DataSource = FontSizeList;

            string[] RowSpaceList =
                new[] { "单倍行距","1.15倍行距","1.5倍行距","2倍行距","2.5倍行距","3倍行距" };
            RowSpace.DataSource = RowSpaceList;
        }
    }
}
