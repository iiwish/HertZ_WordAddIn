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

        //向我的文档写入配置
        ClsThisAddinConfig clsConfig = new ClsThisAddinConfig(Environment.GetFolderPath(Environment.SpecialFolder.CommonDocuments));

        public SettingFormatForm()
        {
            InitializeComponent();
        }

        //启动时加载
        private void SettingFormatForm_Load(object sender, EventArgs e)
        {
            //遍历系统字体并添加到字体选择框
            foreach (FontFamily i in objFont.Families)
            {
                CnFont.Items.Add(i.Name.ToString());
                NumFont.Items.Add(i.Name.ToString());
            }
            
            //设置字号窗体选项
            string[] FontSizeList = 
                new[]{ "初号", "小初", "一号","小一","二号","小二","三号","小三","四号","小四","五号",
                   "小五", "六号", "小六","七号","八号","5","5.5","6.5","7.5","8","9","10","10.5","11",
                   "12","14","16","18","20","22","24","26","28","36","48","72" };
            FontSize.DataSource = FontSizeList;
            TableFontSize.DataSource = FontSizeList;

            //设置行距选项
            string[] RowSpaceList =
                new[] { "单倍行距","1.15倍行距","1.5倍行距","2倍行距","2.5倍行距","3倍行距" };
            RowSpace.DataSource = RowSpaceList;
            TableRowSpace.DataSource = RowSpaceList;

            //从父节点SettingFormatForm中读取配置名为FontGroupCheck的值，字体字号复选框，默认为true
            FontGroupCheck.Checked = clsConfig.ReadConfig<bool>("SettingFormatForm", "FontGroupCheck", true);
            //中文字体
            CnFont.SelectedItem = clsConfig.ReadConfig<string>("SettingFormatForm", "CnFont", "仿宋_GB2312");
            //数字和英文字体
            NumFont.SelectedItem = clsConfig.ReadConfig<string>("SettingFormatForm", "NumFont", "Arial Narrow");
            //字号
            FontSize.SelectedItem = clsConfig.ReadConfig<string>("SettingFormatForm", "FontSize", "小四");

            //标题段落
            TitleGroupCheck.Checked = clsConfig.ReadConfig<bool>("SettingFormatForm", "TitleGroupCheck", true);
            //标题行左缩进
            TitleIndentCkeck.Checked = clsConfig.ReadConfig<bool>("SettingFormatForm", "TitleIndentCkeck", true);
            //一级标题、 表格下一行段前0.5倍行距
            FirstTitleCheck.Checked = clsConfig.ReadConfig<bool>("SettingFormatForm", "FirstTitleCheck", true);
            //段前
            BeforeMainBody.Value = clsConfig.ReadConfig<decimal>("SettingFormatForm", "BeforeMainBody", 0m);
            //段后
            AfterMainBody.Value = clsConfig.ReadConfig<decimal>("SettingFormatForm", "AfterMainBody", 0.9m);
            //行间距
            RowSpace.SelectedItem = clsConfig.ReadConfig<string>("SettingFormatForm", "RowSpace", "单倍行距");

            //表格间距
            SpaceGroupCheck.Checked = clsConfig.ReadConfig<bool>("SettingFormatForm", "SpaceGroupCheck", true);
            //首行缩进
            TableIndentCheck.Checked = clsConfig.ReadConfig<bool>("SettingFormatForm", "TableIndentCheck", true);
            //段前
            TableBeforeMainBody.Value = clsConfig.ReadConfig<decimal>("SettingFormatForm", "TableBeforeMainBody", 0m);
            //段后
            TableAfterMainBody.Value = clsConfig.ReadConfig<decimal>("SettingFormatForm", "TableAfterMainBody", 0m);
            //行间距
            TableRowSpace.SelectedItem = clsConfig.ReadConfig<string>("SettingFormatForm", "TableRowSpace", "单倍行距");

            //表格边框
            BorderGroupCheck.Checked = clsConfig.ReadConfig<bool>("SettingFormatForm", "BorderGroupCheck", true);
            //上下边框宽度1磅
            TableTBBorderCheck.Checked = clsConfig.ReadConfig<bool>("SettingFormatForm", "TableTBBorderCheck", true);
            //左右无边框
            TableLRBorderCheck.Checked = clsConfig.ReadConfig<bool>("SettingFormatForm", "TableLRBorderCheck", true);

            //表格其他选项
            OtherGroupCheck.Checked = clsConfig.ReadConfig<bool>("SettingFormatForm", "OtherGroupCheck", true);
            //垂直居中
            TableVerticalCenterCheck.Checked = clsConfig.ReadConfig<bool>("SettingFormatForm", "TableVerticalCenterCheck", true);
            //标题及合计行加粗
            TableTitleWiderCheck.Checked = clsConfig.ReadConfig<bool>("SettingFormatForm", "TableTitleWiderCheck", true);
            //表格字号
            TableFontSize.SelectedItem = clsConfig.ReadConfig<string>("SettingFormatForm", "TableFontSize", "小四");

        }

        //保存按钮
        private void SaveButton_Click(object sender, EventArgs e)
        {
            //验证数据是否有效
            TestEffect();

            //从父节点SettingFormatForm中读取配置名为FontGroupCheck的值，字体字号复选框，默认为true
            clsConfig.WriteConfig("SettingFormatForm", "FontGroupCheck", FontGroupCheck.Checked.ToString());
            //中文字体
            clsConfig.WriteConfig("SettingFormatForm", "CnFont", CnFont.SelectedItem.ToString());
            //数字和英文字体
            clsConfig.WriteConfig("SettingFormatForm", "NumFont", NumFont.SelectedItem.ToString());
            //字号
            clsConfig.WriteConfig("SettingFormatForm", "FontSize", FontSize.SelectedItem.ToString());

            //标题段落
            clsConfig.WriteConfig("SettingFormatForm", "TitleGroupCheck", TitleGroupCheck.Checked.ToString());
            //标题行左缩进
            clsConfig.WriteConfig("SettingFormatForm", "TitleIndentCkeck", TitleIndentCkeck.Checked.ToString());
            //一级标题、 表格下一行段前0.5倍行距
            clsConfig.WriteConfig("SettingFormatForm", "FirstTitleCheck", FirstTitleCheck.Checked.ToString());
            //段前
            clsConfig.WriteConfig("SettingFormatForm", "BeforeMainBody", BeforeMainBody.Value.ToString());
            //段后
            clsConfig.WriteConfig("SettingFormatForm", "AfterMainBody", AfterMainBody.Value.ToString());
            //行间距
            clsConfig.WriteConfig("SettingFormatForm", "RowSpace", RowSpace.SelectedItem.ToString());

            //表格间距
            clsConfig.WriteConfig("SettingFormatForm", "SpaceGroupCheck", SpaceGroupCheck.Checked.ToString());
            //首行缩进
            clsConfig.WriteConfig("SettingFormatForm", "TableIndentCheck", TableIndentCheck.Checked.ToString());
            //段前
            clsConfig.WriteConfig("SettingFormatForm", "TableBeforeMainBody", TableBeforeMainBody.Value.ToString());
            //段后
            clsConfig.WriteConfig("SettingFormatForm", "TableAfterMainBody", TableAfterMainBody.Value.ToString());
            //行间距
            clsConfig.WriteConfig("SettingFormatForm", "TableRowSpace", TableRowSpace.SelectedItem.ToString());

            //表格边框
            clsConfig.WriteConfig("SettingFormatForm", "BorderGroupCheck", BorderGroupCheck.Checked.ToString());
            //上下边框宽度1磅
            clsConfig.WriteConfig("SettingFormatForm", "TableTBBorderCheck", TableTBBorderCheck.Checked.ToString());
            //左右无边框
            clsConfig.WriteConfig("SettingFormatForm", "TableLRBorderCheck", TableLRBorderCheck.Checked.ToString());

            //表格其他选项
            clsConfig.WriteConfig("SettingFormatForm", "OtherGroupCheck", OtherGroupCheck.Checked.ToString());
            //垂直居中
            clsConfig.WriteConfig("SettingFormatForm", "TableVerticalCenterCheck", TableVerticalCenterCheck.Checked.ToString());
            //标题及合计行加粗
            clsConfig.WriteConfig("SettingFormatForm", "TableTitleWiderCheck", TableTitleWiderCheck.Checked.ToString());
            //表格字号
            clsConfig.WriteConfig("SettingFormatForm", "TableFontSize", TableFontSize.SelectedItem.ToString());
            
            //关闭当前窗体
            this.Close();
        }

        //检查数据是否有效
        private bool TestEffect()
        {
            bool returnValue = true;
            string msg = "";

            //正文字体
            if(CnFont.SelectedItem == null) 
            {
                CnFont.SelectedItem = clsConfig.ReadConfig<string>("SettingFormatForm", "CnFont", "仿宋_GB2312");
                msg = "正文字体选择有误,已恢复默认";
                returnValue = false;
            }

            //数字字体
            if (NumFont.SelectedItem == null) 
            {
                NumFont.SelectedItem = clsConfig.ReadConfig<string>("SettingFormatForm", "NumFont", "Arial Narrow");
                if (returnValue)
                {
                    msg = "数字字体选择有误,已恢复默认";
                }
                else
                {
                    msg = msg + Environment.NewLine + "数字字体选择有误,已恢复默认";
                }
                returnValue = false;
            }

            //字号
            if (FontSize.SelectedItem == null) 
            {
                FontSize.SelectedItem = clsConfig.ReadConfig<string>("SettingFormatForm", "FontSize", "小四");
                if (returnValue)
                {
                    msg = "正文字号选择有误,已恢复默认";
                }
                else
                {
                    msg = msg + Environment.NewLine + "正文字号选择有误,已恢复默认";
                }
                returnValue = false;
            }

            //行间距
            if (RowSpace.SelectedItem == null) 
            {
                RowSpace.SelectedItem = clsConfig.ReadConfig<string>("SettingFormatForm", "RowSpace", "单倍行距");
                if (returnValue)
                {
                    msg = "正文行间距选择有误,已恢复默认";
                }
                else
                {
                    msg = msg + Environment.NewLine + "正文行间距选择有误,已恢复默认";
                }
                returnValue = false;
            }

            //表格行间距
            if (TableRowSpace.SelectedItem == null) 
            {
                TableRowSpace.SelectedItem = clsConfig.ReadConfig<string>("SettingFormatForm", "TableRowSpace", "单倍行距");
                if (returnValue)
                {
                    msg = "表格行间距选择有误,已恢复默认";
                }
                else
                {
                    msg = msg + Environment.NewLine + "表格行间距选择有误,已恢复默认";
                }
                returnValue = false;
            }

            //表格字号
            if (TableFontSize.SelectedItem == null) 
            {
                TableFontSize.SelectedItem = clsConfig.ReadConfig<string>("SettingFormatForm", "TableFontSize", "小四");
                if (returnValue)
                {
                    msg = "表格字号选择有误,已恢复默认";
                }
                else
                {
                    msg = msg + Environment.NewLine + "表格字号选择有误,已恢复默认";
                }
                returnValue = false;
            }

            if (!returnValue) { MessageBox.Show(msg); }
            //返回值
            return returnValue;
        }
    }
}
