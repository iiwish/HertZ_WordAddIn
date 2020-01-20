using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Microsoft.Office.Tools.Ribbon;
using Word = Microsoft.Office.Interop.Word;
using System.Threading;

namespace HertZ_WordAddIn
{
    public partial class HertZ_Ribbon
    {
        private Word.Application WordApp;
        private Word.Document WordDoc;

        //引入调整格式用的函数模块
        private readonly FormatFunC FormatFunC = new FormatFunC();

        //用于显示进度条窗体
        private delegate bool IncreaseHandle(int nValue);

        private void HertZ_Ribbon_Load(object sender, RibbonUIEventArgs e)
        {

        }

        /// <summary>
        /// 修改文本样式
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void TextFormat_Click(object sender, RibbonControlEventArgs e)
        {
            WordApp = Globals.ThisAddIn.Application;
            WordDoc = WordApp.ActiveDocument;

            WordApp.ScreenUpdating = false;//关闭屏幕刷新

            //读取配置文件
            ClsThisAddinConfig clsConfig = new ClsThisAddinConfig(Environment.GetFolderPath(Environment.SpecialFolder.CommonDocuments));

            //正文字体字号
            bool FontGroupCheck = clsConfig.ReadConfig<bool>("SettingFormatForm", "FontGroupCheck", true);
            string CnFont = clsConfig.ReadConfig<string>("SettingFormatForm", "CnFont", "仿宋_GB2312");
            string NumFont = clsConfig.ReadConfig<string>("SettingFormatForm", "NumFont", "Arial Narrow");
            decimal FontSize = FormatFunC.FontSize(clsConfig.ReadConfig<string>("SettingFormatForm", "FontSize", "小四"));

            //标题行间距
            bool TitleGroupCheck = clsConfig.ReadConfig<bool>("SettingFormatForm", "TitleGroupCheck", true);
            bool TitleIndentCkeck = clsConfig.ReadConfig<bool>("SettingFormatForm", "TitleIndentCkeck", true);
            bool FirstTitleCheck = clsConfig.ReadConfig<bool>("SettingFormatForm", "FirstTitleCheck", true);
            decimal BeforeMainBody = clsConfig.ReadConfig<decimal>("SettingFormatForm", "BeforeMainBody", 0m);
            decimal AfterMainBody = clsConfig.ReadConfig<decimal>("SettingFormatForm", "AfterMainBody", 0.9m);
            decimal RowSpace = FormatFunC.RowSpace(clsConfig.ReadConfig<string>("SettingFormatForm", "RowSpace", "单倍行距"));


            //段落大纲级别
            Word.WdOutlineLevel PgLevel;

            
            if (FontGroupCheck || TitleGroupCheck)
            {
                //在状态栏显示进度
                WordApp.StatusBar = "当前进度:0%";

                int i3 = WordDoc.Paragraphs.Count;

                if (FontGroupCheck)
                {
                    WordApp.Selection.WholeStory();
                    WordApp.Selection.Font.Name = CnFont;
                    WordApp.Selection.Font.Name = NumFont;
                }

                int i4 = 1;
                foreach (Word.Paragraph Pg in WordDoc.Paragraphs)
                {
                    if (Pg.Range.Information[Word.WdInformation.wdWithInTable]) { continue; }
                    PgLevel = Pg.OutlineLevel;

                    //如果勾选了标题段落，调整段前段后行间距
                    if (TitleGroupCheck)
                    {
                        //段前间距
                        Pg.SpaceBefore = float.Parse(BeforeMainBody.ToString());
                        //段后间距
                        Pg.SpaceAfter = float.Parse(AfterMainBody.ToString());
                        //段落行间距
                        Pg.LineSpacing = WordApp.LinesToPoints(float.Parse(RowSpace.ToString()));
                    }

                    if (PgLevel == Word.WdOutlineLevel.wdOutlineLevelBodyText)
                    {
                        if (FontGroupCheck)
                        {
                            Pg.Range.Font.Size = float.Parse(FontSize.ToString());
                        }
                    }
                    else
                    {
                        if (TitleGroupCheck)
                        {
                            if (TitleIndentCkeck)
                            {
                                Pg.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
                                Pg.CharacterUnitLeftIndent = -2;
                            }
                            if (PgLevel == Word.WdOutlineLevel.wdOutlineLevel1 && FirstTitleCheck)
                            {
                                Pg.SpaceBefore = 0.5f;
                            }
                        }
                    }

                    //显示进度
                    WordApp.StatusBar = "当前进度:" + Math.Round((i4*100d/i3),2) + "%";
                    i4++;
                }
            }

            WordApp.ScreenUpdating = true;//关闭屏幕刷新

        }

        /// <summary>
        /// 修改表格样式
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void TableFormat_Click(object sender, RibbonControlEventArgs e)
        {
            WordApp = Globals.ThisAddIn.Application;
            WordDoc = WordApp.ActiveDocument;

            WordApp.ScreenUpdating = false;//关闭屏幕刷新

            #region //读取配置文件
            //读取配置文件
            ClsThisAddinConfig clsConfig = new ClsThisAddinConfig(Environment.GetFolderPath(Environment.SpecialFolder.CommonDocuments));

            //表格间距
            bool SpaceGroupCheck = clsConfig.ReadConfig<bool>("SettingFormatForm", "SpaceGroupCheck", true);
            //首行缩进
            bool TableIndentCheck = clsConfig.ReadConfig<bool>("SettingFormatForm", "TableIndentCheck", true);
            //段前
            decimal TableBeforeMainBody = clsConfig.ReadConfig<decimal>("SettingFormatForm", "TableBeforeMainBody", 0m);
            //段后
            decimal TableAfterMainBody = clsConfig.ReadConfig<decimal>("SettingFormatForm", "TableAfterMainBody", 0m);
            //行间距
            decimal TableRowSpace = FormatFunC.RowSpace(clsConfig.ReadConfig<string>("SettingFormatForm", "TableRowSpace", "单倍行距"));

            //表格边框
            bool BorderGroupCheck = clsConfig.ReadConfig<bool>("SettingFormatForm", "BorderGroupCheck", true);
            //上下边框宽度1磅
            bool TableTBBorderCheck = clsConfig.ReadConfig<bool>("SettingFormatForm", "TableTBBorderCheck", true);
            //左右无边框
            bool TableLRBorderCheck = clsConfig.ReadConfig<bool>("SettingFormatForm", "TableLRBorderCheck", true);

            //表格其他选项
            bool OtherGroupCheck = clsConfig.ReadConfig<bool>("SettingFormatForm", "OtherGroupCheck", true);
            //垂直居中
            bool TableVerticalCenterCheck = clsConfig.ReadConfig<bool>("SettingFormatForm", "TableVerticalCenterCheck", true);
            //标题及合计行加粗
            bool TableTitleWiderCheck = clsConfig.ReadConfig<bool>("SettingFormatForm", "TableTitleWiderCheck", true);
            //表格字号
            decimal TableFontSize = FormatFunC.FontSize(clsConfig.ReadConfig<string>("SettingFormatForm", "TableFontSize", "小四"));
            #endregion

            if(SpaceGroupCheck || BorderGroupCheck || OtherGroupCheck)
            {
                
                //在状态栏显示进度
                WordApp.StatusBar = "当前进度:0%";
                int i3 = WordDoc.Tables.Count;

                int i4 = 1;
                foreach (Word.Table Tb in WordDoc.Tables)
                {
                    Tb.Select();

                    //间距
                    if (SpaceGroupCheck)
                    {
                        if (TableIndentCheck)
                        {
                            WordApp.Selection.ParagraphFormat.CharacterUnitLeftIndent = 0;
                            WordApp.Selection.ParagraphFormat.CharacterUnitRightIndent = 0;
                            WordApp.Selection.ParagraphFormat.CharacterUnitFirstLineIndent = 0;
                        }
                        WordApp.Selection.ParagraphFormat.SpaceBefore = float.Parse(TableBeforeMainBody.ToString());
                        WordApp.Selection.ParagraphFormat.SpaceAfter = float.Parse(TableAfterMainBody.ToString());
                        WordApp.Selection.ParagraphFormat.LineSpacing = WordApp.LinesToPoints(float.Parse(TableRowSpace.ToString()));
                        WordApp.Selection.ParagraphFormat.DisableLineHeightGrid = 0;
                    }

                    //边框
                    if (BorderGroupCheck)
                    {
                        //上下边框1磅
                        if (TableTBBorderCheck)
                        {
                            Tb.Borders[Word.WdBorderType.wdBorderTop].LineStyle = Word.WdLineStyle.wdLineStyleSingle;
                            Tb.Borders[Word.WdBorderType.wdBorderTop].LineWidth = Word.WdLineWidth.wdLineWidth100pt;
                            Tb.Borders[Word.WdBorderType.wdBorderTop].Color = Word.WdColor.wdColorAutomatic;

                            Tb.Borders[Word.WdBorderType.wdBorderBottom].LineStyle = Word.WdLineStyle.wdLineStyleSingle;
                            Tb.Borders[Word.WdBorderType.wdBorderBottom].LineWidth = Word.WdLineWidth.wdLineWidth100pt;
                            Tb.Borders[Word.WdBorderType.wdBorderBottom].Color = Word.WdColor.wdColorAutomatic;
                        }
                        //左右无边框
                        if (TableLRBorderCheck)
                        {
                            Tb.Borders[Word.WdBorderType.wdBorderLeft].LineStyle = Word.WdLineStyle.wdLineStyleNone;
                            Tb.Borders[Word.WdBorderType.wdBorderRight].LineStyle = Word.WdLineStyle.wdLineStyleNone;
                        }
                    }

                    //其他选项
                    if (OtherGroupCheck)
                    {
                        //垂直居中
                        if (TableVerticalCenterCheck)
                        {
                            WordApp.Selection.Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;
                        }
                        //
                        if (TableTitleWiderCheck)
                        {

                        }
                        //
                        Tb.Range.Font.Size = float.Parse(TableFontSize.ToString());
                    }

                    //显示进度
                    WordApp.StatusBar = "当前进度:" + Math.Round((i4 * 100d / i3), 2) + "%";
                    i4++;
                }
            }
            WordApp.ScreenUpdating = true;//关闭屏幕刷新
        }

        /// <summary>
        /// 设置默认格式
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void SettingFormat_Click(object sender, RibbonControlEventArgs e)
        {
            Form SettingFormatForm = new SettingFormatForm();
            SettingFormatForm.StartPosition = FormStartPosition.CenterScreen;
            SettingFormatForm.Show();
        }


    }
}
