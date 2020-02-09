using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Microsoft.Office.Tools.Ribbon;
using Word = Microsoft.Office.Interop.Word;
using Excel = Microsoft.Office.Interop.Excel;
using System.Threading;
using Microsoft.VisualBasic;
using System.IO;
using System.Runtime.InteropServices;
using System.Diagnostics;

namespace HertZ_WordAddIn
{
    public partial class HertZ_Ribbon
    {
        private Word.Application WordApp;
        private Word.Document WordDoc;

        private Excel.Application ExcelApp;

        //引入调整格式用的函数模块
        private readonly FormatFunC FormatFunC = new FormatFunC();
        private readonly GlobalFunC FunC = new GlobalFunC();

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
            //弹出窗体提示
            DialogResult IsWait = MessageBox.Show("如果文本行较多运行较为缓慢，进度可在Word左下角查看" + Environment.NewLine + "是否继续？", "请选择", MessageBoxButtons.YesNo);
            if(IsWait != DialogResult.Yes) { return; }

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
            //跳过的段落数
            int SkipPgNum = FunC.TI(clsConfig.ReadConfig<decimal>("SettingFormatForm", "SkipPgs", 0m));

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

                int i4 = 0;
                foreach (Word.Paragraph Pg in WordDoc.Paragraphs)
                {
                    i4++;
                    if(i4 <= SkipPgNum ) { continue; }
                    if (Pg.Range.Information[Word.WdInformation.wdWithInTable]) { continue; }
                    PgLevel = Pg.OutlineLevel;

                    //如果勾选了标题段落，调整段前段后行间距
                    if (TitleGroupCheck)
                    {
                        //段前间距
                        Pg.SpaceBefore = 0;
                        Pg.LineUnitBefore = float.Parse(BeforeMainBody.ToString());
                        //段后间距
                        Pg.SpaceAfter = 0;
                        Pg.LineUnitAfter = float.Parse(AfterMainBody.ToString());
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

                            if (FirstTitleCheck)
                            {
                                if (PgLevel == Word.WdOutlineLevel.wdOutlineLevel1)
                                {
                                    Pg.SpaceBefore = 0;
                                    Pg.LineUnitBefore = 0.5f;
                                }
                            }
                        }
                    }

                    //调整表格下一行间距
                    if (TitleGroupCheck && FirstTitleCheck && i4 > 1)
                    {
                        if (WordDoc.Paragraphs[i4 - 1].Range.Information[Word.WdInformation.wdWithInTable])
                        {
                            Pg.SpaceBefore = 0;
                            Pg.LineUnitBefore = 0.5f;
                        }
                    }

                    //显示进度
                    WordApp.StatusBar = "当前进度:" + Math.Round((i4*100d/i3),2) + "%";
                    
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
            //弹出窗体提示
            DialogResult IsWait = MessageBox.Show("如果表格较多运行较为缓慢，进度可在Word左下角查看" + Environment.NewLine + "是否继续？", "请选择", MessageBoxButtons.YesNo);
            if (IsWait != DialogResult.Yes) { return; }

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
                            WordApp.Selection.ParagraphFormat.FirstLineIndent = WordApp.CentimetersToPoints(0);
                        }
                        WordApp.Selection.ParagraphFormat.SpaceBefore = 0;
                        WordApp.Selection.ParagraphFormat.LineUnitBefore = float.Parse(TableBeforeMainBody.ToString());
                        WordApp.Selection.ParagraphFormat.SpaceAfter = 0;
                        WordApp.Selection.ParagraphFormat.LineUnitAfter = float.Parse(TableAfterMainBody.ToString());
                        WordApp.Selection.ParagraphFormat.LineSpacing = WordApp.LinesToPoints(float.Parse(TableRowSpace.ToString()));
                        WordApp.Selection.ParagraphFormat.DisableLineHeightGrid = -1;//行距设置中，不勾选“如果定义了文档网格，则于网格对齐”
                        WordApp.Selection.ParagraphFormat.AutoAdjustRightIndent = -1;//行距设置中，不勾选“如果定义了文档网格，则自动调整右缩进”
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
                        //表格字体
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

        /// <summary>
        /// 修改附注关联的Excel路径
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ChangeExcelPath_Click(object sender, RibbonControlEventArgs e)
        {
            //弹出窗体提示
            DialogResult IsWait = MessageBox.Show("请在使用HertZ_Excel插件生成的Word附注中使用该功能" + Environment.NewLine + "是否继续？", "请选择", MessageBoxButtons.YesNo);
            if (IsWait != DialogResult.Yes) { return; }

            //选择新Excel文件
            string FilePath;
            OpenFileDialog dialog = new OpenFileDialog();
            dialog.Multiselect = false;//该值确定是否可以选择多个文件
            dialog.Title = "请选择（加工完的）久其导出的Excel文件";
            dialog.Filter = "Excel文件(*.xls;*.xlsx;*.xlsm)|*.xls;*.xlsx;*.xlsm";
            if (dialog.ShowDialog() == DialogResult.OK)
            {
                FilePath = dialog.FileName;
            }
            else
            {
                return;
            }

            WordApp = Globals.ThisAddIn.Application;
            WordDoc = WordApp.ActiveDocument;

            WordApp.ScreenUpdating = false;//关闭屏幕刷新

            int i5 = WordDoc.Fields.Count;
            int i4 = 0;
            foreach (Word.Field TempField in WordDoc.Fields)
            {
                i4++;

                if(TempField.Type != Word.WdFieldType.wdFieldLink) 
                {
                    //显示进度
                    WordApp.StatusBar = "当前进度:" + Math.Round((i4 * 100d / i5), 2) + "%";

                    continue; 
                }

                //TempField.Locked = true;
                TempField.LinkFormat.SourceFullName = ExcelApp.ActiveWorkbook.FullName;

                //显示进度
                WordApp.StatusBar = "当前进度:" + Math.Round((i4 * 100d / i5), 2) + "%";
            }

            WordApp.ScreenUpdating = true;//打开屏幕刷新

            MessageBox.Show("久其路径修改完成！");

        }

        /// <summary>
        /// 更新全部域
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void UpdateLink_Click(object sender, RibbonControlEventArgs e)
        {
            //弹出窗体提示
            DialogResult IsWait = MessageBox.Show("请在使用HertZ_Excel插件生成的Word附注中使用该功能" + Environment.NewLine + "是否继续？", "请选择", MessageBoxButtons.YesNo);
            if (IsWait != DialogResult.Yes) { return; }

            WordApp = Globals.ThisAddIn.Application;
            WordDoc = WordApp.ActiveDocument;

            Excel.Workbook WBK = null;
            Word.Table TempTable;
            bool QuitExcel = false;
            bool CloseWBK = false;
            int ExcelRows;
            int ExcelColumns;
            int WordRows;
            int WordColumns;
            int TempInt;

            //WordDoc.Fields.Update();
            WordApp.ScreenUpdating = false;//关闭屏幕刷新

            WordApp.StatusBar = "当前进度:0.00%";

            string TempStr = null;
            //检查Excel文件
            foreach (Word.Field TempField in WordDoc.Fields)
            {
                if (TempField.Type == Word.WdFieldType.wdFieldLink)
                {
                    TempStr = TempField.LinkFormat.SourceFullName;
                    //检查文件是否存在
                    if (!File.Exists(TempStr))
                    {
                        MessageBox.Show("未发现" + TempStr +"，请检查");
                        return;
                    }
                    break;
                }
            }

            if(TempStr == null) { MessageBox.Show("未发现表格域，请在久其生成的附注中使用该功能"); return; }

            //获取区域的行列数
            if (FunC.IsFileInUse(TempStr))//如果目标文件已被打开
            {
                try
                {
                    ExcelApp = (Excel.Application)Marshal.GetActiveObject("Excel.Application");
                    WBK = ExcelApp.Workbooks[Path.GetFileName(TempStr)];
                }
                catch
                {
                    MessageBox.Show("链接的Excel文件已被后台程序打开，请检查并清理后台程序");
                    return;
                }
            }
            else
            {
                CloseWBK = true;
                try
                {
                    ExcelApp = (Excel.Application)Marshal.GetActiveObject("Excel.Application");
                    QuitExcel = false;
                }
                catch
                {
                    //弹出窗体提示
                    IsWait = MessageBox.Show("当前版本插件建议打开Excel程序再刷新域" + Environment.NewLine + "否则可能需要清理后台Excel程序，是否继续？", "请选择", MessageBoxButtons.YesNo);
                    if (IsWait != DialogResult.Yes) { WordApp.ScreenUpdating = true; return; }

                    ExcelApp = new Excel.Application();
                    ExcelApp.Visible = false;
                    QuitExcel = true;
                }

                WBK = ExcelApp.Workbooks.Open(TempStr);
            }

            //遍历WBK工作表名，加入字典
            Dictionary<string, string> wstDic = new Dictionary<string, string> { };
            object[,] ORG ;
            foreach (Excel.Worksheet wst in WBK.Worksheets)
            {
                ExcelColumns = FunC.AllColumns(wst, 4, 2);
                ExcelRows = FunC.AllRows(wst, "A", 2);
                ORG = wst.Range["A1:A" + ExcelRows].Value2;
                for(int i = ExcelRows; i > Math.Max(ExcelRows - 5, 4); i--)
                {
                    if(FunC.TS(ORG[i,1]).Contains("注：") || FunC.TS(ORG[i, 1]).Contains("注:"))
                    {
                        ExcelRows = i - 1;
                    }
                }

                wstDic.Add(wst.Name, ExcelRows + ":" + ExcelColumns);
            }

            //关闭workbook
            if (CloseWBK) { WBK.Close(); }
            //退出Excel程序
            if (QuitExcel)
            {
                ExcelApp.Visible = true;
                ExcelApp.Quit();
                ExcelApp = null;
            }

            int i5 = WordDoc.Fields.Count;
            int i4 = 0;
            foreach (Word.Field TempField in WordDoc.Fields)
            {
                i4++;

                if (TempField.Type != Word.WdFieldType.wdFieldLink)
                {
                    //显示进度
                    WordApp.StatusBar = "当前进度:" + Math.Round((i4 * 100d / i5), 2) + "%";

                    continue;
                }

                TempStr = TempField.Code.Text;

                if (TempStr.Contains('"'))
                {
                    if (wstDic.ContainsKey(FunC.LinkSheet(TempStr)))
                    {
                        TempField.Select();
                        TempTable = WordApp.Selection.Tables[1];
                        WordRows = TempTable.Rows.Count;
                        WordColumns = TempTable.Columns.Count;
                        ExcelRows = int.Parse(wstDic[FunC.LinkSheet(TempStr)].Split(':')[0]);
                        ExcelColumns = int.Parse(wstDic[FunC.LinkSheet(TempStr)].Split(':')[1]);

                        TempInt = WordRows - ExcelRows;
                        //检查行数
                        if (TempInt < 0)
                        {
                            for (int i = 1; i <= Math.Abs(TempInt); i++)
                            {
                                TempTable.Cell(Math.Max(WordRows - 1, 1), 1).Range.Rows.Add(TempTable.Cell(Math.Max(WordRows - 1, 1), 1).Range.Rows);
                            }
                        }
                        else if (TempInt > 0)
                        {
                            for (int i = TempInt; i >= 1; i--)
                            {
                                TempTable.Cell(Math.Max(WordRows - i, 1), 1).Range.Rows.Delete();
                            }
                        }

                        TempInt = WordColumns - ExcelColumns;
                        //检查列数
                        if (TempInt < 0)
                        {
                            for (int i = 1; i <= Math.Abs(TempInt); i++)
                            {
                                TempTable.Cell(ExcelRows, Math.Max(WordColumns - 2, 1)).Range.Columns.Add(TempTable.Cell(ExcelRows, Math.Max(WordColumns - 2, 1)).Range.Columns);
                            }
                        }
                        else if (TempInt > 0)
                        {
                            for (int i = TempInt; i >= 1; i--)
                            {
                                TempTable.Cell(ExcelRows, Math.Max(WordColumns - 1, 1)).Range.Columns.Delete();
                            }
                        }


                        TempField.Update();
                    }
                }

                //显示进度
                WordApp.StatusBar = "当前进度:" + Math.Round((i4 * 100d / i5), 2) + "%";
            }

            WordApp.ScreenUpdating = true;//打开屏幕刷新

            MessageBox.Show("域更新完成！");

        }

        /// <summary>
        /// 断开全部表格的域链接
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void DiscAllLink_Click(object sender, RibbonControlEventArgs e)
        {
            //弹出窗体提示
            DialogResult IsWait = MessageBox.Show("请在使用HertZ_Excel插件生成的Word附注中使用该功能" + Environment.NewLine + "是否继续？", "请选择", MessageBoxButtons.YesNo);
            if (IsWait != DialogResult.Yes) { return; }

            WordApp = Globals.ThisAddIn.Application;
            WordDoc = WordApp.ActiveDocument;

            int i5 = WordDoc.Fields.Count;
            int i4 = 0;
            //WordDoc.Fields.Update();
            WordApp.ScreenUpdating = false;//关闭屏幕刷新
            foreach (Word.Field TempField in WordDoc.Fields)
            {
                i4++;
                if (TempField.Type != Word.WdFieldType.wdFieldLink)
                {
                    //显示进度
                    WordApp.StatusBar = "当前进度:" + Math.Round((i4 * 100d / i5), 2) + "%";
                    continue;
                }

                //断开链接
                TempField.Unlink();
                //显示进度
                WordApp.StatusBar = "当前进度:" + Math.Round((i4 * 100d / i5), 2) + "%";
            }

            WordApp.ScreenUpdating = true;//打开屏幕刷新

            MessageBox.Show("Link域已全部断开链接！");

        }

        /// <summary>
        /// 断开当前表格的域链接
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void DiscLink_Click(object sender, RibbonControlEventArgs e)
        {
            WordApp = Globals.ThisAddIn.Application;
            WordDoc = WordApp.ActiveDocument;

            //选中整个表格
            Word.Table TempTable = WordApp.Selection.Tables[1];
            TempTable.Select();

            //获取域
            WordApp.Selection.Next(Word.WdUnits.wdWord, 2).Select();
            WordApp.Selection.PreviousField();
            WordApp.Selection.Fields[1].Unlink();

            TempTable.Select();
        }

        /// <summary>
        /// 版本信息
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void VerInfo_Click(object sender, RibbonControlEventArgs e)
        {
            Form InfoForm = new VerInfo
            {
                StartPosition = FormStartPosition.CenterScreen
            };
            InfoForm.Show();
        }

        /// <summary>
        /// 刷新当前域
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void UpdateOneLink_Click(object sender, RibbonControlEventArgs e)
        {
            
            WordApp = Globals.ThisAddIn.Application;
            WordDoc = WordApp.ActiveDocument;

            Word.Field TempField;
            Word.Table TempTable;
            int WordRows = 0;
            int WordColumns = 0;
            int ExcelRows  = 0;
            int ExcelColumns = 0;
            string TempStr;

            //如果选中的区域不包含表格则退出
            if (WordApp.Selection.Tables.Count == 0) { return; }

            //如果选中的区域不在域中则退出
            if (!WordApp.Selection.Information[Word.WdInformation.wdInFieldResult]) { return; }

            WordApp.ScreenUpdating = false;//关闭屏幕刷新

            //选中整个表格
            TempTable = WordApp.Selection.Tables[1];
            TempTable.Select();

            //获取Word中表格行列数
            WordRows = TempTable.Rows.Count;
            WordColumns = TempTable.Columns.Count;

            //获取域
            WordApp.Selection.Next(Word.WdUnits.wdWord,2).Select();
            WordApp.Selection.PreviousField();
            TempField = WordApp.Selection.Fields[1];
            
            TempStr = TempField.LinkFormat.SourceFullName;

            //检查文件是否存在
            if (!File.Exists(TempStr))
            {
                MessageBox.Show("未发现域链接的Excel文件，请检查！");
            }

            Excel.Application ExcelApp = null;
            Excel.Workbook WBK = null;
            Excel.Worksheet WST = null;

            //获取区域的行列数
            if (FunC.IsFileInUse(TempStr))//如果目标文件已被打开
            {
                try
                {
                    ExcelApp = (Excel.Application)Marshal.GetActiveObject("Excel.Application");
                    WBK = ExcelApp.Workbooks[Path.GetFileName(TempStr)];
                }
                catch
                {
                    MessageBox.Show("链接的Excel文件已被后台程序打开，请检查并清理后台程序");
                    return;
                }

                //如果没有工作表
                if (!FunC.SheetExist(WBK, FunC.LinkSheet(TempField.Code.Text)))
                {
                    MessageBox.Show("链接的Excel文件中未发现该工作表，请检查");
                    return;
                }

                WST = WBK.Worksheets[FunC.LinkSheet(TempField.Code.Text)];

                try
                {
                    ExcelRows = WST.Range[FunC.LinkArea(TempField.Code.Text)].Rows.Count;
                    ExcelColumns = WST.Range[FunC.LinkArea(TempField.Code.Text)].Columns.Count;
                }
                catch
                {
                    MessageBox.Show("链接的Excel文件中未发现该域的命名区域，请检查");
                    return;
                }
            }
            else
            {
                bool QuitExcel;
                try
                {
                    ExcelApp = (Excel.Application)Marshal.GetActiveObject("Excel.Application");
                    QuitExcel = false;
                }
                catch
                {
                    //弹出窗体提示
                    DialogResult IsWait = MessageBox.Show("当前版本插件建议打开Excel程序再刷新域" + Environment.NewLine + "否则可能需要清理后台Excel程序，是否继续？", "请选择", MessageBoxButtons.YesNo);
                    if (IsWait != DialogResult.Yes) { WordApp.ScreenUpdating = true; return; }

                    ExcelApp = new Excel.Application();
                    ExcelApp.Visible = false;
                    QuitExcel = true;
                }

                WBK = ExcelApp.Workbooks.Open(TempStr);
                
                //如果没有工作表
                if (!FunC.SheetExist(WBK, FunC.LinkSheet(TempField.Code.Text)))
                {
                    MessageBox.Show("链接的Excel文件中未发现该工作表，请检查");
                    return;
                }

                WST = WBK.Worksheets[FunC.LinkSheet(TempField.Code.Text)];

                try
                {
                    ExcelRows = WST.Range[FunC.LinkArea(TempField.Code.Text)].Rows.Count;
                    ExcelColumns = WST.Range[FunC.LinkArea(TempField.Code.Text)].Columns.Count;
                }
                catch
                {
                    MessageBox.Show("链接的Excel文件中未发现该域的命名区域，请检查");
                    return;
                }
                
                WBK.Close();
                WBK = null;
                WST = null;

                //退出Excel程序
                if (QuitExcel) 
                {
                    ExcelApp.Visible = true;
                    ExcelApp.Quit();
                    ExcelApp = null;
                }

            }

            int TempInt = WordRows - ExcelRows;
            //检查行数
            if (TempInt < 0)
            {
                for (int i = 1; i <= Math.Abs(TempInt); i++)
                {
                    TempTable.Cell(Math.Max(WordRows - 1, 1), 1).Range.Rows.Add(TempTable.Cell(Math.Max(WordRows - 1, 1), 1).Range.Rows);
                }
            }
            else if (TempInt > 0)
            {
                for (int i = TempInt; i >= 1; i--)
                {
                    TempTable.Cell(Math.Max(WordRows - i, 1), 1).Range.Rows.Delete();
                }
            }

            TempInt = WordColumns - ExcelColumns;
            //检查列数
            if (TempInt < 0)
            {
                for (int i = 1; i <= Math.Abs(TempInt); i++)
                {
                    TempTable.Cell(ExcelRows, Math.Max(WordColumns - 2, 1)).Range.Columns.Add(TempTable.Cell(ExcelRows, Math.Max(WordColumns - 2, 1)).Range.Columns);
                }
            }
            else if (TempInt > 0)
            {
                for (int i = TempInt; i >= 1; i--)
                {
                    TempTable.Cell(ExcelRows, Math.Max(WordColumns - 1, 1)).Range.Columns.Delete();
                }
            }

            TempField.Update();

            TempField.Select();

            GC.Collect();
            WordApp.ScreenUpdating = true;//打开屏幕刷新
        }
    }
}
