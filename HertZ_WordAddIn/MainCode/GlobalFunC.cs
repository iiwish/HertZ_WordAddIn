using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace HertZ_WordAddIn
{
    class GlobalFunC
    {

        /// <summary>
        /// 判断字符串是否是字母
        /// </summary>
        public bool IsLetter(string str)
        {
            bool returnValue;
            if (System.Text.RegularExpressions.Regex.IsMatch(str, @"(?i)^[A-Za-z]+$"))
            {
                returnValue = true;
            }
            else
            {
                returnValue = false;
            }
            return returnValue;
        }

        /// <summary>
        /// 判断字符串是否是数字
        /// </summary>
        public bool IsNumber(string str)
        {
            bool returnValue;
            try
            {
                double OutN = double.Parse(str);
                returnValue = true;
            }
            catch
            {
                returnValue = false;
            }
            return returnValue;
        }

        /// <summary>
        /// 将object转换为double,保留四位小数
        /// </summary>
        /// <param name="Value"></param>
        /// <returns></returns>
        public double TD(object Value)
        {
            double returnValue = 0d;
            if (Value == null)
            {
                return Math.Round(returnValue, 4);
            }
            string inputValue = Value.ToString();
            double.TryParse(inputValue, out returnValue);
            returnValue = Math.Round(returnValue, 4);
            return returnValue;
        }

        /// <summary>
        /// 将object转换为int
        /// </summary>
        /// <param name="Value"></param>
        /// <returns></returns>
        public int TI(object Value)
        {
            int returnValue = 0;
            if (Value == null)
            {
                return 0;
            }
            string inputValue = Value.ToString();
            int.TryParse(Value.ToString(), out returnValue);
            return returnValue;
        }

        /// <summary>
        /// 返回Link域中的文件区域
        /// </summary>
        /// <param name="CodeText"></param>
        /// <returns></returns>
        public string LinkArea(string CodeText)
        {
            string TempStr = CodeText.Split('!')[1];
            TempStr = TempStr.Split('"')[0];
            return TempStr;
        }

        /// <summary>
        /// 返回Link域中的表格名称
        /// </summary>
        /// <param name="CodeText"></param>
        /// <returns></returns>
        public string LinkSheet(string CodeText)
        {
            string TempStr = CodeText.Split('!')[0];//CodeText.Split('"')[3];
            TempStr = TempStr.Substring(TempStr.LastIndexOf('"') + 1);
            return TempStr;
        }

        /// <summary>
        /// 返回Link域中的路径
        /// </summary>
        /// <param name="CodeText"></param>
        /// <returns></returns>
        public string LinkPath(string CodeText)
        {
            string TempStr;
            if (CodeText.Length - CodeText.Replace("\"","").Length == 2)
            {
                TempStr = CodeText.Split('"')[0];
                TempStr = TempStr.Substring(0, TempStr.Length - 1);
                TempStr = TempStr.Replace(" LINK Excel.Sheet.12 ", "");
                TempStr = TempStr.Replace(" LINK Excel.Sheet.8 ", "");
                TempStr = TempStr.Replace(@"\\", @"\");
            }
            else
            {
                TempStr = CodeText.Split('"')[1];
                TempStr = TempStr.Replace(@"\\", @"\");
            }
            return TempStr;
        }

        /// <summary>
        /// 修改Link域中的路径
        /// </summary>
        /// <param name="CodeText"></param>
        /// <returns></returns>
        public string LinkPath(string CodeText,string NewPath)
        {
            string TempStr;
            if (CodeText.Length - CodeText.Replace("\"", "").Length == 2)
            {
                TempStr = CodeText.Split('"')[1] + '"' + CodeText.Split('"')[2];
                TempStr = " LINK Excel.Sheet.12 " + NewPath.Replace(@"\", @"\\") + " \"" + TempStr;
            }
            else
            {
                TempStr = CodeText.Split('"')[3] + '"' + CodeText.Split('"')[4];
                TempStr = " LINK Excel.Sheet.12 " + NewPath.Replace(@"\", @"\\") + " \"" + TempStr;
            }
            return TempStr;
        }

        /// <summary>
        /// 判断文件是否被占用
        /// </summary>
        /// <param name="fileName"></param>
        /// <returns></returns>
        public bool IsFileInUse(string fileName)
        {
            bool inUse = true;

            FileStream fs = null;
            try
            {

                fs = new FileStream(fileName, FileMode.Open, FileAccess.Read,

                FileShare.None);

                inUse = false;
            }
            catch
            {

            }
            finally
            {
                if (fs != null)

                    fs.Close();
            }
            return inUse;//true表示正在使用,false没有使用
        }

        /// <summary>
        /// 判断工作表是否存在
        /// </summary>
        public bool SheetExist(Excel.Workbook wbk, string SheetName)
        {
            bool returnValue = false;

            foreach(Excel.Worksheet wst in wbk.Worksheets)
            {
                if(wst.Name == SheetName)
                {
                    returnValue = true;
                    break;
                }
            }

            return returnValue;
        }

        /// <summary>
        /// 返回指定列的行数
        /// </summary>
        public int AllRows(Excel.Worksheet wst,string ColumnName = "A", int ColumnsTotal = 1)
        {
            int returnValue = 0;
            int StartColumn = CNumber(ColumnName);
            int NewRows;
            String Column;

            for (int i = StartColumn; i < StartColumn + ColumnsTotal; i++)
            {
                Column = CName(i);
                NewRows = ((Excel.Range)(wst.Cells[wst.Rows.Count, Column])).End[Excel.XlDirection.xlUp].Row;
                returnValue = Math.Max(returnValue, NewRows);
            }

            return returnValue;
        }

        /// <summary>
        /// 返回指定行的列数
        /// </summary>
        public int AllColumns(Excel.Worksheet wst,int RowName = 1, int RowsTotal = 1)
        {
            int returnValue = 0;
            int NewColumns;

            for (int i = RowName; i < RowName + RowsTotal; i++)
            {
                NewColumns = ((Excel.Range)(wst.Cells[i, "IV"])).End[Excel.XlDirection.xlToLeft].Column;
                returnValue = Math.Max(returnValue, NewColumns);
            }

            return returnValue;
        }

        /// <summary>
        /// 数字转列字母
        /// </summary>
        public string CName(int ColumnNumber)
        {
            int dividend = ColumnNumber;
            string columnName = String.Empty;
            int modulo;

            while (dividend > 0)
            {
                modulo = (dividend - 1) % 26;
                columnName = Convert.ToChar(65 + modulo).ToString() + columnName;
                dividend = (int)((dividend - modulo) / 26);
            }

            return columnName;
        }

        /// <summary>
        /// 列名转换数字
        /// </summary>
        public int CNumber(string ColumnName)
        {
            int index = 0;
            char[] chars = ColumnName.ToUpper().ToCharArray();
            for (int i = 0; i < chars.Length; i++)
            {
                index += ((int)chars[i] - (int)'A' + 1) * (int)Math.Pow(26, chars.Length - i - 1);
            }
            return index;
        }

        /// <summary>
        /// 将object转换为string
        /// </summary>
        /// <param name="Value"></param>
        /// <returns></returns>
        public string TS(object Value)
        {
            string returnValue = "";
            if (Value == null)
            {
                return returnValue;
            }
            returnValue = Value.ToString();
            return returnValue;
        }

    }
}
