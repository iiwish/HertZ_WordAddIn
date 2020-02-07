using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

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
        /// 返回Link域中的文件路径
        /// </summary>
        /// <param name="CodeText"></param>
        /// <returns></returns>
        public string LinkPath(string CodeText)
        {
            string TempStr = CodeText.Split('"')[1];
            TempStr = TempStr.Replace(@"\\", @"\");
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
    }
}
