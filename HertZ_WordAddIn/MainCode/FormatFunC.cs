using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace HertZ_WordAddIn
{
    class FormatFunC
    {

        /// <summary>
        /// 根据字体名称查字号
        /// </summary>
        /// <param name="OriginalFontSize"></param>
        /// <returns></returns>
        public decimal FontSize(string OriginalFontSize)
        {
            decimal ReturnValue = 0;
            Dictionary<string, decimal> FontSizeDic = new Dictionary<string, decimal>
            {
                {"初号", 42},{"小初",36},{"一号", 26},{"小一",24},{"二号",22},{"小二",18},
                {"三号", 16},{"小三",15},{"四号",14},{"小四",12},{"五号",10.5m},{"小五",9},
                {"六号", 7.5m},{"小六", 6.5m},{"七号",5.5m},{"八号",5},{"5",5},{"5.5",5.5m},
                {"6.5",6.5m},{"7.5",7.5m},{"8",8},{"9",9},{"10",10},{"10.5",10.5m},
                {"11",11},{"12",12},{"14",14},{"16",16},{"18",18},{"20",20},
                {"22",22},{"24",24},{"26",26},{"28",28},{"36",36},{"48",48},{"72",72 }
            };
            if (FontSizeDic.ContainsKey(OriginalFontSize))
            {
                ReturnValue = FontSizeDic[OriginalFontSize];
            }
            return ReturnValue;
        }

        /// <summary>
        /// 根据行间距名称查行间距
        /// </summary>
        /// <param name="OriginalRowSpace"></param>
        /// <returns></returns>
        public decimal RowSpace(string OriginalRowSpace)
        {
            decimal ReturnValue = 0;
            Dictionary<string, decimal> RowSpaceDic = new Dictionary<string, decimal>
            {
                {"单倍行距",1 },{"1.15倍行距",1.15m },{"1.5倍行距",1.5m},{"2倍行距",2},{"2.5倍行距",2.5m},{"3倍行距",3}
            };
            if (RowSpaceDic.ContainsKey(OriginalRowSpace))
            {
                ReturnValue = RowSpaceDic[OriginalRowSpace];
            }
            return ReturnValue;
        }
    }
}
