using System;
using System.Collections.Generic;
using System.Text;
using System.Text.RegularExpressions;

namespace ExcelTextReplace
{
    public class ExcelTools
    {

        private static String wordStr = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";

        public static int getWordIndex(char word)
        {
            int index = wordStr.IndexOf(word);
            if (index >= 0)
                return index + 1;
            else
                throw new ExcelTextReplaceException("列[" + word + "不存在]");
        }

        /// <summary>
        /// 获取cell位置
        /// A		26*0  + 1
        /// AA		26*1  + 1
        /// BA		26*2  + 1
        /// ZA		26*26 + 1
        /// AAA		26*26*1 + 26*1 + 1
        /// ABA		26*26*1 + 26*2 + 1
        /// BAA		26*26*2	+ 26*1 + 1
        /// 返回结果:[rowIndex,colIndex]
        /// </summary>
        /// <param name="pstr"></param>
        /// <returns></returns>
        public static int[] getCellPoint(string pstr)
        {
            int[] indexs = new int[2];

            try
            {
                //示例：B24 ：column:1,row:24 ; AB24 : column:28,row:24
                pstr = pstr.ToUpper();
                Regex rowReg = new Regex(@"[0-9]+");
                int rowIndex = int.Parse(rowReg.Match(pstr).Value.Trim());

                Regex colReg = new Regex(@"[A-Z]+");
                string columnStr = colReg.Match(pstr).Value.Trim();
                int colIndex = 0;
                int i = columnStr.Length;

                foreach (char word in columnStr)
                {
                    if (i == 1)
                    {
                        colIndex += getWordIndex(word);
                    }
                    else
                    {
                        colIndex += ((int)Math.Pow(26, (i - 1))) * getWordIndex(word);
                    }

                    i--;
                }

                indexs[0] = rowIndex;
                indexs[1] = colIndex;
            }
            catch (System.Exception ex)
            {
                throw new ExcelTextReplaceException("单元格定位字符串[" + pstr + "]不正确!!" + ex.Message, ex);
            }

            return indexs;
        }
    }
}
