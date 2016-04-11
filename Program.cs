using System;
using System.Collections.Generic;


namespace ExcelTextReplace
{
    class Program
    {
        /// <summary>
        /// excel文本替换程序
        /// exitCode:0  正常
        /// exitCode:1  错误
        /// </summary>
        /// <param name="args"></param>
        public static int Main(string[] args)
        {
            string source =  "c:\\123.xls";// "c:\\1234.xlsx";
            string target = "c:\\12345.xls";// "c:\\12345.xlsx";
            string dataPath = "c:\\data.json";

            try
            {
                /*string source = args[0];// "c:\\1234.xlsx";
                string target = args[1];// "c:\\12345.xlsx";
                string dataPath = args[2];*/

                bool isExcel03 = false;

                string suffix = source.Substring(source.LastIndexOf(".") + 1).ToLower();

                if (suffix.Equals("xlsx") || suffix.Equals("xls"))
                {
                    
                    if (suffix.Equals("xls"))
                    {
                        isExcel03 = true;
                    }
                    else
                    {
                        isExcel03 = false;
                    }

                    if (suffix.Equals("xls"))
                    {
                        isExcel03 = true;
                    }
                    else
                    {
                        isExcel03 = false;
                    }

                    String content = FileHelper.Read(dataPath);
                    List<ReplaceDataVO> dataLs = JSONTools.parseJsonStrToVos(content);

                    if (isExcel03)
                    {
                        Excel2003.replaceTextAllSheet(source, target, dataLs);
                    }
                    else
                    {

                        //Excel2007.replaceTextAllSheet(source, target, replaceDatas);
                    }

                }
                else
                {
                    throw new ExcelTextReplaceException("文件后缀名不正确,应该是xlsx或者xls,当前的文件名后缀[" + suffix + "]");
                }

            }
            catch (System.Exception ex)
            {
                Console.WriteLine(ex.Message);
                return 1;
            }

            Console.ReadKey();
            return 0;
         
        }
    }
}
