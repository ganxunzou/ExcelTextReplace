using System;
using System.Collections.Generic;
using System.Text;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;


namespace ExcelTextReplace
{
    class Program
    {
        /// <summary>
        /// excel文本替换程序
        /// </summary>
        /// <param name="args"></param>
        static void Main(string[] args)
        {
            
            string source ="c:\\123.xls";
            string target ="c:\\456.xls";
            string dataPath = "c:\\a.json";

            bool isExcel03 = false;

            try
            {
                string suffix = source.Substring(source.LastIndexOf(".")+1).ToLower();

                Console.WriteLine("suffix:{0}", suffix);

                if (suffix.Equals("xlsx") || suffix.Equals("xls"))
                {
                    //TODO:后缀名不正确,输出异常
                }
                else
                {
                    if (suffix.Equals("xls"))
                    {
                        isExcel03 = true;
                    }
                    else
                    {
                        isExcel03 = false;
                    }
                }

                String content = FileHelper.Read(dataPath);

                List<Dictionary<string,string>> replaceDatas = JsonTools.JsonToList(content);

                if (isExcel03)
                {
                    Excel2003.replaceTextAllSheet(source, target, replaceDatas);
                }
                else
                {
                    Excel2007.replaceTextAllSheet(source, target, replaceDatas);
                }
              
            }
            catch (System.Exception ex)
            {
                //TODO 异常退出处理
                Console.WriteLine(ex.Message);
            }

            Console.ReadKey();
            
        }
    }
}
