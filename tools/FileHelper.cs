using System;
using System.Collections.Generic;
using System.Text;
using System.IO;

namespace ExcelTextReplace
{
    public class FileHelper
    {
        /// <summary>
        ///  读取文件
        ///  author ganxz
        ///  2016-04-08
        /// </summary>
        /// <param name="path"></param>
        /// <returns></returns>
        public static String Read(string path)
        {
            StreamReader sr = new StreamReader(path,Encoding.Default);
            String content = "";
            String line;
            while ((line = sr.ReadLine()) != null) 
            {
                content += line;
            }

            return content;
        }
    }
}
