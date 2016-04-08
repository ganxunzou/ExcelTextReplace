using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Text;
using System.IO;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using Newtonsoft.Json;

namespace ExcelTextReplace
{
    public class Excel2007
    {
        public static string replaceTextAllSheet(string source, string target, List<Dictionary<string, string>> replaceDatas)
        {
            using (FileStream fs = new FileStream(source, FileMode.Open, FileAccess.Read))
            {
                XSSFWorkbook xssfworkbook = new XSSFWorkbook(fs);

                if (xssfworkbook.NumberOfSheets != replaceDatas.Count)
                {
                    throw new ExcelTextReplaceException("excel文件sheet和要替换的数组长度不一致");
                }

                for (int j = 0; j < xssfworkbook.NumberOfSheets; j++)
                {
                    //XSSFSheet 
                    ISheet sheet = xssfworkbook.GetSheetAt(j);
                    //ROW
                    var rows = sheet.GetRowEnumerator();
                    //replaceData
                    Dictionary<string, string> data = replaceDatas[j];

                    while (rows.MoveNext())
                    {
                        var row = (XSSFRow)rows.Current;

                        for (var i = 0; i < row.LastCellNum; i++)
                        {
                            XSSFCell cell = (XSSFCell)row.GetCell(i);

                            if (cell != null)
                            {
                                String cellValue = cell.ToString();
                                foreach (var item in data)
                                {
                                    if (item.Key.Equals(cellValue))
                                    {
                                        cell.SetCellValue(item.Value);
                                    }
                                }
                            }
                        }
                    }
                }

                //另存为
                MemoryStream stream = new MemoryStream();
                xssfworkbook.Write(stream);
                var buf = stream.ToArray();

                //保存为Excel文件  
                using (FileStream fs1 = new FileStream(target, FileMode.Create, FileAccess.Write))
                {
                    fs1.Write(buf, 0, buf.Length);
                    fs1.Flush();

                    return target;
                }
            }
        }
    }
}
