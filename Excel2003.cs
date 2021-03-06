﻿using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Text;
using System.IO;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using Newtonsoft.Json;


namespace ExcelTextReplace
{
    public class Excel2003
    {
        public static DataTable ExcelToTableForXLS(string file)
        {
            DataTable dt = new DataTable();
            using (FileStream fs = new FileStream(file, FileMode.Open, FileAccess.Read))
            {
                HSSFWorkbook hssfworkbook = new HSSFWorkbook(fs);
                ISheet sheet = hssfworkbook.GetSheetAt(0);

                //表头  
                IRow header = sheet.GetRow(sheet.FirstRowNum);
                List<int> columns = new List<int>();
                for (int i = 0; i < header.LastCellNum; i++)
                {
                    object obj = GetValueTypeForXLS(header.GetCell(i) as HSSFCell);
                    if (obj == null || obj.ToString() == string.Empty)
                    {
                        dt.Columns.Add(new DataColumn("Columns" + i.ToString()));
                        //continue;  
                    }
                    else
                        dt.Columns.Add(new DataColumn(obj.ToString()));
                    columns.Add(i);
                }
                //数据  
                for (int i = sheet.FirstRowNum + 1; i <= sheet.LastRowNum; i++)
                {
                    DataRow dr = dt.NewRow();
                    bool hasValue = false;
                    foreach (int j in columns)
                    {
                        dr[j] = GetValueTypeForXLS(sheet.GetRow(i).GetCell(j) as HSSFCell);
                        if (dr[j] != null && dr[j].ToString() != string.Empty)
                        {
                            hasValue = true;
                        }
                    }
                    if (hasValue)
                    {
                        dt.Rows.Add(dr);
                    }
                }
            }
            return dt;
        }

        /// <summary>  
        /// 将DataTable数据导出到Excel文件中(xls)  
        /// </summary>  
        /// <param name="dt"></param>  
        /// <param name="file"></param>  
        public static void TableToExcelForXLS(DataTable dt, string file)
        {
            HSSFWorkbook hssfworkbook = new HSSFWorkbook();
            ISheet sheet = hssfworkbook.CreateSheet("Test");

            //表头  
            IRow row = sheet.CreateRow(0);
            for (int i = 0; i < dt.Columns.Count; i++)
            {
                ICell cell = row.CreateCell(i);
                cell.SetCellValue(dt.Columns[i].ColumnName);
            }

            //数据  
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                IRow row1 = sheet.CreateRow(i + 1);
                for (int j = 0; j < dt.Columns.Count; j++)
                {
                    ICell cell = row1.CreateCell(j);
                    cell.SetCellValue(dt.Rows[i][j].ToString());
                }
            }

            //转为字节数组  
            MemoryStream stream = new MemoryStream();
            hssfworkbook.Write(stream);
            var buf = stream.ToArray();

            //保存为Excel文件  
            using (FileStream fs = new FileStream(file, FileMode.Create, FileAccess.Write))
            {
                fs.Write(buf, 0, buf.Length);
                fs.Flush();
            }
        }

        /// <summary>  
        /// 获取单元格类型(xls)  
        /// </summary>  
        /// <param name="cell"></param>  
        /// <returns></returns>  
        private static object GetValueTypeForXLS(HSSFCell cell)
        {
            if (cell == null)
                return null;
            switch (cell.CellType)
            {
                case CellType.Blank: //BLANK:  
                    return null;
                case CellType.Boolean: //BOOLEAN:  
                    return cell.BooleanCellValue;
                case CellType.Numeric: //NUMERIC:  
                    return cell.NumericCellValue;
                case CellType.String: //STRING:  
                    return cell.StringCellValue;
                case CellType.Error: //ERROR:  
                    return cell.ErrorCellValue;
                case CellType.Formula: //FORMULA:  
                default:
                    return "=" + cell.CellFormula;
            }
        }

        public static void replaceText(string source,string target,Dictionary<string,string> replaceData,int sheetIndex = 0)
        {
            using (FileStream fs = new FileStream(source, FileMode.Open, FileAccess.Read))
            {
                HSSFWorkbook hssfworkbook = new HSSFWorkbook(fs);

                Console.WriteLine("NumberOfSheets:" + hssfworkbook.NumberOfSheets);

                if (sheetIndex > hssfworkbook.NumberOfSheets)
                {
                    throw new ExcelTextReplaceException("sheetIndex 下表超出,表格总共有[" + hssfworkbook.NumberOfSheets + "]sheet");
                }
                
                //HSSFSheet 
                ISheet sheet = hssfworkbook.GetSheetAt(sheetIndex);

                var rows = sheet.GetRowEnumerator();

                while (rows.MoveNext())
                {
                    var row = (HSSFRow)rows.Current;
                    for (var i = 0; i < row.LastCellNum; i++)
                    {
                        HSSFCell cell = (HSSFCell)row.GetCell(i);
                        String cellValue = cell.ToString();

                        foreach (var item in replaceData)
                        {
                            if (item.Key.Equals(cellValue))
                            {
                                cell.SetCellValue(item.Value);
                            }
                        }
                    }
                }

                //另存为
                MemoryStream stream = new MemoryStream();
                hssfworkbook.Write(stream);
                var buf = stream.ToArray();

                //保存为Excel文件  
                using (FileStream fs1 = new FileStream("c:\\456.xls", FileMode.Create, FileAccess.Write))
                {
                    fs1.Write(buf, 0, buf.Length);
                    fs1.Flush();
                }
            }
        }

        public static string replaceTextAllSheet(string source,string target, List<Dictionary<string, string>> replaceDatas)
        {
            using (FileStream fs = new FileStream(source, FileMode.Open, FileAccess.Read))
            {
                HSSFWorkbook hssfworkbook = new HSSFWorkbook(fs);

                if (hssfworkbook.NumberOfSheets != replaceDatas.Count)
                {
                    throw new ExcelTextReplaceException("excel文件sheet和要替换的数组长度不一致");
                }

                for (int j = 0; j < hssfworkbook.NumberOfSheets;j++ )
                {
                    //HSSFSheet 
                    ISheet sheet = hssfworkbook.GetSheetAt(j);
                    //ROW
                    var rows = sheet.GetRowEnumerator();
                    //replaceData
                    Dictionary<string, string> data = replaceDatas[j];

                    while (rows.MoveNext())
                    {
                        var row = (HSSFRow)rows.Current;

                        for (var i = 0; i < row.LastCellNum; i++)
                        {
                            HSSFCell cell = (HSSFCell)row.GetCell(i);

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
                hssfworkbook.Write(stream);
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


        public static string replaceTextAllSheet(string source, string target, List<ReplaceDataVO> rpDatas)
        {
            using (FileStream fs = new FileStream(source, FileMode.Open, FileAccess.Read))
            {
                HSSFWorkbook hssfworkbook = new HSSFWorkbook(fs);
                foreach (ReplaceDataVO rpData in rpDatas)
                {
                    if (rpData.SheetIndex > hssfworkbook.NumberOfSheets)
                    {
                        throw new ExcelTextReplaceException("Sheet索引超出,最大索引值：[" + hssfworkbook.NumberOfSheets + "],当前传入的索引值：["+rpData.SheetIndex+"]");
                    }

                    //HSSFSheet 
                    HSSFSheet sheet = (HSSFSheet)hssfworkbook.GetSheetAt(rpData.SheetIndex - 1);

                    //新增行
                    foreach (RowDataVO rowData in rpData.RowData)
                    {
                        InsertRows(ref sheet, rowData.From, rowData.Rows);
                    }

                    var rows = sheet.GetRowEnumerator();
                    while (rows.MoveNext())
                    {
                        var row = (HSSFRow)rows.Current;
                        for (var i = 0; i < row.LastCellNum; i++)
                        {
                            HSSFCell cell = (HSSFCell)row.GetCell(i);
                            if (cell != null)
                            {
                                String cellValue = cell.ToString();
                                //游离文本替换
                                foreach (PlaceHolderVO phVo in rpData.PlaceHolderData)
                                {
                                    if (phVo.Holder.Equals(cellValue))
                                    {
                                        cell.SetCellValue(phVo.Value);
                                    }
                                }
                            }
                        }
                    }

                    //插入指定的单元格
                    foreach (LocationDataVO ltDataVo in rpData.LocationData)
                    {
                        if (ltDataVo.Localtion == null)
                        {
                            Console.WriteLine("11");
                        }
                        int[] arr = ExcelTools.getCellPoint(ltDataVo.Localtion);
                        IRow row = sheet.GetRow(arr[0] -1);
                        ICell cell;
                        if (row == null)
                        {
                            row = sheet.CreateRow(arr[0] - 1);
                            cell = row.CreateCell(arr[1] - 1);
                        }
                        else
                        {
                            cell = row.GetCell(arr[1] - 1);
                            if (cell == null)
                            {
                                cell = row.CreateCell(arr[1] - 1);
                            }
                        }

                        cell.SetCellValue(ltDataVo.Value);
                    }
                }

                //另存为
                MemoryStream stream = new MemoryStream();
                hssfworkbook.Write(stream);
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

        private static void InsertRows(ref HSSFSheet sheet1, int fromRowIndex, int rowCount)
        {
            sheet1.ShiftRows(fromRowIndex, sheet1.LastRowNum, rowCount, true, false, true);

            for (int rowIndex = fromRowIndex; rowIndex < fromRowIndex + rowCount; rowIndex++)
            {
                HSSFRow rowSource = (HSSFRow)sheet1.GetRow((rowIndex + rowCount));
                HSSFRow rowInsert = null;
                if (rowSource == null)
                {
                    rowInsert = (HSSFRow)sheet1.CreateRow(rowIndex);
                }
                else
                {
                    rowInsert = (HSSFRow)sheet1.CreateRow(rowIndex);
                    rowInsert.Height = rowSource.Height;
                    for (int colIndex = 0; colIndex < rowSource.LastCellNum; colIndex++)
                    {
                        HSSFCell cellSource = (HSSFCell)rowSource.GetCell(colIndex);
                        HSSFCell cellInsert = (HSSFCell)rowInsert.CreateCell(colIndex);
                        if (cellSource != null)
                        {
                            cellInsert.CellStyle = cellSource.CellStyle;
                        }
                    }
                }
            }
        }
    }
}
