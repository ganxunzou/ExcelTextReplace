using System;
using System.Collections.Generic;
using System.Text;

namespace ExcelTextReplace
{
    /// <summary>
    /// Excel 替换数据VO
    /// </summary>
    public class ReplaceDataVO
    {
        private List<PlaceHolderVO> placeholderData;

        public List<PlaceHolderVO> PlaceHolderData
        {
            get { return placeholderData; }
            set { placeholderData = value; }
        }

        private List<RowDataVO> rowData;

        public List<RowDataVO> RowData
        {
            get { return rowData; }
            set { rowData = value; }
        }

        private List<LocationDataVO> locationData;

        public List<LocationDataVO> LocationData
        {
            get { return locationData; }
            set { locationData = value; }
        }

        private int sheetIndex;

        public int SheetIndex
        {
            get { return sheetIndex; }
            set { sheetIndex = value; }
        }
    }
}
