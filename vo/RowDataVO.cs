using System;
using System.Collections.Generic;
using System.Text;

namespace ExcelTextReplace
{
    public class RowDataVO
    {
        private int from;
        private int rows;

        public int Rows
        {
            get { return rows; }
            set { rows = value; }
        }

        public int From
        {
            get { return from; }
            set { from = value; }
        }
        
    }
}
