using System;
using System.Collections.Generic;
using System.Text;

namespace ExcelTextReplace
{
    public class PlaceHolderVO
    {
        //占位符
        private string holder;

        public string Holder
        {
            get { return holder; }
            set { holder = value; }
        }

        private string value;

        public string Value
        {
            get { return this.value; }
            set { this.value = value; }
        }
    }
}
