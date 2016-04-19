using System;
using System.Collections.Generic;
using System.Text;

namespace ExcelTextReplace
{
    /// <summary>
    /// Location Data VO
    /// </summary>
    public class LocationDataVO
    {
        private string localtion;

        public string Localtion
        {
            get { return localtion; }
            set { localtion = value; }
        }
        private string value;

        public string Value
        {
            get { return this.value; }
            set { this.value = value; }
        }
    }
}
