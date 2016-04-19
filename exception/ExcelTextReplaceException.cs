using System;
using System.Collections.Generic;
using System.Text;
using System.Runtime.Serialization;

namespace ExcelTextReplace
{
    [Serializable]
    class ExcelTextReplaceException : ApplicationException
    {
       public ExcelTextReplaceException() { }  
        public ExcelTextReplaceException(string message)  
            : base(message) { }
        public ExcelTextReplaceException(string message, Exception inner)  
            : base(message, inner) { }  
    }
}
