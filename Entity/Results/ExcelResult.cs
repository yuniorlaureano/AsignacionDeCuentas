using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Entity.Results
{
    public class ExcelResult
    {
        public int Id { get; set; }
        public string Message { get; set; }
        public string[] Sheets { get; set; }
        public string StackTrace { get; set; }
        public bool IsError { get; set; }
        public bool IsSucess { get; set; }
    }
}
