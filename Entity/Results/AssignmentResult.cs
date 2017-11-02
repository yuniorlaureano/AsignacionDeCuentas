using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Entity.Results
{
    public class AssignmentResult
    {
        public int Id { get; set; }
        public bool IsAssined { get; set; }
        public string Message { get; set; }
        public string StackTrace { get; set; }
        public bool IsError { get; set; }
        public int InsertedRows { get; set; }
        public int SpectedRows { get; set; }
    }
}
