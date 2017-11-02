using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Entity.Results
{
    public class AssignedResult
    {
        public string Id { get; set; }
        public string Status { get; set; }
        public string ChargeOut { get; set; }
        public string ChargeIn { get; set; }
        public string CanvCode { get; set; }
        public string CanvEdition { get; set; }
        public string SubscrId { get; set; }
        public string UserAssigned { get; set; }
        public string UserToAssign { get; set; }
        public string IsAssigned { get; set; }
    }
}
