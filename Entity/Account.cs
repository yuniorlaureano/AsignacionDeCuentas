using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Entity
{
    public class Account
    {
        public int Id { get; set; }
        public string Status { get; set; }
        public int ChargeOut { get; set; }
        public int ChargeIn { get; set; }
        public string CanvCode { get; set; }
        public int CanvEdition { get; set; }
        public int SubscrId { get; set; }
        public string UserAssigned { get; set; }
        public string UserToAssign { get; set; }
    }
}
