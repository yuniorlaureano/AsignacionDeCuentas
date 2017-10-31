using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Entity
{
    public class User
    {//select usr_nivel from evaluaciones.dbo.evausrd where usr_codigo = '{0}' and pyc_codigo = 'sra'"
        public int UserCode { get; set; }
        public string UserName { get; set; }
        public string[] Roles { get; set; }
    }
}
