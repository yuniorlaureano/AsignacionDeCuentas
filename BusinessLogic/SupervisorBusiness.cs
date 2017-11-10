using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Entity;
using DataAccess;

namespace BusinessLogic
{
    public class SupervisorBusiness
    {
        public List<Supervisor> GetSupervisors()
        {
            return new SupervisorData().GetSupervisors();
        }

        public void WriteToExcel(List<Supervisor> supervisors, string sheetName, string fileName, string savingPath)
        {
            SupervisorData supervisorData = new SupervisorData();
            supervisorData.WriteToExcel(supervisors, sheetName, fileName, savingPath);
        }
    }
}
