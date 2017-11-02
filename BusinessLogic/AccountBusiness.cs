using Entity;
using Entity.Results;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DataAccess;

namespace BusinessLogic
{
    public class AccountBusiness
    {
        public List<Account> GetInvalid(string userCode)
        {
            return new AccountData().GetInvalid(userCode);
        }
        
        public AssignedCountResult Assign(string userCode)
        {
            return new AccountData().Assign(userCode);
        }

        public CountResult ValidInvalidAccount(string userCode)
        {            
           return new AccountData().ValidInvalidAccount(userCode);
        }

        public List<AssignedResult> GetSpectedAssigment(string userCode)
        {
            return new AccountData().GetSpectedAssigment(userCode);
        }

        public string WriteToExcel(List<Account> assignments, string sheetName, string fileName, string savingPath)
        {
            return new AccountData().WriteToExcel(assignments, sheetName, fileName, savingPath);
        }

        public string WriteAssignedToExcel(List<AssignedResult> assignments, string sheetName, string fileName, string savingPath)
        {
            return new AccountData().WriteAssignedToExcel(assignments, sheetName, fileName, savingPath);
        }
    }
}
