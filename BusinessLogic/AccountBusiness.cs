using Entity;
using Entity.Results;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DataAccess;
using Entity.Common;

namespace BusinessLogic
{
    public class AccountBusiness
    {
        /// <summary>
        /// Obtiene las cuentas invalidas.
        /// </summary>
        /// <param name="userCode">usuario logueado en el sisteama</param>
        /// <returns>List<Account></returns>
        public List<Account> GetInvalid(string userCode)
        {
            return new AccountData().GetInvalid(userCode);
        }
        
        /// <summary>
        /// Corre el procedure que asigna las cuentas.
        /// </summary>
        /// <param name="userCode">usuario logueado en el sistema</param>
        /// <returns>AssignedCountResult</returns>
        public AssignedCountResult Assign(string userCode)
        {
            return new AccountData().Assign(userCode);
        }

        /// <summary>
        /// Obtiene la calidad de cuentas validas e invalidas.
        /// </summary>
        /// <param name="userCode">usuario logueado en el sistema</param>
        /// <returns>CountResult</returns>
        public CountResult ValidInvalidAccount(string userCode)
        {            
           return new AccountData().ValidInvalidAccount(userCode);
        }

        /// <summary>
        /// Obtiene las cuentas invalidas, desde un archivo de excel.
        /// </summary>
        /// <param name="fileNameWithLocation">Localizacion del archivo de excel</param>
        /// <param name="provider">Si es xls, o xlsx</param>
        /// <param name="sheet">la hoja de excel</param>
        /// <returns>List<Account></returns>
        public List<Account> GetInvalalidAccountFromExcel(string fileNameWithLocation, Provider provider, string sheet)
        {
            return new AccountData().GetInvalalidAccountFromExcel(fileNameWithLocation, provider, sheet);
        }

        /// <summary>
        /// Obtine una listga con las cuentas se supone debieron asignarse
        /// </summary>
        /// <param name="userCode">usuerio loguado en el sistema</param>
        /// <returns>List<AssignedResult></returns>
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
