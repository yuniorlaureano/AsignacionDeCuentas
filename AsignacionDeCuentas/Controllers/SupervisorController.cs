using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using Entity;
using BusinessLogic;

namespace AsignacionDeCuentas.Controllers
{
    public class SupervisorController : Controller
    {
        /// <summary>
        /// Permite obtener los supervisores.
        /// </summary>
        /// <returns>ActionResult</returns>
        public ActionResult Index()
        {
            SupervisorBusiness supervisorBusiness = new SupervisorBusiness();
            List<Supervisor> supervisors = supervisorBusiness.GetSupervisors();
            supervisorBusiness.WriteToExcel(supervisors, "Supervisors", "Supervisors", Server.MapPath("~/Content/Files/Excel/Supervisor/"));
            return View(supervisors);
        }

        /// <summary>
        /// Obtiene el archivo excel con los supervisores
        /// </summary>
        /// <returns>FileResult</returns>
        public FileResult GetSupervisorFile()
        {
            return File("~/Content/Files/Excel/Supervisor/Supervisors.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
        }
	}
}