using BusinessLogic;
using Entity.Common;
using Entity.Results;
using System;
using System.Collections.Generic;
using System.Data.OleDb;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace AsignacionDeCuentas.Controllers
{
    public class AccountController : Controller
    {
        public ActionResult Index()
        {
            ViewBag.acccountCount = new AccountBusiness();
            return View(new AssignmentResult());
        }

        [HttpPost]
        public ActionResult Index(string sheet)
        {
            AssigmentBusiness assBusness = new AssigmentBusiness();
            AssignmentResult asscResult = null;
            StatementResult statement = null;
            string userCode = Session["UserCode"].ToString();

            string location = Server.MapPath("~/Content/Files/Excel/Assigment_" + userCode + ".xlsx");
            Provider provider = Provider.XLSX;

            try
            {
                if (!System.IO.File.Exists(location))
                {
                    location = Server.MapPath("~/Content/Files/Excel/Assigment_" + userCode + ".xls");
                    provider = Provider.XLS;
                }

                statement = assBusness.GenerateAssigncSentences(location, provider, sheet);
                asscResult = assBusness.InsertAssignment(statement, userCode);

            }
            catch (ArgumentException except)
            {
                asscResult = new AssignmentResult { IsError = true, Message = except.Message, StackTrace = except.StackTrace };
                return View(asscResult);
            }
            catch (OleDbException except)
            {
                asscResult = new AssignmentResult { IsError = true, Message = except.Message, StackTrace = except.StackTrace };
                return View(asscResult);
            }

            ViewBag.acccountCount = new AccountBusiness().ValidInvalidAccount(userCode);

            return View(asscResult);
        
        }

	}
}