using BusinessLogic;
using Entity.Common;
using Entity.Results;
using System;
using System.Collections.Generic;
using System.Data.OleDb;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using Entity;

namespace AsignacionDeCuentas.Controllers
{
    public class AccountController : Controller
    {
        public ActionResult Index()
        {
            ViewBag.acccountCount = new CountResult();
            return View(new AssignmentResult());
        }

        [HttpPost]
        public ActionResult Index(string sheet)
        {
            AssigmentBusiness assBusness = new AssigmentBusiness();
            AccountBusiness accountBusiness = new AccountBusiness();
            AssignmentResult asscResult = null;
            StatementResult statement = null;
            CountResult countResult = null;
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

                assBusness.DeleteAssignment(userCode);
                statement = assBusness.GenerateAssigncSentences(location, provider, sheet);
                asscResult = assBusness.InsertAssignment(statement, userCode);

            }
            catch (ArgumentException except)
            {
                asscResult = new AssignmentResult { IsError = true, Message = except.Message};
                return View(asscResult);
            }
            catch (OleDbException except)
            {
                asscResult = new AssignmentResult { IsError = true, Message = except.Message};
                return View(asscResult);
            }

            countResult = accountBusiness.ValidInvalidAccount(userCode);//obtiene las cuentas de las cuentas validas e invalidas.
            ViewBag.acccountCount = countResult;

            if (countResult.Invalid > 0)
            {
                List<Account> acc = accountBusiness.GetInvalid(userCode);
                accountBusiness.WriteToExcel(acc, "Invalid", "Invalid_" + userCode, Server.MapPath("~/Content/Files/Excel/Invalid/"));
            }
            else 
            {
                accountBusiness.WriteToExcel(new List<Account>(), "Invalid", "Invalid_" + userCode, Server.MapPath("~/Content/Files/Excel/Invalid/"));
            }

            return View(asscResult);        
        }

        [HttpGet]
        public FileResult GetInvalidFromExcel()
        {
            return File("~/Content/Files/Excel/Invalid/Invalid_" + Session["UserCode"].ToString() + ".xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
        }
        
        [HttpGet]
        public FileResult GetAssignedExcel()
        {
            return File("~/Content/Files/Excel/Assigned/Assigned_" + Session["UserCode"].ToString() + ".xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
        }

        [HttpGet]
        public FileResult Documentation()
        {
            return File("~/Content/Files/Docs/Assignation de cuentas.pdf", "application/pdf");
        }

        [HttpPost]
        public ActionResult Assign()
        {
            string userCode = Session["UserCode"].ToString();
            string location = Server.MapPath("~/Content/Files/Excel/Assigned/");

            AccountBusiness accountBusiness = new AccountBusiness();            
            AssignmentResult asscResult = null;
            CountResult countResult = null;
            AssignedCountResult acc = null;

            try
            {
                acc = accountBusiness.Assign(userCode);
                accountBusiness.WriteAssignedToExcel(
                        accountBusiness.GetSpectedAssigment(userCode), "assigned", "Assigned_" + userCode, location
                    );

                countResult = accountBusiness.ValidInvalidAccount(userCode);//Get the cont of valid and invalid records.
            }
            catch (Exception except)
            {
                asscResult = new AssignmentResult { IsError = true, Message = except.Message};
            }
            finally
            {
                if (countResult == null)
                {
                    countResult = new CountResult();
                }
            }

            ViewBag.acccountCount = countResult;

            if (acc != null)
            {
                if (acc.Unassined > 0)
                {
                    asscResult = new AssignmentResult {
                        IsError = true,
                        IsAssined = false,
                        Message = string.Format("Error. Se han asignado solo {0}, de {1}. Intente nuevamente.", acc.Assigned, countResult.Valid)
                    };
                }
                else
                {
                    asscResult = new AssignmentResult
                    {
                        IsError = false,
                        IsAssined = true,
                        Message = string.Format("Un total de {0} asignadas de {1}", acc.Assigned, countResult.Valid)
                    };
                }
            }

            return View("Index", asscResult);
        }
    }
}