using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using Entity;
using Entity.Common;
using System.Data.OleDb;
using System.IO;

namespace AsignacionDeCuentas.Controllers
{
    public class AssignmentController : Controller
    {
        public ActionResult Index()
        {
            return View();
        }

        [HttpPost]
        public ActionResult UploadExcel(HttpPostedFile excel)
        {
            ExcelOperation excelOp = new ExcelOperation();

            if (excel.ContentLength == 0)
            {
                ViewBag.Error = "Debe proveer el archivo.";
                return View("Index");
            }

            try
            {
                excelOp.SetProvider(excel.FileName);
            }
            catch (ArgumentException except)
            {
                ViewBag.Error = except.Message;
                return View("Index");
            }
            catch (OleDbException except)
            {
                ViewBag.Error = except.Message;
                return View("Index");
            }

            string location = Server.MapPath("~/Content/Files/Excel/Assigment_"+73156+""+excelOp.FileExtension);

            if (System.IO.File.Exists(location))
            {
                System.IO.File.Delete(location);
            }

            excel.SaveAs(location);

            string[] sheets = excelOp.GetSheetNames(excelOp.Provider, excelOp.FileName);

            return View("Index", sheets);
        }
	}
}