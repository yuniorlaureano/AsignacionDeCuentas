using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using Entity;
using Entity.Results;
using Entity.Common;
using System.Data.OleDb;
using System.IO;
using BusinessLogic;

namespace AsignacionDeCuentas.Controllers
{
    public class AssignmentController : Controller
    {
        public ActionResult Index()
        {
            ExcelResult excelResult = new ExcelResult();
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

                if (System.IO.File.Exists(location))
                {
                    excelResult.IsError = false;
                    excelResult.IsSucess = false;
                    excelResult.Sheets = new ExcelOperation().GetSheetNames(provider, location);
                }                
            }
            catch (ArgumentException except)
            {
                excelResult = new ExcelResult { IsError = true, Message = except.Message, StackTrace = except.StackTrace };
                return View(excelResult);
            }
            catch (OleDbException except)
            {
                excelResult = new ExcelResult { IsError = true, Message = except.Message, StackTrace = except.StackTrace };
                return View(excelResult);
            }

            return View(excelResult);
        }

        [HttpPost]
        public ActionResult UploadExcel(HttpPostedFile excel)
        {
            HttpPostedFileBase file = Request.Files[0];
            ExcelOperation excelOp = new ExcelOperation();
            ExcelResult excelResult = null;

            if (file.ContentLength == 0)
            {
                excelResult = new ExcelResult { IsError = true, Message = "Debe proveer el archivo." };
                return View("Index", excelResult);
            }

            try
            {
                excelOp.SetProvider(file.FileName);
            }
            catch (ArgumentException except)
            {
                excelResult = new ExcelResult { IsError = true, Message = except.Message, StackTrace = except.StackTrace };
                return View("Index", excelResult);
            }
            catch (OleDbException except)
            {
                excelResult = new ExcelResult { IsError = true, Message = except.Message, StackTrace = except.StackTrace };
                return View("Index", excelResult);
            }

            string location = Server.MapPath("~/Content/Files/Excel/Assigment_"+Session["UserCode"]+""+excelOp.FileExtension);
            
            try
            {
                if (System.IO.File.Exists(location))
                {
                    System.IO.File.Delete(location);
                }

                file.SaveAs(location);

                excelResult = new ExcelResult
                {
                    IsError = false,
                    IsSucess = true,
                    Message = "Elija la hoja de excel.",
                    Sheets = excelOp.GetSheetNames(excelOp.Provider, location)
                };
            }
            catch (OleDbException except)
            {
                excelResult = new ExcelResult { IsError = true, Message = except.Message, StackTrace = except.StackTrace };
                return View("Index", excelResult);
            }
            catch(IOException except)
            {
                excelResult = new ExcelResult { IsError = true, Message = except.Message, StackTrace = except.StackTrace };
                return View("Index", excelResult);
            }
            catch (ArgumentException except)
            {
                excelResult = new ExcelResult { IsError = true, Message = except.Message, StackTrace = except.StackTrace };
                return View("Index", excelResult);
            }
            finally
            {
                excelOp.CloseConnection();
            }
            
            return View("Index", excelResult);
        }        
	}
}