using CellServe.ExcelHandler;
using CellServe.ExcelHandler.Interfaces;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace CellServe.Web.Controllers
{
    public class DataController : Controller
    {
        private readonly IWorkbookRepository _workbookRepository;

        public DataController(IWorkbookRepository workbookRepository)
        {
            _workbookRepository = workbookRepository;
        }

        [HttpGet, ActionName("Index")]
        public ActionResult GetData(string table)
        {
            Response.TrySkipIisCustomErrors = true;
            var formData = Request.QueryString.AllKeys.ToDictionary(k => k, v => Request.QueryString[v]);

            List<Dictionary<string,string>> results;
            try
            {
                results = _workbookRepository.Read(table, formData);
            }
            catch (CellServeException cex)
            {
                Response.StatusCode = 400;
                return Json(new { cex.Message });
            }

            object response = new
            {
                Table = table,
                Operation = "Read",
                Filter = formData,
                Results = results
            };

            return Json(response, JsonRequestBehavior.AllowGet);
        }

        [HttpPost, ActionName("Index")]
        public ActionResult PostData(string table, FormCollection form)
        {
            Response.TrySkipIisCustomErrors = true;
            var formData = form.AllKeys.ToDictionary(k => k, v => form[v]);

            object req = new
            {
                Table = table,
                Operation = "Write",
                Values = formData
            };

            try
            {
                _workbookRepository.Add(table, formData);
            }
            catch (CellServeException cex)
            {
                Response.StatusCode = 400;
                return Json(new { cex.Message });
            }

            Response.StatusCode = 201;
            return Json(req);
        }

        [HttpGet, ActionName("Suggestions")]
        public ActionResult GetSuggestions(string table)
        {
            Response.TrySkipIisCustomErrors = true;
            var formData = Request.QueryString.AllKeys.ToDictionary(k => k, v => Request.QueryString[v]);

            List<Dictionary<string, string>> results;
            try
            {
                results = _workbookRepository.Read(table, formData);
            }
            catch (CellServeException cex)
            {
                Response.StatusCode = 400;
                return Json(new { cex.Message });
            }

            object response = new
            {
                Table = table,
                Operation = "Read",
                Filter = formData,
                Results = results
            };

            return Json(response, JsonRequestBehavior.AllowGet);
        }

    }
}