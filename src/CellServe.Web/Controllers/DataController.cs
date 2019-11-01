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

            if (formData.Count < 1)
            {
                Response.StatusCode = 400;
                return Json(new { Message = "Please pass in a Key:Value representing a Column:SearchTerm" });
            }
            else if (formData.Count > 1)
            {
                Response.StatusCode = 400;
                return Json(new { Message = "Please pass in one field only to get Suggestions for that field. You sent multiple fields." });
            }

            var searchColumn = formData.FirstOrDefault().Key;
            var searchTerm = formData.FirstOrDefault().Value.ToLower();

            List<string> results;
            if (searchTerm.Length > 2)
            {
                try
                {
                    var allColumnValues = _workbookRepository.ColumnValues(table, searchColumn);
                    results = allColumnValues
                        .Where(colValue => colValue.ToLower().Contains(searchTerm))
                        .Take(20)
                        .ToList();
                }
                catch (CellServeException cex)
                {
                    Response.StatusCode = 400;
                    return Json(new { cex.Message });
                }
            }
            else
            {
                results = new List<string>();
            }
            
            object response = new
            {
                Table = table,
                Operation = "Suggestions",
                Filter = formData,
                Results = results
            };

            return Json(response, JsonRequestBehavior.AllowGet);
        }

    }
}