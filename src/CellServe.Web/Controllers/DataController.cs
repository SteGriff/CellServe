using CellServe.ExcelHandler;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace CellServe.Web.Controllers
{
    public class DataController : Controller
    {
        [HttpGet, ActionName("Index")]
        public ActionResult GetData(string table)
        {
            Response.TrySkipIisCustomErrors = true;
            var formData = Request.QueryString.AllKeys.ToDictionary(k => k, v => Request.QueryString[v]);

            List<Dictionary<string,string>> results;
            try
            {
                var repo = new WorkbookRepository();
                results = repo.Read(table, formData);
            }
            catch (CellServeException cex)
            {
                Response.StatusCode = 400;
                return Json(new { Message = cex.Message });
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
                var repo = new WorkbookRepository();
                repo.Add(table, formData);
            }
            catch (CellServeException cex)
            {
                Response.StatusCode = 400;
                return Json(new { Message = cex.Message });
            }

            Response.StatusCode = 201;
            return Json(req);
        }
    }
}