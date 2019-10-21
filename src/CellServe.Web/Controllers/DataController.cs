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
            var formData = Request.QueryString.AllKeys.ToDictionary(k => k, v => Request.QueryString[v]);

            object req = new
            {
                Table = table,
                Operation = "Read",
                Filter = formData
            };

            return Json(req, JsonRequestBehavior.AllowGet);
        }

        [HttpPost, ActionName("Index")]
        public ActionResult PostData(string table, FormCollection form)
        {
            //var formData = form.AllKeys.Select(k => )

            var formData = form.AllKeys.ToDictionary(k => k, v => form[v]);

            object req = new
            {
                Table = table,
                Operation = "Write",
                Values = formData
            };

            return Json(req);
        }
    }
}