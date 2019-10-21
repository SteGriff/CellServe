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
            object req = new
            {
                Table = table,
                Operation = "Read"
            };
            return Json(req, JsonRequestBehavior.AllowGet);
        }

        [HttpPost, ActionName("Index")]
        public ActionResult PostData(string table, FormCollection form)
        {
            //var formData = form.AllKeys.Select(k => )

            object req = new
            {
                Table = table,
                Operation = "Write",
                Values = form
            };
            return Json(req);
        }
    }
}