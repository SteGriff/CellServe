using CellServe.ExcelHandler;
using CellServe.ExcelHandler.Interfaces;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace CellServe.Web.Controllers
{
    public class SchemaController : Controller
    {
        private readonly IWorkbookRepository _workbookRepository;

        public SchemaController(IWorkbookRepository workbookRepository)
        {
            _workbookRepository = workbookRepository;
        }

        // GET: Schema
        public ActionResult Index()
        {
            
            return View();
        }
    }
}