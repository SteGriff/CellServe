﻿using CellServe.ExcelHandler.Interfaces;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace CellServe.Web.Controllers
{
    public class HomeController : Controller
    {
        private readonly IWorkbookRepository _workbookRepository;

        public HomeController(IWorkbookRepository workbookRepository)
        {
            _workbookRepository = workbookRepository;
        }

        public ActionResult Index()
        {
            var schema = _workbookRepository.Schema();
            return View(schema);
        }

        public ActionResult Tester()
        {
            return View();
        }
    }
}