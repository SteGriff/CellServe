using CellServe.ExcelHandler.Interfaces;
using Microsoft.AspNetCore.Mvc;

namespace CellServe.Web.Controllers;

public sealed class HomeController : Controller
{
    private readonly IWorkbookRepository _workbookRepository;

    public HomeController(IWorkbookRepository workbookRepository) => _workbookRepository = workbookRepository;

    public IActionResult Index() => View(_workbookRepository.Schema());

    public IActionResult Tester() => View();

    [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
    public IActionResult Error() => Problem("An unexpected error occurred.");
}
