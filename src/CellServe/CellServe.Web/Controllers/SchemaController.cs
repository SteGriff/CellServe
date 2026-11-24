using CellServe.ExcelHandler.Interfaces;
using Microsoft.AspNetCore.Mvc;

namespace CellServe.Web.Controllers;

[Route("[controller]")]
public sealed class SchemaController : Controller
{
    private readonly IWorkbookRepository _workbookRepository;

    public SchemaController(IWorkbookRepository workbookRepository) => _workbookRepository = workbookRepository;

    [HttpGet("")]
    public IActionResult Index() => Ok(new { Operation = "Schema", Results = _workbookRepository.Schema() });
}
