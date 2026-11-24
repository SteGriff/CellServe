using CellServe.ExcelHandler;
using CellServe.ExcelHandler.Interfaces;
using Microsoft.AspNetCore.Mvc;

namespace CellServe.Web.Controllers;

[Route("[controller]/{table}")]
public sealed class DataController : Controller
{
    private readonly IWorkbookRepository _workbookRepository;

    public DataController(IWorkbookRepository workbookRepository) => _workbookRepository = workbookRepository;

    [HttpGet]
    public IActionResult Read(string table)
    {
        try
        {
            var filters = Request.Query.ToDictionary(item => item.Key, item => item.Value.ToString());
            return Ok(new { Table = table, Operation = "Read", Filter = filters, Results = _workbookRepository.Read(table, filters) });
        }
        catch (CellServeException exception)
        {
            return BadRequest(new { message = exception.Message });
        }
    }

    [HttpPost]
    public IActionResult Add(string table, [FromForm] IFormCollection form)
    {
        try
        {
            var values = form.ToDictionary(item => item.Key, item => item.Value.ToString());
            _workbookRepository.Add(table, values);
            return StatusCode(StatusCodes.Status201Created, new { Table = table, Operation = "Write", Values = values });
        }
        catch (CellServeException exception)
        {
            return BadRequest(new { message = exception.Message });
        }
    }

    [HttpGet("Suggestions")]
    public IActionResult Suggestions(string table)
    {
        var filters = Request.Query.ToDictionary(item => item.Key, item => item.Value.ToString());
        if (filters.Count != 1)
        {
            return BadRequest(new
            {
                message = filters.Count == 0
                    ? "Please pass a single Column:SearchTerm query value."
                    : "Please pass exactly one field to get suggestions."
            });
        }

        var filter = filters.Single();
        if (filter.Value.Length < 3)
        {
            return Ok(new { Table = table, Operation = "Suggestions", Filter = filters, Results = Array.Empty<string>() });
        }

        try
        {
            var results = _workbookRepository.ColumnValues(table, filter.Key)
                .Where(value => value.Contains(filter.Value, StringComparison.OrdinalIgnoreCase))
                .Take(20)
                .ToList();
            return Ok(new { Table = table, Operation = "Suggestions", Filter = filters, Results = results });
        }
        catch (CellServeException exception)
        {
            return BadRequest(new { message = exception.Message });
        }
    }
}
