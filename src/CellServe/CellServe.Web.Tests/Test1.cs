using CellServe.ExcelHandler.Interfaces;
using CellServe.ExcelHandler.Models;
using CellServe.Web.Controllers;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;

namespace CellServe.Web.Tests;

[TestClass]
public sealed class Test1
{
    [TestMethod]
    public void SuggestionsRejectsMultipleFields()
    {
        var controller = new DataController(new FakeWorkbookRepository());
        controller.ControllerContext = new ControllerContext
        {
            HttpContext = new DefaultHttpContext()
        };
        controller.HttpContext.Request.QueryString = new QueryString("?Name=Test&Country=UK");

        var result = controller.Suggestions("Customers");

        Assert.IsInstanceOfType<BadRequestObjectResult>(result);
    }

    [TestMethod]
    public void SuggestionsReturnEmptyForShortSearchTerms()
    {
        var controller = new DataController(new FakeWorkbookRepository());
        controller.ControllerContext = new ControllerContext
        {
            HttpContext = new DefaultHttpContext()
        };
        controller.HttpContext.Request.QueryString = new QueryString("?Name=Te");

        var result = controller.Suggestions("Customers") as OkObjectResult;

        Assert.IsNotNull(result);
        Assert.Contains("Results", result.Value!.ToString()!);
    }

    private sealed class FakeWorkbookRepository : IWorkbookRepository
    {
        public WorkbookSchema Schema() => throw new NotImplementedException();

        public List<string> ColumnValues(string table, string column) => new() { "Test User" };

        public List<Dictionary<string, string>> Read(string table, IReadOnlyDictionary<string, string> filterDictionary) =>
            new();

        public void Add(string table, IReadOnlyDictionary<string, string> valuesDictionary)
        {
        }
    }
}
