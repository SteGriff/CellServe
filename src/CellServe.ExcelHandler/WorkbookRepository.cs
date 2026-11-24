using CellServe.ExcelHandler.Interfaces;
using CellServe.ExcelHandler.Models;
using Microsoft.Extensions.Caching.Memory;
using Microsoft.Extensions.Options;
using OfficeOpenXml;

namespace CellServe.ExcelHandler;

public sealed class WorkbookRepository : IWorkbookRepository
{
    private readonly IMemoryCache _cache;
    private readonly string _fileLocation;
    private readonly object _workbookLock = new();
    private readonly ISheetFilterStrategy _sheetFilterStrategy;
    private readonly IRowModelingStrategy _rowModelingStrategy;

    public WorkbookRepository(
        ISheetFilterStrategy sheetFilterStrategy,
        IRowModelingStrategy rowModelingStrategy,
        IMemoryCache cache,
        IOptions<WorkbookOptions> options)
    {
        _sheetFilterStrategy = sheetFilterStrategy;
        _rowModelingStrategy = rowModelingStrategy;
        _cache = cache;
        _fileLocation = WorkbookConfig.GetWorkbookFilePath(options.Value);

        var directory = Path.GetDirectoryName(_fileLocation);
        if (!string.IsNullOrWhiteSpace(directory))
        {
            Directory.CreateDirectory(directory);
        }
    }

    public WorkbookSchema Schema() =>
        _cache.GetOrCreate("schema", entry =>
        {
            entry.SlidingExpiration = TimeSpan.FromMinutes(10);
            lock (_workbookLock)
            {
                using var excel = GetExcelPackage();
                return new WorkbookSchema(excel);
            }
        })!;

    public List<Dictionary<string, string>> Read(string table, IReadOnlyDictionary<string, string> filterDictionary)
    {
        lock (_workbookLock)
        {
            using var excel = GetExcelPackage();
            var sheet = GetWorksheet(excel, table);
            return _sheetFilterStrategy.FilterSheet(sheet, sheet.Cells.GetHeaders(), filterDictionary);
        }
    }

    public void Add(string table, IReadOnlyDictionary<string, string> valuesDictionary)
    {
        lock (_workbookLock)
        {
            using var excel = GetExcelPackage();
            var sheet = GetWorksheet(excel, table);
            var headers = sheet.Cells.GetHeaders();
            var row = GetNextFreeRow(sheet);
            foreach (var value in valuesDictionary)
            {
                var column = _rowModelingStrategy.GetColumnAlphaForField(headers, value.Key);
                row[$"{column}{row.Start.Row}"].Value = value.Value;
            }

            excel.Save();
            _cache.Remove("schema");
            _cache.Remove($"colvals-{table}");
        }
    }

    public List<string> ColumnValues(string table, string column) =>
        _cache.GetOrCreate($"colvals-{table}-{column}", entry =>
        {
            entry.SlidingExpiration = TimeSpan.FromMinutes(10);
            lock (_workbookLock)
            {
                using var excel = GetExcelPackage();
                var sheet = GetWorksheet(excel, table);
                return _sheetFilterStrategy.GetAllValuesInColumn(sheet, sheet.Cells.GetHeaders(), column);
            }
        })!;

    private ExcelPackage GetExcelPackage()
    {
        try
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            return new ExcelPackage(new FileInfo(_fileLocation));
        }
        catch (Exception ex)
        {
            throw new CellServeException($"Failed to open the workbook at {_fileLocation}", ex);
        }
    }

    private static ExcelWorksheet GetWorksheet(ExcelPackage excel, string name) =>
        excel.Workbook.Worksheets.FirstOrDefault(sheet =>
            string.Equals(sheet.Name, name, StringComparison.OrdinalIgnoreCase))
        ?? throw new CellServeException($"No such table as {name}");

    private static ExcelRange GetNextFreeRow(ExcelWorksheet sheet)
    {
        var usedRows = sheet.Dimension?.Rows ?? 1;
        for (var row = 2; row <= usedRows; row++)
        {
            if (sheet.Cells[$"A{row}"].Value.ToCellValueString() is null)
            {
                return sheet.Cells[$"{row}:{row}"];
            }
        }

        return sheet.Cells[$"{usedRows + 1}:{usedRows + 1}"];
    }
}
