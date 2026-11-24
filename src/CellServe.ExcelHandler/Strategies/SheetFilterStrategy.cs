using CellServe.ExcelHandler.Interfaces;
using OfficeOpenXml;

namespace CellServe.ExcelHandler.Strategies;

public sealed class SheetFilterStrategy : ISheetFilterStrategy
{
    private readonly IRowModelingStrategy _rowModelingStrategy;

    public SheetFilterStrategy(IRowModelingStrategy rowModelingStrategy) => _rowModelingStrategy = rowModelingStrategy;

    public List<Dictionary<string, string>> FilterSheet(
        ExcelWorksheet sheet,
        Dictionary<string, string> headers,
        IReadOnlyDictionary<string, string> filterDictionary)
    {
        var usedRows = sheet.Dimension?.Rows ?? 1;
        var matchingRows = new List<int>();
        for (var row = 2; row <= usedRows; row++)
        {
            var matches = true;
            foreach (var filter in filterDictionary)
            {
                if (string.IsNullOrWhiteSpace(filter.Value))
                {
                    continue;
                }

                var column = _rowModelingStrategy.GetColumnAlphaForField(headers, filter.Key);
                var cellValue = sheet.Cells[$"{column}{row}"].Value.ToCellValueString();
                if (!string.Equals(cellValue, filter.Value, StringComparison.OrdinalIgnoreCase))
                {
                    matches = false;
                    break;
                }
            }

            if (matches)
            {
                matchingRows.Add(row);
            }
        }

        return matchingRows
            .Select(row => _rowModelingStrategy.RowToModel(sheet.Cells[$"{row}:{row}"], headers))
            .ToList();
    }

    public List<string> GetAllValuesInColumn(ExcelWorksheet sheet, Dictionary<string, string> headers, string fieldName)
    {
        var column = _rowModelingStrategy.GetColumnAlphaForField(headers, fieldName);
        var usedRows = sheet.Dimension?.Rows ?? 1;
        return Enumerable.Range(2, Math.Max(0, usedRows - 1))
            .Select(row => sheet.Cells[$"{column}{row}"].Value.ToCellValueString())
            .Where(value => value is not null)
            .Cast<string>()
            .ToList();
    }
}
