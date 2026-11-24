using CellServe.ExcelHandler.Interfaces;
using OfficeOpenXml;

namespace CellServe.ExcelHandler.Strategies;

public sealed class RowModelingStrategy : IRowModelingStrategy
{
    public Dictionary<string, string> RowToModel(ExcelRange row, Dictionary<string, string> headers)
    {
        var rowModel = new Dictionary<string, string>
        {
            ["$ExcelRow"] = row.Start.Row.ToString()
        };

        foreach (var cell in row)
        {
            var column = cell.Address.AddressToColumnAlpha();
            if (headers.TryGetValue(column, out var fieldName))
            {
                rowModel[fieldName] = cell.Value.ToCellValueString() ?? string.Empty;
            }
        }

        return rowModel;
    }

    public string GetColumnAlphaForField(Dictionary<string, string> headers, string fieldName) =>
        headers.FirstOrDefault(header => string.Equals(header.Value, fieldName, StringComparison.OrdinalIgnoreCase)).Key
        ?? throw new CellServeException($"No such column as {fieldName}");
}
