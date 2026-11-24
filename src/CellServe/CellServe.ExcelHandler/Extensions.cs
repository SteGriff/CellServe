using OfficeOpenXml;

namespace CellServe.ExcelHandler;

public static class Extensions
{
    public static string AddressToColumnAlpha(this string address) =>
        new(address.Where(char.IsLetter).ToArray());

    public static string? ToCellValueString(this object? value)
    {
        if (value is null)
        {
            return null;
        }

        return value is DateTime date ? date.ToString("dd/MM/yyyy") : value.ToString()?.Trim();
    }

    public static Dictionary<string, string> GetHeaders(this ExcelRange cells)
    {
        var headers = new Dictionary<string, string>();
        foreach (var cell in cells["1:1"])
        {
            if (cell.Value is null)
            {
                break;
            }

            headers[cell.Address.AddressToColumnAlpha()] = cell.Value.ToString()!.Trim();
        }

        return headers;
    }
}
