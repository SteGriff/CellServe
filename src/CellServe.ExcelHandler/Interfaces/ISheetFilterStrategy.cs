using OfficeOpenXml;

namespace CellServe.ExcelHandler.Interfaces;

public interface ISheetFilterStrategy
{
    List<Dictionary<string, string>> FilterSheet(ExcelWorksheet sheet, Dictionary<string, string> headers, IReadOnlyDictionary<string, string> filterDictionary);
    List<string> GetAllValuesInColumn(ExcelWorksheet sheet, Dictionary<string, string> headers, string fieldName);
}
