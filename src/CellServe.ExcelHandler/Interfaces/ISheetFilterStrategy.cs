using OfficeOpenXml;
using System.Collections.Generic;

namespace CellServe.ExcelHandler.Interfaces
{
    public interface ISheetFilterStrategy
    {
        List<Dictionary<string, string>> FilterSheet(ExcelWorksheet sheet, Dictionary<string, string> headers, Dictionary<string, string> filterDictionary);
    }
}
