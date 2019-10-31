using OfficeOpenXml;
using System.Collections.Generic;

namespace CellServe.ExcelHandler.Interfaces
{
    public interface IRowModelingStrategy
    {
        Dictionary<string, string> RowToModel(ExcelRange row, Dictionary<string, string> headers);

        string GetColumnAlphaForField(Dictionary<string, string> headers, string fieldName);
    }
}
