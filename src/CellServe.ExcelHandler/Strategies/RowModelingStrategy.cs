using CellServe.ExcelHandler.Interfaces;
using OfficeOpenXml;
using System.Collections.Generic;
using System.Linq;

namespace CellServe.ExcelHandler.Strategies
{
    public class RowModelingStrategy : IRowModelingStrategy
    {
        public Dictionary<string, string> RowToModel(ExcelRange row, Dictionary<string, string> headers)
        {
            var rowModel = new Dictionary<string, string>
            {
                { "$ExcelRow", row.Start.Row.ToString() }
            };

            foreach (var cell in row)
            {
                var cellColumnAlpha = cell.Address.AddressToColumnAlpha();
                if (headers.ContainsKey(cellColumnAlpha))
                {
                    var fieldName = headers[cellColumnAlpha];
                    var cellValue = cell.Value.ToCellValueString();
                    rowModel.Add(fieldName, cellValue);
                }
            }

            return rowModel;
        }

        public string GetColumnAlphaForField(Dictionary<string, string> headers, string fieldName)
        {
            var valueToWriteLowercaseName = fieldName.ToLower();
            var header = headers.FirstOrDefault(h => h.Value.ToLower() == valueToWriteLowercaseName);
            var column = header.Key;
            return column;
        }
    }
}
