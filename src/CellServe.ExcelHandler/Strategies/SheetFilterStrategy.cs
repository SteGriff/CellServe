using CellServe.ExcelHandler.Interfaces;
using LazyCache;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Dynamic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CellServe.ExcelHandler.Strategies
{
    public class SheetFilterStrategy : ISheetFilterStrategy
    {
        private readonly IRowModelingStrategy _rowModelingStrategy;

        public SheetFilterStrategy(IRowModelingStrategy rowModelingStrategy)
        {
            _rowModelingStrategy = rowModelingStrategy;
        }

        public List<Dictionary<string, string>> FilterSheet(ExcelWorksheet sheet, Dictionary<string, string> headers, Dictionary<string, string> filterDictionary)
        {
            var used = sheet.Dimension.Rows;
            var matchingRows = new List<int>();

            // Start at 2; don't scan header
            for (int row = 2; row <= used; row++)
            {
                // Check this row against every filter passed in
                bool currentRowMatchesFilters = true;

                foreach (var filterColumn in filterDictionary)
                {
                    // Skip empty filter criteria
                    if (string.IsNullOrEmpty(filterColumn.Value)) { continue; }

                    var columnAlpha = _rowModelingStrategy.GetColumnAlphaForField(headers, filterColumn.Key);
                    var cell = sheet.Cells[$"{columnAlpha}{row}"];
                    var cellValue = cell.Value.ToCellValueString();
                    if (cellValue.ToLower() != filterColumn.Value.ToLower())
                    {
                        // If any filter fails, quit
                        currentRowMatchesFilters = false;
                        break;
                    }
                }

                if (currentRowMatchesFilters)
                {
                    matchingRows.Add(row);
                }
            }

            // Compose the matching rows into objects and return a list of them
            var rowModels = new List<Dictionary<string, string>>();
            foreach (var matchingRowNumber in matchingRows)
            {
                var row = sheet.Cells[$"{matchingRowNumber}:{matchingRowNumber}"];
                var rowModel = _rowModelingStrategy.RowToModel(row, headers);
                rowModels.Add(rowModel);
            }
            return rowModels;
        }

        public List<string> GetAllValuesInColumn(ExcelWorksheet sheet, Dictionary<string, string> headers, string fieldName)
        {
            try
            {
                var used = sheet.Dimension.Rows;

                var columnAlpha = _rowModelingStrategy.GetColumnAlphaForField(headers, fieldName);
                var entireColumn = sheet.Cells[$"{columnAlpha}:{columnAlpha}"];

                var results = entireColumn
                    .Skip(1)
                    .Take(used)
                    .Where(c => c.Value != null)
                    .Select(c => c.Value.ToCellValueString())
                    .ToList();

                return results;
            }
            catch (Exception)
            {
                return new List<string>();
            }
        }
    }
}
