using CellServe.ExcelHandler.Interfaces;
using CellServe.ExcelHandler.Strategies;
using LazyCache;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Dynamic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CellServe.ExcelHandler
{
    public class WorkbookRepository : IWorkbookRepository
    {
        private readonly IAppCache _cache;
        private readonly string _fileLocation;
        private readonly object _workbookLock;
        private readonly ISheetFilterStrategy _sheetFilterStrategy;
        private readonly IRowModelingStrategy _rowModelingStrategy;

        public WorkbookRepository(ISheetFilterStrategy sheetFilterStrategy, IRowModelingStrategy rowModelingStrategy, IAppCache cache)
        {
            _sheetFilterStrategy = sheetFilterStrategy;
            _rowModelingStrategy = rowModelingStrategy;
            _cache = cache;

            // If we can't find the user's wb location in web.config, pop it here:
            const string tempDirectory = @"C:\Temp";

            // In .Net, a lock object has to be a reference type
            _workbookLock = new { Name = "WorkbookLock" };

            try
            {
                // Grab or fudge the workbook file location
                _fileLocation = WorkbookConfig.GetWorkbookFilePath();
                if (string.IsNullOrEmpty(_fileLocation))
                {
                    Directory.CreateDirectory(tempDirectory);
                    _fileLocation = tempDirectory + @"\CellServeData.xlsx";
                }
            }
            catch (Exception ex)
            {
                throw new CellServeException("Failed to load the location of the workbook, or to create the temporary directory " + tempDirectory, ex);
            }
        }

        public List<Dictionary<string, string>> Read(string table, Dictionary<string, string> filterDictionary)
        {
            lock (_workbookLock)
            {
                using (var excel = GetExcelPackage())
                {
                    ExcelWorksheet sheet;
                    Dictionary<string, string> headers;
                    try
                    {
                        sheet = GetWorksheet(excel, table);
                        headers = GetHeaders(sheet.Cells);
                    }
                    catch (Exception)
                    {
                        throw new CellServeException("No such table as " + table);
                    }

                    var rowModels = _sheetFilterStrategy.FilterSheet(sheet, headers, filterDictionary);
                    return rowModels;
                }
            }
        }

        public void Add(string table, Dictionary<string, string> valuesDictionary)
        {
            lock (_workbookLock)
            {
                using (var excel = GetExcelPackage())
                {
                    ExcelWorksheet sheet;
                    Dictionary<string, string> headers;
                    try
                    {
                        sheet = GetWorksheet(excel, table);
                        headers = GetHeaders(sheet.Cells);
                    }
                    catch (Exception)
                    {
                        throw new CellServeException("No such table as " + table);
                    }

                    var newRow = GetNextFreeRow(sheet);
                    WriteRow(newRow, headers, valuesDictionary);
                    excel.Save();
                }
            }
        }

        public List<string> ColumnValues(string table, string column)
        {
            var values = _cache.GetOrAdd($"colvals-{table}-{column}", () => GetColumnValues(table, column));
            return values;
        }

        private List<string> GetColumnValues(string table, string fieldName)
        {
            lock (_workbookLock)
            {
                using (var excel = GetExcelPackage())
                {
                    ExcelWorksheet sheet;
                    Dictionary<string, string> headers;
                    try
                    {
                        sheet = GetWorksheet(excel, table);
                        headers = GetHeaders(sheet.Cells);
                    }
                    catch (Exception)
                    {
                        throw new CellServeException("No such table as " + table);
                    }

                    var values = _sheetFilterStrategy.GetAllValuesInColumn(sheet, headers, fieldName);
                    return values;
                }
            }
        }

        private ExcelPackage GetExcelPackage()
        {
            try
            {
                // Open or create a workbook at that location
                return new ExcelPackage(new FileInfo(_fileLocation));
            }
            catch (Exception ex)
            {
                throw new CellServeException("Failed to open the workbook at " + _fileLocation, ex);
            }
        }

        private void WriteRow(ExcelRange row, Dictionary<string, string> headers, Dictionary<string, string> valuesDictionary)
        {
            var rowNum = row.Start.Row;

            foreach (var valueToWrite in valuesDictionary)
            {
                var column = _rowModelingStrategy.GetColumnAlphaForField(headers, valueToWrite.Key);
                var cell = row[$"{column}{rowNum}"];
                cell.Value = valueToWrite.Value;
            }
        }

        private ExcelWorksheet GetWorksheet(ExcelPackage excel, string name)
        {
            name = name.ToLower();
            return excel.Workbook.Worksheets.FirstOrDefault(w => w.Name.ToLower() == name);
        }

        /// <summary>
        /// Get a dictionary of header address to field names, i.e. A => Name, B => DoB
        /// </summary>
        /// <param name="cells">The worksheet cells to scan</param>
        /// <returns>A dictionary of header address to field name</returns>
        private Dictionary<string, string> GetHeaders(ExcelRange cells)
        {
            var row = cells["1:1"];
            var headers = new Dictionary<string, string>();

            foreach (var cell in row)
            {
                // Stop scanning header row as soon as we get to an empty cell
                if (cell.Value == null) break;

                headers.Add(
                    cell.Address.AddressToColumnAlpha(),
                    cell.Value.ToString().Trim()
                );
            }

            return headers;
        }

        private ExcelRange GetNextFreeRow(ExcelWorksheet sheet)
        {
            try
            {
                // Get the used range (Dimension)
                int firstUnusedRow = 1;
                bool foundEmpty = false;
                var used = sheet.Dimension.Rows;
                for (int row = 1; row < used; row++)
                {
                    var firstColCell = sheet.Cells[$"A{row}"];
                    var cellValue = firstColCell.ToCellValueString();
                    if (cellValue == null)
                    {
                        firstUnusedRow = row;
                        foundEmpty = true;
                    }
                }

                if (!foundEmpty)
                {
                    firstUnusedRow = used + 1;
                }

                // Return the row range
                return sheet.Cells[$"{firstUnusedRow}:{firstUnusedRow}"];
            }
            catch (Exception ex)
            {
                throw new CellServeException("Couldn't find an empty row", ex);
            }
        }
    }
}
