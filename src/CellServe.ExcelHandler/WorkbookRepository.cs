﻿using LazyCache;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CellServe.ExcelHandler
{
    public class WorkbookRepository
    {
        private IAppCache _queueCache;
        private string _fileLocation;
        private object _workbookLock;

        public WorkbookRepository()
        {
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

        public void Add(string table, Dictionary<string,string> valuesDictionary)
        {
            lock (_workbookLock)
            {
                var excel = GetExcelPackage();
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

        private void WriteRow(ExcelRange row, Dictionary<string, string> headers, Dictionary<string, string> valuesDictionary)
        {
            var rowNum = row.Start.Row;

            foreach (var valueToWrite in valuesDictionary)
            {
                var valueToWriteLowercaseName = valueToWrite.Key.ToLower();
                var header = headers.FirstOrDefault(h => h.Value.ToLower() == valueToWriteLowercaseName);
                var column = header.Key;
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
                    AddressToColumnAlpha(cell.Address),
                    cell.Value.ToString().Trim()
                );
            }

            return headers;
        }

        private string AddressToColumnAlpha(string address)
        {
            return new string(address.Where(char.IsLetter).ToArray());
        }

        private ExcelRange GetNextFreeRow(ExcelWorksheet sheet)
        {
            try
            {
                // Get the used range (Dimension)
                int firstUnusedRow = 1;
                bool foundEmpty = false;
                var used = sheet.Dimension.Rows;
                for(int row = 1; row < used; row++)
                {
                    var firstColCell = sheet.Cells[$"A{row}"];
                    var cellValue = GetCellValue(firstColCell);
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

                //var firstColumn = cells["A:A"];
                //var firstColumnFirstEmptyCell = firstColumn.FirstOrDefault(c => string.IsNullOrEmpty(c.Value.ToString()));
                //var rowNumber = firstColumnFirstEmptyCell.Start.Row;

                // Return the row range
                return sheet.Cells[$"{firstUnusedRow}:{firstUnusedRow}"];
            }
            catch (Exception ex)
            {
                throw new CellServeException("Couldn't find an empty row", ex);
            }
        }

        private string GetCellValue(object value)
        {
            if (value == null)
            {
                return null;
            }
            else if (value as DateTime? != null)
            {
                return ((DateTime)value).ToString("dd/MM/yyyy");
            }
            return value.ToString().Trim();
        }
    }
}
