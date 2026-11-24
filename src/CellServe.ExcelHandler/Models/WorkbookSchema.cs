using OfficeOpenXml;
using System.Collections.Generic;

namespace CellServe.ExcelHandler.Models
{
    public class WorkbookSchema
    {
        public string FilePath { get; set; } = string.Empty;
        public List<SpreadsheetSchema> Sheets { get; set; } = new();

        public int NumberOfSheets
        {
            get
            {
                return Sheets.Count;
            }
        }

        public WorkbookSchema(ExcelPackage excel)
        {
            FilePath = excel.File.FullName;
            Sheets = new List<SpreadsheetSchema>();

            foreach (var sheet in excel.Workbook.Worksheets)
            {
                var ws = new SpreadsheetSchema(sheet.Name)
                {
                    Rows = Math.Max(0, (sheet.Dimension?.Rows ?? 1) - 1),
                    Headers = sheet.Cells.GetHeaders()
                };
                Sheets.Add(ws);
            }
        }
    }
}
