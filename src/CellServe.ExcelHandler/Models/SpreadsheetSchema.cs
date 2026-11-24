using System.Collections.Generic;

namespace CellServe.ExcelHandler.Models
{
    public class SpreadsheetSchema
    {
        public SpreadsheetSchema(string name)
        {
            Name = name;
        }

        public string Name { get; set; }
        public Dictionary<string, string> Headers { get; set; } = new();
        public int Rows { get; set; }
    }
}
