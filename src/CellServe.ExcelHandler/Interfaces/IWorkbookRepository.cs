using CellServe.ExcelHandler.Models;
using System.Collections.Generic;

namespace CellServe.ExcelHandler.Interfaces
{
    public interface IWorkbookRepository
    {
        WorkbookSchema Schema();
        List<string> ColumnValues(string table, string column);
        List<Dictionary<string, string>> Read(string table, Dictionary<string, string> filterDictionary);
        void Add(string table, Dictionary<string, string> valuesDictionary);
    }
}
