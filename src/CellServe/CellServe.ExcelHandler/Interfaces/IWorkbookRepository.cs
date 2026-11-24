using CellServe.ExcelHandler.Models;

namespace CellServe.ExcelHandler.Interfaces;

public interface IWorkbookRepository
{
    WorkbookSchema Schema();
    List<string> ColumnValues(string table, string column);
    List<Dictionary<string, string>> Read(string table, IReadOnlyDictionary<string, string> filterDictionary);
    void Add(string table, IReadOnlyDictionary<string, string> valuesDictionary);
}
