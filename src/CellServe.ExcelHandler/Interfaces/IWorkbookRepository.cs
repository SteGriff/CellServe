using System.Collections.Generic;

namespace CellServe.ExcelHandler.Interfaces
{
    public interface IWorkbookRepository
    {
        List<Dictionary<string, string>> Read(string table, Dictionary<string, string> filterDictionary);
        void Add(string table, Dictionary<string, string> valuesDictionary);
    }
}
