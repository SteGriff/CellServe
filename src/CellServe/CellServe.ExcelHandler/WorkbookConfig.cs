namespace CellServe.ExcelHandler;

public static class WorkbookConfig
{
    public static string GetWorkbookFilePath(WorkbookOptions options) =>
        string.IsNullOrWhiteSpace(options.FilePath)
            ? Path.Combine(Path.GetTempPath(), "CellServeData.xlsx")
            : options.FilePath;
}
