using System;
using System.Configuration;

namespace CellServe.ExcelHandler
{
    public static class WorkbookConfig
    {
        public static string GetWorkbookFilePath()
        {
            try
            {
                return ConfigurationManager.AppSettings["WorkbookFilePath"];
            }
            catch (Exception)
            {
                return null;
            }
        }
    }
}
