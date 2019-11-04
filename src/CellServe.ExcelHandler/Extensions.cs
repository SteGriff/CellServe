using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CellServe.ExcelHandler
{
    public static class Extensions
    {
        public static string AddressToColumnAlpha(this string address)
        {
            return new string(address.Where(char.IsLetter).ToArray());
        }

        public static string ToCellValueString(this object value)
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


        /// <summary>
        /// Get a dictionary of header address to field names, i.e. A => Name, B => DoB
        /// </summary>
        /// <param name="cells">The worksheet cells to scan</param>
        /// <returns>A dictionary of header address to field name</returns>
        public static Dictionary<string, string> GetHeaders(this ExcelRange cells)
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
    }
}
