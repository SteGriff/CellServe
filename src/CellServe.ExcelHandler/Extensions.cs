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
    }
}
