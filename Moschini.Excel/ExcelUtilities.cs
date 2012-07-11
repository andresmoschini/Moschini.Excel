using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace Moschini.Excel
{
    public static class ExcelUtilities
    {
        public static int ColumnNameToOrdinal(string columnName)
        {
            if (columnName == string.Empty)
                return 0;
            var lastLetter = columnName[columnName.Length - 1];
            var restLetters = columnName.Substring(0, columnName.Length - 1);
            int value = (int)lastLetter - 64;
            return value + 26 * ColumnNameToOrdinal(restLetters);
        }

        public static string OrdinalToColumnName(int ordinal)
        {
            if (ordinal == 0)
                return string.Empty;
            ordinal--;

            var module = (ordinal % 26) + 1;
            char lastLetter = (char)(module + 64);
            int restLetters = ordinal / 26;
            return OrdinalToColumnName(restLetters) + lastLetter;
        }

        public static string GetColumnName(string cellReference)
        {
            Regex regex = new Regex("[A-Za-z]+");
            Match match = regex.Match(cellReference);
            return match.Value;
        }

        public static int GetRowNo(string cellReference)
        {
            Regex regex = new Regex("[0-9]+");
            Match match = regex.Match(cellReference);
            return int.Parse(match.Value);
        }

        public static void ReadAndDo<T>(this IExcelReader reader, string column, Action<T> action)
        {
            if (column != null)
            {
                var value = reader.GetValue<T>(column);
                action(value);
            }
        }
    }
}
