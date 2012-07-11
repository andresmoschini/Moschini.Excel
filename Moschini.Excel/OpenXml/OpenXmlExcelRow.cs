using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Text.RegularExpressions;

namespace Moschini.Excel.OpenXml
{
    internal class OpenXmlExcelRow : IExcelRow, IDisposable
    {
        Dictionary<string, object> data;

        public int FieldCount { get; private set; }

        internal OpenXmlExcelRow(OpenXmlExcelConnection connection, Row row)
        {
            data = new Dictionary<string, object>();
            foreach (var cell in row.Elements<Cell>())
            {
                object value; //TODO: test it and improve
                if (cell.DataType == null)
                {
                    //Capture default DataTypes here
                    value = string.IsNullOrEmpty(cell.InnerText) ? null : cell.InnerText;
                }
                else
                {
                    switch (cell.DataType.Value)
                    {
                        case CellValues.Boolean: value = bool.Parse(cell.InnerText); break;
                        case CellValues.Date: value = DateTime.Parse(cell.InnerText); break;
                        case CellValues.Number: value = decimal.Parse(cell.InnerText); break;
                        case CellValues.SharedString: value = connection.GetSharedString(int.Parse(cell.InnerText)); break;
                        default: value = cell.InnerText; break;
                    }
                }
                data.Add(ExcelUtilities.GetColumnName(cell.CellReference.Value), value);
            }
            FieldCount = data.Any() ? ExcelUtilities.ColumnNameToOrdinal(data.Last().Key) : 0;
        }

        public object this[int i]
        {
            get { return GetValue(i); }
        }

        public object this[string columnName]
        {
            get { return GetValue(columnName); }
        }

        private DateTime? ObjectToDateTime(object value)
        {
            if (value == null)
                return null;
            else if (value is DateTime || value is DateTime?)
                    return (DateTime?)value;
            else
                return (DateTime?)DateTime.FromOADate(double.Parse(value.ToString()));
        }

        private T Convert<T>(object value, string columnName, Action<string, Type, Object> onError)
        {
            if (default(T) != null && value == null)
            {
                if (onError != null)
                    onError(columnName, typeof(T), value);
            }
            else if (value != null)
            {
                try
                {
                    if (typeof(T) == typeof(int) || typeof(T) == typeof(int?))
                        return (T)(object)int.Parse(value.ToString());
                    if (typeof(T) == typeof(DateTime?) || typeof(T) == typeof(DateTime))
                        return (T)(object)ObjectToDateTime(value);
                    if (typeof(T) == typeof(string))
                        return (T)(object)value.ToString();
                    else
                        return (T)value;
                }
                catch
                {
                    if (onError != null)
                        onError(columnName, typeof(T), value);
                }
            }
            return default(T);
        }

        public T GetValue<T>(int i)
        {
            return Convert<T>(GetValue(i), null, null);
        }

        public T GetValue<T>(string columnName, Action<string, Type, Object> action)
        {
            return Convert<T>(GetValue(columnName), columnName, action);
        }

        public T GetValue<T>(string columnName)
        {
            return GetValue<T>(columnName, null);
        }
        

        public object GetValue(int i)
        {
            if (i >= FieldCount)
                throw new IndexOutOfRangeException();
            else
                return GetValue(ExcelUtilities.OrdinalToColumnName(i + 1));
        }

        public object GetValue(string columnName)
        {
            //return data[columnName];
            object result;
            if (data.TryGetValue(columnName, out result))
            {
                string value = result as string;
                return value != null ? value.Trim() : result;
            }
            else
                return null;

        }

        public void Dispose()
        {
            if (data != null)
                data = null;
        }

        #region IExcelRow Members


        public bool IsEmpty
        {
            get
            {
                foreach (var value in data.Values)
                    if (value != null && value.ToString() != string.Empty)
                        return false;
                    return true;
            }
        }

        #endregion
    }
}
