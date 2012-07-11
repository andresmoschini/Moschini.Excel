using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Moschini.Excel
{
    public interface IExcelRow
    {
        object this[int i] { get; }
        object this[string columnName] { get; }
        //string GetString(int i);
        object GetValue(int i);
        object GetValue(string columnName);
        T GetValue<T>(int i);
        T GetValue<T>(string columnName);
        T GetValue<T>(string columnName, Action<string, Type, Object> onError);
        int FieldCount { get; }
        bool IsEmpty { get; }
    }
}
