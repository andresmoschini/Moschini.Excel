using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Moschini.Excel
{
    public interface IExcelReader : IExcelRow, IDisposable
    {
        string WorksheetName { get; }
        int ActualPos { get; }
        bool Read();
    }
}
