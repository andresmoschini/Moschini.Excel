using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Moschini.Excel
{
    public interface IExcelConnection : IDisposable
    {
        IExcelReader CreateReader(string worksheetName);
        IExcelReader CreateReader(int worksheetIndex);
        int? CountRows(string worksheetName);
        int? CountRows(int worksheetIndex);
        IExcelRandomWriter CreateRandomWriter(string worksheetName);
        IExcelRandomWriter CreateRandomWriter(string worksheetName, bool createSheetIfNotExists);
        IExcelRandomWriter CreateRandomWriter(int worksheetIndex);
    }
}
