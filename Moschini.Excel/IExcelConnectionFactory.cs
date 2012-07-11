using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Moschini.Excel
{
    public interface IExcelConnectionFactory
    {
        IExcelConnection CreateConnection(string filepath);
    }
}
