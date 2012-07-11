using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Collections.ObjectModel;

namespace Moschini.Excel.OpenXml
{
    public class OpenXmlExcelConnectionFactory : IExcelConnectionFactory
    {
        public IExcelConnection CreateConnection(string filepath)
        {
            return new OpenXmlExcelConnection(filepath);
        }
    }
}
