using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Text.RegularExpressions;

namespace Moschini.Excel.OpenXml
{
    internal class OpenXmlExcelReader : IExcelReader
    {
        public string WorksheetName { get; private set; }
        private SheetData sheetdata;
        readonly OpenXmlExcelConnection connection;
        
        Row currentRow = null;
        OpenXmlExcelRow currentExcelRow = null;
        public int ActualPos { get; private set; }

        internal OpenXmlExcelReader(OpenXmlExcelConnection connection, string worksheetName)
        {
            WorksheetName = worksheetName;
            this.connection = connection;
            this.sheetdata = connection.GetSheet(worksheetName, false).GetFirstChild<SheetData>();
            ActualPos = -1;
            //FirstRowValues = null;
        }

        public bool Read()
        {
            ActualPos++;
            //if (ActualPos == 20) return false;
            //if (currentExcelRow != null)
            //    currentExcelRow.Dispose();

            if (currentRow == null)
                currentRow = sheetdata.Elements<Row>().First();
            else
                currentRow = currentRow.NextSibling<Row>();

            if (currentRow == null)
                return false;
            else
            {
                currentExcelRow = new OpenXmlExcelRow(connection, currentRow);
//                if (_firstRowValues == null)
//                    _firstRowValues = new OpenXmlExcelRow(connection, currentRow); ;
                return true;
            }
        }

        public virtual void Dispose()
        {
            if (sheetdata != null)
                sheetdata = null;
        }

        #region IExcelRow Members

        public object this[int i]
        {
            get { return currentExcelRow[i]; }
        }

        public object this[string columnName]
        {
            get { return currentExcelRow[columnName]; }
        }

        public T GetValue<T>(int i)
        {
            return currentExcelRow.GetValue<T>(i);
        }

        public T GetValue<T>(string columnName)
        {
            return currentExcelRow.GetValue<T>(columnName);
        }

        public T GetValue<T>(string columnName, Action<string, Type, object> onError)
        {
            return currentExcelRow.GetValue<T>(columnName, onError);
        }

        public object GetValue(int i)
        {
            return currentExcelRow.GetValue(i);
        }

        public object GetValue(string columnName)
        {
            return currentExcelRow.GetValue(columnName);
        }

        public int FieldCount
        {
            get { return currentExcelRow.FieldCount; }
        }

        public bool IsEmpty
        {
            get { return currentExcelRow.IsEmpty; }
        }

        #endregion
    }
}
