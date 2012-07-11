using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Moschini.Excel
{
    public class ExcelDumpper : IDisposable
    {
        //TODO: refactor it to a more OO approach with less dictionaries and tuples
        private Dictionary<string, IExcelConnection> connections = new Dictionary<string, IExcelConnection>();
        private Dictionary<string, Dictionary<string, IExcelRandomWriter>> writers = new Dictionary<string, Dictionary<string, IExcelRandomWriter>>();
        private IExcelConnectionFactory excelConnectionFactory;
        
        public ExcelDumpper(IExcelConnectionFactory excelConnectionFactory)
        {
            this.excelConnectionFactory = excelConnectionFactory;
        }

        public void Write(string filepath, string sheetname, uint rowIndex, string columnName, string value)
        {
            var writter = GetWritter(filepath, sheetname);
            writers[filepath][sheetname].Write(rowIndex, columnName, value);
        }

        protected IExcelRandomWriter GetWritter(string filepath, string sheetname)
        {
            var connection = GetConnection(filepath);

            if (!writers[filepath].ContainsKey(sheetname))
            {
                var newWritter = connection.CreateRandomWriter(sheetname);
                writers[filepath][sheetname] = newWritter;
            }

            return writers[filepath][sheetname];
        }

        protected IExcelConnection GetConnection(string filepath)
        {
            if (!connections.ContainsKey(filepath))
            {
                var newConnection = excelConnectionFactory.CreateConnection(filepath);
                connections[filepath] = newConnection;
                writers[filepath] = new Dictionary<string, IExcelRandomWriter>();
            }

            return connections[filepath];
        }

        public void Dispose()
        {
            foreach (var writterDictionary in writers.Values)
            {
                foreach (var writter in writterDictionary.Values)
                {
                    writter.Dispose();
                }
            }

            foreach (var connection in connections.Values)
            {
                connection.Dispose();
            }
        }
    }
}
