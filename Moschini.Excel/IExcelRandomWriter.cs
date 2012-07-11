using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Drawing;

namespace Moschini.Excel
{
    public interface IExcelRandomWriter : IDisposable
    {
        void Write(uint rowIndex, string columnName, string value);
        void DeleteRow(uint rowIndex);
        void DeleteRow(int rowIndex);
        void ColorizeRow(uint rowIndex, Color color);
        void MoveRow(uint oldIndex, uint newIndex);
        void MoveRow(int oldIndex, int newIndex);
    }
}
