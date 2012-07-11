using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;

namespace Moschini.Excel.OpenXml
{
    //TODO: remove not referenced common strings
    internal class OpenXmlExcelRandomWriter : IExcelRandomWriter
    {
        public string WorksheetName { get; private set; }
        private Worksheet worksheet;
        private SheetData sheetData;
        readonly OpenXmlExcelConnection connection;
        private Dictionary<System.Drawing.Color, UInt32Value> cellFormats = new Dictionary<System.Drawing.Color, UInt32Value>();
        private Dictionary<System.Drawing.Color, UInt32Value> fillPatterns = new Dictionary<System.Drawing.Color, UInt32Value>();

        internal OpenXmlExcelRandomWriter(OpenXmlExcelConnection connection, string worksheetName, bool createSheetIfNotExists)
        {
            WorksheetName = worksheetName;
            this.connection = connection;
            this.worksheet = connection.GetSheet(worksheetName, createSheetIfNotExists);
            this.sheetData = worksheet.GetFirstChild<SheetData>();
        }

        public void Write(uint rowIndex, string columnName, string value)
        {
            string cellReference = columnName + rowIndex;
            var row = GetRow(sheetData, rowIndex);
            var cell = GetCell(row, columnName);

            int oldkey;
            if (cell.DataType == null
                || cell.DataType.Value != CellValues.SharedString
                || !int.TryParse(cell.CellValue.Text, out oldkey))
                oldkey = -1;
            var sharedStringIndex = connection.GetSharedStringIndex(value);
            cell.CellValue = new CellValue(sharedStringIndex.ToString());
            cell.DataType = new DocumentFormat.OpenXml.EnumValue<CellValues>(CellValues.SharedString);
            if (oldkey >= 0)
                connection.DeferredSubmitSharedStringTableDeletes();
            else
                connection.DeferredSubmit();
        }

        public void DeleteRow(uint rowIndex)
        {
            var row = GetRow(sheetData, rowIndex);
            row.Remove();
            connection.DeferredSubmitSharedStringTableDeletes();
        }

        public void DeleteRow(int rowIndex)
        {
            if (rowIndex <= 0)
                throw new ArgumentOutOfRangeException("rowIndex has to be greater than zero");
            DeleteRow((uint)rowIndex);
        }

        private UInt32Value GetFillPatternId(System.Drawing.Color color)
        {
            //TODO: enhance it
            if (!fillPatterns.ContainsKey(color))
            {
                var newindex = connection.Stylesheet.Fills.Count();
                connection.Stylesheet.Fills.Count = (UInt32Value)(uint)newindex + 1;
                fillPatterns.Add(color, (UInt32Value)(uint)newindex);
                var fill = new Fill(new PatternFill()
                {
                    PatternType = PatternValues.Solid,
                    BackgroundColor = new BackgroundColor() { Indexed = (UInt32Value)64U },
                    ForegroundColor = new ForegroundColor() { Rgb = new HexBinaryValue(System.Drawing.ColorTranslator.ToHtml(color).Replace("#", "")) }
                });
                connection.Stylesheet.Fills.Append(fill);
                connection.DeferredSubmitStyle();
            }
            return fillPatterns[color];
        }

        private UInt32Value GetCellFormatId(System.Drawing.Color color)
        {
            //TODO: enhance it, maybe with differential formatting record (dxf)
            if (!cellFormats.ContainsKey(color))
            {
                var newindex = connection.Stylesheet.CellFormats.Count();
                connection.Stylesheet.CellFormats.Count = (UInt32Value)(uint)newindex + 1;
                cellFormats.Add(color, (UInt32Value)(uint)newindex);
                var format = new CellFormat() { NumberFormatId=0, FontId = 0, FillId = GetFillPatternId(color), FormatId=0, ApplyFill = true };
                connection.Stylesheet.CellFormats.Append(format);
                connection.DeferredSubmitStyle();
            }
            return cellFormats[color];
        }

        public void ColorizeRow(uint rowIndex, System.Drawing.Color color)
        {
            var row = GetRow(sheetData, rowIndex);
            row.CustomFormat = true;
            row.StyleIndex = GetCellFormatId(color);
            foreach (var cell in row.Elements<Cell>())
                cell.StyleIndex = GetCellFormatId(color);
            connection.DeferredSubmitStyle();
        }

        public void MoveRow(int oldIndex, int newIndex)
        {
            if (oldIndex <= 0)
                throw new ArgumentOutOfRangeException("oldIndex parameter has to be greater than zero");
            if (newIndex <= 0)
                throw new ArgumentOutOfRangeException("oldIndex parameter has to be greater than zero");
            MoveRow((uint)oldIndex, (uint)newIndex);
        }

        public void MoveRow(uint oldIndex, uint newIndex)
        {
            if (oldIndex == newIndex)
                return;
            var rowInNewPlace = GetRow(sheetData, newIndex, false);
            if (rowInNewPlace != null)
                throw new ArgumentException("newIndex refers to an existing row");
            var row = GetRow(sheetData, oldIndex, false);
            if (row == null)
                throw new ArgumentException("oldIndex refers to a not existing row");
            row.RowIndex = newIndex;
            foreach (var cell in row.Elements<Cell>())
            {
                var column = ExcelUtilities.GetColumnName(cell.CellReference);
                cell.CellReference = string.Format("{0}{1}", column, newIndex);
            }
            connection.DeferredSubmit();
        }

        private static Cell GetCell(Row row, string columnName)
        {
            string cellReference = columnName + row.RowIndex;
            int columnIndex = ExcelUtilities.ColumnNameToOrdinal(columnName);

            var cell = row.Elements<Cell>().Where(c => c.CellReference.Value == cellReference).FirstOrDefault();
            if (cell == null)
            {
                Cell refCell = null;
                foreach (Cell otherCell in row.Elements<Cell>())
                {
                    int otherColumnIndex = ExcelUtilities.ColumnNameToOrdinal(ExcelUtilities.GetColumnName(otherCell.CellReference.Value));
                    if (otherColumnIndex > columnIndex)
                    {
                        refCell = otherCell;
                        break;
                    }
                }
                cell = new Cell() { CellReference = cellReference };
                row.InsertBefore(cell, refCell);
                if (refCell == null)
                {
                    if (row.Spans == null)
                        row.Spans = new ListValue<StringValue>();
                    row.Spans.InnerText = string.Format("1:{0}", columnIndex);
                }
                //connection.DeferredSubmit();
            }
            return cell;
        }

        private static Row GetRow(SheetData sheetData, uint excelRowIndex)
        {
            return GetRow(sheetData, excelRowIndex, true);
        }

        private static Row GetRow(SheetData sheetData, uint excelRowIndex, bool autoAdd)
        {
            var row = sheetData.Elements<Row>().Where(r => r.RowIndex.Value == excelRowIndex).FirstOrDefault();
            if (row == null && autoAdd)
            {
                row = new Row() { RowIndex = excelRowIndex };
                sheetData.Append(row);
            }
            return row;
        }

        public virtual void Dispose()
        {
            if (sheetData != null)
                sheetData = null;
        }
    }
}
