using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Collections.ObjectModel;

namespace Moschini.Excel.OpenXml
{
    public class OpenXmlExcelConnection : IExcelConnection
    {
        private class SharedStringsStorage : KeyedCollection<string, string>
        {
            public SharedStringsStorage() 
                : base()
            { 
            }

            public SharedStringsStorage(IEnumerable<string> items)
                : this()
            {
                foreach (var item in items)
                {
                    this.Add(item);
                }
            }

            protected new void Add(string item)
            {
                if (this.Contains(item))
                    this.Add(item + "'");
                else
                    base.Add(item);
            }

            protected override string GetKeyForItem(string item)
            {
                return item;
            }

            public bool Add(string value, out int index)
            {
                index = IndexOf(value);
                if (index == -1)
                {
                    Add(value);
                    index = IndexOf(value);
                    return true;
                }
                return false;                
            }
        }

        SpreadsheetDocument spreadsheetDocument;
        
        //string[] sharedStrings;
        SharedStringsStorage sharedStrings;

        internal OpenXmlExcelConnection(string filepath) 
        {
            try
            {
                spreadsheetDocument = SpreadsheetDocument.Open(filepath, true);
                //TODO: optimize it
                sharedStrings = new SharedStringsStorage(spreadsheetDocument.WorkbookPart.GetPartsOfType<SharedStringTablePart>().First().SharedStringTable.Select(x => x.InnerText));
                //sharedStrings = spreadsheetDocument.WorkbookPart.GetPartsOfType<SharedStringTablePart>().First().SharedStringTable.Select(x => x.InnerText).ToArray();
            }
            catch
            {
                Dispose();
                throw;
            }
        }

        internal string GetSharedString(int index)
        {
            return sharedStrings.ElementAt(index);
        }

        internal int GetSharedStringIndex(string value)
        {
            int index;
            if (sharedStrings.Add(value, out index))
            {
                spreadsheetDocument.WorkbookPart.GetPartsOfType<SharedStringTablePart>().First().SharedStringTable.AppendChild(new SharedStringItem(new DocumentFormat.OpenXml.Spreadsheet.Text(value)));
                DeferredSubmitSharedStringTableAddings();
            }
            return index;
        }

        public void Dispose()
        {
            DeferredSubmit(true);

            if (spreadsheetDocument != null)
            {
                spreadsheetDocument.Dispose();
                spreadsheetDocument = null;
            }
            if (sharedStrings != null)
                sharedStrings = null;
        }

        public IExcelReader CreateReader(string worksheetName)
        {
            return new OpenXmlExcelReader(this, worksheetName);
        }

        public IExcelReader CreateReader(int worksheetIndex)
        {
            return new OpenXmlExcelReader(this, GetSheetName(worksheetIndex));
        }

        internal string GetSheetName(int index)
        {
            return spreadsheetDocument.WorkbookPart.Workbook.Sheets.Elements<Sheet>().Skip(index).FirstOrDefault().Name;
        }

        internal WorkbookPart WorkbookPart
        {
            get { return spreadsheetDocument.WorkbookPart; }
        }

        internal Stylesheet Stylesheet
        {
            get { return WorkbookPart.GetPartsOfType<WorkbookStylesPart>().First().Stylesheet; }
        }

        internal Worksheet GetSheet(string worksheetName, bool createSheetIfNotExists)
        {
            var sheet = WorkbookPart.Workbook.Sheets.Elements<Sheet>().Where(s => s.Name == worksheetName).FirstOrDefault();
            string relationshipId;
            WorksheetPart worksheetPart;
            if (sheet == null)
            {
                if (!createSheetIfNotExists)
                {
                    throw new ApplicationException("Worksheet doesn't exist");
                }
                else
                {
                    // Add a blank WorksheetPart.
                    worksheetPart = WorkbookPart.AddNewPart<WorksheetPart>();
                    worksheetPart.Worksheet = new Worksheet(new SheetData());
                    worksheetPart.Worksheet.Save();

                    relationshipId = WorkbookPart.GetIdOfPart(worksheetPart);

                    uint sheetId = 1;
                    if (WorkbookPart.Workbook.Sheets.Elements<Sheet>().Count() > 0)
                        sheetId = WorkbookPart.Workbook.Sheets.Elements<Sheet>().Select(s => s.SheetId.Value).Max() + 1;

                    sheet = new Sheet() { Id = relationshipId, SheetId = sheetId, Name = worksheetName };
                    WorkbookPart.Workbook.Sheets.Append(sheet);
                    WorkbookPart.Workbook.Save();
                }
            }
            else
            {
                relationshipId = sheet.Id.Value;
                worksheetPart = (WorksheetPart)spreadsheetDocument.WorkbookPart.GetPartById(relationshipId);
            }
            
            return worksheetPart.Worksheet;
        }

        public IExcelRandomWriter CreateRandomWriter(string worksheetName)
        {
            return CreateRandomWriter(worksheetName, false);
        }

        public IExcelRandomWriter CreateRandomWriter(string worksheetName, bool createSheetIfNotExists)
        {
            return new OpenXmlExcelRandomWriter(this, worksheetName, createSheetIfNotExists);
        }

        public IExcelRandomWriter CreateRandomWriter(int worksheetIndex)
        {
            return new OpenXmlExcelRandomWriter(this, GetSheetName(worksheetIndex), false);
        }

        public int? CountRows(string worksheetName)
        {
            try
            {
                var sheet = GetSheet(worksheetName, false);
                var txt = sheet.SheetDimension.InnerText;
                var endCellReference = txt.Split(':').Last();
                return ExcelUtilities.GetRowNo(endCellReference);
            }
            catch
            {
                return null;
            }
        }

        public int? CountRows(int worksheetIndex)
        {
            return CountRows(GetSheetName(worksheetIndex));
        }

        private void RemakeSharedStringTable()
        {
            var newtable = new SharedStringsStorage();
            foreach (var part in spreadsheetDocument.WorkbookPart.GetPartsOfType<WorksheetPart>())
            {
                Worksheet worksheet = part.Worksheet;
                foreach (var cell in worksheet.GetFirstChild<SheetData>().Descendants<Cell>())
                {
                    int oldkey;
                    if (cell.DataType != null
                        && cell.DataType.Value == CellValues.SharedString
                        && int.TryParse(cell.CellValue.Text, out oldkey))
                    {
                        var strValue = sharedStrings[oldkey];
                        int newkey;
                        newtable.Add(strValue, out newkey);
                        cell.CellValue.Text = newkey.ToString();
                    }
                }
            }
            var table = spreadsheetDocument.WorkbookPart.GetPartsOfType<SharedStringTablePart>().First().SharedStringTable;
            table.RemoveAllChildren();
            foreach (var strValue in newtable)
                table.AppendChild(new SharedStringItem(new DocumentFormat.OpenXml.Spreadsheet.Text(strValue)));
        }

        #region Saving file
        const int BUNCH_SIZE = 5000;
        int bunchCount = BUNCH_SIZE;
        bool pendingChanges = false;
        bool pendingSharedStringDeletes = false;
        bool pendingSharedStringAddings = false;
        bool pendingStyleChanges = false;

        private void DeferredSubmit(bool doNow)
        {
            var doSubmit = doNow || bunchCount-- <= 1;
            bunchCount = doSubmit ? BUNCH_SIZE : bunchCount;
            /*
            if (doSubmit && pendingSharedStringDeletes)
                RemakeSharedStringTable();
             */
            if (doSubmit && (pendingChanges || pendingSharedStringDeletes))
                spreadsheetDocument.WorkbookPart.Workbook.Save();
            if (doSubmit && (pendingSharedStringAddings || pendingSharedStringDeletes))
                spreadsheetDocument.WorkbookPart.SharedStringTablePart.SharedStringTable.Save();
            if (doSubmit && pendingStyleChanges)
                Stylesheet.Save();
            pendingChanges = pendingChanges && !doSubmit;
            pendingSharedStringAddings = pendingSharedStringAddings && !doSubmit;
            pendingSharedStringDeletes = pendingSharedStringDeletes && !doSubmit;
            pendingStyleChanges = pendingStyleChanges && !doSubmit;
        }

        internal void DeferredSubmit()
        {
            pendingChanges = true;
            DeferredSubmit(false);
        }

        internal void DeferredSubmitSharedStringTableAddings()
        {
            pendingSharedStringAddings = true;
            DeferredSubmit(false);
        }

        internal void DeferredSubmitSharedStringTableDeletes()
        {
            pendingSharedStringDeletes = true;
            DeferredSubmit(false);
        }

        internal void DeferredSubmitStyle()
        {
            pendingStyleChanges = true;
            DeferredSubmit(false);
        }
        #endregion
    }
}
