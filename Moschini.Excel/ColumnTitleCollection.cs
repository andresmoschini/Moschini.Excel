using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Moschini.Excel
{
    //TODO: remove this class it is so much complex and it show be enough with a dictionary
    public class ColumnTitleCollection: List<ColumnTitle>
    {
        public ColumnTitle this[string titleName]
        {
            get { return this.FirstOrDefault(x => x.TitleText == titleName); }
            set
            {
                var existing = this[titleName];
                if (existing != null) this.Remove(existing);

                this.Add(value);
            }
        }

        public static ColumnTitleCollection ReadTitlesRow(IExcelReader rdr)
        {
            var titles = new ColumnTitleCollection();

            if (rdr.Read())
            {
                titles.AddRange(Enumerable
                    .Range(0, rdr.FieldCount)
                    .Select(i => new ColumnTitle()
                    {
                        ColumnLocation = ExcelUtilities.OrdinalToColumnName(i + 1),
                        TitleText = rdr.GetValue<string>(i)
                    }));
            }
            return titles;
        }

        public string GetLocationByColumnTitle(string columnTitle, bool throwException = true)
        {
            if (!this.ContainsColumnTitle(columnTitle))
            {
                var trimed = columnTitle.TrimEnd(new char[] { '\'' });
                if (columnTitle.Length - trimed.Length > 10)
                {
                    if (throwException)
                    {
                        throw new ArgumentOutOfRangeException("columnTitle", string.Format("Column title '{0}' does not exist. All titles: {1}", trimed, string.Join(", ", this.Select(x => x.TitleText).ToArray())));
                    }
                    else
                    {
                        return null;
                    }
                }
                else
                {
                    return GetLocationByColumnTitle(columnTitle + "'", throwException);
                }
            }

            return this[columnTitle].ColumnLocation;
        }

        public bool ContainsColumnTitle(string columnTitle)
        {
            return this[columnTitle] != null;
        }
    }
}
