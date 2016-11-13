using System.Collections.Generic;

namespace ExcelReader.Models
{
    public class ExcelSpreadSheetModel
    {
        private readonly List<ExcelRow> _rows;
        public IEnumerable<ExcelRow> Rows => _rows;
        public ExcelSpreadSheetModel()
        {
            _rows = new List<ExcelRow>();
        }

        internal void AddRow(ExcelRow row)
        {
            _rows.Add(row);
        }
    }
}
