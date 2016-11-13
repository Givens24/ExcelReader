using System.Collections.Generic;

namespace ExcelReader.Models
{
    public class ExcelRow
    {
        public int Index { get; set; }
        private readonly List<ExcelCell> _cells;
        public IEnumerable<ExcelCell> Cells => _cells;
        public ExcelRow()
        {
            _cells = new List<ExcelCell>();
        }

        internal void AddCell(ExcelCell cell)
        {
            _cells.Add(cell);
        }
    }
}
