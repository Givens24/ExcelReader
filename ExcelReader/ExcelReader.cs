using ExcelReader.Models;
using ExcelReader.Resources;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelReader
{
    public class ExcelReader
    {
        private readonly Excel._Application _excelApplication;
        private readonly Excel._Workbook _workBook;
        private readonly ExcelSpreadSheetModel _excelWorkSheet;

        public ExcelReader(string excelFileName, int workSheetIndex)
        {
            if (string.IsNullOrEmpty(excelFileName))
            {
                throw new InvalidOperationException(ExcelReaderErrorMessages.FILE_NAME_IS_EMPTY_ERROR);
            }

            if (!Path.GetExtension(excelFileName).Contains(".xls"))
            {
                throw new InvalidOperationException(ExcelReaderErrorMessages.FILE_INCORRECT_FORMAT_ERROR);
            }

            if (!File.Exists(excelFileName))
            {
                throw new InvalidOperationException(ExcelReaderErrorMessages.FILE_DOES_NOT_EXIST_ERROR);
            }

            _excelApplication = new Excel.Application();
            _workBook = _excelApplication.Workbooks.Open(excelFileName);
            _excelWorkSheet = new ExcelSpreadSheetModel();
            ReadWorkSheetIntoInternalModel(workSheetIndex);
        }

        private void ReadWorkSheetIntoInternalModel(int workSheetIndex)
        {
            Excel.Worksheet workSheet;
            if (!TryFindWorksheet(out workSheet, workSheetIndex))
            {
                throw new InvalidOperationException(string.Format(ExcelReaderErrorMessages.EXCEL_WORKSHEET_NOT_FOUND_ERROR, workSheetIndex));
            }

            workSheet.UsedRange.Rows.Cast<Excel.Range>().ToList().ForEach(row =>
            {
                var excelRow = new ExcelRow {Index = row.Column};
                row.Cells.Cast<Excel.Range>().ToList().ForEach(cell =>
                {
                    var excelCell = new ExcelCell
                    {
                        Row = excelRow,
                        Value = Convert.ToString(cell.Value2)
                    };
                    excelRow.AddCell(excelCell);
                });
                _excelWorkSheet.AddRow(excelRow);
            });
            _workBook.Close();
            _excelApplication.Quit();
        }

        private bool TryFindWorksheet(out Excel.Worksheet workSheet, int worksheetIndex)
        {
            try
            {
               workSheet = _workBook.Worksheets[worksheetIndex];
               return true;
            }
            catch (Exception)
            {
                workSheet = null;
                return false;
            }
        }

        public IEnumerable<ExcelRow> FindRows(Func<ExcelRow, bool> whereExpression)
        {
            return _excelWorkSheet.Rows.Where(whereExpression);
        }
    }
}
