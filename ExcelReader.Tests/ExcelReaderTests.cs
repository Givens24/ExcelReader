using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Linq;

namespace ExcelReader.Tests
{
    [TestClass]
    public class ExcelReaderTests
    {
        [TestMethod]
        [TestCategory("ExcelReader")]
        [TestCategory("Exception Testing")]
        [ExpectedException(typeof(InvalidOperationException), "File name cannot be empty.")]
        public void ExcelReader_Exception_Expected_When_File_Name_Is_Empty()
        {
            var excelReader = new ExcelReader("", 1);
        }

        [TestMethod]
        [TestCategory("ExcelReader")]
        [TestCategory("Exception Testing")]
        [ExpectedException(typeof(InvalidOperationException), "File provided was not an excel file.")]
        public void ExcelReader_Exception_Expected_When_File_Name_Provided_Is_Not_In_Excel_Format()
        {
            var excelReader = new ExcelReader("C:\\TestBook.docx", 1);
        }

        [TestMethod]
        [TestCategory("ExcelReader")]
        [TestCategory("Exception Testing")]
        [ExpectedException(typeof(InvalidOperationException), "The file provided does not exist.")]
        public void ExcelReader_Exception_Expected_When_File_Does_Not_Exist()
        {
            var excelReader = new ExcelReader("C:\\TestBook1.xlsx", 1);
        }

        [TestMethod]
        [TestCategory("ExcelReader")]
        [TestCategory("Exception Testing")]
        [ExpectedException(typeof(InvalidOperationException), "Excel worksheet at index 0 was not found")]
        public void ExcelReader_Exception_Expected_When_Worksheet_Cannot_Be_Found()
        {
            var excelReader = new ExcelReader("C:\\TestBook.xlsx", 2);
        }

        [TestMethod]
        [TestCategory("ExcelReader")]
        [TestCategory("Integration")]
        [TestCategory("File IO")]
        public void FindRows_Successfully_Find_Rows_By_Cell_Value()
        {
            var excelReader = new ExcelReader("C:\\TestBook.xlsx", 1);
            var results = excelReader.FindRows(x => x.Cells.Any(cell => cell.Value == "1234"));

            Assert.IsTrue(results.Count() == 2);
        }

        [TestMethod]
        [TestCategory("ExcelReader")]
        [TestCategory("Integration")]
        [TestCategory("File IO")]
        public void FindRows_Successfully_Find_Rows_By_Row_Index()
        {
            var excelReader = new ExcelReader("C:\\TestBook.xlsx", 1);
            var results = excelReader.FindRows(x => x.Index == 1);

            Assert.IsTrue(results.Any(x => x.Cells.Any(cell => cell.Value == "Grant")));
        }
    }
}
