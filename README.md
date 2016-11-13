# ExcelReader
A library for quickly reading excel spreadsheets with LINQ.

This library has one function called **FindRows** that provides a quick and easy way of searching an Excel spreadsheet data.

##FindRows
The **ExcelReader** takes an excel file path and a worksheet index. The FindRows function takes a ```Func<ExcelRow, bool>``` in order to search the spreadsheet by the library's internal models **ExcelRow** and **ExcelCell**.
The following example shows the proper usage of the library and **FindRows** function.

```C#
var excelReader = new ExcelReader("C:\\TestBook.xlsx", 1);
var results = excelReader.FindRows(x => x.Cells.Any(cell => cell.Value == "1234"));

Assert.IsTrue(results.Count() == 2);
```
