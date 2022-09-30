using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Linq;
using TemplateGOTests;

namespace TemplateGO.Utils.Tests
{
    [TestClass()]
    public class CellUtilsTests
    {
        [DataRow("A1", 1)]
        [DataRow("A10", 10)]
        [DataRow("Z2", 2)]
        [TestMethod()]
        public void RowValueTest(string cell, int column)
        {
            Assert.AreEqual(column, CellUtils.RowValue(cell));
        }

        [DataRow("A1", 1)]
        [DataRow("B10", 2)]
        [DataRow("Z2", 26)]
        [DataRow("AA1", 27)]
        [DataRow("AB1", 28)]
        [DataRow("ZA2", 677)]
        [TestMethod()]
        public void ColumnValueTest(string cell, int column)
        {
            Assert.AreEqual(column, CellUtils.ColumnValue(cell));
        }

        [TestMethod()]
        public void EmptyValueTest()
        {
            Assert.ThrowsException<ArgumentException>(() => CellUtils.RowValue(""));
            Assert.ThrowsException<ArgumentException>(() => CellUtils.ColumnValue(""));
        }

        [TestMethod()]
        public void CellWidthTestNotSet()
        {
            using var doc = SpreadsheetDocument.Open(R.FullPath("data/cell-size.xlsx"), false);
            Assert.IsNotNull(doc);

            var sheet1 = doc.WorkbookPart?.Workbook.Descendants<Sheet>().Where(r => r.Name == "Sheet1").FirstOrDefault();
            Assert.IsNotNull(sheet1);

            var sheetPart = doc?.WorkbookPart?.GetPartById(sheet1.Id!) as WorksheetPart;
            Assert.IsNotNull(sheetPart);

            var cell = sheetPart.Worksheet.Descendants<Cell>().Where(r => r.CellReference == "A1").FirstOrDefault();
            Assert.IsNotNull(cell);

            //var res = CellUtils.CellWidth(cell, sheetPart);
            //Assert.AreEqual(12, res);
        }

        [TestMethod()]
        public void CellWidthTest()
        {
            using var doc = SpreadsheetDocument.Open(R.FullPath("data/cell-size.xlsx"), false);
            Assert.IsNotNull(doc);

            var sheet1 = doc.WorkbookPart?.Workbook.Descendants<Sheet>().Where(r => r.Name == "Sheet2").FirstOrDefault();
            Assert.IsNotNull(sheet1);

            var sheetPart = doc?.WorkbookPart?.GetPartById(sheet1.Id!) as WorksheetPart;
            Assert.IsNotNull(sheetPart);

            var cell = sheetPart.Worksheet.Descendants<Cell>().Where(r => r.CellReference == "A1").FirstOrDefault();
            Assert.IsNotNull(cell);

            //var res = CellUtils.CellWidth(cell, sheetPart);
            //Assert.AreEqual(12, res);
        }
    }
}