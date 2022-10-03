using TemplateGO.Utils;
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

        [DataRow("A1:C10", "B5:B6", true)]
        [DataRow("A1:C10", "C10:C15", true)]
        [DataRow("A1:C10", "C11:C15", false)]
        [DataRow("A1:C10", "D10:D15", false)]
        [DataRow("A1:C10", "D20:D25", false)]
        [TestMethod()]
        public void IsIntersectTest(string range1, string range2, bool expected)
        {
            var res1 = CellUtils.IsIntersect(range1, range2);
            Assert.AreEqual(expected, res1);

            var res2 = CellUtils.IsIntersect(range2, range1);
            Assert.AreEqual(expected, res2);
        }

        [DataRow("2:4", "3:3", true)]
        [DataRow("2:4", "1:5", true)]
        [DataRow("2:4", "1:1", false)]
        [TestMethod()]
        public void IsRowIntersectTest(string range1, string range2, bool expected)
        {
            var res1 = CellUtils.IsRowIntersect(range1, range2);
            Assert.AreEqual(expected, res1);

            var res2 = CellUtils.IsRowIntersect(range2, range1);
            Assert.AreEqual(expected, res2);
        }

        [DataRow("B:D", "B:B", true)]
        [DataRow("B:D", "C:C", true)]
        [DataRow("B:D", "A:E", true)]
        [DataRow("B:D", "A:A", false)]
        [TestMethod()]
        public void IsColumnIntersectTest(string range1, string range2, bool expected)
        {
            var res1 = CellUtils.IsColumnIntersect(range1, range2);
            Assert.AreEqual(expected, res1);

            var res2 = CellUtils.IsColumnIntersect(range2, range1);
            Assert.AreEqual(expected, res2);
        }

        [DataRow("", "")]
        [DataRow("AF", "12")]
        [TestMethod()]
        public void IsIntersectTestBad(string range1, string range2)
        {
            Assert.ThrowsException<ArgumentException>(() => CellUtils.IsIntersect(range1, range2));
            Assert.ThrowsException<ArgumentException>(() => CellUtils.IsRowIntersect(range1, range2));
            Assert.ThrowsException<ArgumentException>(() => CellUtils.IsColumnIntersect(range1, range2));
        }

        [DataRow("A1:D10", 1, "A1:D11")]
        [DataRow("A1:D10", 0, "A1:D10")]
        [DataRow("A1:D10", -1, "A1:D9")]
        [TestMethod()]
        public void UpdateToReferenceTest(string origin, int shift, string expected)
        {
            var res = CellUtils.UpdateToReference(origin, shift);
            Assert.AreEqual(expected, res);
        }

        [TestMethod()]
        public void UpdateToReferenceTestBad()
        {
            Assert.ThrowsException<ArgumentException>(() => CellUtils.UpdateToReference("", 0));
            Assert.ThrowsException<ArgumentException>(() => CellUtils.UpdateToReference("A2D10", 0));

            Assert.ThrowsException<ArgumentOutOfRangeException>(() => CellUtils.UpdateToReference("A2:D10", -10));
            Assert.ThrowsException<ArgumentOutOfRangeException>(() => CellUtils.UpdateToReference("A2:D10", -11));
        }

        [DataRow("'te st'!$C$10:$F$20", 0, "'te st'!$C$10:$F$20")]
        [DataRow("'te st'!$C$10:$F$20", 10, "'te st'!$C$10:$F$30")]
        [DataRow("'te st'!$C$10:$F$20", -10, "'te st'!$C$10:$F$10")]
        [DataRow("测试!C10:F20", 0, "测试!C10:F20")]
        [DataRow("测试!C10:F20", 10, "测试!C10:F30")]
        [DataRow("测试!C10:F20", -10, "测试!C10:F10")]
        [TestMethod()]
        public void ExpandRowReferenceTest(string input, int shift, string expected)
        {
            Assert.AreEqual(expected, CellUtils.ExpandRowReference(input, shift));
        }

        [DataRow("测试!C10F20", 10)]
        [DataRow("测试!C10:F20T", 10)]
        [DataRow("测试!C10:FT", 0)]
        [DataRow("测试!C10:$10", 0)]
        [TestMethod()]
        public void ExpandRowReferenceTestBad(string input, int shift)
        {
            Assert.ThrowsException<ArgumentException>(() => CellUtils.ExpandRowReference(input, shift));
        }
    }
}