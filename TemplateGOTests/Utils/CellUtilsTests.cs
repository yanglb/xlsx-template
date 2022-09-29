using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace TemplateGO.Utils.Tests
{
    [TestClass()]
    public class CellUtilsTests
    {
        [DataRow(null, null)]
        [DataRow("", null)]
        [DataRow("A1", 1)]
        [DataRow("B10", 2)]
        [DataRow("Z2", 26)]
        [DataRow("AA1", 27)]
        [DataRow("AB1", 28)]
        [DataRow("ZA2", 677)]
        [TestMethod()]
        public void RowValueTest(string? cell, int? column)
        {
            Assert.AreEqual(column, CellUtils.RowValue(cell));
        }

        [DataRow(null, null)]
        [DataRow("", null)]
        [DataRow("A1", 1)]
        [DataRow("A10", 10)]
        [DataRow("Z2", 2)]
        [TestMethod()]
        public void ColumnValueTest(string? cell, int? column)
        {
            Assert.AreEqual(column, CellUtils.ColumnValue(cell));
        }
    }
}