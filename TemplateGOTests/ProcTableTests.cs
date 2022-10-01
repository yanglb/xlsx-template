using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using TemplateGOTests;

namespace TemplateGO.Tests
{
    [TestClass()]
    public class ProcTableTests
    {
        [TestMethod()]
        public void TableTest()
        {
            TemplateGO.Render(R.FullPath("data/table.xlsx"), R.JsonFromFile("data/table.json"), "table-out.xlsx");

            // 应该能打开文档
            using var doc = SpreadsheetDocument.Open("table-out.xlsx", false);
            Assert.IsNotNull(doc);
            Assert.IsNotNull(doc.WorkbookPart);
            var sheets = doc.WorkbookPart.Workbook.Descendants<Sheet>();
            Assert.IsNotNull(sheets);
        }
    }
}
