using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using TemplateGOTests;

namespace TemplateGO.Tests
{
    [TestClass()]
    public class ProcImageTests
    {
        [TestMethod()]
        public void ImageTest()
        {
            TemplateGO.Render(R.FullPath("data/image.xlsx"), R.JsonFromFile("data/image.json"), "image-out.xlsx");

            // 应该能打开文档
            using var doc = SpreadsheetDocument.Open("image-out.xlsx", false);
            Assert.IsNotNull(doc);
            Assert.IsNotNull(doc.WorkbookPart);
            var sheets = doc.WorkbookPart.Workbook.Descendants<Sheet>();
            Assert.IsNotNull(sheets);
        }
    }
}
