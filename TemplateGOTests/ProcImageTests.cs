using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Collections.Generic;
using System.Text.Json;
using TemplateGOTests;

namespace TemplateGO.Tests
{
    [TestClass()]
    public class ProcImageTests
    {
        private static readonly Dictionary<string, string> TestData = new()
        {
            { "url", "https://www.baidu.com/img/PCtm_d9c8750bed0b3c7d089fa7d55720d6cf.png" },
            { "file", R.FullPath("data/image.jpg") }
        };

        [TestMethod()]
        public void ImageTest()
        {
            var json = JsonSerializer.Serialize(TestData);
            TemplateGO.Render(R.FullPath("data/image.xlsx"), json, "image-out.xlsx");

            // 应该能打开文档
            using var doc = SpreadsheetDocument.Open("image-out.xlsx", false);
            Assert.IsNotNull(doc);
            Assert.IsNotNull(doc.WorkbookPart);
            var sheets = doc.WorkbookPart.Workbook.Descendants<Sheet>();
            Assert.IsNotNull(sheets);
        }
    }
}
