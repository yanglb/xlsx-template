using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Collections.Generic;
using System.Text.Json;
using TemplateGOTests;

namespace TemplateGO.Tests
{
    [TestClass()]
    public class ProcQRCodeTests
    {
        private static readonly Dictionary<string, string> TestData = new()
        {
            { "url", "https://yanglb.com/" },
            { "string", "这是二维码内容" }
        };

        [TestMethod()]
        public void QrTest()
        {
            var json = JsonSerializer.Serialize(TestData);
            TemplateGO.Render(R.FullPath("data/qr.xlsx"), json, "qr-out.xlsx");

            // 应该能打开文档
            using var doc = SpreadsheetDocument.Open("qr-out.xlsx", false);
            Assert.IsNotNull(doc);
            Assert.IsNotNull(doc.WorkbookPart);
            var sheets = doc.WorkbookPart.Workbook.Descendants<Sheet>();
            Assert.IsNotNull(sheets);
        }
    }
}
