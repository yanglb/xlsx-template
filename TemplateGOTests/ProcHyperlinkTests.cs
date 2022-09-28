using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Linq;
using TemplateGOTests;

namespace TemplateGO.Tests
{
    [TestClass()]
    public class ProcHyperlinkTests
    {
        [TestMethod()]
        public void HyperlinkTest()
        {
            TemplateGO.Render(R.FullPath("data/hyperlink.xlsx"), R.JsonFromFile("data/hyperlink.json"), "hyperlink-out.xlsx");

            // 应该能打开文档
            using (var doc = SpreadsheetDocument.Open("hyperlink-out.xlsx", false))
            {
                Assert.IsNotNull(doc);
                Assert.IsNotNull(doc.WorkbookPart);
                var sheets = doc.WorkbookPart.Workbook.Descendants<Sheet>();
                Assert.IsNotNull(sheets);

                // 所有的连接应该为 hyperink.json 中的
                var siPart = doc.WorkbookPart.WorksheetParts.GetEnumerator();
                while (siPart.MoveNext())
                {
                    foreach (var link in siPart.Current?.HyperlinkRelationships!)
                    {
                        var links = siPart.Current.Worksheet.Descendants<Hyperlinks>().FirstOrDefault();
                        var hLink = links?.Descendants<Hyperlink>().Where(r => r.Id == link.Id).FirstOrDefault();
                        Assert.AreEqual(
                            link.Uri.ToString(),
                            "http://yanglb.com/from-test",
                            $"Sheet = {siPart.Current.Worksheet.LocalName}, Cell = {hLink?.Reference} 数据不正确"
                        );
                    }
                }
            }
        }
    }
}
