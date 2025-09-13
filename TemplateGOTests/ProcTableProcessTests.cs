using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Collections.Generic;
using System.Linq;
using TemplateGOTests;

namespace TemplateGO.Tests
{
    [TestClass()]
    public class ProcTableProcessTests
    {
        // 表格
        [TestMethod()]
        public void TableProcessTest()
        {
            var outFile = R.OutFullPath("table-process-out.xlsx");
            TemplateRender.Render(R.FullPath("data/table-process.xlsx"), R.JsonFromFile("data/table.json"), outFile, new TemplateOptions()
            {
                Transforms = new Dictionary<string, TransformDelegate>()
                {
                    { "avatar", AvatarTransform },
                    { "userProfile", UserProfileTransform  }
                }
            });

            // 应该能打开文档
            using var doc = SpreadsheetDocument.Open(outFile, false);
            Assert.IsNotNull(doc);
            Assert.IsNotNull(doc.WorkbookPart);
            var sheets = doc.WorkbookPart.Workbook.Descendants<Sheet>();
            Assert.IsNotNull(sheets);

            // 检查B列数据是否正常
            var sheetPart = doc.WorkbookPart.GetPartById(sheets.FirstOrDefault()!.Id!) as WorksheetPart;
            var shareStringTable = doc.WorkbookPart.SharedStringTablePart?.SharedStringTable;
            Assert.IsNotNull(sheetPart);

            var nameExpected = "贾阳煦,钟惠玲,孟三春,边念蕾,终苏凌,勾融雪,甄映寒,濮高寒,范弘致,公冰洁";
            var actualNames = R.ColumnStrings(sheetPart.Worksheet, shareStringTable, "B", 4, 13);
            var names = string.Join(',', actualNames);
            Assert.AreEqual(nameExpected, names);

            // 所有的连接应该正确
            var siPart = doc.WorkbookPart.WorksheetParts.GetEnumerator();
            List<string> linkList = [];
            while (siPart.MoveNext())
            {
                foreach (var link in siPart.Current?.HyperlinkRelationships!)
                {
                    linkList.Add(link.Uri.ToString());
                }
            }
            CollectionAssert.AreEqual(linkList, actualNames.Select(ProfileLink).ToList());
        }

        private object? AvatarTransform(object? value, TransformOptions options)
        {
            if (value is string stringValue) return R.FullPath(stringValue);
            return value;
        }

        private object? UserProfileTransform(object? value, TransformOptions options)
        {
            if (value is string stringValue) return ProfileLink(stringValue);
            return value;
        }

        private static string ProfileLink(string name) => $"https://yanlgb.com/profile/{name}";
    }
}
