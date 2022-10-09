using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using TemplateGOTests;

namespace TemplateGO.Tests
{
    /// <summary>
    /// 测试数据转换
    /// </summary>
    [TestClass()]
    public class TransformTests
    {
        [TestMethod()]
        public void Test()
        {
            var outFile = R.OutFullPath("transform-out.xlsx");
            TemplateRender.Render(R.FullPath("data/transform.xlsx"), R.JsonFromFile("data/transform.json"), outFile, new TemplateOptions()
            {
                Transforms = new System.Collections.Generic.Dictionary<string, TransformDelegate>() {
                    {"gender", GenderTransform },
                    { "cm2m", Cm2MTransform },
                    { "score", ScoreTransform },
                    { "name1", NameTransform },
                    { "name2", NameTransform },
                }
            });

            // 应该能打开文档
            using var doc = SpreadsheetDocument.Open(outFile, false);
            Assert.IsNotNull(doc);
            Assert.IsNotNull(doc.WorkbookPart);
            var sheets = doc.WorkbookPart.Workbook.Descendants<Sheet>();
            Assert.IsNotNull(sheets);

            var sheetName = "数据转换";

            // 检查单元格内容
            Assert.AreEqual("张三", R.CellStringValue(doc, sheetName, "C3"));
            Assert.AreEqual("李四", R.CellStringValue(doc, sheetName, "D3"));

            // 性别
            Assert.AreEqual("male", R.CellStringValue(doc, sheetName, "C4"));
            Assert.AreEqual("female", R.CellStringValue(doc, sheetName, "D4"));
            Assert.AreEqual("男孩", R.CellStringValue(doc, sheetName, "C5"));
            Assert.AreEqual("女孩", R.CellStringValue(doc, sheetName, "D5"));

            // 身高
            Assert.AreEqual("160", R.CellStringValue(doc, sheetName, "C6"));
            Assert.AreEqual("175", R.CellStringValue(doc, sheetName, "D6"));
            Assert.AreEqual("1.6", R.CellStringValue(doc, sheetName, "C7"));
            Assert.AreEqual("1.75", R.CellStringValue(doc, sheetName, "D7"));

            // 分数
            Assert.AreEqual("85.5", R.CellStringValue(doc, sheetName, "C8"));
            Assert.AreEqual("70", R.CellStringValue(doc, sheetName, "D8"));
            Assert.AreEqual("优秀", R.CellStringValue(doc, sheetName, "C9"));
            Assert.AreEqual("一般", R.CellStringValue(doc, sheetName, "D9"));

            // 不存在的转换器
            Assert.AreEqual("85.5", R.CellStringValue(doc, sheetName, "C10"));
            Assert.AreEqual("70", R.CellStringValue(doc, sheetName, "D10"));
            Assert.AreEqual("张三|T|T", R.CellStringValue(doc, sheetName, "C11"));
            Assert.AreEqual("李四|T", R.CellStringValue(doc, sheetName, "D11"));
        }

        private object? GenderTransform(object? value, TransformOptions options)
        {
            if (value == null) return null;

            if (value.ToString() == "male") return "男孩";
            if (value.ToString() == "female") return "女孩";

            return "未知";
        }

        private object? NameTransform(object? value, TransformOptions options)
        {
            if (value == null) return null;
            return $"{value}|T";
        }

        private object? Cm2MTransform(object? value, TransformOptions options)
        {
            if (value == null) return null;
            var cm = (int)value;
            var m = cm / 100.0;
            return Math.Round(m, 2);
        }

        private object? ScoreTransform(object? value, TransformOptions options)
        {
            if (value == null) return null;
            double score = double.Parse(value.ToString()!);
            if (score >= 80) return "优秀";
            return "一般";
        }
    }
}