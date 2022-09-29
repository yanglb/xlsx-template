using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Linq;
using TemplateGO.Utils;
using TemplateGOTests;

namespace TemplateGO.Tests
{
    [TestClass()]
    public class PropValueTests
    {
        [TestMethod()]
        public void ValueTest()
        {
            TemplateGO.Render(R.FullPath("data/value.xlsx"), R.JsonFromFile("data/value.json"), "value-out.xlsx");

            // 应该能打开文档
            using (var doc = SpreadsheetDocument.Open("value-out.xlsx", false))
            {
                Assert.IsNotNull(doc);
                Assert.IsNotNull(doc.WorkbookPart);
                var sheets = doc.WorkbookPart.Workbook.Descendants<Sheet>();
                Assert.IsNotNull(sheets);

                var ssTable = doc.WorkbookPart?.SharedStringTablePart?.SharedStringTable;

                // 属性替换 工作表
                var sheet1 = doc.WorkbookPart?.Workbook.Descendants<Sheet>().Where(r =>r.Name == "属性替换").FirstOrDefault();
                Assert.IsNotNull(sheet1);

                var sheet1Part = doc.WorkbookPart?.GetPartById(sheet1.Id!) as WorksheetPart;
                Assert.IsNotNull(sheet1Part);

                // B列数据应该如下
                var cell = sheet1Part.Worksheet.Descendants<Cell>().Where(cell => cell.CellReference == "B3").FirstOrDefault();
                Assert.IsNotNull(cell);
                Assert.IsTrue(CellValues.String == cell.DataType!, cell.CellReference);
                Assert.AreEqual("测试名称", CellUtils.GetCellString(cell, ssTable), cell.CellReference);

                cell = sheet1Part.Worksheet.Descendants<Cell>().Where(cell => cell.CellReference == "B4").FirstOrDefault();
                Assert.IsNotNull(cell);
                Assert.IsTrue(CellValues.String == cell.DataType!, cell.CellReference);
                Assert.AreEqual("测试名称测试名称", CellUtils.GetCellString(cell, ssTable), cell.CellReference);

                cell = sheet1Part.Worksheet.Descendants<Cell>().Where(cell => cell.CellReference == "B5").FirstOrDefault();
                Assert.IsNotNull(cell);
                Assert.IsTrue(CellValues.Number == cell.DataType!, cell.CellReference);
                Assert.AreEqual("32", cell.InnerText, cell.CellReference);

                cell = sheet1Part.Worksheet.Descendants<Cell>().Where(cell => cell.CellReference == "B6").FirstOrDefault();
                Assert.IsNotNull(cell);
                Assert.IsTrue(CellValues.String == cell.DataType!, cell.CellReference);
                Assert.AreEqual("3232", cell.InnerText, cell.CellReference);

                cell = sheet1Part.Worksheet.Descendants<Cell>().Where(cell => cell.CellReference == "B7").FirstOrDefault();
                Assert.IsNotNull(cell);
                Assert.IsTrue(CellValues.Number == cell.DataType!, cell.CellReference);
                Assert.AreEqual("46.5", cell.InnerText, cell.CellReference);

                cell = sheet1Part.Worksheet.Descendants<Cell>().Where(cell => cell.CellReference == "B8").FirstOrDefault();
                Assert.IsNotNull(cell);
                Assert.IsTrue(CellValues.Boolean == cell.DataType!, cell.CellReference);
                Assert.AreEqual("true", cell.InnerText, cell.CellReference);

                cell = sheet1Part.Worksheet.Descendants<Cell>().Where(cell => cell.CellReference == "B9").FirstOrDefault();
                Assert.IsNotNull(cell);
                Assert.IsTrue(CellValues.String == cell.DataType!, cell.CellReference);
                Assert.AreEqual("学校名称", cell.InnerText, cell.CellReference);

                cell = sheet1Part.Worksheet.Descendants<Cell>().Where(cell => cell.CellReference == "B10").FirstOrDefault();
                Assert.IsNotNull(cell);
                Assert.IsTrue(CellValues.String == cell.DataType!, cell.CellReference);
                Assert.AreEqual("父亲姓名", cell.InnerText, cell.CellReference);

                cell = sheet1Part.Worksheet.Descendants<Cell>().Where(cell => cell.CellReference == "B11").FirstOrDefault();
                Assert.IsNotNull(cell);
                Assert.IsTrue(CellValues.String == cell.DataType!, cell.CellReference);
                Assert.AreEqual("https://yanglb.com", cell.InnerText, cell.CellReference);

                cell = sheet1Part.Worksheet.Descendants<Cell>().Where(cell => cell.CellReference == "B12").FirstOrDefault();
                Assert.IsNotNull(cell);
                Assert.IsTrue(CellValues.String == cell.DataType!, cell.CellReference);
                Assert.AreEqual("我得了46.5分", cell.InnerText, cell.CellReference);

                cell = sheet1Part.Worksheet.Descendants<Cell>().Where(cell => cell.CellReference == "B14").FirstOrDefault();
                Assert.IsNotNull(cell);
                Assert.IsTrue(CellValues.String == cell.DataType!, cell.CellReference);
                Assert.AreEqual("测试名称 的母亲叫 母亲姓名", cell.InnerText, cell.CellReference);

                // 处理程序不存在的属性
                cell = sheet1Part.Worksheet.Descendants<Cell>().Where(cell => cell.CellReference == "B17").FirstOrDefault();
                Assert.IsNotNull(cell);
                Assert.IsTrue(CellValues.SharedString == cell.DataType!, cell.CellReference);
                Assert.AreEqual("${name|nono}", CellUtils.GetCellString(cell, ssTable), cell.CellReference);

                cell = sheet1Part.Worksheet.Descendants<Cell>().Where(cell => cell.CellReference == "B18").FirstOrDefault();
                Assert.IsNotNull(cell);
                Assert.IsTrue(CellValues.SharedString == cell.DataType!, cell.CellReference);
                Assert.AreEqual("${name|nono}原样保留", CellUtils.GetCellString(cell, ssTable), cell.CellReference);

                // 属性为存在时
                cell = sheet1Part.Worksheet.Descendants<Cell>().Where(cell => cell.CellReference == "B19").FirstOrDefault();
                Assert.IsNotNull(cell);
                Assert.IsNull(cell.DataType, cell.CellReference);
                Assert.IsNull(cell.CellValue, cell.CellReference);

                cell = sheet1Part.Worksheet.Descendants<Cell>().Where(cell => cell.CellReference == "B20").FirstOrDefault();
                Assert.IsNotNull(cell);
                Assert.IsTrue(CellValues.String == cell.DataType!, cell.CellReference);
                Assert.AreEqual("不存在的属性", cell.InnerText, cell.CellReference);
            }
        }
    }
}