using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace TemplateGO.Parser.Tests
{
    [TestClass()]
    public class TableOptionsTests
    {
        [DataRow("${t|table}", 1U, 0U)]
        [DataRow("${t|table: titleCount=2}", 2U, 0U)]
        [DataRow("${t|table: sampleCount=2}", 1U, 2U)]
        [DataRow("${t|table:sampleCount=10, titleCount=3}", 3U, 10U)]
        [TestMethod()]
        public void TableOptionsTest(string input, uint titlaCount, uint sampleCount)
        {
            var grammar = new Grammar(input);
            var res = new TableOptions(grammar.Options);
            Assert.IsNotNull(res);
            Assert.AreEqual(titlaCount, res.TitleCount);
            Assert.AreEqual(sampleCount, res.SampleCount);
        }
    }
}