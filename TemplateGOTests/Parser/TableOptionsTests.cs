using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace TemplateGO.Parser.Tests
{
    [TestClass()]
    public class TableOptionsTests
    {
        [DataRow("${t|table}", 1, false)]
        [DataRow("${t|table: titleCount=2}", 2, false)]
        [DataRow("${t|table: keepExists}", 1, true)]
        [DataRow("${t|table:keepExists=dd, titleCount=3}", 3, true)]
        [TestMethod()]
        public void TableOptionsTest(string input, int titlaCount, bool keepExists)
        {
            var grammar = new Grammar(input);
            var res = new TableOptions(grammar.Options);
            Assert.IsNotNull(res);
            Assert.AreEqual(titlaCount, res.TitleCount);
            Assert.AreEqual(keepExists, res.KeepExists);
        }
    }
}