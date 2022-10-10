using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Linq;
using TemplateGO.Parser;

namespace TemplateGO.Tests.Parser
{
    [TestClass()]
    public class GrammarTests
    {
        [DataRow("")]
        [DataRow("hello")]
        [DataRow("hello}")]
        [DataRow("{hello}")]
        [DataRow("$hello}")]
        [TestMethod()]
        public void GrammarTestBadContent(string value)
        {
            Assert.ThrowsException<ArgumentException>(() => new Grammar(value));
        }

        [DataRow(null, false)]
        [DataRow("", false)]
        [DataRow("hello", false)]
        [DataRow("${", false)]
        [DataRow("${}", true)]
        [DataRow("abc${}test", true)]
        [DataRow("abc${ab|te:jj}test", true)]
        [TestMethod()]
        public void IsMatchTest(string? input, bool expected)
        {
            Assert.AreEqual(expected, Grammar.IsMatch(input));
        }

        [DataRow("", null)]
        [DataRow(null, null)]
        [DataRow("${hel", null)]
        [DataRow("${}", "${}")]
        [DataRow("${prop|proc|t1|t2:k=v,k2=v2}", "${prop|proc|t1|t2:k=v,k2=v2}")]
        [DataRow("T${prop|proc|t1|t2:k=v,k2=v2}", "${prop|proc|t1|t2:k=v,k2=v2}")]
        [DataRow("${prop|proc|t1|t2:k=v,k2=v2}S", "${prop|proc|t1|t2:k=v,k2=v2}")]
        [DataRow("T${prop|proc|t1|t2:k=v,k2=v2}S", "${prop|proc|t1|t2:k=v,k2=v2}")]
        [DataRow("Hello ${user}, this is your ${gitf|image}", "${user},${gitf|image}")]
        [DataRow("Hello ${user[1].a1}, this is your ${gitf[2].s1|image|cm2m}", "${user[1].a1},${gitf[2].s1|image|cm2m}")]
        [TestMethod()]
        public void MatchesTest(string input, string? matchs)
        {
            var res = Grammar.Matches(input);
            if (matchs == null)
            {
                Assert.IsNull(res);
            }
            else
            {
                Assert.IsNotNull(res);
                Assert.AreEqual(matchs, string.Join(',', res.Select(r => r.Value)));
            }
        }

        [TestMethod()]
        public void GrammarTestEmpty()
        {
            var parser = new Grammar("${}");
            Assert.IsNotNull(parser);
            Assert.AreEqual("${}", parser.Origin);
            Assert.AreEqual("", parser.Property);
            Assert.AreEqual(Processor.ProcessorType.Value, parser.Processor);
            Assert.AreEqual(0, parser.Options.Count);
        }

        [TestMethod()]
        public void GrammarTestSimple()
        {
            var parser = new Grammar("${hello}");
            Assert.IsNotNull(parser);
            Assert.AreEqual("${hello}", parser.Origin);
            Assert.AreEqual("hello", parser.Property);
            Assert.AreEqual(Processor.ProcessorType.Value, parser.Processor);
            Assert.AreEqual(0, parser.Options.Count);
        }

        [TestMethod()]
        public void GrammarTestHaveProcessor()
        {
            var parser = new Grammar("${ hello | test }");
            Assert.IsNotNull(parser);
            Assert.AreEqual("${ hello | test }", parser.Origin);
            Assert.AreEqual("hello", parser.Property);
            Assert.AreEqual("test", parser.Processor);
            Assert.AreEqual(0, parser.Options.Count);
        }

        [TestMethod()]
        public void GrammarTestHaveProcessorAndOptions()
        {
            var parser = new Grammar("${hello| test : flag, user = yang , name = 123}");
            Assert.IsNotNull(parser);
            Assert.AreEqual("${hello| test : flag, user = yang , name = 123}", parser.Origin);
            Assert.AreEqual("hello", parser.Property);
            Assert.AreEqual("test", parser.Processor);

            // 检查选项
            Assert.AreEqual(3, parser.Options.Count);
            Assert.AreEqual("", parser.Options["flag"]);
            Assert.AreEqual("yang", parser.Options["user"]);
            Assert.AreEqual("123", parser.Options["name"]);
        }

        [TestMethod()]
        public void GrammarTestHaveProcessorAndOptionsNoProp()
        {
            var parser = new Grammar("${|test:flag,user=yang,name=123}");
            Assert.IsNotNull(parser);
            Assert.AreEqual("${|test:flag,user=yang,name=123}", parser.Origin);
            Assert.AreEqual("", parser.Property);
            Assert.AreEqual("test", parser.Processor);

            // 检查选项
            Assert.AreEqual(3, parser.Options.Count);
            Assert.AreEqual("", parser.Options["flag"]);
            Assert.AreEqual("yang", parser.Options["user"]);
            Assert.AreEqual("123", parser.Options["name"]);
        }

        [DataRow("${ prop }", "prop", "value", "", 0)]
        [DataRow("${}", "", "value", "", 0)]
        [DataRow("${|test}", "", "test", "", 0)]
        [DataRow("${|test|t1|t2}", "", "test", "t1,t2", 0)]
        [DataRow("${ prop | test | t1 | t2 : flag }", "prop", "test", "t1,t2", 1)]
        [DataRow("${ prop | | t1 | t2 : flag,user=yang }", "prop", "value", "t1,t2", 2)]
        [TestMethod()]
        public void GrammarTestWithTransform(string input, string property, string processor, string transforms, int optionCount)
        {
            var parser = new Grammar(input);
            Assert.IsNotNull(parser);
            Assert.AreEqual(property, parser.Property);
            Assert.AreEqual(processor, parser.Processor);
            Assert.AreEqual(optionCount, parser.Options.Count);

            // 转换器
            Assert.AreEqual(transforms, string.Join(',', parser.Transforms));
        }

        [DataRow("prop", "prop", "value", "", 0)]
        [DataRow("=prop|link|eval|t2:k=v,k2=v2", "=prop|link|eval|t2:k=v,k2=v2", "value", "", 0)]
        [DataRow("=:prop|link|eval|t2:k=v,k2=v2", "=:prop|link|eval|t2:k=v,k2=v2", "value", "", 0)]
        [DataRow("prop|link|eval|t2:k=v,k2=v2", "prop", "link", "eval,t2", 2)]
        [DataRow("#prop|image|eval|t2:k,k2=v2", "#prop", "image", "eval,t2", 2)]
        [TestMethod()]
        public void GrammarTestWithTable(string input, string property, string processor, string transforms, int optionCount)
        {
            var parser = new Grammar(input, true);
            Assert.IsNotNull(parser);
            Assert.AreEqual(property, parser.Property);
            Assert.AreEqual(processor, parser.Processor);
            Assert.AreEqual(optionCount, parser.Options.Count);

            // 转换器
            Assert.AreEqual(transforms, string.Join(',', parser.Transforms));
        }
    }
}