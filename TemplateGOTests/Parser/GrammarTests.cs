using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
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
    }
}