using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;

namespace TemplateGO.Tests
{
    [TestClass()]
    public class ParserTests
    {
        [TestMethod()]
        public void ParserTestBadContent()
        {
            Assert.ThrowsException<ArgumentException>(() => new Parser(""));
            Assert.ThrowsException<ArgumentException>(() => new Parser("hello"));
            Assert.ThrowsException<ArgumentException>(() => new Parser("hello}"));
            Assert.ThrowsException<ArgumentException>(() => new Parser("{hello}"));
            Assert.ThrowsException<ArgumentException>(() => new Parser("$hello}"));
        }

        [TestMethod()]
        public void ParserTestEmpty()
        {
            var parser = new Parser("${}");
            Assert.IsNotNull(parser);
            Assert.AreEqual(parser.Origin, "${}");
            Assert.AreEqual(parser.Property, "");
            Assert.AreEqual(parser.Processor, Processor.ProcessorType.Value);
            Assert.AreEqual(parser.Options.Count, 0);
        }

        [TestMethod()]
        public void ParserTestSimple()
        {
            var parser = new Parser("${hello}");
            Assert.IsNotNull(parser);
            Assert.AreEqual(parser.Origin, "${hello}");
            Assert.AreEqual(parser.Property, "hello");
            Assert.AreEqual(parser.Processor, Processor.ProcessorType.Value);
            Assert.AreEqual(parser.Options.Count, 0);
        }

        [TestMethod()]
        public void ParserTestHaveProcessor()
        {
            var parser = new Parser("${hello|test}");
            Assert.IsNotNull(parser);
            Assert.AreEqual(parser.Origin, "${hello|test}");
            Assert.AreEqual(parser.Property, "hello");
            Assert.AreEqual(parser.Processor, "test");
            Assert.AreEqual(parser.Options.Count, 0);
        }

        [TestMethod()]
        public void ParserTestHaveProcessorAndOptions()
        {
            var parser = new Parser("${hello|test:flag,user=yang,name=123}");
            Assert.IsNotNull(parser);
            Assert.AreEqual(parser.Origin, "${hello|test:flag,user=yang,name=123}");
            Assert.AreEqual(parser.Property, "hello");
            Assert.AreEqual(parser.Processor, "test");

            // 检查选项
            Assert.AreEqual(parser.Options.Count, 3);
            Assert.AreEqual(parser.Options["flag"], "");
            Assert.AreEqual(parser.Options["user"], "yang");
            Assert.AreEqual(parser.Options["name"], "123");
        }

        [TestMethod()]
        public void ParserTestHaveProcessorAndOptionsNoProp()
        {
            var parser = new Parser("${|test:flag,user=yang,name=123}");
            Assert.IsNotNull(parser);
            Assert.AreEqual(parser.Origin, "${|test:flag,user=yang,name=123}");
            Assert.AreEqual(parser.Property, "");
            Assert.AreEqual(parser.Processor, "test");

            // 检查选项
            Assert.AreEqual(parser.Options.Count, 3);
            Assert.AreEqual(parser.Options["flag"], "");
            Assert.AreEqual(parser.Options["user"], "yang");
            Assert.AreEqual(parser.Options["name"], "123");
        }
    }
}