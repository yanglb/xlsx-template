using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using TemplateGO.Parser;

namespace TemplateGO.Tests.Parser
{
    [TestClass()]
    public class GrammarTests
    {
        [TestMethod()]
        public void GrammarTestBadContent()
        {
            Assert.ThrowsException<ArgumentException>(() => new Grammar(""));
            Assert.ThrowsException<ArgumentException>(() => new Grammar("hello"));
            Assert.ThrowsException<ArgumentException>(() => new Grammar("hello}"));
            Assert.ThrowsException<ArgumentException>(() => new Grammar("{hello}"));
            Assert.ThrowsException<ArgumentException>(() => new Grammar("$hello}"));
        }

        [TestMethod()]
        public void GrammarTestEmpty()
        {
            var parser = new Grammar("${}");
            Assert.IsNotNull(parser);
            Assert.AreEqual(parser.Origin, "${}");
            Assert.AreEqual(parser.Property, "");
            Assert.AreEqual(parser.Processor, Processor.ProcessorType.Value);
            Assert.AreEqual(parser.Options.Count, 0);
        }

        [TestMethod()]
        public void GrammarTestSimple()
        {
            var parser = new Grammar("${hello}");
            Assert.IsNotNull(parser);
            Assert.AreEqual(parser.Origin, "${hello}");
            Assert.AreEqual(parser.Property, "hello");
            Assert.AreEqual(parser.Processor, Processor.ProcessorType.Value);
            Assert.AreEqual(parser.Options.Count, 0);
        }

        [TestMethod()]
        public void GrammarTestHaveProcessor()
        {
            var parser = new Grammar("${hello|test}");
            Assert.IsNotNull(parser);
            Assert.AreEqual(parser.Origin, "${hello|test}");
            Assert.AreEqual(parser.Property, "hello");
            Assert.AreEqual(parser.Processor, "test");
            Assert.AreEqual(parser.Options.Count, 0);
        }

        [TestMethod()]
        public void GrammarTestHaveProcessorAndOptions()
        {
            var parser = new Grammar("${hello|test:flag,user=yang,name=123}");
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
        public void GrammarTestHaveProcessorAndOptionsNoProp()
        {
            var parser = new Grammar("${|test:flag,user=yang,name=123}");
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