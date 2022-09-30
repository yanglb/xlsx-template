using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace TemplateGO.Parser.Tests
{
    [TestClass()]
    public class ImageOptionsTests
    {
        [DataRow("${t|image}", 0, null, null)]
        [DataRow("${t|image:padding=1.6in}", (long)1463040, null, null)]
        [DataRow("${t|image:fw=20px}", null, (long)190500, null)]
        [DataRow("${t|image:fh=2.5cm}", null, null, (long)900000)]
        [DataRow("${t|image:padding=1.6in,fw=20px,fh=2.5cm}", (long)1463040, (long)190500, (long)900000)]
        [DataRow("${t|image:padding=1.6in,fw=20px,fh=2.5cm,haha=34,ttt=55}", (long)1463040, (long)190500, (long)900000)]
        [TestMethod()]
        public void ImageOptionsTest(string input, long padding, long? fw, long? fh)
        {
            var grammar = new Grammar(input);
            var res = new ImageOptions(grammar.Options);
            Assert.IsNotNull(res);
            Assert.AreEqual(padding, res.Padding);
            Assert.AreEqual(fw, res.FrameWidth);
            Assert.AreEqual(fh, res.FrameHeight);
        }
    }
}