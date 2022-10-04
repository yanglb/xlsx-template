using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace TemplateGO.Parser.Tests
{
    [TestClass()]
    public class ImageOptionsTests
    {
        [DataRow("${t|image}", 0, null, null, false)]
        [DataRow("${t|image:padding=1.6in}", (long)1463040, null, null, false)]
        [DataRow("${t|image:fw=20px}", null, (long)190500, null, false)]
        [DataRow("${t|image:fh=2.5cm,deleteMarked}", null, null, (long)900000, true)]
        [DataRow("${t|image:padding=1.6in,fw=20px,fh=2.5cm}", (long)1463040, (long)190500, (long)900000, false)]
        [DataRow("${t|image:padding=1.6in,fw=20px,fh=2.5cm,haha=34,ttt=55,deleteMarked}", (long)1463040, (long)190500, (long)900000, true)]
        [TestMethod()]
        public void ImageOptionsTest(string input, long padding, long? fw, long? fh, bool deleteMarked)
        {
            var grammar = new Grammar(input);
            var res = new ImageOptions(grammar.Options);
            Assert.IsNotNull(res);
            Assert.AreEqual(padding, res.Padding);
            Assert.AreEqual(fw, res.FrameWidth);
            Assert.AreEqual(fh, res.FrameHeight);
            Assert.AreEqual(deleteMarked, res.DeleteMarked);
        }

        [TestMethod()]
        public void ImageOptionsTestEmpty()
        {
            var res = new ImageOptions();
            Assert.IsNotNull(res);
            Assert.AreEqual(0, res.Padding);
            Assert.IsNull(res.FrameWidth);
            Assert.IsNull(res.FrameHeight);
            Assert.IsFalse(res.DeleteMarked);
        }
    }
}