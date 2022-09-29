using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.IO;
using System.Text;

namespace TemplateGO.Tests
{
    [TestClass()]
    public class ImageUtilsTests
    {
        [TestMethod()]
        public void LocalFileTest()
        {
            var file = ImageUtils.ToLocalFile("local.png");
            Assert.AreEqual("local.png", file);
        }

        [TestMethod()]
        public void Base64Test()
        {
            var data = Encoding.UTF8.GetBytes("hello");
            var base64Data = $"data:image/png;base64,{Convert.ToBase64String(data)}";

            var file = ImageUtils.ToLocalFile(base64Data);
            Assert.IsTrue(file.EndsWith(".png"));
            Assert.IsTrue(File.Exists(file));

            var fileData = Encoding.UTF8.GetString(File.ReadAllBytes(file));
            Assert.AreEqual("hello", fileData);
        }

        [TestMethod()]
        public void Base64Test2()
        {
            var data = Encoding.UTF8.GetBytes("hello");
            var base64Data = $"data:image/x-pcx,{Convert.ToBase64String(data)}";

            var file = ImageUtils.ToLocalFile(base64Data);
            Assert.IsTrue(file.EndsWith(".pcx"));
            Assert.IsTrue(File.Exists(file));

            var fileData = Encoding.UTF8.GetString(File.ReadAllBytes(file));
            Assert.AreEqual("hello", fileData);
        }

        [TestMethod()]
        public void Base64TestBad()
        {
            var data = Encoding.UTF8.GetBytes("hello");
            var base64Data = $"data:image/unknow;base64,{Convert.ToBase64String(data)}";
            Assert.ThrowsException<ArgumentOutOfRangeException>(() => ImageUtils.ToLocalFile(base64Data));
        }

        [TestMethod()]
        public void HttpsTest()
        {
            var file = ImageUtils.ToLocalFile("https://www.baidu.com/img/PCtm_d9c8750bed0b3c7d089fa7d55720d6cf.png");
            Assert.IsTrue(file.EndsWith(".png"));
            Assert.IsTrue(File.Exists(file));
        }

        public void HttpTest()
        {
            var file = ImageUtils.ToLocalFile("http://www.baidu.com/img/PCtm_d9c8750bed0b3c7d089fa7d55720d6cf.png");
            Assert.IsTrue(file.EndsWith(".png"));
            Assert.IsTrue(File.Exists(file));
        }
    }
}