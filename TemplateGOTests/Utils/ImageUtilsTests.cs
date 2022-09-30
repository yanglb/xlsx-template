using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.IO;
using System.Text;
using TemplateGO.Parser;
using TemplateGOTests;

namespace TemplateGO.Utils.Tests
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

        // 默认情况下应该为图片大小
        [DataRow("${t|image}", 0, 0, 3048000L, 4572000L)]

        // 指定边距
        [DataRow("${t|image:padding=1cm}", 360000L, 360000L, 3048000L, 4572000L)]

        // 指定宽度
        [DataRow("${t|image:padding=1cm,fw=10cm}", 360000L, 360000L, 2880000L, 4320000L)]  // >
        [DataRow("${t|image:fw=320px}", 0L, 0L, 3048000L, 4572000L)]                       // =
        [DataRow("${t|image:padding=1cm,fw=3in}", 360000L, 360000L, 2023200L, 3034800L)]   // <

        // 指定高度
        [DataRow("${t|image:padding=1cm,fh=14cm}", 360000L, 360000L, 2880000L, 4320000L)]  // >
        [DataRow("${t|image:padding=1cm,fh=480px}", 360000L, 360000L, 2568000L, 3852000L)] // =
        [DataRow("${t|image:padding=1cm,fh=4in}", 360000L, 360000L, 1958400L, 2937600L)]   // <

        // 宽度 + 高度
        [DataRow("${t|image:padding=10px,fw=300px,fh=300px}", 95250L, 95250L, 1778000L, 2667000L)]   // 宽占满
        [DataRow("${t|image:padding=10px,fw=600px,fh=300px}", 95250L, 95250L, 1778000L, 2667000L)]   // 高占满
        [DataRow("${t|image:padding=10px,fw=820px,fh=1220px}", 95250L, 95250L, 7620000L, 11430000L)] // 等比放大
        [DataRow("${t|image:padding=0px,fw=160px,fh=240px}", 0L, 0L, 1524000L, 2286000L)]            // 等比缩小

        [TestMethod()]
        public void GetImageShapeTest(string input, long x, long y, long w, long h)
        {
            var grammar = new Grammar(input);
            var options = new ImageOptions(grammar.Options);
            var res = ImageUtils.GetImageShape(options, 320, 480);

            Assert.AreEqual(x, res.X);
            Assert.AreEqual(y, res.Y);
            Assert.AreEqual(w, res.W);
            Assert.AreEqual(h, res.H);
        }

        [TestMethod()]
        public void GetImageShapeTestFile()
        {
            var file = R.FullPath("data/image.jpg");
            var options = new ImageOptions();
            var res = ImageUtils.GetImageShape(options, file);
            Assert.IsNotNull(res);
        }
    }
}