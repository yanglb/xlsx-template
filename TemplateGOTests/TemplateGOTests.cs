using Microsoft.VisualStudio.TestTools.UnitTesting;
using TemplateGOTests;

namespace TemplateGO.Tests
{
    [TestClass()]
    public class TemplateGOTests
    {
        [TestMethod()]
        public void RenderTestXltx()
        {
            TemplateGO.Render(R.FullPath("data/test.xlsx"), R.JsonFromFile("data/test.json"), "out.xlsx");
            Assert.Fail();
        }
    }
}