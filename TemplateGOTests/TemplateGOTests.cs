using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.IO;
using System.Text.Json;
using TemplateGOTests;

namespace TemplateGO.Tests
{
    [TestClass()]
    public class TemplateGOTests
    {
        private JsonElement JsonFromFile(string file)
        {
            var fullPath = R.FullPath(file);
            var jsonString = File.ReadAllText(fullPath);
            return JsonDocument.Parse(jsonString)!.RootElement;
        }

        [TestMethod()]
        public void RenderTestXltx()
        {
            TemplateGO.Render(R.FullPath("data/test.xlsx"), JsonFromFile("data/test.json"), "out.xlsx");
            Assert.Fail();
        }
    }
}