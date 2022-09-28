using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using TemplateGOTests;

namespace TemplateGO.Tests
{
    [TestClass()]
    public class UtilsTests
    {
        private JsonElement Json { get { return JsonFromFile("data/utils.json"); } }

        private JsonElement JsonFromFile(string file)
        {
            var fullPath = R.FullPath(file);
            var jsonString = File.ReadAllText(fullPath);
            return JsonDocument.Parse(jsonString)!.RootElement;
        }

        [TestMethod()]
        public void GetValueTestL1()
        {
            Assert.AreEqual(Utils.GetValue(Json, "phone"), "18602507915");
            Assert.AreEqual(Utils.GetValue(Json, "height"), 48.5);
            Assert.AreEqual(Utils.GetValue(Json, "weight"), 3);

            Assert.AreEqual(Utils.GetValue(Json, "isNamed"), false);
            Assert.AreEqual(Utils.GetValue(Json, "hasPhone"), true);

            Assert.AreEqual(Utils.GetValue(Json, "school"), null);
            Assert.ThrowsException<KeyNotFoundException>(() => Utils.GetValue(Json, "lastUpdate"));

            // 其它情况应该返回 Object/Array
            var obj = Utils.GetValue(Json, "hospital") as JsonElement?;
            Assert.AreEqual(obj?.ValueKind, JsonValueKind.Object);

            var array = Utils.GetValue(Json, "parents") as JsonElement?;
            Assert.AreEqual(array?.ValueKind, JsonValueKind.Array);
        }

        [TestMethod()]
        public void GetValueTestL2()
        {
            Assert.AreEqual(Utils.GetValue(Json, "hospital.name"), "测试医院");
            Assert.ThrowsException<KeyNotFoundException>(() => Utils.GetValue(Json, "hospital.notName"));
        }

        [TestMethod()]
        public void GetValueTestL2Array()
        {
            Assert.AreEqual(Utils.GetValue(Json, "parents[0].name"), "父亲姓名");
            Assert.AreEqual(Utils.GetValue(Json, "parents[1].age"), 50);
            Assert.ThrowsException<KeyNotFoundException>(() => Utils.GetValue(Json, "hospital.notName"));
        }
    }
}