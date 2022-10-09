using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Collections.Generic;
using System.Text.Json;
using TemplateGO.Utils;
using TemplateGOTests;

namespace TemplateGO.Tests
{
    [TestClass()]
    public class JsonUtilsTest
    {
        private JsonElement Json { get { return R.JsonFromFile("data/utils.json"); } }

        [TestMethod()]
        public void GetValueTestL1()
        {
            Assert.AreEqual(JsonUtils.GetValue(Json, "phone"), "18602507915");
            Assert.AreEqual(JsonUtils.GetValue(Json, "height"), 48.5);
            Assert.AreEqual(JsonUtils.GetValue(Json, "weight"), 3);

            Assert.AreEqual(JsonUtils.GetValue(Json, "isNamed"), false);
            Assert.AreEqual(JsonUtils.GetValue(Json, "hasPhone"), true);

            Assert.AreEqual(JsonUtils.GetValue(Json, "school"), null);
            Assert.ThrowsException<KeyNotFoundException>(() => JsonUtils.GetValue(Json, "lastUpdate"));

            // 其它情况应该返回 Object/Array
            var obj = JsonUtils.GetValue(Json, "hospital") as JsonElement?;
            Assert.AreEqual(obj?.ValueKind, JsonValueKind.Object);

            var array = JsonUtils.GetValue(Json, "parents") as JsonElement?;
            Assert.AreEqual(array?.ValueKind, JsonValueKind.Array);
        }

        [TestMethod()]
        public void GetValueFromArrayTest()
        {
            // 数组对象
            var arr1 = Json.GetProperty("parents");
            Assert.AreEqual("父亲姓名", JsonUtils.GetValue(arr1, "[0].name"));
            Assert.AreEqual(50, JsonUtils.GetValue(arr1, "[1].age"));

            Assert.AreEqual("父亲姓名", JsonUtils.GetValue(Json, "parents[0].name"));
            Assert.AreEqual(50, JsonUtils.GetValue(Json, "parents[1].age"));

            // 经数组元素
            var arr2 = Json.GetProperty("favorites");
            Assert.AreEqual("music", JsonUtils.GetValue(arr2, "[0]"));
            Assert.AreEqual("move", JsonUtils.GetValue(arr2, "[1]"));

            Assert.AreEqual("music", JsonUtils.GetValue(Json, "favorites[0]"));
            Assert.AreEqual("move", JsonUtils.GetValue(Json, "favorites[1]"));
        }

        [TestMethod()]
        public void GetValueTestL2()
        {
            Assert.AreEqual(JsonUtils.GetValue(Json, "hospital.name"), "测试医院");
            Assert.ThrowsException<KeyNotFoundException>(() => JsonUtils.GetValue(Json, "hospital.notName"));
        }

        [TestMethod()]
        public void GetValueTestL2Array()
        {
            Assert.AreEqual(JsonUtils.GetValue(Json, "parents[0].name"), "父亲姓名");
            Assert.AreEqual(JsonUtils.GetValue(Json, "parents[1].age"), 50);
            Assert.ThrowsException<KeyNotFoundException>(() => JsonUtils.GetValue(Json, "hospital.notName"));
        }
    }
}