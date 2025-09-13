using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;

namespace XlsxTemplate.Utils.Tests
{
    [TestClass()]
    public class ParserUtilsTests
    {
        [DataRow("1cm", "1", "cm")]
        [DataRow("1.25kg", "1.25", "kg")]
        [DataRow("-35km/h", "-35", "km/h")]
        [DataRow("17.56", "17.56", "")]
        [TestMethod()]
        public void ValueWithUnitTest(string input, string value, string unit)
        {
            string actualValue;
            string actualUnit;
            ParserUtils.ValueWithUnit(input, out actualValue, out actualUnit);

            Assert.AreEqual(value, actualValue);
            Assert.AreEqual(unit, actualUnit);
        }

        [DataRow("1", "cm", 360000)]
        [DataRow("1", "in", 914400)]
        [DataRow("1", "px", 9525)]

        [TestMethod()]
        public void ValueToEmuTest(string value, string unit, long result)
        {
            var actual = ParserUtils.ValueToEmu(value, unit);
            Assert.AreEqual(result, actual);
        }

        [DataRow("1", "cm+")]
        [DataRow("1", "in+")]
        [DataRow("1", "")]
        [DataRow("1", "kg")]

        [TestMethod()]
        public void ValueToEmuTestBad(string value, string unit)
        {
            Assert.ThrowsException<ArgumentException>(() => ParserUtils.ValueToEmu(value, unit));
        }

        [DataRow("1cm", 360000)]
        [DataRow("1in", 914400)]
        [DataRow("1px", 9525)]
        [TestMethod()]
        public void ParserValueToEmuTest(string input, long result)
        {
            var actual = ParserUtils.ParseValueToEmu(input);
            Assert.AreEqual(result, actual);
        }
    }
}