using BioMed.Parsers;
using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace BioMed.Parsers.Tests
{
    [TestClass()]
    public class ExcelParserTests
    {
        [TestMethod()]
        public void LoadFile()
        {
            var excelparserclient = new BioMed.Parsers.ExcelParser(@"C:\BioMed\Patient Monitors - Examples. V1.1.xlsx");
            var output = excelparserclient.toXML();
            if (string.IsNullOrEmpty(output))
                Assert.Fail();
        }

        [TestMethod()]
        public void TabToListTest()
        {
            var excelparserclient = new BioMed.Parsers.ExcelParser(@"C:\BioMed\Patient Monitors - Examples. V1.1.xlsx");
            var output = excelparserclient.TabToList("ECG");
            if (output == null)
                Assert.Fail();
        }

        [TestMethod()]
        public void TabsToListTest()
        {
            var excelparserclient = new BioMed.Parsers.ExcelParser(@"C:\BioMed\Patient Monitors - Examples. V1.1.xlsx");
            var output = excelparserclient.TabsToList();
            if (output == null)
                Assert.Fail();
        }
    }
}
