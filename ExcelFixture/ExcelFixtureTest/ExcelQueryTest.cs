// Copyright 2015-2020 Rik Essenius
//
//   Licensed under the Apache License, Version 2.0 (the "License"); you may not use this file 
//   except in compliance with the License. You may obtain a copy of the License at
//
//       http://www.apache.org/licenses/LICENSE-2.0
//
//   Unless required by applicable law or agreed to in writing, software distributed under the License 
//   is distributed on an "AS IS" BASIS WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
//   See the License for the specific language governing permissions and limitations under the License.

using System.Collections.ObjectModel;
using System.Diagnostics.CodeAnalysis;
using ExcelFixture;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace ExcelFixtureTest
{
    [TestClass]
    public class ExcelQueryTest
    {
        private static Excel _excel;

        [ClassCleanup]
        public static void ClassCleanup()
        {
            _excel.CloseExcel();
        }

        [ClassInitialize, DeploymentItem("ExcelFixtureTest\\ExcelFixtureTest.xlsm"),
         SuppressMessage("Usage", "CA1801:Review unused parameters", Justification = "False positive"),
         SuppressMessage("Style", "IDE0060:Remove unused parameter", Justification = "False positive")]
        public static void ClassInitialize(TestContext testContext)
        {
            _excel = new Excel();
            Assert.IsTrue(_excel.LoadWorkbookReadOnly("ExcelFixtureTest.xlsm"));
            Assert.IsTrue(_excel.SelectWorksheet("Sheet1"), "Select worksheet 1");
        }

        [TestMethod]
        public void ExcelQueryWithHeaderTest()
        {
            TestTable("C4:D8", "UseHeaders", new[]
            {
                new[] {new object[] {"Input", 0.0}, new object[] {"Fibonacci", 0.0}},
                new[] {new object[] {"Input", 1.0}, new object[] {"Fibonacci", 1.0}},
                new[] {new object[] {"Input", 2.0}, new object[] {"Fibonacci", 1.0}},
                new[] {new object[] {"Input", 3.0}, new object[] {"Fibonacci", 2.0}}
            });
        }

        [TestMethod]
        public void ExcelQueryWithoutHeaderTest()
        {
            TestTable("C4:D8", string.Empty, new[]
            {
                new[] {new object[] {"Column 3", "Input"}, new object[] {"Column 4", "Fibonacci"}},
                new[] {new object[] {"Column 3", 0.0}, new object[] {"Column 4", 0.0}},
                new[] {new object[] {"Column 3", 1.0}, new object[] {"Column 4", 1.0}},
                new[] {new object[] {"Column 3", 2.0}, new object[] {"Column 4", 1.0}},
                new[] {new object[] {"Column 3", 3.0}, new object[] {"Column 4", 2.0}}
            });
        }

        private static void TestTable(string range, string options, object[][][] expectedValues)
        {
            var eq = string.IsNullOrEmpty(options) ? new ExcelQuery(_excel, range) : new ExcelQuery(_excel, range, options);
            var table = eq.Query();
            Assert.IsNotNull(table);
            Assert.AreEqual(expectedValues.Length, table.Count, "Row Count");
            for (var row = 0; row < table.Count; row++)
            {
                var rowCollection = table[row] as Collection<object>;
                Assert.IsNotNull(rowCollection);
                Assert.AreEqual(expectedValues[row].Length, rowCollection.Count, "Column Count");
                for (var column = 0; column < rowCollection.Count; column++)
                {
                    var columnCollection = rowCollection[column] as Collection<object>;
                    Assert.IsNotNull(columnCollection);
                    Assert.AreEqual(2, expectedValues[row][column].Length, "Cell Count");
                    Assert.AreEqual(expectedValues[row][column][0], columnCollection[0], "{0}/{1}({2},{3},{4})", range, options, row, column, 0);
                    Assert.AreEqual(expectedValues[row][column][1], columnCollection[1], "{0}/{1}({2},{3},{4})", range, options, row, column, 1);
                }
            }
        }
    }
}
