// Copyright 2015-2019 Rik Essenius
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

        [ClassInitialize, DeploymentItem("ExcelFixtureTest\\ExcelFixtureTest.xlsm")]
        public static void ClassInitialize(TestContext testContext)
        {
            _excel = new Excel();
            Assert.IsTrue(_excel.LoadWorkbookReadOnly("ExcelFixtureTest.xlsm"));
            Assert.IsTrue(_excel.SelectWorksheet("Sheet1"), "Select worksheet 1");
        }

        private static void TestTable(string range, string options, object[,,] expectedValues)
        {
            var eq = string.IsNullOrEmpty(options)
                ? new ExcelQuery(_excel, range)
                : new ExcelQuery(_excel, range, options);
            var table = eq.Query();
            Assert.IsNotNull(table);
            Assert.AreEqual(expectedValues.GetLength(0), table.Count, "Row Count");
            for (var row = 0; row < table.Count; row++)
            {
                var rowCollection = table[row] as Collection<object>;
                Assert.IsNotNull(rowCollection);
                Assert.AreEqual(expectedValues.GetLength(1), rowCollection.Count, "Column Count");
                for (var column = 0; column < rowCollection.Count; column++)
                {
                    var columnCollection = rowCollection[column] as Collection<object>;
                    Assert.IsNotNull(columnCollection);
                    Assert.AreEqual(2, expectedValues.GetLength(2), "Cell Count");
                    Assert.AreEqual(expectedValues[row, column, 0], columnCollection[0], "{0}/{1}({2},{3},{4})", range,
                        options, row, column, 0);
                    Assert.AreEqual(expectedValues[row, column, 1], columnCollection[1], "{0}/{1}({2},{3},{4})", range,
                        options, row, column, 1);
                }
            }
        }

        [TestMethod]
        public void ExcelQueryWithHeaderTest()
        {
            TestTable("C4:D8", "UseHeaders", new object[,,]
            {
                {{"Input", 0.0}, {"Fibonacci", 0.0}},
                {{"Input", 1.0}, {"Fibonacci", 1.0}},
                {{"Input", 2.0}, {"Fibonacci", 1.0}},
                {{"Input", 3.0}, {"Fibonacci", 2.0}}
            });
        }

        [TestMethod]
        public void ExcelQueryWithoutHeaderTest()
        {
            TestTable("C4:D8", string.Empty, new object[,,]
            {
                {{"Column 3", "Input"}, {"Column 4", "Fibonacci"}},
                {{"Column 3", 0.0}, {"Column 4", 0.0}},
                {{"Column 3", 1.0}, {"Column 4", 1.0}},
                {{"Column 3", 2.0}, {"Column 4", 1.0}},
                {{"Column 3", 3.0}, {"Column 4", 2.0}}
            });
        }
    }
}