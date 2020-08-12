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

using System;
using System.Data;
using System.Diagnostics;
using System.Diagnostics.CodeAnalysis;
using System.IO;
using System.Runtime.InteropServices;
using ExcelFixture;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace ExcelFixtureTest
{
    [TestClass]
    public class ExcelTest
    {
        private static Excel _excel;

        private static void CheckCosts(double inputHours, double monthlyCosts, double yearlyCosts)
        {
            Assert.IsTrue(_excel.SetValueOfCellTo("BuildTime", inputHours), "Build Time set");
            Assert.AreEqual(monthlyCosts, _excel.ValueOfCell("CostsPerMonth"), "Monthly costs OK");
            Assert.IsTrue(_excel.ClickButton("Button 1"), "Click Button 1");
            Assert.AreEqual(yearlyCosts, _excel.ValueOfCell("CostsPerYear"), "Yearly costs OK");
        }

        [ClassCleanup]
        public static void ClassCleanup() => _excel.CloseExcel();

        [SuppressMessage("Usage", "CA1801:Review unused parameters", Justification = "False positive"),
         SuppressMessage("Style", "IDE0060:Remove unused parameter", Justification = "False positive"), ClassInitialize,
         DeploymentItem("ExcelFixtureTest\\ExcelFixtureTest.xlsm"), DeploymentItem("ExcelFixtureTest\\TestSheet.xlsx")]
        public static void ClassInitialize(TestContext testContext)
        {
            _excel = new Excel();
            Assert.IsTrue(_excel.LoadWorkbookReadOnly(@"ExcelFixtureTest.xlsm"));
            Assert.IsTrue(_excel.SelectWorksheet("Sheet2"), "Select worksheet");
        }

        [TestMethod]
        public void ExcelButtonTest()
        {
            Assert.IsFalse(_excel.ClickButton("Button 2"), "Press Button 2 fails");
            Assert.IsFalse(_excel.ClickButton(@"Nonexisting"), "Press non-existing button fails");
        }

        [TestMethod]
        public void ExcelCellTest()
        {
            var cell = _excel.LastCell;
            Assert.AreEqual(347, Math.Round(Convert.ToDouble(_excel.ValueOfCell(cell)), 0), "Value of Cell OK");
            Debug.Print(_excel.ValueOfCell(cell).ToString());
            Debug.Print(_excel.FormulaOfCell("CostsPerMonth").ToString());
        }

        [TestMethod]
        public void ExcelCellWithTextTest()
        {
            Assert.AreEqual("$A$7", _excel.CellWithText("", "Costs per year"));
            Assert.AreEqual(null, _excel.CellWithText("", @"notpresent"));
            Assert.AreEqual("$A$2", _excel.CellWithText("partial", "Cost"));
            Assert.AreEqual("$A$3", _excel.CellWithText("partial", "MAX"));
        }

        [TestMethod]
        public void ExcelCheckCostsTest()
        {
            CheckCosts(0, 0, 0);
            CheckCosts(1, 0, 0);
            CheckCosts(2, 3, 36);
            CheckCosts(20, 57, 684);
            CheckCosts(100, 105, 1260);
            CheckCosts(1000, 645, 7740);
            CheckCosts(5110, 3111, 37332);
        }

        [TestMethod]
        public void ExcelExecuteExpectExceptionTest()
        {
            try
            {
                _excel.Execute("0/0");
                Assert.Fail("No exception thrown");
            }
            catch (EvaluateException e)
            {
                Assert.AreEqual("Error executing [0/0]: #Div/0!", e.Message);
            }
        }

        [TestMethod]
        public void ExcelExecuteTest()
        {
            Assert.AreEqual(32D, _excel.Execute("2^5"), "Value evaluation works");
            Assert.AreEqual(34, _excel.Execute("Add(21, 13)"), "Function call works");
            Assert.AreEqual(32767, _excel.Execute("Add(32766, 1)"), "return int until max int16");
            Assert.AreEqual(32768D, _excel.Execute("Add(32767, 1)"), "switch to double after max int16");
            Assert.AreEqual(-2146826252D, _excel.Execute("-2146826252"), "CVError values do not erroneously raise an exception");

            Assert.IsTrue(_excel.SetValueOfCellTo("BuildTime", "50"), "Set build time value to 50");
            _excel.Execute("UpdateCostsPerYear()");
            Assert.AreEqual(900D, _excel.ValueOfCell("CostsPerYear"), "CostsPerYear correct. so sub executed");
            Assert.IsTrue(_excel.SetValueOfCellTo("BuildTime", "110"), "Set value to 110");
            _excel.Execute("UpdateCostsPerYear()");
            Assert.AreEqual(1332D, _excel.ValueOfCell("CostsPerYear"), "CostsPerYear correct again");
        }

        [TestMethod]
        public void ExcelFormulaOfCellTeast() => Assert.AreEqual("=B11/60", _excel.FormulaOfCell("$B$12"));

        [TestMethod]
        public void ExcelOffsetTest() => Assert.AreEqual("$B$12", _excel.OffsetByRowsAndColumns("$A$9", 3, 1));

        [TestMethod]
        public void ExcelOpenCloseWorkbookTest()
        {
            Assert.IsTrue(_excel.LoadWorkbook("TestSheet.xlsx"), "Open second sheet");
            Assert.AreEqual("Test sheet", _excel.ValueOfCell("A1"), "Right sheet is active on second workbook");
            Assert.IsTrue(_excel.SelectWorkbook("ExcelFixtureTest.xlsm"), "Switch to first sheet");
            Assert.AreEqual("Free hours", _excel.ValueOfCell("A1"), "Right sheet is active on first workbook");
            Assert.IsTrue(_excel.CloseWorkbook("TestSheet.xlsx"), "Close inactive sheet");
            Assert.IsFalse(_excel.CloseWorkbook("TestSheet.xlsx"), "Close inactive sheet again");
            Assert.AreEqual("Free hours", _excel.ValueOfCell("A1"), "First workbook is still current");
            Assert.IsFalse(_excel.SelectWorkbook("TestSheet.xlsx"));
            Assert.IsTrue(_excel.LoadWorkbook("TestSheet.xlsx"), "Open second sheet again");
            Assert.IsTrue(_excel.CloseWorkbook(), "Immediately close the current workbook");
            Assert.IsFalse(_excel.CloseWorkbook(), "Immediately close the current workbook again");
            Assert.IsFalse(_excel.SelectWorksheet("Sheet1"), "No current workbook, so can't select a sheet");
            Assert.IsTrue(_excel.SelectWorkbook("ExcelFixtureTest.xlsm"), "Switch to first sheet");
            Assert.AreEqual("Free hours", _excel.ValueOfCell("A1"), "Right sheet is active on first workbook");
        }

        [TestMethod]
        public void ExcelPropertiesTest() => Console.WriteLine(_excel.Properties("CostsPerMonth"));

        [TestMethod]
        public void ExcelTextOfCellTest() => Assert.AreEqual("Free hours", _excel.TextOfCell("A1"));

        [TestMethod]
        public void ExcelWorkbookProtectionTest()
        {
            Assert.IsFalse(_excel.WorkbookIsProtected, "Workbook is not protected");
            Assert.IsTrue(_excel.UnprotectWorkbookWithPassword(@"fout"), "unprotection is ignored for unprotected workbooks");
            Assert.IsTrue(_excel.ProtectWorkbookWithPassword("secret"), "protect unprotected sheet");
            Assert.IsTrue(_excel.WorkbookIsProtected, "Workbook protected after Protect");
            Assert.IsFalse(_excel.ProtectWorkbookWithPassword("secret"), "protect already protected sheet fails");
            Assert.IsFalse(_excel.UnprotectWorkbookWithPassword(@"fout"), "Unprotecting with wrong password fails");
            Assert.IsTrue(_excel.UnprotectWorkbookWithPassword("secret"), "unprotect protected sheet");
            Assert.IsFalse(_excel.WorkbookIsProtected, "Workbook unprotected after Unprotect");
            var fileName = Path.GetTempFileName();
            _excel.SaveWorkbookAsWithPassword(fileName, "secret");
            Assert.IsTrue(_excel.CloseWorkbook(), "Close initial workbook");
            try
            {
                _excel.LoadWorkbookWithPassword(fileName, @"fout");
                Assert.Fail("No exception thrown with wrong password");
            }
            catch (COMException)
            {
                // expected; continue
            }
            Assert.IsTrue(_excel.LoadWorkbookWithPassword(fileName, "secret"), "Load protected workbook");
            Assert.IsFalse(_excel.WorkbookIsReadOnly, "Workbook is not read-only");
            Assert.IsTrue(_excel.WorkbookIsProtected, "Loaded workbook is protected");
            Assert.IsTrue(_excel.UnprotectWorkbookWithPassword("secret"), "Unprotect loaded workbook");
            Assert.IsFalse(_excel.WorkbookIsProtected, "Loaded workbook is not protected afterwards");
            var fileName2 = Path.GetTempFileName();
            _excel.SaveWorkbook(); // without password so should now be unprotected when reloaded
            _excel.SaveWorkbookAs(fileName2); // also without password
            Assert.IsTrue(_excel.CloseWorkbook(), "Close protected workbook");
            Assert.IsTrue(_excel.LoadWorkbook(fileName), "Load deprotected workbook");
            Assert.IsFalse(_excel.WorkbookIsReadOnly, "Deprotected Workbook is not read-only");
            Assert.IsFalse(_excel.WorkbookIsProtected, "Deprotected workbook is not protected");
            _excel.SaveWorkbookWithPassword("secret2");
            Assert.IsTrue(_excel.CloseWorkbook(), "Close reprotected workbook");
            Assert.IsTrue(_excel.LoadWorkbookReadOnlyWithPassword(fileName, "secret2"), "Load reprotected workbook readonly");
            Assert.IsTrue(_excel.WorkbookIsReadOnly, "Workbook is read-only");
            Assert.IsTrue(_excel.CloseWorkbook(), "Close reprotected readonly workbook");
            Assert.IsTrue(_excel.LoadWorkbookWithPassword(fileName, "secret2"), "Load reprotected workbook");
            Assert.IsFalse(_excel.WorkbookIsReadOnly, "Workbook is not read-only");
            Assert.IsTrue(_excel.CloseWorkbook(), "Close reprotected workbook");
            File.Delete(fileName);
            Assert.IsTrue(_excel.LoadWorkbookReadOnly(fileName2), "Load unprotected workbook readonly");
            Assert.IsFalse(_excel.WorkbookIsProtected, "Loaded unprotected Workbook is not protected");
            Assert.IsTrue(_excel.ProtectWorkbookWithPassword("secret"), "Protect loaded unprotected Workbook");
            Assert.IsTrue(_excel.WorkbookIsProtected, "Loaded unprotected Workbook is protected after Protect");
            Assert.IsTrue(_excel.CloseWorkbook(), "Close unprotected workbook");
            File.Delete(fileName2);

            // bring the original state back for the next test
            Assert.IsTrue(_excel.LoadWorkbookReadOnly("ExcelFixtureTest.xlsm"));
            Assert.IsTrue(_excel.SelectWorksheet("Sheet2"), "Select worksheet");
        }

        [TestMethod]
        public void ExcelWorksheetProtectedTest()
        {
            var cellA1 = _excel.ValueOfCell("A1");

            Assert.IsFalse(_excel.WorksheetIsProtected, "Worksheet is not protected before");
            Assert.IsTrue(_excel.ProtectWorksheetWithPassword("secret"), "Protect with password succeeds");
            Assert.IsTrue(_excel.WorksheetIsProtectedWithPassword, "Worksheet is protected with password");
            Assert.IsFalse(_excel.SetValueOfCellTo("A1", "1"), "Can't change protected sheet (1)");
            Assert.IsFalse(_excel.UnprotectWorksheetWithPassword(string.Empty), "Unprotecting without password fails");
            Assert.IsTrue(_excel.WorksheetIsProtected, "Worksheet is still protected after failed unprotect");
            Assert.IsFalse(_excel.SetValueOfCellTo("A1", "2"), "Can't change protected sheet (2)");
            Assert.IsFalse(_excel.UnprotectWorksheetWithPassword("wrong"), "Unprotecting with wrong password fails");
            Assert.IsTrue(_excel.WorksheetIsProtected, "Worksheet is still protected after using wrong password");
            Assert.IsTrue(_excel.UnprotectWorksheetWithPassword("secret"), "Unprotecting with password succeeds");
            Assert.IsFalse(_excel.WorksheetIsProtected, "Worksheet is not protected after unprotecting");
            Assert.IsTrue(_excel.SetValueOfCellTo("A1", "3"), "Can change unprotected sheet (3)");
            Assert.IsTrue(_excel.ProtectWorksheetWithPassword(string.Empty), "Protecting with empty password succeeds");
            Assert.IsTrue(_excel.WorksheetIsProtected, "Worksheet is protected with empty password");
            Assert.IsFalse(_excel.WorksheetIsProtectedWithPassword, "Worksheet is not protected with password");
            Assert.IsFalse(_excel.SetValueOfCellTo("A1", "4"), "Can't change protected sheet (4)");
            Assert.IsFalse(_excel.ProtectWorksheetWithPassword("test"), "Re-protecting fails");
            Assert.IsFalse(_excel.WorksheetIsProtectedWithPassword, "Worksheet is protected without password after reprotection");
            Assert.IsTrue(_excel.UnprotectWorksheetWithPassword(string.Empty), "Unprotection without password succeeds");
            Assert.IsFalse(_excel.WorksheetIsProtected, "Worksheet is not protected afterwards");
            Assert.IsTrue(_excel.SetValueOfCellTo("A1", cellA1), "Can change A1 of unprotected sheet to original value");
        }
    }
}
