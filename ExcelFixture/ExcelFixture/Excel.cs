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

using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Excel;

namespace ExcelFixture
{
    [Documentation("Enabling testing Excel spreadsheets")]
    public class Excel
    {
        private static readonly Dictionary<int, string> CvErrors = new Dictionary<int, string>
        {
            {-2146826281, "#Div/0!"},
            {-2146826246, "#N/A"},
            {-2146826259, "#Name?"},
            {-2146826288, "#Null!"},
            {-2146826252, "#Num!"},
            {-2146826265, "#Ref!"},
            {-2146826273, "#Value!"}
        };

        private Workbook _currentWorkbook;
        private Application _excel;

        internal Worksheet CurrentWorksheet { get; private set; }

        private Application ExcelApplication => _excel ?? (_excel = new Application {Visible = false, DisplayAlerts = false});

        [Documentation("Get the address of the bottom right cell of the sheet")]
        public string LastCell => CurrentWorksheet.Cells.SpecialCells(XlCellType.xlCellTypeLastCell).Address;

        [Documentation("Check if the current workbook is protected")]
        public bool WorkbookIsProtected => _currentWorkbook.HasPassword;

        [Documentation("Check if the workbook is opened read-only")]
        public bool WorkbookIsReadOnly => _currentWorkbook.ReadOnly;

        [Documentation("Check if the current worksheet is protected (with or without password)")]
        public bool WorksheetIsProtected => CurrentWorksheet.ProtectContents;

        [Documentation("Check if the current worksheet is protected with a password")]
        public bool WorksheetIsProtectedWithPassword
        {
            get
            {
                if (!CurrentWorksheet.ProtectContents) return false;
                // somewhat tricky. We need to try and unprotect without a password to check if it has a password.
                // we take string.Empty as not using an argument shows a dialog, which we don't want.
                // If an error occurs, there is a password. 
                // If the un-protection succeeds there was none, and we simply re-protect without password. 
                var returnValue = false;
                try
                {
                    CurrentWorksheet.Unprotect(string.Empty);
                    CurrentWorksheet.Protect();
                }
                catch (COMException)
                {
                    returnValue = true;
                }
                return returnValue;
            }
        }

        private static bool IsError(object obj) => obj is int i && CvErrors.ContainsKey(i);

        [Documentation("Find a cell containing the specified text. Scope can be Partial or Whole.")]
        public object CellWithText(string scope, string dataToSearch)
        {
            var lookAt = scope.StartsWith("part", StringComparison.InvariantCultureIgnoreCase)
                ? XlLookAt.xlPart
                : XlLookAt.xlWhole;
            var cell = CurrentWorksheet.Cells.Find(dataToSearch, LookAt: lookAt, MatchCase: false);
            return cell?.Address;
        }

        [Documentation("Click a button")]
        public bool ClickButton(string name)
        {
            var buttons = CurrentWorksheet.Buttons();

            foreach (Button button in buttons)
            {
                if (button.Name != name && button.Text != name) continue;
                var procedure = button.OnAction;
                if (string.IsNullOrEmpty(procedure))
                {
                    return false;
                }
                ExcelApplication.Run(procedure);
                return true;
            }
            return false;
        }

        [Documentation("Close the Excel application (without saving)")]
        public void CloseExcel()
        {
            if (_excel == null) return;
            foreach (Workbook workbook in _excel.Workbooks)
            {
                workbook.Close(false, Type.Missing, Type.Missing);
            }
            _excel.Quit();
        }

        [Documentation("Close the indicated workbook")]
        public bool CloseWorkbook(string workbookPath)
        {
            var workbook = FindWorkbook(workbookPath);
            if (workbook == null) return false;
            if (_currentWorkbook == workbook)
            {
                _currentWorkbook = null;
                CurrentWorksheet = null;
            }
            workbook.Close(false, Type.Missing, Type.Missing);
            return FindWorkbook(workbookPath) == null;
        }

        [Documentation("Close the current workbook")]
        public bool CloseWorkbook() => _currentWorkbook != null && CloseWorkbook(_currentWorkbook.FullName);

        [Documentation("Execute a macro, function or expression")]
        public object Execute(string expression)
        {
            object returnValue1 = _excel.Evaluate(expression);
            if (!IsError(returnValue1)) return returnValue1;
            var message = "Error executing [" + expression + "]: " + CvErrors[Convert.ToInt32(returnValue1)];
            throw new EvaluateException(message);
        }

        private Workbook FindWorkbook(string workbookPath)
        {
            var fullPath = Path.GetFullPath(workbookPath);
            return ExcelApplication.Workbooks.Cast<Workbook>()
                .FirstOrDefault(
                    workbook => workbook.FullName.Equals(fullPath, StringComparison.InvariantCultureIgnoreCase));
        }

        [Documentation("Get the formula of a certain cell")]
        public object FormulaOfCell(string cellLocation)
        {
            var cell = CurrentWorksheet.Range[cellLocation];
            return cell.Formula;
        }

        private bool LoadWorkbook(string path, bool readOnly, string password)
        {
            var fullPath = Path.GetFullPath(path);
            _currentWorkbook = ExcelApplication.Workbooks.Open(fullPath, ReadOnly: readOnly, Password: password,
                IgnoreReadOnlyRecommended: !readOnly);
            CurrentWorksheet = _currentWorkbook.ActiveSheet;
            return _currentWorkbook != null;
        }

        [Documentation("Load a workbook")]
        public bool LoadWorkbook(string path) => LoadWorkbook(path, false, null);

        [Documentation("Load a workbook in read-only mode (bypass read/write password dialog)")]
        public bool LoadWorkbookReadOnly(string path) => LoadWorkbook(path, true, null);

        [Documentation("Load a workbook in read only mode, and provide the password")]
        public bool LoadWorkbookReadOnlyWithPassword(string path, string password) => LoadWorkbook(path, true, password);

        [Documentation("Load a workbook, providing a password")]
        public bool LoadWorkbookWithPassword(string path, string password) => LoadWorkbook(path, false, password);

        [Documentation("Get the address of the range that is offset a number of columns and rows from the input range")]
        public object OffsetByRowsAndColumns(object cellLocation, object rows, object cols) =>
            CurrentWorksheet.Range[cellLocation].Offset[rows, cols]?.Address;

        internal object Properties(object cellLocation)
        {
            var retVal = string.Empty;
            var cell = CurrentWorksheet.Range[cellLocation];
            var cellType = typeof(Range);
            foreach (var prop in cellType.GetProperties())
            {
                retVal += prop.Name + "=";
                if (prop.GetIndexParameters().Length == 0)
                {
                    retVal += prop.GetValue(cell) + "; ";
                }
                else
                {
                    retVal += prop.GetIndexParameters().Length + " params;";
                    if (prop.Name == "Address")
                    {
                        retVal += prop.GetValue(cell, new object[] {null, null, null, null, null}) + "; ";
                    }
                }
            }
            return retVal;
        }

        [Documentation("Protect the current workbook")]
        public bool ProtectWorkbookWithPassword(string password)
        {
            if (WorkbookIsProtected) return false;

            // Protect doesn't automatically set Password. So do that separately.
            // We need this since WorkbookIsProtected relies on HasPassword.
            _currentWorkbook.Protect(password);
            _currentWorkbook.Password = password;
            return true;
        }

        [Documentation("Protect the current worksheet")]
        public bool ProtectWorksheetWithPassword(string password)
        {
            // Protect ignores re-protection so let's not allow that.
            if (WorksheetIsProtected) return false;

            // if we have an empty password, protect without password.
            CurrentWorksheet.Protect(string.IsNullOrEmpty(password) ? Type.Missing : password);

            // we are successful if the contents are now protected
            return CurrentWorksheet.ProtectContents;
        }

        // don't announce these for now. Testing shouldn't require saving Excel sheets.
        internal void SaveWorkbook() => SaveWorkbookAsWithPassword(_currentWorkbook.FullName, string.Empty);

        internal void SaveWorkbookAs(string path) => SaveWorkbookAsWithPassword(path, string.Empty);

        internal void SaveWorkbookAsWithPassword(string path, string password) => _currentWorkbook.SaveAs(path, Password: password);

        internal void SaveWorkbookWithPassword(string password) => SaveWorkbookAsWithPassword(_currentWorkbook.FullName, password);

        [Documentation("Switch to an already open workbook")]
        public bool SelectWorkbook(string workbookPath)
        {
            var workbook = FindWorkbook(workbookPath);
            if (workbook == null) return false;
            _currentWorkbook = workbook;
            CurrentWorksheet = _currentWorkbook.ActiveSheet;
            return true;
        }

        [Documentation("Switch to a worksheet of the current workbook")]
        public bool SelectWorksheet(string sheetName)
        {
            if (_currentWorkbook == null) return false;
            CurrentWorksheet = int.TryParse(sheetName, out var sheetNumber)
                ? _currentWorkbook.Sheets[sheetNumber]
                : _currentWorkbook.Sheets[sheetName];
            return CurrentWorksheet != null;
        }

        [Documentation("Set the value of a certain cell")]
        public bool SetValueOfCellTo(string cellLocation, object value)
        {
            try
            {
                var cell = CurrentWorksheet.Range[cellLocation];
                cell.Value2 = value;
                return cell.Value2.ToString().Equals(value.ToString());
            }
            catch (COMException)
            {
                return false;
            }
        }

        [Documentation("Get the text in a certain cell (displayed text, not necessarily actual value)")]
        public object TextOfCell(string cellLocation)
        {
            var cell = CurrentWorksheet.Range[cellLocation];
            return cell.Text;
        }

        [Documentation("Unprotect the current workbook")]
        public bool UnprotectWorkbookWithPassword(string password)
        {
            // setting the password property only works automatically when loading a sheet.
            // So we need to clear it ourselves after a successful unprotect.
            // We need that since WorkbookIsProtected relies on HasPassword.
            try
            {
                _currentWorkbook.Unprotect(password);
            }
            catch (COMException)
            {
                return false;
            }
            _currentWorkbook.Password = string.Empty;
            return true;
        }

        [Documentation("Unprotect the current worksheet")]
        public bool UnprotectWorksheetWithPassword(string password)
        {
            const int passwordError = -2146827284;
            try
            {
                // Don't unprotect without password as that may cause an (invisible) dialog.
                // if the sheet is protected without password, the argument is ignored anyway.
                CurrentWorksheet.Unprotect(password);
            }
            catch (COMException ce)
            {
                if (ce.ErrorCode == passwordError) return false;
                throw;
            }
            return true;
        }

        [Documentation("Get the value of a certain cell")]
        public object ValueOfCell(string cellLocation)
        {
            var cell = CurrentWorksheet.Range[cellLocation];
            return cell.Value2;
        }
    }
}