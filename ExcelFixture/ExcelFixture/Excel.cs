// Copyright 2015-2019 Rik Essenius
// Licensed under the Apache License, Version 2.0 (the "License");
// you may not use this file except in compliance with the License. You may obtain a copy of the License at
//
//   http://www.apache.org/licenses/LICENSE-2.0
//
// Unless required by applicable law or agreed to in writing, software distributed under the License is 
// distributed on an "AS IS" BASIS, WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
// See the License for the specific language governing permissions and limitations under the License.

using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics.CodeAnalysis;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Excel;

namespace ExcelFixture
{
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

        [SuppressMessage("ReSharper", "UnusedMember.Global", Justification = "Used by FixtureExplorer")]
        public static Dictionary<string, string> FixtureDocumentation { get; } = new Dictionary<string, string>
        {
            {string.Empty, "Enabling testing Excel spreadsheets"},
            {nameof(CellWithText), "Find a cell containing the specified text. Scope can be Partial or Whole."},
            {nameof(ClickButton), "Click a button"},
            {nameof(CloseExcel), "Close the Excel application (without saving)"},
            {nameof(CloseWorkbook), "Close the current workbook"},
            {nameof(CloseWorkbook) + "`1", "Close the indicated workbook"},
            {nameof(Execute), "Execute a macro, function or expression"},
            {nameof(FormulaOfCell), "Get the formula of a certain cell"},
            {nameof(LastCell), "Get the address of the bottom right cell of the sheet"},
            {nameof(LoadWorkbook), "Load a workbook"},
            {nameof(LoadWorkbookReadOnly), "Load a workbook in read-only mode (bypass read/write password dialog)"},
            {nameof(LoadWorkbookReadOnlyWithPassword), "Load a workbook in read only mode, and provide the password"},
            {nameof(LoadWorkbookWithPassword), "Load a workbook, providing a password"},
            {
                nameof(OffsetByRowsAndColumns),
                "Get the address of the range that is offset a number of columns and rows from the input range"
            },
            {nameof(ProtectWorkbookWithPassword), "Protect the current workbook"},
            {nameof(ProtectWorksheetWithPassword), "Protect the current worksheet"},
            {nameof(SelectWorkbook), "Switch to an already open workbook"},
            {nameof(SelectWorksheet), "Switch to a worksheet of the current workbook"},
            {nameof(SetValueOfCellTo), "Set the value of a certain cell"},
            {nameof(TextOfCell), "Get the text in a certain cell (displayed text, not necessarily actual value)"},
            {nameof(UnprotectWorkbookWithPassword), "Unprotect the current workbook"},
            {nameof(UnprotectWorksheetWithPassword), "Unprotect the current worksheet"},
            {nameof(ValueOfCell), "Get the value of a certain cell"},
            {nameof(WorkbookIsProtected), "Check if the current workbook is protected"},
            {nameof(WorkbookIsReadOnly), "Check if the workbook is opened read-only"},
            {nameof(WorksheetIsProtected), "Check if the current worksheet is protected (with or without password)"},
            {nameof(WorksheetIsProtectedWithPassword), "Check if the current worksheet is protected with a password"}
        };

        internal Worksheet CurrentWorksheet { get; private set; }

        private Application ExcelApplication => _excel ?? (_excel = new Application {Visible = false, DisplayAlerts = false});

        public string LastCell => CurrentWorksheet.Cells.SpecialCells(XlCellType.xlCellTypeLastCell).Address;

        public bool WorkbookIsProtected => _currentWorkbook.HasPassword;

        public bool WorkbookIsReadOnly => _currentWorkbook.ReadOnly;

        public bool WorksheetIsProtected => CurrentWorksheet.ProtectContents;

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

        public object CellWithText(string scope, string dataToSearch)
        {
            var lookAt = scope.StartsWith("part", StringComparison.InvariantCultureIgnoreCase)
                ? XlLookAt.xlPart
                : XlLookAt.xlWhole;
            var cell = CurrentWorksheet.Cells.Find(dataToSearch, LookAt: lookAt, MatchCase: false);
            return cell?.Address;
        }

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

        public void CloseExcel()
        {
            if (_excel == null) return;
            foreach (Workbook workbook in _excel.Workbooks)
            {
                workbook.Close(false, Type.Missing, Type.Missing);
            }
            _excel.Quit();
        }

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

        public bool CloseWorkbook() => _currentWorkbook != null && CloseWorkbook(_currentWorkbook.FullName);

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

        public bool LoadWorkbook(string path) => LoadWorkbook(path, false, null);

        public bool LoadWorkbookReadOnly(string path) => LoadWorkbook(path, true, null);

        public bool LoadWorkbookReadOnlyWithPassword(string path, string password) => LoadWorkbook(path, true, password);

        public bool LoadWorkbookWithPassword(string path, string password) => LoadWorkbook(path, false, password);

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

        public bool ProtectWorkbookWithPassword(string password)
        {
            if (WorkbookIsProtected) return false;

            // Protect doesn't automatically set Password. So do that separately.
            // We need this since WorkbookIsProtected relies on HasPassword.
            _currentWorkbook.Protect(password);
            _currentWorkbook.Password = password;
            return true;
        }

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

        public bool SelectWorkbook(string workbookPath)
        {
            var workbook = FindWorkbook(workbookPath);
            if (workbook == null) return false;
            _currentWorkbook = workbook;
            CurrentWorksheet = _currentWorkbook.ActiveSheet;
            return true;
        }

        public bool SelectWorksheet(string sheetName)
        {
            if (_currentWorkbook == null) return false;
            CurrentWorksheet = int.TryParse(sheetName, out var sheetNumber)
                ? _currentWorkbook.Sheets[sheetNumber]
                : _currentWorkbook.Sheets[sheetName];
            return CurrentWorksheet != null;
        }

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

        public object TextOfCell(string cellLocation)
        {
            var cell = CurrentWorksheet.Range[cellLocation];
            return cell.Text;
        }

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

        public object ValueOfCell(string cellLocation)
        {
            var cell = CurrentWorksheet.Range[cellLocation];
            return cell.Value2;
        }
    }
}