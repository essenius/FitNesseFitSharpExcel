﻿// Copyright 2016=5-2017 Rik Essenius
// Licensed under the Apache License, Version 2.0 (the "License");
// you may not use this file except in compliance with the License. You may obtain a copy of the License at
//
//   http://www.apache.org/licenses/LICENSE-2.0
//
// Unless required by applicable law or agreed to in writing, software distributed under the License is 
// distributed on an "AS IS" BASIS, WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
// See the License for the specific language governing permissions and limitations under the License.

using Microsoft.Office.Interop.Excel;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace ExcelFixtureTest
{
    [TestClass]
    public class CoverageIssueTest
    {
        [TestMethod] //Test showing bug in ExcelInterop + code coverage. Does not seem to occur in Excel 2016
        public void CodeCoverageTest()
        {
            var excel = new Application {Visible = true};
            Assert.IsNotNull(excel);
            excel.Quit();
        }
    }
}