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
using System.Collections.ObjectModel;
using Microsoft.Office.Interop.Excel;

namespace ExcelFixture
{
    /// <summary>Fixture to query a range of an Excel sheet</summary>
    public class ExcelQuery
    {
        private readonly string _options;
        private readonly Range _range;

        /// <summary>Query a range of an excel sheet. Parameters: script fixture, range</summary>
        /// <param name="excel">Fixture</param>
        /// <param name="range">the range to use in Excel format (e.g. C4:D8)</param>
        public ExcelQuery(Excel excel, object range) : this(excel, range, string.Empty)
        {
        }

        /// <summary>Query a range of an excel sheet. Parameters: script fixture, range, useheaders</summary>
        /// <param name="excel">script fixture</param>
        /// <param name="range">range spec in Excel format (e.g. C4:D8)</param>
        /// <param name="options">useheaders or nothing. If useheaders: the first row is expected to contain the headers</param>
        public ExcelQuery(Excel excel, object range, string options)
        {
            _range = excel.CurrentWorksheet.Range[range];
            _options = options;
        }

        /// <returns>the selected data range as a query result</returns>
        public Collection<object> Query()
        {
            var headerCollection = new Collection<string>();
            var useHeaders = _options.Equals(@"useheaders", StringComparison.InvariantCultureIgnoreCase);
            for (var i = 1; i <= _range.Columns.Count; i++)
            {
                headerCollection.Add(useHeaders ? _range[1, i].Value : "Column " + (i + _range.Column - 1));
            }
            var startRow = useHeaders ? 2 : 1;
            var rowCollection = new Collection<object>();
            for (var row = startRow; row <= _range.Rows.Count; row++)
            {
                var cellCollection = new Collection<object>();
                for (var column = 1; column <= _range.Columns.Count; column++)
                {
                    cellCollection.Add(new Collection<object> {headerCollection[column - 1], _range[row, column].Value});
                }
                rowCollection.Add(cellCollection);
            }
            return rowCollection;
        }
    }
}
