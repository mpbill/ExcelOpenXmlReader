using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Linq;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Packaging;
using System.IO;

namespace ExcelOpenXmlReader
{
    public class ExcelOpenXmlReader
    {
        private Regex alphaRegex;
        private Regex numRegex;
        private string path;
        private List<string> SharedStrings;
        private List<ExcelOpenXmlRow> MyRows;
        public ExcelOpenXmlReader(string path)
        {
            this.path = path;
            alphaRegex = new Regex("[a-zA-Z]+", RegexOptions.Compiled);
            numRegex = new Regex("[0-9]+", RegexOptions.Compiled);
            
            SpreadsheetDocument doc = SpreadsheetDocument.Open(File.Open(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite), false);
            SharedStrings =
                (from item in
                    doc.WorkbookPart.GetPartsOfType<SharedStringTablePart>()
                        .First()
                        .SharedStringTable.Descendants<SharedStringItem>()
                    select item.InnerText).ToList();
            var Rows = doc.WorkbookPart.WorksheetParts.First().Worksheet.Descendants<Row>();
            MyRows = new List<ExcelOpenXmlRow>(Rows.Count());
            foreach (var row in Rows)
            {
                var cells = row.Descendants<Cell>();
                ExcelOpenXmlRow myRow = new ExcelOpenXmlRow(cells, SharedStrings);
                MyRows.Add(myRow);
            }


        }

        public DataSet AsDataSet()
        {
            return new DataSet();
        }

    }

    public class ExcelOpenXmlRow : IDisposable
    {
        public List<object> ItemList { get; set; }
        public List<string> SharedStrings; 
        public List<TheCell> TheCells { get; set; }
        
        public ExcelOpenXmlRow(IEnumerable<Cell> cells, List<string> SharedStrings )
        {
            this.SharedStrings = SharedStrings;
            TheCells = (from cell in cells select new TheCell(cell, SharedStrings)).ToList();
            ItemList = new List<object>(TheCells.Last().Column);
            foreach (TheCell theCell in TheCells)
            {
                while (ItemList.Count < theCell.ColumnZeroBase)
                {
                    ItemList.Add(System.DBNull.Value);
                }
                ItemList.Add(theCell.Value);
            }
        }

        
        

        public void Dispose()
        {
            throw new NotImplementedException();
        }
    }

    

    public struct NumLetters
    {
        public IEnumerable<char> Letters
        {
            get
            {
                for (char i = 'A'; i <= 'Z'; i++)
                {
                    for (int j = 0; j <= 1; j++)
                    {
                        if(j==0)
                            yield return i;
                        else
                        {
                            yield return i.ToString().ToLowerInvariant().ToCharArray()[0];
                        }
                    }

                }
                for (char i = 'a'; i < 'Z'; i++)
                {
                    yield return i;
                }
            }
        }

        public IEnumerable<char> Numbers
        {
            get
            {
                foreach (int i in Enumerable.Range(0,10))
                {
                    yield return i.ToString().ToCharArray()[0];
                }
            }
        } 
    }
    public class TheCell
    {
        public int Row { get; set; }
        public int Column { get; set; }
        public int ColumnZeroBase { get { return Column - 1; } }
        public object Value { get; set; }
        public string ColumnName { get; set; }
        public Type Type { get; set; }
        private NumLetters NumLetters;

        
        public TheCell(Cell cell, List<string> SharedStrings)
        {
            
            NumLetters = new NumLetters();
            ParseCellName(cell.CellReference);
            Column = ExcelColumnNameToNumber(ColumnName);
            ParseCellValue(cell, SharedStrings);

        }

        public void ParseCellValue(Cell cell, List<string> SharedStrings)
        {
            if (cell.DataType != null && cell.CellValue != null)
            {
                if (cell.DataType == CellValues.Date)
                {
                    Value =  DateTime.FromOADate(double.Parse(cell.CellValue.InnerText));
                    Type = typeof (DateTime);
                }
                else if (cell.DataType == CellValues.SharedString)
                {
                    Value = SharedStrings[int.Parse(cell.CellValue.InnerText)];
                    Type = typeof (string);
                }
                else if (cell.DataType == CellValues.String || cell.DataType == CellValues.InlineString)
                {
                    Value = cell.CellValue.InnerText;
                    Type = typeof (string);
                }
                else if (cell.DataType == CellValues.Boolean)
                {
                    Debugger.Break();
                    Value =  System.DBNull.Value;
                    Type = typeof (System.DBNull);

                }
                else if (cell.DataType == CellValues.Number)
                {
                    Value = double.Parse(cell.CellValue.InnerText);
                    Type = typeof (double);
                }
                else if (cell.DataType == CellValues.Error)
                {
                    Value = cell.CellValue.InnerText;
                    Type = typeof (ErrorItem);
                }

                else
                {
                    Value = System.DBNull.Value;
                    Type = typeof (System.DBNull);
                    Debugger.Break();
                }
            }
            else if (cell.CellValue == null)
            {
                Value = System.DBNull.Value;
                Type = typeof (DBNull);
            }
            else if (cell.CellValue != null)
            {
                Value = cell.CellValue.InnerText;
                Type = typeof (string);
            }
            else
            {
                Debugger.Break();
                Value = System.DBNull.Value;
                Type = typeof (DBNull);
            }
        }
        private void ParseCellName(string cellName)
        {
            string row = string.Empty;
            foreach (char c in cellName)
            {
                if (NumLetters.Letters.Contains(c))
                    ColumnName += c;
                else
                    row += c;

            }
            Row = int.Parse(row);
        }
        public int ExcelColumnNameToNumber(string columnName)
        {
            if (string.IsNullOrEmpty(columnName)) throw new ArgumentNullException("columnName");

            columnName = columnName.ToUpperInvariant();

            int sum = 0;

            for (int i = 0; i < columnName.Length; i++)
            {
                sum *= 26;
                sum += (columnName[i] - 'A' + 1);
            }

            return sum;
        }
    }
}