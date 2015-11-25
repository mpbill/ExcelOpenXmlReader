using System;
using System.Collections;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Linq;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Packaging;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using System.Xml;

namespace ExcelOpenXmlReader
{
    public class OpenXmlWorkbook
    {
        public List<string> SharedStrings { get; set; }
        private SpreadsheetDocument Document { get; set; }
        public BlockingCollection<OpenXmlWorksheet> OpenXmlWorksheets { get; set; }
        private MemoryStream ms;
        public OpenXmlWorkbook(string path)
        {
            LoadStream(path);
            InitializeDocument();
            InitializeSharedStrings();
            OpenXmlWorksheets = new BlockingCollection<OpenXmlWorksheet>();
            Document.WorkbookPart.WorksheetParts.ToList().ForEach(part => {ThreadMethod(part);});
        }
        private void InitializeSharedStrings()
        {
            SharedStrings =
                (from item in
                    Document.WorkbookPart.GetPartsOfType<SharedStringTablePart>()
                        .First()
                        .SharedStringTable.Descendants<SharedStringItem>()
                 select item.InnerText).ToList();
        }
        private void AddWorksheet(OpenXmlWorksheet sheet)
        {
            OpenXmlWorksheets.Add(sheet);
        }
        private void LoadStream(string path)
        {
            using (var fs = File.Open(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                ms = new MemoryStream();
                fs.CopyTo(ms);
                

            }
        }
        private void InitializeDocument()
        {
            Document = SpreadsheetDocument.Open(ms, false);
        }
        private void ThreadMethod(WorksheetPart param)
        {
            
            try
            {
                var ws = new OpenXmlWorksheet(param, SharedStrings);
                OpenXmlWorksheets.Add(ws);
            }
            catch (XmlException)
            {
                return;
            }
            catch (InvalidDataException)
            {
                return;
            }
            catch (InvalidOperationException)
            {
                return;
            }
        }
    }
    public class OpenXmlWorksheet
    {
        public List<string> SharedStrings { get; set; }
        public WorksheetPart WorksheetPart { get; set; }
        private BlockingCollection<ExcelOpenXmlRow> MyRows { get; set; }
        public DataTable DataTable { get; private set; }
        public OpenXmlWorksheet(WorksheetPart part, List<string> sharedStrings)
        {
            WorksheetPart = part;
            SharedStrings = sharedStrings;
            InitializeMyRows();
            WorksheetPart.Worksheet.Descendants<Row>().AsParallel().ForAll(row =>
            {
                AddRow(row);
            });


        }
        private void AddRow(Row descendant)
        {
            MyRows.Add(new ExcelOpenXmlRow(descendant.Descendants<Cell>(), SharedStrings));
        }
        private void InitializeMyRows()
        {
            var matches = CompiledRegexPatterns.ParseDeminsions.Matches(WorksheetPart.Worksheet.Descendants<SheetDimension>().First().Reference.InnerText);
            MyRows = new BlockingCollection<ExcelOpenXmlRow>(int.Parse(matches[1].Value) - int.Parse(matches[0].Value));
        }
        public static async Task InitAsync(WorksheetPart part, List<string> sharedString, Action<OpenXmlWorksheet> Callback)
        {
            
            var t =  Task.Factory.StartNew(() =>
            {
                try
                {
                    var ws = new OpenXmlWorksheet(part, sharedString);
                    return ws;
                }
                catch(XmlException)
                {
                    return null;
                }
                catch(InvalidDataException)
                {
                    return null;
                }
                catch(InvalidOperationException)
                {
                    return null;
                }
                
            }, CancellationToken.None, TaskCreationOptions.LongRunning, TaskScheduler.Default);
            await t;
            Callback.Invoke(t.Result);
            
            
        }


    }

    public class ExcelOpenXmlRow : IDisposable
    {
        public List<object> ItemList { get; set; }
        public List<string> SharedStrings; 
        public List<TheCell> TheCells { get; set; }
        
        public ExcelOpenXmlRow(IEnumerable<Cell> cells, List<string> SharedStrings )
        {
            if (!cells.Any())
                return;
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
    
    public static class CompiledRegexPatterns
    {
        public static Regex ParseDeminsions = new Regex("[0-9]+", RegexOptions.Compiled);
    }

    public static class NumLetters
    {
        public static char[] Numbers = "1234567890".ToCharArray();

        public static char[] Letters = "qQwWeErRtTyYuUiIoOpPaAsSdDfFgGhHjJkKlLzZxXcCvVbBnNmM".ToCharArray();

    }
    public class TheCell
    {
        public int Row { get; set; }
        public int Column { get; set; }
        public int ColumnZeroBase { get { return Column - 1; } }
        public object Value { get; set; }
        public string ColumnName { get; set; }
        public Type Type { get; set; }
        

        
        public TheCell(Cell cell, List<string> SharedStrings)
        {
            
            
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
                    Type = typeof(System.DBNull);
                    if (cell.CellValue.InnerText == "0")
                        Value = false;
                    else if (cell.CellValue.InnerText == "1")
                        Value = true;
                    else
                    {
                        Debugger.Break();
                    }

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