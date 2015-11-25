using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Spreadsheet;
using ExcelOpenXmlReader;


namespace test
{
    class Program
    {
        static void Main(string[] args)
        {
            OpenXmlWorkbook wb = new OpenXmlWorkbook("NorthwindPlus.xlsx");
        }
    }
}
