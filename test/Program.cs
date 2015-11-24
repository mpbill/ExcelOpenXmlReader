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
            ExcelOpenXmlReader.ExcelOpenXmlReader reader =
                new ExcelOpenXmlReader.ExcelOpenXmlReader(
                    @"D:\tempGitRepository\Projects\DanEmailGenerator\ABFS_EmailGeneratorUnitTest\110515-111915.xlsm");
            
        }
    }
}
