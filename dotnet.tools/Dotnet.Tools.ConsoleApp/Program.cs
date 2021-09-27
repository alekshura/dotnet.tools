using Dotnet.Tools.OpenXml.Excel;
using System;

namespace Dotnet.Tools.ConsoleApp
{
    class Program
    {
        static void Main(string[] args)
        {
            var excelCache = ExcelHelper.CachExcelFile(@"Path\to\you\excel.xlsx");
            Console.ReadKey();
        }
    }
}
