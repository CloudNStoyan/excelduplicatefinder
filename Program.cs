using System;
using System.Linq;
using LinqToExcel;

namespace ExcelDuplicateFinder
{
    class Program
    {
        static void Main(string[] args)
        {
            const string excelFilePath = "./data.xlsx";

            var excel = new ExcelQueryFactory(excelFilePath);

            var cols = excel.WorksheetNoHeader();

            var rows = cols.Select(x => x[0]).ToArray();
            Console.WriteLine(rows[1]);
        }
    }
}
