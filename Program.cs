using System;
using System.Collections.Generic;
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

            var dictionary = new Dictionary<string, int>();

            foreach (var cell in rows)
            {
                string text = cell.Cast<string>();

                if (!dictionary.ContainsKey(text))
                {
                    dictionary.Add(text, 1);
                }
                else
                {
                    dictionary[text]++;
                }
            }

            foreach (var pair in dictionary)
            {
                if (pair.Value > 1)
                {
                    Console.WriteLine($"\"{pair.Key}\" was found {pair.Value} times");
                }
            }
        }
    }
}
