using OpenXmlServer;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenXML
{
    class Program
    {
        static void Main(string[] args)
        {
            var wordHelper = new WordHelper();
            wordHelper.Init(@"C:\Users\Lee\Desktop\Doc1.docx");
            var excelHelper = new ExcelHelper();
            excelHelper.Init(@"C:\Users\Lee\Desktop\Sheet.xlsx", @"C:\Users\Lee\Desktop\Sheet1.xlsx", "{}");
            Console.ReadLine();
        }
    }
}
