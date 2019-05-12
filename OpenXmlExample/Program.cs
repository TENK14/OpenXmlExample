using System;
using System.IO;

namespace OpenXmlExample
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Hello World!");

            //using (FileStream fileStream = File.Create(Path.Combine(@"C:\Users\wagner.HAVIT\Downloads\168.UPB", "excel.xlsx")))
            using (FileStream fileStream = new FileStream(Path.Combine(@"C:\Users\wagner.HAVIT\Downloads\168.UPB", "excel.xlsx"), FileMode.Create))
            {
                var service = new ExcelService();
                service.CreateExcel(fileStream);
            }

            Console.WriteLine("End");
        }
    }
}
