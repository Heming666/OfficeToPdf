using System;
using System.IO;

namespace ConsoleApp1
{
    class Program
    {
        static void Main(string[] args)
        {
            var current = AppDomain.CurrentDomain.BaseDirectory;
            var source = Path.Combine(current, "22.pptx");
            var target = Path.Combine(current, "22.pdf");
            Common.OfficeToPDF.ConvertToPdf(source, target);
            source = Path.Combine(current, "fabdcf2f-5cbc-47ee-ad28-1c492b73323a.ppt");
             target = Path.Combine(current, "11.pdf");
            Common.OfficeToPDF.ConvertToPdf(source, target);


        }
    }
}
