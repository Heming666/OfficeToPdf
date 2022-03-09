using System.IO;
using Spire.Doc;
using Spire.Doc.Documents;

namespace ConsoleApp1
{
    class Program
    {
        static void Main(string[] args)
        {
            //var current = AppDomain.CurrentDomain.BaseDirectory;
            //var source = Path.Combine(current, "22.pptx");
            //var target = Path.Combine(current, "22.pdf");
            //Common.OfficeToPDF.ConvertToPdf(source, target);
            //source = Path.Combine(current, "fabdcf2f-5cbc-47ee-ad28-1c492b73323a.ppt");
            // target = Path.Combine(current, "11.pdf");
            //Common.OfficeToPDF.ConvertToPdf(source, target);
            string path = @"C:\Users\hm80\Desktop\web前端求职简历.html";
                Document document = new Document();
                document.LoadFromFile(path,FileFormat.Html,XHTMLValidationType.None);
                document.SaveToFile(@"C:\Users\hm80\Desktop\11.pdf",FileFormat.PDF);
                document.Close();
                document.Dispose();

        }
    }
}
