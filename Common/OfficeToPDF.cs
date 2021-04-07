using Aspose.Cells;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Common
{
   public static  class OfficeToPDF
    {
        public static void ConvertToPdf(string sourcePath, string savePath)
        {
            var extension = Path.GetExtension(sourcePath);
            switch (extension)
            {
                case ".doc":
                case ".docx":
                    Aspose.Words.Document document = new Aspose.Words.Document(sourcePath);
                    document.Save(savePath,Aspose.Words.SaveFormat.Pdf);
                    break;
                case ".xls":
                case ".xlsx":
                    var book = new Workbook(sourcePath);
                    book.Save(savePath, SaveFormat.Pdf);
                    break;
                case ".ppt":
                case ".pptx":
                    Aspose.Slides.Presentation presentation = new Aspose.Slides.Presentation(sourcePath);
                    presentation.Save(savePath, Aspose.Slides.Export.SaveFormat.Pdf);
                    break;
                default:
                    WriteLog.AddLog("待转换的文件不是office类型的文件");
                    break;
            }
        }
    }
}
