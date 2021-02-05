using Aspose.Slides;
using Aspose.Words;
using Aspose.Words.Saving;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WordToPdf
{
    class Wlppt
    {

        public  void pptToPdf(string from, string to)
        {
            Presentation ppt = new Presentation(from);
            ppt.Save(to, Aspose.Slides.Export.SaveFormat.Pdf);
            Console.WriteLine("成功!");
        }




    }
}
