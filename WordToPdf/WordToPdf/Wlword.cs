using Aspose.Words;
using Aspose.Words.Saving;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WordToPdf
{
    class Wlword
    {

        public void WordToPDF(String from,String to)
        {
            try
            {
                Document doc = new Document(from);
                //保存为PDF文件，此处的SaveFormat支持很多种格式，如图片，epub,rtf 等等

                //权限这块的设置成不可复制
                PdfSaveOptions saveOptions = new PdfSaveOptions();
                // Create encryption details and set owner password.
                PdfEncryptionDetails encryptionDetails = new PdfEncryptionDetails(string.Empty, "password", PdfEncryptionAlgorithm.RC4_128);
                // Start by disallowing all permissions.
                encryptionDetails.Permissions = PdfPermissions.DisallowAll;
                // Extend permissions to allow editing or modifying annotations.
                encryptionDetails.Permissions = PdfPermissions.ModifyAnnotations | PdfPermissions.DocumentAssembly;
                saveOptions.EncryptionDetails = encryptionDetails;
                // Render the document to PDF format with the specified permissions.
                doc.Save(to , saveOptions);
            
                //doc.Save(to, SaveFormat.Pdf);
                Console.WriteLine("成功!");
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
                Console.WriteLine("强行报错!");
            }
         
        }
    }
}
