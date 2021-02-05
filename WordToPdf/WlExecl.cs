using Aspose.Words.Saving;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WordToPdf
{
    class WlExecl
    {

        public void ExeclToPDF(String from, String to)
        {
            try
            {

                Aspose.Cells.Workbook xls = new Aspose.Cells.Workbook(from);
                Aspose.Cells.PdfSaveOptions xlsSaveOption = new Aspose.Cells.PdfSaveOptions();
                xlsSaveOption.SecurityOptions = new Aspose.Cells.Rendering.PdfSecurity.PdfSecurityOptions();
                #region pdf 加密
                //Set the user password
                //PDF加密功能
                //xlsSaveOption.SecurityOptions.UserPassword = "pdfKey";
                //Set the owner password
                //xlsSaveOption.SecurityOptions.OwnerPassword = "sxbztxmgzxt";
                #endregion
                //Disable extracting content permission
                xlsSaveOption.SecurityOptions.ExtractContentPermission = false;
                //Disable print permission
                xlsSaveOption.SecurityOptions.PrintPermission = false;
                xlsSaveOption.AllColumnsInOnePagePerSheet = true;

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
                //doc.Save(to, saveOptions);

                xls.Save(to, xlsSaveOption);

                Console.WriteLine("转换成功!");
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
                Console.WriteLine("强行报错!");
            }

        }

    }
}
