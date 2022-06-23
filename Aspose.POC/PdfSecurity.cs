using Aspose.Pdf;
using Aspose.Pdf.Facades;
using Aspose.Words;
using System;
using Document = Aspose.Words.Document;

namespace Aspose.POC
{
    public class PdfSecurity
    {
        public static void Run()
        {
            Console.WriteLine("\nSetPdfPrivileges Processing started... ");


            SecureDocumentRequest request = new SecureDocumentRequest
            {
                DocumentId = Common.FilePath + "setsecurityInputPdf.pdf",
                Password = "qwerty",
                AllowPrinting = true,
                AllowCommenting = false,
                AllowCopying = false,
                AllowSigning = false,
                AllowFormFilling = false,
                DocFormat = "pdf",               
            };

            if (request.DocFormat == "pdf")
            {
                SetPdfPrivileges(request);
            }
            else if (request.DocFormat == "docx")
            {
                ApplyProtectionLevel(request);
            }
            else if (request.DocFormat == "docx" && request.SecurityEnabled == true)
            {
                RemoveProtectionLevel();
            }
            else
            {
                Console.WriteLine("\nDocument Format not valid\n");
            }

            Console.WriteLine("\nSetPdfPrivileges processing completed...");

        }

        public static void SetPdfPrivileges(SecureDocumentRequest obj)
        {
            PdfFileSecurity security = new PdfFileSecurity();

            DocumentPrivilege privilege = DocumentPrivilege.ForbidAll;
            privilege.ChangeAllowLevel = 3;
            privilege.AllowPrint = obj.AllowPrinting;
            privilege.AllowCopy = obj.AllowCopying;
            privilege.AllowFillIn = obj.AllowFormFilling;
            privilege.AllowModifyAnnotations = obj.AllowCommenting;
           
            security.BindPdf(obj.DocumentId);
            security.SetPrivilege(privilege);
            security.Save(Common.FilePath + "setsecurityInputPdf_out_.pdf");

        }

        public static void ApplyProtectionLevel(SecureDocumentRequest obj)
        {
            string docFilePath = Common.FilePath + "ApplyProtectionLevelInputDoc.docx";
            Document doc = new Document(docFilePath);
            doc.Protect(ProtectionType.ReadOnly, obj.Password);
            doc.Save(Common.FilePath + "ApplyProtectionLevelInputDoc_out_.docx");
        }
        public static void RemoveProtectionLevel()
        {
            string docFilePath = Common.FilePath + "ApplyProtectionLevelInputDoc_out_.docx";
            Document doc = new Document(docFilePath);
            if (doc.ProtectionType == ProtectionType.ReadOnly)
            {
                doc.Unprotect();
                doc.Save(Common.FilePath + "ApplyProtectionLevelInputDoc_out_.docx");
            }
        }

    }
}

