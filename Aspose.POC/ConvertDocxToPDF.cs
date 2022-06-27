using Aspose.Pdf.Facades;
using Aspose.Words;
using System;
using System.IO;

namespace Aspose.POC
{
    internal class ConvertDocxToPDF
    {
        #region "Action Handlers (command handler)"

        private static void ConvertDocAction(ConvertDocumentRequest obj)
        {

            Document doc = new Document(obj.DocumentId);

            // Apply Protection to Document
            ApplyProtection(doc, obj.OutputOptions.DocumentOptions.ProtectionType, obj.OutputOptions.DocumentOptions.Password);

            // Remove Comments
            RemoveComments(doc);

            //AcceptAllRevisions
            AcceptAllRevisions(doc);

            // Apply Header
            AddHeader(doc, obj.OutputOptions.DocumentOptions.HeaderOptions.Text, obj.OutputOptions.DocumentOptions.HeaderOptions.TextColor,
                  obj.OutputOptions.DocumentOptions.HeaderOptions.IsBold, obj.OutputOptions.DocumentOptions.HeaderOptions.IsItalic,
                  obj.OutputOptions.DocumentOptions.HeaderOptions.FontSize, obj.OutputOptions.DocumentOptions.HeaderOptions.HorizontalAlignment);

            // Apply footer
            AddFooter(doc, obj.OutputOptions.DocumentOptions.FooterOptions.Text, obj.OutputOptions.DocumentOptions.FooterOptions.TextColor, 
                  obj.OutputOptions.DocumentOptions.FooterOptions.IsBold, obj.OutputOptions.DocumentOptions.FooterOptions.IsItalic, 
                  obj.OutputOptions.DocumentOptions.FooterOptions.FontSize, obj.OutputOptions.DocumentOptions.FooterOptions.HorizontalAlignment);

            // Apply watermark
            if (obj.OutputOptions.DocumentOptions.ApplyWatermark)
            {
                AddWatermark(doc, obj.OutputOptions.DocumentOptions.WatermarkOptions.Text, obj.OutputOptions.DocumentOptions.WatermarkOptions.TextColor);
            }

            //var abc = doc.Save("testfile", SaveFormat.pdf);
            var dataDir = Common.FilePath + "ConvertDocAction_Output.pdf";

            // Save the finished document to disk.
            doc.Save(dataDir);

            // Pdf Security
            if (obj.OutputOptions.PDFOptions.ApplySecurity)
            {
                SetPdfPrivileges(doc, obj.OutputOptions.PDFOptions.ApplySecurity, obj.OutputOptions.PDFOptions.AllowPrinting,
                obj.OutputOptions.PDFOptions.AllowCopying, obj.OutputOptions.PDFOptions.AllowCommenting, obj.OutputOptions.PDFOptions.AllowFormFilling,
                obj.OutputOptions.PDFOptions.DocumentFormat);
            }
            var dataDir1 = Common.FilePath + "ConvertDocAction2_Output.pdf";

            // Save the finished document to disk.
            doc.Save(dataDir1);
        }
        #endregion

        #region Common Utility Methods (common library)
        public static void ApplyProtection(Document doc, ProtectionType protectionType, string password)
        {
            if (protectionType != ProtectionType.NoProtection)
            {
                doc.Protect(protectionType, password);
            }
        }

        public static void RemoveComments(Document doc)
        {
            // Collect all comments in the document
            NodeCollection comments = doc.GetChildNodes(NodeType.Comment, true);
            // Remove all comments.
            comments.Clear();
        }

        public static void AcceptAllRevisions(Document doc)
        {
            doc.AcceptAllRevisions();
        }

        public static void AddHeader(Document doc, string text, string textColor, bool isBold, bool isItalic, string fontSize, string horizontalAlignment)
        {
            DocumentBuilder builder = new DocumentBuilder(doc);
            Section currentSection = builder.CurrentSection;
            PageSetup pageSetup = currentSection.PageSetup;
            pageSetup.DifferentFirstPageHeaderFooter = false;
            pageSetup.OddAndEvenPagesHeaderFooter = false;
            builder.Font.Bold = isBold;
            builder.Font.Italic = isItalic;
            builder.Font.Size = Convert.ToDouble(fontSize);
            builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
            builder.Write(text);
        }

        public static void AddFooter(Document doc, string text, string textColor, bool isBold, bool isItalic, string fontSize, string horizontalAlignment)
        {
            DocumentBuilder builder = new DocumentBuilder(doc);
            Section currentSection = builder.CurrentSection;
            PageSetup pageSetup = currentSection.PageSetup;
            pageSetup.DifferentFirstPageHeaderFooter = false;
            pageSetup.OddAndEvenPagesHeaderFooter = false;
            builder.Font.Bold = isBold;
            builder.Font.Italic = isItalic;
            builder.Font.Size = Convert.ToDouble(fontSize);
            builder.MoveToHeaderFooter(HeaderFooterType.FooterPrimary);
            builder.Write(text);
        }

        public static void AddWatermark(Document doc, string text, string textColor)
        {
            // If we wish to edit the text formatting using it as a watermark,
            // we can do so by passing a TextWatermarkOptions object when creating the watermark.
            TextWatermarkOptions textWatermarkOptions = new TextWatermarkOptions();
            textWatermarkOptions.FontFamily = "Arial";
            textWatermarkOptions.FontSize = 36;
            textWatermarkOptions.Layout = WatermarkLayout.Diagonal;
            //textWatermarkOptions.Layout = WatermarkLayout.Horizontal;
            textWatermarkOptions.IsSemitrasparent = false;

            doc.Watermark.SetText(text, textWatermarkOptions);
        }

        public static void SetPdfPrivileges(Document doc, bool applySecurity, bool allowPrinting, bool allowCopying, bool allowCommenting, 
            bool allowFormFillling, string documentFormat)
        {
            // Create DocumentPrivileges object
            DocumentPrivilege privilege = DocumentPrivilege.ForbidAll;
            privilege.ChangeAllowLevel = 1;
            privilege.AllowPrint = allowPrinting;
            privilege.AllowCopy = allowCopying;
            privilege.AllowModifyAnnotations = allowCommenting;

            // Create PdfFileSecurity object
            var dataDir = Common.FilePath + "ConvertDocAction_Output.pdf";
           
            //// DataDir1 to stream
            //var templateStream = new MemoryStream();
            //doc.Save(templateStream);

            //// Rewind the stream position back to zero.
            //templateStream.Position = 0;

            PdfFileSecurity fileSecurity = new PdfFileSecurity();
            fileSecurity.BindPdf(dataDir);
            fileSecurity.SetPrivilege(privilege);
        }
        #endregion

        public static void Run()
        {
            Console.WriteLine("\nApplying Common Utility Methods and converting docx to pdf...");

            string fileName = "convertWordDoc.docx";

            var model = new ConvertDocumentRequest
            {
                DocumentId = Common.FilePath + fileName,
                OutputOptions = new OutputOptions
                {
                    FileName = "convertWordDoc.docx",
                    OutputFormat = "pdf",
                    DocumentOptions = new DocumentOptions
                    {
                        ProtectionType = ProtectionType.NoProtection,
                        Password = "password",
                        ApplyWatermark = true,
                        HeaderOptions = new HeaderOptions
                        {
                            Text = "Header text goes here",
                            TextColor = "black",
                            IsBold = true,
                            IsItalic = true,
                            FontSize = "10",
                        },
                        FooterOptions = new FooterOptions
                        {
                            Text = "Footer text goes here",
                            TextColor = "black",
                            IsBold = true,
                            IsItalic = true,
                            FontSize = "10",
                        },
                        WatermarkOptions = new WatermarkOptions
                        {
                            Text = "Aspose Watermark",
                            TextColor = "black"
                        }
                    },
                    PDFOptions = new PDFOptions
                    {
                        ApplySecurity = true,
                        AllowPrinting = true,
                        AllowCommenting = false,
                        AllowCopying = false,
                        AllowFormFilling = false,
                        DocumentFormat = "pdf",
                    }
                }
            };

            ConvertDocAction(model);

            Console.WriteLine("\nSuccessfully Converted docx to pdf...");
        }
    }

    #region Models (request models)
    public class ConvertDocumentRequest
    {
        /// <summary>
        /// The document id.
        /// </summary>
        public string DocumentId { get; set; }

        public OutputOptions OutputOptions { get; set; }
    }

    public class OutputOptions
    {
        /// <summary>
        /// The output format of the generated document.
        /// </summary>
        public string OutputFormat { get; set; }

        /// <summary>
        /// The output file name of the generated document.
        /// </summary>
        public string FileName { get; set; }
        public DocumentOptions DocumentOptions { get; set; }
        public PDFOptions PDFOptions { get; set; }

    }

    public class DocumentOptions
    {
        public bool ApplyWatermark { get; set; }
        public bool ApplyHeader { get; set; }
        public bool ApplyFooter { get; set; }
        public ProtectionType ProtectionType { get; set; }
        public HeaderOptions HeaderOptions { get; set; }
        public FooterOptions FooterOptions { get; set; }
        public WatermarkOptions WatermarkOptions { get; set; }
        public string Password { get; set; }
    }
    public class PDFOptions
    {
        public bool ApplySecurity { get; set; }
        public bool AllowPrinting { get; set; }
        public bool AllowCopying { get; set; }
        public bool AllowCommenting { get; set; }
        public bool AllowFormFilling { get; set; }
        public string DocumentFormat { get; set; }
        public string Password { get; set; }
    }

    public class HeaderOptions
    {
        public string Text { get; set; }
        public string TextColor { get; set; }
        public bool IsBold { get; set; }
        public bool IsItalic { get; set; }
        public string FontSize { get; set; }
        public string HorizontalAlignment { get; set; }
    }
    public class FooterOptions
    {
        public string Text { get; set; }
        public string TextColor { get; set; }
        public bool IsBold { get; set; }
        public bool IsItalic { get; set; }
        public string FontSize { get; set; }
        public string HorizontalAlignment { get; set; }
    }

    public class WatermarkOptions
    {
        public string Text { get; set; }
        public string TextColor { get; set; }
    }
    #endregion
}



