using Aspose.Words;
using System;

namespace Aspose.POC
{
    internal class ConvertDocxToPDF
    {
        #region "Action Handlers (command handler)"
        private static void ConvertDocAction(ConvertDocumentRequest obj)
        {

            Document doc = new Document(obj.DocumentId);
            // 1) preprocessing -->unprotect , watermark, header and footer

            // Apply Protection to Document
            ApplyProtection(doc, ProtectionType.NoProtection, "pwd");

            // Remove Comments
            RemoveComments(doc);

            // Apply Header
            // AddHeader(doc, "header text", "red", "10", "10", "vertical");

            // apply footer
              //AddFooter(doc, obj.OutputOptions.DocumentOptions.FooterOptions.Text, );
            // Apply watermark
            if (obj.OutputOptions.DocumentOptions.ApplyWatermark)
            {
                AddWatermark(doc, obj.OutputOptions.DocumentOptions.WatermarkOptions.Text, obj.OutputOptions.DocumentOptions.WatermarkOptions.TextColor);
            }

            //2) do convesion to pdf

            var dataDir = Common.FilePath + "ConvertDocAction_Output.pdf";

            // Save the finished document to disk.
            doc.Save(dataDir);

            //  3) post ---> output is pdf then use pdfoptions


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
            builder.Write("Header Text goes here...");
        }

        public static void AddFooter(Document doc, string text, string textColor, string fontSize, string horizontalAlignment)
        {
            DocumentBuilder builder = new DocumentBuilder(doc);
            Section currentSection = builder.CurrentSection;
            PageSetup pageSetup = currentSection.PageSetup;
            pageSetup.DifferentFirstPageHeaderFooter = false;
            pageSetup.OddAndEvenPagesHeaderFooter = false;
            builder.Font.Size = Convert.ToDouble(fontSize);
            builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
            builder.Write("Footer Text goes here...");

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

            //doc.Watermark.SetText("Aspose Watermark", textWatermarkOptions);
            doc.Watermark.SetText(text, textWatermarkOptions);
        }
        #endregion

        public static void Run()
        {
            Console.WriteLine("\nprocessing started...");

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

                    },

                }
            };

            ConvertDocAction(model);

            Console.WriteLine("\nGenerateDocument processing completed...");
        }
    }

    #region Models (request models)
    public class ConvertDocumentRequest
    {
        /// <summary>
        /// The document id.
        /// </summary>
        public string DocumentId { get; set; }

        /// <summary>
        /// Optional file name for the converted document.
        /// </summary>
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
        public HeaderOptions HeaderOptions { get; set; }
        public FooterOptions FooterOptions { get; set; }
        public WatermarkOptions WatermarkOptions { get; set; }
        public string Password { get; set; }
    }
    public class PDFOptions
    {
        public string Version { get; set; }
        public bool QuickRender { get; set; }
        public bool ScreenReaders { get; set; }
        public bool EmbedFullFontSet { get; set; }
        public string ConformanceLevel { get; set; }
        public short LockOutputFile { get; set; }
        public bool FlattenForms { get; set; }
        public short ImageQuality { get; set; }
        public bool ApplySecurity { get; set; }
        public bool AllowPrinting { get; set; }
        public bool AllowCopying { get; set; }
        public bool AllowPageExtraction { get; set; }
        public bool AllowCommenting { get; set; }
        public bool AllowFormFilling { get; set; }
        public bool AllowSigning { get; set; }
        public string Password { get; set; }
    }

    public class HeaderOptions
    {
        public string Text { get; set; }
        public string TextColor { get; set; }
        public string FontFamily { get; set; }
        public string FontSize { get; set; }
        public string FontWeight { get; set; }
        public string HorizontalAlignment { get; set; }
    }
    public class FooterOptions
    {
        public string Text { get; set; }
        public string TextColor { get; set; }
        public string FontFamily { get; set; }
        public string FontSize { get; set; }
        public string FontWeight { get; set; }
        public string HorizontalAlignment { get; set; }
    }

    public class WatermarkOptions
    {
        public string Text { get; set; }
        public string TextColor { get; set; }
    }
    #endregion




}



