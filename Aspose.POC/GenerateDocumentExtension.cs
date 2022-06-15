using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;
using Aspose.Words.Tables;
using System;

namespace Aspose.POC
{
    internal static class GenerateDocumentExtension
    {
        public static void TableOfContent(Document doc)
        {
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert a table of contents at the beginning of the document.
            builder.InsertTableOfContents("\\o \"1-3\" \\h \\z \\u");

            //// Start the actual document content on the second page.
            //builder.InsertBreak(BreakType.PageBreak);

            //// Build a document with complex structure by applying different heading styles thus creating TOC entries.
            //builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;

            //builder.Writeln("Heading 1");

            //builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading2;

            //// Insert a TC field at the current document builder position.
            //builder.InsertField("TC \"External Doc\" \\f t");

            // The newly inserted table of contents will be initially empty.
            // It needs to be populated by updating the fields in the document.
            doc.UpdateFields();
        }

        public static void AddingHeader_Footer(Document doc)
        {
            DocumentBuilder builder = new DocumentBuilder(doc);

            Section currentSection = builder.CurrentSection;
            PageSetup pageSetup = currentSection.PageSetup;

            // Specify if we want headers/footers of the first page to be different from other pages.
            // You can also use PageSetup.OddAndEvenPagesHeaderFooter property to specify
            // Different headers/footers for odd and even pages.
            pageSetup.DifferentFirstPageHeaderFooter = true;

            // --- Create header for the first page. ---
            pageSetup.HeaderDistance = 20;
            builder.MoveToHeaderFooter(HeaderFooterType.HeaderFirst);
            builder.ParagraphFormat.Alignment = ParagraphAlignment.Center;

            // Set font properties for header text.
            builder.Font.Name = "Arial";
            builder.Font.Bold = true;
            builder.Font.Size = 14;
            // Specify header title for the first page.
            builder.Write("Aspose.Words Header/Footer Creation Primer - Title Page.");

            // --- Create header for pages other than first. ---
            pageSetup.HeaderDistance = 20;
            builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);

            // Insert absolutely positioned image into the top/left corner of the header.
            // Distance from the top/left edges of the page is set to 10 points.
            string imageFileName = Common.FilePath + "Technology-Transparent.png";
            builder.InsertImage(imageFileName, RelativeHorizontalPosition.Page, 10, RelativeVerticalPosition.Page, 10, 35, 35, WrapType.Through);

            builder.ParagraphFormat.Alignment = ParagraphAlignment.Right;
            // Specify another header title for other pages.
            builder.Write("Aspose.Words Header/Footer Creation Primer.");

            // --- Create footer for pages other than first. ---
            builder.MoveToHeaderFooter(HeaderFooterType.FooterPrimary);

            // We use table with two cells to make one part of the text on the line (with page numbering)
            // To be aligned left, and the other part of the text (with copyright) to be aligned right.
            builder.StartTable();

            // Clear table borders.
            builder.CellFormat.ClearFormatting();

            builder.InsertCell();

            // Set first cell to 1/3 of the page width.
            builder.CellFormat.PreferredWidth = PreferredWidth.FromPercent(100 / 3);

            // Insert page numbering text here.
            // It uses PAGE and NUMPAGES fields to auto calculate current page number and total number of pages.
            builder.Write("Page ");
            builder.InsertField("PAGE", "");
            builder.Write(" of ");
            builder.InsertField("NUMPAGES", "");

            // Align this text to the left.
            builder.CurrentParagraph.ParagraphFormat.Alignment = ParagraphAlignment.Left;

            builder.InsertCell();
            // Set the second cell to 2/3 of the page width.
            builder.CellFormat.PreferredWidth = PreferredWidth.FromPercent(100 * 2 / 3);

            builder.Write("(C) 2022 Aspose Pty Ltd. All rights reserved.");

            // Align this text to the right.
            builder.CurrentParagraph.ParagraphFormat.Alignment = ParagraphAlignment.Right;

            builder.EndRow();
            builder.EndTable();

            builder.MoveToDocumentEnd();
            // Make page break to create a second page on which the primary headers/footers will be seen.
            builder.InsertBreak(BreakType.PageBreak);

            // Make section break to create a third page with different page orientation.
            builder.InsertBreak(BreakType.SectionBreakNewPage);

            // Get the new section and its page setup.
            currentSection = builder.CurrentSection;
            pageSetup = currentSection.PageSetup;

            //// Configure the section to have the page count that PAGE fields display start from 5.
            //// Also, configure all PAGE fields to display their page numbers using uppercase Roman numerals.
            ////pageSetup.RestartPageNumbering = true;
            //pageSetup.PageStartingNumber = 1;
            //pageSetup.PageNumberStyle = NumberStyle.UppercaseRoman;

            // Set page orientation of the new section to landscape.
            pageSetup.Orientation = Orientation.Landscape;

            // This section does not need different first page header/footer.
            // We need only one title page in the document and the header/footer for this page
            // Has already been defined in the previous section
            pageSetup.DifferentFirstPageHeaderFooter = false;

            // This section displays headers/footers from the previous section by default.
            // Call currentSection.HeadersFooters.LinkToPrevious(false) to cancel this.
            // Page width is different for the new section and therefore we need to set 
            // A different cell widths for a footer table.
            currentSection.HeadersFooters.LinkToPrevious(false);

            // If we want to use the already existing header/footer set for this section 
            // But with some minor modifications then it may be expedient to copy headers/footers
            // From the previous section and apply the necessary modifications where we want them.
            CopyHeadersFootersFromPreviousSection(currentSection);

            // Find the footer that we want to change.
            HeaderFooter primaryFooter = currentSection.HeadersFooters[HeaderFooterType.FooterPrimary];

            Row row = primaryFooter.Tables[0].FirstRow;
            row.FirstCell.CellFormat.PreferredWidth = PreferredWidth.FromPercent(100 / 3);
            row.LastCell.CellFormat.PreferredWidth = PreferredWidth.FromPercent(100 * 2 / 3);
        }

        private static void CopyHeadersFootersFromPreviousSection(Section section)
        {
            Section previousSection = (Section)section.PreviousSibling;

            if (previousSection == null)
                return;

            section.HeadersFooters.Clear();

            foreach (HeaderFooter headerFooter in previousSection.HeadersFooters)
                section.HeadersFooters.Add(headerFooter.Clone(true));
        }

        public static void ApplyDocumentFormating(Document doc)
        {
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Specify font formatting
            Font font = builder.Font;
            font.Size = 16;
            font.Bold = true;
            font.Color = System.Drawing.Color.Blue;
            font.Name = "Arial";
            font.Underline = Underline.Dash;

            // Specify paragraph formatting
            ParagraphFormat paragraphFormat = builder.ParagraphFormat;
            paragraphFormat.FirstLineIndent = 8;
            paragraphFormat.Alignment = ParagraphAlignment.Justify;
            paragraphFormat.KeepTogether = true;

            builder.Writeln("A whole paragraph.");
        }

        public static void AddingOrRemovingSectionsInDocument(Document doc)
        {
            //Adding Section
            Section sectionToAdd = new Section(doc);
            doc.Sections.Add(sectionToAdd);

            //Removing Section - ether one
            doc.Sections.RemoveAt(0);
            //doc.Sections.Clear();

            //Adding Content to Section
            // This is the section that we will append and prepend to.
            Section section = doc.Sections[2];

            // This copies content of the 1st section and inserts it at the beginning of the specified section.
            Section sectionToPrepend = doc.Sections[0];
            section.PrependContent(sectionToPrepend);

            //Section Breaks
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.Writeln("This is page 1.");
            builder.InsertBreak(BreakType.PageBreak);

            builder.Writeln("This is page 2.");
            builder.InsertBreak(BreakType.PageBreak);
        }

        private static void Createbookmark(Document doc)
        {
            DocumentBuilder builder = new DocumentBuilder(doc);

            //BookmarkTable
            BookmarkTable(builder);

            builder.StartBookmark("My Bookmark");
            builder.Writeln("Text inside a bookmark.");

            builder.StartBookmark("Nested Bookmark");
            builder.Writeln("Text inside a NestedBookmark.");
            builder.EndBookmark("Nested Bookmark");

            builder.Writeln("Text after Nested Bookmark.");
            builder.EndBookmark("My Bookmark");


            PdfSaveOptions options = new PdfSaveOptions();
            options.OutlineOptions.BookmarksOutlineLevels.Add("My Bookmark", 1);
            options.OutlineOptions.BookmarksOutlineLevels.Add("Nested Bookmark", 2);
        }

        private static void BookmarkTable(DocumentBuilder builder)
        {
            Table table = builder.StartTable();

            // Insert a cell
            builder.InsertCell();

            // Start bookmark here after calling InsertCell
            builder.StartBookmark("MyBookmark");

            builder.Write("This is row 1 cell 1");

            // Insert a cell
            builder.InsertCell();
            builder.Write("This is row 1 cell 2");

            builder.EndRow();

            // Insert a cell
            builder.InsertCell();
            builder.Writeln("This is row 2 cell 1");

            // Insert a cell
            builder.InsertCell();
            builder.Writeln("This is row 2 cell 2");

            builder.EndRow();

            builder.EndTable();
            // End of bookmark
            builder.EndBookmark("MyBookmark");
        }
    }
}
