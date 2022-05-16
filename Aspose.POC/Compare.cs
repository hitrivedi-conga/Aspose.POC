using Aspose.Words;
using Aspose.Words.Comparing;
using System;

namespace Aspose.POC {
    internal class Compare {
        public static void Run() {
            Console.WriteLine("\nCompare processing started...");

            Document docOriginal = new Document(Common.FilePath + "Compare-Original.docx");
            Document docEdited = new Document(Common.FilePath + "Compare-Revised.docx");

            // Apply different comparing options.
            CompareOptions compareOptions = new CompareOptions();
            compareOptions.IgnoreFormatting = true;
            compareOptions.IgnoreCaseChanges = false;
            compareOptions.IgnoreComments = false;
            compareOptions.IgnoreTables = false;
            compareOptions.IgnoreFields = false;
            compareOptions.IgnoreFootnotes = false;
            compareOptions.IgnoreTextboxes = false;
            compareOptions.IgnoreHeadersAndFooters = false;
            compareOptions.Target = ComparisonTargetType.New;

            // compare document must not have revisions.
            if(docOriginal.Revisions.Count > 0) docOriginal.Revisions.AcceptAll();
            if(docEdited.Revisions.Count > 0) docEdited.Revisions.AcceptAll();

            // compare both documents.
            docOriginal.Compare(docEdited, "HIREN T", DateTime.Now, compareOptions);
            docOriginal.Save(Common.FilePath + "Compare_out_.docx");

            Console.WriteLine("\nCompare processing completed...");
        }
    }
}
