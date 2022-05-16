using Aspose.Words;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Aspose.POC {
    internal class CreateWordPackage {
        public static void Run() {
            Console.WriteLine("\nCreateWordPackage processing started...");
            
            Document dstDoc = new Document();

            // The destination document is not actually empty which often causes a blank page to appear before the appended document
            // This is due to the base document having an empty section and the new document being started on the next page.
            // Remove all content from the destination document before appending.
            dstDoc.RemoveAllChildren();

            string[] sourceDocuments = new string[] {
                Common.FilePath + "InputDocument-1.docx",
                Common.FilePath + "InputDocument-2.docx",
                Common.FilePath + "InputDocument-3.docx"
            };

            foreach(string sourceDocument in sourceDocuments) {
                Document srcDoc = new Document(sourceDocument);
                ImportFormatMode mode = ImportFormatMode.KeepSourceFormatting;

                // Loop through all sections in the source document. 
                // Section nodes are immediate children of the Document node so we can just enumerate the Document.
                foreach(Section srcSection in srcDoc) {
                    // Because we are copying a section from one document to another, 
                    // it is required to import the Section node into the destination document.
                    // This adjusts any document-specific references to styles, lists, etc.
                    //
                    // Importing a node creates a copy of the original node, but the copy
                    // is ready to be inserted into the destination document.
                    Node dstSection = dstDoc.ImportNode(srcSection, true, mode);

                    // Now the new section node can be appended to the destination document.
                    dstDoc.AppendChild(dstSection);
                }
            }

            var outPut = Common.FilePath + "CreateWordPackage_out_.docx";
            
            // Save the joined document
            dstDoc.Save(outPut);

            //ExEnd:AppendDocumentManually
            Console.WriteLine("\nCreateWordPackage processing completed...");
        }
    }
}
