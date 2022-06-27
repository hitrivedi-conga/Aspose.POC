using Aspose.Words;
using Aspose.Words.Reporting;
using System;

namespace Aspose.POC {
    internal class Program {

        static void Main(string[] args) {
            Console.WriteLine("Main program started...");

            Console.WriteLine("=====================================================");
            //GenerateDocument.Run();
            ConvertDocxToPDF.Run();
            //CreateWordPackage.Run();
            //Compare.Run();
            //PdfSecurity.Run();
            Console.WriteLine("=====================================================");

            Console.WriteLine("\n\nMain program finished. Press any key to exit....");
            Console.ReadKey();
        }        

        static void HelloWorld() {
            Console.WriteLine("\n\nHelloWorld processing started...");            

            string fileName = "HelloWorld.docx";
            // Load the template document.
            Document doc = new Document(Common.FilePath + fileName);

            // Create an instance of sender class to set it's properties.
            Sender sender = new Sender {
                To = "Xyz",
                FromName = "Hiren Trivedi",
                Subject = "LINQ Reporting Engine",
                Message = "God is great!!"
            };

            // Create a Reporting Engine.
            ReportingEngine engine = new ReportingEngine();

            // Execute the build report.
            engine.BuildReport(doc, sender, "sender");

            var dataDir = Common.GetOutputFilePath(fileName);

            // Save the finished document to disk.
            doc.Save(dataDir);

            Console.WriteLine("\nHelloWorld processing completed...");
            Console.WriteLine("=====================================================");
        }

        static void ComplexRepeat() {
            Console.WriteLine("\n\nComplexRepeat processing started...");

            string fileName = "ComplexRepeat.docx";
            // Load the template document.
            Document doc = new Document(Common.FilePath + fileName);

            // Create a Reporting Engine.
            ReportingEngine engine = new ReportingEngine();

            // Execute the build report.
            engine.BuildReport(doc, Common.GetMyOrders(), "orders");

            var dataDir = Common.GetOutputFilePath(fileName);

            // Save the finished document to disk.
            doc.Save(dataDir);

            Console.WriteLine("\nComplexRepeat processing completed...");
            Console.WriteLine("=====================================================");
        }

        static void InsertOLEObject() {
            Console.WriteLine("\nInsertOLEObject processing started...");

            var oleFilePath = Common.FilePath + "RLP.pptx";
            string fileName = "Template.docx";
            
            // Load the template document.
            Document doc = new Document(Common.FilePath + fileName);

            var jsonStream = Common.PreProcessMergeData(Common.FilePath + "MergeData.json");
            JsonDataSource jsonDataSource = new JsonDataSource(jsonStream);

            // Create a Reporting Engine.
            ReportingEngine engine = new ReportingEngine();

            // Execute the build report.
            engine.BuildReport(doc, jsonDataSource);

            DocumentBuilder documentBuilder = new DocumentBuilder(doc);
            documentBuilder.InsertOleObject(oleFilePath, "RLPPresentation", true, false, null);

            var dataDir = Common.GetOutputFilePath(fileName);

            // Save the finished document to disk.
            doc.Save(dataDir);

            Console.WriteLine("\nInsertOLEObject processing completed...");
        }        
    }
}
