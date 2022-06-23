using Aspose.Words;
using Aspose.Words.Reporting;
using System;

namespace Aspose.POC {
    internal class GenerateDocument {
        public static void Run() {
            Console.WriteLine("\nGenerateDocument processing started...");

            string fileName = "IGTD New equipment and disposable.docx";
            // Load the template document.
            Document doc = new Document(Common.FilePath + fileName);

            var jsonStream = Common.PreProcessMergeData(Common.FilePath + "MergeDataPhilips.json");
            JsonDataSource jsonDataSource = new JsonDataSource(jsonStream);

            // Create a Reporting Engine.
            ReportingEngine engine = new ReportingEngine();

            // Execute the build report.
            engine.BuildReport(doc, jsonDataSource);

            var dataDir = Common.GetOutputFilePath(fileName);

            // Save the finished document to disk.
            doc.Save(dataDir);

            Console.WriteLine("\nGenerateDocument processing completed...");
        }        
    }
}
