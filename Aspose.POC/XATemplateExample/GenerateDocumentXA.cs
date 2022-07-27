using Aspose.Words;
using Aspose.Words.Reporting;
using System;
using System.Collections.Generic;

namespace Aspose.POC.XATemplateExample {
    internal class GenerateDocumentXA {
        public static void Run() {
            string[] templates = { "xa2.0_aspose_hybrid_template.docx", "AsposeTemplate_For_DS.docx", "book_mark-template.docx" };
            Console.WriteLine("\nGenerateDocument XA+Aspose template merging started...");

            // Retrieve data set schema and prepare data set.
            var dataSource = new TemplateDataSource();
            dataSource.FillSmartSet(Common.GetDataSetXml(Common.FilePath + "DataSet-NewDataSet.xml"));

            foreach(var template in templates) {
                BuildReport(template, dataSource);
            }

            Console.WriteLine("\nGenerateDocument XA+Aspose template merging completed...");
        }

        private static void BuildReport(string fileName, TemplateDataSource dataSource) {
            Console.WriteLine($"\n Merging - \"{fileName}\" Started...");
            
            // Load the template document.
            Document template = new Document(Common.FilePath + fileName);

            try {
                // Merge data
                ReportingEngine engine = new ReportingEngine();
                engine.Options = ReportBuildOptions.AllowMissingMembers;
                engine.BuildReport(template, dataSource, "NewDataSet");

                // Save the finished document to disk.
                var outputFile = Common.GetOutputFilePath(fileName);
                template.Save(outputFile);
            } catch(Exception) {

                throw;
            }

            Console.WriteLine($"\n Merging - \"{fileName}\" Completed...");
        }
    }
}
