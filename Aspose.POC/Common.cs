using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Reflection;

namespace Aspose.POC {
    public class Common {
        private static readonly string path;
        public static string FilePath { get { return path; } }

        static Common() {
            path = GetFilesDirectoryPath();
        }

        private static string GetFilesDirectoryPath() {
            string assemblyLocation = Assembly.GetExecutingAssembly().Location;
            var startIndex = assemblyLocation.IndexOf("bin");
            var substr = assemblyLocation.Substring(startIndex);
            return assemblyLocation.Replace(substr, string.Empty) + "Files\\";
        }

        public static string GetOutputFilePath(String inputFilePath, string extension = null) {
            if (string.IsNullOrEmpty(extension))
            {
                  extension = Path.GetExtension(inputFilePath);
            }
            string filename = Path.GetFileNameWithoutExtension(inputFilePath);
            return path + filename + "_out_" + extension;
        }

        public static MemoryStream PreProcessMergeData(string jsonFilePath) {
            string json = File.ReadAllText(jsonFilePath);
            dynamic jsonObj = JsonConvert.DeserializeObject(json);

            //MasterAsposeTemplate
            //var clauseId = jsonObj["MergeData"]["Clause"]["ClauseId"];
            //jsonObj["MergeData"]["Clause"]["ClauseId"] = Common.FilePath + clauseId;

            //var clauseIdDar = jsonObj["MergeData"]["DAR"]["ClauseId"];
            //jsonObj["MergeData"]["DAR"]["ClauseId"] = Common.FilePath + clauseIdDar;

            //MSA Template
            //var clauseIdDar = jsonObj["MergeDataMSA"]["DAR"]["ClauseId"];
            //jsonObj["MergeDataMSA"]["DAR"]["ClauseId"] = Common.FilePath + clauseIdDar;

            //var clauseId = jsonObj["MergeDataMSA"]["Clause"]["ClauseId"];
            //jsonObj["MergeDataMSA"]["Clause"]["ClauseId"] = Common.FilePath + clauseId;

            ////var clauseDar = jsonObj["MergeDataMSA"]["DA"]["ClauseId"];
            //jsonObj["MergeDataMSA"]["DA"]["ClauseId"] = Common.FilePath + clauseDar;

            //IGTD NEW Equipment and disposable Template
            var clauseId = jsonObj["MergeDataPhilips"]["Clause"]["ClauseId"];
            jsonObj["MergeDataPhilips"]["Clause"]["ClauseId"] = Common.FilePath + clauseId;

            var clauseIdDar = jsonObj["MergeDataPhilips"]["DAR"]["ClauseId"];
            jsonObj["MergeDataPhilips"]["DAR"]["ClauseId"] = Common.FilePath + clauseIdDar;

            var disposable = jsonObj["MergeDataPhilips"]["Disposable"]["ClauseId"];
            jsonObj["MergeDataPhilips"]["Disposable"]["ClauseId"] = Common.FilePath + disposable;

            var equipment = jsonObj["MergeDataPhilips"]["Equipment"]["ClauseId"];
            jsonObj["MergeDataPhilips"]["Equipment"]["ClauseId"] = Common.FilePath + equipment;

            var returns = jsonObj["MergeDataPhilips"]["Returns"]["ClauseId"];
            jsonObj["MergeDataPhilips"]["Returns"]["ClauseId"] = Common.FilePath + returns;

            var returns2 = jsonObj["MergeDataPhilips"]["Returns2"]["ClauseId"];
            jsonObj["MergeDataPhilips"]["Returns2"]["ClauseId"] = Common.FilePath + returns2;

            var returns3 = jsonObj["MergeDataPhilips"]["Returns3"]["ClauseId"];
            jsonObj["MergeDataPhilips"]["Returns3"]["ClauseId"] = Common.FilePath + returns3;

            var returns4 = jsonObj["MergeDataPhilips"]["Returns4"]["ClauseId"];
            jsonObj["MergeDataPhilips"]["Returns4"]["ClauseId"] = Common.FilePath + returns4;

            var returns5 = jsonObj["MergeDataPhilips"]["Returns5"]["ClauseId"];
            jsonObj["MergeDataPhilips"]["Returns5"]["ClauseId"] = Common.FilePath + returns5;

            var Notices = jsonObj["MergeDataPhilips"]["Notices"]["ClauseId"];
            jsonObj["MergeDataPhilips"]["Notices"]["ClauseId"] = Common.FilePath + Notices;

            var Confident = jsonObj["MergeDataPhilips"]["Confident"]["ClauseId"];
            jsonObj["MergeDataPhilips"]["Confident"]["ClauseId"] = Common.FilePath + Confident;

            var Idemnity = jsonObj["MergeDataPhilips"]["Idemnity"]["ClauseId"];
            jsonObj["MergeDataPhilips"]["Idemnity"]["ClauseId"] = Common.FilePath + Idemnity;

            var Limitation = jsonObj["MergeDataPhilips"]["Limitation"]["ClauseId"];
            jsonObj["MergeDataPhilips"]["Limitation"]["ClauseId"] = Common.FilePath + Limitation;

            var output = JsonConvert.SerializeObject(jsonObj, Formatting.Indented);

            var stream = new MemoryStream();
            var writer = new StreamWriter(stream);
            writer.Write(output);
            writer.Flush();
            stream.Position = 0;
            return stream;
        }

        public static IEnumerable<Order> GetMyOrders() {
            Order order = new Order() { OrderID= "O0016", BillTo = "John Smith", OrderDate = new DateTime(2021, 03, 01) };
            order.Products = GetProducts().Take(2).ToArray();
            yield return order;

            order = new Order() { OrderID = "O0024", BillTo = "John Smith", OrderDate = new DateTime(2022, 01, 15) };
            order.Products = GetProducts().Skip(4).ToArray();
            yield return order;
        }

        public static IEnumerable<Product> GetProducts() {
            Product product = new Product() { Name = "Composer", UnitPrice = 100, Quantity = 2 };
            product.Category = new Category() { Name = "Document Processing", Discount = 5 };
            yield return product;

            product = new Product() { Name = "MergeService", UnitPrice = 550, Quantity = 1 };
            product.Category = new Category() { Name = "Document Processing", Discount = 5 };
            yield return product;

            product = new Product() { Name = "CLM System", UnitPrice = 1050, Quantity = 4 };
            product.Category = new Category() { Name = "CRM", Discount = 8 };
            yield return product;

            product = new Product() { Name = "CPQ System", UnitPrice = 2050, Quantity = 2 };
            product.Category = new Category() { Name = "CRM", Discount = 8 };
            yield return product;

            product = new Product() { Name = "CFS System", UnitPrice = 780, Quantity = 1 };
            product.Category = new Category() { Name = "CRM", Discount = 10 };
            yield return product;

            product = new Product() { Name = "XAutor", UnitPrice = 500, Quantity = 1 };
            product.Category = new Category() { Name = "Template Builder", Discount = 10 };
            yield return product;
        }

        public static string GetDataSetXml(string dataSetSchemaPath) {
            return File.ReadAllText(dataSetSchemaPath);
        }

        public static DataSet PrepareDataSet() {
            var dataSet = new DataSet();
            var agreementTable = new DataTable("Apttus__APTS_Agreement__c");
            dataSet.Tables.Add(agreementTable);

            var agreementIdColumn = new DataColumn("apts_agreement_id", typeof(string));
            agreementIdColumn.Unique = true;
            agreementTable.Columns.Add(agreementIdColumn);

            var agreementNameColumn = new DataColumn("apts_agreement_name", typeof(string));
            agreementTable.Columns.Add(agreementNameColumn);

            var agreementStartDateColumn = new DataColumn("apts_agreement_contract_start_date", typeof(DateTime));
            agreementTable.Columns.Add(agreementStartDateColumn);

            var agreementEndDateColumn = new DataColumn("apts_agreement_contract_end_date", typeof(DateTime));
            agreementTable.Columns.Add(agreementEndDateColumn);

            var agreementNumberColumn = new DataColumn("apts_agreement_ff_agreement_number", typeof(string));
            agreementTable.Columns.Add(agreementNumberColumn);

            var agreementContractValueColumn = new DataColumn("apts_agreement_total_contract_value", typeof(Decimal));
            agreementTable.Columns.Add(agreementContractValueColumn);

            var agreementTermMonthsColumn = new DataColumn("apts_agreement_term_months", typeof(int));
            agreementTable.Columns.Add(agreementTermMonthsColumn);

            var agreementAutoRenewableColumn = new DataColumn("apts_agreement_auto_renewal", typeof(bool));
            agreementTable.Columns.Add(agreementAutoRenewableColumn);

            agreementTable.Rows.Add("123456", "Test Agreement", DateTime.Now, DateTime.Now.AddYears(5), "A123456", 1250000, 24, true);

            return dataSet;
        }
    }
}
