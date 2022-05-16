using Newtonsoft.Json;
using System;
using System.Collections.Generic;
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

        public static string GetOutputFilePath(String inputFilePath) {
            string extension = Path.GetExtension(inputFilePath);
            string filename = Path.GetFileNameWithoutExtension(inputFilePath);
            return path + filename + "_out_" + extension;
        }

        public static MemoryStream PreProcessMergeData(string jsonFilePath) {
            string json = File.ReadAllText(jsonFilePath);
            dynamic jsonObj = JsonConvert.DeserializeObject(json);

            var clauseId = jsonObj["MergeData"]["Clause"]["ClauseId"];
            jsonObj["MergeData"]["Clause"]["ClauseId"] = Common.FilePath + clauseId;

            var clauseIdDar = jsonObj["MergeData"]["DAR"]["ClauseId"];
            jsonObj["MergeData"]["DAR"]["ClauseId"] = Common.FilePath + clauseIdDar;

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
    }
}
