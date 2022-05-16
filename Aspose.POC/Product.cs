using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Aspose.POC {
    public class Order {
        public string OrderID { get; set; }
        public string BillTo { get; set; }
        public DateTime OrderDate { get; set; }
        public Product[] Products { get; set; }        
    }
    public class Product {
        public string Name { get; set; }
        public Category Category { get; set; }
        public double UnitPrice { get; set; }
        public int Quantity { get; set; }
    }

    public class Category {
        public string Name { get; set; }
        public double Discount { get; set; }
    }
}
