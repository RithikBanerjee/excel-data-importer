using System;
using System.Collections.Generic;

namespace ExcelDataImporter.Model
{
    //demo data type for example
    public class DemoTable
    {
        public int Id { get; set; }
        public string Name { get; set; }
        public string Type { get; set; }
        public DateTime Date { get; set; }
        public string Category { get; set; }
        public DemoDetails Details { get; set; }
        public List<DemoItems> Items { get; set; }
    }
    //demo child data type with 1 to 1 relation
    public class DemoDetails
    {
        public string TradeName { get; set; }
        public string Address { get; set; }
        public string City { get; set; }
        public string State { get; set; }
        public string PhoneNumber { get; set; }
    }
    //demo child data type with 1 to many relation
    public class DemoItems
    {
        public string SerialNo { get; set; }
        public long Quantity { get; set; }
        public double UnitPrice { get; set; }
        public double GSTRate { get; set; }
        public double GSTAmount { get; set; }
        public double Amount { get; set; }
    }
}
