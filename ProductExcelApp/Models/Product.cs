namespace ProductExcelApp.Models
{
    public class Product
    {
        public int Id { get; set; }
        public string Name { get; set; }
        public ProductType ProductType { get; set; }
        public double Price { get; set; }
    }

    public enum ProductType
    {
        Kilo,
        Unit,
        Litre
    }
}