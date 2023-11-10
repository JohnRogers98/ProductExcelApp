namespace ProductExcelApp.Models
{
    public class Order
    {
        public int Id { get; set; }

        public int ProductId { get; set; }
        public Product Product { get; set; }

        public int ClientId { get; set; }
        public Client Client { get; set; }

        public int NumberOfOrder { get; set; }

        public int Count { get; set; }

        public DateOnly DeliveryDate { get; set; }
    }
}
