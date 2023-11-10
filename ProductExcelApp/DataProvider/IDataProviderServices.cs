using ProductExcelApp.Models;

namespace ProductExcelApp.DataProvider
{
    public interface IDataProviderServices
    {
        IEnumerable<Product> GetProducts();
        IEnumerable<Client> GetClients();
        IEnumerable<Order> GetOrders();

        Boolean ChangeContactPersonByClientId(Int32 clientId, String newContactPerson);
    }
}
