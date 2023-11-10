using ProductExcelApp;
using ProductExcelApp.DataProvider;
using ProductExcelApp.Resources;
using System.Text;

public class Program
{
    private static void Main(string[] args)
    {
        Console.OutputEncoding = Encoding.UTF8;

        var path = Path.Combine(StaticResources.RootPath, "Products.xlsx");

        IDataProviderServices dataProvider = ExcelContext.CreateExcelContext(path);

        var repository = Repository.CreateFilledRepository(dataProvider);

        var productInfo = repository.GetProductInfoByName("Сыр");

        foreach(var product in productInfo)
        {
            Console.WriteLine(product);
        }

        var goldClient = repository.GetGoldClient(2023, 1);
        Console.WriteLine(goldClient.Name);
    }
}