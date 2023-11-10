using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using ProductExcelApp.Models;
using ProductExcelApp.Resources;
using System.Data;

namespace ProductExcelApp.DataProvider
{
    public class ExcelContext : IDataProviderServices
    {
        private string _pathToFile;

        private ExcelContext(string pathToFile)
        {
            _pathToFile = pathToFile;
        }

        public static ExcelContext CreateExcelContext(string pathToFile)
        {
            if (File.Exists(pathToFile))
                return new ExcelContext(pathToFile);
            return null;
        }

        #region Actions

        public IEnumerable<Product> GetProducts()
        {
            DataTable productTable = ReadSheetOfName(StaticResources.Product_Rus);

            var productList = productTable.AsEnumerable().Select(
                x => new Product
                {
                    Id = Convert.ToInt32(x[0]),
                    Name = (string)x[1],
                    ProductType = GetProductTypeFromRus((string)x[2]),
                    Price = Convert.ToDouble(x[3])
                });
            return productList;
        }
        private ProductType GetProductTypeFromRus(string rusName)
        {
            return rusName switch
            {
                "Килограмм" => ProductType.Kilo,
                "Литр" => ProductType.Kilo,
                "Штука" => ProductType.Unit,
                _ => throw new ArgumentException()
            };
        }

        public IEnumerable<Client> GetClients()
        {
            DataTable clientTable = ReadSheetOfName(StaticResources.Client_Rus);

            var clientList = clientTable.AsEnumerable().Select(
                x => new Client
                {
                    Id = Convert.ToInt32(x[0]),
                    Name = (string)x[1],
                    Adress = (string)x[2],
                    ContactPerson = (string)x[3]
                });
            return clientList;
        }

        public IEnumerable<Order> GetOrders()
        {
            DataTable orderTable = ReadSheetOfName(StaticResources.Order_Rus);

            var orderList = orderTable.AsEnumerable().Select(
                x => new Order
                {
                    Id = Convert.ToInt32(x[0]),
                    ProductId = Convert.ToInt32(x[1]),
                    ClientId = Convert.ToInt32(x[2]),
                    NumberOfOrder = Convert.ToInt32(x[3]),
                    Count = Convert.ToInt32(x[4]),
                    DeliveryDate = GetDateFromOADaysNumber(Convert.ToInt32(x[5]))
                });
            return orderList;
        }
        /// <summary>
        /// Excell uses number of days from O.A.
        /// </summary>
        /// <returns></returns>
        private DateOnly GetDateFromOADaysNumber(int oaNumber)
        {
            return DateOnly.FromDateTime(DateTime.FromOADate(oaNumber));
        }

        #endregion Acrtions

        private DataTable ReadSheetOfName(string nameOfSheet)
        {
            DataTable table = new DataTable();

            using var doc = SpreadsheetDocument.Open(_pathToFile, false);

            WorkbookPart workbookPart = doc.WorkbookPart;
            Sheets sheetCollection = workbookPart.Workbook.GetFirstChild<Sheets>();

            var sheet = sheetCollection.OfType<Sheet>().First(sh => sh.Name == nameOfSheet);

            Worksheet worksheet = ((WorksheetPart)workbookPart.GetPartById(sheet.Id)).Worksheet;
            SheetData sheetData = worksheet.GetFirstChild<SheetData>();

            for (int row = 0; row < sheetData.ChildElements.Count(); row++)
            {
                var columnOfRowList = new List<string>();

                for (int column = 0; column < sheetData.ElementAt(row).ChildElements.Count(); column++)
                {
                    Cell currentCell = (Cell)sheetData.ElementAt(row).ChildElements.ElementAt(column);
                    string currentValue = string.Empty;

                    if (currentCell.CellValue == null)
                    {
                        return table;
                    }

                    if (currentCell.DataType != null)
                    {
                        if (currentCell.DataType == CellValues.SharedString)
                        {
                            if (int.TryParse(currentCell.InnerText, out int id))
                            {
                                SharedStringItem item = workbookPart.SharedStringTablePart.SharedStringTable.Elements<SharedStringItem>().ElementAt(id);
                                if (item.Text != null)
                                {
                                    if (row == 0)
                                    {
                                        table.Columns.Add(item.Text.Text);
                                    }
                                    else
                                    {
                                        columnOfRowList.Add(item.Text.Text);
                                    }
                                }
                            }
                        }
                    }
                    else
                    {
                        if (row != 0)
                        {
                            columnOfRowList.Add(currentCell.InnerText);
                        }
                    }
                }
                if (row != 0)
                    table.Rows.Add(columnOfRowList.ToArray());
            }

            return table;
        }
    }
}
