﻿using ProductExcelApp;
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
    }
}