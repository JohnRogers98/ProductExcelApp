﻿using ProductExcelApp.DataProvider;
using ProductExcelApp.Models;

namespace ProductExcelApp
{
    public class Repository
    {
        private IDataProviderServices _dataProvider;

        private Dictionary<Int32, Product> _productDictionary = new Dictionary<Int32, Product>();

        private Dictionary<Int32, Client> _clientDictionary = new Dictionary<Int32, Client>();

        private Dictionary<Int32, Order> _orderDictionary = new Dictionary<Int32, Order>();

        private Repository(IDataProviderServices dataProvider)
        {
            _dataProvider = dataProvider;

            SetClientDictionary();
            SetProductDictionary();
            SetOrderDictionary();
        }

        public static Repository CreateFilledRepository(IDataProviderServices dataProvider)
        {
            return new Repository(dataProvider);
        }

        public Boolean ChangeClientContactPerson(String nameOfClient, String newContactPerson)
        {
            var client = _clientDictionary.Values.Where(client => client.Name == nameOfClient).FirstOrDefault();
            if(client == null)
            {
                return false;
            }

            client.ContactPerson = newContactPerson;
            var status = _dataProvider.ChangeContactPersonByClientId(client.Id, client.ContactPerson);
            return status;
        }

        public IEnumerable<Object> GetProductInfoByName(String productName)
        {
            var productId = _productDictionary.Values.First(product => product.Name == productName).Id;

            return _orderDictionary.Values.Where(order => order.ProductId == productId)
                .Select(order => new { order.Client.Name, order.Count, order.DeliveryDate });
        }

        public Client GetGoldClient(Int32 year)
        {
            var from =  new DateOnly(year, 1, 1);
            var to =  new DateOnly(year, 12, 31);

            return GetGoldClient(from, to);
        }
        public Client GetGoldClient(Int32 year, Int32 month)
        {
            var from = new DateOnly(year, month, 1);
            var to = new DateOnly(year, month, DateTime.DaysInMonth(year, month));

            return GetGoldClient(from, to);
        }
        private Client? GetGoldClient(DateOnly from, DateOnly to)
        {
            var ordersInDateRange = _orderDictionary.Values.Where(order => order.DeliveryDate >= from && order.DeliveryDate <= to);

            var ordersGroupByClientId = ordersInDateRange.GroupBy(order => order.ClientId);

            var groupOfOrdersMaxCount = ordersGroupByClientId.OrderByDescending(x => x.Count()).FirstOrDefault();

            var goldClient = groupOfOrdersMaxCount?.First().Client;

            return goldClient;
        }
        
        private void SetProductDictionary()
        {
            var products = _dataProvider.GetProducts();

            foreach(var product in products)
            {
                var isAdded = _productDictionary.TryAdd(product.Id, product);
                if(!isAdded)
                {
                    throw new ArgumentException();
                }
            }
        }

        private void SetClientDictionary()
        {
            var clients = _dataProvider.GetClients();

            foreach (var client in clients)
            {
                var isAdded = _clientDictionary.TryAdd(client.Id, client);
                if (!isAdded)
                {
                    throw new ArgumentException();
                }
            }
        }

        /// <summary>
        /// Sets orders values from data provider with filling empty product and client references
        /// </summary>
        /// <exception cref="ArgumentException"></exception>
        private void SetOrderDictionary()
        {
            var orders = _dataProvider.GetOrders();

            foreach (var order in orders)
            {
                SetOrder_ProductLink(in order);
                SetOrder_ClientLink(in order);

                var isAdded = _orderDictionary.TryAdd(order.Id, order);
                if (!isAdded)
                {
                    throw new ArgumentException();
                }
            }
        }
        private void SetOrder_ProductLink(in Order order)
        {
            Product p;
            Boolean existProduct = _productDictionary.TryGetValue(order.ProductId, out p);
            if (existProduct)
            {
                order.Product = p;
            }
            else
            {
                //
            }
        }
        private void SetOrder_ClientLink(in Order order)
        {
            Client c;
            Boolean existClient = _clientDictionary.TryGetValue(order.ClientId, out c);
            if (existClient)
            {
                order.Client = c;
            }
            else
            {
                //
            }
        }

    }
}