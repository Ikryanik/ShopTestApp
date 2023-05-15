using ClosedXML.Excel;

namespace ShopTestApp.Services;

public class ClientService
{
    private readonly List<Product> _products;
    private readonly List<Client> _clients;
    private readonly List<Order> _orders;

    public ClientService()
    {
        _products = Storage.Products;
        _clients = Storage.Clients;
        _orders = Storage.Orders;
    }

    public void PrintClientsInformationByProductName()
    {
        if (string.IsNullOrWhiteSpace(Storage.Path))
        {
            CustomConsole.WriteLine("Источник данных не обнаружен\n", ConsoleColor.Red);
            return;
        }

        var isFinished = false;
        while (!isFinished)
        {
            var product = EnterAProductName(_products);

            if (product == null)
            {
                CustomConsole.WriteLine("Товара с таким названием не существует!", ConsoleColor.Red);
                return;
            }

            var productOrders = _orders.Where(x => x.ProductCode == product.Code);

            var productOrderEntries = productOrders.Join(_clients,
                po => po.ClientCode,
                c => c.Code,
                (po, c) => new ProductOrderEntry(c.Name, po.Count, product.Price, po.PlacementDate));

            PrintProductOrders(productOrderEntries);
            isFinished = true;
        }
    }

    private void PrintProductOrders(IEnumerable<ProductOrderEntry> enumerable)
    {
        if (string.IsNullOrWhiteSpace(Storage.Path))
        {
            CustomConsole.WriteLine("Источник данных не обнаружен\n", ConsoleColor.Red);
            return;
        }

        CustomConsole.WriteLine(
            $"|{"Название организации",25}|{"Количество товара",20}|{"Цена",10}|{"Дата заказа",15}|",
            ConsoleColor.Blue);

        foreach (var order in enumerable)
        {
            CustomConsole.WriteLine(
                $"|{order.Name,25}|{order.Count,20}|{order.Price,10: 0.00}|{order.PlacementDate,15: d}|",
                ConsoleColor.Cyan);
        }

        Console.WriteLine();
    }

    private Product? EnterAProductName(List<Product> list)
    {
        CustomConsole.Write("Введите наименование товара: ", ConsoleColor.DarkGray);
        var nameOfProduct = CustomConsole.ReadLine();

        if (string.IsNullOrWhiteSpace(nameOfProduct))
        {
            Console.WriteLine();
            return null;
        }

        var product1 = list.FirstOrDefault(x => x.Name.ToLower() == nameOfProduct.ToLower());
        return product1;
    }

    public void FindAGoldClient()
    {
        if (string.IsNullOrWhiteSpace(Storage.Path))
        {
            CustomConsole.WriteLine("Источник данных не обнаружен\n", ConsoleColor.Red);
            return;
        }

        var periodString = string.Empty;
        var isCommand = false;

        var startDate = new DateTime();
        var endDate = new DateTime();

        while (!isCommand)
        {
            CustomConsole.Write("Введите год (текущий - '-'): ", ConsoleColor.DarkGray);
            var wroteYearValue = CustomConsole.ReadLine();

            CustomConsole.Write("Введите месяц в численном формате (текущий - '-', без месяца - ''): ", ConsoleColor.DarkGray);
            var wroteMonthValue = CustomConsole.ReadLine();

            if (wroteYearValue == "-")
            {
                wroteYearValue = DateTime.Now.Year.ToString();
            }

            if (wroteMonthValue == "-")
            {
                wroteMonthValue = DateTime.Now.Month.ToString();
            }

            var isYear = int.TryParse(wroteYearValue, out var year);
            if (!isYear || year < 1900 || year > 2200)
            {
                CustomConsole.WriteLine("Неверно введен год");
                continue;
            }

            var month = 0;

            if (!string.IsNullOrWhiteSpace(wroteMonthValue))
            {
                var isMonth = int.TryParse(wroteMonthValue, out month);
                if (!isMonth || month < 1 || month > 12)
                {
                    CustomConsole.WriteLine("Неверно введен месяц");
                    continue;
                }
            }

            if (month == 0)
            {
                startDate = new DateTime(year, 1, 1);
                endDate = new DateTime(year, 12, 31);

                periodString = $"за {year} г.";
            }
            else
            {
                startDate = new DateTime(year, month, 1);
                var day = DateTime.DaysInMonth(year, month);
                endDate = new DateTime(year, month, day);

                periodString = $" за {month}.{year}";
            }

            isCommand = true;
        }

        var maxOrderCountClients = GetClientWithMaxOrders(startDate, endDate);
        if (maxOrderCountClients == null)
        {
            CustomConsole.WriteLine($"Заказы за {periodString} не были совершены\n", ConsoleColor.Blue);
            return;
        }

        var maxOrderCount = maxOrderCountClients.First().CountOrders;

        CustomConsole.WriteLine($"\nСПИСОК ЗОЛОТЫХ КЛИЕНТОВ {periodString}", ConsoleColor.DarkBlue);
        CustomConsole.WriteLine($"Количество заказов: {maxOrderCount}", ConsoleColor.Blue);
        foreach (var client in maxOrderCountClients)
        {
            var name = _clients.FirstOrDefault(x => x.Code == client.ClientCode)?.Name;
            CustomConsole.WriteLine(name, ConsoleColor.Cyan);
        }

        Console.WriteLine();
    }

    private List<OrderClient>? GetClientWithMaxOrders(DateTime startDate, DateTime endDate)
    {
        var ordersByDate = _orders.Where(x => x.PlacementDate >= startDate && x.PlacementDate <= endDate).ToList();

        if (!ordersByDate.Any()) return null;

        var ordersClientsCodes = ordersByDate.DistinctBy(x => x.ClientCode).Select(x => x.ClientCode);
        var orderClients = ordersClientsCodes.Select(x => new OrderClient(x, ordersByDate.Count(order => order.ClientCode == x)))
            .ToList();

        if (!orderClients.Any()) return null;

        var maxOrderCount = orderClients.Max(x => x.CountOrders);

        var maxOrderCountClients = orderClients.Where(x => x.CountOrders == maxOrderCount).ToList();
        return maxOrderCountClients;
    }
}