using ClosedXML.Excel;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Drawing.ChartDrawing;
using ShopTestApp;

var products = new List<Product>();
var clients = new List<Client>();
var orders = new List<Order>();
var path = string.Empty;

var commands = new List<Command>
{
    new(1, "Выбрать файл с данными", OpenFile),
    new(2, "Вывести информацию о клиентах, заказавших товар", PrintClientsInformationByProductName),
    new(3, "Изменить контактное лицо", ChangeContactPerson),
    new(4, "Золотой клиент", FindAGoldClient),
    new(5, "Выход", () => { })
};

while (true)
{
    ShowCommands(commands);

    var command = EnterACommand(commands);
    
    if (command == null)
    {
        CustomConsole.WriteLine("Введена неверная команда. Повторите попытку!\n", ConsoleColor.Red);
        continue;
    }

    if (command.Id == 5) return;

    command.Run();
}

void ShowCommands(List<Command> commandsList)
{
    foreach (var command in commandsList)
    {
        CustomConsole.WriteLine($"{command.Id}. {command.Name}");
    }
}

Command? EnterACommand(List<Command> commandsList)
{
    CustomConsole.Write("Введите номер операции: ", ConsoleColor.DarkGray);

    var wroteValue = CustomConsole.ReadLine();
    var isNumeric = int.TryParse(wroteValue, out var commandNumber);

    if (!isNumeric) return null;

    var command = commandsList.FirstOrDefault(x => x.Id == commandNumber);

    return command;
}

void OpenFile()
{
    path = EnterAPathToFile();

    if (path == "Выход")
    {
        Console.WriteLine();
        return;
    }

    try
    {
        var workbook = new XLWorkbook(path);
        var worksheet = workbook.Worksheet("Товары");
        var rows = worksheet.RangeUsed().RowsUsed().Skip(1);
        ProductsExport(rows);

        worksheet = workbook.Worksheet("Клиенты");
        rows = worksheet.RangeUsed().RowsUsed().Skip(1);
        ClientsExport(rows);

        worksheet = workbook.Worksheet("Заявки");
        rows = worksheet.RangeUsed().RowsUsed().Skip(1);
        OrdersExport(rows);
    }
    catch
    {
        CustomConsole.WriteLine("Данный файл не подходит для экспорта\n", ConsoleColor.Red);
        path = string.Empty;
        return;
    }

    CustomConsole.WriteLine("Файл успешно загружен!\n", ConsoleColor.Green);
}

string? EnterAPathToFile()
{
    var isFinished = false;
    path = string.Empty;

    while (!isFinished)
    {
        CustomConsole.Write("Введите путь до файла с данными: ", ConsoleColor.DarkGray);

        path = CustomConsole.ReadLine();

        if (string.IsNullOrWhiteSpace(path)) return "Выход";

        if (!IsRightPath(path)) continue;

        isFinished = true;
    }

    return path;
}

bool IsRightPath(string? path)
{
    if (path == null || !File.Exists(path))
    {
        CustomConsole.WriteLine("Неверный путь!", ConsoleColor.Red);
        return false;
    }

    var fileInfo = new FileInfo(path);
    if (fileInfo.Extension.ToUpper() != ".XLSX")
    {
        CustomConsole.WriteLine("Неверный тип файла!", ConsoleColor.Red);
        return false;
    }

    return true;
}

void ProductsExport(IEnumerable<IXLRangeRow> rows)
{
    foreach (var row in rows)
    {
        row.Cell(1).TryGetValue(out int code);
        row.Cell(4).TryGetValue(out decimal price);
        var name = row.Cell(2).GetText();
        var unit = row.Cell(3).GetText();

        products.Add(new Product
        {
            Code = code,
            Name = name,
            Unit = unit,
            Price = price
        });
    }
}

void ClientsExport(IEnumerable<IXLRangeRow> rows)
{
    foreach (var row in rows)
    {
        row.Cell(1).TryGetValue(out int code);
        var name = row.Cell(2).GetText();
        var address = row.Cell(3).GetText();
        var contactPerson = row.Cell(4).GetText();

        clients.Add(new Client
        {
            Code = code,
            Name = name,
            Address = address,
            ContactPerson = contactPerson
        });
    }
}

void OrdersExport(IEnumerable<IXLRangeRow> rows)
{
    foreach (var row in rows)
    {
        row.Cell(1).TryGetValue(out int code);
        row.Cell(2).TryGetValue(out int productCode);
        row.Cell(3).TryGetValue(out int clientCode);
        row.Cell(4).TryGetValue(out int orderNumber);
        row.Cell(5).TryGetValue(out int count);
        row.Cell(6).TryGetValue(out DateTime placementDate);

        orders.Add(new Order
        {
            Code = code,
            ProductCode = productCode,
            ClientCode = clientCode,
            OrderNumber = orderNumber,
            Count = count,
            PlacementDate = placementDate
        });
    }
}

void PrintClientsInformationByProductName()
{
    if (string.IsNullOrWhiteSpace(path))
    {
        CustomConsole.WriteLine("Источник данных не обнаружен\n", ConsoleColor.Red);
        return;
    }

    var isFinished = false;
    while (!isFinished)
    {
        CustomConsole.Write("Введите наименование товара: ", ConsoleColor.DarkGray);
        var nameOfProduct = CustomConsole.ReadLine();

        if (string.IsNullOrWhiteSpace(nameOfProduct))
        {
            Console.WriteLine();
            break;
        }

        var product = products.FirstOrDefault(x => x.Name.ToLower() == nameOfProduct?.ToLower());

        if (product == null)
        {
            CustomConsole.WriteLine("Товара с таким названием не существует!", ConsoleColor.Red);
            return;
        }

        var productOrders = orders.Where(x => x.ProductCode == product.Code);

        var productOrderEntries = productOrders.Join(clients,
                po => po.ClientCode,
                c => c.Code,
            (po, c) => new ProductOrderEntry(c.Name, po.Count, product.Price, po.PlacementDate));

        PrintProductOrders(productOrderEntries);
        isFinished = true;
    }

}

void ChangeContactPerson()
{
    if (string.IsNullOrWhiteSpace(path))
    {
        CustomConsole.WriteLine("Источник данных не обнаружен\n", ConsoleColor.Red);
        return;
    }

    var isFinished = false;

    while (!isFinished)
    {
        CustomConsole.Write("Введите название организации: ", ConsoleColor.DarkGray);
        var clientName = CustomConsole.ReadLine();

        if (string.IsNullOrWhiteSpace(clientName))
        {
            Console.WriteLine();
            return;
        }

        var client = clients.FirstOrDefault(x => x.Name == clientName);

        if (client == null)
        {
            CustomConsole.WriteLine("Организации с таким названием не существует!", ConsoleColor.Red);
            continue;
        }

        CustomConsole.Write("Введите ФИО нового контактного лица: ", ConsoleColor.DarkGray);
        var newContactPerson = CustomConsole.ReadLine();

        var workbook = new XLWorkbook(path);
        var worksheet = workbook.Worksheet("Клиенты");
        var rows = worksheet.RangeUsed().RowsUsed().Skip(1);

        var row = rows.FirstOrDefault(x => x.Cell(2).GetText() == clientName);

        if (row == null)
        {
            CustomConsole.WriteLine("Что-то пошло не так", ConsoleColor.Red);
            return;
        }

        try
        {
            row.Cell(4).Value = newContactPerson;
            workbook.Save();

            client.ContactPerson = newContactPerson;
        }
        catch
        {
            CustomConsole.WriteLine("При сохранении данных произошла ошибка", ConsoleColor.Red);
            throw;
        }

        CustomConsole.WriteLine($"Изменение контактного лица для организации {clientName} выполнено успешно!\n", ConsoleColor.Green);
        isFinished = true;
    }
}

void FindAGoldClient()
{
    if (string.IsNullOrWhiteSpace(path))
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
        var name = clients.FirstOrDefault(x => x.Code == client.ClientCode)?.Name;
        CustomConsole.WriteLine(name, ConsoleColor.Cyan);
    }

    Console.WriteLine();
}

void PrintProductOrders(IEnumerable<ProductOrderEntry> enumerable)
{
    if (string.IsNullOrWhiteSpace(path))
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

List<OrderClient>? GetClientWithMaxOrders(DateTime startDate, DateTime endDate)
{
    var ordersByDate = orders.Where(x => x.PlacementDate >= startDate && x.PlacementDate <= endDate).ToList();

    if (!ordersByDate.Any()) return null;

    var ordersClientsCodes = ordersByDate.DistinctBy(x => x.ClientCode).Select(x => x.ClientCode);
    var orderClients = ordersClientsCodes.Select(x => new OrderClient(x, ordersByDate.Count(order => order.ClientCode == x)))
        .ToList();

    if (!orderClients.Any()) return null;

    var maxOrderCount = orderClients.Max(x => x.CountOrders);

    var maxOrderCountClients = orderClients.Where(x => x.CountOrders == maxOrderCount).ToList();
    return maxOrderCountClients;
}

public record OrderClient(int ClientCode, int CountOrders);