using ClosedXML.Excel;

namespace ShopTestApp.Services;

public class FileService
{
    private readonly List<Product> _products;
    private readonly List<Client> _clients;
    private readonly List<Order> _orders;

    public FileService()
    {
        _products = Storage.Products;
        _clients = Storage.Clients;
        _orders = Storage.Orders;
    }

    public void OpenFile()
    {
        Storage.Path = EnterAPathToFile();

        if (Storage.Path == "Выход")
        {
            Console.WriteLine();
            return;
        }

        try
        {
            var workbook = new XLWorkbook(Storage.Path);
            var worksheet = workbook.Worksheet("Товары");
            var rows = worksheet.RangeUsed().RowsUsed().Skip(1);
            ProductsImport(rows);

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
            Storage.Path = string.Empty;
            return;
        }

        CustomConsole.WriteLine("Файл успешно загружен!\n", ConsoleColor.Green);
    }

    private string EnterAPathToFile()
    {
        var isFinished = false;
        Storage.Path = string.Empty;

        while (!isFinished)
        {
            CustomConsole.Write("Введите путь до файла с данными: ", ConsoleColor.DarkGray);

            Storage.Path = CustomConsole.ReadLine();

            if (string.IsNullOrWhiteSpace(Storage.Path)) return "Выход";

            if (!IsRightPath(Storage.Path)) continue;

            isFinished = true;
        }

        return Storage.Path;
    }

    private bool IsRightPath(string? path)
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

    private void ProductsImport(IEnumerable<IXLRangeRow> rows)
    {
        foreach (var row in rows)
        {
            row.Cell(1).TryGetValue(out int code);
            row.Cell(4).TryGetValue(out decimal price);
            var name = row.Cell(2).GetText();
            var unit = row.Cell(3).GetText();

            _products.Add(new Product
            {
                Code = code,
                Name = name,
                Unit = unit,
                Price = price
            });
        }
    }

    private void ClientsExport(IEnumerable<IXLRangeRow> rows)
    {
        foreach (var row in rows)
        {
            row.Cell(1).TryGetValue(out int code);
            var name = row.Cell(2).GetText();
            var address = row.Cell(3).GetText();
            var contactPerson = row.Cell(4).GetText();

            _clients.Add(new Client
            {
                Code = code,
                Name = name,
                Address = address,
                ContactPerson = contactPerson
            });
        }
    }

    private void OrdersExport(IEnumerable<IXLRangeRow> rows)
    {
        foreach (var row in rows)
        {
            row.Cell(1).TryGetValue(out int code);
            row.Cell(2).TryGetValue(out int productCode);
            row.Cell(3).TryGetValue(out int clientCode);
            row.Cell(4).TryGetValue(out int orderNumber);
            row.Cell(5).TryGetValue(out int count);
            row.Cell(6).TryGetValue(out DateTime placementDate);

            _orders.Add(new Order
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

    public void ChangeContactPerson()
    {
        if (string.IsNullOrWhiteSpace(Storage.Path))
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

            var client = _clients.FirstOrDefault(x => x.Name == clientName);

            if (client == null)
            {
                CustomConsole.WriteLine("Организации с таким названием не существует!", ConsoleColor.Red);
                continue;
            }

            CustomConsole.Write("Введите ФИО нового контактного лица: ", ConsoleColor.DarkGray);
            var newContactPerson = CustomConsole.ReadLine();

            var workbook = new XLWorkbook(Storage.Path);
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
}