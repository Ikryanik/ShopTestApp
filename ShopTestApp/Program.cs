using ShopTestApp;

while (true)
{
    ShowMenu();

    CustomConsole.Write("Введите номер операции: ", ConsoleColor.DarkGray);
    
    var isNumeric = int.TryParse(CustomConsole.ReadLine(ConsoleColor.White), out var commandNumber);

    if (!isNumeric)
    {
        CustomConsole.WriteLine("Введена неверная команда. Повторите попытку!\n", ConsoleColor.Red);
        continue;
    }

    switch (commandNumber)
    {
        case 1:
            OpenFile();
            break;
        case 2:
            CustomConsole.Write("Введите наименование товара: ", ConsoleColor.DarkGray);
            
            var nameOfProduct = CustomConsole.ReadLine(ConsoleColor.White);

            PrintClientsInformationByProductName(nameOfProduct);
            break;
        case 3:
            break;
        case 4:
            break;
        default:
            CustomConsole.WriteLine("Введена неверная команда. Повторите попытку!\n", ConsoleColor.Red);
            continue;
    }
}

void ShowMenu()
{
    CustomConsole.WriteLine("1. Выбрать файл с данными\n" +
                      "2. Вывести информацию о клиентах, заказавших товар\n" +
                      "3. Изменить контактное лицо\n" +
                      "4. Золотой клиент", 
                        ConsoleColor.White);
}

void OpenFile()
{
    var isFinished = false;

    while (!isFinished)
    {
        CustomConsole.Write("Введите путь до файла с данными: ", ConsoleColor.DarkGray);
        
        var path = CustomConsole.ReadLine(ConsoleColor.White);
        FileInfo fileInfo = new FileInfo(path);
        //fileInfo.Extension;
        if (!IsRightPath(path)) continue;
        
        isFinished = true;
    }
    
    CustomConsole.WriteLine("Файл успешно загружен!\n", ConsoleColor.Green);
}

bool IsRightPath(string? path)
{
    if (path == null || !File.Exists(path))
    {
        CustomConsole.WriteLine("Неверный путь!", ConsoleColor.Red);
        return false;
    }

    if (!path.ToLower().EndsWith(".xlsx"))
    {
        CustomConsole.WriteLine("Неверный тип файла!", ConsoleColor.Red);
        return false;
    }

    return true;
}

void PrintClientsInformationByProductName(string productName)
{

}