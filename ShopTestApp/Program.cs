using ShopTestApp;
using ShopTestApp.Services;

var fileService = new FileService();
var clientService = new ClientService();

var commands = new List<Command>
{
    new(1, "Выбрать файл с данными", fileService.OpenFile),
    new(2, "Вывести информацию о клиентах, заказавших товар", clientService.PrintClientsInformationByProductName),
    new(3, "Изменить контактное лицо", fileService.ChangeContactPerson),
    new(4, "Золотой клиент", clientService.FindAGoldClient),
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