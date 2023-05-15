namespace ShopTestApp;

public static class CustomConsole
{
    public static void WriteLine(string? message, ConsoleColor color = ConsoleColor.White)
    {
        Console.ForegroundColor = color;
        Console.WriteLine(message);
    }

    public static void Write(string? message, ConsoleColor color = ConsoleColor.White)
    {
        Console.ForegroundColor = color;
        Console.Write(message);
    }

    public static void Write(int message, ConsoleColor color = ConsoleColor.White)
    {
        Console.ForegroundColor = color;
        Console.Write(message);
    }

    public static string? ReadLine(ConsoleColor color = ConsoleColor.White)
    {
        Console.ForegroundColor = color;
        return Console.ReadLine();
    }
}