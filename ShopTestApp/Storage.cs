namespace ShopTestApp;

public static class Storage
{
    public static string? Path { get; set; } = string.Empty;

    public static List<Product> Products = new();
    public static List<Client> Clients = new();
    public static List<Order>  Orders = new();
}