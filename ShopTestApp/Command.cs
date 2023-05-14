namespace ShopTestApp;

public class Command
{
    private readonly Action _method;

    public int Id { get; set; }
    public string Name { get; set; }

    public Command(int id, string name, Action method)
    {
        Id = id;
        Name = name;
        _method = method;
    }

    public void Run()
    {
        _method.Invoke();
    }
}