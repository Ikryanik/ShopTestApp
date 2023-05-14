namespace ShopTestApp;

public class Order
{
    public int Code { get; set; }
    public int ProductCode { get; set; }
    public int ClientCode { get; set; }
    public int OrderNumber { get; set; }
    public int Count { get; set; }
    public DateTime PlacementDate { get; set; }
}