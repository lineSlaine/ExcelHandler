namespace Task3.Models;

public class Request
{
    public Request(int id, int productId, int userId, int number, int productCount, string placementDate)
    {
        Id = id;
        ProductId = productId;
        UserId = userId;
        Number = number;
        ProductCount = productCount;
        var date = placementDate.Split('.', ' ');
        PlacementDate = new DateOnly(Convert.ToInt16(date[2]), Convert.ToInt16(date[1]), Convert.ToInt16(date[0]));
    }

    public int Id { get; set; }
    public int ProductId { get; set; }
    public int UserId { get; set; }
    public int Number { get; set; }
    public int ProductCount { get; set; }
    public DateOnly PlacementDate { get; set; }
}
