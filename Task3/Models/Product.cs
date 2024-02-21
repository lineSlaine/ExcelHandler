namespace Task3.Models;

public class Product
{
    public Product(int id, string productName, string measureUnit, decimal productPrice)
    {
        Id = id;
        ProductName = productName;
        MeasureUnit = measureUnit;
        ProductPrice = productPrice;
    }

    public int Id { get; set; }
    public string ProductName { get; set; }
    public string MeasureUnit { get; set; }
    public decimal ProductPrice { get; set; }
}
