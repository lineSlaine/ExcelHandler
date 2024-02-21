using ClosedXML.Excel;
using Task3.Models;

namespace Task3.Services;

public class ProductService
{
    public IEnumerable<Product> Products { get; private set; }
    public ProductService(string file)
    {
        Products = Get(file);
    }

    private IEnumerable<Product> Get(string file)
    {
        var products = new List<Product>();

        using var workbook = new XLWorkbook(file);
        var worksheet = workbook.Worksheet("Товары");
        int i = 2;
        while (!string.IsNullOrWhiteSpace(worksheet.Cell(i, 1).Value.ToString()))
        {
            products.Add(new Product(Convert.ToInt16(worksheet.Cell(i, 1).Value.ToString()), worksheet.Cell(i, 2).Value.ToString(),
                worksheet.Cell(i, 3).Value.ToString(), Convert.ToDecimal(worksheet.Cell(i, 4).Value.ToString())));
            i++;
        }
        return products;
    }
    public Product? GetProductIdByName(string name)
    {
        var lowerName = name.ToLower();
        foreach (var product in Products)
        {
            if (product.ProductName.ToLower().Equals(lowerName))
            {
                return product;
            }
        }
        return null;
    }
}
