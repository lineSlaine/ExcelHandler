using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Text.RegularExpressions;

namespace Task3;

internal class Program
{
    static string GetColumnName(string? cellName)
    {
        if (cellName is null)
        {
            return string.Empty;
        }

        // Create a regular expression to match the column name portion of the cell name. [A-Za-z]
        Regex regex = new Regex(@"[A-Za-z]|\d+");
        Match match = regex.Match(cellName);

        return match.Value;
    }

    // Given a cell name, parses the specified cell to get the row index.
    static uint? GetRowIndex(string? cellName)
    {
        if (cellName is null)
        {
            return null;
        }

        // Create a regular expression to match the row index portion the cell name.
        Regex regex = new Regex(@"\d+");
        Match match = regex.Match(cellName);

        return uint.Parse(match.Value);
    }
    static void Main(string[] args)
    {
        SpreadsheetDocument spreadSheet = null;
        bool check = true;
        while (check)
        {
            //Console.Write("Введите путь к файлу: ");
            string? path = @"C:\Users\musfo\Desktop\тест\Практическое задание для кандидата.xlsx"; //Console.ReadLine();
            if (!string.IsNullOrWhiteSpace(path))
            {
                if (File.Exists(path))
                {
                    try
                    {
                        using (SpreadsheetDocument document = SpreadsheetDocument.Open(path, false))
                        {
                            

                            //orksheetPart worksheetPart = (WorksheetPart)document.WorkbookPart!.GetPartById(id);


                            var shareStringPart = document.WorkbookPart.GetPartsOfType<SharedStringTablePart>();
                            var items = shareStringPart.First().SharedStringTable.Elements<SharedStringItem>();
                            foreach (var item in items)
                            {
                                Console.WriteLine(item.InnerText);
                            }
                        }




                        check = false;
                    }
                    catch(Exception ex)
                    {
                        Console.Clear();
                        Console.WriteLine("Что-то пошло не так! Попробуйте еще раз." + ex.ToString());
                        break;
                    }
                }
                else
                {
                    Console.Clear();
                    Console.WriteLine("Путь указан неверно! Попробуйте еще раз.");
                }


            }
            else
            {
                Console.Clear();
                Console.WriteLine("Что-то пошло не так! Попробуйте еще раз.");
            }
        }
        //spreadSheet.Dispose();
    }
}

class User
{
    public User(int id, string organizationName, string address, string contactPerson)
    {
        Id = id;
        OrganizationName = organizationName;
        Address = address;
        ContactPerson = contactPerson;
    }

    public int Id { get; set; }
    public string OrganizationName { get; set; }
    public string Address { get; set; }
    public string ContactPerson { get; set; }
}
class UserService
{

}


class Product
{
    public Product(int id, string productName, string measureUnit, int productPrice)
    {
        Id = id;
        ProductName = productName;
        MeasureUnit = measureUnit;
        ProductPrice = productPrice;
    }

    public int Id { get; set; }
    public string ProductName { get; set; }
    public string MeasureUnit { get; set; }
    public int ProductPrice { get; set; }
}
class ProductService
{

}



class Request
{
    public Request(int id, int productId, int userId, int number, int productCount, DateOnly placementDate)
    {
        Id = id;
        ProductId = productId;
        UserId = userId;
        Number = number;
        ProductCount = productCount;
        PlacementDate = placementDate;
    }

    public int Id { get; set; }
    public int ProductId { get; set; }
    public int UserId { get; set; }
    public int Number { get; set; }
    public int ProductCount { get; set; }
    public DateOnly PlacementDate {  get; set; }
}
class RequestService
{

}
