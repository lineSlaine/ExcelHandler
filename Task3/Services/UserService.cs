using ClosedXML.Excel;
using DocumentFormat.OpenXml.Spreadsheet;
using Task3.Models;

namespace Task3.Services;

public class UserService
{
    private string _filePath;
    public List<User> Users { get; private set; }

    public UserService(string file)
    {
        Users = Get(file);
        _filePath = file;
    }

    private List<User> Get(string file)
    {
        var users = new List<User>();

        using var workbook = new XLWorkbook(file);
        var worksheet = workbook.Worksheet("Клиенты");
        int i = 2;
        while (!string.IsNullOrWhiteSpace(worksheet.Cell(i, 1).Value.ToString()))
        {

            users.Add(new User(Convert.ToInt16(worksheet.Cell(i, 1).Value.ToString()), worksheet.Cell(i, 2).Value.ToString(),
                worksheet.Cell(i, 3).Value.ToString(), worksheet.Cell(i, 4).Value.ToString()));
            i++;
        }
        return users;
    }
    public string GetContactPersonNameById(int id)
    {
        string name = string.Empty;
        foreach (var user in Users)
        {
            if (user.Id == id)
            {
                name = user.ContactPerson;
                break;
            }
        }
        return name;
    }
    public bool Update(int idInList, string contactName, string organizationName)
    {
        using var workbook = new XLWorkbook(_filePath);
        var worksheet = workbook.Worksheet("Клиенты");
        int i = 2;
        
        while (!string.IsNullOrWhiteSpace(worksheet.Cell(i, 1).Value.ToString()))
        {
            if (Users[idInList-1].Id == Convert.ToInt16(worksheet.Cell(i,1).Value.ToString()))
            {
                try
                {
                    worksheet.Cell(i, 2).Value = organizationName;
                    worksheet.Cell(i, 4).Value = contactName;
                    Users[idInList - 1].ContactPerson = contactName;
                    Users[idInList - 1].OrganizationName = organizationName;
                    workbook.Save();

                    return true;
                }
                catch { return false; }
            }
            i++;
        }

        return false;
    }
}
