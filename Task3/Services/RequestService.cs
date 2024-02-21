using ClosedXML.Excel;
using Task3.Models;

namespace Task3.Services;

public class RequestService
{
    public IEnumerable<Request> Requests { get; private set; }
    public RequestService(string file)
    {
        Requests = Get(file);
    }
    private IEnumerable<Request> Get(string file)
    {
        var request = new List<Request>();

        using var workbook = new XLWorkbook(file);
        var worksheet = workbook.Worksheet("Заявки");
        int i = 2;
        while (!string.IsNullOrWhiteSpace(worksheet.Cell(i, 1).Value.ToString()))
        {
            request.Add(new Request(
                Convert.ToInt16(worksheet.Cell(i, 1).Value.ToString()), Convert.ToInt16(worksheet.Cell(i, 2).Value.ToString()),
                Convert.ToInt16(worksheet.Cell(i, 3).Value.ToString()), Convert.ToInt16(worksheet.Cell(i, 4).Value.ToString()),
                Convert.ToInt16(worksheet.Cell(i, 5).Value.ToString()), worksheet.Cell(i, 6).Value.ToString()));
            i++;
        }
        return request;
    }
    public IEnumerable<Request> GetByProductId(int id)
    {
        var requests = new List<Request>();
        foreach (var request in Requests) 
        {
            if (request.ProductId == id) requests.Add(request);
        }
        return requests;
    }

    public IEnumerable<Request> GetByData(int year, int month)
    {
        var requests = new List<Request>();
        foreach (var request in Requests)
        {
            if (request.PlacementDate.Month == month && request.PlacementDate.Year == year) requests.Add(request);
        }
        return requests;
    }
    public int RequestsCountByDate(int id, IEnumerable<Request> requests)
    {
        int count = 0;
        if (requests.Any())
        {
            foreach (var request in requests)
            {
                if(request.UserId == id) count++;
            }
        }
        else return 0;
        return count;
    }
}
