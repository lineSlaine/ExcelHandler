using DocumentFormat.OpenXml.Office2013.Excel;
using System.ComponentModel;
using System.Text;
using Task3.Services;

namespace Task3;

internal class Program
{
    public static string? TextConstructor(string productName, UserService userService, 
        ProductService productService, RequestService requestService)
    {
        var product = productService.GetProductIdByName(productName);
        if (product == null)
        {
            return "Такого продукта нет в списке!";
        }
        var requests = requestService.GetByProductId(product.Id);
        if (!requests.Any())
        {
            
            return "Такой продукт еще никто не заказывал!";
        }
        var stringBuilder = new StringBuilder("Контакное лицо\tКол-во\tСумма\tДата\n");
        foreach (var request in requests)
        {
            stringBuilder.Append($"{userService.GetContactPersonNameById(request.UserId)} \t {request.ProductCount} \t " +
                $"{request.ProductCount * product.ProductPrice} \t {request.PlacementDate} \n");
        }
        return stringBuilder.ToString();

    }

    public static string? TextConstructor(int year, int month, RequestService requestService, UserService userService)
    {
        var requests = requestService.GetByData(year, month);
        if (!requests.Any())
        {
            return "Такого клиента нет!";
        }
        int userId = -1, count = 0;
        foreach (var request in requests)
        {
            var countRequests = requestService.RequestsCountByDate(request.UserId, requests);
            if (countRequests > count)
            {
                userId = request.UserId;
                count = countRequests;
            }
        }
        return count == 0 ? "Такого клиента нет!" : $"Клиент {userService.GetContactPersonNameById(userId)} является золотым клиентом совершив {count} заказ(ов)";
    }

    static void Main(string[] args)
    {
        UserService userService = null;
        ProductService productService = null;
        RequestService requestService = null;

        bool isInitializationNotCompleted = true;
        while (isInitializationNotCompleted)
        {
            Console.Write("Введите путь к файлу: ");
            string? file = Console.ReadLine();
            Console.Clear();
            if (string.IsNullOrWhiteSpace(file))
            {
                Console.WriteLine("Что-то пошло не так! Попробуйте еще раз.");
                continue;
            }
            if (!File.Exists(file))
            {
                Console.WriteLine("Путь указан неверно! Попробуйте еще раз.");
                continue;
            }
            try
            {
                //инициализация таблиц
                userService = new UserService(file);
                productService = new ProductService(file);
                requestService = new RequestService(file);
                isInitializationNotCompleted = false;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Что-то пошло не так! Попробуйте еще раз. Ошибка: {ex.Message}");
                break;
            }

        }

        while (true)
        {
            Console.WriteLine("Доступные действия: \n " +
                "1. По названию товара вывести список клиентов, заказавших этот товар \n " +
                "2. Запрос на изменение контактного лица с указанием параметров(ФИО, название организации) \n " +
                "3. Определение золотого клиента, с наибольшим кол-вом заказов за указанный год или месяц \n" +
                "Введите цифру действия, которое хотите совершить: ");
            int choice;
            try
            {
                choice = Convert.ToInt16(Console.ReadLine());
            }
            catch
            {
                Console.Clear();
                Console.WriteLine("Что-то пошло не так! Попробуйте еще раз!");
                continue;
            }
            switch(choice)
            {
                case 1:
                    Console.Clear();
                    Console.WriteLine("Введите название товара: ");
                    string? name = Console.ReadLine();
                    Console.WriteLine(TextConstructor(name, userService, productService, requestService));
                    break;
                case 2:
                    Console.Clear();
                    var stringBuilder = new StringBuilder("Контакное лицо\tНаименование организации\n");
                    for(var i = 0; i < userService.Users.Count; i++)
                    {
                        stringBuilder.Append($"{i}. {userService.Users[i].ContactPerson}\t{userService.Users[i].OrganizationName} \n");
                    }
                    Console.WriteLine(stringBuilder.ToString());
                    Console.Write("Введите номер клиента, данные которые хотите изменить: ");
                    try
                    {
                        int clientNumber = Convert.ToInt16(Console.ReadLine());
                        Console.Clear();
                        Console.Write("Введите новое ФИО: ");
                        string newContactName = Console.ReadLine();
                        Console.Write("\nВведите новое название организации: ");
                        string newOrganizationName = Console.ReadLine();
                        Console.Clear();
                        if (string.IsNullOrWhiteSpace(newContactName) || string.IsNullOrWhiteSpace(newOrganizationName))
                        {
                            Console.WriteLine("Введенные вами данные некорректны!");
                            continue;
                        }
                        if (!userService.Update(clientNumber, newContactName, newOrganizationName))
                        {
                            Console.WriteLine("Введенные вами данные некорректны!");
                            continue;
                        }
                        Console.WriteLine("Успешно!");
                    }
                    catch
                    {
                        Console.WriteLine("Введенные вами данные некорректны!");
                    }
                    break;
                case 3:
                    Console.Clear();
                    Console.Write("Введите год: ");
                    try
                    {
                        int year = Convert.ToInt16(Console.ReadLine());
                        Console.Write("\n Введите месяц: ");
                        int month = Convert.ToInt16(Console.ReadLine());
                        Console.Clear();
                        Console.WriteLine(TextConstructor(year, month, requestService, userService));
                    }
                    catch 
                    { 
                        Console.Clear();
                        Console.WriteLine("Введенные вами данные некорректны!");
                    }

                    break;
                default:
                    Console.Clear();
                    Console.WriteLine("Попробуйте написать число от 1 до 3.");
                    break;
            }
        }
    }
}
