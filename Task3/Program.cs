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
        if (product != null)
        {
            var requests = requestService.GetByProductId(product.Id);
            if (requests.Any())
            {
                var stringBuilder = new StringBuilder("Контакное лицо\tКол-во\tСумма\tДата\n");
                foreach (var request in requests)
                {
                    stringBuilder.Append($"{userService.GetContactPersonNameById(request.UserId)} \t {request.ProductCount} \t " +
                        $"{request.ProductCount*product.ProductPrice} \t {request.PlacementDate} \n");
                }
                return stringBuilder.ToString();
            }
            else return "Такой продукт еще никто не заказывал!";
        }
        else return "Такого продукта нет в списке!";
    }

    public static string? TextConstructor(int year, int month, RequestService requestService, UserService userService)
    {
        var requests = requestService.GetByData(year, month);
        int userId = -1, count = 0;
        if (requests.Any())
        {
            foreach (var request in requests)
            {
                var countRequests = requestService.RequestsCountByDate(request.UserId, requests);
                if (countRequests > count)
                {
                    userId = request.UserId;
                    count = countRequests;
                }
            }
            if (count == 0) 
            {
                return "Такого клиента нет!";
            }
            else return $"Клиент {userService.GetContactPersonNameById(userId)} является золотым клиентом совершив {count} заказ(ов)";
        }
        else
        {  
            return "Такого клиента нет!";
        }
    }

    static void Main(string[] args)
    {
        UserService userService = null;
        ProductService productService = null;
        RequestService requestService = null;

        bool check = false;
        while (!check)
        {
            Console.Write("Введите путь к файлу: ");
            //string? file = @"C:\Users\musfo\Desktop\тест\Практическое задание для кандидата.xlsx";
            string? file = Console.ReadLine();
            if (!string.IsNullOrWhiteSpace(file))
            {
                if (File.Exists(file))
                {
                    try
                    {
                        //инициализация таблиц
                        userService = new UserService(file);
                        productService = new ProductService(file);
                        requestService = new RequestService(file);
                        check = true;
                    }
                    catch (Exception ex)
                    {
                        Console.Clear();
                        Console.WriteLine($"Что-то пошло не так! Попробуйте еще раз. Ошибка: {ex}");
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
                    Console.WriteLine(TextConstructor(name, userService, productService,requestService));
                    break;
                case 2:
                    Console.Clear();
                    int i = 1;
                    var stringBuilder = new StringBuilder("Контакное лицо\tНаименование организации\n");
                    foreach (var user in userService.Users)
                    {
                        stringBuilder.Append($"{i}. {user.ContactPerson}\t{user.OrganizationName} \n");
                        i++;
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
                        if(!string.IsNullOrWhiteSpace(newContactName) && !string.IsNullOrWhiteSpace(newOrganizationName))
                        {
                            if (userService.Update(clientNumber, newContactName, newOrganizationName))
                            {
                                Console.Clear();
                                Console.WriteLine("Успешно!");
                            }
                            else
                            {
                                Console.Clear();
                                Console.WriteLine("Введенные вами данные некорректны!");
                            }
                        }
                        else
                        {
                            Console.Clear();
                            Console.WriteLine("Введенные вами данные некорректны!");
                        }
                    }
                    catch
                    {
                        Console.Clear();
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
