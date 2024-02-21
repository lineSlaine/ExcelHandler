namespace Task3.Models;

public class User
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
