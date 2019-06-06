using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExportExcelDemo
{
    class Data
    {
        public static List<User> ListUser { get; set; } = new List<User>
        {
            new User{Name="Kha", Gender=true, BirthDate = DateTime.Now, Address="GL", PhoneNumber="123456"},
            new User{Name="Định", Gender=true,BirthDate = DateTime.Now, Address="ĐL", PhoneNumber="4321"},
            new User{Name="Cường", Gender=true,BirthDate = DateTime.Now, Address="HCM", PhoneNumber="abcd"},
            new User{Name="Trang", Gender=false,BirthDate = DateTime.Now, Address="ĐN", PhoneNumber="98765"}
        };

        public static List<Role> ListRole { get; set; } = new List<Role>
        {
            new Role{Name="Administrator", Description="Admin is a Manager"},
            new Role{Name="Guest", Description="Guest is a normal user"}
        };

    }

    public class User
    {
        public string Name { get; set; }
        public bool Gender { get; set; }
        public DateTime BirthDate { get; set; }
        public string Address { get; set; }
        public string PhoneNumber { get; set; }
    }

    public class Role
    {
        public string Name { get; set; }
        public string Description { get; set; }
    }
}
