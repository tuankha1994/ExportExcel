using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExportExcelDemo
{
    class Program
    {
        static void Main(string[] args)
        {
            ExcelHelper.ExportExcelFile(@"D:\users.xlsx", Data.ListUser);
            ExcelHelper.ExportExcelFile(@"D:\roles.xlsx", Data.ListRole);          
        }
    }
}
