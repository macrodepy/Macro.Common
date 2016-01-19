using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Macro.Common.API.Excel;

namespace Macro.Common.Demo
{
    class Program
    {
        static void Main(string[] args)
        {
            TestListToExcel();
        }

        //Require Administrator rights
        public static void TestListToExcel()
        {
            List<Name> names = new List<Name>();

            for (int i = 0; i < 5; i++)
            {
                Name name = new Name();
                name.FirstName = "Mustafa" + i;

                names.Add(name);
            }

            names.ToExcel("c:\\musta.xls");
        }
    }

    class Name
    {
        public string FirstName { get; set; }
    }
}
