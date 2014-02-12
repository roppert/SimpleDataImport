using System;
using System.Collections.Generic;
using SimpleDataImport;

namespace Examples
{
    public class NamedEntry
    {
        public string Description { get; set; }
        public string Name { get; set; }
        public float Price { get; set; }
        public bool Active { get; set; }
        public int Number { get; set; }
        public DateTime Created { get; set; }

        public NamedEntry()
        {
            // Default for string is null but we need it to be empty or
            // else ToString() returns nothing due to string.Join and null.
            this.Description = string.Empty;
            this.Name = string.Empty;
        }

        /// <summary>
        /// This override is just to print example result nicely
        /// </summary>
        /// <returns></returns>
        override public string ToString()
        {
            return string.Join("|", Description, Name, Price, Active, Number, Created);
        }
    }

    public class NamelessEntry
    {
        public string A { get; set; }
        public string B { get; set; }
        public double C { get; set; }
        public int D { get; set; }

        public NamelessEntry()
        {
            this.A = string.Empty;
            this.B = string.Empty;
        }

        override public string ToString()
        {
            return string.Join("|", A, B, C, D);
        }
    }

    // NOTE: When all strings the format for numbers etc will be en-US
    //       so you need to convert this to CurrentUICulture before
    //       presenting data to the user.
    public class AllStringsEntry
    {
        public string A { get; set; }
        public string B { get; set; }
        public string C { get; set; }
        public string D { get; set; }

        public AllStringsEntry()
        {
            A = string.Empty;
            B = string.Empty;
            C = string.Empty;
            D = string.Empty;
        }

        override public string ToString()
        {
            return string.Join("|", A, B, C, D);
        }
    }

    class Program
    {
        static void Main(string[] args)
        {
            string path = "../../test.xlsx";
            System.Console.WriteLine(args.Length);
            if (args.Length == 1)
                path = args[0];
            System.Console.WriteLine(path);

            var edi = new OpenXmlImport<NamedEntry>(path);
            List<NamedEntry> data = edi.Import();
            foreach (NamedEntry row in data)
                System.Console.WriteLine(row);

            System.Console.WriteLine("---");

            var edi2 = new OpenXmlImport<NamelessEntry>(path, hasColumnNames: false);
            List<NamelessEntry> data2 = edi2.Import();
            foreach (NamelessEntry row in data2)
                System.Console.WriteLine(row);

            System.Console.WriteLine("---");

            var edi3 = new OpenXmlImport<AllStringsEntry>(path, hasColumnNames: false);
            List<AllStringsEntry> data3 = edi3.Import();
            foreach (AllStringsEntry row in data3)
                System.Console.WriteLine(row);

            System.Console.Read();
        }
    }
}
