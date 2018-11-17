using System;
using System.Collections.Generic;

namespace Tests
{
    class Program
    {
        static void Main(string[] args)
        {
            ExcelTest.GenerateExcel();

            Console.WriteLine("生成完毕");
            Console.ReadKey();
        }
    }
}
