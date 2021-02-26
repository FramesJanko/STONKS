using System;
using System.IO;
using Microsoft.Office.Interop.Excel;

namespace Stonks
{
    class Program
    {
        static void Main(string[] args)
        {
            
            string stockName = "GME";
            float stockPrice = 44.25f;
            if (stockName == "GME" && stockPrice == 44.25) {
                Console.WriteLine("Buy now");
            }

            
            File.WriteAllText(@"C:\Users\rheck\OneDrive\Documents\GitHub\Stonks\Stonks\Data\Stonks.txt", stockName);
            
        }
    }
}
