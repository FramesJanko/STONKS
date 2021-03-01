using System;
using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using OfficeOpenXml;
using Excel = Microsoft.Office.Interop.Excel;

namespace Stonks
{
    class Program
    {
        
        
        
        static async Task Main(string[] args)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            string workingDirectory = Environment.CurrentDirectory;
            string path = Directory.GetParent(workingDirectory).Parent.Parent.FullName + "\\DemoStonkData.xlsx";
            Console.WriteLine("File located at " + path);
            //C:\Users\Robby\Documents\GitHub\STONKS\
            var file = new FileInfo(path);

            var stocks = GetSetupData();

            await SaveExcelFile(stocks, file);
        }

        private static async Task SaveExcelFile(List<StockModel> stocks, FileInfo file)
        {
            DeleteIfExists(file);

            using var package = new ExcelPackage(file);

            string time = DateTime.Today.ToString();
            var xlWorksheet = package.Workbook.Worksheets.Add("Charts: " + time);
            var range = xlWorksheet.Cells["A1"].LoadFromCollection(stocks, true);
            range.AutoFitColumns();
            await package.SaveAsync();
        }

        private static void DeleteIfExists(FileInfo file)
        {
            if (file.Exists)
            {
                file.Delete();
            }
        }

        private static List<StockModel> GetSetupData()
        {
            List<StockModel> output = new()
            {
                new() { Id = 1, Name = "Gamestop", Ticker = "GME", Open = 161.5f, Close = 101.4f, High = 176.98f, Low = 84.54f },
                new() { Id = 2, Name = "Wingstop", Ticker = "WING", Open = 130.5f, Close = 135.8f, High = 136.75f, Low = 129.87f},
                new() { Id = 3,  Name = "Aphria", Ticker = "APHA", Open = 19.88f, Close = 20.74f, High = 20.74f, Low = 18.99f },
            };

            return output;
        }

    }
    
}
