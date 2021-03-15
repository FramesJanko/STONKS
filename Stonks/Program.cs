using System;
using System.Collections.Generic;
using System.IO;
using System.Net;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using OfficeOpenXml;
using Excel = Microsoft.Office.Interop.Excel;

namespace Stonks
{
    partial class Program
    {
        
        static async Task Main(string[] args)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            string workingDirectory = Environment.CurrentDirectory;
            string InputPath = Directory.GetParent(workingDirectory).Parent.Parent.Parent.FullName + "\\StonkInput.xlsx";
            string OutputPath = Directory.GetParent(workingDirectory).Parent.Parent.Parent.FullName + "\\StonkData.xlsx";

            Console.WriteLine("Reading from " + InputPath);
            Console.WriteLine("Writing to " + OutputPath);

            
            var InputFile = new FileInfo(InputPath);
            var OutputFile = new FileInfo(OutputPath);

            List<StockModel> stocks = await ReadStockDataFromExcel(InputFile);
            
            await WriteStockDataToExcel(stocks, OutputFile);
            foreach(StockModel s in stocks)
            {
                Console.WriteLine(s.Name);
            }

        }

        private static async Task WriteStockDataToExcel(List<StockModel> stocks, FileInfo file)
        {
            //DeleteIfExists(file);
            using var package = new ExcelPackage(file);

            string time = DateTime.Now.ToString();
            var xlWorksheet = package.Workbook.Worksheets.Add("Prices from: " + time);
            var range = xlWorksheet.Cells["A1"].LoadFromCollection(stocks, true);
            range.AutoFitColumns();
            await package.SaveAsync();
        }

        //Don't need this delete method, because we're just adding a new sheet with the current time in the name of the sheet.
        //private static void DeleteIfExists(FileInfo file)
        //{
        //    if (file.Exists)
        //    {
        //        file.Delete();
        //    }
        //}

        private static List<StockModel> DummyMethod()
        {
            List<StockModel> output = new ()
            {
                new () { Id = 1, Name = "Gamestop", Ticker = "GME", Open = 162.5f, PreviousClose = 101.4f, High = 176.98f, Low = 84.54f },
                new () { Id = 2, Name = "Wingstop", Ticker = "WING", Open = 130.5f, PreviousClose = 135.8f, High = 136.75f, Low = 129.87f},
                new () { Id = 3,  Name = "Aphria", Ticker = "APHA", Open = 19.50f, PreviousClose = 20.74f, High = 20.74f, Low = 19.50f }
            };

            return output;
        }

        private static List<RealStockModel> GetRealStockSetupData()
        {
            List<RealStockModel> output = new()
            {
                new() { Id = 1, Name = "Gamestop", Ticker = "GME", Open = "=B2.Open", Close = "=B2.Close", High = "=B2.High", Low = "=B2.Low" },
                new() { Id = 2, Name = "Wingstop", Ticker = "WING", Open = "=B3.Open", Close = "=B3.Close", High = "=B3.High", Low = "=B3.Low" },
                new() { Id = 3, Name = "Aphria", Ticker = "APHA", Open = "=B4.Open", Close = "=B4.Close", High = "=B4.High", Low = "=B4.Low" }
            };
            

            return output;
        }

        private static async Task<List<StockModel>> ReadStockDataFromExcel(FileInfo file)
        {
            List<StockModel> output = new();
            using var package = new ExcelPackage(file);
            await package.LoadAsync(file);
            var worksheet = package.Workbook.Worksheets[0];

            int row = 2;
            int col = 1;
            
            while (string.IsNullOrWhiteSpace(worksheet.Cells[row, col].Value?.ToString()) == false)
            {
                StockModel stock = new();
                stock.Id = int.Parse(worksheet.Cells[row, col].Value.ToString());
                stock.Name = worksheet.Cells[row, col + 1].Value.ToString();
                stock.Ticker = worksheet.Cells[row, col + 2].Value.ToString();
                stock.Open = float.Parse(worksheet.Cells[row, col + 3].Value.ToString());
                stock.PreviousClose = float.Parse(worksheet.Cells[row, col + 4].Value.ToString());
                stock.High = float.Parse(worksheet.Cells[row, col + 5].Value.ToString());
                stock.Low = float.Parse(worksheet.Cells[row, col + 6].Value.ToString());
                output.Add(stock);
                row++;
            }
            return output;
        }
    }
    
}
