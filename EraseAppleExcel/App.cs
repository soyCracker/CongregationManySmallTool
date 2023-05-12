using Microsoft.Extensions.Configuration;
using SmallTool.Lib.Services;
using SmallTool.Lib.Utils;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EraseAppleExcel
{
    public class App
    {
        private readonly EraseAppleExcelService eaeService;
        private readonly string version = "1.0";

        public App(EraseAppleExcelService eaeService)
        {
            this.eaeService=eaeService;
        }

        public void Run()
        {
            try
            {
                Console.WriteLine($"EraseAppleExcel V{version}");
                Console.WriteLine("Directory: " + Directory.GetCurrentDirectory());

                eaeService.Start();
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
                Console.Read();
            }
        }
    }
}
