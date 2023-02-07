using Microsoft.Extensions.Configuration;
using SmallTool.Lib.Models.CongregationPreachReport;
using SmallTool.Lib.Services;
using SmallTool.Lib.Utils;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CongregationPreachReport
{
    public class App
    {
        private readonly IConfiguration config;
        private readonly CongregationPreachReportService cprService;

        public App(IConfiguration config, CongregationPreachReportService cprService)
        {
            this.config = config;
            this.cprService=cprService;
        }

        public void Run()
        {
            Console.WriteLine("CongregationPreachReport");
            Console.WriteLine("ReleaseMode: " + config.GetValue<bool>("ReleaseMode"));
            Console.WriteLine("Directory: " + Directory.GetCurrentDirectory());

            InputModel inputModel = new InputModel();
            Console.Write("輸入目標年份: ");
            inputModel.Year = Console.ReadLine();
            Console.Write("輸入目標月份: ");
            inputModel.Month = Console.ReadLine();

            //Console.WriteLine(inputModel.Month.GetMonthEngName());
            cprService.Start(inputModel);

            Console.WriteLine("程式結束");
            Console.ReadLine();
        }
    }
}
