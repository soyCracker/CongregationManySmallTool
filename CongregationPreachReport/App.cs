using Microsoft.Extensions.Configuration;
using SmallTool.Lib.Models.CongregationPreachReport;
using SmallTool.Lib.Services;

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
            Console.WriteLine($"CPR 會眾傳道報告程式 V{config.GetValue<string>("Version")}");
            Console.WriteLine($"ReleaseMode 正式模式: {config.GetValue<bool>("ReleaseMode")}");
            Console.WriteLine($"Directory 檔案目錄: {Directory.GetCurrentDirectory()}");

            InputModel inputModel = new InputModel();
            Console.Write("輸入目標西元年份: ");
            while (string.IsNullOrEmpty(inputModel.Year))
            {
                inputModel.Year = Console.ReadLine();
                int temp = 0;
                if (inputModel.Year.Length!=4 || !int.TryParse(inputModel.Year, out temp))
                {
                    Console.Write("格式錯誤，輸入目標西元年份: ");
                    inputModel.Year="";
                }
            }

            Console.Write("輸入目標月份(數字): ");
            while (string.IsNullOrEmpty(inputModel.Month))
            {
                inputModel.Month = Console.ReadLine();
                int temp = 0;
                if (!int.TryParse(inputModel.Month, out temp))
                {
                    Console.Write("格式錯誤，輸入目標月份(數字): ");
                    inputModel.Month="";
                }
            }

            //Console.WriteLine(inputModel.Month.GetMonthEngName());
            cprService.Start(inputModel);

            Console.WriteLine("程式結束");
            Console.ReadLine();
        }
    }
}
