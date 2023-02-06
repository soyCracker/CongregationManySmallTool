using Microsoft.Extensions.Configuration;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using SmallTool.Lib.Models.CongregationPreachReport;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SmallTool.Lib.Services
{
    public class CongregationPreachReportService
    {
        private readonly IConfiguration config;

        public CongregationPreachReportService(IConfiguration config)
        {
            this.config = config;
        }

        public void Start(InputModel input)
        {
            Console.WriteLine("CongregationPreachReportService Start!");
            string congregationXls = GetCongregationXls(input.Year);
            List<PreachReportModel> preachReportModels = ReadAllPreachReport(congregationXls, input.Month);
        }

        private string GetCongregationXls(string year)
        {
            string[] targetXlsFolderParam = { };
            //string targetXlsFolder = Path.Combine(targetXlsFolderParam);
            var rootFolder = config.GetValue<string>("TargetRootFolder");
            var firstLayerFolder = Directory.GetDirectories(rootFolder)[0];
            var wholeCongregationFolder = config.GetValue<string>("WholeCongregationFolder");
            var yearFolder = Directory.GetDirectories(Path.Combine(new string[] { firstLayerFolder, wholeCongregationFolder })).FirstOrDefault(x => x.Contains(year));
            var xls = Directory.GetFiles(yearFolder).Where(f => f.EndsWith(".xls") || f.EndsWith(".xlsx")).FirstOrDefault();
            Console.WriteLine($"全會眾傳道報告檔案: {xls}");
            return xls;
        }

        private List<PreachReportModel> ReadAllPreachReport(string xls, string month)
        {
            FileStream stream = new FileStream(xls, FileMode.Open);
            using (stream)
            {
                IWorkbook workbook;
                // 避免報異常EOF in header
                stream.Position = 0;
                if (stream.Name.ToLower().EndsWith(".xlsx"))
                {
                    Console.WriteLine("XSSFWorkbook");
                    workbook = new XSSFWorkbook(stream);
                }
                else
                {
                    Console.WriteLine("HSSFWorkbook");
                    workbook = new HSSFWorkbook(stream);
                }
                ISheet sheet = workbook.GetSheet(month + "月");

                int startRow = config.GetValue<int>("CongregationXlsStartRow");
                int startCol = config.GetValue<int>("CongregationXlsStartCol");
                int endCol = config.GetValue<int>("CongregationXlsEndCol");
                for (int r = startRow; r<sheet.LastRowNum; r++)
                {
                    for (int c = startCol; c<=endCol; c++) //8
                    {
                        Console.WriteLine(sheet.GetRow(r).GetCell(c).ToString().Trim() + "\n");
                    }
                }



                workbook.Close();
                return null;
            }
        }
    }
}