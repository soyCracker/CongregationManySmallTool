using iText.Forms;
using iText.Forms.Fields;
using iText.IO.Font;
using iText.Kernel.Font;
using iText.Kernel.Pdf;
using Microsoft.Extensions.Configuration;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using SmallTool.Lib.Models.CongregationPreachReport;
using SmallTool.Lib.Models.CPR_Maui_App;
using SmallTool.Lib.Utils;
using System.Text.RegularExpressions;

namespace SmallTool.Lib.Services
{
    public class CPRMauiService
    {
        private readonly IConfiguration config;

        public CPRMauiService(IConfiguration config)
        {
            this.config = config;
        }

        public async Task Start(IndexModel model)
        {
            await Task.Run(() =>
            {
                Console.WriteLine("\n##### CongregationPreachReportService Start! #####\n");
                Console.WriteLine("##### 確認全會眾傳道報告檔案 #####\n");
                string congregationXls = GetCongregationXls(model.Year, model.RootFolder);
                Console.WriteLine("##### 讀取所有人的報告 #####\n");
                List<PreachReportModel> preachReportModels = ReadAllPreachReport(congregationXls, model.Month, model.RootFolder);
                Console.WriteLine("##### 開始寫入PDF #####\n");
                WriteToEveryonePdf(preachReportModels, model);
            });
        }

        //確認全會眾傳道報告檔案
        private string GetCongregationXls(string year, string rootFolder)
        {
            var wholeCongregationFolder = config.GetValue<string>("WholeCongregationFolder");
            var yearFolder = Directory.GetDirectories(System.IO.Path.Combine(new string[] { rootFolder, wholeCongregationFolder })).FirstOrDefault(x => x.Contains(year));
            //稍微避免奇怪的檔案
            var xls = Directory.GetFiles(yearFolder).Where(f => !System.IO.Path.GetFileName(f).StartsWith(".") && (f.EndsWith(".xls") || f.EndsWith(".xlsx"))).FirstOrDefault();
            Console.WriteLine($"全會眾傳道報告檔案: {xls}\n");
            return xls;
        }

        //讀取所有人的報告
        private List<PreachReportModel> ReadAllPreachReport(string xls, string month, string rootFolder)
        {
            using FileStream stream = new FileStream(xls, FileMode.Open);
            IWorkbook workbook;
            // 避免報異常EOF in header
            stream.Position = 0;
            if (stream.Name.ToLower().EndsWith(".xlsx"))
            {
                Console.WriteLine("Excel type: xlsx XSSFWorkbook\n");
                workbook = new XSSFWorkbook(stream);
            }
            else
            {
                Console.WriteLine("Excel type: xls HSSFWorkbook\n");
                workbook = new HSSFWorkbook(stream);
            }
            ISheet sheet = workbook.GetSheet(month + "月");

            List<PreachReportModel> prModels = GetEveryoneBasicData(rootFolder);
            int startRow = config.GetValue<int>("CongregationXlsStartRow");
            for (int r = startRow; r<sheet.LastRowNum; r++)
            {
                string tempName = sheet.GetRow(r).GetCell(2).SafeTrim();
                if (!string.IsNullOrEmpty(tempName))
                {
                    PreachReportModel prModel = prModels.FirstOrDefault(x => x.Name==tempName);
                    if (prModel != null)
                    {
                        prModel.Distribution = sheet.GetRow(r).GetCell(3).SafeTrim();
                        prModel.Video = sheet.GetRow(r).GetCell(4).SafeTrim();
                        prModel.Hour = sheet.GetRow(r).GetCell(5).SafeTrim();
                        prModel.Review = sheet.GetRow(r).GetCell(6).SafeTrim();
                        prModel.Study = sheet.GetRow(r).GetCell(7).SafeTrim();
                        prModel.Remark = sheet.GetRow(r).GetCell(8).SafeTrim();
                        Console.WriteLine($"{prModel.Name}, {prModel.Team}, {prModel.Type}, {prModel.Distribution}, {prModel.Video}, {prModel.Hour}, "
                            + $"{prModel.Review}, {prModel.Study}, {prModel.Remark}");
                    }
                }
            }
            workbook.Close();
            return prModels;
        }

        //先從資料夾結構讀取所有人的名字、所屬小組、傳道員類別
        private List<PreachReportModel> GetEveryoneBasicData(string rootFolder)
        {
            var teams = Directory.GetDirectories(rootFolder).Where(x => x != config.GetValue<string>("WholeCongregationFolder"));
            List<PreachReportModel> prModels = new List<PreachReportModel>();
            //小組資料夾
            foreach (var team in teams)
            {
                //先驅與傳道員資料夾
                var preacherTypes = Directory.GetDirectories(team);
                foreach (var pType in preacherTypes)
                {
                    //各傳道員資料夾
                    var persons = Directory.GetDirectories(pType);
                    foreach (var person in persons)
                    {
                        PreachReportModel model = new PreachReportModel
                        {
                            Name = Path.GetFileName(person),
                            Type = Path.GetFileName(pType),
                            Team = Path.GetFileName(team)
                        };
                        prModels.Add(model);
                    }
                }
            }
            return prModels;
        }

        //寫入每個人的PDF
        private void WriteToEveryonePdf(List<PreachReportModel> prModels, IndexModel indexModel)
        {
            foreach (var prModel in prModels)
            {
                string pdfPath = GetManPdf(prModel, indexModel.Year, indexModel.RootFolder);
                if (!string.IsNullOrEmpty(pdfPath))
                {
                    using FileStream fs = new FileStream(pdfPath, FileMode.Open);
                    using PdfDocument pdfDoc = new PdfDocument(new PdfReader(fs), new PdfWriter(pdfPath + config.GetValue<string>("TempPdfStr")));
                    try
                    {
                        PdfAcroForm form = PdfAcroForm.GetAcroForm(pdfDoc, true);
                        IDictionary<string, PdfFormField> fields = form.GetFormFields();
                        SetPdfFields(pdfDoc, fields, prModel, indexModel);
                        pdfDoc.Close();
                        fs.Close();
                        File.Delete(pdfPath);
                        File.Move(pdfPath + config.GetValue<string>("TempPdfStr"), pdfPath);
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"{prModel.Name} 的作業發生錯誤， {ex.ToString()}\n");
                        pdfDoc.Close();
                        fs.Close();
                        DelTempPdf(pdfPath);
                    }
                }
            }
        }

        private void DelTempPdf(string pdfPath)
        {
            if (pdfPath.EndsWith(config.GetValue<string>("TempPdfStr")) && File.Exists(pdfPath))
            {
                File.Delete(pdfPath);
            }
        }

        private string GetManPdf(PreachReportModel prModel, string year, string rootFolder)
        {
            var pdfFolder = System.IO.Path.Combine(rootFolder, prModel.Team, prModel.Type, prModel.Name);
            if (Directory.Exists(pdfFolder))
            {
                DelAllDelMePdf(pdfFolder);
                //盡量去除奇怪的檔案名稱
                string pattern = "[0-9]{4}-[0-9]{2}.[Pp][Dd][Ff]";
                Regex regex = new Regex(pattern);
                var pdf = Directory.GetFiles(pdfFolder).Where(x => regex.IsMatch(System.IO.Path.GetFileName(x))
                            && System.IO.Path.GetFileName(x).Length==11
                            //包含年份後2碼，比如2023=>23
                            && System.IO.Path.GetFileName(x).Substring(2).Contains(year.Substring(2, 2))).Max();
                if (!string.IsNullOrEmpty(pdf))
                {
                    Console.WriteLine($"{prModel.Name} {System.IO.Path.GetFileName(pdf)}");
                    return pdf;
                }
            }
            Console.WriteLine($"##### {prModel.Team} {prModel.Type} {prModel.Name} 的PDF名稱怪怪的會跳過 #####");
            return "";
        }

        private void DelAllDelMePdf(string pdfFolder)
        {
            var delmes = Directory.GetFiles(pdfFolder).Where(x => System.IO.Path.GetFileName(x).Contains(config.GetValue<string>("TempPdfStr")));
            foreach (var delme in delmes)
            {
                try
                {
                    File.Delete(delme);
                }
                catch (Exception) { }
            }
        }

        //寫入PDF各欄位
        private void SetPdfFields(PdfDocument pdfDoc, IDictionary<string, PdfFormField> fields, PreachReportModel prModel, IndexModel model)
        {
            int pageNum = 0;
            string? year1 = fields[config.GetValue<string>("PDF_Year")].GetValueAsString();
            string? year2 = fields[config.GetValue<string>("PDF_Year") + "_2"].GetValueAsString();
            PdfFont pdfFont = MyFont(model.Font);
            if (string.IsNullOrEmpty(year1) || year1 == model.Year)
            {
                if (string.IsNullOrEmpty(year1))
                {
                    Console.WriteLine($"檢測到 {prModel.Name} 的PDF似乎是空白無輸入的PDF，已寫入資料但建議檢查一下({year1})");
                }
                pageNum = 1;
                fields[config.GetValue<string>("PDF_Year")].SetFontAndSize(pdfFont, 12f).SetValue(model.Year);
            }
            else if (string.IsNullOrEmpty(year2) || year2 == model.Year)
            {
                pageNum = 2;
                fields[config.GetValue<string>("PDF_Year") + "_2"].SetFontAndSize(pdfFont, 12f).SetValue(model.Year);
            }

            //9月是開頭1
            int month = int.Parse(model.Month);
            int fieldNum = month>=9 ? month-8 : month+4;

            string distribution = $"{pageNum}-{config.GetValue<string>("PDF_Distribution")}_{fieldNum}";
            fields[distribution].SetFontAndSize(pdfFont, 12f).SetValue(prModel.Distribution);

            string video = $"{pageNum}-{config.GetValue<string>("PDF_Video")}_{fieldNum}";
            fields[video].SetFontAndSize(pdfFont, 12f).SetValue(prModel.Video);

            string hour = $"{pageNum}-{config.GetValue<string>("PDF_Hours")}_{fieldNum}";
            fields[hour].SetFontAndSize(pdfFont, 12f).SetValue(prModel.Hour);

            string rv = $"{pageNum}-{config.GetValue<string>("PDF_RV")}_{fieldNum}";
            fields[rv].SetFontAndSize(pdfFont, 12f).SetValue(prModel.Review);

            string study = $"{pageNum}-{config.GetValue<string>("PDF_Study")}_{fieldNum}";
            fields[study].SetFontAndSize(pdfFont, 12f).SetValue(prModel.Study);

            //RemarksOctober_2 RemarksFebruary
            string remark = $"{config.GetValue<string>("PDF_Remark")}{model.Month.GetMonthEngName()}{(pageNum==2 ? "_2" : "")}";
            fields[remark].SetFontAndSize(pdfFont, 10f).SetValue(prModel.Remark);

            //foreach (var field in fields)
            //{
            //    Console.WriteLine("PDF field: " + field.Key.ToString());
            //}
        }

        //itext 7.1.14修改字體功能正常
        public PdfFont MyFont(byte[] bytes)
        {
            return PdfFontFactory.CreateFont(bytes, PdfEncodings.IDENTITY_H);
        }
    }
}
