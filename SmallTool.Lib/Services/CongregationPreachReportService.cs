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
using SmallTool.Lib.Utils;

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
            Console.WriteLine("\n##### CongregationPreachReportService Start! #####\n");
            Console.WriteLine("##### 確認全會眾傳道報告檔案 #####\n");
            string congregationXls = GetCongregationXls(input.Year);
            Console.WriteLine("##### 讀取所有人的報告 #####\n");
            List<PreachReportModel> preachReportModels = ReadAllPreachReport(congregationXls, input.Month);
            Console.WriteLine("##### 開始寫入PDF #####\n");
            WriteToEveryonePdf(preachReportModels, input);
        }

        //確認全會眾傳道報告檔案
        private string GetCongregationXls(string year)
        {
            var rootFolder = config.GetValue<string>("TargetRootFolder");
            var firstLayerFolder = Directory.GetDirectories(rootFolder)[0];
            var wholeCongregationFolder = config.GetValue<string>("WholeCongregationFolder");
            var yearFolder = Directory.GetDirectories(System.IO.Path.Combine(new string[] { firstLayerFolder, wholeCongregationFolder })).FirstOrDefault(x => x.Contains(year));
            var xls = Directory.GetFiles(yearFolder).Where(f => f.EndsWith(".xls") || f.EndsWith(".xlsx")).FirstOrDefault();
            Console.WriteLine($"全會眾傳道報告檔案: {xls}\n");
            return xls;
        }

        //讀取所有人的報告
        private List<PreachReportModel> ReadAllPreachReport(string xls, string month)
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

            List<PreachReportModel> prModels = GetEveryoneBasicData();
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
                        //Console.WriteLine($"{prModel.Name}, {prModel.Team}, {prModel.Type}, {prModel.Distribution}, {prModel.Video}, {prModel.Hour}, "
                        //    + $"{prModel.Review}, {prModel.Study}, {prModel.Remark}\n");
                    }
                }
            }
            workbook.Close();
            return prModels;
        }

        //先從資料夾結構讀取所有人的名字、所屬小組、傳道員類別
        private List<PreachReportModel> GetEveryoneBasicData()
        {
            var rootFolder = config.GetValue<string>("TargetRootFolder");
            var firstLayerFolder = Directory.GetDirectories(rootFolder)[0];
            var teams = Directory.GetDirectories(firstLayerFolder).Where(x => x != config.GetValue<string>("WholeCongregationFolder"));
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
        private void WriteToEveryonePdf(List<PreachReportModel> prModels, InputModel input)
        {
            foreach (var prModel in prModels)
            {
                string pdfPath = GetManPdf(prModel, input.Year);
                if (!string.IsNullOrEmpty(pdfPath))
                {
                    using FileStream fs = new FileStream(pdfPath, FileMode.Open);
                    using PdfDocument pdfDoc = new PdfDocument(new PdfReader(fs), new PdfWriter(pdfPath + config.GetValue<string>("TempPdfStr")));
                    try
                    {
                        PdfAcroForm form = PdfAcroForm.GetAcroForm(pdfDoc, true);
                        IDictionary<string, PdfFormField> fields = form.GetFormFields();
                        SetPdfFields(pdfDoc, fields, prModel, input);
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

        private string GetManPdf(PreachReportModel prModel, string year)
        {
            var rootFolder = config.GetValue<string>("TargetRootFolder");
            var firstLayerFolder = Directory.GetDirectories(rootFolder)[0];
            var pdfFolder = System.IO.Path.Combine(firstLayerFolder, prModel.Team, prModel.Type, prModel.Name);
            if (Directory.Exists(pdfFolder))
            {
                var pdf = Directory.GetFiles(pdfFolder).Where(x => x.EndsWith(".pdf") && System.IO.Path.GetFileName(x).Length==11
                            && x.Contains(year.Substring(2, 2))).Max();
                if (!string.IsNullOrEmpty(pdf))
                {
                    Console.WriteLine(".");
                    //Console.WriteLine(pdf);
                    return pdf;
                }
            }
            Console.WriteLine($"{prModel.Team} {prModel.Type} {prModel.Name} 的PDF名稱怪怪的會跳過");
            return "";
        }

        //寫入PDF各欄位
        private void SetPdfFields(PdfDocument pdfDoc, IDictionary<string, PdfFormField> fields, PreachReportModel prModel, InputModel input)
        {
            int pageNum = 0;
            string? year1 = fields[config.GetValue<string>("PDF_Year")].GetValueAsString();
            string? year2 = fields[config.GetValue<string>("PDF_Year") + "_2"].GetValueAsString();
            if (string.IsNullOrEmpty(year1) || year1 == input.Year)
            {
                if (string.IsNullOrEmpty(year1))
                {
                    Console.WriteLine($"檢測到 {prModel.Name} 的PDF似乎是空白無輸入的PDF，已寫入資料但建議檢查一下({year1})");
                }
                pageNum = 1;
                fields[config.GetValue<string>("PDF_Year")].SetFontAndSize(MyFont(), 12f).SetValue(input.Year);
            }
            else if (string.IsNullOrEmpty(year2) || year2 == input.Year)
            {
                pageNum = 2;
                fields[config.GetValue<string>("PDF_Year") + "_2"].SetFontAndSize(MyFont(), 12f).SetValue(input.Year);
            }

            //9月是開頭1
            int month = int.Parse(input.Month);
            int fieldNum = month>=9 ? month-8 : month+4;

            string distribution = $"{pageNum}-{config.GetValue<string>("PDF_Distribution")}_{fieldNum}";
            fields[distribution].SetFontAndSize(MyFont(), 12f).SetValue(prModel.Distribution);

            string video = $"{pageNum}-{config.GetValue<string>("PDF_Video")}_{fieldNum}";
            fields[video].SetFontAndSize(MyFont(), 12f).SetValue(prModel.Video);

            string hour = $"{pageNum}-{config.GetValue<string>("PDF_Hours")}_{fieldNum}";
            fields[hour].SetFontAndSize(MyFont(), 12f).SetValue(prModel.Hour);

            string rv = $"{pageNum}-{config.GetValue<string>("PDF_RV")}_{fieldNum}";
            fields[rv].SetFontAndSize(MyFont(), 12f).SetValue(prModel.Review);

            string study = $"{pageNum}-{config.GetValue<string>("PDF_Study")}_{fieldNum}";
            fields[study].SetFontAndSize(MyFont(), 12f).SetValue(prModel.Study);

            //RemarksOctober_2 RemarksFebruary
            string remark = $"{config.GetValue<string>("PDF_Remark")}{input.Month.GetMonthEngName()}{(pageNum==2 ? "_2" : "")}";
            fields[remark].SetFontAndSize(MyFont(), 10f).SetValue(prModel.Remark);

            //foreach (var field in fields)
            //{
            //    Console.WriteLine("PDF field: " + field.Key.ToString());
            //}
        }

        //itext 7.1.14修改字體功能正常
        public PdfFont MyFont()
        {
            string path = System.IO.Path.Combine(config.GetValue<string>("FontDir"), "TaipeiSansTCBeta-Regular.ttf");
            return PdfFontFactory.CreateFont(path, PdfEncodings.IDENTITY_H);
        }
    }
}