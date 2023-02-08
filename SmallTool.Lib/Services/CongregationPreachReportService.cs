using iText.Forms;
using iText.Forms.Fields;
using iText.IO.Font;
using iText.Kernel.Font;
using iText.Kernel.Geom;
using iText.Kernel.Pdf;
using iText.Kernel.Pdf.Annot;
using iText.Kernel.Pdf.Canvas;
using iText.Layout;
using iText.Layout.Element;
using MathNet.Numerics.Distributions;
using Microsoft.Extensions.Configuration;
using NPOI.HSSF.EventUserModel.DummyRecord;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using SmallTool.Lib.Models.CongregationPreachReport;
using SmallTool.Lib.Utils;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static NPOI.HSSF.Util.HSSFColor;

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
            Console.WriteLine("CongregationPreachReportService Start!\n\n");
            string congregationXls = GetCongregationXls(input.Year);
            List<PreachReportModel> preachReportModels = ReadAllPreachReport(congregationXls, input.Month);
            WriteToEveryonePdf(preachReportModels, input);
        }

        private string GetCongregationXls(string year)
        {
            var rootFolder = config.GetValue<string>("TargetRootFolder");
            var firstLayerFolder = Directory.GetDirectories(rootFolder)[0];
            var wholeCongregationFolder = config.GetValue<string>("WholeCongregationFolder");
            var yearFolder = Directory.GetDirectories(System.IO.Path.Combine(new string[] { firstLayerFolder, wholeCongregationFolder })).FirstOrDefault(x => x.Contains(year));
            var xls = Directory.GetFiles(yearFolder).Where(f => f.EndsWith(".xls") || f.EndsWith(".xlsx")).FirstOrDefault();
            Console.WriteLine($"全會眾傳道報告檔案: {xls}");
            return xls;
        }

        private List<PreachReportModel> ReadAllPreachReport(string xls, string month)
        {
            using FileStream stream = new FileStream(xls, FileMode.Open);
            IWorkbook workbook;
            // 避免報異常EOF in header
            stream.Position = 0;
            if (stream.Name.ToLower().EndsWith(".xlsx"))
            {
                Console.WriteLine("Excel type: xlsx XSSFWorkbook");
                workbook = new XSSFWorkbook(stream);
            }
            else
            {
                Console.WriteLine("Excel type: xls HSSFWorkbook");
                workbook = new HSSFWorkbook(stream);
            }
            ISheet sheet = workbook.GetSheet(month + "月");

            List<PreachReportModel> preachReportModels = new List<PreachReportModel>();
            int startRow = config.GetValue<int>("CongregationXlsStartRow");
            string currentTeam = "";
            for (int r = startRow; r<sheet.LastRowNum; r++)
            {
                PreachReportModel prModel = new PreachReportModel();
                prModel.Name = sheet.GetRow(r).GetCell(2).SafeTrim();
                if (!string.IsNullOrEmpty(prModel.Name))
                {
                    string team = sheet.GetRow(r).GetCell(0).SafeTrim();
                    prModel.Team = ChkTeam(team, currentTeam);
                    currentTeam = prModel.Team;

                    string type = sheet.GetRow(r).GetCell(1).SafeTrim();
                    prModel.Type = type=="p" ? config.GetValue<string>("Pioneer") : config.GetValue<string>("Preacher");

                    prModel.Distribution = sheet.GetRow(r).GetCell(3).SafeTrim();
                    prModel.Video = sheet.GetRow(r).GetCell(4).SafeTrim();
                    prModel.Hour = sheet.GetRow(r).GetCell(5).SafeTrim();
                    prModel.Review = sheet.GetRow(r).GetCell(6).SafeTrim();
                    prModel.Study = sheet.GetRow(r).GetCell(7).SafeTrim();
                    prModel.Remark = sheet.GetRow(r).GetCell(8).SafeTrim();

                    /*Console.WriteLine(string.Format("{0} {1} {2} {3} {4} {5} {6} {7} {8}", prModel.Name, prModel.Team,
                        prModel.Type, prModel.Distribution, prModel.Video, prModel.Hour, prModel.Review, prModel.Study, prModel.Remark));*/
                    preachReportModels.Add(prModel);
                    //Console.WriteLine("\n");
                }
            }
            workbook.Close();
            return preachReportModels;
        }

        private string ChkTeam(string team, string lastTeam)
        {
            if (!string.IsNullOrEmpty(team))
            {
                string[] teamArr = config.GetSection("Teams").Get<string[]>();
                if (teamArr.Contains(team))
                {
                    return team;
                }
            }
            return lastTeam;
        }

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
                        Console.WriteLine(prModel.Name + " 的作業發生錯誤，" + ex.ToString());
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