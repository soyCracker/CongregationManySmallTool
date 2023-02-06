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
            Console.WriteLine("CongregationPreachReportService Start!");
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
                Console.WriteLine("XSSFWorkbook");
                workbook = new XSSFWorkbook(stream);
            }
            else
            {
                Console.WriteLine("HSSFWorkbook");
                workbook = new HSSFWorkbook(stream);
            }
            ISheet sheet = workbook.GetSheet(month + "月");

            List<PreachReportModel> preachReportModels = new List<PreachReportModel>();
            int startRow = config.GetValue<int>("CongregationXlsStartRow");
            string currentTeam = "";
            for (int r = startRow; r<sheet.LastRowNum; r++)
            {
                PreachReportModel prModel = new PreachReportModel();
                prModel.Name = StringUtil.SafeTrim(sheet.GetRow(r).GetCell(2).ToString());
                if (!string.IsNullOrEmpty(prModel.Name))
                {
                    string team = StringUtil.SafeTrim(sheet.GetRow(r).GetCell(0).ToString());
                    prModel.Team = ChkTeam(team, currentTeam);
                    currentTeam = prModel.Team;

                    string type = StringUtil.SafeTrim(sheet.GetRow(r).GetCell(1).ToString());
                    prModel.Type = type=="p" ? config.GetValue<string>("Pioneer") : config.GetValue<string>("Preacher");

                    prModel.Publish = StringUtil.SafeTrim(sheet.GetRow(r).GetCell(3).ToString());
                    prModel.Video = StringUtil.SafeTrim(sheet.GetRow(r).GetCell(4).ToString());
                    prModel.Hour = StringUtil.SafeTrim(sheet.GetRow(r).GetCell(5).ToString());
                    prModel.Review = StringUtil.SafeTrim(sheet.GetRow(r).GetCell(6).ToString());
                    prModel.Study = StringUtil.SafeTrim(sheet.GetRow(r).GetCell(7).ToString());
                    prModel.Remark = StringUtil.SafeTrim(sheet.GetRow(r).GetCell(8).ToString());

                    Console.WriteLine(string.Format("{0} {1} {2} {3} {4} {5} {6} {7} {8}", prModel.Name, prModel.Team,
                        prModel.Type, prModel.Publish, prModel.Video, prModel.Hour, prModel.Review, prModel.Study, prModel.Remark));
                    preachReportModels.Add(prModel);
                }
                Console.WriteLine("\n");
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
                string pdfPath = GetManPdf(prModel.Name, prModel.Type, prModel.Team);
                if (!string.IsNullOrEmpty(pdfPath))
                {
                    using FileStream fs = new FileStream(pdfPath, FileMode.Open);
                    using PdfDocument pdfDoc = new PdfDocument(new PdfReader(fs), new PdfWriter(pdfPath + ".new.pdf"));//
                    try
                    {
                        PdfAcroForm form = PdfAcroForm.GetAcroForm(pdfDoc, true);
                        IDictionary<string, PdfFormField> fields = form.GetFormFields();
                        SetPdfFields(pdfDoc, fields, prModel, input);
                        pdfDoc.Close();
                        fs.Close();
                        File.Delete(pdfPath);
                        File.Move(pdfPath + ".new.pdf", pdfPath);
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(prModel.Name + " 的作業發生錯誤，" + ex.ToString());
                        pdfDoc.Close();
                        fs.Close();
                        if (File.Exists(pdfPath + ".new.pdf"))
                        {
                            File.Delete(pdfPath + ".new.pdf");
                        }
                    }
                }
            }
        }

        private string GetManPdf(string name, string type, string team)
        {
            var rootFolder = config.GetValue<string>("TargetRootFolder");
            var firstLayerFolder = Directory.GetDirectories(rootFolder)[0];
            var pdf = Directory.GetFiles(System.IO.Path.Combine(firstLayerFolder, team, type, name)).Max();
            Console.WriteLine(pdf);
            string onlyFileName = System.IO.Path.GetFileName(pdf);
            if (!onlyFileName.EndsWith(".pdf") || onlyFileName.Length!=11)
            {
                Console.WriteLine(name + "的PDF怪怪的會跳過");
                return "";
            }
            return pdf;
        }

        private void SetPdfFields(PdfDocument pdfDoc, IDictionary<string, PdfFormField> fields, PreachReportModel prModel, InputModel input)
        {
            int pageNum = 0;
            string year1 = fields[config.GetValue<string>("PDF_Year")].GetValue().ToString();
            string year2 = fields[config.GetValue<string>("PDF_Year") + "_2"].GetValue().ToString();
            if (string.IsNullOrEmpty(year1) || year1 == input.Year)
            {
                pageNum = 1;
            }
            else if (string.IsNullOrEmpty(year2) || year2 == input.Year)
            {
                pageNum = 2;
            }

            //9月是開頭1
            int month = int.Parse(input.Month);
            int fieldNum = month>=9 ? month-8 : month+4;
            string fieldName = pageNum + "-" + config.GetValue<string>("PDF_Publish") + "_" + fieldNum;
            fields[fieldName].SetFont(GetMsjhbdFont()).SetValue(prModel.Publish);
            //SetPdfFeldValueCenter(pdfDoc, fields, pageNum + "-" + config.GetValue<string>("PDF_Publish") + "_" + fieldNum, prModel.Publish);

            //foreach (var field in fields)
            //{
            //    Console.WriteLine("PDF field: " + field.Key.ToString());
            //}
        }

        //微軟正黑體 itext 7.1.5修改字體功能正常
        public PdfFont GetMsjhbdFont()
        {
            string path = System.IO.Path.Combine("Font", "msjhbd.ttc,0");
            return PdfFontFactory.CreateFont(path, PdfEncodings.IDENTITY_H, true);
        }

        //寫入文字在輸入框中間
        //public void SetPdfFeldValueCenter(PdfDocument pdfDoc, IDictionary<string, PdfFormField> fields, string fieldsKey, string content)
        //{
        //    foreach (PdfAnnotation annotation in fields[fieldsKey].GetWidgets())
        //    {
        //        PdfPage page = annotation.GetPage();
        //        Rectangle rectangle = annotation.GetRectangle().ToRectangle();
        //        Canvas canvas = new Canvas(new PdfCanvas(page), pdfDoc, rectangle);
        //        Paragraph p = new Paragraph(content).SetFont(GetMsjhbdFont());
        //        p.SetFixedPosition(rectangle.GetX() + (rectangle.GetWidth() / 3), rectangle.GetY() - 4, rectangle.GetWidth()).SetFont(GetMsjhbdFont());
        //        canvas.Add(p);
        //        canvas.Close();
        //        page.RemoveAnnotation(annotation);
        //    }
        //}
    }
}