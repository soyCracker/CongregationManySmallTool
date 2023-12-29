using iText.Forms;
using iText.Forms.Fields;
using iText.Kernel.Pdf;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using SmallTool.Lib.Exceptions.DelegationTool;
using SmallTool.Lib.Models.DelegationTool;
using SmallTool.Lib.Utils;

namespace SmallTool.Lib.Services
{
    public class DtExportService
    {
        private string blockDate = "";
        private string lastDelagation = "";
        private string lastDate = "";
        private int sameDelegationCount = 1;
        private readonly DtPdfService pdfService;

        public DtExportService(DtPdfService pdfService)
        {
            this.pdfService = pdfService;
        }

        public Dictionary<string, byte[]> Start(string formFile, string s89chFile, string s89jpFile,
            string descStr, string descJPStr, string JPFlagStr)
        {
            FileStream fs = new FileStream(formFile, FileMode.Open);
            return Process(fs, s89chFile, s89jpFile, descStr, descJPStr, JPFlagStr);
        }

        private Dictionary<string, byte[]> Process(FileStream stream, string s89chFile, string s89jpFile,
            string descStr, string descJPStr, string JPFlagStr)
        {
            //DateTime start = DateTime.Now;
            //Console.WriteLine("Export Process start:" + start);
            try
            {
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
                    ISheet sheet = workbook.GetSheetAt(0);
                    Dictionary<string, byte[]> dict = new Dictionary<string, byte[]>();
                    for (int j = 1; j <= 2; j++)
                    {
                        for (int i = 4; i <= sheet.LastRowNum; i++)
                        {
                            //DateTime ls = DateTime.Now;
                            //Console.WriteLine("Delegation Read & Export start:" + ls);
                            var vm = GetEachDelegation(sheet, i, j);
                            if (vm != null)
                            {

                                var tuple = ExportDelegation(vm, s89chFile, s89jpFile, descStr,
                                    descJPStr, JPFlagStr);
                                dict.Add(tuple.Item1, tuple.Item2);
                            }
                            //DateTime le = DateTime.Now;
                            //Console.WriteLine("Delegation Read & Export end:" + le);
                            //Console.WriteLine("Delegation Read & Export diff:" + le.Subtract(ls));
                        }
                    }
                    workbook.Close();
                    //DateTime end = DateTime.Now;
                    //Console.WriteLine("Export Process end:" + end);
                    //Console.WriteLine("Export Process diff:" + end.Subtract(start));
                    return dict;
                }
            }
            catch (IOException)
            {
                throw;
            }
        }

        private DelegationVM GetEachDelegation(ISheet sheet, int rowNum, int classInt)
        {
            try
            {
                DelegationVM vm = new DelegationVM();
                vm.Name = sheet.GetRow(rowNum).GetCell(2 + (classInt - 1) * 2).SafeTrim();
                if (string.IsNullOrEmpty(vm.Name))
                {
                    return null;
                }

                vm.Assistant = sheet.GetRow(rowNum).GetCell(3 + (classInt - 1) * 2).SafeTrim();
                vm.Header = StringUtil.GetChinesePrintAble(sheet.GetRow(rowNum).GetCell(1).SafeTrim());
                vm.Class = classInt.ToString();
                vm.Date = sheet.GetRow(rowNum).GetCell(0).SafeTrim();

                if (string.IsNullOrEmpty(vm.Date))
                {
                    vm.Date = blockDate;
                }
                else
                {
                    blockDate = vm.Date;
                }
                return vm;
            }
            catch (Exception)
            {
                return null;
            }
        }

        private Tuple<string, byte[]> ExportDelegation(DelegationVM vm, string s89chFile,
            string s89jpFile, string descStr, string descJPStr, string JPFlagStr)
        {
            //DateTime start = DateTime.Now;
            //Console.WriteLine("ExportDelegation start:" + start);
            Console.WriteLine("-------------------------------\n");
            string s89 = s89chFile;
            string description = descStr;
            //識別日文委派
            if (vm.Name.Contains(JPFlagStr))
            {
                s89 = s89jpFile;
                description = description + descJPStr;
                //識別日文委派的字符替換掉
                vm.Name = vm.Name.Replace(JPFlagStr, "");
            }
            Tuple<string, byte[]> tuple = WritePdf(s89, vm, description);
            //DateTime end = DateTime.Now;
            //Console.WriteLine("ExportDelegation end:" + end);
            //Console.WriteLine("ExportDelegation diff:" + end.Subtract(start));
            return tuple;
        }

        private Tuple<string, byte[]> WritePdf(string s89, DelegationVM delegation, string description)
        {
            using (FileStream fs = new FileStream(s89, FileMode.Open))
            {
                MemoryStream ms = new MemoryStream();
                PdfWriter pdfWriter = new PdfWriter(ms);
                pdfWriter.SetCloseStream(false);
                PdfDocument pdfDoc = new PdfDocument(new PdfReader(fs), pdfWriter);
                PdfAcroForm form = PdfAcroForm.GetAcroForm(pdfDoc, true);
                IDictionary<string, PdfFormField> fields = form.GetFormFields();
                SetPdfField(fields, delegation, pdfDoc);
                pdfDoc.Close();
                string fileName = Path.Combine(TimeUtil.CovertDateToFileNameStr(delegation.Date) + description
                    + delegation.Name + ".pdf");
                byte[] data = ms.ToArray();
                ms.Close();
                return Tuple.Create(fileName, data);
            }
        }

        private void SetPdfField(IDictionary<string, PdfFormField> fields, DelegationVM delegation, PdfDocument pdfDoc)
        {
            if (pdfService.IsMyFontExist())
            {
                //我她X的，SetFont要在SetValue之前
                Console.WriteLine("學生:" + delegation.Name + "\n");
                //pdfService.SetPdfFeldValueCenter(fields, pdfDoc, S89PdfField.Name2, delegation.Name);
                pdfService.SetPdfFeldValueCenter(pdfDoc, delegation.Name, 100, 270);

                Console.WriteLine("助手:" + delegation.Assistant + "\n");
                //pdfService.SetPdfFeldValueCenter(fields, pdfDoc, S89PdfField.Ass2, delegation.Assistant);
                pdfService.SetPdfFeldValueCenter(pdfDoc, delegation.Assistant, 100, 245);

                //日期
                pdfService.SetPdfFeldValueCenter(pdfDoc, delegation.Date, 100, 220);

                //題目
                pdfService.SetPdfFeldValueCenter(pdfDoc, delegation.Header, 100, 190);
            }
            else
            {
                throw new NoFontException();
            }
        }
    }
}
