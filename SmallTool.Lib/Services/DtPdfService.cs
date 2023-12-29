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

namespace SmallTool.Lib.Services
{
    public class DtPdfService
    {
        private readonly IConfiguration config;

        public DtPdfService(IConfiguration config)
        {
            this.config = config;
        }

        //思源黑體 itext 7.1.5、7.1.14修改字體功能正常
        public PdfFont GetMyFont()
        {
            string path = System.IO.Path.Combine(config.GetValue<string>("FontFolder"), "TaipeiSansTCBeta-Regular.ttf");
            PdfFont font = PdfFontFactory.CreateFont(path, PdfEncodings.IDENTITY_H, true);
            return font;
        }

        public bool IsMyFontExist()
        {
            if (File.Exists(System.IO.Path.Combine(config.GetValue<string>("FontFolder"), "TaipeiSansTCBeta-Regular.ttf")))
            {
                return true;
            }
            return false;
        }

        //勾選框
        public void SetPdfCheckBoxSelected(IDictionary<string, PdfFormField> fields, PdfDocument pdfDoc, string fieldsKey)
        {
            foreach (PdfAnnotation annotation in fields[fieldsKey].GetWidgets())
            {
                PdfPage page = annotation.GetPage();
                Rectangle rectangle = annotation.GetRectangle().ToRectangle();
                Canvas canvas = new Canvas(new PdfCanvas(page), pdfDoc, rectangle);
                // 設定字體、粗體
                Paragraph p = new Paragraph("v").SetFont(GetMyFont()).SetBold();
                p.SetFixedPosition(rectangle.GetX(), rectangle.GetY() - 4, rectangle.GetWidth());
                canvas.Add(p);
                canvas.Close();
                page.RemoveAnnotation(annotation);
            }
        }

        //寫入文字在輸入框中間
        public void SetPdfFeldValueCenter(PdfDocument pdfDoc, string content, int left, int bottom)
        {
            PdfPage page = pdfDoc.GetPage(pdfDoc.GetNumberOfPages());
            //Rectangle rectangle = annotation.GetRectangle().ToRectangle();
            Canvas canvas = new Canvas(page, page.GetPageSize());
            // 設定字體、粗體
            Paragraph p = new Paragraph(content).SetFont(GetMyFont());
            p.SetFixedPosition(left, bottom, page.GetPageSize().GetWidth());
            canvas.Add(p);
            canvas.Close();
        }
    }
}
