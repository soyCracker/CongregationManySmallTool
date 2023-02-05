using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.Util;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;

namespace SmallTool.Lib.Services
{
    public class EraseAppleExcelService
    {
        public EraseAppleExcelService()
        {

        }

        public void Start(string source, string outputFolder)
        {
            FileStream stream = new FileStream(source, FileMode.Open);
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

                for (int r = 0; r<sheet.LastRowNum; r++)
                {
                    for (int c = 0; c<sheet.GetRow(r).LastCellNum; c++)
                    {
                        string cellStr = sheet.GetRow(r).GetCell(c).ToString().Trim();
                        string[] names = cellStr.Split(' ', StringSplitOptions.RemoveEmptyEntries);
                        string newStr = "";
                        for (int i = 0; i<names.Length; i++)
                        {
                            newStr += names[i];
                            if (i<names.Length-1)
                            {
                                newStr+='\n';
                            }
                        }
                        sheet.GetRow(r).GetCell(c).SetCellValue(newStr);
                    }
                }

                FileStream result = new FileStream(outputFolder + "/new.xlsx", FileMode.Create);
                workbook.Write(result, false);
                workbook.Close();

            }
        }
    }
}
