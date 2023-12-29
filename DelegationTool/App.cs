using Microsoft.Extensions.Configuration;
using SmallTool.Lib.Exceptions.DelegationTool;
using SmallTool.Lib.Services;
using SmallTool.Lib.Utils;

namespace DelegationTool
{
    public class App
    {
        private readonly IConfiguration config;
        private string delegationFormFolder = "";
        private string assignmentFolder = "";
        private string s89chFile = "";
        private string s89jpFile = "";
        private readonly DtExportService exportService;
        private readonly DtRecordService recordService;

        public App(IConfiguration config, DtExportService exportService, DtRecordService recordService)
        {
            this.config = config;
            this.exportService = exportService;
            this.recordService = recordService;
        }

        public void Run()
        {
            delegationFormFolder = Path.Combine(config.GetValue<string>("FileFolder"), config.GetValue<string>("DelegationFormFolder"));
            assignmentFolder = Path.Combine(config.GetValue<string>("FileFolder"), config.GetValue<string>("AssignmentFolder"));
            s89chFile = Path.Combine(config.GetValue<string>("FileFolder"), config.GetValue<string>("S89CH"));
            if (File.Exists(s89chFile))
            {
                Console.WriteLine("YESSSSSSSSSSSS");
            }
            else
            {
                Console.WriteLine("NOOOOOOOOOOOOOOO");
            }
            s89jpFile = Path.Combine(config.GetValue<string>("FileFolder"), config.GetValue<string>("S89J"));

            while (true)
            {
                FucDescOutput();
                string fucKey = Console.ReadLine();
                Fuchub(fucKey);
            }
        }

        private void FucDescOutput()
        {
            Console.WriteLine($"DelegationTool 委派單程式 V{config.GetValue<string>("Version")}");
            Console.WriteLine($"Release Mode: {config.GetValue<bool>("ReleaseMode")}");
            Console.WriteLine("Delegation Console Tool 功能:");
            Console.WriteLine("委派紀錄填寫 - 1");
            Console.WriteLine("輸出委派單 - 3");
            Console.Write("輸入功能:");
        }

        private void Fuchub(string fucKey)
        {
            if (fucKey == "1")
            {
                RecordWork();
            }
            else if (fucKey == "3")
            {
                ExportWork();
            }
            else if (fucKey == "99")
            {
                TestWork();
            }
        }

        private void RecordWork()
        {
            Console.WriteLine("委派紀錄填寫");
            string outputFolder = CreateOutputFolder(config.GetValue<string>("OutputFolder"));
            string delegationXlsx = PrepareAndGetTempXlsx(delegationFormFolder);
            string assignRecordXlsx = PrepareAndGetTempXlsx(assignmentFolder);
            byte[] data = recordService.Start(delegationXlsx, assignRecordXlsx);
            File.Delete(delegationXlsx);
            File.Delete(assignRecordXlsx);
            using (FileStream fs = new FileStream(Path.Combine(outputFolder, "assignmentNext.xlsx"), FileMode.Create))
            {
                fs.Write(data, 0, data.Length);
            }
        }

        private void ExportWork()
        {
            try
            {
                Console.WriteLine("輸出委派單");
                string xlsx = PrepareAndGetTempXlsx(delegationFormFolder);
                if (string.IsNullOrEmpty(xlsx))
                {
                    Console.WriteLine(Path.GetFullPath(delegationFormFolder) + " 沒有委派單可輸出\n");
                    Console.WriteLine(Directory.GetCurrentDirectory() + " 沒有委派單可輸出\n");
                }
                else
                {
                    //ExportService exportService = new ExportService(new PDFService(Constant.FONT_FOLDER));
                    Dictionary<string, byte[]> dict = exportService.Start(xlsx, s89chFile,
                        s89jpFile, config.GetValue<string>("DescStr"), config.GetValue<string>("DescJP_Str"), config.GetValue<string>("JP_FlagStr"));
                    string outputFolder = CreateOutputFolder(config.GetValue<string>("OutputFolder"));
                    SavePdfDict(dict, outputFolder);
                }
                File.Delete(xlsx);
            }
            catch (IOException ex)
            {
                Console.WriteLine("可能excel或pdf檔案不存在、config內的檔名填錯，或是檔案正被使用中請關閉所有檔案再試一次\n");
                Console.WriteLine(ex + "\n");
            }
            catch (NoFontException ex)
            {
                Console.WriteLine("字型檔案不存在，請檢查Font資料夾或重新下載程式\n");
                Console.WriteLine(ex + "\n");
            }
            catch (MultipleXlsException ex)
            {
                Console.WriteLine("存在多個xls或xlsx於資料夾中\n");
                Console.WriteLine(ex + "\n");
            }
            catch (Exception ex)
            {
                Console.WriteLine("未知錯誤，請複製以下資訊給開發者:" + ex + "\n");
            }
        }

        private void SavePdfDict(Dictionary<string, byte[]> dict, string outputFolder)
        {
            foreach (var item in dict)
            {
                using (FileStream fs = new FileStream(Path.Combine(outputFolder, item.Key), FileMode.Create))
                {
                    fs.Write(item.Value, 0, item.Value.Length);
                }
            }
        }

        private static void TestWork()
        {
            Console.WriteLine("TestWork() !!!\n");
        }

        private string PrepareAndGetTempXlsx(string fileFolder)
        {
            string[] files = Directory.GetFiles(fileFolder).Where(f => f.EndsWith(".xls") || f.EndsWith(".xlsx")).ToArray();
            if (files.Length > 1)
            {
                throw new MultipleXlsException();
            }
            else if (files.Length == 1)
            {
                string file = files[0];
                string fileExt = file.ToLower().EndsWith("xlsx") ? "xlsx" : "xls";
                string tempXls = string.Format("{0}_temp.{1}", file, fileExt);
                if (File.Exists(tempXls))
                {
                    File.Delete(tempXls);
                }
                File.Copy(file, tempXls);
                Console.WriteLine("tempXls:" + tempXls + "\n");
                return tempXls;
            }
            return "";
        }

        private string CreateOutputFolder(string outputFolder)
        {
            string[] pathParam = { outputFolder, TimeUtil.GetTimeNow() };
            string targetPath = Path.Combine(pathParam);
            if (!Directory.Exists(targetPath))
            {
                Directory.CreateDirectory(targetPath);
            }
            return targetPath;
        }
    }
}
