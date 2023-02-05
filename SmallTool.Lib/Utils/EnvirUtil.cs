using Microsoft.Extensions.Configuration;
using System.Text;

namespace SmallTool.Lib.Utils
{
    public class EnvirUtil
    {
        public void InitEnvironment(bool releaseMode)
        {
            if (!releaseMode)
            {
                Directory.SetCurrentDirectory(Directory.GetCurrentDirectory() + "../../../../");
                Console.WriteLine("work dir:" + Directory.GetCurrentDirectory() + "\n");
            }
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
            Console.OutputEncoding = Encoding.Unicode;
        }

        public string PrepareAndGetTempXlsx(string fileFolder)
        {
            string[] files = Directory.GetFiles(fileFolder).Where(f => f.EndsWith(".xls") || f.EndsWith(".xlsx")).ToArray();
            if (files.Length > 1)
            {
                //throw new MultipleXlsException();
            }
            else if (files.Length == 1)
            {
                string file = files[0];
                return file;
            }
            return "";
        }

        public string CreateOutputFolder(string outputFolder)
        {
            string[] pathParam = { outputFolder, TimeUtil.GetTimeNow() };
            string targetPath = Path.Combine(pathParam);
            if (!Directory.Exists(targetPath))
            {
                Directory.CreateDirectory(targetPath);
            }
            return targetPath;
        }

        public IConfiguration ReadAppSetting()
        {
            return new ConfigurationBuilder()
                .AddJsonFile("appsettings.json", optional: true, reloadOnChange: true)
                .Build();
        }
    }
}
