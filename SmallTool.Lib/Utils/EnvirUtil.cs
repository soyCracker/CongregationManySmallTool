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
                //Console.WriteLine("work dir:" + Directory.GetCurrentDirectory() + "\n");
            }
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
            Console.OutputEncoding = Encoding.UTF8;
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
            //讀取環境設定，讀不同的appsettings
            var environmentName = Environment.GetEnvironmentVariable("ASPNETCORE_ENVIRONMENT");
            return new ConfigurationBuilder()
                .AddJsonFile("appsettings.json", optional: true, reloadOnChange: true)
                .AddJsonFile($"appsettings.{environmentName}.json", true, true)
                .AddEnvironmentVariables()
                .Build();
        }
    }
}
