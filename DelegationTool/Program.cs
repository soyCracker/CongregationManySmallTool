using DelegationTool;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using SmallTool.Lib.Services;
using SmallTool.Lib.Utils;

try
{
    EnvirUtil envirUtil = new EnvirUtil();
    IConfiguration configuration = envirUtil.ReadAppSetting();
    envirUtil.InitEnvironment(configuration.GetValue<bool>("ReleaseMode"));

    // 1. 建立依賴注入的容器
    var service = new ServiceCollection();
    // 2. 註冊服務
    service.AddTransient<App>();
    service.AddSingleton(configuration);
    service.AddScoped<DtExportService>();
    service.AddScoped<DtRecordService>();
    service.AddScoped<DtPdfService>();

    // 建立依賴服務提供者
    var serviceProvider = service.BuildServiceProvider();

    // 3. 執行主服務
    serviceProvider.GetRequiredService<App>().Run();
}
catch (Exception ex)
{
    Console.WriteLine($"\n\n遭遇致命錯誤: {ex}\n\n");
    Console.ReadLine();
}