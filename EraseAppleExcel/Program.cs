// See https://aka.ms/new-console-template for more information
using EraseAppleExcel;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using SmallTool.Lib.Services;
using SmallTool.Lib.Utils;

EnvirUtil envirUtil = new EnvirUtil();
IConfiguration configuration = envirUtil.ReadAppSetting();
envirUtil.InitEnvironment(configuration.GetValue<bool>("ReleaseMode"));

// 1. 建立依賴注入的容器
var service = new ServiceCollection();
// 2. 註冊服務
service.AddTransient<App>();
service.AddSingleton(configuration);
service.AddScoped<EraseAppleExcelService>();

// 建立依賴服務提供者
var serviceProvider = service.BuildServiceProvider();

// 3. 執行主服務
serviceProvider.GetRequiredService<App>().Run();