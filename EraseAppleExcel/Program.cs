// See https://aka.ms/new-console-template for more information
using EraseAppleExcel.Base;
using SmallTool.Lib.Services;
using SmallTool.Lib.Utils;

Console.WriteLine("Hello, World!");


EnvirUtil envirUtil = new EnvirUtil();
envirUtil.InitEnvironment(false);

string xlsx = envirUtil.PrepareAndGetTempXlsx(Config.FILE_FOLDER);
EraseAppleExcelService eaeService = new EraseAppleExcelService();
eaeService.Start(xlsx, Config.OUTPUT_FOLDER);