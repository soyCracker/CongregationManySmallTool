using CommunityToolkit.Maui.Storage;
using CommunityToolkit.Mvvm.ComponentModel;
using Microsoft.AspNetCore.Components.Forms;
using SmallTool.Lib.Models.CPR_Maui_App;
using SmallTool.Lib.Services;

namespace CPR_Maui_App.ViewModels
{
    public partial class IndexViewModel : ObservableObject
    {
        [ObservableProperty]
        public IndexModel indexModel;
        private readonly CPRMauiService cprMauiService;
        public EditContext timeEditContext;

        public IndexViewModel(CPRMauiService cprMauiService)
        {
            this.cprMauiService = cprMauiService;
            indexModel = new IndexModel();
            timeEditContext = new EditContext(IndexModel);
            IndexModel.FontPath = Path.Combine(FileSystem.Current.AppDataDirectory, "TaipeiSansTCBeta-Regular.ttf");
        }

        public async Task Start()
        {
            IndexModel.Step = 2;
            await SetFont();
            await cprMauiService.Start(IndexModel);
            IndexModel.Step = 0;
        }

        public void SetInputYearAndMonth(string val)
        {
            var datetimeArr = val.Split('-');
            IndexModel.Year = datetimeArr[0];

            string tempMonth = datetimeArr[1];
            if (!string.IsNullOrEmpty(tempMonth) && tempMonth.Length==2 && tempMonth[0] == '0')
            {
                tempMonth = tempMonth.Substring(1);
            }

            IndexModel.Month = tempMonth;
        }

        public void GoSelectFolder()
        {
            if (timeEditContext.Validate())
            {
                IndexModel.Step = 1;
            }
        }

        public async Task SelectFolder(CancellationToken cancellationToken)
        {
            var result = await FolderPicker.Default.PickAsync(cancellationToken);
            if (result.IsSuccessful)
            {
                Console.WriteLine($"The folder was picked: Name - {result.Folder.Name}, Path - {result.Folder.Path}");
                IndexModel.RootFolder = result.Folder.Path;
            }
            else
            {
                Console.WriteLine($"The folder was not picked with error: {result.Exception.Message}");
            }
        }

        public async Task SetFont()
        {
            byte[] bytes;
            using var stream = await FileSystem.OpenAppPackageFileAsync("TaipeiSansTCBeta-Regular.ttf");
            using (var binaryReader = new BinaryReader(stream))
            {
                bytes = binaryReader.ReadBytes((int)stream.Length);
            }
            IndexModel.Font = bytes;
        }
    }
}
