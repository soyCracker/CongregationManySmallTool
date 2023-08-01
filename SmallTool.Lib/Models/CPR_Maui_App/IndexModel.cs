using CommunityToolkit.Mvvm.ComponentModel;
using System.ComponentModel.DataAnnotations;

namespace SmallTool.Lib.Models.CPR_Maui_App
{
    public class IndexModel : ObservableObject
    {
        [Required(ErrorMessage = "年份，錯誤格式")]
        public string? Year { get; set; }
        [Required(ErrorMessage = "月份，錯誤格式")]
        public string? Month { get; set; }
        public string? RootFolder { get; set; }
        public int Step { get; set; } = 0;
        [Required(ErrorMessage = "字體缺失，嚴重錯誤")]
        public string? FontPath { get; set; }
        public byte[] Font { get; set; }
    }
}
