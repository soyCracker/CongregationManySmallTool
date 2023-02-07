using System.Globalization;

namespace SmallTool.Lib.Utils
{
    public static class TimeUtil
    {
        public static string GetTimeNow()
        {
            return DateTime.Now.ToString("yyyyMMddHHmmss");
        }

        public static string CovertDateToFileNameStr(string dateStr)
        {
            return DateTime.Parse(dateStr).ToString("yyyyMMdd");
        }

        public static string GetMonthEngName(this string monthNum)
        {
            return DateTime.Parse($"2023-{monthNum}-1").ToString("MMMM", CultureInfo.GetCultureInfo("en-us"));
        }
    }
}
