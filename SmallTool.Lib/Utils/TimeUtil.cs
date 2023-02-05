using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SmallTool.Lib.Utils
{
    public class TimeUtil
    {
        public static string GetTimeNow()
        {
            return DateTime.Now.ToString("yyyyMMddHHmmss");
        }

        public static string CovertDateToFileNameStr(string dateStr)
        {
            return DateTime.Parse(dateStr).ToString("yyyyMMdd");
        }
    }
}
