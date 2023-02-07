using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SmallTool.Lib.Utils
{
    public static class StringUtil
    {
        public static string SafeTrim(string str)
        {
            if (string.IsNullOrEmpty(str))
            {
                return "";
            }
            return str.Trim();
        }

        public static string SafeTrim(this ICell cell)
        {
            var str = cell.ToString();
            if (string.IsNullOrEmpty(str))
            {
                return "";
            }
            return str.Trim();
        }
    }
}
