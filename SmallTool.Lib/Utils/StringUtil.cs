using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SmallTool.Lib.Utils
{
    public class StringUtil
    {
        public static string SafeTrim(string str)
        {
            if (string.IsNullOrEmpty(str))
            {
                return "";
            }
            return str.Trim();
        }
    }
}
