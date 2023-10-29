using NPOI.SS.UserModel;

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
            if (cell==null)
            {
                return "";
            }
            var str = cell.ToString();
            if (string.IsNullOrEmpty(str))
            {
                return "";
            }
            return str.Trim();
        }

        /// <summary>
        /// get chinese or printable string
        /// </summary>
        /// <param name="source"></param>
        /// <returns>chinese or printable string</returns>
        public static string GetChinesePrintAble(string source)
        {
            try
            {
                string result = "";
                for (int i = 0; i < source.Length; i++)
                {
                    if (char.ConvertToUtf32(source, i) >= Convert.ToInt32("4e00", 16) &&
                        char.ConvertToUtf32(source, i) <= Convert.ToInt32("9fff", 16) ||
                        (source[i] >= 32 && source[i] <= 126))
                    {
                        result += source.Substring(i, 1);
                    }
                }
                return result;
            }
            catch (Exception)
            {
                return source;
            }
        }
    }
}
