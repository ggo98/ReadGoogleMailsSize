using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace readgooglemailssize
{
    public static class StringExtensions
    {
        public static string SmartSubString(this string s, int startIndex, int length)
        {
            int slen = s.Length;
            if (startIndex > slen)
                return string.Empty;
            int len = slen - startIndex;
            if (len > length)
                len = length;
            if (len < 0)
                return string.Empty;
            string ret = s.Substring(startIndex, len);
            return ret;
        }
    }
}
