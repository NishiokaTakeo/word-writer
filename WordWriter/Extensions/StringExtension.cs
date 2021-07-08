using System;
using System.Text.RegularExpressions;

namespace WordWriter.Extensions
{
    public static class StringExtension
    {
        /// <summary>
        /// Remove all white space without inside of text.
        /// </summary>
        /// <param name="str"></param>
        /// <returns></returns>
        public static string TrimWhiteSpace(this String str)
        {
            string reped = (new Regex(@"^(\s|\t|\n|\b)+")).Replace(str, "");
            
            reped  = (new Regex(@"(\s|\t|\n|\b)+$")).Replace(reped, "");

            return reped;
        }

        public static System.Collections.Generic.IEnumerable<string> SplitByLength(this string str, int maxLength)
        {
            for (int index = 0; index < str.Length; index += maxLength)
            {
                yield return str.Substring(index, Math.Min(maxLength, str.Length - index));
            }
        }
    }   


}