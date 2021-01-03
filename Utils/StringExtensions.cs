using System.Linq;
using System.Collections.Generic;
using System.Globalization;
using System.Text.RegularExpressions;
using ZoomController.Utils;

namespace ZoomController
{
    public static class StringExtensions
    {
        private static readonly TextInfo TextInfo = new CultureInfo("en-US", false).TextInfo;

        /// <summary>
        /// Based on : https://www.dotnetperls.com/uppercase-first-letter
        /// </summary>
        public static string UppercaseFirst(this string value)
        {
            if (string.IsNullOrEmpty(value)) return value;

            char[] a = value.ToCharArray();
            a[0] = char.ToUpper(a[0]);
            return new string(a);
        }

        /// <summary>
        /// Based on: https://www.dotnetperls.com/uppercase-first-letter
        /// </summary>
        public static string UppercaseWords(this string value)
        {
            if (string.IsNullOrEmpty(value)) return value;

            char[] array = value.ToCharArray();
            // Handle the first letter in the string.
            if (array.Length >= 1)
            {
                if (char.IsLower(array[0]))
                {
                    array[0] = char.ToUpper(array[0]);
                }
            }
            // Scan through the letters, checking for spaces.
            // ... Uppercase the lowercase letters following spaces.
            for (int i = 1; i < array.Length; i++)
            {
                if (array[i - 1] == ' ')
                {
                    if (char.IsLower(array[i]))
                    {
                        array[i] = char.ToUpper(array[i]);
                    }
                }
            }
            return new string(array);
        }

        public static string ToTitleCase(this string value)
        {
            return TextInfo.ToTitleCase(value);
        }

        private static readonly Regex ReStripHTML = RegexCache.Get(@"<.*?>", RegexOptions.Compiled);
        /// <summary>
        /// Naïve removal of HTML tags from a string.  Based on:
        /// https://stackoverflow.com/a/18154046.
        /// </summary>
        /// <param name="s"></param>
        /// <returns></returns>
        public static string StripHTML(this string s)
        {
            return ReStripHTML.Replace(s, string.Empty);
        }

        /// <summary>
        /// Splits up a given sentence into words and returns them.  Adapted from: https://stackoverflow.com/a/16734675.
        /// </summary>
        /// <param name="text">The sentence to parse.</param>
        /// <returns>A string[] of words extracted from the given sentence.</returns>
        public static string[] GetWordsInSentence(this string text)
        {
            return text.Split().Select(x => x.Trim(text.Where(char.IsPunctuation).Distinct().ToArray())).ToArray();
        }
        

        /// <summary>
        /// Strips leading/trailing from each line and removes blank lines from multi-line strings. Normalizes line delimiter to a carriage return.
        /// </summary>
        public static string StripBlankLinesAndTrimSpace(this string s)
        {
            var lines = s.Split(ZCUtils.CRLFDelim);
            var ret = new List<string>();
            foreach (var line in lines)
            {
                var temp = line.Trim();
                if (temp.Length > 0)
                {
                    ret.Add(temp);
                }
            }
            return string.Join("\n", ret);
        }
    }
}
