using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Globalization;

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
    }
}
