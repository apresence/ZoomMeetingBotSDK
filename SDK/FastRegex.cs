namespace ZoomMeetingBotSDK
{
    using System;
    using System.Collections.Generic;
    using System.Text.RegularExpressions;

    /// <summary>
    /// These classes are intended to make working with large numbers of regular expressions efficient and convenient.  To do so normally, one would
    /// have to instantiate a static Regex object for each regular expression which requires additional lines of code and is messy to track,
    /// especially as regular expressions fall in and out of use.
    /// 
    /// HOW IT WORKS
    /// When a pattern and option pair are encountered that are not in the cache, they are wrapped in a new Regex object with RegexOptions.Compiled
    /// set and added to the cache, then the new Regex is returned.  Otherwise, the cached RegEx is returned.  The size of the Regex MSIL object
    /// cache is also increased as needed such that it always has at least CACHE_INCREMENT_SIZE elements free.
    /// </summary>
    public static class RegexCache
    {
        static Dictionary<string, Regex> cache = new Dictionary<string, Regex>();

        // Size to increase global Regex cache when it is full
        static int CACHE_INCREMENT_SIZE = 15;

        // Timeout for regex calculation.  This is to prevent a DoS for malformed regular expressions
        static TimeSpan MATCH_TIMEOUT = new TimeSpan(0, 0, 5);

        public static Regex Get(string pattern, RegexOptions options)
        {
            Regex ret = null;

            // Always compile RE to MSIL for efficiency
            options |= RegexOptions.Compiled;

            // Simple compound key with pattern & options
            var key = pattern + ":" + options.ToString();

            lock (cache)
            {
                if (!cache.TryGetValue(key, out ret))
                {
                    ret = new Regex(pattern, options, MATCH_TIMEOUT);

                    int targetCacheSize = CACHE_INCREMENT_SIZE * ((int)(cache.Count / CACHE_INCREMENT_SIZE) + 2);
                    if (Regex.CacheSize < targetCacheSize)
                    {
                        Regex.CacheSize = targetCacheSize;
                    }

                    cache[key] = ret;
                }
            }
            return ret;
        }

        public static Regex Get(string pattern)
        {
            return Get(pattern, RegexOptions.None);
        }
    }

    public class FastRegex
    {
        private Regex re;

        public FastRegex(string pattern, RegexOptions options)
        {
            re = RegexCache.Get(pattern, options);
        }

        public FastRegex(string pattern)
        {
            re = RegexCache.Get(pattern, RegexOptions.None);
        }

        public static bool IsMatch(string input, string pattern, RegexOptions options)
        {
            return RegexCache.Get(pattern, options).IsMatch(input);
        }

        public static bool IsMatch(string input, string pattern)
        {
            return RegexCache.Get(pattern, RegexOptions.None).IsMatch(input);
        }

        public static string Replace(string input, string pattern, string replacement, RegexOptions options)
        {
            return RegexCache.Get(pattern, options).Replace(input, replacement);
        }

        public static string Replace(string input, string pattern, string replacement)
        {
            return RegexCache.Get(pattern, RegexOptions.None).Replace(input, replacement);
        }
    }
}
