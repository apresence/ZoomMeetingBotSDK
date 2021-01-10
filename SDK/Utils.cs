namespace ZoomMeetingBotSDK
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Reflection;
    using System.Security.Cryptography;
    using System.Text;
    using System.Text.RegularExpressions;
    using System.Web.Script.Serialization;
    using Newtonsoft.Json;

    /// <summary>
    /// Common utilities useful in implementing bots.
    /// </summary>
    public static class Utils
    {
        public static readonly char[] CRLFDelim = new char[] { '\r', '\n' };

        /// <summary>
        /// Shared global serializer.  This is a tradeoff between instantiation and lock contention time.
        /// Since I don't plan on using it often, I've chosen the shared/locking method.
        /// </summary>
        private static JavaScriptSerializer serializer = new JavaScriptSerializer();

        public static T DeserializeJson<T>(string json)
        {
            //System.Text.Json version
            //return JsonSerializer.Deserialize<T>(json);

            return JsonConvert.DeserializeObject<T>(json);
        }

        /// <summary>
        /// Reference implementation for returning a string representation of an object, similar to Python's native repr() function.
        /// </summary>
        public static string repr(object obj)
        {
            if (obj == null)
            {
                return "(null)";
            }

            if (!(obj is Enum))
            {
                try
                {
                    return (new JavaScriptSerializer()).Serialize(obj);
                }
                catch
                {
                }
            }

            try
            {
                return obj.ToString();
            }
            catch
            {
            }

            return "(repr: Failed to convert object " + obj.GetType().ToString() + " to string)";
        }

        private static readonly HashSet<string> SkipPropertyNames = new HashSet<string> { "ControlType", "ProcessId", "Orientation" };

        public static string GetObjStrs(object o)
        {
            List<string> l = new List<string>();
            foreach (var prop in o.GetType().GetProperties())
            {
                if (SkipPropertyNames.Contains(prop.Name))
                {
                    continue;
                }

                var val = prop.GetValue(o, null);
                if (val == null)
                {
                    continue;
                }

                if ((val is string s) && (s.Length == 0))
                {
                    continue;
                }

                if ((val is bool b) && (b == false))
                {
                    continue;
                }

                if ((val is int i) && (i == 0))
                {
                    continue;
                }

                l.Add(string.Format("{0}:{1}", repr(prop.Name), repr(val)));
            }
            return string.Join(",", l);
        }

        public static string GetObjHash(object o)
        {
            MD5 hash = MD5.Create();
            StringBuilder sb = new StringBuilder();
            byte[] data = hash.ComputeHash(Encoding.UTF8.GetBytes(GetObjStrs(o)));

            for (int i = 0; i < data.Length; i++)
            {
                sb.Append(data[i].ToString("x2"));
            }

            return sb.ToString();
        }

        public static string RepeatString(string s, int count)
        {
            return new StringBuilder().Insert(0, s, count).ToString();
        }

        public static string GetFirstRegExGroupMatch(Regex re, string text, string default_value = null)
        {
            try
            {
                MatchCollection matches = re.Matches(text);
                GroupCollection groups = matches[0].Groups;
                return groups[1].Value;
            }
            catch
            {
                return default_value;
            }
        }

        public static void ExpandDictionaryPipes(Dictionary<string, string> dic)
        {
            if (dic is null)
            {
                return;
            }
            
            string[] keys = new string[dic.Count];
            dic.Keys.CopyTo(keys, 0);

            foreach (string key in keys)
            {
                var a = key.Split('|');

                // If this key doesn't have any pipes, we can skip it
                if (a.Length == 1)
                {
                    continue;
                }

                var val = dic[key];
                dic.Remove(key);
                foreach (var subkey in a)
                {
                    // Skip blank keys
                    var cleanSubkey = subkey.Trim();
                    if (cleanSubkey.Length == 0)
                    {
                        continue;
                    }

                    dic.Add(subkey, val);
                }
            }
        }
    }
}
