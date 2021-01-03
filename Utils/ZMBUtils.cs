using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;
using System.Web.Script.Serialization;

namespace ZoomMeetngBotSDK.Utils
{
    /// <summary>
    /// We make this class abstract so that some members can be overriden with more tailored implementations.
    /// </summary>
    public abstract class ZMBUtils
    {
        public static readonly char[] CRLFDelim = new char[] { '\r', '\n' };

        public class JsonDictionary : Dictionary<string, dynamic>
        {
        }

        private static JavaScriptSerializer serializer = null;

        public static JsonDictionary JsonToDict(string json)
        {
            if (serializer == null)
            {
                serializer = new JavaScriptSerializer();
            }

            lock (serializer)
            {
                return (JsonDictionary)serializer.DeserializeObject(json);
            }
        }

        /// <summary>
        /// Reference implementation for returning a string representation of an object, similar to Python's native repr() function.
        /// </summary>
        public static string repr(object o)
        {
            if (o is null)
            {
                return "(null)";
            }

            var ret = new StringBuilder();
            var val = o.ToString();

            foreach (char c in val)
            {
                var i = (int)c;
                if (i <= 32)
                {
                    ret.Append('^');
                    ret.Append((char)(i + 64));
                }
                else if ("^\"\\".Contains(c))
                {
                    ret.Append('\\');
                    ret.Append(c);
                }
                else
                {
                    ret.Append(c);
                }
            }

            return ret.ToString();
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

        /*
        private static readonly Random _GetRandomIndex_rand = new Random();
        /// <summary>
        /// Returns a random value from the given Dictionary object.
        /// </summary>
        public static TValue GetRandomDictionaryValue<TKey, TValue>(IDictionary<TKey, TValue> dic)
        {
            return dic.ElementAt(_GetRandomIndex_rand.Next(0, dic.Count - 1)).Value;
        }

        public static string GetRandomStringFromArray(string[] ary)
        {
            return ary.ElementAt(_GetRandomIndex_rand.Next(0, ary.Length - 1));
        }
        */
    }
}
