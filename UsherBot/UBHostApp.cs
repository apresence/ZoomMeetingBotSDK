namespace ZoomMeetingBotSDK
{
    using System;
    using System.Collections.Generic;
    using System.IO;
    using System.Text;
    using System.Threading;
    using global::ZoomMeetingBotSDK.Interop.HostApp;
    using global::ZoomMeetingBotSDK.Utils;

    public class UBHostApp : CHostApp
    {
        private static readonly object LogLock = new object();
        private static readonly char[] DotSep = { '.' };

        public override dynamic GetSetting(string key, dynamic default_value)
        {
            Global.cfgDic.TryGetValue(key, out dynamic obj);

            if (obj is null)
            {
                return default_value;
            }

            if (obj is object[])
            {
                var ret = new List<string>();
                foreach (var item in obj)
                {
                    ret.Add(Convert.ToString(item));
                }

                return ret;
            }

            if (obj is Dictionary<string, object>)
            {
                var ret = new Dictionary<string, string>();
                foreach (var kvp in obj)
                {
                    ret.Add(kvp.Key, Convert.ToString(kvp.Value));
                }

                ZMBUtils.ExpandDictionaryPipes(ret);

                return ret;
            }

            return obj;
        }

        public override void Log(string sMessage)
        {
            string n = "UsherBot.Log";

            lock (LogLock)
            {
                StreamWriter sw = null;
                for (int attempt = 0; attempt < 3; attempt++)
                {
                    try
                    {
                        sw = File.Exists(n) ? File.AppendText(n) : File.CreateText(n);
                    }
                    catch (IOException ex)
                    {
                        Console.WriteLine(string.Format("ERR Failed to write log file; Trying again in 1s (Attempt #{0}); Exception: {1}", attempt, ZMBUtils.repr(ex.ToString())));
                        Thread.Sleep(1000);

                        continue;
                    }

                    break;
                }

                if (sw == null)
                {
                    Console.WriteLine(string.Format("ERR Max attempts trying to write to log file; Giving up"));
                    return;
                }

                sw.WriteLine(sMessage);
                sw.Close();
            }
        }

        public override void Log(LogType nLogType, string sMessage)
        {
            var s = new StringBuilder(DateTime.UtcNow.ToString("yyyy-MM-dd HH:mm:ss.fff"))
                .Append(' ')
                .Append(nLogType.ToString())
                .Append(' ')
                .Append(sMessage)
                .ToString();

            if ((nLogType != LogType.DBG) || Global.cfg.DebugLoggingEnabled)
            {
                lock (LogLock)
                {
                    Console.WriteLine(s);
                }
            }

            Log(s);
        }
    }
}
