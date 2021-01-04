namespace ZoomMeetingBotSDK
{
    using System;
    using System.Collections.Generic;
    using System.IO;
    using System.Threading;
    using global::ZoomMeetingBotSDK.Interop.HostApp;
    using global::ZoomMeetingBotSDK.Utils;

    public class UBHostApp : IHostApp
    {
        private static readonly object LogLock = new object();
        private static readonly char[] DotSep = { '.' };

        dynamic IHostApp.GetSetting(string key)
        {
            Global.cfgDic.TryGetValue(key, out dynamic obj);

            if (obj is null)
            {
                return obj;
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

        void IHostApp.Log(Interop.HostApp.LogType nLogType, string sMessage, params object[] values)
        {
            string s = string.Format(
                "{0} {1} {2}",
                DateTime.UtcNow.ToString("yyyy-MM-dd HH:mm:ss.fff"),
                nLogType.ToString(),
                string.Format(sMessage, values));

            string n = "UsherBot.Log";

            lock (LogLock)
            {
                if ((nLogType != LogType.DBG) || Global.cfg.DebugLoggingEnabled)
                {
                    Console.WriteLine(s);
                }

                StreamWriter sw = null;
                for (int attempt = 0; attempt < 3; attempt++)
                {
                    try
                    {
                        sw = File.Exists(n) ? File.AppendText(n) : File.CreateText(n);
                    }
                    catch (IOException ex)
                    {
                        Console.WriteLine(string.Format("ERR Failed to write log file; Trying again in 1s (Attempt #{0}); Exception: {1}", attempt, Global.repr(ex.ToString())));
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

                sw.WriteLine(s);
                sw.Close();
            }
        }
    }
}
