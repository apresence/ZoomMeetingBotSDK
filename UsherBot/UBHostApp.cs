namespace ZoomMeetngBotSDK
{
    using System;
    using System.IO;
    using System.Threading;
    using global::ZoomMeetngBotSDK.Interop.HostApp;

    public class UBHostApp : IHostApp
    {
        private static readonly object LogLock = new object();
        private static readonly char[] DotSep = { '.' };

        dynamic IHostApp.GetSetting(string key)
        {
            if (key.StartsWith("broadcast."))
            {
                string value = null;

                try
                {
                    return Global.cfg.BroadcastCommands[key.Split(DotSep, 2)[1]];
                }
                catch
                {
                    // pass
                }

                return value;
            }

            switch (key)
            {
                case "bot.name":
                    return Global.cfg.MyParticipantName;
                case "bot.gender":
                    return Global.cfg.BotGender;
            }

            ((IHostApp)this).Log(LogType.WRN, $"HostApp.GetSetting: Setting with key {Global.repr(key)} does not exist");

            return null;
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
