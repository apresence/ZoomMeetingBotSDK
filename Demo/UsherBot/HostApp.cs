namespace ZoomMeetingBotSDK
{
    using System;
    using System.Collections.Generic;
    using System.IO;
    using System.Linq;
    using System.Text;
    using System.Threading;
    using static Utils;

    public class HostApp : CHostApp
    {
        public class HostAppSettings
        {
            public HostAppSettings()
            {
                DebugLoggingEnabled = false;
            }

            public bool DebugLoggingEnabled { get; set; }
        }

        private static readonly object SettingsLock = new object();
        private static readonly object LogLock = new object();
        private static readonly object IdleTimerLock = new object();

        private static readonly char[] DotSep = { '.' };

        private static Dictionary<string, dynamic> cfgDic = new Dictionary<string, dynamic>();

        private static System.Threading.Timer idleTimer = null;
        private static DateTime lastSettingsModDT = DateTime.MinValue;

        private static volatile bool shouldExit = false;
        private static volatile int idleTimerTicks = 0;

        private static string jsonSettings = null;

        private static HostAppSettings hostAppSettings = null;

        public override event EventHandler SettingsChanged;

        public void Init()
        {
            LoadSettings();
        }

        public void Start()
        {
            idleTimer = new System.Threading.Timer(IdleTimerHandler, null, 0, 5000);
        }

        public void Stop()
        {
            shouldExit = true;
        }

        public void LoadSettings()
        {
            lock (SettingsLock)
            {
                string sPath = @"settings.json";

                if (!File.Exists(sPath))
                {
                    SaveSettings();
                    return;
                }

                DateTime lastModDT = File.GetLastWriteTimeUtc(sPath);

                // Don't load/reload unless changed
                if (lastModDT == lastSettingsModDT)
                {
                    return;
                }

                lastSettingsModDT = lastModDT;

                Log(LogType.INF, "(Re-)loading settings.json");

                jsonSettings = File.ReadAllText(@"settings.json");
                hostAppSettings = DeserializeJson<HostAppSettings>(jsonSettings);
            }

            SettingsChanged?.Invoke(this, null);
        }

        public static void SaveSettings()
        {
            /*
            hostApp.Log(LogType.INF, "Saving settings.json");
            var options = new JsonSerializerOptions
            {
                WriteIndented = true,
            };
            File.WriteAllText(@"settings.json", JsonSerializer.Serialize(cfg, options));
            */
            throw new NotImplementedException();
        }

        public override string GetSettingsAsJSON()
        {
            lock (SettingsLock)
            {
                return jsonSettings;
            }
        }

        /*
        public override dynamic GetSetting(string key, dynamic default_value = null)
        {
            lock (SettingsLock)
            {
                cfgDic.TryGetValue(key, out dynamic obj);

                if (obj is null)
                {
                    return default_value;
                }

                /-*
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

                    return ret;
                }
                *-/

                return obj;
            }
        }

        public override Dictionary<string, dynamic> GetSettingsDic()
        {
            lock (SettingsLock)
            {
                return cfgDic.ToDictionary(kvp => kvp.Key, kvp => kvp.Value);  //kvp => kvp.Value.Clone());
            }
        }
        */

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
                        Console.WriteLine(string.Format("ERR Failed to write log file; Trying again in 1s (Attempt #{0}); Exception: {1}", attempt, repr(ex.ToString())));
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

            if ((nLogType != LogType.DBG) || hostAppSettings.DebugLoggingEnabled)
            {
                lock (LogLock)
                {
                    Console.WriteLine(s);
                }
            }

            Log(s);
        }

        private void IdleTimerHandler(object o)
        {
            if (shouldExit)
            {
                return;
            }

            idleTimerTicks += 1;

            if (!Monitor.TryEnter(IdleTimerLock))
            {
                Log(LogType.WRN, "HostApp.IdleTimerHandler {0:X4} - Busy; Will try again later", idleTimerTicks);
                return;
            }

            try
            {
                LoadSettings();

                // Zoom is really bad about moving/resizing it's windows, so keep it in check
                Controller.LayoutWindows();
            }
            catch (Exception ex)
            {
                Log(LogType.ERR, "HostApp.IdleTimerHandler {0:X4} - Unhandled Exception: {1}", idleTimerTicks, ex.ToString());
            }
            finally
            {
                Monitor.Exit(IdleTimerLock);
            }
        }
    }
}
