namespace ZoomMeetingBotSDK
{
    using System;
    using System.Collections.Generic;
    using System.Diagnostics;
    using System.Text;

    public enum LogType
    {
        INF = 0,
        WRN,
        ERR,
        CRT,
        DBG,
    }

    public interface IHostApp
    {
        /// <summary>
        /// Fired when one or more settings have been changed.
        /// </summary>
        event EventHandler SettingsChanged;

        /// <summary>
        /// Log a single string.
        /// </summary>
        /// <param name="sMessage"></param>
        void Log(string sMessage);

        /// <summary>
        /// Log a string with a log type.
        /// </summary>
        void Log(LogType nLogType, string sMessage);

        /// <summary>
        /// Log a formatted string with a log type.
        /// </summary>
        void Log(LogType nLogType, string sMessage, params object[] values);

        /// <summary>
        /// Retrieve JSON string of all settings.
        /// </summary>
        string GetSettingsAsJSON();

        /// <summary>
        /// Retrieve a setting by key.
        /// </summary>
        dynamic GetSetting(string key, dynamic default_value = null);

        /// <summary>
        /// Returns a copy of all settings as a key/value pair dictionary.
        /// </summary>
        Dictionary<string, dynamic> GetSettingsDic();
    }

    public class CHostApp : IHostApp
    {
        // Provide a lock across all instances of this object.
        private static readonly object logLock = new object();

        public virtual event EventHandler SettingsChanged;

        /// <summary>
        /// Reference implementation for retrieving settings as JSON. Returns an empty string.
        /// </summary>
        public virtual string GetSettingsAsJSON()
        {
            return "";
        }

        /// <summary>
        /// Reference implementation for retrieving settings.  Returns null for all keys.
        /// </summary>
        public virtual dynamic GetSetting(string key, dynamic default_value)
        {
            return default_value;
        }

        public virtual Dictionary<string, dynamic> GetSettingsDic()
        {
            throw new NotImplementedException();
        }

        /// <summary>
        /// Reference implementation for logging function.  Logs to the console and debug output window.
        /// </summary>
        public virtual void Log(string sMessage)
        {
            lock (logLock)
            {
                Debug.WriteLine(sMessage);
                Console.WriteLine(sMessage);
            }
        }


        /// <summary>
        /// Prepends timestamp and log type to log message, then passes it to Log(sMessage).
        /// </summary>
        public virtual void Log(LogType nLogType, string sMessage)
        {
            Log(new StringBuilder(DateTime.UtcNow.ToString("yyyy-MM-dd HH:mm:ss.fff"))
                .Append(' ')
                .Append(nLogType.ToString())
                .Append(' ')
                .Append(sMessage)
                .ToString());
        }

        /// <summary>
        /// Prepends timestamp and log type to log message, performing string format expansion, then passing the result along to Log(nLogType, sMessage).
        /// </summary>
        public virtual void Log(LogType nLogType, string sMessage, params object[] values)
        {
            Log(nLogType, string.Format(sMessage, values));
        }
    }
}
