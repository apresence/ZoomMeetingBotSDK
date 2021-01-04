namespace ZoomMeetingBotSDK
{
    using System;
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
        void Log(string sMessage);

        void Log(LogType nLogType, string sMessage);

        void Log(LogType nLogType, string sMessage, params object[] values);
        dynamic GetSetting(string key, dynamic default_value = null);
    }

    public class CHostApp : IHostApp
    {
        // Provide a lock across all instances of this object.
        private static readonly object logLock = new object();

        /// <summary>
        /// Reference implementation for retrieving settings.  Returns null for all keys.
        /// </summary>
        public virtual dynamic GetSetting(string key, dynamic default_value)
        {
            return default_value;
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
