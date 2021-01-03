using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ZoomController.Interop.HostApp
{
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

        void Log(LogType nLogType, string sMessage, params object[] values);
        dynamic GetSetting(string key);
    }

    public abstract class HostApp : IHostApp
    {
        // Provide a lock across all instances of this object.
        private static readonly object logLock = new object();

        /// <summary>
        /// Reference implementation for retrieving settings.  Returns null for all keys.
        /// </summary>
        dynamic IHostApp.GetSetting(string key)
        {
            return null;
        }

        /// <summary>
        /// Reference implementation for logging function.  Logs to the console.
        /// </summary>
        void IHostApp.Log(LogType nLogType, string sMessage, params object[] values)
        {
            string s = string.Format(
                    "{0} {1} {2}",
                    DateTime.UtcNow.ToString("yyyy-MM-dd HH:mm:ss.fff"),
                    nLogType.ToString(),
                    string.Format(sMessage, values));

            lock (logLock)
            {
                Console.WriteLine(s);
            }
        }
    }
}
