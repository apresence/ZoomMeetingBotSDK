using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;
using System.Threading;
using ZoomMeetingBotSDK.Interop.HostApp;
using ZoomMeetingBotSDK.Utils;
using static ZoomMeetingBotSDK.Utils.ZMBUtils;

namespace ZoomMeetingBotSDK
{
    public class Global
    {
        static public IHostApp hostApp = null;

        [Flags]
        public enum BotAutomationFlag
        {
            None = 0b0000000000,
            RenameMyself = 0b0000000001,
            ReclaimHost = 0b0000000010,
            ProcessParticipants = 0b0000000100,
            ProcessChat = 0b0000001000,
            CoHostKnown = 0b0000010000,
            AdmitKnown = 0b0000100000,
            AdmitOthers = 0b0001000000,
            Converse = 0b0010000000,
            Speak = 0b0100000000,
            UnmuteMyself = 0b1000000000,
            All = 0b1111111111
        }

        public class EmailCommandArgs
        {
            /// <summary>
            /// Example for arguments.
            /// </summary>
            public string ArgsExample { get; set; }

            /// <summary>
            /// Subject for the email.
            /// </summary>
            public string Subject { get; set; }

            /// <summary>
            /// Body for the email.
            /// </summary>
            public string Body { get; set; }
        }

        public class ConfigurationSettings
        {
            /// <summary>
            /// Absolute path to web browser executable.
            /// </summary>
            public string BrowserExecutable { get; set; }

            /// <summary>
            /// Optional command line arguments to pass to web browser.
            /// </summary>
            public string BrowserArguments { get; set; }

            /// <summary>
            /// If true, enables additional debug logging. If false, suppresses that logging.
            /// </summary>
            public bool DebugLoggingEnabled { get; set; }

            /// <summary>
            /// If true, the BOT will stop all automated operations.  Primarily useful for debugging purposes.
            /// </summary>
            public bool IsPaused { get; set; }

            /// <summary>
            /// If true, the BOT will prompt for input just after startup and before beginning automated actions.  Primarily useful for debugging purposes.
            /// </summary>
            public bool PromptOnStartup { get; set; }

            /// <summary>
            /// If true, the program will wait for a debugger to attach right after startup.
            /// </summary>
            public bool WaitForDebuggerAttach { get; set; }

            /// <summary>
            /// Number of seconds to wait before admitting an unknown participant to the meeting.
            /// </summary>
            public int UnknownParticipantWaitSecs { get; set; }

            /// <summary>
            /// Number of seconds to pause between adding unknown participants.  This helps to mitigate against Zoom Bombers which tend to flood the meeting all at once.
            /// </summary>
            public int UnknownParticipantThrottleSecs { get; set; }

            /// <summary>
            /// Name to use when joining the Zoom meeting.  If the default name does not match, a rename is done after joining the meeting.
            /// </summary>
            public string MyParticipantName { get; set; }

            /// <summary>
            /// A set of flags that controls which Bot automation is enabled and disabled.  See BotAutomationFlag enum for further details.
            /// </summary>
            public BotAutomationFlag BotAutomationFlags { get; set; }

            /// <summary>
            /// Sets the gender of the Bot.
            /// </summary>
            public Interop.ChatBot.Gender BotGender { get; set; }

            /// <summary>
            /// ID of the meeting to join.
            /// </summary>
            public string MeetingID { get; set; }

            /// <summary>
            /// A list of commands which when invoked will send a predefined message to Everyone in the chat.
            /// </summary>
            public Dictionary<string, string> BroadcastCommands { get; set; }

            /// <summary>
            /// Number of seconds to delay before allowing the same broadcast message to be sent again.
            ///   <0 : Infinite delay (Only send broadcast message once)
            ///    0 : No delay
            ///   >0 : Delay in seconds
            /// </summary>
            public int BroadcastCommandGuardTimeSecs { get; set; }

            /// <summary>
            /// A list of commands that will send an email.  The format is /command to-address whatever-you-want-that-is-put-in-{0}-in-the-message
            /// </summary>
            public Dictionary<string, EmailCommandArgs> EmailCommands { get; set; }

            /// <summary>
            /// A list of pre-defined one-time greetings based on received private chat messages.  Use pipes to allow more than one query to produce the same response.  TBD: Move into RemedialBot
            /// </summary>
            public Dictionary<string, string> OneTimeHiSequences { get; set; }

            /// <summary>
            /// A list of pre-defined small talk responses based on received private chat messages.  Use pipes to allow more than one query to produce the same response.  TBD: Move into RemedialBot
            /// </summary>
            public Dictionary<string, string> SmallTalkSequences { get; set; }

            /// <summary>
            /// Random bot responses.  TBD: Move into RemedialBot
            /// </summary>
            public string[] RandomTalk { get; set; }

            /// <summary>
            /// If set, this message is sent to participants in the waiting room every WaitingRoomAnnouncementDelaySecs seconds.
            /// </summary>
            public string WaitingRoomAnnouncementMessage { get; set; }

            /// <summary>
            /// Controls the sending frequency for WaitingRoomAnnouncementMessage.
            /// </summary>
            public int WaitingRoomAnnouncementDelaySecs { get; set; }

            /// <summary>
            /// Number of ms to delay after performing an action which requires an update to the Zoom UI.
            /// </summary>
            public int UIActionDelayMilliseconds { get; set; }

            /// <summary>
            /// Number of ms to delay after moving mouse to target location before sending click event.
            /// </summary>
            public int ClickDelayMilliseconds { get; set; }

            /// <summary>
            /// Number of ms to delay after sending keyboard input to the remote app.
            /// </summary>
            public int KeyboardInputDelayMilliseconds { get; set; }


            /// <summary>
            /// Disables using clipboard to paste text into target apps; Falls back on sending individual keystrokes instead.
            /// </summary>
            public bool DisableClipboardPasteText { get; set; }

            /// <summary>
            /// Number of times to retry parsing participant list until count matches what is in the window title.
            /// </summary>
            public int ParticipantCountMismatchRetries { get; set; }

            /// <summary>
            /// Disable paging up/down in the Participants window. If screen resolution is sufficiently big (height), it may not be needed.
            /// </summary>
            public bool DisableParticipantPaging { get; set; }

            /// <summary>
            /// Duration of time over which to do mouse moves when simulating input. This helps the target app "see" the mouse movement more reliably. <= 0 will move the mouse instantly.
            /// </summary>
            public double MouseMovementRate { get; set; }

            /// <summary>
            /// Seconds to wait between opening the meeting options menu to get status from the UI.
            /// Special values:
            ///   -1  Only poll when meeting is first started, then poll on demand (when one of the options needs to be changed)
            ///    0  Poll as fast as possible
            ///   >0  Delay between polls
            /// </summary>
            public int UpdateMeetingOptionsDelaySecs { get; set; }

            /// <summary>
            /// Sets the Windows Speech TTS Voice to use based on it's name.
            /// </summary>
            public string TTSVoice { get; set; }

            /// <summary>
            /// Configures which display should be used to run Zoom, ZoomMeetingBotSDK, UsherBot, etc.  The default is whichever display is set as the "main" screen
            /// </summary>
            public string Screen { get; set; }

            /// <summary>
            /// Enables the function WalkRawElementsToString() which walks the entire AutomationElement tree and returns a string.  The default is True.  This is an
            /// extremely expensive operation, often taking a minute or more, so disabling it on live meetings may be a good idea.
            /// </summary>
            public bool EnableWalkRawElementsToString { get; set; }

            /// <summary>
            /// Full path to Zoom executable
            /// </summary>
            public string ZoomExecutable { get; set; }

            /// <summary>
            /// User name for Zoom account
            /// </summary>
            public string ZoomUsername { get; set; }

            /// <summary>
            /// Password for Zoom account (encrypted - can only be decrypted by the current user on the current machine)
            /// </summary>
            public string ZoomPassword { get; set; }

            public ConfigurationSettings()
            {
                BrowserExecutable = @"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe";
                BrowserArguments = "https://zoom.us/signin";
                DebugLoggingEnabled = false;
                IsPaused = false;
                PromptOnStartup = false;
                WaitForDebuggerAttach = false;
                UnknownParticipantThrottleSecs = 15;
                UnknownParticipantWaitSecs = 30;
                MyParticipantName = "ZoomBot";
                BotAutomationFlags = BotAutomationFlag.All;
                BotGender = Interop.ChatBot.Gender.Female;
                MeetingID = null;
                BroadcastCommands = new Dictionary<string, string>();
                BroadcastCommandGuardTimeSecs = 300;
                EmailCommands = new Dictionary<string, EmailCommandArgs>();
                OneTimeHiSequences = new Dictionary<string, string>();
                SmallTalkSequences = new Dictionary<string, string>();
                WaitingRoomAnnouncementMessage = null;
                WaitingRoomAnnouncementDelaySecs = 60;
                DisableClipboardPasteText = false;
                ParticipantCountMismatchRetries = 3;
                UIActionDelayMilliseconds = 250;
                ClickDelayMilliseconds = UIActionDelayMilliseconds;
                KeyboardInputDelayMilliseconds = UIActionDelayMilliseconds;
                DisableParticipantPaging = false;
                MouseMovementRate = 100;
                UpdateMeetingOptionsDelaySecs = -1;
                TTSVoice = null;
                Screen = null;
                ZoomExecutable = @"%AppData%\Zoom\bin\Zoom.exe";
                ZoomUsername = null;
                ZoomPassword = null;
            }
        }

        public static ConfigurationSettings cfg = new ConfigurationSettings();
        public static Dictionary<string, dynamic> cfgDic = new Dictionary<string, dynamic>();

        public static T DeserializeJson<T>(string json)
        {
            return JsonSerializer.Deserialize<T>(json);
        }

        // Similar to python repr() function
        public static string repr(object o)
        {
            try
            {
                //return JsonSerializer.Serialize(o);
                return JSONSpecialDoubleHandler.HandleSpecialDoublesAsStrings.Serialize(o);
            }
            catch (Exception ex)
            {
                throw new FormatException(string.Format("Could not serialize object {0}", o.GetType().Name), ex);
            }
        }

        private static DateTime dtLastSettingsMod = DateTime.MinValue;

        public static void LoadSettings()
        {
            string sPath = @"settings.json";

            if (!File.Exists(sPath))
            {
                SaveSettings();
                return;
            }

            DateTime dtLastMod = File.GetLastWriteTimeUtc(sPath);

            // Don't load/reload unless changed
            if (dtLastMod == dtLastSettingsMod)
            {
                return;
            }

            dtLastSettingsMod = dtLastMod;

            hostApp.Log(LogType.INF, "(Re-)loading settings.json");

            var json = File.ReadAllText(@"settings.json");
            cfg = JsonSerializer.Deserialize<ConfigurationSettings>(json);
            cfgDic = JsonToDict(json);

            ZMBUtils.ExpandDictionaryPipes(cfg.BroadcastCommands);
            ZMBUtils.ExpandDictionaryPipes(cfg.OneTimeHiSequences);
            ZMBUtils.ExpandDictionaryPipes(cfg.SmallTalkSequences);

            UsherBot.SettingsUpdated();
        }

        public static void SaveSettings()
        {
            hostApp.Log(LogType.INF, "Saving settings.json");
            var options = new JsonSerializerOptions
            {
                WriteIndented = true,
            };
            File.WriteAllText(@"settings.json", JsonSerializer.Serialize(cfg, options));
        }

        public static void WaitDebuggerAttach()
        {
            Console.WriteLine("Waiting for debugger to attach");
            while (!Debugger.IsAttached)
            {
                Thread.Sleep(100);
            }
            Console.WriteLine("Debugger attached");
        }

        /// <summary>
        /// Class providing Enum of strings. Uses description attribute and reflection. Caches values for efficiency.
        /// Based on this article: https://stackoverflow.com/questions/4367723/get-enum-from-description-attribute.
        /// </summary>
        public static class EnumEx
        {
            private static readonly Dictionary<(Type, dynamic), string> DescCache = new Dictionary<(Type, dynamic), string>();

            public static string GetDescriptionFromValue<T>(T value)
            {
                var key = (typeof(T), value);

                if (!DescCache.TryGetValue(key, out string ret))
                {
                    ret = (Attribute.GetCustomAttribute(value.GetType().GetField(value.ToString()), typeof(DescriptionAttribute)) is DescriptionAttribute attribute) ? attribute.Description : value.ToString();
                    DescCache[key] = ret;
                }

                return ret;
            }

            private static readonly Dictionary<(Type, string), dynamic> ValueCache = new Dictionary<(Type, string), dynamic>();

            public static T GetValueFromDescription<T>(string description)
            {
                var type = typeof(T);
                var key = (typeof(T), description);

                if (!ValueCache.TryGetValue(key, out dynamic ret))
                {
                    if (!type.IsEnum)
                    {
                        throw new InvalidOperationException();
                    }

                    bool bFound = false;
                    foreach (var field in type.GetFields())
                    {
                        if (Attribute.GetCustomAttribute(field, typeof(DescriptionAttribute)) is DescriptionAttribute attribute)
                        {
                            if (attribute.Description == description)
                            {
                                bFound = true;
                                ret = field.GetValue(null);
                                break;
                            }
                        }
                        else
                        {
                            if (field.Name == description)
                            {
                                bFound = true;
                                ret = field.GetValue(null);
                                break;
                            }
                        }
                    }
                    if (!bFound)
                    {
                        throw new KeyNotFoundException(string.Format("Description/Name {0} not found in enum {1}", Global.repr(description), Global.repr(type.ToString())));
                    }
                }

                return (T)ret;
            }
        }
    }
}
