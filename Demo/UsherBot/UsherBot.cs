﻿// TBD:
//   - Move all try/catch logic to controller

namespace ZoomMeetingBotSDK
{
    using System;
    using System.Collections.Generic;
    using System.Diagnostics;
    using System.IO;
    using System.Linq;
    using System.Reflection;
    using System.Runtime.CompilerServices;
    using System.Text;
    using System.Text.RegularExpressions;
    using System.Threading;

    using ZoomMeetingBotSDK.ChatBot;
    using ZoomMeetingBotSDK.ControlBot;

    using static ZoomMeetingBotSDK.Utils;

    public class UsherBot : IControlBot
    {
#pragma warning disable SA1401 // Fields should be private
        public static volatile bool shouldExit = false;
#pragma warning restore SA1401 // Fields should be private
        private static volatile bool leaveMeeting = false;
        private static volatile bool leavingMeeting = false;
        private static volatile bool endForAll = true;
        private static string speaker = null;

        private static IHostApp hostApp;

        private enum UserLevel
        {
            Unknown,
            Known,
            CoHost,
            AutoCoHost,
            Admin,
        }

        /// <summary>
        /// This is a list of userIDs and names that we have kicked, banned, etc. for the meeting.  We use this to
        /// prevent re-admitting/co-hosting etc. bad actors.
        /// </summary>
        private static readonly Dictionary<uint, string> BadUsers = new Dictionary<uint, string>();

        private static readonly Dictionary<string, UserLevel> UserLevels = new Dictionary<string, UserLevel>();
        private static readonly object _lock_eh = new object();

        private static DateTime dtLastWaitingRoomAnnouncement = DateTime.MinValue;

        private static DateTime dtLastAdmission = DateTime.MinValue;
        private static DateTime dtLastGoodUserMod = DateTime.MinValue;

        /// <summary>
        /// Used to record the last time a broadcast message was sent in order to prevent a specific broadcast message from being requested & sent in rapid succession.
        /// </summary>
        private static Dictionary<string, DateTime> BroadcastSentTime = new Dictionary<string, DateTime>();

        private static Dictionary<uint, HashSet<string>> trackers = new Dictionary<uint, HashSet<string>>();

        /// <summary>
        /// Topic of the current meeting. Set with "/topic ..." command (available to admins only), sent to new participants as they join, and also retreived on-demand by "/topic".
        /// </summary>
        private static string Topic = null;

        [Flags]
        public enum BotAutomationFlag
        {
            None                 = 0b00000000000,
            SendTopicOnJoin      = 0b00000000001,
            RenameMyself         = 0b00000000010,
            ReclaimHost          = 0b00000000100,
            ProcessParticipants  = 0b00000001000,
            ProcessChat          = 0b00000010000,
            CoHostKnown          = 0b00000100000,
            AdmitKnown           = 0b00001000000,
            AdmitOthers          = 0b00010000000,
            Converse             = 0b00100000000,
            Speak                = 0b01000000000,
            UnmuteMyself         = 0b10000000000,
            All                  = 0b11111111111,
        }

        public class EmailCommandArgs
        {
            /// <summary>
            /// Example for arguments.
            /// </summary>
            public string ArgsExample;

            /// <summary>
            /// Subject for the email.
            /// </summary>
            public string Subject;

            /// <summary>
            /// Body for the email.
            /// </summary>
            public string Body;
        }

        public class BotConfigurationSettings
        {
            public BotConfigurationSettings()
            {
                DebugLoggingEnabled = false;
                IsPaused = false;
                UnknownParticipantThrottleSecs = 15;
                UnknownParticipantWaitSecs = 30;
                MyParticipantName = "UsherBot";
                BotAutomationFlags = BotAutomationFlag.All;
                MeetingID = null;
                BroadcastCommands = new Dictionary<string, string>();
                BroadcastCommandGuardTimeSecs = 300;
                EmailCommands = new Dictionary<string, EmailCommandArgs>();
                WaitingRoomAnnouncementMessage = null;
                WaitingRoomAnnouncementDelaySecs = 60;
                IncludeUserIdsInLists = false;
                BadNameRegex = null;
            }

            /// <summary>
            /// If true, enables additional debug logging. If false, suppresses that logging.
            /// </summary>
            public bool DebugLoggingEnabled { get; set; }

            /// <summary>
            /// If true, the BOT will stop all automated operations.  Primarily useful for debugging purposes.
            /// </summary>
            public bool IsPaused { get; set; }

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
            /// If set, this message is sent to participants in the waiting room every WaitingRoomAnnouncementDelaySecs seconds.
            /// </summary>
            public string WaitingRoomAnnouncementMessage { get; set; }

            /// <summary>
            /// Controls the sending frequency for WaitingRoomAnnouncementMessage.
            /// </summary>
            public int WaitingRoomAnnouncementDelaySecs { get; set; }

            /// <summary>
            /// Controls including user ids in lists or not.  If on, users are listed like "\"User Name\"#123456"; If off just "User Name".
            /// </summary>
            public bool IncludeUserIdsInLists { get; set; }

            /// <summary>
            /// A regular expression defining bad participant names. Any participant matching this pattern will not be admitted to the meeting.
            /// </summary>
            public string BadNameRegex { get; set; }
        }

        public static BotConfigurationSettings cfg = new BotConfigurationSettings();
        private static Regex badNameRegex = null;

        public static bool SetMode(string sName, bool bNewState)
        {
            if (sName == "citadel")
            {
                // In Citadel mode, we do not automatically admit unknown participants
                bool bCitadelMode = (cfg.BotAutomationFlags & BotAutomationFlag.AdmitOthers) == 0;
                if (bCitadelMode == bNewState)
                {
                    return false;
                }

                if (bNewState)
                {
                    cfg.BotAutomationFlags ^= BotAutomationFlag.AdmitOthers;
                }
                else
                {
                    cfg.BotAutomationFlags |= BotAutomationFlag.AdmitOthers;
                }

                hostApp.Log(LogType.INF, "Citadel mode {0}", bNewState ? "on" : "off");
                return true;
            }

            if (sName == "lockdown")
            {
                // In lockdown mode, lock the meeting and don't admit ayone
                var botLockdownFlags = BotAutomationFlag.AdmitOthers | BotAutomationFlag.AdmitKnown | BotAutomationFlag.CoHostKnown;
                bool bLockdownMode = (cfg.BotAutomationFlags & botLockdownFlags) == 0;
                if (bLockdownMode == bNewState)
                {
                    return false;
                }

                if (bNewState)
                {
                    cfg.BotAutomationFlags ^= botLockdownFlags;
                }
                else
                {
                    cfg.BotAutomationFlags |= botLockdownFlags;
                }

                hostApp.Log(LogType.INF, $"[UsherBot] {(bNewState ? "Locking" : "Unlocking")} meeting");
                _ = Controller.LockMeeting(bNewState);

                hostApp.Log(LogType.INF, $"Lockdown mode {(bNewState ? "on" : "off")}");
                return true;
            }

            if (sName == "debug")
            {
                if (cfg.DebugLoggingEnabled == bNewState)
                {
                    return false;
                }

                cfg.DebugLoggingEnabled = bNewState;
                hostApp.Log(LogType.INF, "Debug mode {0}", cfg.DebugLoggingEnabled ? "on" : "off");
                return true;
            }

            if (sName == "pause")
            {
                if (cfg.IsPaused == bNewState)
                {
                    return false;
                }

                cfg.IsPaused = bNewState;
                hostApp.Log(LogType.INF, "Pause mode {0}", cfg.IsPaused ? "on" : "off");
                return true;
            }

            if (sName == "passive")
            {
                var bPassive = cfg.BotAutomationFlags == BotAutomationFlag.None;
                if (bPassive == bNewState)
                {
                    return false;
                }

                cfg.BotAutomationFlags = bPassive ? BotAutomationFlag.None : BotAutomationFlag.All;
                hostApp.Log(LogType.INF, "Passive mode {0}", bPassive ? "on" : "off");
                return true;
            }

            throw new Exception(string.Format("Unknown mode: {0}", sName));
        }

        public static void ClearRemoteCommands()
        {
            string sPath = @"command_file.txt";
            if (!File.Exists(sPath))
            {
                return;
            }

            File.Delete(sPath);
        }

        public static string GetTopic(bool useDefault = true)
        {
            if (string.IsNullOrEmpty(Topic)) {
                return useDefault ? "The topic has not been set." : null;
            }

            // Prepend "Today's topic: " or "Tonight's topic: " only if it is not already present
            var dayName = GetTodayTonight();
            return Topic.ToLower().StartsWith(dayName) ? Topic : dayName.UppercaseFirst() + "'s topic: " + Topic;
        }

        public static bool SendTopic(Controller.Participant recipient, bool useDefault = true)
        {
            // Skip pure phone users since we can't send chat messages to them
            if (recipient.IsPurePhoneUser)
            {
                return false;
            }

            // Only send automated topic (useDefault = false) once per unique userId/name.  This prevents us from sending the topic to someone more than once if they leave/re-join the meeting
            if (!useDefault && (DicTopicSends.ContainsKey(recipient.UserId) || DicTopicSends.Values.Contains(recipient.Name)))
            {
                return false;
            }

            var topic = GetTopic(useDefault);

            if (topic == null)
            {
                return false;
            }

            var response = ChatBotConverse("_onetimehi_", recipient, "SimpleBot"); // Say "hi"
            response = response == null ? topic : response + " " + topic;

            DicTopicSends[recipient.UserId] = recipient.Name;

            return Controller.SendChatMessage(recipient, response);
        }

        /// <summary>
        /// Return normalized particpant name in an attempt to minimize variance:
        ///   1. Convert to lowercase
        ///   2. Remove one-letter last name abbreviations ("John D." -> "John D")
        ///   3. Replace runs of spaces with a single space ("John  D" -> "John D")
        ///   4. Remove known suffixes: (Usher), (DL), (Chair), (Speaker) ("John D (Usher)" -> "John D").
        /// </summary>
        private static string CleanUserName(string s)
        {
            if (s == null)
            {
                return null;
            }

            //return Regex.Replace(Regex.Replace(s.ToLower().Replace(".", string.Empty), @"\s+", " "), @"\s*\((?:Usher|DL|Chair|Speaker)\)\s*$", string.Empty, RegexOptions.IgnoreCase).Trim();
            return FastRegex.Replace(
                FastRegex.Replace(
                    FastRegex.Replace(
                        s.ToLower(),
                        @"(\s[a-z])\.",
                        "$1"),
                    @"\s+",
                    " "),
                @"\s*\((?:Usher|DL|Chair|Speaker)\)\s*$",
                string.Empty).Trim();
        }

        private static readonly HashSet<string> HsParticipantMessages = new HashSet<string>();

        private static string FirstParticipantGreeted = null;

        private static void DoParticipantActions()
        {
            if ((cfg.BotAutomationFlags & BotAutomationFlag.ProcessParticipants) == 0)
            {
                return;
            }

            // TBD: Could UpdateParticipants() here, but we should be good with updates provided by fired SDK events

            if (Controller.me != null)
            {
                // If I've got my own participant object, do any self-automation needed

                if (((cfg.BotAutomationFlags & BotAutomationFlag.ReclaimHost) != 0) && (!Controller.me.IsHost))
                {
                    // TBD: Throttle ReclaimHost attempts?
                    if (Controller.me.IsCoHost)
                    {
                        hostApp.Log(LogType.WRN, "BOT I'm Co-Host instead of Host; Trying to reclaim host");
                    }
                    else
                    {
                        hostApp.Log(LogType.WRN, "BOT I'm not Host or Co-Host; Trying to reclaim host");
                    }

                    if (Controller.ReclaimHost())
                    {
                        hostApp.Log(LogType.INF, "BOT Reclaim host successful");
                    }
                    else
                    {
                        hostApp.Log(LogType.WRN, "BOT Failed to reclaim host");
                    }
                }

                if (((cfg.BotAutomationFlags & BotAutomationFlag.RenameMyself) != 0) && (Controller.me.Name != cfg.MyParticipantName))
                {
                    // Rename myself.  Event handler will type in the name when the dialog pops up
                    hostApp.Log(LogType.INF, $"BOT Renaming myself from {repr(Controller.me.Name)} to {repr(cfg.MyParticipantName)}");
                    _ = Controller.RenameParticipant(Controller.me, cfg.MyParticipantName);
                }

                if (((cfg.BotAutomationFlags & BotAutomationFlag.UnmuteMyself) != 0) && Controller.me.IsAudioMuted)
                {
                    // Unmute myself
                    hostApp.Log(LogType.INF, "BOT Unmuting myself");
                    _ = Controller.UnmuteParticipant(Controller.me);
                }

            }

            // TBD: Could update meeting options here to see if everyone is muted, etc...

            int numWaiting = 0;
            int numAttending = 0;

            bool bAdmitOthers = (cfg.BotAutomationFlags & BotAutomationFlag.AdmitOthers) != 0;
            DateTime dtNow = DateTime.UtcNow;

            // Get a safe copy of participant list
            List<Controller.Participant> plist = null;
            lock (Controller.participants)
            {
                plist = Controller.participants.Values.ToList<Controller.Participant>();
            }

            foreach (Controller.Participant p in plist)
            {
                // Skip over my own participant record; We handled that earlier.  Also, skip over anyone not in the waiting room
                if (p.IsMe)
                {
                    continue;
                }

                switch (p.Status)
                {
                    case Controller.ParticipantStatus.Attending:
                        // Do attending actions

                        numAttending++;

                        _ = PromoteIfNeeded(p);

                        continue;
                    case Controller.ParticipantStatus.Waiting:
                        // Fall through to do waiting actions below

                        if (AdmitIfNeeded(p))
                        {
                            numAttending++;
                        }
                        else
                        {
                            numWaiting++;
                        }

                        continue;
                    default:
                        continue;
                }
            }

            if ((numAttending == 0) && (numWaiting > 0))
            {
                string waitMsg = cfg.WaitingRoomAnnouncementMessage;

                if (waitMsg == null)
                {
                    return;
                }

                if (waitMsg.Length == 0)
                {
                    return;
                }

                if (cfg.WaitingRoomAnnouncementDelaySecs <= 0)
                {
                    return;
                }

                // At least one person is in the waiting room.  If we're configured to make annoucements to them, do so now
                dtNow = DateTime.UtcNow;
                if (dtNow >= dtLastWaitingRoomAnnouncement.AddSeconds(cfg.WaitingRoomAnnouncementDelaySecs))
                {
                    // TBD: Sending to everyone in the waiting room is not yet available via the SDK.  For more details, see:
                    //   https://devforum.zoom.us/t/how-to-send-chat-messages-to-everyone-in-waiting-room/39538
                    _ = Controller.SendChatMessage(Controller.SpecialParticipant.everyoneInWaitingRoom, waitMsg);
                    dtLastWaitingRoomAnnouncement = dtNow;
                }
            }

            // Greet the first person to join the meeting, but only if we started Zoom
            if ((!Controller.ZoomAlreadyRunning) && (FirstParticipantGreeted == null) && (numAttending > 0))
            {
                // Looking for a participant that is not me, is using computer audio, and is a known good user
                var idx = plist.FindIndex(x => (
                    (!x.IsMe) &&
                    (x.Status == Controller.ParticipantStatus.Attending) &&
                    (x.audioDevice == Controller.ControllerAudioType.AUDIOTYPE_VOIP) &&
                    //UserLevels.ContainsKey(CleanUserName(x.Name))
                    CheckUserLevel(x.Name, UserLevel.Known)
                ));
                if (idx != -1)
                {
                    FirstParticipantGreeted = plist[idx].Name;
                    var msg = ChatBotConverse("_onetimehi_", plist[idx], "SimpleBot"); // Say "hi"

                    Thread.Sleep(3000);

                    Sound.Play("bootup");

                    Thread.Sleep(3000);

                    Sound.Speak(cfg.MyParticipantName + " online.");

                    if (Controller.SendChatMessage(Controller.SpecialParticipant.everyoneInMeeting, msg))
                    {
                        Sound.Speak(msg);
                    }
                }
            }
        }

        private static int nTimerIterationID = 0;

        /// <summary>Changes the specified mode to the specified state.</summary>
        /// <returns>Returns true if the state was changed.
        /// If the specified mode is already in the specified state, return false.</returns>
        private static void ReadRemoteCommands()
        {
            string sPath = @"command_file.txt";
            string line;

            if (!File.Exists(sPath))
            {
                return;
            }

            hostApp.Log(LogType.INF, "Processing Remote Commands");

            using (StreamReader sr = File.OpenText(sPath))
            {
                while ((line = sr.ReadLine()) != null)
                {
                    line = line.Trim();
                    if (line.Length == 0)
                    {
                        continue;
                    }

                    if ((line == "citadel:on") || (line == "citadel:off"))
                    {
                        SetMode("citadel", line.EndsWith(":on"));
                    }
                    if ((line == "lockdown:on") || (line == "lockdown:off"))
                    {
                        SetMode("lockdown", line.EndsWith(":on"));
                    }
                    else if ((line == "debug:on") || (line == "debug:off"))
                    {
                        SetMode("debug", line.EndsWith(":on"));
                    }
                    else if ((line == "pause:on") || (line == "pause:off"))
                    {
                        SetMode("pause", line.EndsWith(":on"));
                    }
                    else if ((line == "passive:on") || (line == "passive:off"))
                    {
                        SetMode("passive", line.EndsWith(":on"));
                    }
                    else if ((line == "exit") || (line == "leave"))
                    {
                        hostApp.Log(LogType.INF, "Received {0} command", line);
                        UsherBot.LeaveMeeting(false);
                    }
                    else if ((line == "kill") || (line == "end"))
                    {
                        hostApp.Log(LogType.INF, "Received {0} command", line);
                        UsherBot.LeaveMeeting(true);
                    }
                    else
                    {
                        hostApp.Log(LogType.ERR, "Unknown command: {0}", line);
                    }
                }
            }

            File.Delete(sPath);
        }

        public static void WriteRemoteCommands(string[] commands)
        {
            string sPath = @"command_file.txt";
            File.WriteAllText(sPath, string.Join(System.Environment.NewLine, commands));
        }

        private static void LoadUserLevels()
        {
            lock (UserLevels)
            {

                string sPath = $"{hostApp.GetWorkDir()}\\users.txt";

                if (!File.Exists(sPath))
                {
                    return;
                }

                DateTime dtLastMod = File.GetLastWriteTimeUtc(sPath);

                // Don't load/reload unless changed
                if (dtLastMod == dtLastGoodUserMod)
                {
                    return;
                }

                dtLastGoodUserMod = dtLastMod;

                hostApp.Log(LogType.INF, "(Re-)loading UserLevels");

                UserLevels.Clear();
                using (StreamReader sr = File.OpenText(sPath))
                {
                    string line = null;
                    while ((line = sr.ReadLine()) != null)
                    {
                        line = line.Trim();
                        if (line.Length == 0)
                        {
                            continue;
                        }

                        UserLevel userLevel = UserLevel.Known;

                        // CoHost entries end with "^"
                        if (line.EndsWith("^"))
                        {
                            line = line.TrimEnd('^');
                            userLevel = UserLevel.CoHost;
                        }

                        // Auto-CoHost entries that end with "%"
                        if (line.EndsWith("%"))
                        {
                            line = line.TrimEnd('%');
                            userLevel = UserLevel.AutoCoHost;
                        }

                        // Admin entries end with "@"
                        if (line.EndsWith("@"))
                        {
                            line = line.TrimEnd('@');
                            userLevel = UserLevel.Admin;
                        }

                        // Allow alises, delimited by "|"
                        string[] names = line.Split('|');
                        foreach (string name in names)
                        {
                            var cleanName = CleanUserName(name);
                            _ = UserLevels.TryGetValue(cleanName, out UserLevel oldUserLevel);

                            // If an entry for the user already exists, remove the duplicate and use the higher user level
                            if (userLevel > oldUserLevel)
                            {
                                UserLevels[cleanName] = userLevel;
                            }
                        }
                    }
                }
            }
        }

        private static void SaveUserLevels()
        {
            lock (UserLevels)
            {
                string sPath = $"{hostApp.GetWorkDir()}\\users.txt";

                hostApp.Log(LogType.INF, "Saving UserLevels");

                using (StreamWriter sw = File.CreateText(sPath))
                {
                    var keys = UserLevels.Keys.ToList();
                    keys.Sort();

                    foreach (var key in keys)
                    {
                        var userLevel = UserLevels[key];
                        var line = key;

                        if (userLevel == UserLevel.Admin)
                        {
                            line += '@';
                        }
                        else if (userLevel == UserLevel.CoHost)
                        {
                            line += '^';
                        }

                        sw.WriteLine(line);
                    }
                }

                dtLastGoodUserMod = File.GetLastWriteTimeUtc(sPath);
            }
        }

        /// <summary>
        /// Looks up user by 'name' and returns the user flags.  Returns UserFlags.UNKNOWN if the user does not exist.
        /// </summary>
        private static UserLevel GetUserLevel(string name)
        {
            lock (UserLevels)
            {
                UserLevels.TryGetValue(CleanUserName(name), out UserLevel ret);
                return ret;
            }
        }

        /// <summary>
        /// Looks up user by 'name' and returns True if the user is the given level or higher.
        /// </summary>
        private static bool CheckUserLevel(string name, UserLevel minLevel)
        {
            lock (UserLevels)
            {
                return GetUserLevel(name) >= minLevel;
            }
        }

        /// <summary>
        /// Updates the userType for the given user, adding to UserLevels if missing, or removing if newType = UNKNOWN.
        /// </summary>
        private static bool SetUserLevel(string name, UserLevel newUserLevel)
        {
            lock (UserLevels)
            {
                var cleanName = CleanUserName(name);
                var exists = UserLevels.TryGetValue(cleanName, out UserLevel oldUserLevel);

                if (oldUserLevel == newUserLevel)
                {
                    return false;
                }

                if (newUserLevel == UserLevel.Unknown)
                {
                    UserLevels.Remove(cleanName);
                }
                else
                {
                    UserLevels[cleanName] = newUserLevel;
                }

                SaveUserLevels();

                return true;
            }
        }

        private static readonly char[] SpaceDelim = new char[] { ' ' };

        /// <summary>
        /// Get rid of annoying iPhone stuff.  iPhone users joining Zoom are named "User's iPhone" or "iPhoneUser" etc.
        /// Example input/output: "User's iPhone" => "User".
        /// </summary>
        // Get rid of annoying iPhone stuff... X's iPhone iPhoneX etc.
        private static string RemoveIPhoneStuff(string name)
        {
            if (string.IsNullOrEmpty(name)) { return name; }

            var ret = FastRegex.Replace(name, @"’s iPhone", string.Empty, RegexOptions.IgnoreCase);
            ret = FastRegex.Replace(ret, @"â€™s iPhone", string.Empty, RegexOptions.IgnoreCase);
            ret = FastRegex.Replace(ret, @"�s iPhone", string.Empty, RegexOptions.IgnoreCase);
            ret = FastRegex.Replace(ret, @"'s iPhone", string.Empty, RegexOptions.IgnoreCase);
            ret = FastRegex.Replace(ret, @"\s*iPhone\s*", string.Empty, RegexOptions.IgnoreCase);

            // If the only thing in the name is "iPhone", leave it
            return (ret.Length == 0) ? name : ret;
        }

        private static string GetFirstName(string name)
        {
            if (string.IsNullOrEmpty(name))
            {
                return null;
            }

            name = name.Trim();
            if (name.Length == 0)
            {
                return null;
            }

            name = RemoveIPhoneStuff(name);

            string firstName = name.Split(SpaceDelim)[0];

            // If the first letter is capitalized, and it's not *all* uppercase, assume it's cased correctly.  Ex: Joe, JoeAnne
            if ((firstName.Substring(0, 1) == firstName.Substring(0, 1).ToUpper()) && (firstName != firstName.ToUpper()))
            {
                return firstName;
            }

            // It's either all uppercase or all lower case, so title case it
            return firstName.ToTitleCase();
        }

        private static string GetDayTime()
        {
            DateTime dtNow = DateTime.Now;
            if (dtNow.Hour >= 12 && dtNow.Hour < 17)
            {
                return "afternoon";
            }
            else if (dtNow.Hour < 12)
            {
                return "morning";
            }

            return "evening";
        }

        private static string GetTodayTonight()
        {
            DateTime dtNow = DateTime.Now;

            return (dtNow.Hour < 17) ? "today" : "tonight";
        }

        private static string ResolveVars(string text, string to)
        {
            StringBuilder ret = new StringBuilder();
            StringBuilder accum = null;

            foreach (var ch in text)
            {
                switch (ch)
                {
                    case '{':
                        if (accum != null)
                        {
                            throw new FormatException($"Cannot have open brace inside braces: {repr(text)}");
                        }

                        accum = new StringBuilder();

                        break;
                    case '}':
                        if (accum == null)
                        {
                            throw new FormatException($"Cannot have close brace without open brace: {repr(text)}");
                        }

                        // TBD: Could load all of this into a Dictionary<string, string> ...
                        string key = accum.ToString();
                        string val = null;
                        if ((key == "0") || (key == "firstname"))
                        {
                            val = GetFirstName(to);
                        }
                        else if ((key == "1") || (key == "timeofday"))
                        {
                            val = GetDayTime();
                        }
                        else if (key == "fullname")
                        {
                            val = to;
                        }
                        else if (key == "topic")
                        {
                            val = GetTopic(true);
                        }
                        else
                        {
                            if (cfg.BroadcastCommands != null)
                            {
                                cfg.BroadcastCommands.TryGetValue(key, out val);
                            }
                        }

                        if (val == null)
                        {
                            throw new ArgumentException($"Variable {key} could not be resolved: {repr(text)}");
                        }

                        ret.Append(val);
                        accum = null;

                        break;
                    default:
                        if (accum != null)
                        {
                            accum.Append(ch);
                        }
                        else
                        {
                            ret.Append(ch);
                        }

                        break;
                }
            }

            return ret.ToString();
        }

        private static readonly Dictionary<uint, string> DicTopicSends = new Dictionary<uint, string>();

        private static void SetSpeaker(Controller.Participant newSpeaker, Controller.Participant from)
        {
            if (newSpeaker == null)
            {
                if (speaker == null)
                {
                    _ = Controller.SendChatMessage(from, "Speaker mode is already off");
                    return;
                }

                _ = Controller.SendChatMessage(from, "Turning speaker mode off");

                hostApp.Log(LogType.INF, "Allowing participants to unmute themselves");
                _ = Controller.SetAllowParticipantsToUnmuteSelf(true);

                speaker = null;
                return;
            }

            if (newSpeaker.Name == speaker)
            {
                _ = Controller.SendChatMessage(from, $"Speaker is already {newSpeaker}");
                return;
            }

            _ = Controller.SendChatMessage(from, "Turning speaker mode on");

            hostApp.Log(LogType.INF, "Disallowing participants to unmute themselves");
            _ = Controller.SetAllowParticipantsToUnmuteSelf(false);

            // Unmute speaker and mute everyone else
            foreach (var p in Controller.participants.Values)
            {
                // Skip over me, host & cohost
                if (p.IsMe || p.IsHost || p.IsCoHost)
                {
                    continue;
                }

                if (p.UserId == newSpeaker.UserId)
                {
                    if (p.IsAudioMuted)
                    {
                        hostApp.Log(LogType.INF, $"Unmuting {p}");
                        Controller.UnmuteParticipant(p);
                    }
                }
                else
                {
                    if (!p.IsAudioMuted)
                    {
                        hostApp.Log(LogType.INF, $"Muting {p}");
                        Controller.MuteParticipant(p);
                    }
                }
            }

            _ = Controller.SendChatMessage(from, $"Speaker set to {newSpeaker}");
            speaker = newSpeaker.Name;
        }

        private static List<IChatBot> chatBots = null;

        /// <summary>
        /// Searches for ChatBot plugins under plugins\ChatBots\{BotName}\ZoomMeetingBotSDK.ChatBot.{BotName}.dll and tries to instantiate them,
        /// returning a list of ones that succeeded.  The list is ordered ascending by "order".
        ///
        /// NOTE: We put the plugins in their own directories to allow them to use whatever .Net verison and dependency libraries they'd
        /// like without confliciting with those used by the main process.
        /// </summary>
        public static List<IChatBot> GetChatBots()
        {
            var bots = new List<Tuple<int, IChatBot>>();
            var botPluginDir = new DirectoryInfo(Path.Combine(Environment.CurrentDirectory, @"plugins\ChatBot"));
            if (!botPluginDir.Exists)
            {
                return null;
            }

            foreach (var subdir in botPluginDir.GetDirectories())
            {
                FileInfo[] files = subdir.GetFiles("ZoomMeetingBotSDK.ChatBot.*.dll");
                if (files.Length > 1)
                {
                    hostApp.Log(LogType.WRN, $"Cannot load bot in {repr(subdir.FullName)}; More than one DLL found");
                }
                else if (files.Length == 0)
                {
                    hostApp.Log(LogType.WRN, $"Cannot load bot in {repr(subdir.FullName)}; No DLL found");
                }
                else
                {
                    var file = files[0];
                    try
                    {
                        hostApp.Log(LogType.DBG, $"Loading {file.Name}");
                        var assembly = Assembly.LoadFile(file.FullName);
                        var types = assembly.GetTypes();
                        foreach (var type in types)
                        {
                            List<Type> interfaceTypes = new List<Type>(type.GetInterfaces());
                            if (interfaceTypes.Contains(typeof(IChatBot)))
                            {
                                var chatBot = Activator.CreateInstance(type) as IChatBot;
                                var chatBotInfo = chatBot.GetChatBotInfo();
                                chatBot.Init(new ChatBotInitParam()
                                {
                                    hostApp = hostApp,
                                });
                                hostApp.Log(LogType.DBG, $"Loaded {repr(chatBotInfo.Name)} chatbot with order {chatBotInfo.DefaultOrder}");
                                chatBot.Start();
                                bots.Add(new Tuple<int, IChatBot>(chatBotInfo.DefaultOrder, chatBot));
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        hostApp.Log(LogType.ERR, $"Failed to load {repr(file.FullName)}: {repr(ex.ToString())}");
                    }
                }
            }

            if (bots.Count == 0)
            {
                return null;
            }

            return bots.OrderBy(o => o.Item1).Select(x => x.Item2).ToList();
        }

        public static void LeaveMeeting(bool endMeetingForAll = true)
        {
            if (leaveMeeting)
            {
                hostApp.Log(LogType.WRN, "[UsherBot] LeaveMeeting - Already leaving");
                return;
            }
            leaveMeeting = true;

            hostApp.Log(LogType.WRN, $"[UsherBot] LeaveMeeting endForAll={endMeetingForAll}");
            endForAll = endMeetingForAll;
        }

        /// <summary>
        /// Leaves the meeting, optionally ending meeting or passing off Host role to another participant.
        /// </summary>
        private static void ReallyLeaveMeeting()
        {
            if (leavingMeeting)
            {
                hostApp.Log(LogType.WRN, "[UsherBot] ReallyLeaveMeeting - Already leaving");
                return;
            }

            leavingMeeting = true;

            if (!endForAll)
            {
                if (!Controller.me.IsHost)
                {
                    hostApp.Log(LogType.DBG, "[UsherBot] ReallyLeaveMeeting - I am not host");
                }
                else
                {
                    hostApp.Log(LogType.DBG, "[UsherBot] ReallyLeaveMeeting - I am host; Trying to find someone to pass it to");

                    // Choose the highest ranking participant to pass it to
                    //   Rank: 10 * UserLevel + 1 if they are CoHost
                    // In other words, we'll pick the best alt host in order of preference:
                    //   Admin + IsCoHost
                    //   CoHost + IsCoHost
                    //   Known + IsCoHost
                    //   IsCoHost
                    int altHostRank = 0;
                    Controller.Participant altHost = null;
                    foreach (Controller.Participant p in Controller.participants.Values)
                    {
                        var pUserLevel = GetUserLevel(p.Name);
                        var pRank = (10 * (int)pUserLevel) + (p.IsCoHost ? 1 : 0);

                        if (pRank > altHostRank)
                        {
                            altHostRank = pRank;
                            altHost = p;
                            hostApp.Log(LogType.DBG, $"[UserBot] ReallyLeaveMeeting - Hunting for alt host; Best so far: {repr(p.Name)} userLevel={pUserLevel} isCoHost={p.IsCoHost} rank={pRank}");
                        }
                    }

                    if (altHostRank == 0) {
                        hostApp.Log(LogType.ERR, "[UsherBot] ReallyLeaveMeeting - Could not find an alternative host; Ending meeting");
                        endForAll = true;
                    }
                    else
                    {
                        hostApp.Log(LogType.INF, $"[UsherBot] ReallyLeaveMeeting - Passing Host to {altHost}");
                        if (Controller.PromoteParticipant(altHost, Controller.ParticipantRole.Host))
                        {
                            hostApp.Log(LogType.INF, $"[UsherBot] ReallyLeaveMeeting - Passed Host to {altHost}");
                        }
                        else
                        {
                            hostApp.Log(LogType.ERR, $"[UsherBot] ReallyLeaveMeeting - Failed to pass Host to {altHost}; Ending meeting");
                            endForAll = true;
                        }
                    }
                }
            }

            hostApp.Log(LogType.INF, "[UsherBot] ReallyLeaveMeeting - Really Leaving Meeting");
            _ = Controller.LeaveMeeting(endForAll);
        }

        /*
        protected virtual void Dispose(bool disposing)
        {
            if (!disposing)
            {
                return;
            }

            tmrIdle.Dispose();
        }
        */

        private static GmailSenderLib.GmailSender gmailSender = null;
        private static bool SendEmail(string subject, string body, string to)
        {
            try
            {
                if (gmailSender is null)
                {
                    gmailSender = new GmailSenderLib.GmailSender(System.Reflection.Assembly.GetCallingAssembly().GetName().Name, hostApp.GetWorkDir());
                }

                hostApp.Log(LogType.ERR, "SendEmail - Sending email to {0} with subject {1}", repr(to), repr(subject));
                gmailSender.Send(new GmailSenderLib.SimpleMailMessage(subject, body, to));

                return true;
            }
            catch (Exception ex)
            {
                hostApp.Log(LogType.ERR, "SendEmail - Failed; Exception: {0}", repr(ex.ToString()));
                return false;
            }
        }

        public void Init(ControlBotInitParam param)
        {
            hostApp = param.hostApp;
            LoadSettings();

            hostApp.SettingsChanged += new EventHandler(SettingsChanged);
        }

        /// <summary>
        /// Returns a delmited list of participants.
        /// </summary>
        /// <param name="participants">The list of participants.</param>
        /// <param name="includeIDs">If true, includes names and IDs, otherwise just names.</param>
        /// <param name="delim">Delimiter to use.</param>
        /// <returns>A delimited list string.</returns>
        private string ParticipantListToString(List<Controller.Participant> participants, string delim = " | ")
        {
            if (participants.Count() == 0)
            {
                return "None";
            }

            return cfg.IncludeUserIdsInLists ?
                participants.Select(x => x.ToString()).Aggregate((a, b) => a + delim + b) :
                participants.Select(x => x.Name).Aggregate((a, b) => a + delim + b);
        }

        public void Start()
        {
            Controller.Init(hostApp);
            Sound.Init(hostApp);

            if ((cfg.BotAutomationFlags & BotAutomationFlag.Converse) != 0)
            {
                chatBots = GetChatBots();
            }

            Controller.OnChatMessageReceive += Controller_OnChatMessageReceive;
            Controller.OnParticipantJoinWaitingRoom += Controller_OnParticipantJoinWaitingRoom;
            Controller.OnParticipantLeaveWaitingRoom += Controller_OnParticipantLeaveWaitingRoom;
            Controller.OnParticipantJoinMeeting += Controller_OnParticipantJoinMeeting;
            Controller.OnParticipantLeaveMeeting += Controller_OnParticipantLeaveMeeting;
            Controller.OnParticipantActiveAudioChange += Controller_OnParticipantActiveAudioChange;
            Controller.OnParticipantRaisedHandsChange += Controller_OnParticipantRaisedHandsChange;
            Controller.OnActionTimerTick += Controller_OnActionTimerTick;
            Controller.OnExit += Controller_OnExit;

            Controller.Start();

            //tmrIdle = new System.Threading.Timer(ActionTimer, null, 0, 5000);

            return;
        }

        private static List<Controller.Participant> activeAudioParticipants = new List<Controller.Participant>();
        private static List<Controller.Participant> raisedHandParticipants = new List<Controller.Participant>();

        private void Controller_OnParticipantActiveAudioChange(object sender, Controller.OnParticipantActiveAudioChangeArgs e)
        {
            activeAudioParticipants = e.activeAudioParticipants;
            NotifyTrackers("talkers", ParticipantListToString(activeAudioParticipants));
        }

        private void Controller_OnParticipantRaisedHandsChange(object sender, Controller.OnParticipantRaisedHandsChangeArgs e)
        {
            raisedHandParticipants = e.raisedHandParticipants;
            NotifyTrackers("hands", ParticipantListToString(raisedHandParticipants));
        }

        private void Controller_OnActionTimerTick(object sender, EventArgs e)
        {
            if (leaveMeeting)
            {
                ReallyLeaveMeeting();
                return;
            }

            if (shouldExit)
            {
                return;
            }

            Interlocked.Increment(ref nTimerIterationID);

            if (!Monitor.TryEnter(_lock_eh))
            {
                hostApp.Log(LogType.WRN, "ActionTimer {0:X4} - Busy; Will try again later", nTimerIterationID);
                return;
            }

            try
            {
                //hostApp.Log(LogType.DBG, "ActionTimer {0:X4} - Enter");

                LoadUserLevels();
                ReadRemoteCommands();

                if (cfg.IsPaused)
                {
                    return;
                }

                //hostApp.Log(LogType.DBG, "ActionTimer {0:X4} - DoParticipantActions", nTimerIterationID);
                DoParticipantActions();
            }
            /* TBD: Do something about this?
            catch (Controller.ZoomClosedException ex)
            {
                hostApp.Log(LogType.INF, ex.ToString());
                ShouldExit = true;
            }
            */
            catch (Exception ex)
            {
                hostApp.Log(LogType.ERR, "ActionTimer {0:X4} - Unhandled Exception: {1}", nTimerIterationID, ex.ToString());
            }
            finally
            {
                //hostApp.Log(LogType.DBG, "ActionTimer {0:X4} - Exit", nTimerIterationID);
                Monitor.Exit(_lock_eh);
            }
        }

        private void Controller_OnExit(object sender, EventArgs e)
        {
            Stop();
        }

        private static bool AdmitIfNeeded(Controller.Participant p)
        {
            var bAdmitKnown = (cfg.BotAutomationFlags & BotAutomationFlag.AdmitKnown) != 0;
            var bAdmitOthers = (cfg.BotAutomationFlags & BotAutomationFlag.AdmitOthers) != 0;
            var waitMsg = String.Empty;

            if (p.IsMe || (!bAdmitKnown && !bAdmitOthers))
            {
                // Nothing to do
                return false;
            }

            //var sCleanName = CleanUserName(p.Name);

            bool blockUser = false;
            if (BadUsers.ContainsKey(p.UserId) || BadUsers.Values.Contains(p.Name))
            {
                blockUser = true;
                waitMsg = $"BOT Not admitting {p} : BAD USER";
            }

            if ((badNameRegex != null) && badNameRegex.IsMatch(p.Name))
            {
                blockUser = true;
                waitMsg = $"BOT Not admitting {p} : BAD NAME";
            }

            if (blockUser)
            {
                // Make sure we don't display the message more than once
                if (!HsParticipantMessages.Contains(waitMsg))
                {
                    hostApp.Log(LogType.INF, waitMsg);
                    HsParticipantMessages.Add(waitMsg);
                }

                return false;
            }

            if (HsParticipantMessages.Contains(waitMsg))
            {
                HsParticipantMessages.Remove(waitMsg);
            }

            //if (UserLevels.ContainsKey(sCleanName) && bAdmitKnown)
            if (CheckUserLevel(p.Name, UserLevel.Known) && bAdmitKnown)
            {
                hostApp.Log(LogType.INF, "BOT Admitting {p} : KNOWN");
                return Controller.AdmitParticipant(p);
            }

            if (!bAdmitOthers)
            {
                // Nothing to do
                return false;
            }

            // Admitting an unknown user

            bool bAdmit = false;

            DateTime dtNow = DateTime.UtcNow;
            DateTime dtWhenToAdmit = p.WaitingRoomEntryDT.AddSeconds(cfg.UnknownParticipantWaitSecs);
            if (dtWhenToAdmit < dtNow)
            {
                // Too early to admit this participant
                return false;
            }

            dtWhenToAdmit = dtLastAdmission.AddSeconds(cfg.UnknownParticipantWaitSecs);
            bAdmit = dtNow >= dtWhenToAdmit;

            waitMsg = $"BOT Admit {p} : Unknown participant waiting room time reached";
            if (bAdmit)
            {
                waitMsg += " : Admitting";
            }

            // Make sure we don't display the message more than once
            if (!HsParticipantMessages.Contains(waitMsg))
            {
                hostApp.Log(LogType.INF, waitMsg);
                HsParticipantMessages.Add(waitMsg);
            }

            if (bAdmit && Controller.AdmitParticipant(p))
            {
                // User was successfully admitted; Remove the message from the queue
                HsParticipantMessages.Remove(waitMsg);

                // Caculate next admission time
                dtLastAdmission = dtNow;

                return true;
            }

            return false;
        }

        private static bool PromoteIfNeeded(Controller.Participant p)
        {
            // Handle automatically co-hosting folks here if needed

            if ((p.IsMe) || (p.IsCoHost))
            {
                // I can't promote myself, and I can't co-host someone who is already co-hosted
                return false;
            }

            if ((cfg.BotAutomationFlags & BotAutomationFlag.CoHostKnown) == 0)
            {
                // Nothing to do
                return false;
            }

            //var cleanName = CleanUserName(p.Name);
            //UserLevels.TryGetValue(cleanName, out bool bUserShouldBeCoHost);
            //if (!CheckUserLevel(p.Name, UserLevel.CoHost))
            if (!CheckUserLevel(p.Name, UserLevel.AutoCoHost))
            {
                // Nothing to do
                return false;
            }

            var msg = $"BOT Not co-hosting {p} : BAD USER";
            if (BadUsers.ContainsKey(p.UserId) || BadUsers.Values.Contains(p.Name))
            {
                // Make sure we don't display the message more than once
                if (!HsParticipantMessages.Contains(msg))
                {
                    hostApp.Log(LogType.INF, msg);
                    HsParticipantMessages.Add(msg);
                }

                return false;
            }

            if (HsParticipantMessages.Contains(msg))
            {
                HsParticipantMessages.Remove(msg);
            }

            if ((!Controller.me.IsHost) && (!Controller.me.IsCoHost))
            {
                hostApp.Log(LogType.WRN, $"BOT Participant {p} should be Co-Host, but I am not Co-Host or Host");
                return false;
            }

            hostApp.Log(LogType.INF, $"BOT Promoting {p} to Co-host");
            return Controller.PromoteParticipant(p, Controller.ParticipantRole.CoHost);
        }

        private void Controller_OnParticipantJoinMeeting(object sender, Controller.OnParticipantJoinMeetingArgs e)
        {
            var p = e.participant;

            // Send the topic if configured to do so
            if ((cfg.BotAutomationFlags & BotAutomationFlag.SendTopicOnJoin) != 0)
            {
                _ = SendTopic(p, false);
            }

            PromoteIfNeeded(p);
        }

        private void Controller_OnParticipantLeaveMeeting(object sender, Controller.OnParticipantLeaveMeetingArgs e)
        {
            // Nothing to do yet ...
        }

        private void Controller_OnParticipantJoinWaitingRoom(object sender, Controller.OnParticipantJoinWaitingRoomArgs e)
        {
            AdmitIfNeeded(e.participant);
        }

        private void Controller_OnParticipantLeaveWaitingRoom(object sender, Controller.OnParticipantLeaveWaitingRoomArgs e)
        {
            // Nothing to do for now ...
        }

        private static string EnableTracking(uint userId, string trackingItem)
        {
            lock (trackers)
            {
                var bEnabled = trackers.TryGetValue(userId, out HashSet<string> items) && items.Contains(trackingItem);

                if (!bEnabled)
                {
                    if (items == null)
                    {
                        items = new HashSet<string>();
                    }

                    items.Add(trackingItem);

                    trackers[userId] = items;
                }

                return $"Tracking of {trackingItem} {(bEnabled ? "is already enabled" : "enabled")}";
            }
        }

        private static void NotifyTrackers(string item, string value, uint dontSendToUserId = 0)
        {
            lock (trackers)
            {
                foreach (var tracker in trackers.ToList())
                {
                    if (tracker.Value.Contains(item))
                    {
                        if (!Controller.GetParticipantById(tracker.Key, out Controller.Participant p))
                        {
                            // We couldn't get the participant object; Remove from tracker list
                            trackers.Remove(tracker.Key);
                        }
                        else
                        {
                            if ((dontSendToUserId == 0) || (dontSendToUserId != tracker.Key))
                            {
                                _ = Controller.SendChatMessage(p, $"{item.UppercaseFirst()}: {value}");
                            }
                        }
                    }
                }
            }

        }

        private static readonly string[] scrubTTSBadTexts = new string[] { "@", "http", "Forecast" };
        private static readonly Regex scrubTTSRegex = new Regex(@"^(.*?)<tts>(.*?)</tts>(.*?)$", RegexOptions.IgnoreCase | RegexOptions.Singleline);
        /// <summary>
        /// Pre-processes TTS text, removing excessive or difficult to pronounce things such as email addresses, web pages, etc.
        /// </summary>
        private static void ScrubTTS(string input, out string text, out string tts)
        {

            // If text contains <tts>...</tts>, then use what's inside of the tags
            var match = scrubTTSRegex.Match(input);
            if ((match != null) && (match.Groups.Count == 4))
            {
                text = (match.Groups[1].Value + " " + match.Groups[3].Value).Trim();
                tts = match.Groups[2].Value.Trim();
                return;
            }

            text = input;
            var lines = input.GetLines();
            var ret = new List<string>();
            foreach (var line in lines)
            {
                var done = false;
                foreach (var badText in scrubTTSBadTexts)
                {
                    if (line.Contains(badText))
                    {
                        done = true;
                        break;
                    }
                }

                if (done)
                {
                    break;
                }

                ret.Add(line.Trim());
            }

            tts = string.Join("\n", ret);
        }

        private static bool SendChatAndSpeak(Controller.Participant to, string input, bool speak = true)
        {
            ScrubTTS(input, out string text, out string tts);

            var ret = ((text != null) && (text.Length > 0)) ? Controller.SendChatMessage(to, text) : true;

            if (ret && speak && (tts != null) && (tts.Length > 0))
            {
                Sound.Speak(tts);
            }

            return ret;
        }

        private class ChatBotSender : IChatBotUser
        {
            public ChatBotSender(Controller.Participant p)
            {
                Name = p.Name;
                UserId = p.UserId;
            }

            public string Name { get; set; }
            public uint UserId { get; set; }
        }

        private static string ChatBotConverse(string text, Controller.Participant from, string specificBot = null, bool isToEveryone = false)
        {
            string response = null;
            string botName = null;

            // No sense in trying to do anything if we don't have any ChatBots available
            if ((chatBots == null) || (chatBots.Count == 0))
            {
                return response;
            }

            // We'll try each bot in order until one of them provides a response
            foreach (var chatBot in chatBots)
            {
                botName = chatBot.GetChatBotInfo().Name;
                if ((specificBot != null) && (botName != specificBot))
                {
                    // Looking for a specific bot, and this one ain't it!
                    continue;
                }

                string failureMsg = null;
                try
                {
                    response = chatBot.Converse(text, new ChatBotSender(isToEveryone ? Controller.SpecialParticipant.everyoneInMeeting : from));
                    if (response == null)
                    {
                        failureMsg = "Response is null";
                    }
                }
                catch (Exception ex)
                {
                    failureMsg = "Exception occurred: " + ex.ToString();
                    response = null;
                }

                if (response != null)
                {
                    break;
                }

                //hostApp.Log(LogType.WRN, $"ChatBot converse with {repr(chatBot.GetChatBotInfo().Name)} failed: {repr(failureMsg)}");
            }

            if (response == null)
            {
                hostApp.Log(LogType.ERR, "No ChatBot was able to produce a response");
                return response;
            }

            hostApp.Log(LogType.INF, $"chatBot {botName} > {from}: {repr(response)}");

            response = ResolveVars(response, from.Name);

            /*
            var idx = response.IndexOf(" /");
            if (idx > 1)
            {
                // Allow for the case where we have a message with a command.  For example: "Hi there! /donate".
                _ = SendChatAndSpeak(from, response.Substring(0, idx), false);
                response = response.Substring(idx+1);
            }
            */

            return response;
        }

        private void HandleChatMessageToMe(Controller.Participant from, Controller.Participant replyTo, string text, bool isToEveryone, bool onlyTwoAttendees)
        {
            var isPrivate = !isToEveryone;

            // ====
            // Handle chit-chat
            // ====

            // All commands start with "/"; Treat everything else as small talk
            if (!text.StartsWith("/"))
            {
                var response = ChatBotConverse(text, from, null, isToEveryone);

                if (response == null)
                {
                    // No ChatBot response, so just ignore the incoming message
                    return;
                }

                if (!response.StartsWith("/"))
                {
                    // If the bot is addressed publically or if there are only two people in the meeting, then reply with TTS
                    var speak = isToEveryone || onlyTwoAttendees;

                    _ = SendChatAndSpeak(replyTo, response, speak);

                    return;
                }

                // The ChatBot wants the control bot to run a command; Pass it along
                text = response;
                isPrivate = true;
            }

            // ====
            // Handle non-priviledged commands
            // ====

            if (!isPrivate)
            {
                // Ignore commands outside of private messages
                return;
            }

            string[] a = text.Split(SpaceDelim, 2);

            string command = a[0].ToLower().Substring(1);

            // Parse out command argument if given
            string target = (a.Length == 1) ? null : (a[1].Length == 0 ? null : a[1]);

            // Determine if sender is cohost or not
            //UserLevels.TryGetValue(CleanUserName(from.Name), out bool bAdmin);

            // Non-privileged commands (User is no a CoHost or Admin)
            if (!CheckUserLevel(from.Name, UserLevel.CoHost))
            {
                // Always reply directly to sender here
                replyTo = from;

                if (command == "topic")
                {
                    SendTopic(replyTo, true);
                    return;
                }

                if (cfg.BroadcastCommands.TryGetValue(command, out string response))
                {
                    //_ = Controller.SendChatMessage(replyTo, response);
                    _ = SendChatAndSpeak(replyTo, response, false);
                    return;
                }

                _ = Controller.SendChatMessage(from, $"I'm sorry, you are not authorized to run that command.");

                return;
            }

            // ====
            // Handle priviledged commands
            // ====

            if (cfg.BroadcastCommands.TryGetValue(command, out string broadcastMsg))
            {
                DateTime dtNow = DateTime.UtcNow;

                if (BroadcastSentTime.TryGetValue(command, out DateTime dtSentTime))
                {
                    int guardTime = cfg.BroadcastCommandGuardTimeSecs;

                    if (guardTime < 0)
                    {
                        _ = Controller.SendChatMessage(replyTo, $"{command}: This broadcast message was already sent.");
                        return;
                    }

                    if ((guardTime > 0) && (dtNow <= dtSentTime.AddSeconds(cfg.BroadcastCommandGuardTimeSecs)))
                    {
                        _ = Controller.SendChatMessage(replyTo, $"{command}: This broadcast message was already sent recently. Please try again later.");
                        return;
                    }
                }

                //if (Controller.SendChatMessage(Controller.SpecialParticipant.everyoneInMeeting, broadcastMsg))
                if (SendChatAndSpeak(Controller.SpecialParticipant.everyoneInMeeting, broadcastMsg, false))
                {
                    BroadcastSentTime[command] = dtNow;
                }

                return;
            }

            // Priv retrival or set of topic
            if (command == "topic")
            {
                if (target == null)
                {
                    _ = SendTopic(replyTo, true);
                    return;
                }

                bool broadcast = false;
                string reply;

                string[] b = target.Split(SpaceDelim, 2);

                string cmd = b[0].ToLower().TrimStart('/');

                if (cmd == "force")
                {
                    if ((b[1] != null) && (!b[1].EndsWith(".")))
                    {
                        b[1] += ".";
                    }

                    Topic = b[1];

                    reply = "Topic forced to: " + Topic;
                    broadcast = true;
                }
                else if ((cmd == "clear") || (cmd == "off"))
                {
                    if (Topic == null)
                    {
                        reply = "The topic has not been set; There is nothing to clear";
                    }
                    else
                    {
                        reply = "Topic cleared";
                        Topic = null;
                    }
                }
                else if (string.Compare(Topic, target, true) == 0)
                {
                    reply = "The topic is already set to: " + target;
                }
                else if (Topic == null)
                {
                    if ((target != null) && (!target.EndsWith(".")))
                    {
                        target += ".";
                    }

                    reply = "Topic set to: " + target;
                    Topic = target;
                    broadcast = true;
                }
                else
                {
                    reply = "Topic is already set; Use /topic force to change it";
                }

                _ = Controller.SendChatMessage(replyTo, reply);

                if (broadcast)
                {
                    _ = SendChatAndSpeak(Controller.SpecialParticipant.everyoneInMeeting, GetTopic(), false);
                }

                return;
            }

            // All of the following commands require options

            if (target == null)
            {
                return;
            }

            // Do remember/forget stuff:
            //   /remember Participant - Remember participant (Will be added to UserLevels)
            //   /forget Participant - Forget participant (Will be removed from UserLevels completely)
            //   /remember /admin Participant - Remember participant is admin
            //   /remember /cohost Participant - Remember participant is co-host
            //   /remember /cohost /auto Participant - Remember participant is auto co-host

            var before = text;
            var rftext = FastRegex.Replace(before, @"\s*/remember\s*", string.Empty, RegexOptions.IgnoreCase);
            var remember = rftext != before;

            before = rftext;
            rftext = FastRegex.Replace(before, @"\s*/forget\s*", string.Empty, RegexOptions.IgnoreCase);
            var forget = rftext != before;

            before = rftext;
            rftext = FastRegex.Replace(before, @"\s*/auto\s*", string.Empty, RegexOptions.IgnoreCase);
            var auto = rftext != before;

            if (remember || forget)
            {
                // These commands require admin
                if (!CheckUserLevel(from.Name, UserLevel.Admin))
                {
                    _ = Controller.SendChatMessage(replyTo, $"I'm sorry, but you're not authorized to run that command.");
                    return;
                }

                if (remember && forget)
                {
                    _ = Controller.SendChatMessage(replyTo, $"I'm pretty smart, but I'm not sure how to remember AND forget at the same time!");
                    return;
                }

                before = rftext;
                rftext = FastRegex.Replace(before, @"\s*/cohost\s*", string.Empty, RegexOptions.IgnoreCase);
                var cohost = rftext != before;

                before = rftext;
                rftext = FastRegex.Replace(before, @"\s*/admin\s*", string.Empty, RegexOptions.IgnoreCase);
                var admin = rftext != before;

                target = rftext;
                if (rftext.Length == 0)
                {
                    _ = Controller.SendChatMessage(replyTo, $"Who do you want me to {(remember ? "remember" : "forget")}?");
                    return;
                }

                var oldUserLevel = GetUserLevel(target);
                var newUserLevel = oldUserLevel;

                if (remember)
                {
                    if (admin)
                    {
                        newUserLevel = UserLevel.Admin;
                    }
                    else if (cohost)
                    {
                        if (auto)
                        {
                            newUserLevel = UserLevel.AutoCoHost;
                        }
                        else
                        {
                            newUserLevel = UserLevel.CoHost;
                        }
                    }
                    else
                    {
                        newUserLevel = UserLevel.Known;
                    }
                }
                else if (forget)
                {
                    if (admin)
                    {
                        newUserLevel = UserLevel.CoHost;
                    }
                    else if (cohost)
                    {
                        if (auto)
                        {
                            newUserLevel = UserLevel.CoHost;
                        }
                        else
                        {
                            newUserLevel = UserLevel.Known;
                        }
                    }
                    else
                    {
                        newUserLevel = UserLevel.Unknown;
                    }
                }
                else
                {
                    throw new Exception("Should never get here");
                }

                string response;
                if (oldUserLevel == newUserLevel)
                {
                    response = $"{repr(target)} is already {newUserLevel}.";
                }
                else
                {
                    SetUserLevel(target, newUserLevel);

                    response = $"OK, I'll remember that {repr(target)} is {newUserLevel} instead of {oldUserLevel}.";
                }

                _ = Controller.SendChatMessage(replyTo, response);

                return;
            }

            // Do email commands

            if (cfg.EmailCommands != null)
            {
                if (cfg.EmailCommands.TryGetValue(command, out EmailCommandArgs emailCommandArgs))
                {
                    string[] args = target.Trim().Split(SpaceDelim, 2);
                    string subject = emailCommandArgs.Subject;
                    string body = emailCommandArgs.Body;
                    string toAddress = args[0];

                    // If the command has any variables in it, assume more arguments are needed (ex: recipient name)
                    if (subject.Contains("{") || body.Contains("{"))
                    {
                        if (args.Length <= 1)
                        {
                            _ = Controller.SendChatMessage(replyTo, $"Error: The format of the command is incorrect; Correct example: /{command} {emailCommandArgs.ArgsExample}");
                            return;
                        }

                        string toName = args[1].Trim();
                        subject = ResolveVars(subject, toName);
                        body = ResolveVars(body, toName);
                    }

                    if (SendEmail(subject, body, toAddress))
                    {
                        _ = Controller.SendChatMessage(replyTo, $"{command}: Successfully sent email to {toAddress}");
                    }
                    else
                    {
                        _ = Controller.SendChatMessage(replyTo, $"{command}: Failed to send email to {toAddress}");
                    }

                    return;
                }
            }

            // Do security commands

            if ((command == "citadel") || (command == "lockdown") || (command == "passive") || (command == "lock"))
            {
                if (command == "lock")
                {
                    command = "lockdown";
                }

                string sNewMode = target.ToLower().Trim();
                bool bNewMode;

                if (sNewMode == "on")
                {
                    bNewMode = true;
                }
                else if (sNewMode == "off")
                {
                    bNewMode = false;
                }
                else
                {
                    _ = Controller.SendChatMessage(replyTo, $"Sorry, the {command} command requires either on or off as a parameter");
                    return;
                }

                if (SetMode(command, bNewMode))
                {
                    _ = Controller.SendChatMessage(replyTo, $"{command} mode has been changed to {sNewMode}");
                }
                else
                {
                    _ = Controller.SendChatMessage(replyTo, $"{command} mode is already {sNewMode}");
                }

                return;
            }

            // ...

            if (command == "ids")
            {
                string arg = target.ToLower().Trim();
                bool newModeBool = arg == "on" || arg == "enabled";
                string newModeStr = newModeBool ? "on" : "off";


                var response = string.Empty;
                if (newModeBool != cfg.IncludeUserIdsInLists)
                {
                    response = $"IncludeUserIdsInLists has been turned {newModeStr}";
                    cfg.IncludeUserIdsInLists = newModeBool;
                }
                else
                {
                    response = $"IncludeUserIdsInLists is already {newModeStr}";
                }

                _ = Controller.SendChatMessage(replyTo, response);
            }

            if (command == "waitmsg")
            {
                var sWaitMsg = target.Trim();

                if ((sWaitMsg.Length == 0) || (sWaitMsg.ToLower() == "off"))
                {
                    if ((cfg.WaitingRoomAnnouncementMessage != null) && (cfg.WaitingRoomAnnouncementMessage.Length > 0))
                    {
                        cfg.WaitingRoomAnnouncementMessage = null;
                        _ = Controller.SendChatMessage(replyTo, "Waiting room message has been turned off");
                    }
                    else
                    {
                        _ = Controller.SendChatMessage(replyTo, "Waiting room message is already off");
                    }
                }
                else if (sWaitMsg == cfg.WaitingRoomAnnouncementMessage)
                {
                    _ = Controller.SendChatMessage(replyTo, $"Waiting room message is already set to:\n{target}");
                }
                else
                {
                    cfg.WaitingRoomAnnouncementMessage = target.Trim();
                    _ = Controller.SendChatMessage(replyTo, $"Waiting room message has set to:\n{target}");
                }

                return;
            }

            // Pre-processing for rename action
            string newName = null;
            if (command == "rename")
            {
                string[] renameArgs = target.Split(new string[] { " to " }, StringSplitOptions.RemoveEmptyEntries);
                if (renameArgs.Length != 2)
                {
                    _ = Controller.SendChatMessage(replyTo, $"Please use the format: /{command} Old Name to New Name\nExample: /{command} iPad User to John Doe");
                    return;
                }

                target = renameArgs[0];
                newName = renameArgs[1];
            }

            // Handle special "/speaker off" command
            if ((command == "speaker") && (target == "off"))
            {
                SetSpeaker(null, from);
                return;
            }

            if ((command == "speak") || (command == "say"))
            {
                if (Controller.SendChatMessage(Controller.SpecialParticipant.everyoneInMeeting, target))
                {
                    Sound.Speak(target);
                }

                return;
            }

            if (command == "play")
            {
                if (Controller.SendChatMessage(replyTo, $"Playing: {repr(target)}"))
                {
                    Sound.Play(target);
                }

                return;
            }

            if (command == "list")
            {
                string arg = target.ToLower();
                string response = null;

                if ((arg == "hand") || (arg == "hands"))
                {
                    response = $"Raised Hands: {(raisedHandParticipants.Count == 0 ? "None" : raisedHandParticipants.ToDelimString())}";
                }
                else if ((arg == "talking") || (arg == "talkers") || (arg == "talk") || (arg == "activeaudio") || (arg == "audio"))
                {
                    response = $"Talkers: {(activeAudioParticipants.Count == 0 ? "None" : activeAudioParticipants.ToDelimString())}";
                }
                else
                {
                    response = $"Usage: /list hands|talkers";
                }

                _ = Controller.SendChatMessage(replyTo, response);

                return;
            }

            if ((command == "track") || (command == "tracking"))
            {
                string arg = target.ToLower();
                string response = null;

                if ((arg == "off") || (arg == "stop") || (arg == "disable"))
                {
                    lock (trackers)
                    {
                        if (trackers.ContainsKey(from.UserId))
                        {
                            trackers.Remove(from.UserId);
                            response = $"Tracking disabled";
                        }
                        else
                        {
                            response = $"Tracking is not enabled";
                        }
                    }
                }
                else if ((arg == "hand") || (arg == "hands"))
                {
                    _ = Controller.SendChatMessage(replyTo, EnableTracking(from.UserId, "hands"));
                    response = $"Raised Hands: {(raisedHandParticipants.Count == 0 ? "None" : raisedHandParticipants.ToDelimString())}";
                }
                else if ((arg == "talking") || (arg == "talkers") || (arg == "talk") || (arg == "activeaudio") || (arg == "audio"))
                {
                    _ = Controller.SendChatMessage(replyTo, EnableTracking(from.UserId, "talkers"));
                    response = $"Talkers: {(activeAudioParticipants.Count == 0 ? "None" : activeAudioParticipants.ToDelimString())}";
                }
                else if ((arg == "chat") || (arg == "chats"))
                {
                    response = EnableTracking(from.UserId, "chat");
                }
                else
                {
                    response = $"Usage: /track hands|talkers|chat|off";
                }

                _ = Controller.SendChatMessage(replyTo, response);

                return;
            }

            if (command == "who")
            {
                _ = Controller.SendChatMessage(replyTo, $"{repr(target)} is {GetUserLevel(target)}.");
                return;
            }

            // All of the following commands require a target participant

            // If the sender refers to themselves as "me", resolve this to their actual participant name
            Controller.Participant targetParticipant = null;
            if (target.ToLower() == "me")
            {
                targetParticipant = from;
            }
            else
            {
                try
                {
                    if (target.StartsWith("#"))
                    {
                        if (uint.TryParse(target.Substring(1), out uint id))
                        {
                            Controller.GetParticipantById(id, out targetParticipant, false);
                        }
                    }

                    if (targetParticipant == null)
                    {
                        targetParticipant = Controller.GetParticipantByName(target);
                    }
                }
                catch (ArgumentException)
                {
                    // TBD: Return userId or some other unique info?
                    _ = Controller.SendChatMessage(replyTo, $"Sorry, there is more than one participant here named {repr(target)}. I'm not sure which one you mean...");
                    return;
                }
            }

            if (targetParticipant == null)
            {
                // TBD: Try regex/partial match, returning results?
                _ = Controller.SendChatMessage(replyTo, $"Sorry, I don't see anyone named here named {repr(target)}. Remember, Case Matters!");
                return;
            }

            // All of the following require a participant target

            // Make sure I'm not the target :p
            if (targetParticipant.IsMe)
            {
                _ = Controller.SendChatMessage(replyTo, "U Can't Touch This\n* MC Hammer Music *\nhttps://youtu.be/otCpCn0l4Wo");
                return;
            }

            // Do rename if requested

            // TBD: Can you rename someone in the waiting room using the SDK?
            if (newName != null)
            {
                if (targetParticipant.Name == from.Name)
                {
                    _ = Controller.SendChatMessage(replyTo, "Why don't you just rename yourself?");
                    return;
                }

                var success = Controller.RenameParticipant(targetParticipant, newName);
                _ = Controller.SendChatMessage(replyTo, $"{(success ? "Successfully renamed" : "Failed to rename")} {repr(targetParticipant.Name)} to {repr(newName)}");

                return;
            }

            if (command == "admit")
            {
                if (targetParticipant.Status != Controller.ParticipantStatus.Waiting)
                {
                    _ = Controller.SendChatMessage(replyTo, $"Sorry, {repr(targetParticipant.Name)} is not in the waiting room");
                }
                else
                {
                    var success = Controller.AdmitParticipant(targetParticipant);
                    _ = Controller.SendChatMessage(replyTo, $"{(success ? "Successfully admitted" : "Failed to admit")} {repr(targetParticipant.Name)}");
                }

                return;
            }

            // Commands after here require the participant to be attending
            if (targetParticipant.Status != Controller.ParticipantStatus.Attending)
            {
                _ = Controller.SendChatMessage(replyTo, $"Sorry, {repr(targetParticipant.Name)} is not attending");
                return;
            }

            if ((command == "cohost") || (command == "promote"))
            {
                if (targetParticipant.IsHost || targetParticipant.IsCoHost)
                {
                    _ = Controller.SendChatMessage(replyTo, $"Sorry, {repr(targetParticipant.Name)} is already Host or Co-Host so cannot be promoted");
                    return;
                }
                else if (!targetParticipant.IsVideoOn)
                {
                    _ = Controller.SendChatMessage(replyTo, $"Sorry, I'm not allowed to Co-Host {repr(targetParticipant.Name)} because their video is off");
                    return;
                }

                var success = Controller.PromoteParticipant(targetParticipant, Controller.ParticipantRole.CoHost);
                _ = Controller.SendChatMessage(replyTo, $"{(success ? "Successfully promoted" : "Failed to promote")} {repr(targetParticipant.Name)}");

                return;
            }

            if (command == "demote")
            {
                if (!targetParticipant.IsCoHost)
                {
                    _ = Controller.SendChatMessage(replyTo, $"Sorry, {repr(targetParticipant.Name)} isn't Co-Host so they cannot be demoted");
                    return;
                }

                // Tag this as a bad user so we don't automatically cohost again
                BadUsers[targetParticipant.UserId] = targetParticipant.Name;

                _ = Controller.SendChatMessage(replyTo, $"{(Controller.DemoteParticipant(targetParticipant) ? "Successfully demoted" : "Failed to demote")} {repr(targetParticipant.Name)}");
                return;
            }

            if (command == "mute")
            {
                // TBD: Add /force option so that they cannot unmute

                _ = Controller.SendChatMessage(replyTo, $"{(Controller.MuteParticipant(targetParticipant) ? "Successfully muted" : "Failed to mute")} {repr(targetParticipant.Name)}");
                return;
            }

            if (command == "unmute")
            {
                _ = Controller.SendChatMessage(replyTo, $"{(Controller.UnmuteParticipant(targetParticipant) ? "Successfully unmuted" : "Failed to unmute")} {repr(targetParticipant.Name)}");
                return;
            }

            if ((command == "expel") || (command == "kick"))
            {
                _ = Controller.SendChatMessage(replyTo, $"{(Controller.ExpelParticipant(targetParticipant) ? "Successfully expelled" : "Failed to expel")} {repr(targetParticipant.Name)}");

                // Tag this as a bad user so we don't automatically re-admit
                BadUsers[targetParticipant.UserId] = targetParticipant.Name;

                return;
            }

            if ((command == "wait") || (command == "putwr") || (command == "waitroom") || (command == "waitingroom") || (command == "putinwait") || (command == "putinwaiting") || (command == "putinwaitingroom"))
            {
                _ = Controller.SendChatMessage(replyTo, $"{(Controller.PutParticipantInWaitingRoom(targetParticipant) ? "Successfully put" : "Failed to put")} {repr(targetParticipant.Name)} in waiting room");

                // Tag this as a bad user so we don't automatically re-admit
                BadUsers[targetParticipant.UserId] = targetParticipant.Name;

                return;
            }

            if (command == "speaker")
            {
                SetSpeaker(targetParticipant, from);
                return;
            }

            _ = Controller.SendChatMessage(replyTo, $"Sorry, I don't know the command {command}");
        }

        private void Controller_OnChatMessageReceive(object sender, Controller.OnChatMessageReceiveArgs e)
        {
            var to = e.to;
            var from = e.from;
            var text = e.text;
            var isToEveryone = e.isToEveryone;

            // TBD: All of this parsing is really messy. It could use a re-write!

            // If the message is from the bot or we're not configured to process chat messages, then bail
            if (e.from.IsMe || ((cfg.BotAutomationFlags & BotAutomationFlag.ProcessChat) == 0))
            {
                return;
            }

            // Special case when only the bot and one attendee are present
            // TBD: Folks in the waiting room can throw this off. Really should count # of actual attendees
            var onlyTwoAttendees = Controller.participants.Count == 2;

            Controller.Participant replyTo = null;
            if (isToEveryone)
            {
                // If there are only two people in the meeting, isPrivate=true and we can assume they are talking to the bot.
                //   If there is more than one person in the meeting, isPrivate=false and we check for the bot's name so we can be sure they are talking to it.
                var withoutMyName = FastRegex.Replace(text, @"\b" + cfg.MyParticipantName + @"\b", string.Empty, RegexOptions.IgnoreCase);
                if ((withoutMyName == text) && (!onlyTwoAttendees))
                {
                    return;
                }

                // My name is in it!  Treat it like a private message to me (sans my name), but reply to everyone in the meeting
                text = withoutMyName;
                replyTo = Controller.SpecialParticipant.everyoneInMeeting;
            }
            else
            {
                replyTo = from;
                NotifyTrackers("chat", $"(from {from}) {text}", from.UserId);

                // Currently the SDK only sends events for chats sent to everyone or sent to me, but it's possible that may change in the future.
                //   We do a sanity check here to make sure we don't respond to messages sent to other participants
                if (!to.IsMe)
                {
                    return;
                }
            }

            // Break the text up into multiple lines
            // TBD: Do some kind of rate throttling in case someone is spamming me?
            var lines = text.GetLines();
            foreach (var line in lines)
            {
                HandleChatMessageToMe(from, replyTo, line, isToEveryone, onlyTwoAttendees);
            }
        }

        public void Stop()
        {
            lock (_lock_eh)
            {
                if (shouldExit)
                {
                    hostApp.Log(LogType.WRN, "[UsherBot] Stop(): Ignoring duplicate request");
                    return;
                }

                if (chatBots != null)
                {
                    foreach (var chatBot in chatBots)
                    {
                        try
                        {
                            chatBot.Stop();
                        }
                        catch (Exception ex)
                        {
                            hostApp.Log(LogType.ERR, $"Exception stopping chatBot: {ex}");
                        }
                    }
                }

                LeaveMeeting();
            }

            // TBD: This is a total hackfest. Need to cleanly end the meeting...
            Controller.SleepWithDoEvents(5000);
            shouldExit = true;

            try
            {
                Controller.Stop();
            }
            catch (Exception ex)
            {
                hostApp.Log(LogType.ERR, $"Exception stopping controller: {ex}");
            }
        }

        public void SettingsChanged(object sender, EventArgs e)
        {
            LoadSettings();
        }

        private void LoadSettings()
        {
            cfg = DeserializeJson<BotConfigurationSettings>(hostApp.GetSettingsAsJSON());
            ExpandDictionaryPipes(cfg.BroadcastCommands);

            badNameRegex = null;
            if (cfg.BadNameRegex != null)
            {
                try
                {
                    badNameRegex = new Regex(cfg.BadNameRegex, RegexOptions.Compiled);
                }
                catch (Exception ex)
                {
                    hostApp.Log(LogType.ERR, $"Failed to compile BadNameRegex {repr(cfg.BadNameRegex)}: {ex}");
                }
            }
        }
    }
}
