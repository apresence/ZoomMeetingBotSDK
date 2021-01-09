namespace ZoomMeetingBotSDK
{
    using System;
    using System.Collections.Generic;
    using System.IO;
    using System.Linq;
    using System.Reflection;
    using System.Text.RegularExpressions;
    using System.Threading;
    using static Utils;
    using ChatBot;
    using ControlBot;

    public class UsherBot : IControlBot
    {
#pragma warning disable SA1401 // Fields should be private
        public static volatile bool ShouldExit = false;
#pragma warning restore SA1401 // Fields should be private

        private static IHostApp hostApp;
        private static Controller controller;

        private static readonly Dictionary<string, bool> GoodUsers = new Dictionary<string, bool>();
        private static readonly object _lock_eh = new object();

        private static DateTime dtLastWaitingRoomAnnouncement = DateTime.MinValue;

        private static DateTime dtNextAdmission = DateTime.MinValue;
        private static DateTime dtLastGoodUserMod = DateTime.MinValue;

        private static System.Threading.Timer tmrIdle = null;

        /// <summary>
        /// Used to record the last time a broadcast message was sent in order to prevent a specific broadcast message from being requested & sent in rapid succession.
        /// </summary>
        private static Dictionary<string, DateTime> BroadcastSentTime = new Dictionary<string, DateTime>();

        /// <summary>
        /// Topic of the current meeting. Set with "/topic ..." command (available to admins only), sent to new participants as they join, and also retreived on-demand by "/topic".
        /// </summary>
        private static string Topic = null;

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
            All = 0b1111111111,
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
            public bool DebugLoggingEnabled = false;
            public bool IsPaused = false;
            public bool PromptOnStartup = false;
            public int UnknownParticipantThrottleSecs = 15;
            public int UnknownParticipantWaitSecs = 30;
            public string MyParticipantName = "ZoomBot";
            public BotAutomationFlag BotAutomationFlags = BotAutomationFlag.All;
            public string MeetingID = null;
            public int BroadcastCommandGuardTimeSecs = 300;
            public string WaitingRoomAnnouncementMessage = null;
            public int WaitingRoomAnnouncementDelaySecs = 60;
            //public string Screen = null;
        }

        public static BotConfigurationSettings cfg = new BotConfigurationSettings();

        private static Dictionary<string, string> oneTimeHiSequences = null;
        private static Dictionary<string, string> broadcastCommands = null;
        private static Dictionary<string, EmailCommandArgs> emailCommands = null;

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
                // In lockdown mode, don't automatically admit or cohost anybody
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
                hostApp.Log(LogType.INF, "Lockdown mode {0}", bNewState ? "on" : "off");
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
                return useDefault ? "The topic has not been set" : null;
            }

            return GetTodayTonight().UppercaseFirst() + "'s topic: " + Topic;
        }

        public static bool SendTopic(string recipient, bool useDefault = true)
        {
            var topic = GetTopic(useDefault);

            if (topic == null)
            {
                return false;
            }

            var response = OneTimeHi("morning", recipient);
            if (response != null)
            {
                response = FormatChatResponse(response, recipient) + " " + topic;
            }
            else
            {
                response = topic;
            }

            Controller.SendChatMessage(recipient, response);

            return true;
        }

        private static string CleanUserName(string s)
        {
            if (s == null)
            {
                return null;
            }
            // 1. Lower-case names for comparison purposes
            // 2. Remove periods (Chris M. -> Chris M)
            // 3. Replace runs of spaces with a single space
            // 4. Remove known suffixes: (Usher), (DL), (Chair), (Speaker)
            return Regex.Replace(Regex.Replace(s.ToLower().Replace(".", string.Empty), @"\s+", " "), @"\s*\((?:Usher|DL|Chair|Speaker)\)\s*$", string.Empty, RegexOptions.IgnoreCase).Trim();
        }

        //static private string sLastChatData = "";
        private static void DoChatActions()
        {
            if ((cfg.BotAutomationFlags & BotAutomationFlag.ProcessChat) == 0)
            {
                return;
            }

            //ZoomMeetingBotSDK.SendQueuedChatMessages();
            _ = Controller.UpdateChat();
            Controller.SendQueuedChatMessages();
        }

        private static readonly HashSet<string> HsParticipantMessages = new HashSet<string>();

        private static string FirstParticipantGreeted = null;
        private static void DoParticipantActions()
        {
            if ((cfg.BotAutomationFlags & BotAutomationFlag.ProcessParticipants) == 0)
            {
                return;
            }

            _ = Controller.UpdateParticipants();

            if (Controller.me != null)
            {
                // If I've got my own participant object, do any self-automation needed

                if (((cfg.BotAutomationFlags & BotAutomationFlag.ReclaimHost) != 0) && (Controller.me.role != Controller.ParticipantRole.Host))
                {
                    // TBD: Throttle ReclaimHost attempts?
                    if (Controller.me.role == Controller.ParticipantRole.CoHost)
                    {
                        hostApp.Log(LogType.WRN, "BOT I'm Co-Host instead of Host; Trying to reclaim host");
                    }
                    else if (Controller.me.role == Controller.ParticipantRole.None)
                    {
                        hostApp.Log(LogType.WRN, "BOT I'm not Host or Co-Host; Trying to reclaim host");
                    }
                    Controller.ReclaimHost();
                }

                if (((cfg.BotAutomationFlags & BotAutomationFlag.RenameMyself) != 0) && (Controller.me.name != cfg.MyParticipantName))
                {
                    // Rename myself.  Event handler will type in the name when the dialog pops up
                    hostApp.Log(LogType.INF, "BOT Renaming myself from {0} to {1}", repr(Controller.me.name), repr(cfg.MyParticipantName));
                    Controller.RenameParticipant(Controller.me, cfg.MyParticipantName);
                }

                if (((cfg.BotAutomationFlags & BotAutomationFlag.UnmuteMyself) != 0) && (Controller.me.audioStatus == Controller.ParticipantAudioStatus.Muted))
                {
                    // Unmute myself
                    hostApp.Log(LogType.INF, "BOT Unmuting myself");
                    Controller.UnmuteParticipant(Controller.me);
                }

                Controller.UpdateMeetingOptions();
            }

            bool bWaiting = false;
            DateTime dtNow = DateTime.UtcNow;
            foreach (Controller.Participant p in Controller.participants.Values)
            {
                // Skip over my own participant record; We handled that earlier
                if (p.isMe)
                {
                    continue;
                }

                string sCleanName = CleanUserName(p.name);
                bool bAdmit = false;
                bool bAdmitKnown = (cfg.BotAutomationFlags & BotAutomationFlag.AdmitKnown) != 0;
                bool bAdmitOthers = (cfg.BotAutomationFlags & BotAutomationFlag.AdmitOthers) != 0;

                if (p.status == Controller.ParticipantStatus.Waiting)
                {
                    bWaiting = true;

                    if (!(bAdmitKnown || bAdmitOthers))
                    {
                        continue; // Nothing to do
                    }

                    if (GoodUsers.ContainsKey(sCleanName))
                    {
                        if (bAdmitKnown)
                        {
                            hostApp.Log(LogType.INF, "BOT Admitting {0} : KNOWN", repr(p.name));
                            if (Controller.AdmitParticipant(p))
                            {
                                //SendTopic(p.name, false);
                            }
                        }

                        continue;
                    }

                    // Unknown user
                    DateTime dtWhenToAdmit = p.dtWaiting.AddSeconds(cfg.UnknownParticipantWaitSecs);
                    dtWhenToAdmit = dtWhenToAdmit > dtNextAdmission ? dtWhenToAdmit : dtNextAdmission;
                    bAdmit = dtWhenToAdmit >= dtNow;

                    if (!bAdmitOthers)
                    {
                        continue;
                    }

                    string sMsg = string.Format("BOT Admit {0} : Unknown participant waiting room time reached", p.name);
                    if (bAdmit)
                    {
                        sMsg += " : Admitting";
                    }

                    // Make sure we don't display the message more than once
                    if (!HsParticipantMessages.Contains(sMsg))
                    {
                        hostApp.Log(LogType.INF, sMsg);
                        HsParticipantMessages.Add(sMsg);
                    }

                    if (bAdmit && Controller.AdmitParticipant(p))
                    {
                        HsParticipantMessages.Remove(sMsg); // After we admit the user, remove the message
                        dtNextAdmission = dtNow.AddSeconds(cfg.UnknownParticipantThrottleSecs);

                        /*
                        // Participant was successfully admitted.  We want to send them the topic if one is set, but we can't do that
                        //   while they are in the waiting room (DMs cannot be sent to waiting room participants, only broadcast messages),
                        //   so queue up the message for later after they are admitted.
                        SendTopic(p.name, false);
                        */
                    }

                    continue;
                }

                if (p.status == Controller.ParticipantStatus.Attending)
                {
                    if (((cfg.BotAutomationFlags & BotAutomationFlag.CoHostKnown) != 0) && (p.role == Controller.ParticipantRole.None) && (Controller.me.role == Controller.ParticipantRole.Host))
                    {
                        // If I'm host, and this user is not co-host, check if they should be
                        if (GoodUsers.TryGetValue(sCleanName, out bool bCoHost) && bCoHost)
                        {
                            // Yep, they should be, so do the promotion
                            hostApp.Log(LogType.INF, "BOT Promoting {0} to Co-host", repr(p.name));
                            Controller.PromoteParticipant(p);
                        }
                    }

                    continue;
                }
            }

            if (bWaiting)
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

                // At least one person is in the waiting room.  If we're configured to make annoucements to
                //   them, then do so now

                dtNow = DateTime.UtcNow;
                if (dtNow >= dtLastWaitingRoomAnnouncement.AddSeconds(cfg.WaitingRoomAnnouncementDelaySecs))
                {
                    Controller.SendChatMessage(Controller.SpecialRecipient.EveryoneInWaitingRoom, waitMsg);
                    dtLastWaitingRoomAnnouncement = dtNow;
                }
            }

            // Greet the first person to join the meeting, but only if we started Zoom
            if ((!Controller.ZoomAlreadyRunning) && (FirstParticipantGreeted == null))
            {
                var plist = Controller.participants.ToList();

                // Looking for a participant that is not me, using computer audio, audio is connected, and is a known good user
                var idx = plist.FindIndex(x => (
                    (!x.Value.isMe) &&
                    (x.Value.device == Controller.ParticipantAudioDevice.Computer) &&
                    (x.Value.audioStatus != Controller.ParticipantAudioStatus.Disconnected) &&
                    GoodUsers.ContainsKey(CleanUserName(x.Value.name))
                ));
                if (idx != -1)
                {
                    FirstParticipantGreeted = plist[idx].Value.name;
                    var msg = FormatChatResponse(OneTimeHi("morning", FirstParticipantGreeted), FirstParticipantGreeted);

                    Sound.Play("bootup");
                    Thread.Sleep(3000);
                    Sound.Speak(cfg.MyParticipantName + " online.");
                    Controller.SendChatMessage(Controller.SpecialRecipient.EveryoneInMeeting, true, msg);
                }
            }
        }

        private static int nTimerIterationID = 0;

        private static void TimerIdleHandler(object o)
        {
            if (ShouldExit)
            {
                return;
            }

            Interlocked.Increment(ref nTimerIterationID);

            if (!Monitor.TryEnter(_lock_eh))
            {
                hostApp.Log(LogType.WRN, "TimerIdleHandler {0:X4} - Busy; Will try again later", nTimerIterationID);
                return;
            }

            try
            {
                //hostApp.Log(LogType.DBG, "TimerIdleHandler {0:X4} - Enter");

                LoadGoodUsers();
                ReadRemoteCommands();

                // Zoom is really bad about moving/resizing it's windows, so keep it in check
                Controller.LayoutWindows();

                if (cfg.IsPaused)
                {
                    return;
                }

                //hostApp.Log(LogType.DBG, "TimerIdleHandler {0:X4} - DoParticipantActions", nTimerIterationID);
                DoParticipantActions();

                //hostApp.Log(LogType.DBG, "TimerIdleHandler {0:X4} - DoChatActions", nTimerIterationID);
                DoChatActions();
            }
            catch (Controller.ZoomClosedException ex)
            {
                hostApp.Log(LogType.INF, ex.ToString());
                ShouldExit = true;
            }
            catch (Exception ex)
            {
                hostApp.Log(LogType.ERR, "TimerIdleHandler {0:X4} - Unhandled Exception: {1}", nTimerIterationID, ex.ToString());
            }
            finally
            {
                //hostApp.Log(LogType.DBG, "TimerIdleHandler {0:X4} - Exit", nTimerIterationID);
                Monitor.Exit(_lock_eh);
            }
        }

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
                    else if (line == "exit")
                    {
                        hostApp.Log(LogType.INF, "Received {0} command", line);
                        Controller.LeaveMeeting(false);
                        ShouldExit = true;
                    }
                    else if (line == "kill")
                    {
                        hostApp.Log(LogType.INF, "Received {0} command", line);
                        Controller.LeaveMeeting(true);
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

        private static void LoadGoodUsers()
        {
            string sPath = @"good_users.txt";

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

            hostApp.Log(LogType.INF, "(Re-)loading GoodUsers");

            GoodUsers.Clear();
            using (StreamReader sr = File.OpenText(sPath))
            {
                string line = null;
                bool bAdmin = false;
                while ((line = sr.ReadLine()) != null)
                {
                    line = line.Trim();
                    if (line.Length == 0)
                    {
                        continue;
                    }

                    // Admin lines end in "^"
                    bAdmin = line.EndsWith("^");
                    if (bAdmin)
                    {
                        line = line.TrimEnd('^');
                    }

                    // Allow alises, delimited by "|"
                    string[] names = line.Split('|');
                    foreach (string name in names)
                    {
                        string sCleanName = CleanUserName(name);
                        if (sCleanName.Length == 0)
                        {
                            continue;
                        }

                        // TBD: Don't allow generic names -- aka, don't allow names without at least one space in them?
                        if (GoodUsers.ContainsKey(sCleanName))
                        {
                            // Duplicate entry; Honor admin flag
                            GoodUsers[sCleanName] = GoodUsers[sCleanName] | bAdmin;
                        }
                        else
                        {
                            GoodUsers.Add(sCleanName, bAdmin);
                        }
                    }
                }
            }
        }

        private static void OnMeetingOptionStateChange(object sender, Controller.MeetingOptionStateChangeEventArgs e)
        {
            hostApp.Log(LogType.INF, "Meeting option {0} changed to {1}", repr(e.optionName), e.newState.ToString());
        }

        private static void OnParticipantAttendanceStatusChange(object sender, Controller.ParticipantEventArgs e)
        {
            Controller.Participant p = e.participant;
            hostApp.Log(LogType.INF, "Participant {0} status {1}", repr(p.name), p.status.ToString());

            // TBD: Could immediately admit recognized attendees
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

        private static string FormatChatResponse(string text, string to)
        {
            return string.Format(text, GetFirstName(to), GetDayTime());
        }

        private static readonly Dictionary<string, string> DicOneTimeHis = new Dictionary<string, string>();

        private static string OneTimeHi(string text, string to)
        {
            string response = null;

            if (oneTimeHiSequences == null)
            {
                return response;
            }

            // Do one-time "hi" only once
            if (DicOneTimeHis.ContainsKey(to))
            {
                return null;
            }

            // Try to give a specific response
            foreach (var word in text.GetWordsInSentence())
            {
                if (oneTimeHiSequences.TryGetValue(word.ToLower(), out response))
                {
                    break;
                }
            }

            if (response != null)
            {
                DicOneTimeHis.Add(to, response); // TBD: Really only need key hash
            }

            return response;
        }

        private static void SetSpeaker(Controller.Participant p, string from)
        {
            Controller.SendChatMessage(from, "Speaker mode is not yet implemented");

            /*
            if (p == null)
            {
                if (ZoomMeetingBotSDK.GetMeetingOption(ZoomMeetingBotSDK.MeetingOption.AllowParticipantsToUnmuteThemselves) == System.Windows.Automation.ToggleState.On)
                {
                    if (from != null)
                    {
                        ZoomMeetingBotSDK.SendChatMessage(from, "Speaker mode is already off");
                    }

                    return;
                }

                ZoomMeetingBotSDK.SetMeetingOption(ZoomMeetingBotSDK.MeetingOption.AllowParticipantsToUnmuteThemselves, System.Windows.Automation.ToggleState.On);
                if (from != null)
                {
                    ZoomMeetingBotSDK.SendChatMessage(from, "Speaker mode turned off");
                }

                return;
            }

            if (from != null)
            {
                ZoomMeetingBotSDK.SendChatMessage(from, $"Setting speaker to {p.name}");
            }

            ZoomMeetingBotSDK.SetMeetingOption(ZoomMeetingBotSDK.MeetingOption.MuteParticipantsUponEntry, System.Windows.Automation.ToggleState.On);
            // - Set by MuteAll dialog - ZoomMeetingBotSDK.SetMeetingOption(ZoomMeetingBotSDK.MeetingOption.AllowParticipantsToUnmuteThemselves, System.Windows.Automation.ToggleState.Off);

            /-*
            _ = ZoomMeetingBotSDK.MuteAll(false);

            // MuteAll does not mute Host or Co-Host participants, so do that now
            foreach (ZoomMeetingBotSDK.Participant participant in ZoomMeetingBotSDK.participants.Values)
            {
                // Skip past folks who are not Host or Co-Host
                if (participant.role == ZoomMeetingBotSDK.ParticipantRole.None)
                {
                    continue;
                }

                // Skip past folks that are not unmuted
                if (participant.audioStatus != ZoomMeetingBotSDK.ParticipantAudioStatus.Unmuted)
                {
                    continue;
                }

                ZoomMeetingBotSDK.MuteParticipant(p);
            }

            ZoomMeetingBotSDK.UnmuteParticipant(p);
            *-/

            // Mute everyone who is not muted (unless they are host or co-host)
            foreach (ZoomMeetingBotSDK.Participant participant in ZoomMeetingBotSDK.participants.Values)
            {
                if (participant.name == p.name)
                {
                    // This is the speaker, make sure he/she is unmuted
                    if (participant.audioStatus == ZoomMeetingBotSDK.ParticipantAudioStatus.Muted)
                    {
                        ZoomMeetingBotSDK.UnmuteParticipant(participant);
                    }

                    continue;
                }

                // Skip past folks who are Host or Co-Host
                if (participant.role != ZoomMeetingBotSDK.ParticipantRole.None)
                {
                    continue;
                }

                // Mute anyone who is off mute
                if (participant.audioStatus == ZoomMeetingBotSDK.ParticipantAudioStatus.Unmuted)
                {
                    ZoomMeetingBotSDK.MuteParticipant(p);
                }
            }
            */
        }

        private static void OnChatMessageReceive(object source, Controller.ChatEventArgs e)
        {
            hostApp.Log(LogType.INF, "New message from {0} to {1}: {2}", repr(e.from), repr(e.to), repr(e.text));

            string sTo = e.to;
            string sFrom = e.from;
            string sMsg = e.text.Trim();
            string sReplyTo = sFrom;

            if (!GoodUsers.TryGetValue(CleanUserName(sFrom), out bool bAdmin))
            {
                bAdmin = false;
            }

            // Ignore messages from me
            if (sFrom.ToLower() == "me")
            {
                return;
            }

            if (!e.isPrivate)
            {
                // Message is to everyone (public), bail if my name is not in it
                var withoutMyName = Regex.Replace(sMsg, @"\b" + cfg.MyParticipantName + @"\b", string.Empty, RegexOptions.IgnoreCase);

                // If strings are the same, it's not to me
                if (withoutMyName == sMsg)
                {
                    return;
                }

                sMsg = withoutMyName;

                // My name is in it, so reply to everyone
                sReplyTo = Controller.SpecialRecipient.EveryoneInMeeting;
            }
            else if (sTo.ToLower() != "me")
            {
                // Ignore it if it's not to me
                return;
            }

            // All commands start with "/"; Treat everything else as small talk
            if (!sMsg.StartsWith("/"))
            {
                // Try to get the best response possible; Fall back on something random if all else fails
                //   TBD: Could make sure we don't say the same thing twice...

                var isToEveryone = Controller.SpecialRecipient.IsEveryone(sReplyTo);

                // If the bot is addressed publically or if there are only two people in the meeting, then reply with TTS
                // TBD: Should be attending count, not participant count.  Some could be in the waiting room
                var speak = isToEveryone || (Controller.participants.Count == 2);

                // We start with a one-time hi.  Various bots may be in different time zones and the
                //   good morning/afternoon/evening throws things off
                var response = OneTimeHi(sMsg, sFrom);

                // Handle canned responses based on broadcast keywords.  TBD: Move this into a bot
                if (broadcastCommands != null)
                {
                    foreach (var broadcastCommand in broadcastCommands)
                    {
                        if (FastRegex.IsMatch($"\\b${broadcastCommand.Key}\\b", sMsg, RegexOptions.IgnoreCase))
                        {
                            response = broadcastCommand.Value;

                            // Don't want to speak broadcast messages
                            speak = false;
                        }
                    }
                }

                // Handle topic request
                if (response == null)
                {
                    if (FastRegex.IsMatch($"\\b(topic|reading)\\b", sMsg, RegexOptions.IgnoreCase))
                    {
                        SendTopic(sReplyTo, true);
                        return;
                    }
                }

                // We did the one time hi, now feed the text to the chat bots!
                if ((response == null) && (chatBots != null))
                {
                    // We'll try each bot in order by intelligence level until one of them works
                    foreach (var chatBot in chatBots)
                    {
                        string failureMsg = null;
                        try
                        {
                            response = chatBot.Converse(sMsg, sFrom);
                            if (response == null)
                            {
                                failureMsg = "Response is null";
                            }
                        }
                        catch (Exception ex)
                        {
                            failureMsg = "Exception occured: " + ex.ToString();
                            response = null;
                        }

                        if (response != null)
                        {
                            break;
                        }

                        hostApp.Log(LogType.WRN, $"ChatBot converse with {repr(chatBot.GetChatBotInfo().Name)} failed: {repr(failureMsg)}");
                    }
                }

                if (response == null)
                {
                    hostApp.Log(LogType.ERR, "No ChatBot was able to produce a response");
                }

                Controller.SendChatMessage(sReplyTo, speak, FormatChatResponse(response, sFrom));

                return;
            }

            // Non-priv retrival of topic
            if (sMsg == "/topic")
            {
                SendTopic(sReplyTo, true);
                return;
            }

            // Everything after here is a command.  Drop any commands not directly addressed to me
            if (sTo.ToLower() != "me")
            {
                return;
            }

            // Only allow admin users to run the following commands
            if (!bAdmin)
            {
                hostApp.Log(LogType.WRN, "Ignoring command {0} from non-admin {1}", repr(sMsg), repr(sFrom));
                return;
            }

            if (!Controller.participants.TryGetValue(sFrom, out Controller.Participant sender))
            {
                hostApp.Log(LogType.ERR, "Received command {0} from {1}, but I don't have a Participant class for them", repr(sMsg), repr(e.from));
                return;
            }

            string[] a = sMsg.Split(SpaceDelim, 2);

            string sCommand = a[0].ToLower().Substring(1);

            // All of the following commands require an argument
            string sTarget = (a.Length == 1) ? null : (a[1].Length == 0 ? null : a[1]);

            if ((broadcastCommands != null) && broadcastCommands.TryGetValue(sCommand, out string sBroadcastMsg))
            {
                DateTime dtNow = DateTime.UtcNow;

                if (BroadcastSentTime.TryGetValue(sCommand, out DateTime dtSentTime))
                {
                    int guardTime = cfg.BroadcastCommandGuardTimeSecs;

                    if (guardTime < 0)
                    {
                        Controller.SendChatMessage(sender.name, $"{sCommand}: This broadcast message was already sent.");
                        return;
                    }

                    if ((guardTime > 0) && (dtNow <= dtSentTime.AddSeconds(cfg.BroadcastCommandGuardTimeSecs)))
                    {
                        Controller.SendChatMessage(sender.name, $"{sCommand}: This broadcast message was already sent recently. Please try again later.");
                        return;
                    }
                }

                Controller.SendChatMessage(Controller.SpecialRecipient.EveryoneInMeeting, sBroadcastMsg);
                BroadcastSentTime[sCommand] = dtNow;

                return;
            }

            // Priv retrival or set of topic
            if (sCommand == "topic")
            {
                if (sTarget == null)
                {
                    SendTopic(sender.name, true);
                    return;
                }

                bool broadcast = false;
                string reply;

                string[] b = sTarget.Split(SpaceDelim, 2);

                string cmd = b[0].ToLower().TrimStart('/');

                if (cmd == "force")
                {
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
                else if (string.Compare(Topic, sTarget, true) == 0)
                {
                    reply = "The topic is already set to: " + sTarget;
                }
                else if (Topic == null)
                {
                    reply = "Topic set to: " + sTarget;
                    Topic = sTarget;
                    broadcast = true;
                }
                else
                {
                    reply = "Topic is already set; Use /topic force to change it";
                }

                Controller.SendChatMessage(sReplyTo, reply);

                if (broadcast)
                {
                    Controller.SendChatMessage(Controller.SpecialRecipient.EveryoneInMeeting, GetTopic());
                }

                return;
            }

            // All of the following commands require options
            if (sTarget == null)
            {
                return;
            }

            if (emailCommands != null)
            {
                if (emailCommands.TryGetValue(sCommand, out EmailCommandArgs emailCommandArgs))
                {
                    string[] args = sTarget.Trim().Split(SpaceDelim, 2);

                    string toAddress = args[0];
                    string subject = emailCommandArgs.Subject;
                    string body = emailCommandArgs.Body;

                    if (subject.Contains("{0}") || body.Contains("{0}"))
                    {
                        if (args.Length <= 1)
                        {
                            Controller.SendChatMessage(sender.name, $"Error: The format of the command is incorrect; Correct example: /{sCommand} {emailCommandArgs.ArgsExample}");
                            return;
                        }

                        string emailArg = args[1].Trim();
                        subject = subject.Replace("{0}", emailArg);
                        body = body.Replace("{0}", emailArg);
                    }

                    if (SendEmail(subject, body, toAddress))
                    {
                        Controller.SendChatMessage(sender.name, $"{sCommand}: Successfully sent email to {toAddress}");
                    }
                    else
                    {
                        Controller.SendChatMessage(sender.name, $"{sCommand}: Failed to send email to {toAddress}");
                    }


                    return;
                }
            }

            if ((sCommand == "citadel") || (sCommand == "lockdown") || (sCommand == "passive"))
            {
                string sNewMode = sTarget.ToLower().Trim();
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
                    Controller.SendChatMessage(sender.name, "Sorry, the {0} command requires either on or off as a parameter", repr(sCommand));
                    return;
                }

                if (SetMode(sCommand, bNewMode))
                {
                    Controller.SendChatMessage(sender.name, "{0} mode has been changed to {1}", GetFirstName(sCommand), sNewMode);
                }
                else
                {
                    Controller.SendChatMessage(sender.name, "{0} mode is already {1}", GetFirstName(sCommand), sNewMode);
                }
                return;
            }

            if (sCommand == "waitmsg")
            {
                var sWaitMsg = sTarget.Trim();

                if ((sWaitMsg.Length == 0) || (sWaitMsg.ToLower() == "off"))
                {
                    if ((cfg.WaitingRoomAnnouncementMessage != null) && (cfg.WaitingRoomAnnouncementMessage.Length > 0))
                    {
                        cfg.WaitingRoomAnnouncementMessage = null;
                        Controller.SendChatMessage(sender.name, "Waiting room message has been turned off");
                    }
                    else
                    {
                        Controller.SendChatMessage(sender.name, "Waiting room message is already off");
                    }
                }
                else if (sWaitMsg == cfg.WaitingRoomAnnouncementMessage)
                {
                    Controller.SendChatMessage(sender.name, "Waiting room message is already set to:\n{0}", sTarget);
                }
                else
                {
                    cfg.WaitingRoomAnnouncementMessage = sTarget.Trim();
                    Controller.SendChatMessage(sender.name, "Waiting room message has set to:\n{0}", sTarget);
                }
                return;
            }

            // Pre-processing for rename action
            string newName = null;
            if (sCommand == "rename")
            {
                string[] renameArgs = sTarget.Split(new string[] { " to " }, StringSplitOptions.RemoveEmptyEntries);
                if (renameArgs.Length != 2)
                {
                    Controller.SendChatMessage(sender.name, "Please use the format: /{0} Old Name to New Name", sCommand);
                    Controller.SendChatMessage(sender.name, "Example: /{0} iPad User to John Doe", sCommand);
                    return;
                }
                sTarget = renameArgs[0];
                newName = renameArgs[1];
            }

            // Handle special "/speaker off" command
            if ((sCommand == "speaker") && (sTarget == "off"))
            {
                SetSpeaker(null, e.from);
                return;
            }

            if ((sCommand == "speak") || (sCommand == "say"))
            {
                Controller.SendChatMessage(Controller.SpecialRecipient.EveryoneInMeeting, sCommand == "speak", sTarget);

                return;
            }

            if (sCommand == "play")
            {
                Controller.SendChatMessage(sender.name, "Playing: {0}", repr(sTarget));
                Sound.Play(sTarget);
                return;
            }

            // If the sender refers to themselves as "me", resolve this to their actual participant name
            if (sTarget.ToLower() == "me")
            {
                sTarget = e.from;
            }

            // All of the following require a participant target
            if (!Controller.participants.TryGetValue(sTarget, out Controller.Participant target))
            {
                Controller.SendChatMessage(sender.name, "Sorry, I don't see anyone named here named {0}. Remember, Case Matters!", repr(sTarget));
                return;
            }

            // Make sure I'm not the target :p
            if (target.isMe)
            {
                Controller.SendChatMessage(sender.name, "U Can't Touch This\n* MC Hammer Music *\nhttps://youtu.be/otCpCn0l4Wo");
                return;
            }

            // Do rename if requested
            if (newName != null)
            {
                if (target.name == sender.name)
                {
                    Controller.SendChatMessage(sender.name, "Why don't you just rename yourself?");
                    return;
                }

                Controller.SendChatMessage(sender.name, "Renaming {0} to {1}", repr(target.name), repr(newName));
                Controller.RenameParticipant(target, newName);
                return;
            }

            if (sCommand == "admit")
            {
                if (target.status != Controller.ParticipantStatus.Waiting)
                {
                    Controller.SendChatMessage(sender.name, "Sorry, {0} is not waiting", repr(target.name));
                }
                else
                {
                    Controller.SendChatMessage(sender.name, "Admitting {0}", repr(target.name));
                    if (Controller.AdmitParticipant(target))
                    {
                        // Participant was successfully admitted.  We want to send them the topic if one is set, but we can't do that
                        //   while they are in the waiting room (DMs cannot be sent to waiting room participants, only broadcast messages),
                        //   so queue up the message for later after they are admitted.
                        //SendTopic(target.name, false);
                    }
                }

                return;
            }

            // Commands after here require the participant to be attending
            if (target.status != Controller.ParticipantStatus.Attending)
            {
                Controller.SendChatMessage(sender.name, "Sorry, {0} is not attending", repr(target.name));
                return;
            }

            if ((sCommand == "cohost") || (sCommand == "promote"))
            {
                if (target.role != Controller.ParticipantRole.None)
                {
                    Controller.SendChatMessage(sender.name, "Sorry, {0} is already Host or Co-Host so cannot be promoted", repr(target.name));
                }
                else if (target.videoStatus != Controller.ParticipantVideoStatus.On)
                {
                    Controller.SendChatMessage(sender.name, "Co-Host name matched for {0}, but video is off", repr(target.name));
                    return;
                }
                else
                {
                    Controller.SendChatMessage(sender.name, "Promoting {0} to Co-Host", repr(target.name));
                    Controller.PromoteParticipant(target);
                }

                return;
            }

            if (sCommand == "demote")
            {
                if (target.role != Controller.ParticipantRole.CoHost)
                {
                    Controller.SendChatMessage(sender.name, "Sorry, {0} isn't Co-Host so cannot be demoted", repr(target.name));
                }
                else
                {
                    Controller.SendChatMessage(sender.name, "Demoting {0}", repr(target.name));
                    Controller.DemoteParticipant(target);
                }

                return;
            }

            if (sCommand == "mute")
            {
                Controller.SendChatMessage(sender.name, "Muting {0}", repr(target.name));
                Controller.MuteParticipant(target);
                return;
            }

            if (sCommand == "unmute")
            {
                Controller.SendChatMessage(sender.name, "Requesting {0} to Unmute", repr(target.name));
                Controller.UnmuteParticipant(target);
                return;
            }

            if (sCommand == "speaker")
            {
                SetSpeaker(target, e.from);
                return;
            }

            Controller.SendChatMessage(sender.name, "Sorry, I don't know the command {0}", repr(sCommand));
        }

        private static List<IChatBot> chatBots = null;

        /// <summary>
        /// Searches for ChatBot plugins under plugins\ChatBot\{BotName}\ZoomMeetingBotSDK.ChatBot.{BotName}.dll and tries to instantiate them,
        /// returning a list of ones that succeeded.  The list is ordered by intelligence level, with the most intelligent bot listed
        /// first.
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
                                hostApp.Log(LogType.DBG, $"Loaded {repr(chatBotInfo.Name)} chatbot with intelligence level {chatBotInfo.IntelligenceLevel}");
                                chatBot.Start();
                                bots.Add(new Tuple<int, IChatBot>(chatBotInfo.IntelligenceLevel, chatBot));
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

            return bots.OrderByDescending(o => o.Item1).Select(x => x.Item2).ToList();
        }

        /// <summary>
        /// Leaves the meeting, optionally ending meeting or passing off Host role to another participant.
        /// </summary>
        public static void LeaveMeeting(bool endForAll = false)
        {
            if (!endForAll)
            {
                if (Controller.me.role != Controller.ParticipantRole.Host)
                {
                    hostApp.Log(LogType.DBG, "BOT LeaveMeeting - I am not host");
                }
                else
                {
                    hostApp.Log(LogType.DBG, "BOT LeaveMeeting - I am host; Trying to find someone to pass it to");

                    Controller.Participant altHost = null;
                    foreach (Controller.Participant p in Controller.participants.Values)
                    {
                        if (p.role == Controller.ParticipantRole.CoHost)
                        {
                            altHost = p;
                            break;
                        }
                    }

                    if (altHost == null)
                    {
                        hostApp.Log(LogType.ERR, "BOT LeaveMeeting - Could not find an alternative host; Ending meeting");
                        endForAll = true;
                    }
                    else
                    {
                        try
                        {
                            hostApp.Log(LogType.INF, "BOT LeaveMeeting - Passing Host to {0}", repr(altHost.name));
                            Controller.PromoteParticipant(altHost, Controller.ParticipantRole.Host);
                            hostApp.Log(LogType.INF, "BOT LeaveMeeting - Passed Host to {0}", repr(altHost.name));
                        }
                        catch (Exception ex)
                        {
                            hostApp.Log(LogType.ERR, "BOT LeaveMeeting - Failed to pass Host to {0}; Ending meeting", repr(altHost.name));
                            endForAll = true;
                        }
                    }
                }
            }

            hostApp.Log(LogType.INF, "BOT LeaveMeeting - Leaving Meeting");
            Controller.LeaveMeeting(endForAll);
        }

        protected virtual void Dispose(bool disposing)
        {
            if (!disposing)
            {
                return;
            }

            tmrIdle.Dispose();
        }

        private static GmailSenderLib.GmailSender gmailSender = null;
        private static bool SendEmail(string subject, string body, string to)
        {
            try
            {
                if (gmailSender is null)
                {
                    gmailSender = new GmailSenderLib.GmailSender(System.Reflection.Assembly.GetCallingAssembly().GetName().Name);
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

            controller = new Controller();
            controller.Init(hostApp);
        }

        public void Start()
        {
            if ((cfg.BotAutomationFlags & BotAutomationFlag.Converse) != 0)
            {
                chatBots = GetChatBots();
            }

            Controller.ParticipantAttendanceStatusChange += OnParticipantAttendanceStatusChange;
            Controller.ChatMessageReceive += OnChatMessageReceive;
            Controller.MeetingOptionStateChange += OnMeetingOptionStateChange;
            Controller.Start();

            tmrIdle = new System.Threading.Timer(TimerIdleHandler, null, 0, 5000);

            return;
        }

        public void Stop()
        {
            ShouldExit = true;
        }

        public void SettingsChanged(object sender, EventArgs e)
        {
            LoadSettings();
        }

        private void LoadSettings()
        {
            var dic = hostApp.GetSettingsDic();
            DeserializeDictToObject<BotConfigurationSettings>(dic, cfg);

            oneTimeHiSequences = DynDicValueToStrDic(dic, "OneTimeHiSequences");
            ExpandDictionaryPipes(oneTimeHiSequences);

            broadcastCommands = DynDicValueToStrDic(dic, "BroadcastCommands");
            ExpandDictionaryPipes(broadcastCommands);

            if (dic.TryGetValue("EmailCommands", out dynamic valueEC))
            {
                Dictionary<string, dynamic> dicEC = valueEC;

                emailCommands = new Dictionary<string, EmailCommandArgs>();
                foreach (var kvp in dicEC)
                {
                    Dictionary<string, dynamic> dicValue = kvp.Value;
                    var emailCommandArgs = new EmailCommandArgs();
                    DeserializeDictToObject<EmailCommandArgs>(dicValue, emailCommandArgs);
                    emailCommands.Add(kvp.Key, emailCommandArgs);
                }
            }
        }
    }
}
