namespace ZoomMeetingBotSDK
{
    using System;
    using System.Collections.Generic;
    using System.Diagnostics.CodeAnalysis;
    using System.Drawing.Imaging;
    using System.IO;
    using System.Linq;
    using System.Reflection;
    using System.Text.RegularExpressions;
    using System.Threading;
    using global::ZoomMeetingBotSDK.Interop.ChatBot;
    using global::ZoomMeetingBotSDK.Interop.HostApp;
    using global::ZoomMeetingBotSDK.Utils;

    internal class UsherBot
    {
#pragma warning disable SA1401 // Fields should be private
        public static volatile bool ShouldExit = false;
#pragma warning restore SA1401 // Fields should be private

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

        public static bool SetMode(string sName, bool bNewState)
        {
            if (sName == "citadel")
            {
                // In Citadel mode, we do not automatically admit unknown participants
                bool bCitadelMode = (Global.cfg.BotAutomationFlags & Global.BotAutomationFlag.AdmitOthers) == 0;
                if (bCitadelMode == bNewState)
                {
                    return false;
                }

                if (bNewState)
                {
                    Global.cfg.BotAutomationFlags ^= Global.BotAutomationFlag.AdmitOthers;
                }
                else
                {
                    Global.cfg.BotAutomationFlags |= Global.BotAutomationFlag.AdmitOthers;
                }

                Global.hostApp.Log(LogType.INF, "Citadel mode {0}", bNewState ? "on" : "off");
                return true;
            }

            if (sName == "lockdown")
            {
                // In lockdown mode, don't automatically admit or cohost anybody
                var botLockdownFlags = Global.BotAutomationFlag.AdmitOthers | Global.BotAutomationFlag.AdmitKnown | Global.BotAutomationFlag.CoHostKnown;
                bool bLockdownMode = (Global.cfg.BotAutomationFlags & botLockdownFlags) == 0;
                if (bLockdownMode == bNewState)
                {
                    return false;
                }

                if (bNewState)
                {
                    Global.cfg.BotAutomationFlags ^= botLockdownFlags;
                }
                else
                {
                    Global.cfg.BotAutomationFlags |= botLockdownFlags;
                }
                Global.hostApp.Log(LogType.INF, "Lockdown mode {0}", bNewState ? "on" : "off");
                return true;
            }
            if (sName == "debug")
            {
                if (Global.cfg.DebugLoggingEnabled == bNewState)
                {
                    return false;
                }

                Global.cfg.DebugLoggingEnabled = bNewState;
                Global.hostApp.Log(LogType.INF, "Debug mode {0}", Global.cfg.DebugLoggingEnabled ? "on" : "off");
                return true;
            }
            if (sName == "pause")
            {
                if (Global.cfg.IsPaused == bNewState)
                {
                    return false;
                }

                Global.cfg.IsPaused = bNewState;
                Global.hostApp.Log(LogType.INF, "Pause mode {0}", Global.cfg.IsPaused ? "on" : "off");
                return true;
            }
            if (sName == "passive")
            {
                var bPassive = Global.cfg.BotAutomationFlags == Global.BotAutomationFlag.None;
                if (bPassive == bNewState)
                {
                    return false;
                }

                Global.cfg.BotAutomationFlags = bPassive ? Global.BotAutomationFlag.None : Global.BotAutomationFlag.All;
                Global.hostApp.Log(LogType.INF, "Passive mode {0}", bPassive ? "on" : "off");
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

            ZoomMeetingBotSDK.SendChatMessage(recipient, response);

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
            if ((Global.cfg.BotAutomationFlags & Global.BotAutomationFlag.ProcessChat) == 0)
            {
                return;
            }

            //ZoomMeetingBotSDK.SendQueuedChatMessages();
            _ = ZoomMeetingBotSDK.UpdateChat();
            ZoomMeetingBotSDK.SendQueuedChatMessages();
        }

        private static readonly HashSet<string> HsParticipantMessages = new HashSet<string>();

        private static string FirstParticipantGreeted = null;
        private static void DoParticipantActions()
        {
            if ((Global.cfg.BotAutomationFlags & Global.BotAutomationFlag.ProcessParticipants) == 0)
            {
                return;
            }

            _ = ZoomMeetingBotSDK.UpdateParticipants();

            if (ZoomMeetingBotSDK.me != null)
            {
                // If I've got my own participant object, do any self-automation needed

                if (((Global.cfg.BotAutomationFlags & Global.BotAutomationFlag.ReclaimHost) != 0) && (ZoomMeetingBotSDK.me.role != ZoomMeetingBotSDK.ParticipantRole.Host))
                {
                    // TBD: Throttle ReclaimHost attempts?
                    if (ZoomMeetingBotSDK.me.role == ZoomMeetingBotSDK.ParticipantRole.CoHost)
                    {
                        Global.hostApp.Log(LogType.WRN, "BOT I'm Co-Host instead of Host; Trying to reclaim host");
                    }
                    else if (ZoomMeetingBotSDK.me.role == ZoomMeetingBotSDK.ParticipantRole.None)
                    {
                        Global.hostApp.Log(LogType.WRN, "BOT I'm not Host or Co-Host; Trying to reclaim host");
                    }
                    ZoomMeetingBotSDK.ReclaimHost();
                }

                if (((Global.cfg.BotAutomationFlags & Global.BotAutomationFlag.RenameMyself) != 0) && (ZoomMeetingBotSDK.me.name != Global.cfg.MyParticipantName))
                {
                    // Rename myself.  Event handler will type in the name when the dialog pops up
                    Global.hostApp.Log(LogType.INF, "BOT Renaming myself from {0} to {1}", Global.repr(ZoomMeetingBotSDK.me.name), Global.repr(Global.cfg.MyParticipantName));
                    ZoomMeetingBotSDK.RenameParticipant(ZoomMeetingBotSDK.me, Global.cfg.MyParticipantName);
                }

                if (((Global.cfg.BotAutomationFlags & Global.BotAutomationFlag.UnmuteMyself) != 0) && (ZoomMeetingBotSDK.me.audioStatus == ZoomMeetingBotSDK.ParticipantAudioStatus.Muted))
                {
                    // Unmute myself
                    Global.hostApp.Log(LogType.INF, "BOT Unmuting myself");
                    ZoomMeetingBotSDK.UnmuteParticipant(ZoomMeetingBotSDK.me);
                }

                ZoomMeetingBotSDK.UpdateMeetingOptions();
            }

            bool bWaiting = false;
            DateTime dtNow = DateTime.UtcNow;
            foreach (ZoomMeetingBotSDK.Participant p in ZoomMeetingBotSDK.participants.Values)
            {
                // Skip over my own participant record; We handled that earlier
                if (p.isMe)
                {
                    continue;
                }

                string sCleanName = CleanUserName(p.name);
                bool bAdmit = false;
                bool bAdmitKnown = (Global.cfg.BotAutomationFlags & Global.BotAutomationFlag.AdmitKnown) != 0;
                bool bAdmitOthers = (Global.cfg.BotAutomationFlags & Global.BotAutomationFlag.AdmitOthers) != 0;

                if (p.status == ZoomMeetingBotSDK.ParticipantStatus.Waiting)
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
                            Global.hostApp.Log(LogType.INF, "BOT Admitting {0} : KNOWN", Global.repr(p.name));
                            if (ZoomMeetingBotSDK.AdmitParticipant(p))
                            {
                                //SendTopic(p.name, false);
                            }
                        }

                        continue;
                    }

                    // Unknown user
                    DateTime dtWhenToAdmit = p.dtWaiting.AddSeconds(Global.cfg.UnknownParticipantWaitSecs);
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
                        Global.hostApp.Log(LogType.INF, sMsg);
                        HsParticipantMessages.Add(sMsg);
                    }

                    if (bAdmit && ZoomMeetingBotSDK.AdmitParticipant(p))
                    {
                        HsParticipantMessages.Remove(sMsg); // After we admit the user, remove the message
                        dtNextAdmission = dtNow.AddSeconds(Global.cfg.UnknownParticipantThrottleSecs);

                        /*
                        // Participant was successfully admitted.  We want to send them the topic if one is set, but we can't do that
                        //   while they are in the waiting room (DMs cannot be sent to waiting room participants, only broadcast messages),
                        //   so queue up the message for later after they are admitted.
                        SendTopic(p.name, false);
                        */
                    }

                    continue;
                }

                if (p.status == ZoomMeetingBotSDK.ParticipantStatus.Attending)
                {
                    if (((Global.cfg.BotAutomationFlags & Global.BotAutomationFlag.CoHostKnown) != 0) && (p.role == ZoomMeetingBotSDK.ParticipantRole.None) && (ZoomMeetingBotSDK.me.role == ZoomMeetingBotSDK.ParticipantRole.Host))
                    {
                        // If I'm host, and this user is not co-host, check if they should be
                        if (GoodUsers.TryGetValue(sCleanName, out bool bCoHost) && bCoHost)
                        {
                            // Yep, they should be, so do the promotion
                            Global.hostApp.Log(LogType.INF, "BOT Promoting {0} to Co-host", Global.repr(p.name));
                            ZoomMeetingBotSDK.PromoteParticipant(p);
                        }
                    }

                    continue;
                }
            }

            if (bWaiting)
            {
                string waitMsg = Global.cfg.WaitingRoomAnnouncementMessage;

                if (waitMsg == null)
                {
                    return;
                }

                if (waitMsg.Length == 0)
                {
                    return;
                }

                if (Global.cfg.WaitingRoomAnnouncementDelaySecs <= 0)
                {
                    return;
                }

                // At least one person is in the waiting room.  If we're configured to make annoucements to
                //   them, then do so now

                dtNow = DateTime.UtcNow;
                if (dtNow >= dtLastWaitingRoomAnnouncement.AddSeconds(Global.cfg.WaitingRoomAnnouncementDelaySecs))
                {
                    ZoomMeetingBotSDK.SendChatMessage(ZoomMeetingBotSDK.SpecialRecipient.EveryoneInWaitingRoom, waitMsg);
                    dtLastWaitingRoomAnnouncement = dtNow;
                }
            }

            // Greet the first person to join the meeting, but only if we started Zoom
            if ((!ZoomMeetingBotSDK.ZoomAlreadyRunning) && (FirstParticipantGreeted == null))
            {
                var plist = ZoomMeetingBotSDK.participants.ToList();

                // Looking for a participant that is not me, using computer audio, audio is connected, and is a known good user
                var idx = plist.FindIndex(x => (
                    (!x.Value.isMe) &&
                    (x.Value.device == ZoomMeetingBotSDK.ParticipantAudioDevice.Computer) &&
                    (x.Value.audioStatus != ZoomMeetingBotSDK.ParticipantAudioStatus.Disconnected) &&
                    GoodUsers.ContainsKey(CleanUserName(x.Value.name))
                ));
                if (idx != -1)
                {
                    FirstParticipantGreeted = plist[idx].Value.name;
                    var msg = FormatChatResponse(OneTimeHi("morning", FirstParticipantGreeted), FirstParticipantGreeted);

                    Sound.Play("bootup");
                    Thread.Sleep(3000);
                    Sound.Speak(Global.cfg.MyParticipantName + " online.");
                    ZoomMeetingBotSDK.SendChatMessage(ZoomMeetingBotSDK.SpecialRecipient.EveryoneInMeeting, true, msg);
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
                Global.hostApp.Log(LogType.WRN, "TimerIdleHandler {0:X4} - Busy; Will try again later", nTimerIterationID);
                return;
            }

            try
            {
                //Global.hostApp.Log(LogType.DBG, "TimerIdleHandler {0:X4} - Enter");

                Global.LoadSettings();
                LoadGoodUsers();
                ReadRemoteCommands();

                // Zoom is really bad about moving/resizing it's windows, so keep it in check
                ZoomMeetingBotSDK.LayoutWindows();

                if (Global.cfg.IsPaused)
                {
                    return;
                }

                //Global.hostApp.Log(LogType.DBG, "TimerIdleHandler {0:X4} - DoParticipantActions", nTimerIterationID);
                DoParticipantActions();

                //Global.hostApp.Log(LogType.DBG, "TimerIdleHandler {0:X4} - DoChatActions", nTimerIterationID);
                DoChatActions();
            }
            catch (ZoomMeetingBotSDK.ZoomClosedException ex)
            {
                Global.hostApp.Log(LogType.INF, ex.ToString());
                ShouldExit = true;
            }
            catch (Exception ex)
            {
                Global.hostApp.Log(LogType.ERR, "TimerIdleHandler {0:X4} - Unhandled Exception: {1}", nTimerIterationID, ex.ToString());
            }
            finally
            {
                //Global.hostApp.Log(LogType.DBG, "TimerIdleHandler {0:X4} - Exit", nTimerIterationID);
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

            Global.hostApp.Log(LogType.INF, "Processing Remote Commands");

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
                        Global.hostApp.Log(LogType.INF, "Received {0} command", line);
                        ZoomMeetingBotSDK.LeaveMeeting(false);
                        ShouldExit = true;
                    }
                    else if (line == "kill")
                    {
                        Global.hostApp.Log(LogType.INF, "Received {0} command", line);
                        ZoomMeetingBotSDK.LeaveMeeting(true);
                    }
                    else
                    {
                        Global.hostApp.Log(LogType.ERR, "Unknown command: {0}", line);
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

            Global.hostApp.Log(LogType.INF, "(Re-)loading GoodUsers");

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

        private static void OnMeetingOptionStateChange(object sender, ZoomMeetingBotSDK.MeetingOptionStateChangeEventArgs e)
        {
            Global.hostApp.Log(LogType.INF, "Meeting option {0} changed to {1}", Global.repr(e.optionName), e.newState.ToString());
        }

        private static void OnParticipantAttendanceStatusChange(object sender, ZoomMeetingBotSDK.ParticipantEventArgs e)
        {
            ZoomMeetingBotSDK.Participant p = e.participant;
            Global.hostApp.Log(LogType.INF, "Participant {0} status {1}", Global.repr(p.name), p.status.ToString());

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

            // Do one-time "hi" only once
            if (DicOneTimeHis.ContainsKey(to))
            {
                return null;
            }

            // Try to give a specific response
            foreach (var word in text.GetWordsInSentence())
            {
                if (Global.cfg.OneTimeHiSequences.TryGetValue(word.ToLower(), out response))
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

        private static void SetSpeaker(ZoomMeetingBotSDK.Participant p, string from)
        {
            ZoomMeetingBotSDK.SendChatMessage(from, "Speaker mode is not yet implemented");

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

        private static void OnChatMessageReceive(object source, ZoomMeetingBotSDK.ChatEventArgs e)
        {
            Global.hostApp.Log(LogType.INF, "New message from {0} to {1}: {2}", Global.repr(e.from), Global.repr(e.to), Global.repr(e.text));

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
                var withoutMyName = Regex.Replace(sMsg, @"\b" + Global.cfg.MyParticipantName + @"\b", string.Empty, RegexOptions.IgnoreCase);

                // If strings are the same, it's not to me
                if (withoutMyName == sMsg)
                {
                    return;
                }

                sMsg = withoutMyName;

                // My name is in it, so reply to everyone
                sReplyTo = ZoomMeetingBotSDK.SpecialRecipient.EveryoneInMeeting;
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

                var isToEveryone = ZoomMeetingBotSDK.SpecialRecipient.IsEveryone(sReplyTo);

                // If the bot is addressed publically or if there are only two people in the meeting, then reply with TTS
                // TBD: Should be attending count, not participant count.  Some could be in the waiting room
                var speak = isToEveryone || (ZoomMeetingBotSDK.participants.Count == 2);

                // We start with a one-time hi.  Various bots may be in different time zones and the
                //   good morning/afternoon/evening throws things off
                var response = OneTimeHi(sMsg, sFrom);

                // Handle canned responses based on broadcast keywords.  TBD: Move this into a bot
                foreach (var broadcastCommand in Global.cfg.BroadcastCommands)
                {
                    if (FastRegex.IsMatch($"\\b${broadcastCommand.Key}\\b", sMsg, RegexOptions.IgnoreCase))
                    {
                        response = broadcastCommand.Value;

                        // Don't want to speak broadcast messages
                        speak = false;
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

                        Global.hostApp.Log(LogType.WRN, $"ChatBot converse with {Global.repr(chatBot.GetChatBotInfo().Name)} failed: {Global.repr(failureMsg)}");
                    }
                }

                if (response == null)
                {
                    Global.hostApp.Log(LogType.ERR, "No ChatBot was able to produce a response");
                }

                ZoomMeetingBotSDK.SendChatMessage(sReplyTo, speak, FormatChatResponse(response, sFrom));

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
                Global.hostApp.Log(LogType.WRN, "Ignoring command {0} from non-admin {1}", Global.repr(sMsg), Global.repr(sFrom));
                return;
            }

            if (!ZoomMeetingBotSDK.participants.TryGetValue(sFrom, out ZoomMeetingBotSDK.Participant sender))
            {
                Global.hostApp.Log(LogType.ERR, "Received command {0} from {1}, but I don't have a Participant class for them", Global.repr(sMsg), Global.repr(e.from));
                return;
            }

            string[] a = sMsg.Split(SpaceDelim, 2);

            string sCommand = a[0].ToLower().Substring(1);

            // All of the following commands require an argument
            string sTarget = (a.Length == 1) ? null : (a[1].Length == 0 ? null : a[1]);

            if (Global.cfg.BroadcastCommands.TryGetValue(sCommand, out string sBroadcastMsg))
            {
                DateTime dtNow = DateTime.UtcNow;

                if (BroadcastSentTime.TryGetValue(sCommand, out DateTime dtSentTime))
                {
                    int guardTime = Global.cfg.BroadcastCommandGuardTimeSecs;

                    if (guardTime < 0)
                    {
                        ZoomMeetingBotSDK.SendChatMessage(sender.name, $"{sCommand}: This broadcast message was already sent.");
                        return;
                    }

                    if ((guardTime > 0) && (dtNow <= dtSentTime.AddSeconds(Global.cfg.BroadcastCommandGuardTimeSecs)))
                    {
                        ZoomMeetingBotSDK.SendChatMessage(sender.name, $"{sCommand}: This broadcast message was already sent recently. Please try again later.");
                        return;
                    }
                }

                ZoomMeetingBotSDK.SendChatMessage(ZoomMeetingBotSDK.SpecialRecipient.EveryoneInMeeting, sBroadcastMsg);
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

                ZoomMeetingBotSDK.SendChatMessage(sReplyTo, reply);

                if (broadcast)
                {
                    ZoomMeetingBotSDK.SendChatMessage(ZoomMeetingBotSDK.SpecialRecipient.EveryoneInMeeting, GetTopic());
                }

                return;
            }

            // All of the following commands require options
            if (sTarget == null)
            {
                return;
            }

            if (Global.cfg.EmailCommands.TryGetValue(sCommand, out Global.EmailCommandArgs emailCommandArgs))
            {
                string[] args = sTarget.Trim().Split(SpaceDelim, 2);

                string toAddress = args[0];
                string subject = emailCommandArgs.Subject;
                string body = emailCommandArgs.Body;

                if (subject.Contains("{0}") || body.Contains("{0}"))
                {
                    if (args.Length <= 1)
                    {
                        ZoomMeetingBotSDK.SendChatMessage(sender.name, $"Error: The format of the command is incorrect; Correct example: /{sCommand} {emailCommandArgs.ArgsExample}");
                        return;
                    }

                    string emailArg = args[1].Trim();
                    subject = subject.Replace("{0}", emailArg);
                    body = body.Replace("{0}", emailArg);
                }

                if (SendEmail(subject, body, toAddress))
                {
                    ZoomMeetingBotSDK.SendChatMessage(sender.name, $"{sCommand}: Successfully sent email to {toAddress}");
                }
                else
                {
                    ZoomMeetingBotSDK.SendChatMessage(sender.name, $"{sCommand}: Failed to send email to {toAddress}");
                }


                return;
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
                    ZoomMeetingBotSDK.SendChatMessage(sender.name, "Sorry, the {0} command requires either on or off as a parameter", Global.repr(sCommand));
                    return;
                }

                if (SetMode(sCommand, bNewMode))
                {
                    ZoomMeetingBotSDK.SendChatMessage(sender.name, "{0} mode has been changed to {1}", GetFirstName(sCommand), sNewMode);
                }
                else
                {
                    ZoomMeetingBotSDK.SendChatMessage(sender.name, "{0} mode is already {1}", GetFirstName(sCommand), sNewMode);
                }
                return;
            }

            if (sCommand == "waitmsg")
            {
                var sWaitMsg = sTarget.Trim();

                if ((sWaitMsg.Length == 0) || (sWaitMsg.ToLower() == "off"))
                {
                    if ((Global.cfg.WaitingRoomAnnouncementMessage != null) && (Global.cfg.WaitingRoomAnnouncementMessage.Length > 0))
                    {
                        Global.cfg.WaitingRoomAnnouncementMessage = null;
                        ZoomMeetingBotSDK.SendChatMessage(sender.name, "Waiting room message has been turned off");
                    }
                    else
                    {
                        ZoomMeetingBotSDK.SendChatMessage(sender.name, "Waiting room message is already off");
                    }
                }
                else if (sWaitMsg == Global.cfg.WaitingRoomAnnouncementMessage)
                {
                    ZoomMeetingBotSDK.SendChatMessage(sender.name, "Waiting room message is already set to:\n{0}", sTarget);
                }
                else
                {
                    Global.cfg.WaitingRoomAnnouncementMessage = sTarget.Trim();
                    ZoomMeetingBotSDK.SendChatMessage(sender.name, "Waiting room message has set to:\n{0}", sTarget);
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
                    ZoomMeetingBotSDK.SendChatMessage(sender.name, "Please use the format: /{0} Old Name to New Name", sCommand);
                    ZoomMeetingBotSDK.SendChatMessage(sender.name, "Example: /{0} iPad User to John Doe", sCommand);
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
                ZoomMeetingBotSDK.SendChatMessage(ZoomMeetingBotSDK.SpecialRecipient.EveryoneInMeeting, sCommand == "speak", sTarget);

                return;
            }

            if (sCommand == "play")
            {
                ZoomMeetingBotSDK.SendChatMessage(sender.name, "Playing: {0}", Global.repr(sTarget));
                Sound.Play(sTarget);
                return;
            }

            // If the sender refers to themselves as "me", resolve this to their actual participant name
            if (sTarget.ToLower() == "me")
            {
                sTarget = e.from;
            }

            // All of the following require a participant target
            if (!ZoomMeetingBotSDK.participants.TryGetValue(sTarget, out ZoomMeetingBotSDK.Participant target))
            {
                ZoomMeetingBotSDK.SendChatMessage(sender.name, "Sorry, I don't see anyone named here named {0}. Remember, Case Matters!", Global.repr(sTarget));
                return;
            }

            // Make sure I'm not the target :p
            if (target.isMe)
            {
                ZoomMeetingBotSDK.SendChatMessage(sender.name, "U Can't Touch This\n* MC Hammer Music *\nhttps://youtu.be/otCpCn0l4Wo");
                return;
            }

            // Do rename if requested
            if (newName != null)
            {
                if (target.name == sender.name)
                {
                    ZoomMeetingBotSDK.SendChatMessage(sender.name, "Why don't you just rename yourself?");
                    return;
                }

                ZoomMeetingBotSDK.SendChatMessage(sender.name, "Renaming {0} to {1}", Global.repr(target.name), Global.repr(newName));
                ZoomMeetingBotSDK.RenameParticipant(target, newName);
                return;
            }

            if (sCommand == "admit")
            {
                if (target.status != ZoomMeetingBotSDK.ParticipantStatus.Waiting)
                {
                    ZoomMeetingBotSDK.SendChatMessage(sender.name, "Sorry, {0} is not waiting", Global.repr(target.name));
                }
                else
                {
                    ZoomMeetingBotSDK.SendChatMessage(sender.name, "Admitting {0}", Global.repr(target.name));
                    if (ZoomMeetingBotSDK.AdmitParticipant(target))
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
            if (target.status != ZoomMeetingBotSDK.ParticipantStatus.Attending)
            {
                ZoomMeetingBotSDK.SendChatMessage(sender.name, "Sorry, {0} is not attending", Global.repr(target.name));
                return;
            }

            if ((sCommand == "cohost") || (sCommand == "promote"))
            {
                if (target.role != ZoomMeetingBotSDK.ParticipantRole.None)
                {
                    ZoomMeetingBotSDK.SendChatMessage(sender.name, "Sorry, {0} is already Host or Co-Host so cannot be promoted", Global.repr(target.name));
                }
                else if (target.videoStatus != ZoomMeetingBotSDK.ParticipantVideoStatus.On)
                {
                    ZoomMeetingBotSDK.SendChatMessage(sender.name, "Co-Host name matched for {0}, but video is off", Global.repr(target.name));
                    return;
                }
                else
                {
                    ZoomMeetingBotSDK.SendChatMessage(sender.name, "Promoting {0} to Co-Host", Global.repr(target.name));
                    ZoomMeetingBotSDK.PromoteParticipant(target);
                }

                return;
            }

            if (sCommand == "demote")
            {
                if (target.role != ZoomMeetingBotSDK.ParticipantRole.CoHost)
                {
                    ZoomMeetingBotSDK.SendChatMessage(sender.name, "Sorry, {0} isn't Co-Host so cannot be demoted", Global.repr(target.name));
                }
                else
                {
                    ZoomMeetingBotSDK.SendChatMessage(sender.name, "Demoting {0}", Global.repr(target.name));
                    ZoomMeetingBotSDK.DemoteParticipant(target);
                }

                return;
            }

            if (sCommand == "mute")
            {
                ZoomMeetingBotSDK.SendChatMessage(sender.name, "Muting {0}", Global.repr(target.name));
                ZoomMeetingBotSDK.MuteParticipant(target);
                return;
            }

            if (sCommand == "unmute")
            {
                ZoomMeetingBotSDK.SendChatMessage(sender.name, "Requesting {0} to Unmute", Global.repr(target.name));
                ZoomMeetingBotSDK.UnmuteParticipant(target);
                return;
            }

            if (sCommand == "speaker")
            {
                SetSpeaker(target, e.from);
                return;
            }

            ZoomMeetingBotSDK.SendChatMessage(sender.name, "Sorry, I don't know the command {0}", Global.repr(sCommand));
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
                    Global.hostApp.Log(LogType.WRN, $"Cannot load bot in {Global.repr(subdir.FullName)}; More than one DLL found");
                }
                else if (files.Length == 0)
                {
                    Global.hostApp.Log(LogType.WRN, $"Cannot load bot in {Global.repr(subdir.FullName)}; No DLL found");
                }
                else
                {
                    var file = files[0];
                    try
                    {
                        Global.hostApp.Log(LogType.DBG, $"Loading {file.Name}");
                        var assembly = Assembly.LoadFile(file.FullName);
                        var types = assembly.GetTypes();
                        foreach (var type in types)
                        {
                            List<Type> interfaceTypes = new List<Type>(type.GetInterfaces());
                            if (interfaceTypes.Contains(typeof(IChatBot)))
                            {
                                var chatBot = Activator.CreateInstance(type) as IChatBot;
                                var chatBotInfo = chatBot.GetChatBotInfo();
                                chatBot.Start(new ChatBotInitParam()
                                {
                                    hostApp = Global.hostApp,
                                });
                                Global.hostApp.Log(LogType.DBG, $"Loaded {Global.repr(chatBotInfo.Name)} chatbot with intelligence level {chatBotInfo.IntelligenceLevel}");
                                bots.Add(new Tuple<int, IChatBot>(chatBotInfo.IntelligenceLevel, chatBot));
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        Global.hostApp.Log(LogType.ERR, $"Failed to load {Global.repr(file.FullName)}: {Global.repr(ex.ToString())}");
                    }
                }
            }

            if (bots.Count == 0)
            {
                return null;
            }

            return bots.OrderByDescending(o => o.Item1).Select(x => x.Item2).ToList();
        }

        public static void SettingsUpdated()
        {
            if (chatBots != null)
            {
                // We'll try each bot in order by intelligence level until one of them works
                foreach (var chatBot in chatBots)
                {
                    try
                    {
                        chatBot.SettingsUpdated();
                    }
                    catch (Exception ex)
                    {
                        Global.hostApp.Log(LogType.ERR, $"SettingsNotify failed: {Global.repr(ex.ToString())}");
                    }
                }
            }
        }

        public static void Run()
        {
            if ((Global.cfg.BotAutomationFlags & Global.BotAutomationFlag.Converse) != 0)
            {
                chatBots = GetChatBots();
            }

            ZoomMeetingBotSDK.ParticipantAttendanceStatusChange += OnParticipantAttendanceStatusChange;
            ZoomMeetingBotSDK.ChatMessageReceive += OnChatMessageReceive;
            ZoomMeetingBotSDK.MeetingOptionStateChange += OnMeetingOptionStateChange;
            ZoomMeetingBotSDK.Start();

            tmrIdle = new System.Threading.Timer(TimerIdleHandler, null, 0, 5000);

            return;
        }

        /// <summary>
        /// Leaves the meeting, optionally ending meeting or passing off Host role to another participant.
        /// </summary>
        public static void LeaveMeeting(bool endForAll = false)
        {
            if (!endForAll)
            {
                if (ZoomMeetingBotSDK.me.role != ZoomMeetingBotSDK.ParticipantRole.Host)
                {
                    Global.hostApp.Log(LogType.DBG, "BOT LeaveMeeting - I am not host");
                }
                else
                {
                    Global.hostApp.Log(LogType.DBG, "BOT LeaveMeeting - I am host; Trying to find someone to pass it to");

                    ZoomMeetingBotSDK.Participant altHost = null;
                    foreach (ZoomMeetingBotSDK.Participant p in ZoomMeetingBotSDK.participants.Values)
                    {
                        if (p.role == ZoomMeetingBotSDK.ParticipantRole.CoHost)
                        {
                            altHost = p;
                            break;
                        }
                    }

                    if (altHost == null)
                    {
                        Global.hostApp.Log(LogType.ERR, "BOT LeaveMeeting - Could not find an alternative host; Ending meeting");
                        endForAll = true;
                    }
                    else
                    {
                        try
                        {
                            Global.hostApp.Log(LogType.INF, "BOT LeaveMeeting - Passing Host to {0}", Global.repr(altHost.name));
                            ZoomMeetingBotSDK.PromoteParticipant(altHost, ZoomMeetingBotSDK.ParticipantRole.Host);
                            Global.hostApp.Log(LogType.INF, "BOT LeaveMeeting - Passed Host to {0}", Global.repr(altHost.name));
                        }
                        catch (Exception ex)
                        {
                            Global.hostApp.Log(LogType.ERR, "BOT LeaveMeeting - Failed to pass Host to {0}; Ending meeting", Global.repr(altHost.name));
                            endForAll = true;
                        }
                    }
                }
            }

            Global.hostApp.Log(LogType.INF, "BOT LeaveMeeting - Leaving Meeting");
            ZoomMeetingBotSDK.LeaveMeeting(endForAll);
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

                Global.hostApp.Log(LogType.ERR, "SendEmail - Sending email to {0} with subject {1}", Global.repr(to), Global.repr(subject));
                gmailSender.Send(new GmailSenderLib.SimpleMailMessage(subject, body, to));

                return true;
            }
            catch (Exception ex)
            {
                Global.hostApp.Log(LogType.ERR, "SendEmail - Failed; Exception: {0}", Global.repr(ex.ToString()));
                return false;
            }
        }
    }
}
