namespace ZoomMeetingBotSDK
{
    using System;
    using System.Collections.Concurrent;
    using System.Collections.Generic;
    using System.ComponentModel;
    using System.Diagnostics;
    using System.Drawing;
    using System.Globalization;
    using System.Text;
    using System.Text.RegularExpressions;
    using System.Threading;
    using System.Windows;
    using System.Windows.Automation;
    using System.Windows.Forms;

    using static Utils;

    public class Controller
    {
        public class ControllerConfigurationSettings
        {
            public ControllerConfigurationSettings()
            {
                BrowserExecutable = @"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe";
                BrowserArguments = "https://zoom.us/signin";
                MyParticipantName = "ZoomBot";            
                MeetingID = null;
                DisableClipboardPasteText = false;
                ParticipantCountMismatchRetries = 3;
                UIActionDelayMilliseconds = 250;
                ClickDelayMilliseconds = UIActionDelayMilliseconds;
                KeyboardInputDelayMilliseconds = UIActionDelayMilliseconds;
                DisableParticipantPaging = false;
                MouseMovementRate = 100;
                UpdateMeetingOptionsDelaySecs = -1;
                Screen = null;
                ZoomExecutable = @"%AppData%\Zoom\bin\Zoom.exe";
                ZoomUsername = null;
                ZoomPassword = null;
                EnableWalkRawElementsToString = false;
            }

            /// <summary>
            /// Absolute path to web browser executable.
            /// </summary>
            public string BrowserExecutable { get; set; }

            /// <summary>
            /// Optional command line arguments to pass to web browser.
            /// </summary>
            public string BrowserArguments { get; set; }

            /// <summary>
            /// Name to use when joining the Zoom meeting.  If the default name does not match, a rename is done after joining the meeting.
            /// </summary>
            public string MyParticipantName { get; set; }

            /// <summary>
            /// ID of the meeting to join.
            /// </summary>
            public string MeetingID { get; set; }

            /// <summary>
            /// Disables using clipboard to paste text into target apps; Falls back on sending individual keystrokes instead.
            /// </summary>
            public bool DisableClipboardPasteText { get; set; }

            /// <summary>
            /// Number of times to retry parsing participant list until count matches what is in the window title.
            /// </summary>
            public int ParticipantCountMismatchRetries { get; set; }

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
            /// Disable paging up/down in the Participants window. If screen resolution is sufficiently big (height), it may not be needed.
            /// </summary>
            public bool DisableParticipantPaging { get; set; }

            /// <summary>
            /// Duration of time over which to do mouse moves when simulating input. This helps the target app "see" the mouse movement more reliably. <= 0 will move the mouse instantly.
            /// </summary>
            public int MouseMovementRate { get; set; }

            /// <summary>
            /// Seconds to wait between opening the meeting options menu to get status from the UI.
            /// Special values:
            ///   -1  Only poll when meeting is first started, then poll on demand (when one of the options needs to be changed)
            ///    0  Poll as fast as possible
            ///   >0  Delay between polls
            /// </summary>
            public int UpdateMeetingOptionsDelaySecs { get; set; }

            /// <summary>
            /// Configures which display should be used to run Zoom, ZoomMeetingBotSDK, UsherBot, etc.  The default is whichever display is set as the "main" screen
            /// </summary>
            public string Screen { get; set; }

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

            /// <summary>
            /// Enables the function WalkRawElementsToString() which walks the entire AutomationElement tree and returns a string.  The default is True.  This is an
            /// extremely expensive operation, often taking a minute or more, so disabling it on live meetings may be a good idea.
            /// </summary>
            public bool EnableWalkRawElementsToString { get; set; }
        }
        public static ControllerConfigurationSettings cfg = new ControllerConfigurationSettings();

        public enum MeetingOption
        {
            [Description("Mute Participants upon Entry")]
            MuteParticipantsUponEntry,
            [Description("Allow Participants to Unmute Themselves")]
            AllowParticipantsToUnmuteThemselves,
            [Description("Allow Participants to Rename Themselves")]
            AllowParticipantsToRenameThemselves,
            [Description("Play sound when someone joins or leaves")]
            PlaySoundWhenSomeoneJoinsOrLeaves,
            [Description("Enable Waiting Room")]
            EnableWaitingRoom,
            [Description("Lock Meeting")]
            LockMeeting,
        }

        public enum ParticipantCanChatWith
        {
            [Description("No One")]
            NoOne,
            [Description("Host Only")]
            HostOnly,
            [Description("Everyone Publicly")]
            EveryonePublically,
            [Description("Everyone Publicly and Privately")]
            EveryonePublicallyAndPrivately,
        }

        /// <summary>
        /// Class providing Enum of strings. Uses description attribute and reflection. Caches values for efficiency.
        /// Based on this article: https://stackoverflow.com/questions/4367723/get-enum-from-description-attribute.
        /// </summary>
        private static class EnumEx
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
                        throw new KeyNotFoundException(string.Format("Description/Name {0} not found in enum {1}", repr(description), repr(type.ToString())));
                    }
                }

                return (T)ret;
            }
        }

        public class MeetingOptionStateChangeEventArgs : EventArgs
        {
            public MeetingOption optionValue;
            public string optionName;
            public ToggleState oldState;
            public ToggleState newState;
        }

        public static event EventHandler<MeetingOptionStateChangeEventArgs> MeetingOptionStateChange = (sender, e) => { };

        // Was the Zoom app already running, or did we start it ourselves
        public static bool ZoomAlreadyRunning = false;

        internal static IHostApp hostApp;

        private static readonly Dictionary<MeetingOption, ToggleState> _globalParticipantOptions = new Dictionary<MeetingOption, ToggleState>();

        private class PromoteParticipantActionArgs
        {
            public ParticipantRole NewRole;
        }

        private class IndividualParticipantActionArgs
        {
            public IndividualParticipantActionArgs(Participant target, IndividualParticipantAction function, object args)
            {
                this.Target = target;
                this.Args = args;
                this.Function = function;

                this.WhenAdded = DateTime.UtcNow;
                this.Attempts = 0;
            }

            public Participant Target;
            public DateTime WhenAdded;
            public object Args;
            public int Attempts;
            public IndividualParticipantAction Function;
        }

        private delegate bool IndividualParticipantAction(IndividualParticipantActionArgs args);

        private static Dictionary<string, Queue<IndividualParticipantActionArgs>> IndividualParticipantActionQueue = new Dictionary<string, Queue<IndividualParticipantActionArgs>>();

        private static DateTime nextMeetingOptionsUpdate = DateTime.MinValue;

        private static void QueueIndividualParticipantAction(IndividualParticipantActionArgs args)
        {
            Queue<IndividualParticipantActionArgs> q;
            var targetName = args.Target.name;
            bool queueExists = IndividualParticipantActionQueue.TryGetValue(targetName, out q);

            if (!queueExists)
            {
                q = new Queue<IndividualParticipantActionArgs>();
            }
            else
            {
                // Check if the requested action is a duplicate
                foreach (var item in q)
                {
                    // TBD: Might want to also check arguments for uniqueness
                    if (item.Function.Method.Name == args.Function.Method.Name)
                    {
                        hostApp.Log(LogType.WRN, "Ignoring duplicate action {0} for participant {1}", item.Function.Method.Name, repr(targetName));
                        return;
                    }
                }
            }

            q.Enqueue(args);

            if (!queueExists)
            {
                IndividualParticipantActionQueue[targetName] = q;
            }
        }

        /// <summary>
        /// Executes any queued individual participant actions for the given participant.  If one of the actions fails, stops processing.
        /// </summary>
        /// <returns>True if one or more actions were successfully processed; False otherwise.</returns>
        private static bool ExecuteQueuedIndividualParticipantActions(Participant p)
        {
            // TBD: Clear out queued items that timeout
            bool ret = false;

            var targetName = p.name;
            if (!IndividualParticipantActionQueue.TryGetValue(targetName, out Queue<IndividualParticipantActionArgs> q))
            {
                // No queue
                return ret;
            }

            while (true)
            {
                var args = q.Peek();
                args.Attempts++;

                // Update AutomationElement for this target (The participants move around in the list pretty regularly; This makes sure we target the right entry if it has moved)
                args.Target._ae = p._ae;

                hostApp.Log(LogType.INF, "Attempt #{2} for participant {0} action {1}", repr(targetName), args.Function.Method.Name, args.Attempts);
                var success = args.Function(args);

                if (success)
                {
                    ret = true;
                }
                else
                {
                    if (args.Attempts == 3)
                    {
                        hostApp.Log(LogType.ERR, "Max attempts reached for participant {0} action {1}", repr(targetName), args.Function.Method.Name);

                        // Fall through to remove this action from the queue
                    }
                    else
                    {
                        hostApp.Log(LogType.WRN, "Attempt #{2} failure for participant {0} action {1}", repr(targetName), args.Function.Method.Name, args.Attempts);

                        // The action failed; We can't process any other actions until this one is done, so stop.  We'll try again later
                        break;
                    }
                }

                _ = q.Dequeue();

                // Remove the queue from the dictionary if there's nothing left in it
                if (q.Count == 0)
                {
                    IndividualParticipantActionQueue.Remove(targetName);

                    // We reached the end of the queue, so stop
                    hostApp.Log(LogType.DBG, "Done processing participant {0} actions", repr(targetName));
                    break;
                }
            }

            return ret;
        }

        private static ToggleState GetMeetingOption(MeetingOption option, bool updateIfNeeded = true)
        {
            if ((updateIfNeeded) && (cfg.UpdateMeetingOptionsDelaySecs == -1))
            {
                UpdateMeetingOptions(true);
            }

            if (!_globalParticipantOptions.TryGetValue(option, out ToggleState ret))
            {
                return ToggleState.Indeterminate;
            }

            return ret;
        }

        private static void SetMeetingOption(MeetingOption option, ToggleState newState)
        {
            VerifyMyRole("SetMeetingOption", ParticipantRole.Host);

            ToggleState currentState = GetMeetingOption(option);

            // If value is already what we want, then there's nothing to do!
            if (newState == currentState)
            {
                return;
            }

            // It's not what we want, so click it to toggle
            //_ = ClickParticipantWindowControl("More options to manage all participants", true, ControlType.SplitButton);
            _ = ClickMoreOptionsToManageAllParticipants();
            InvokeMenuItem(aeParticipantsWindow, EnumEx.GetDescriptionFromValue(option));
        }

        public static void UpdateMeetingOptions(bool force = false)
        {
            /* UpdateMeetingOptionDelaySecs:
             * -1  Only poll when meeting is first started, then poll on demand (when one of the options needs to be changed)
             *  0  Poll as fast as possible
             *  >0  Delay between polls
             */
            if (force)
            {
                hostApp.Log(LogType.DBG, "UpdateMeetingOptions - Forced update");
            }
            else if (cfg.UpdateMeetingOptionsDelaySecs == -1)
            {
                if (nextMeetingOptionsUpdate == DateTime.MinValue)
                {
                    hostApp.Log(LogType.DBG, "UpdateMeetingOptions - Performing initial update");
                    nextMeetingOptionsUpdate = DateTime.MaxValue;
                }
                else
                {
                    hostApp.Log(LogType.DBG, "UpdateMeetingOptions - Already performed initial update");
                    return;
                }
            }
            else if (cfg.UpdateMeetingOptionsDelaySecs == 0)
            {
                hostApp.Log(LogType.DBG, "UpdateMeetingOptions - Polling as fast as possible");
            }
            else if (cfg.UpdateMeetingOptionsDelaySecs > 0)
            {
                if (DateTime.UtcNow < nextMeetingOptionsUpdate)
                {
                    hostApp.Log(LogType.DBG, "UpdateMeetingOptions - Not time to update yet (Next: {0})", repr(nextMeetingOptionsUpdate));
                    return;
                }

                nextMeetingOptionsUpdate = DateTime.UtcNow.AddSeconds(cfg.UpdateMeetingOptionsDelaySecs);
            }

            hostApp.Log(LogType.DBG, "UpdateMeetingOptions - Enter");
            try
            {
                //_ = ClickParticipantWindowControl("More options to manage all participants", true, ControlType.SplitButton);
                _ = ClickMoreOptionsToManageAllParticipants();

                hostApp.Log(LogType.DBG, "UpdateMeetingOptions - WaitPopupMenu");
                var menu = WaitPopupMenu();

                //hostApp.Log(LogType.DBG, "UpdateMeetingOptions - Menu AETree: {0}", UIATools.WalkRawElementsToString(menu));

                hostApp.Log(LogType.DBG, "UpdateMeetingOptions - Caching Menu Items");
                CacheRequest cr = new CacheRequest();
                cr.Add(AutomationElement.NameProperty);
                cr.Add(TogglePatternIdentifiers.Pattern);
                cr.Add(TogglePatternIdentifiers.ToggleStateProperty);
                cr.TreeScope = TreeScope.Element | TreeScope.Children;
                AutomationElement list;

                using (cr.Activate())
                {
                    // Cache the specified properties for menu items
                    list = menu.FindFirst(TreeScope.Element, new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Menu));
                }
                if (list == null)
                {
                    return;
                }

                hostApp.Log(LogType.DBG, "UpdateMeetingOptions - Retrieving CachedChildren");
                var options = list.CachedChildren;

                hostApp.Log(LogType.DBG, "UpdateMeetingOptions - Closing Menu");
                WindowTools.SendKeys("{ESC}"); // Break out of menu. TBD: May be better just to click somewhere?

                //var options = menu.FindAll(TreeScope.Children, new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.CheckBox));

                hostApp.Log(LogType.DBG, "UpdateMeetingOptions - Parsing Menu Items");
                foreach (AutomationElement option in options)
                {
                    // Skip over blank/spacer options
                    var optionName = option.Cached.Name;
                    if (optionName.Length == 0)
                    {
                        continue;
                    }

                    // Get toggle state
                    if (!option.TryGetCachedPattern(TogglePatternIdentifiers.Pattern, out object togglePattern))
                    {
                        continue;
                    }

                    var newState = ((TogglePattern)togglePattern).Cached.ToggleState;

                    // Get current value
                    var optionValue = EnumEx.GetValueFromDescription<MeetingOption>(optionName);

                    // Update value if changed
                    ToggleState oldState = GetMeetingOption(optionValue, false);
                    if (newState != oldState)
                    {
                        _globalParticipantOptions[optionValue] = newState;
                        MeetingOptionStateChange(null, new MeetingOptionStateChangeEventArgs
                        {
                            optionValue = optionValue,
                            optionName = optionName,
                            oldState = oldState,
                            newState = newState,
                        });
                    }

                    //MeetingOptionStateChange
                }
            }
            finally
            {
                hostApp.Log(LogType.DBG, "UpdateMeetingOptions - Exit");
            }
        }

        /// <summary>
        /// Subscribes to UIAutomation events from the Zoom app's windows.  See more details here:
        /// https://docs.microsoft.com/en-us/dotnet/framework/ui-automation/subscribe-to-ui-automation-events.
        /// </summary>
        private class EventWatcher : IDisposable
        {
            private readonly EventWaitHandle eventWaitHandle = new EventWaitHandle(false, EventResetMode.ManualReset);
            private readonly object queueLock = new object();

            private bool isDisposed = false;
            private int processID = 0;
            private string lastEventKey = null;
            private Queue<UIEvent> eventQueue = new Queue<UIEvent>();
            private bool isWatching = false;
            private bool isEventHandlerAdded = false;

            public EventWatcher(int pid)
            {
                this.processID = pid;
                hostApp.Log(LogType.DBG, "Setting Up Event Handlers");

                System.Windows.Automation.Automation.AddAutomationFocusChangedEventHandler(
                    new AutomationFocusChangedEventHandler(OnFocusChangedEvent));
                isEventHandlerAdded = true;
            }

            private void OnFocusChangedEvent(object src, AutomationEventArgs e)
            {
                // Sometimes we get null events... strange
                if (src is null)
                {
                    return;
                }

                try
                {
                    AutomationElement ae;
                    AutomationElement.AutomationElementInformation aei;

                    try
                    {
                        ae = src as AutomationElement;
                        aei = ae.Current;

                        var nCPID = aei.ProcessId;
                        if (nCPID != this.processID)
                        {
                            //hostApp.Log(LogType.WRN, "Got {0} for {1} (Wrong process)", UIATools.GetEventShortName(e), UIATools.AEToString(ae));
                            return;
                        }
                    }
                    catch (ElementNotAvailableException)
                    {
                        return;
                    }

                    lock (this.queueLock)
                    {
                        /*
                         * Sometimes we can get dup events with slightly different data, and this can throw us off.  For example:
                         * 2020-06-24 23:47:32.059 DBG *** EVENT ENQUEUE Focus "LocalizedControlType":"list item","Name":"John Doe","HasKeyboardFocus":true,"IsEnabled":true,"BoundingRectangle":{"IsEmpty":false,"Location":{"X":1603,"Y":740},"Size":{"IsEmpty":false,"Width":200,"Height":27},"X":1603,"Y":740,"Width":200,"Height":27,"Left":1603,"Top":740,"Right":1803,"Bottom":767,"TopLeft":{"X":1603,"Y":740},"TopRight":{"X":1803,"Y":740},"BottomLeft":{"X":1603,"Y":767},"BottomRight":{"X":1803,"Y":767}},"IsControlElement":true,"IsContentElement":true
                         * 2020-06-24 23:47:32.494 DBG *** EVENT ENQUEUE Focus "LocalizedControlType":"list item","Name":"John Doe","IsEnabled":true,"BoundingRectangle":{"IsEmpty":false,"Location":{"X":1603,"Y":740},"Size":{"IsEmpty":false,"Width":200,"Height":27},"X":1603,"Y":740,"Width":200,"Height":27,"Left":1603,"Top":740,"Right":1803,"Bottom":767,"TopLeft":{"X":1603,"Y":740},"TopRight":{"X":1803,"Y":740},"BottomLeft":{"X":1603,"Y":767},"BottomRight":{"X":1803,"Y":767}},"IsControlElement":true,"IsContentElement":true
                         *
                         * For this reason, only look at LocalizedControlType and Name
                         */
                        // TBD: Could use reis?
                        var sHash = UIATools.GetControlTypeShortName(aei.ControlType) + ":" + repr(aei.Name);

                        if (sHash == this.lastEventKey)
                        {
                            //hostApp.Log(LogType.DBG, "Ignoring duplicate event {0} {1}", UIATools.GetEventShortName(e), UIATools.AEToString(ae));
                            hostApp.Log(LogType.DBG, "Ignoring duplicate event");
                            return;
                        }
                        this.lastEventKey = sHash;

                        // Unexpected/unwanted dialog box squasher - Prompting for audio, host wants you to unmute, etc.
                        if ((aei.ControlType == ControlType.Window) && (aei.ClassName == "zChangeNameWndClass") && (!Controller.bWaitingForChangeNameDialog))
                        {
                            hostApp.Log(LogType.WRN, "Closing unexpected dialog 0x{0:X8} {1}", (uint)aei.NativeWindowHandle, UIATools.AEToString(ae));
                            WindowTools.CloseWindow((IntPtr)aei.NativeWindowHandle);

                            return;
                        }

                        // Hunt for first chat message
                        if ((aeCurrentChatMessageListItem == null) && (sChatListRid != null) && (aei.ControlType == ControlType.ListItem))
                        {
                            var thisItemRid = UIATools.AERuntimeIDToString(ae.GetRuntimeId());
                            if (thisItemRid.StartsWith(sChatListRid))
                            {
                                hostApp.Log(LogType.DBG, "Chat < Got first message!");
                                aeCurrentChatMessageListItem = ae;

                                // EXPERIMENTAL -- Unload event handler!
                                Stop(true);

                                return;
                            }
                        }

                        if (!this.isWatching)
                        {
                            hostApp.Log(LogType.DBG, "Ignoring event {0} (not watching)", sHash);
                            return;
                        }

                        var evt = new Controller.UIEvent
                        {
                            e = e.EventId,
                            ae = ae,
                            aei = aei,
                        };

                        this.eventQueue.Enqueue(evt);
                        this.eventWaitHandle.Set();

                        hostApp.Log(LogType.DBG, "*** EVENT ENQUEUE {0} {1}", UIATools.GetEventShortName(e), UIATools.AEToString(ae));
                    }
                }
                catch (Exception ex)
                {
                    hostApp.Log(LogType.ERR, "OnFocusChangedEvent - Unhandled Exception: {0}", ex.ToString());
                }
            }

            /// <summary>
            /// Waits for an event with a timeout.  Returns null if the timeout is reached.
            /// </summary>
            public UIEvent WaitEvent(int millisecondsTimeout)
            {
                UIEvent ret = null;
                if (!eventWaitHandle.WaitOne(millisecondsTimeout))
                {
                    return ret;
                }

                lock (queueLock)
                {
                    ret = eventQueue.Dequeue();
                    if (eventQueue.Count == 0)
                    {
                        eventWaitHandle.Reset();
                    }

                    return ret;
                }
            }

            /// <summary>
            /// Removes and returns the first UIEvent from the queue, or null if there are none.
            /// </summary>
            public UIEvent TryGetNextEvent()
            {
                return WaitEvent(0);
            }

            private void InternalClear()
            {
                eventQueue = new Queue<UIEvent>();
                lastEventKey = null;
                eventWaitHandle.Reset();
            }

            /// <summary>
            /// Clears the event queue.
            /// </summary>
            public void Clear()
            {
                lock (queueLock)
                {
                    InternalClear();
                }
            }

            /// <summary>
            /// Start listening for events.  Since hooking/unhooking an event handler takes time, simply clear the queue and set/clear a flag.
            /// </summary>
            public void Start()
            {
                lock (queueLock)
                {
                    if (!isEventHandlerAdded) {
                        hostApp.Log(LogType.WRN, "EventWatcher.Start() called with no event handler in place.");
                    }

                    InternalClear();
                    isWatching = true;
                }
            }

            /// <summary>
            /// Start listening for events.  Since hooking/unhooking an event handler takes time, simply clear the queue and set/clear a flag.
            /// </summary>
            public void Stop(bool force = false)
            {
                lock (queueLock)
                {
                    InternalClear();
                    if (force)
                    {
                        try
                        {
                            if (isEventHandlerAdded)
                            {
                                hostApp.Log(LogType.DBG, "Removing AE focus event handler");
                                System.Windows.Automation.Automation.RemoveAutomationFocusChangedEventHandler(this.OnFocusChangedEvent);
                                isEventHandlerAdded = false;
                            }
                        }
                        catch (Exception ex)
                        {
                            hostApp.Log(LogType.WRN, "Failed to remove AE focus event handler: {0}", repr(ex.ToString()));
                        }
                    }
                    isWatching = false;
                }
            }

            public void Dispose() => Dispose(true);

            protected virtual void Dispose(bool disposing)
            {
                if (isDisposed)
                {
                    return;
                }

                if (disposing)
                {
                    /*
                    lock (queueLock)
                    {
                        InternalClear();
                        Automation.RemoveAutomationFocusChangedEventHandler(OnFocusChangedEvent);
                    }
                    */
                    Stop(true);
                }

                isDisposed = true;
            }
        }

        public static class SpecialRecipient
        {
            public static string EveryoneInMeeting = "Everyone (in Meeting)";
            public static string EveryoneInWaitingRoom = "Everyone (in Waiting Room)";

            /// <summary>This method returns true if the given recipient is one of the special Everyone options; false otherwise.</summary>
            public static bool IsEveryone(string sRecipient)
            {
                return sRecipient.StartsWith("Everyone");
            }

            /// <summary>When there is nobody in the waiting room, the "Everyone (in Meeting)" selection item is renamed to "Everyone".  This
            /// method normalizes the value.</summary>
            public static string Normalize(string sRecipient)
            {
                return sRecipient == "Everyone" ? EveryoneInMeeting : sRecipient;
            }
        }

        public enum ParticipantListType
        {
            Detect = -1,
            Waiting = 0,
            Attending = 2,
        }

        public enum ParticipantStatus
        {
            Waiting = 0,
            Joining = 1,
            Attending = 2,
            Leaving = 3,
        }

        public enum ParticipantRole
        {
            None = 0,
            Host = 1,
            CoHost = 2,
        }

        public enum ParticipantAudioDevice
        {
            Unknown = 0,
            Computer = 1,
            Telephone = 2,
        }

        public enum ParticipantAudioStatus
        {
            Disconnected = 0,
            Unmuted = 1,
            Muted = 2,
        }

        public enum ParticipantVideoStatus
        {
            Disconnected = 0,
            On = 1,
            Off = 2,
        }

        public class Participant
        {
            public string name = null;
            public ParticipantStatus status = ParticipantStatus.Waiting;
            public ParticipantRole role = ParticipantRole.None;
            public ParticipantAudioDevice device = ParticipantAudioDevice.Unknown;
            public ParticipantAudioStatus audioStatus = ParticipantAudioStatus.Disconnected;
            public ParticipantVideoStatus videoStatus = ParticipantVideoStatus.Disconnected;
            public bool isMe = false;
            public DateTime dtWaiting = DateTime.MinValue;
            public DateTime dtAttending = DateTime.MinValue;
            internal AutomationElement _ae;
            public bool isSharing = false;
        }

        public static Dictionary<string, Participant> participants = new Dictionary<string, Participant>();
        public static Participant me = null;

        // v5.2 public static Regex reChatMessage = new Regex(@"^From (.+) to (.+):\s*(\(Privately\)|)\s*(\d+:\d+ [AP]M)\r\n(.*)$", RegexOptions.Compiled);
        private static Regex reChatMessage = new Regex(@"^From (.+) to (.+):\s*(\(Direct Message\)|)\s*(\d+:\d+ [AP]M)\r\n(.*)$", RegexOptions.Compiled | RegexOptions.Singleline);
        private static Regex reChatRecipient = new Regex(@"^To: (.+)$", RegexOptions.Compiled);
        private static string sLastChatMessage = null;

        private static IntPtr hZoomMainWindow = IntPtr.Zero;
        private static int nZoomPID = 0;
        private static AutomationElement aeZoomMainWindow = null;

        public class ParticipantEventArgs : EventArgs
        {
            public Participant participant;
        }

        public static event EventHandler<ParticipantEventArgs> ParticipantAttendanceStatusChange = (sender, e) => { };

        public class ChatEventArgs : EventArgs
        {
            //private readonly DateTime dt = DateTime.MinValue;
            public string from = null;
            public string to = null;
            public string text = null;
            public bool isPrivate = false;
        }

        public static event EventHandler<ChatEventArgs> ChatMessageReceive = (sender, e) => { };

        private static bool bFirstChatUpdate = true;

        private class UIEvent
        {
            public AutomationEvent e;
            public AutomationElement ae;
            public AutomationElement.AutomationElementInformation aei;
        }

        private static Rectangle chatRect = new Rectangle(0, 0, 0, 0);
        private static Rectangle partRect = new Rectangle(0, 0, 0, 0);
        private static Rectangle zoomRect = new Rectangle(0, 0, 0, 0);
        public static Rectangle AppRect = new Rectangle(0, 0, 0, 0);

        //private static readonly Regex reWaitingPanelMsg = new Regex(@"^(\d+) (?:person|people) (?:is|are) waiting", RegexOptions.Compiled);
        //private static readonly Regex reParticipantPanelMsg = new Regex(@"^(\d+) (?:participant|participants) in the meeting", RegexOptions.Compiled);

        private static readonly string SZoomLoginWindowClass = "ZPFTEWndClass";
        private static readonly Regex ReZoomLoginWindowTitle = new Regex(@"^Zoom Cloud Meetings$", RegexOptions.Compiled);

        private static readonly string SZoomMenuWindowClass = "ZPPTMainFrmWndClassEx";
        private static readonly Regex ReZoomMenuWindowTitle = new Regex(@"^Zoom$", RegexOptions.Compiled);

        private static readonly string SZoomJoinWindowClass = "zWaitHostWndClass";
        private static readonly Regex ReZoomJoinWindowTitle = new Regex(@"^Zoom$", RegexOptions.Compiled);

        private static readonly string SZoomMeetingClass = "ZPContentViewWndClass";
        private static readonly Regex ReZoomMeetingWindow = new Regex(@"^Zoom Meeting", RegexOptions.Compiled);

        // Chat window title can be "Chat" or "Zoom Group Chat" based on client version
        private static readonly string SZoomChatWindowClass = "ZPConfChatWndClass";
        private static readonly Regex ReZoomChatWindowTitle = new Regex(@"^(?:Zoom Group |)Chat$", RegexOptions.Compiled);

        private static readonly string SZoomParticipantsWindowClass = "zPlistWndClass";
        private static readonly Regex ReZoomParticipantsWindowTitle = new Regex(@"^Participants \((\d+)\)$", RegexOptions.Compiled);

        private static readonly string SZoomRenameWindowClass = "zChangeNameWndClass";
        private static readonly Regex ReZoomRenameWindowTitle = new Regex(@"^Rename$", RegexOptions.Compiled);

        private static readonly string SZoomConfirmCoHostWindowClass = "zChangeNameWndClass";
        private static readonly Regex ReZoomConfirmCoHostWindowTitle = new Regex(@"^Zoom$", RegexOptions.Compiled);
        private static readonly Regex ReZoomConfirmCoHostPrompt = new Regex(@"^Zoom,Do you want to make (.+) the co-host of this meeting\?,$", RegexOptions.Compiled);
        private static readonly Regex ReZoomConfirmHostPrompt = new Regex(@"^Zoom,Do you want to change the host to (.+)\?,$", RegexOptions.Compiled);

        private static readonly string SChromeZoomWindowClass = "Chrome_WidgetWin_1";
        private static readonly Regex ReChromeZoomWindowTitle = new Regex(@" \- Zoom \- Google Chrome$", RegexOptions.Compiled);

        private static readonly string SJoinAudioWindowClass = "zJoinAudioWndClass";

        private static readonly string SChangeNameWindowClass = "zChangeNameWndClass";
        private static readonly Regex ReChangeNameWindowTitle = new Regex(@"^Zoom$", RegexOptions.Compiled);

        private static readonly string SZoomConfirmMuteAllWindowClass = "zChangeNameWndClass";
        private static readonly Regex ReZoomConfirmMuteAllWindowTitle = new Regex(@"^Mute All$", RegexOptions.Compiled);

        private static readonly string SZoomPopUpMenuWindowClass = "WCN_ModelessWnd";
        private static readonly Regex ReZoomPopUpMenuTitle = new Regex(@"^$", RegexOptions.Compiled);

        private static readonly Regex ReAnyWindowTitle = new Regex(@".*", RegexOptions.Compiled);

        public static volatile bool bWaitingForChangeNameDialog = false;

        public class ZoomClosedException : Exception
        {
            public ZoomClosedException() { }

            public ZoomClosedException(string message)
                : base(message) { }

            public ZoomClosedException(string message, Exception inner)
                : base(message, inner) { }
        }

        private static IntPtr WaitZoomMeetingWindow(out string sWindowTitle, int timeout = 10000, int poll = 250)
        {
            sWindowTitle = string.Empty;

            try
            {
                return WindowTools.WaitWindow(SZoomMeetingClass, ReZoomMeetingWindow, out sWindowTitle, timeout, poll);
            }
            catch (TimeoutException ex)
            {
                hostApp.Log(LogType.WRN, "Zoom was closed; Killing off any hung Zoom processes");
                Kill();

                throw new ZoomClosedException("Zoom Closed", ex);
            }
        }

        public static IntPtr GetZoomMeetingWindowHandle()
        {
            return WaitZoomMeetingWindow(out _, 0);
        }

        public static IntPtr GetChatPanelWindowHandle()
        {
            IntPtr hWnd = WindowTools.FindWindow(SZoomChatWindowClass, ReZoomChatWindowTitle, out _); // ZPConfChatWndClass
            WindowTools.ShowWindow(hWnd, (uint)WindowTools.ShowCmd.SW_RESTORE); // TBD: Check if it needs to be shown or not
            if (hWnd == IntPtr.Zero)
            {
                hostApp.Log(LogType.INF, "GetChatPanelWindowHandle: Opening Panel");
                WindowTools.SendKeys(GetZoomMeetingWindowHandle(), "%h");

                hWnd = WindowTools.FindWindow(SZoomChatWindowClass, ReZoomChatWindowTitle, out _);
            }

            //WindowTools.FocusWindow(hWnd);
            return hWnd;
        }

        public static IntPtr GetParticipantsPanelWindowHandle()
        {
            IntPtr hWnd = WindowTools.FindWindow(SZoomParticipantsWindowClass, ReZoomParticipantsWindowTitle, out _);
            WindowTools.ShowWindow(hWnd, (uint)WindowTools.ShowCmd.SW_RESTORE); // TBD: Check if it needs to be shown or not
            if (hWnd == IntPtr.Zero)
            {
                hostApp.Log(LogType.INF, "GetParticipantsPanelWindowHandle: Opening Panel");
                WindowTools.SendKeys(GetZoomMeetingWindowHandle(), "%u");

                hWnd = WindowTools.FindWindow(SZoomParticipantsWindowClass, ReZoomParticipantsWindowTitle, out _);
            }

            // TBD: Could do something about the title changing; Use it to update cache?
            //WindowTools.FocusWindow(hWnd);
            return hWnd;
        }

        // CAUTION: This Regex is suceptible to attack by the attendee renaming themselves with ",(Co-host)" at the end, for example
        //private static readonly Regex ReAttendingParticipant = new Regex(@"^(.*?),(?:\(((?:Host|Co\-host|Me)[^\)]*)\),|) (No Audio Connected|(?:Telephone|Computer audio) (?:un|)muted),(No Video Connected|Video (?:on|off))(?:,(.*)|)$");
        private static readonly Regex ReWaitingListStart = new Regex(@"^Waiting Room \(\d+\), (?:expanded|collapsed)$");
        private static readonly Regex ReAttendingListStart = new Regex(@"^In the Meeting \(\d+\), (?:expanded|collapsed)$");
        private static readonly Regex ReAttendingParticipant = new Regex(@"^(.*?),(?:\(((?:Host|Co\-host|Me)[^\)]*)\),|) (?:(Screen sharing), |)(No Audio Connected|(?:Telephone|Computer audio) (?:un|)muted),(No Video Connected|Video (?:on|off))(?:,(.*)|)$");
        private static readonly string SJoiningSuffix = "Joining...";

        private static IntPtr hWndParticipants = IntPtr.Zero;
        private static AutomationElement aeParticipantsWindow = null;
        private static AutomationElement aeParticipantsList = null;

        private static string sChatListRid = null;
        private static AutomationElement aeChatWindow = null;
        private static AutomationElement aeChatMessageList = null;
        private static AutomationElement aeChatInputControl = null;
        private static AutomationElement aeChatRecipientControl = null;

        private static AutomationElement aeAllParticipantsPane = null;

        private static Participant GetParticipantFromListItem(ref ParticipantListType listType, AutomationElement listItem)
        {
            string[] a;

            Participant p = new Participant
            {
                _ae = listItem,
            };
            var sValue = listItem.Cached.Name;

            if (ReWaitingListStart.IsMatch(sValue) || ReAttendingListStart.IsMatch(sValue))
            {
                // This is the beginning of the waiting or attending list; Skip it
                return null;
            }

            var m = ReAttendingParticipant.Match(sValue);
            if (!m.Success)
            {
                // This item does not match the Attending participant format, so it must be a Waiting participant

                /*
                 * TBD: Rationalize this format somehow?
                if (!m.Success)
                {
                    hostApp.Log(LogType.WRN, "Got ListItem that cannot be parsed: {0}", repr(sValue));
                    return null;
                }
                */

                // WAITING
                //   'LocalizedControlType' = 'list item','Name' = 'Reagan'
                //     'LocalizedControlType' = 'text','Name' = 'Reagan'
                if (sValue.EndsWith(SJoiningSuffix))
                {
                    // If they're joining, "Joining..." will be appended to the end of the name
                    p.status = ParticipantStatus.Joining;
                    p.name = sValue.Substring(0, sValue.Length - SJoiningSuffix.Length);
                }
                else
                {
                    p.status = ParticipantStatus.Waiting;
                    p.name = sValue;
                }

                if (p.name.Contains("****"))
                {
                    p.device = ParticipantAudioDevice.Telephone; // NOTE: This only works if masking is enabled
                }
                //ret["state"] = (m.Groups[2].Value.Trim() == "Joining...") ? "Joining" : "Waiting";
                return p;
            }

            // ATTENDING
            //   'LocalizedControlType' = 'list item','Name' = 'Crispy Chris,(Co-host, me), Computer audio muted,Video on' Text: 'Crispy Chris,(Co-host, me), Computer audio muted,Video on'
            //       'LocalizedControlType' = 'text','Name' = 'Crispy Chris' Text: 'Crispy Chris'
            //       'LocalizedControlType' = 'text','Name' = '(Co-host, me)' Text: '(Co-host, me)'
            p.status = ParticipantStatus.Attending;

            var sName = m.Groups[1].Value;
            var sStatus = m.Groups[2].Value;
            var sSharingStatus = m.Groups[3].Value;
            var sAudioStatus = m.Groups[4].Value;
            var sVideoStatus = m.Groups[5].Value;
            //var sExtra = m.Groups[5].Value; // TBD: Parse for hand raised, etc.

            // Set Name
            p.name = sName;

            // Set Status - Example: (Host, me)
            // TBD: If it has a status component, could search subtree for text item that matches to validate.  This would mitigate the security issue with someone renaming themselves to include it
            if (sStatus.Length > 0)
            {
                a = sStatus.Split(',');
                foreach (string i in a)
                {
                    string x = i.Trim();
                    if (x == "Co-host")
                    {
                        p.role = ParticipantRole.CoHost;
                    }
                    else if (x == "Host")
                    {
                        p.role = ParticipantRole.Host;
                    }
                    else if ((x == "me") || (x == "Me"))
                    {
                        p.isMe = true;
                    }
                    else if (x.StartsWith("participant ID: "))
                    {
                        // Ignore this useless info
                    }
                    else
                    {
                        hostApp.Log(LogType.WRN, "GetParticipantFromListItem Status Element Value Unrecognized {0}", repr(i));
                    }
                }
            }

            // Set Sharing Status
            if (sSharingStatus.Length > 0)
            {
                p.isSharing = true;
            }

            // Set Audio Status
            if (sAudioStatus == "No Audio Connected")
            {
                p.audioStatus = ParticipantAudioStatus.Disconnected;
            }
            else
            {
                p.device = sAudioStatus.StartsWith("Computer audio ") ? ParticipantAudioDevice.Computer : ParticipantAudioDevice.Telephone;
                if (sAudioStatus.EndsWith(" muted"))
                {
                    p.audioStatus = ParticipantAudioStatus.Muted;
                }
                else if (sAudioStatus.EndsWith(" unmuted"))
                {
                    p.audioStatus = ParticipantAudioStatus.Unmuted;
                }
            }

            // Set video status
            if (sVideoStatus == "No Video Connected")
            {
                p.videoStatus = ParticipantVideoStatus.Disconnected;
            }
            else if (sVideoStatus == "Video on")
            {
                p.videoStatus = ParticipantVideoStatus.On;
            }
            else if (sVideoStatus == "Video off")
            {
                p.videoStatus = ParticipantVideoStatus.Off;
            }

            // Set me object if this is my participant entry
            if (p.isMe)
            {
                me = p;
            }

            return p;
        }

        private static void WalkParticipantList(AutomationElement aeList, out bool foundMe)
        {
            foundMe = false;
            ParticipantListType listType = ParticipantListType.Detect;

            if (aeList == null)
            {
                return;
            }

            int nProcessed = 0;
            var sw = Stopwatch.StartNew();
            hostApp.Log(LogType.DBG, "WalkParticipantList - Enter");

            try
            {
                // Cache the data we're interested in.  Normally, each element we retreive results in a client<->server round-trip.  This does it all in one pass.
                //   It also specifies which properties we want to retrieve instead of getting all of them
                CacheRequest cr = new CacheRequest();
                cr.Add(AutomationElement.NameProperty);
                cr.TreeScope = TreeScope.Element | TreeScope.Children;
                AutomationElement list;

                using (cr.Activate())
                {
                    // Load the list element and cache the specified properties for its descendants
                    list = aeList.FindFirst(TreeScope.Element, new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.List));
                }

                if (list == null)
                {
                    return;
                }

                //AutomationElement firstListItem = null;
                foreach (AutomationElement listItem in list.CachedChildren)
                {
                    // TBD: How to handle two participants with the same name? There doesn't seem to be any kind of unique ID or "tie breaker" available in Zoom's UIA
                    nProcessed++;
                    hostApp.Log(LogType.DBG, "| GetParticipantFromListItem {0} {1}", listType.ToString(), repr(listItem.Cached.Name));
                    Participant p = GetParticipantFromListItem(ref listType, listItem);

                    /*
                    if ((firstListItem == null) && (listType == ParticipantListType.Attending))
                    {
                        // This should be the first attendee in the list
                        firstListItem = listItem;
                    }
                    */

                    if (p == null)
                    {
                        continue;
                    }

                    // If this is the first time we're encountering this participant this pass, execute any queued actions now
                    if (!participants.ContainsKey(p.name) && ExecuteQueuedIndividualParticipantActions(p))
                    {
                        // One or more actions were successfully executed; Particpant's state likely changed, so update & re-parse
                        hostApp.Log(LogType.DBG, "Reparsing participant after successful action(s)");

                        AutomationElement updatedListItem = null;
                        try
                        {
                            updatedListItem = listItem.GetUpdatedCache(cr);
                        }
                        catch (Exception ex)
                        {
                            hostApp.Log(LogType.DBG, "Failed to re-process {0} {1}: {2}", listType.ToString(), repr(p.name), repr(ex.ToString()));
                            continue;
                        }

                        p = GetParticipantFromListItem(ref listType, updatedListItem);
                        if (p == null)
                        {
                            hostApp.Log(LogType.DBG, "Failed to re-process {0} {1}", listType.ToString(), repr(p.name));
                            continue;
                        }

                        hostApp.Log(LogType.DBG, "| GetParticipantFromListItem {0} {1} (REPARSE)", listType.ToString(), repr(updatedListItem.Cached.Name));
                    }

                    if (p.isMe)
                    {
                        foundMe = true;
                    }

                    participants[p.name] = p;
                }

                // Set focus on the top item in the current page of the user list by clicking on it.  This ensures that when we hit {PGUP} later that we scroll properly.
                //   X + 32 gets us about in the middle of the user's avatar which is a safe area to click.
                /*
                if (listType == ParticipantListType.Attending)
                {
                    if (firstListItem == null)
                    {
                        // This should never happen, but just in case!
                        hostApp.Log(LogType.WRN, "WalkParticipantList - First list item is null; Cannot click");
                    }
                    else
                    {
                        Rect rect = firstListItem.Current.BoundingRectangle;
                        // NOTE: When there are enough people in the meeting, the "Find a participant" input box appears, and we end up clicking on that
                        //   instead of the actual list item.  Compensate for this by clicking on the bottom of the listitem instead of the middle
                        // WindowTools.ClickOnPoint(IntPtr.Zero, new System.Drawing.Point((int)rect.X + 32, (int)(rect.Y + (rect.Height / 2))));
                        WindowTools.ClickOnPoint(IntPtr.Zero, new System.Drawing.Point((int)rect.X + 32, (int)(rect.Y + rect.Height - 1)));
                    }
                }
                */
            }
            finally
            {
                sw.Stop();
                hostApp.Log(LogType.DBG, "WalkParticipantList - Exit - Processed {0} item(s) in {1:0.000}s", nProcessed, sw.Elapsed.TotalSeconds);
            }
        }

        private static void SelectChatRecipient(string sTargetRecipient)
        {
            int nAttempt = 0;

            while (true)
            {
                try
                {
                    nAttempt++;

                    var sCurrentRecipient = aeChatRecipientControl.Current.Name;

                    var m = reChatRecipient.Match(sCurrentRecipient);

                    // This should never happen, but check anyway
                    if (!m.Success)
                    {
                        throw new Exception(string.Format("Failed to parse current recipient value {0}", repr(sCurrentRecipient)));
                    }

                    sCurrentRecipient = SpecialRecipient.Normalize(m.Groups[1].Value);

                    // If the desired recipient is already selected, then there's nothing to do!
                    if (sCurrentRecipient == sTargetRecipient)
                    {
                        return;
                    }

                    // If we tried 3 times unsuccessfully, bail!
                    if (nAttempt == 4)
                    {
                        break;
                    }

                    hostApp.Log(LogType.DBG, "(Attempt #{0}) Changing chat recipient from {1} to {2}", nAttempt, repr(sCurrentRecipient), repr(sTargetRecipient));

                    WindowTools.ClickMiddle(IntPtr.Zero, aeChatRecipientControl.Current.BoundingRectangle);

                    // When we click the chat "To: " Split Button, it opens a new window with a list of participants.
                    //   Get an AE for the list, get all of the List Items in it (list of participants), then select
                    //   and click the one we want.
                    try
                    {
                        // TBD: What does it look like when there's a text box for searching?
                        // Window "ZPConfChatUserWndClass":"" > Pane "" > List "" > List Item 1, List Item 2, ...
                        IntPtr hWnd = WindowTools.WaitWindow("ZPConfChatUserWndClass", string.Empty, 2000);
                        var aeSelectChatRecipientWindow = AutomationElement.FromHandle(hWnd);
                        var aeListParentPane = aeSelectChatRecipientWindow.FindFirst(TreeScope.Children, new AndCondition(
                            new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Pane),
                            new PropertyCondition(AutomationElement.NameProperty, string.Empty)));

                        var aeRecipientList = aeListParentPane.FindFirst(TreeScope.Children, new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.List));

                        // The "Everyone (in Meeting)" recipient is special as it can have two values based if there is anyone in the waiting room or not.
                        //   We compensate for that here
                        // NOTE: We know all children here are ListItem controls, so there's no need to check for it
                        var recipient = aeRecipientList.FindFirst(
                            TreeScope.Children,
                            (sTargetRecipient == SpecialRecipient.EveryoneInMeeting) ?
                                (Condition)new OrCondition(
                                    new PropertyCondition(AutomationElement.NameProperty, "Everyone"),
                                    new PropertyCondition(AutomationElement.NameProperty, sTargetRecipient)) :
                                (Condition)new PropertyCondition(AutomationElement.NameProperty, sTargetRecipient));

                        hostApp.Log(LogType.DBG, "SelectChatRecipient - About to select {0}", sTargetRecipient);
                        var sip = (SelectionItemPattern)recipient.GetCurrentPattern(SelectionItemPatternIdentifiers.Pattern);
                        sip.Select();

                        // CMM 2020.12.24 - This can unselect the current item rather than select it. Sending a SPACE seems to work better
                        //   This works even when the search box is present
                        WindowTools.ClickMiddle(IntPtr.Zero, recipient.Current.BoundingRectangle);
                        //WindowTools.SendKeys(hWnd, " "); - Doesn't work

                        // Loop back around to verify selection
                        continue;
                    }
                    catch (Exception ex)
                    {
                        // If we get an exception, just log it and try again
                        hostApp.Log(LogType.WRN, ex.ToString());
                        continue;
                    }
                }
                catch (Exception ex)
                {
                    hostApp.Log(LogType.WRN, "(Attempt {0}) Failed to select chat recipient {0}: {1}", nAttempt, repr(sTargetRecipient), repr(ex.ToString()));
                }
            }

            // Hit ESC to kill the menu
            WindowTools.SendKeys("{ESC}");

            throw new KeyNotFoundException(string.Format("Could not select chat participant {0}", repr(sTargetRecipient)));
        }

        private class OutboundChatMsg
        {
            public string to = null;
            public string msg = null;
            public int attempt = 0;
            public bool speak = false;
            public DateTime lastAttemptDT = DateTime.MinValue;
        }
        private static readonly int CHAT_MSG_MAX_ATTEMPTS = 3;
        private static readonly int CHAT_MSG_RETRY_DELAY_SECS = 10;

        private static readonly ConcurrentQueue<OutboundChatMsg> QSendMsgs = new ConcurrentQueue<OutboundChatMsg>();
        private static AutomationElement aeCurrentChatMessageListItem = null;
        private static int nMessageScanCount = 0;
        private static bool bFirstChatMessageScan = true;
        private static string sLastChatMessageText = null;
        //private static readonly char[] crDelim = new char[] { '\r' };

        public static void SendQueuedChatMessages()
        {
            // We can't do anything until we have our chat objects
            if ((aeChatWindow == null) || (aeChatInputControl == null))
            {
                return;
            }

            // Nothing to do if we have no messages to send
            if (QSendMsgs.Count == 0)
            {
                return;
            }

            var hWndChat = GetChatPanelWindowHandle();

            ConcurrentQueue<OutboundChatMsg> QRetry = new ConcurrentQueue<OutboundChatMsg>();

            // Send any queued up chat messages
            //   TBD: Could optimize by batching messages to the same participant together.  Can also send any that are to the currently selected recipient immediately
            while (QSendMsgs.TryDequeue(out OutboundChatMsg ocm))
            {
                try
                {
                    // If we've reached the max attempts, don't try again
                    if (ocm.attempt >= CHAT_MSG_MAX_ATTEMPTS)
                    {
                        hostApp.Log(LogType.ERR, "Chat > Giving up on message to {0}: {1}; Max attempts", repr(ocm.to), repr(ocm.msg));
                        continue;
                    }

                    // If it's too soon to send this message, then queue it up for later
                    var nowDT = DateTime.UtcNow;
                    if (nowDT < ocm.lastAttemptDT.AddSeconds(CHAT_MSG_RETRY_DELAY_SECS))
                    {
                        hostApp.Log(LogType.DBG, "Chat > Too soon to re-send {0}: {1}", repr(ocm.to), repr(ocm.msg));
                        QRetry.Enqueue(ocm);
                        continue;
                    }

                    ocm.lastAttemptDT = nowDT;
                    ocm.attempt++;

                    if ((ocm.to != SpecialRecipient.EveryoneInMeeting) && (ocm.to != SpecialRecipient.EveryoneInWaitingRoom))
                    {
                        // Unless we're sending to everyone in the meeting or the waiting room, make sure the participant is actually in the meeting

                        if (!participants.TryGetValue(ocm.to, out Participant p))
                        {
                            hostApp.Log(LogType.WRN, "Chat > Giving up on message to {0}: {1}; Paricipant is not in meeting", repr(ocm.to), repr(ocm.msg));
                            continue;
                        }

                        if (p.device == ParticipantAudioDevice.Telephone)
                        {
                            hostApp.Log(LogType.WRN, "Chat > Giving up on message to {0}: {1}; Paricipant is dial-in only", repr(ocm.to), repr(ocm.msg));
                            continue;
                        }

                        // If the participant isn't attending (ie: they are waiting, joining, etc.), throw an exception so we can try again later
                        if (p.status != ParticipantStatus.Attending)
                        {
                            throw new Exception(string.Format("Participant status should be \"Attending\", not {0}", repr(p.status.ToString())));
                        }
                    }

                    hostApp.Log(LogType.INF, "Chat > (Attempt #{0}) Sending message to {1}: {2}", repr(ocm.attempt), repr(ocm.to), repr(ocm.msg));

                    SelectChatRecipient(ocm.to);

                    // Select chat input box & type the message, then press {ENTER}

                    // 2020.08.10 - This seems to have become unreliable; Try clicking instead.
                    //((InvokePattern)aeChatInputControl.GetCurrentPattern(InvokePattern.Pattern)).Invoke();
                    WindowTools.ClickMiddle(IntPtr.Zero, aeChatInputControl.Current.BoundingRectangle);

                    //WindowTools.SendKeys(hWndChat, ocm.msg + "{ENTER}");

                    // Delete any left-over text.  This can happen when the previous chat msg's recipient was not found
                    WindowTools.SendKeys("^A{DEL}");

                    // Set the chat text.  This tries to use the clipboard if possible.
                    WindowTools.SendText(ocm.msg);

                    // Send the text
                    WindowTools.SendKeys("{ENTER}");

                    if (ocm.speak)
                    {
                        ocm.speak = false;
                        Sound.Speak(ocm.msg);
                    }
                }
                catch (Exception ex)
                {
                    hostApp.Log(LogType.WRN, "Chat > (Attempt #{0}) Failed to send message to {1}: {2}; {3}", repr(ocm.attempt), repr(ocm.to), repr(ocm.msg), repr(ex.ToString()));

                    /*
                    if (ocm.attempt >= 3)
                    {
                        hostApp.Log(LogType.ERR, "Giving up on message to {0}: {1}", repr(ocm.to), repr(ocm.msg));
                    }
                    else
                    {
                        QRetry.Enqueue(ocm);
                    }

                    continue;
                    */
                }

                // We requeue all messages - We only consider the send a success if we pick it up later in the chat history
                QRetry.Enqueue(ocm);
            }

            // Queue up any messages we need to retry
            while (QRetry.TryDequeue(out OutboundChatMsg ocm))
            {
                QSendMsgs.Enqueue(ocm);
            }
        }

        /// <summary>
        /// Converts the given a time string extracted to a DateTime object.  The time string must be in Zoom's chat message time format.
        /// </summary>
        /// <param name="s">Time string; Example values: "12:50 AM", "04:03 PM", "10:16 PM".</param>
        private static DateTime ParseChatMessageTimeValue(string s)
        {
            // TBD: Verify InvariantCulture is appropriate here -- ie: Does Zoom's format change?
            //   If it does, we could either pass CultureInfo, CurrentUICulture or just do DateTime.Parse() which will likely figure it out
            DateTime dt = DateTime.ParseExact(s, "hh:mm tt", CultureInfo.InvariantCulture);
            // Ex: If it's currently 12:00 AM, and the given timestamp is 11:59 PM, we need to subtract one day from the date value
            return (dt > DateTime.Now) ? dt.AddDays(-1) : dt;
        }

        /// <summary>
        /// Scans chat pane for new incoming chat messages.
        /// </summary>
        /// <returns>true if there are any new messages; false otherwise.</returns>
        public static bool UpdateChat()
        {
            var hWndChat = GetChatPanelWindowHandle();
            List<ChatEventArgs> ceas = new List<ChatEventArgs>();
            bool ret = false;

            if ((aeChatWindow == null) || (hWndChat != (IntPtr)aeChatWindow.Current.NativeWindowHandle))
            {
                aeChatWindow = AutomationElement.FromHandle(hWndChat);

                aeChatMessageList = aeChatWindow.FindFirst(
                    TreeScope.Subtree,
                    new AndCondition(
                        new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.List),
                        new PropertyCondition(AutomationElement.NameProperty, "chat text list")));
                sChatListRid = UIATools.AERuntimeIDToString(aeChatMessageList.GetRuntimeId());
                hostApp.Log(LogType.DBG, "Got List RID: {0}", sChatListRid);
                sChatListRid += ",-1,"; // Example: 42,12980066,4,0,1,0,-1,0 - Where the final "0" is the zero-based index of the message number

                aeChatRecipientControl = aeChatWindow.FindFirst(
                    TreeScope.Subtree,
                    new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.SplitButton));

                // TBD: Does the text of this control ever change? Maybe re-write based on tree position instead of doing Subtree search
                aeChatInputControl = aeChatWindow.FindFirst(
                    TreeScope.Subtree,
                    new AndCondition(
                        new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Edit),
                        new PropertyCondition(AutomationElement.NameProperty, "Input chat text Type message here…")));
            }

            nMessageScanCount++;

            if (aeCurrentChatMessageListItem == null)
            {
                if (nMessageScanCount == 1)
                {
                    hostApp.Log(LogType.DBG, "Chat < Scanning for pre-existing messages");
                }
                else if (nMessageScanCount == 4)
                {
                    // It can take up to 5-10 secs for the event to fire, and we poll every 5s.  4 gives a little play at 15s
                    hostApp.Log(LogType.DBG, "Chat < No pre-existing messages found");
                    bFirstChatMessageScan = false;
                }

                // Click where the first message will show up.  If there is a message here, it will trigger an event
                //   which the event handler will pick up and handle, setting aeCurrentChatMessageList item to that
                //   message's item
                var rect = aeChatMessageList.Current.BoundingRectangle;
                WindowTools.ClickOnPoint(IntPtr.Zero, new System.Windows.Point(rect.X + 10, rect.Y + 10));

                return ret;
            }

            // If we got here, we have a chat message object.  Use it to find it's siblings and parse them
            hostApp.Log(LogType.DBG, "Chat < Parsing messages");
            while (true)
            {
                var sCurrentValue = aeCurrentChatMessageListItem.Current.Name;

                var m = reChatMessage.Match(sCurrentValue);
                if (m.Success == false)
                {
                    throw new FormatException($"Chat < Unable to parse message {repr(sCurrentValue)}");
                }

                // Populate message components
                var sFrom = m.Groups[1].Value;
                var sTo = SpecialRecipient.Normalize(m.Groups[2].Value);
                var bIsPrivate = m.Groups[3].Value.Length != 0;
                var dtTime = ParseChatMessageTimeValue(m.Groups[4].Value);
                var sMsgText = m.Groups[5].Value.StripBlankLinesAndTrimSpace();
                string sNewMsgText; // = string.Empty;

                // Additional lines can be appended to the last chat message if they come from the same sender
                if (sLastChatMessageText != null)
                {
                    if (sMsgText == sLastChatMessageText)
                    {
                        // Message has not changed; Move on to see if there are new messages
                        sNewMsgText = string.Empty;
                    }
                    else if (sMsgText.StartsWith(sLastChatMessageText))
                    {
                        // New text has been added onto old message; Continue on down to parse it
                        sNewMsgText = sMsgText.Substring(sLastChatMessageText.Length).Trim();
                    }
                    else
                    {
                        // Text has completely changed - this should not happen!
                        throw new Exception($"Got new message {repr(sCurrentValue)} when last chat message text was not null -- Shouldn't happen?!");
                        //sNewMsgText = sMsgText.Trim();
                    }
                }
                else
                {
                    // Text has completely changed (this is a new message); Move down to parse it
                    sNewMsgText = sMsgText;
                }

                if (sNewMsgText.Length == 0)
                {
                    // Last message has not changed; See if there are any new ones
                    var aeNextChatMessageListItem = TreeWalker.ContentViewWalker.GetNextSibling(aeCurrentChatMessageListItem);

                    // No new messages; We're done!
                    if (aeNextChatMessageListItem == null)
                    {
                        break;
                    }

                    // Loop back around to parse it
                    aeCurrentChatMessageListItem = aeNextChatMessageListItem;
                    sLastChatMessageText = null;
                    continue;
                }

                if (bFirstChatMessageScan)
                {
                    hostApp.Log(LogType.DBG, "Chat < First scan; Skipping pre-existing message from {0} to {1}: {2}", repr(sFrom), repr(sTo), repr(sNewMsgText));
                }
                else
                {
                    // Remove any chat messages we've sent that we see show up in the chat history.  This confirms they were sent.
                    // NOTE: There could be some issues here:
                    //   * What if two messages are combined into one?
                    //   * What if messages get interleaved from multiple senders?
                    //   * What if Zoom munges some of the characters?
                    //   * Trailing, leading and blank lines are removed. If one message has one of these and a second doesn't; we won't detect the difference
                    //   * Repeated message in chat history are merged into one message when we parse -- might have been more than one OB msg
                    //   * I'm sure there is more!
                    // TBD: There has to be a more efficient way that draining and reloading the queue each time.
                    if (sFrom == "Me")
                    {
                        ConcurrentQueue<OutboundChatMsg> qRetry = new ConcurrentQueue<OutboundChatMsg>();
                        while (QSendMsgs.TryDequeue(out OutboundChatMsg ocm))
                        {
                            // We trim the message in inbound text for various reasons; Also trim the outgoing message for comparison
                            var chatMsgText = ocm.msg.StripBlankLinesAndTrimSpace();

                            hostApp.Log(LogType.DBG, $"CHAT TX CHK To {repr(sTo)}/{repr(ocm.to)} Msg {repr(sNewMsgText)}/{repr(chatMsgText)}");

                            // We use Contains here since inbound messages can be merged into one.  Ex: 1> MSG1; 2> MSG2; 3> MSG3 ; < MSG1\nMSG2\nMSG3
                            if ((ocm.to == sTo) && sNewMsgText.Contains(chatMsgText))
                            {
                                hostApp.Log(LogType.INF, $"Chat > Transmission confirmed after {repr(ocm.attempt)} attempt(s) for message to {repr(ocm.to)}: {repr(ocm.msg)}");

                                if (ocm.speak)
                                {
                                    ocm.speak = false;
                                    Sound.Speak(ocm.msg);
                                }
                            }
                            else
                            {
                                qRetry.Enqueue(ocm);
                            }
                        }

                        // Queue up any messages we need to retry
                        while (qRetry.TryDequeue(out OutboundChatMsg ocm))
                        {
                            QSendMsgs.Enqueue(ocm);
                        }
                    }

                    string[] lines = sNewMsgText.Split(CRLFDelim);
                    foreach (var line in lines)
                    {
                        // We've got messages!  Fire an events for them
                        hostApp.Log(LogType.DBG, "Chat < Firing event for message from {0} to {1}: {2}", repr(sFrom), repr(sTo), repr(line));

                        ChatMessageReceive(null, new ChatEventArgs
                        {
                            from = sFrom,
                            to = sTo,
                            isPrivate = bIsPrivate,
                            text = line,
                        });
                    }

                    ret = true;
                }

                sLastChatMessageText = sMsgText;
            }

            // Done processing messages
            bFirstChatMessageScan = false;

            hostApp.Log(LogType.DBG, $"Chat < Done parsing; new={ret}");
            return ret;
        }

        public static void SendChatMessage(string to, bool speak, string msg, params object[] values)
        {
            QSendMsgs.Enqueue(new OutboundChatMsg
            {
                to = to,
                msg = string.Format(msg, values),
                speak = speak,
            });
        }

        public static void SendChatMessage(string to, string msg, params object[] values)
        {
            SendChatMessage(to, false, msg, values);
        }

        // Returns True if participants were updated
        public static bool UpdateParticipants()
        {
            DateTime dtNow = DateTime.UtcNow;
            hWndParticipants = GetParticipantsPanelWindowHandle();

            // TBD: This isn't really the right spot for this ...
            // Dialog squasher: Close any unexpected dialog boxes
            var hUnexpectedDialog = WindowTools.FindWindowByClass(Controller.SZoomRenameWindowClass);
            if (hUnexpectedDialog != IntPtr.Zero)
            {
                var ae = AutomationElement.FromHandle(hUnexpectedDialog);
                var sTree = string.Empty;
                if (ae != null)
                {
                    sTree = UIATools.WalkRawElementsToString(ae, true);
                }

                hostApp.Log(LogType.WRN, "Closing unexpected dialog 0x{0:X8}; AETree {1}", (uint)hUnexpectedDialog, repr(sTree));
                WindowTools.CloseWindow(hUnexpectedDialog);
            }

            if ((aeParticipantsWindow == null) || (hWndParticipants != (IntPtr)aeParticipantsWindow.Current.NativeWindowHandle))
            {
                hostApp.Log(LogType.DBG, "UpdateParticipants - ae == null : {0}", aeParticipantsWindow == null);

                // 'LocalizedControlType' = 'list','Name' = 'Waiting Room list, use arrow key to navigate, and press tab for more options' Text: 'Waiting Room list, use arrow key to navigate, and press tab for more options'
                //   'LocalizedControlType' = 'list item','Name' = 'Reagan' Text: 'Reagan'
                //     'LocalizedControlType' = 'text','Name' = 'Reagan' Text: 'Reagan'
                // 'LocalizedControlType'='list','Name'='participant list, use arrow key to navigate, and press tab for more options' Text:'participant list, use arrow key to navigate, and press tab for more options'
                //   'LocalizedControlType' = 'list item','Name' = 'Crispy Chris,(Co-host, me), Computer audio muted,Video on' Text: 'Crispy Chris,(Co-host, me), Computer audio muted,Video on'
                //       'LocalizedControlType' = 'text','Name' = 'Crispy Chris' Text: 'Crispy Chris'
                //       'LocalizedControlType' = 'text','Name' = '(Co-host, me)' Text: '(Co-host, me)'
                //      (IF SELECTED)
                //     'LocalizedControlType' = 'button','Name' = 'Unmute' Text: 'Unmute'
                // 'LocalizedControlType' = 'split button','Name' = 'More options for Crispy Chris'
                // 'LocalizedControlType'='button','Name'='Mute All (Alt+M)' Text:'Mute All (Alt+M)'
                // 'LocalizedControlType'='split button','Name'='More options to manage all participants'

                // New One: Someone,(Co-host), screen sharing, Telephone muted,Video on

                /* UI Tree for Participants Window:
                    properties="ControlType":{"LocalizedControlType":"window","Id":50032,"ProgrammaticName":"ControlType.Window"},"LocalizedControlType":"window","Name":"Participants (4)","IsKeyboardFocusable":true,"IsEnabled":true,"BoundingRectangle":{"IsEmpty":false,"Location":{"X":-749,"Y":0},"Size":{"IsEmpty":false,"Width":400,"Height":1047},"X":-749,"Y":0,"Width":400,"Height":1047,"Left":-749,"Top":0,"Right":-349,"Bottom":1047,"TopLeft":{"X":-749,"Y":0},"TopRight":{"X":-349,"Y":0},"BottomLeft":{"X":-749,"Y":1047},"BottomRight":{"X":-349,"Y":1047}},"IsControlElement":true,"IsContentElement":true,"ClassName":"zPlistWndClass","NativeWindowHandle":17436438,"ProcessId":7768,"Orientation":0,"FrameworkId":"Win32" supportedPatterns=["WindowPatternIdentifiers.Pattern","TransformPatternIdentifiers.Pattern"]
                        properties="ControlType":{"LocalizedControlType":"title bar","Id":50037,"ProgrammaticName":"ControlType.TitleBar"},"LocalizedControlType":"title bar","Name":"Participants (4)","IsEnabled":true,"BoundingRectangle":{"IsEmpty":false,"Location":{"X":-741,"Y":8},"Size":{"IsEmpty":false,"Width":384,"Height":23},"X":-741,"Y":8,"Width":384,"Height":23,"Left":-741,"Top":8,"Right":-357,"Bottom":31,"TopLeft":{"X":-741,"Y":8},"TopRight":{"X":-357,"Y":8},"BottomLeft":{"X":-741,"Y":31},"BottomRight":{"X":-357,"Y":31}},"IsControlElement":true,"AutomationId":"TitleBar","ProcessId":7768,"Orientation":0,"FrameworkId":"Win32" supportedPatterns=[]
                            properties="ControlType":{"LocalizedControlType":"menu bar","Id":50010,"ProgrammaticName":"ControlType.MenuBar"},"LocalizedControlType":"menu bar","Name":"System Menu Bar","IsKeyboardFocusable":true,"IsEnabled":true,"BoundingRectangle":{"IsEmpty":false,"Location":{"X":-741,"Y":8},"Size":{"IsEmpty":false,"Width":22,"Height":22},"X":-741,"Y":8,"Width":22,"Height":22,"Left":-741,"Top":8,"Right":-719,"Bottom":30,"TopLeft":{"X":-741,"Y":8},"TopRight":{"X":-719,"Y":8},"BottomLeft":{"X":-741,"Y":30},"BottomRight":{"X":-719,"Y":30}},"IsControlElement":true,"AutomationId":"SystemMenuBar","ProcessId":7768,"Orientation":0,"FrameworkId":"Win32" supportedPatterns=[]
                                properties="ControlType":{"LocalizedControlType":"menu item","Id":50011,"ProgrammaticName":"ControlType.MenuItem"},"LocalizedControlType":"menu item","Name":"System","AccessKey":"Alt \u002B Space","IsKeyboardFocusable":true,"IsEnabled":true,"BoundingRectangle":{"IsEmpty":false,"Location":{"X":-741,"Y":8},"Size":{"IsEmpty":false,"Width":22,"Height":22},"X":-741,"Y":8,"Width":22,"Height":22,"Left":-741,"Top":8,"Right":-719,"Bottom":30,"TopLeft":{"X":-741,"Y":8},"TopRight":{"X":-719,"Y":8},"BottomLeft":{"X":-741,"Y":30},"BottomRight":{"X":-719,"Y":30}},"IsControlElement":true,"IsContentElement":true,"AutomationId":"Item 1","ProcessId":7768,"Orientation":0,"FrameworkId":"Win32" supportedPatterns=["ExpandCollapsePatternIdentifiers.Pattern"]
                            properties="ControlType":{"LocalizedControlType":"button","Id":50000,"ProgrammaticName":"ControlType.Button"},"LocalizedControlType":"button","Name":"Minimize","IsEnabled":true,"BoundingRectangle":{"IsEmpty":false,"Location":{"X":-496,"Y":1},"Size":{"IsEmpty":false,"Width":47,"Height":30},"X":-496,"Y":1,"Width":47,"Height":30,"Left":-496,"Top":1,"Right":-449,"Bottom":31,"TopLeft":{"X":-496,"Y":1},"TopRight":{"X":-449,"Y":1},"BottomLeft":{"X":-496,"Y":31},"BottomRight":{"X":-449,"Y":31}},"IsControlElement":true,"AutomationId":"Minimize","ProcessId":7768,"Orientation":0,"FrameworkId":"Win32" supportedPatterns=["InvokePatternIdentifiers.Pattern"]
                            properties="ControlType":{"LocalizedControlType":"button","Id":50000,"ProgrammaticName":"ControlType.Button"},"LocalizedControlType":"button","Name":"Maximize","IsEnabled":true,"BoundingRectangle":{"IsEmpty":false,"Location":{"X":-449,"Y":1},"Size":{"IsEmpty":false,"Width":46,"Height":30},"X":-449,"Y":1,"Width":46,"Height":30,"Left":-449,"Top":1,"Right":-403,"Bottom":31,"TopLeft":{"X":-449,"Y":1},"TopRight":{"X":-403,"Y":1},"BottomLeft":{"X":-449,"Y":31},"BottomRight":{"X":-403,"Y":31}},"IsControlElement":true,"AutomationId":"Maximize","ProcessId":7768,"Orientation":0,"FrameworkId":"Win32" supportedPatterns=["InvokePatternIdentifiers.Pattern"]
                            properties="ControlType":{"LocalizedControlType":"button","Id":50000,"ProgrammaticName":"ControlType.Button"},"LocalizedControlType":"button","Name":"Close","IsEnabled":true,"BoundingRectangle":{"IsEmpty":false,"Location":{"X":-403,"Y":1},"Size":{"IsEmpty":false,"Width":47,"Height":30},"X":-403,"Y":1,"Width":47,"Height":30,"Left":-403,"Top":1,"Right":-356,"Bottom":31,"TopLeft":{"X":-403,"Y":1},"TopRight":{"X":-356,"Y":1},"BottomLeft":{"X":-403,"Y":31},"BottomRight":{"X":-356,"Y":31}},"IsControlElement":true,"AutomationId":"Close","ProcessId":7768,"Orientation":0,"FrameworkId":"Win32" supportedPatterns=["InvokePatternIdentifiers.Pattern"]
                        properties="ControlType":{"LocalizedControlType":"pane","Id":50033,"ProgrammaticName":"ControlType.Pane"},"LocalizedControlType":"pane","IsEnabled":true,"BoundingRectangle":{"IsEmpty":false,"Location":{"X":-741,"Y":31},"Size":{"IsEmpty":false,"Width":384,"Height":1008},"X":-741,"Y":31,"Width":384,"Height":1008,"Left":-741,"Top":31,"Right":-357,"Bottom":1039,"TopLeft":{"X":-741,"Y":31},"TopRight":{"X":-357,"Y":31},"BottomLeft":{"X":-741,"Y":1039},"BottomRight":{"X":-357,"Y":1039}},"IsControlElement":true,"IsContentElement":true,"ProcessId":7768,"Orientation":0 supportedPatterns=[]
                            properties="ControlType":{"LocalizedControlType":"pane","Id":50033,"ProgrammaticName":"ControlType.Pane"},"LocalizedControlType":"pane","IsEnabled":true,"BoundingRectangle":{"IsEmpty":false,"Location":{"X":-741,"Y":31},"Size":{"IsEmpty":false,"Width":384,"Height":1008},"X":-741,"Y":31,"Width":384,"Height":1008,"Left":-741,"Top":31,"Right":-357,"Bottom":1039,"TopLeft":{"X":-741,"Y":31},"TopRight":{"X":-357,"Y":31},"BottomLeft":{"X":-741,"Y":1039},"BottomRight":{"X":-357,"Y":1039}},"IsControlElement":true,"IsContentElement":true,"ProcessId":7768,"Orientation":0 supportedPatterns=[]
                                properties="ControlType":{"LocalizedControlType":"list view","Id":50008,"ProgrammaticName":"ControlType.List"},"LocalizedControlType":"list","Name":"participant list, use arrow key to navigate, and press tab for more options","IsEnabled":true,"BoundingRectangle":{"IsEmpty":false,"Location":{"X":-741,"Y":31},"Size":{"IsEmpty":false,"Width":384,"Height":881},"X":-741,"Y":31,"Width":384,"Height":881,"Left":-741,"Top":31,"Right":-357,"Bottom":912,"TopLeft":{"X":-741,"Y":31},"TopRight":{"X":-357,"Y":31},"BottomLeft":{"X":-741,"Y":912},"BottomRight":{"X":-357,"Y":912}},"IsControlElement":true,"IsContentElement":true,"ProcessId":7768,"Orientation":0 supportedPatterns=["SelectionPatternIdentifiers.Pattern"]
                                    properties="ControlType":{"LocalizedControlType":"list item","Id":50007,"ProgrammaticName":"ControlType.ListItem"},"LocalizedControlType":"list item","Name":"Crispy Chris,(Co-host, me), Computer audio unmuted,Video on","HasKeyboardFocus":true,"IsEnabled":true,"BoundingRectangle":{"IsEmpty":false,"Location":{"X":-739,"Y":33},"Size":{"IsEmpty":false,"Width":380,"Height":41},"X":-739,"Y":33,"Width":380,"Height":41,"Left":-739,"Top":33,"Right":-359,"Bottom":74,"TopLeft":{"X":-739,"Y":33},"TopRight":{"X":-359,"Y":33},"BottomLeft":{"X":-739,"Y":74},"BottomRight":{"X":-359,"Y":74}},"IsControlElement":true,"IsContentElement":true,"ProcessId":7768,"Orientation":0 supportedPatterns=["SelectionItemPatternIdentifiers.Pattern"]
                                        properties="ControlType":{"LocalizedControlType":"pane","Id":50033,"ProgrammaticName":"ControlType.Pane"},"LocalizedControlType":"pane","IsEnabled":true,"BoundingRectangle":{"IsEmpty":false,"Location":{"X":-739,"Y":33},"Size":{"IsEmpty":false,"Width":380,"Height":41},"X":-739,"Y":33,"Width":380,"Height":41,"Left":-739,"Top":33,"Right":-359,"Bottom":74,"TopLeft":{"X":-739,"Y":33},"TopRight":{"X":-359,"Y":33},"BottomLeft":{"X":-739,"Y":74},"BottomRight":{"X":-359,"Y":74}},"IsControlElement":true,"IsContentElement":true,"ProcessId":7768,"Orientation":0 supportedPatterns=[]
                                            properties="ControlType":{"LocalizedControlType":"pane","Id":50033,"ProgrammaticName":"ControlType.Pane"},"LocalizedControlType":"pane","IsEnabled":true,"BoundingRectangle":{"IsEmpty":false,"Location":{"X":-739,"Y":33},"Size":{"IsEmpty":false,"Width":380,"Height":2},"X":-739,"Y":33,"Width":380,"Height":2,"Left":-739,"Top":33,"Right":-359,"Bottom":35,"TopLeft":{"X":-739,"Y":33},"TopRight":{"X":-359,"Y":33},"BottomLeft":{"X":-739,"Y":35},"BottomRight":{"X":-359,"Y":35}},"IsControlElement":true,"IsContentElement":true,"ProcessId":7768,"Orientation":0 supportedPatterns=[]
                                            properties="ControlType":{"LocalizedControlType":"pane","Id":50033,"ProgrammaticName":"ControlType.Pane"},"LocalizedControlType":"pane","IsEnabled":true,"BoundingRectangle":{"IsEmpty":false,"Location":{"X":-739,"Y":35},"Size":{"IsEmpty":false,"Width":380,"Height":36},"X":-739,"Y":35,"Width":380,"Height":36,"Left":-739,"Top":35,"Right":-359,"Bottom":71,"TopLeft":{"X":-739,"Y":35},"TopRight":{"X":-359,"Y":35},"BottomLeft":{"X":-739,"Y":71},"BottomRight":{"X":-359,"Y":71}},"IsControlElement":true,"IsContentElement":true,"ProcessId":7768,"Orientation":0 supportedPatterns=[]
                                                properties="ControlType":{"LocalizedControlType":"pane","Id":50033,"ProgrammaticName":"ControlType.Pane"},"LocalizedControlType":"pane","IsEnabled":true,"BoundingRectangle":{"IsEmpty":false,"Location":{"X":-739,"Y":35},"Size":{"IsEmpty":false,"Width":10,"Height":36},"X":-739,"Y":35,"Width":10,"Height":36,"Left":-739,"Top":35,"Right":-729,"Bottom":71,"TopLeft":{"X":-739,"Y":35},"TopRight":{"X":-729,"Y":35},"BottomLeft":{"X":-739,"Y":71},"BottomRight":{"X":-729,"Y":71}},"IsControlElement":true,"IsContentElement":true,"ProcessId":7768,"Orientation":0 supportedPatterns=[]
                                                properties="ControlType":{"LocalizedControlType":"pane","Id":50033,"ProgrammaticName":"ControlType.Pane"},"LocalizedControlType":"pane","IsEnabled":true,"BoundingRectangle":{"IsEmpty":false,"Location":{"X":-726,"Y":37},"Size":{"IsEmpty":false,"Width":32,"Height":32},"X":-726,"Y":37,"Width":32,"Height":32,"Left":-726,"Top":37,"Right":-694,"Bottom":69,"TopLeft":{"X":-726,"Y":37},"TopRight":{"X":-694,"Y":37},"BottomLeft":{"X":-726,"Y":69},"BottomRight":{"X":-694,"Y":69}},"IsControlElement":true,"IsContentElement":true,"ProcessId":7768,"Orientation":0 supportedPatterns=[]
                                                    properties="ControlType":{"LocalizedControlType":"pane","Id":50033,"ProgrammaticName":"ControlType.Pane"},"LocalizedControlType":"pane","IsEnabled":true,"BoundingRectangle":{"IsEmpty":false,"Location":{"X":-726,"Y":37},"Size":{"IsEmpty":false,"Width":32,"Height":32},"X":-726,"Y":37,"Width":32,"Height":32,"Left":-726,"Top":37,"Right":-694,"Bottom":69,"TopLeft":{"X":-726,"Y":37},"TopRight":{"X":-694,"Y":37},"BottomLeft":{"X":-726,"Y":69},"BottomRight":{"X":-694,"Y":69}},"IsControlElement":true,"IsContentElement":true,"ProcessId":7768,"Orientation":0 supportedPatterns=[]
                                                properties="ControlType":{"LocalizedControlType":"pane","Id":50033,"ProgrammaticName":"ControlType.Pane"},"LocalizedControlType":"pane","IsEnabled":true,"BoundingRectangle":{"IsEmpty":false,"Location":{"X":-693,"Y":35},"Size":{"IsEmpty":false,"Width":8,"Height":36},"X":-693,"Y":35,"Width":8,"Height":36,"Left":-693,"Top":35,"Right":-685,"Bottom":71,"TopLeft":{"X":-693,"Y":35},"TopRight":{"X":-685,"Y":35},"BottomLeft":{"X":-693,"Y":71},"BottomRight":{"X":-685,"Y":71}},"IsControlElement":true,"IsContentElement":true,"ProcessId":7768,"Orientation":0 supportedPatterns=[]
                                                properties="ControlType":{"LocalizedControlType":"pane","Id":50033,"ProgrammaticName":"ControlType.Pane"},"LocalizedControlType":"pane","IsEnabled":true,"BoundingRectangle":{"IsEmpty":false,"Location":{"X":-685,"Y":35},"Size":{"IsEmpty":false,"Width":316,"Height":36},"X":-685,"Y":35,"Width":316,"Height":36,"Left":-685,"Top":35,"Right":-369,"Bottom":71,"TopLeft":{"X":-685,"Y":35},"TopRight":{"X":-369,"Y":35},"BottomLeft":{"X":-685,"Y":71},"BottomRight":{"X":-369,"Y":71}},"IsControlElement":true,"IsContentElement":true,"ProcessId":7768,"Orientation":0 supportedPatterns=[]
                                                    properties="ControlType":{"LocalizedControlType":"pane","Id":50033,"ProgrammaticName":"ControlType.Pane"},"LocalizedControlType":"pane","IsEnabled":true,"BoundingRectangle":{"IsEmpty":false,"Location":{"X":-685,"Y":35},"Size":{"IsEmpty":false,"Width":186,"Height":36},"X":-685,"Y":35,"Width":186,"Height":36,"Left":-685,"Top":35,"Right":-499,"Bottom":71,"TopLeft":{"X":-685,"Y":35},"TopRight":{"X":-499,"Y":35},"BottomLeft":{"X":-685,"Y":71},"BottomRight":{"X":-499,"Y":71}},"IsControlElement":true,"IsContentElement":true,"ProcessId":7768,"Orientation":0 supportedPatterns=[]
                                                        properties="ControlType":{"LocalizedControlType":"text","Id":50020,"ProgrammaticName":"ControlType.Text"},"LocalizedControlType":"text","Name":"Crispy Chris","IsEnabled":true,"BoundingRectangle":{"IsEmpty":false,"Location":{"X":-685,"Y":35},"Size":{"IsEmpty":false,"Width":69,"Height":36},"X":-685,"Y":35,"Width":69,"Height":36,"Left":-685,"Top":35,"Right":-616,"Bottom":71,"TopLeft":{"X":-685,"Y":35},"TopRight":{"X":-616,"Y":35},"BottomLeft":{"X":-685,"Y":71},"BottomRight":{"X":-616,"Y":71}},"IsControlElement":true,"IsContentElement":true,"ProcessId":7768,"Orientation":0 supportedPatterns=[]
                                                        properties="ControlType":{"LocalizedControlType":"text","Id":50020,"ProgrammaticName":"ControlType.Text"},"LocalizedControlType":"text","Name":"(Co-host, me)","IsEnabled":true,"BoundingRectangle":{"IsEmpty":false,"Location":{"X":-616,"Y":35},"Size":{"IsEmpty":false,"Width":83,"Height":36},"X":-616,"Y":35,"Width":83,"Height":36,"Left":-616,"Top":35,"Right":-533,"Bottom":71,"TopLeft":{"X":-616,"Y":35},"TopRight":{"X":-533,"Y":35},"BottomLeft":{"X":-616,"Y":71},"BottomRight":{"X":-533,"Y":71}},"IsControlElement":true,"IsContentElement":true,"ProcessId":7768,"Orientation":0 supportedPatterns=[]
                                                    properties="ControlType":{"LocalizedControlType":"button","Id":50000,"ProgrammaticName":"ControlType.Button"},"LocalizedControlType":"button","Name":"Mute","IsEnabled":true,"BoundingRectangle":{"IsEmpty":false,"Location":{"X":-495,"Y":41},"Size":{"IsEmpty":false,"Width":56,"Height":24},"X":-495,"Y":41,"Width":56,"Height":24,"Left":-495,"Top":41,"Right":-439,"Bottom":65,"TopLeft":{"X":-495,"Y":41},"TopRight":{"X":-439,"Y":41},"BottomLeft":{"X":-495,"Y":65},"BottomRight":{"X":-439,"Y":65}},"IsControlElement":true,"IsContentElement":true,"ProcessId":7768,"Orientation":0 supportedPatterns=["InvokePatternIdentifiers.Pattern"]
                                                    properties="ControlType":{"LocalizedControlType":"split button","Id":50031,"ProgrammaticName":"ControlType.SplitButton"},"LocalizedControlType":"split button","Name":"More options for Crispy Chris","IsEnabled":true,"BoundingRectangle":{"IsEmpty":false,"Location":{"X":-435,"Y":41},"Size":{"IsEmpty":false,"Width":66,"Height":24},"X":-435,"Y":41,"Width":66,"Height":24,"Left":-435,"Top":41,"Right":-369,"Bottom":65,"TopLeft":{"X":-435,"Y":41},"TopRight":{"X":-369,"Y":41},"BottomLeft":{"X":-435,"Y":65},"BottomRight":{"X":-369,"Y":65}},"IsControlElement":true,"IsContentElement":true,"ProcessId":7768,"Orientation":0 supportedPatterns=["InvokePatternIdentifiers.Pattern"]
                                                properties="ControlType":{"LocalizedControlType":"pane","Id":50033,"ProgrammaticName":"ControlType.Pane"},"LocalizedControlType":"pane","IsEnabled":true,"BoundingRectangle":{"IsEmpty":false,"Location":{"X":-369,"Y":35},"Size":{"IsEmpty":false,"Width":10,"Height":36},"X":-369,"Y":35,"Width":10,"Height":36,"Left":-369,"Top":35,"Right":-359,"Bottom":71,"TopLeft":{"X":-369,"Y":35},"TopRight":{"X":-359,"Y":35},"BottomLeft":{"X":-369,"Y":71},"BottomRight":{"X":-359,"Y":71}},"IsControlElement":true,"IsContentElement":true,"ProcessId":7768,"Orientation":0 supportedPatterns=[]
                                            properties="ControlType":{"LocalizedControlType":"pane","Id":50033,"ProgrammaticName":"ControlType.Pane"},"LocalizedControlType":"pane","IsEnabled":true,"BoundingRectangle":{"IsEmpty":false,"Location":{"X":-739,"Y":71},"Size":{"IsEmpty":false,"Width":380,"Height":2},"X":-739,"Y":71,"Width":380,"Height":2,"Left":-739,"Top":71,"Right":-359,"Bottom":73,"TopLeft":{"X":-739,"Y":71},"TopRight":{"X":-359,"Y":71},"BottomLeft":{"X":-739,"Y":73},"BottomRight":{"X":-359,"Y":73}},"IsControlElement":true,"IsContentElement":true,"ProcessId":7768,"Orientation":0 supportedPatterns=[]
                                            properties="ControlType":{"LocalizedControlType":"pane","Id":50033,"ProgrammaticName":"ControlType.Pane"},"LocalizedControlType":"pane","IsEnabled":true,"BoundingRectangle":{"IsEmpty":false,"Location":{"X":-739,"Y":73},"Size":{"IsEmpty":false,"Width":380,"Height":1},"X":-739,"Y":73,"Width":380,"Height":1,"Left":-739,"Top":73,"Right":-359,"Bottom":74,"TopLeft":{"X":-739,"Y":73},"TopRight":{"X":-359,"Y":73},"BottomLeft":{"X":-739,"Y":74},"BottomRight":{"X":-359,"Y":74}},"IsControlElement":true,"IsContentElement":true,"ProcessId":7768,"Orientation":0 supportedPatterns=[]
                                    properties="ControlType":{"LocalizedControlType":"list item","Id":50007,"ProgrammaticName":"ControlType.ListItem"},"LocalizedControlType":"list item","Name":"UsherBot,(Host), No Audio Connected,Video on","IsEnabled":true,"BoundingRectangle":{"IsEmpty":false,"Location":{"X":-739,"Y":74},"Size":{"IsEmpty":false,"Width":380,"Height":41},"X":-739,"Y":74,"Width":380,"Height":41,"Left":-739,"Top":74,"Right":-359,"Bottom":115,"TopLeft":{"X":-739,"Y":74},"TopRight":{"X":-359,"Y":74},"BottomLeft":{"X":-739,"Y":115},"BottomRight":{"X":-359,"Y":115}},"IsControlElement":true,"IsContentElement":true,"ProcessId":7768,"Orientation":0 supportedPatterns=["SelectionItemPatternIdentifiers.Pattern"]
                                        properties="ControlType":{"LocalizedControlType":"pane","Id":50033,"ProgrammaticName":"ControlType.Pane"},"LocalizedControlType":"pane","IsEnabled":true,"BoundingRectangle":{"IsEmpty":false,"Location":{"X":-739,"Y":74},"Size":{"IsEmpty":false,"Width":380,"Height":41},"X":-739,"Y":74,"Width":380,"Height":41,"Left":-739,"Top":74,"Right":-359,"Bottom":115,"TopLeft":{"X":-739,"Y":74},"TopRight":{"X":-359,"Y":74},"BottomLeft":{"X":-739,"Y":115},"BottomRight":{"X":-359,"Y":115}},"IsControlElement":true,"IsContentElement":true,"ProcessId":7768,"Orientation":0 supportedPatterns=[]
                                            properties="ControlType":{"LocalizedControlType":"pane","Id":50033,"ProgrammaticName":"ControlType.Pane"},"LocalizedControlType":"pane","IsEnabled":true,"BoundingRectangle":{"IsEmpty":false,"Location":{"X":-739,"Y":74},"Size":{"IsEmpty":false,"Width":380,"Height":2},"X":-739,"Y":74,"Width":380,"Height":2,"Left":-739,"Top":74,"Right":-359,"Bottom":76,"TopLeft":{"X":-739,"Y":74},"TopRight":{"X":-359,"Y":74},"BottomLeft":{"X":-739,"Y":76},"BottomRight":{"X":-359,"Y":76}},"IsControlElement":true,"IsContentElement":true,"ProcessId":7768,"Orientation":0 supportedPatterns=[]
                                            properties="ControlType":{"LocalizedControlType":"pane","Id":50033,"ProgrammaticName":"ControlType.Pane"},"LocalizedControlType":"pane","IsEnabled":true,"BoundingRectangle":{"IsEmpty":false,"Location":{"X":-739,"Y":76},"Size":{"IsEmpty":false,"Width":380,"Height":36},"X":-739,"Y":76,"Width":380,"Height":36,"Left":-739,"Top":76,"Right":-359,"Bottom":112,"TopLeft":{"X":-739,"Y":76},"TopRight":{"X":-359,"Y":76},"BottomLeft":{"X":-739,"Y":112},"BottomRight":{"X":-359,"Y":112}},"IsControlElement":true,"IsContentElement":true,"ProcessId":7768,"Orientation":0 supportedPatterns=[]
                                                properties="ControlType":{"LocalizedControlType":"pane","Id":50033,"ProgrammaticName":"ControlType.Pane"},"LocalizedControlType":"pane","IsEnabled":true,"BoundingRectangle":{"IsEmpty":false,"Location":{"X":-739,"Y":76},"Size":{"IsEmpty":false,"Width":10,"Height":36},"X":-739,"Y":76,"Width":10,"Height":36,"Left":-739,"Top":76,"Right":-729,"Bottom":112,"TopLeft":{"X":-739,"Y":76},"TopRight":{"X":-729,"Y":76},"BottomLeft":{"X":-739,"Y":112},"BottomRight":{"X":-729,"Y":112}},"IsControlElement":true,"IsContentElement":true,"ProcessId":7768,"Orientation":0 supportedPatterns=[]
                                                properties="ControlType":{"LocalizedControlType":"pane","Id":50033,"ProgrammaticName":"ControlType.Pane"},"LocalizedControlType":"pane","IsEnabled":true,"BoundingRectangle":{"IsEmpty":false,"Location":{"X":-726,"Y":78},"Size":{"IsEmpty":false,"Width":32,"Height":32},"X":-726,"Y":78,"Width":32,"Height":32,"Left":-726,"Top":78,"Right":-694,"Bottom":110,"TopLeft":{"X":-726,"Y":78},"TopRight":{"X":-694,"Y":78},"BottomLeft":{"X":-726,"Y":110},"BottomRight":{"X":-694,"Y":110}},"IsControlElement":true,"IsContentElement":true,"ProcessId":7768,"Orientation":0 supportedPatterns=[]
                                                    properties="ControlType":{"LocalizedControlType":"pane","Id":50033,"ProgrammaticName":"ControlType.Pane"},"LocalizedControlType":"pane","IsEnabled":true,"BoundingRectangle":{"IsEmpty":false,"Location":{"X":-726,"Y":78},"Size":{"IsEmpty":false,"Width":32,"Height":32},"X":-726,"Y":78,"Width":32,"Height":32,"Left":-726,"Top":78,"Right":-694,"Bottom":110,"TopLeft":{"X":-726,"Y":78},"TopRight":{"X":-694,"Y":78},"BottomLeft":{"X":-726,"Y":110},"BottomRight":{"X":-694,"Y":110}},"IsControlElement":true,"IsContentElement":true,"ProcessId":7768,"Orientation":0 supportedPatterns=[]
                                                properties="ControlType":{"LocalizedControlType":"pane","Id":50033,"ProgrammaticName":"ControlType.Pane"},"LocalizedControlType":"pane","IsEnabled":true,"BoundingRectangle":{"IsEmpty":false,"Location":{"X":-693,"Y":76},"Size":{"IsEmpty":false,"Width":8,"Height":36},"X":-693,"Y":76,"Width":8,"Height":36,"Left":-693,"Top":76,"Right":-685,"Bottom":112,"TopLeft":{"X":-693,"Y":76},"TopRight":{"X":-685,"Y":76},"BottomLeft":{"X":-693,"Y":112},"BottomRight":{"X":-685,"Y":112}},"IsControlElement":true,"IsContentElement":true,"ProcessId":7768,"Orientation":0 supportedPatterns=[]
                                                properties="ControlType":{"LocalizedControlType":"pane","Id":50033,"ProgrammaticName":"ControlType.Pane"},"LocalizedControlType":"pane","IsEnabled":true,"BoundingRectangle":{"IsEmpty":false,"Location":{"X":-685,"Y":76},"Size":{"IsEmpty":false,"Width":316,"Height":36},"X":-685,"Y":76,"Width":316,"Height":36,"Left":-685,"Top":76,"Right":-369,"Bottom":112,"TopLeft":{"X":-685,"Y":76},"TopRight":{"X":-369,"Y":76},"BottomLeft":{"X":-685,"Y":112},"BottomRight":{"X":-369,"Y":112}},"IsControlElement":true,"IsContentElement":true,"ProcessId":7768,"Orientation":0 supportedPatterns=[]
                                                    properties="ControlType":{"LocalizedControlType":"pane","Id":50033,"ProgrammaticName":"ControlType.Pane"},"LocalizedControlType":"pane","IsEnabled":true,"BoundingRectangle":{"IsEmpty":false,"Location":{"X":-685,"Y":76},"Size":{"IsEmpty":false,"Width":292,"Height":36},"X":-685,"Y":76,"Width":292,"Height":36,"Left":-685,"Top":76,"Right":-393,"Bottom":112,"TopLeft":{"X":-685,"Y":76},"TopRight":{"X":-393,"Y":76},"BottomLeft":{"X":-685,"Y":112},"BottomRight":{"X":-393,"Y":112}},"IsControlElement":true,"IsContentElement":true,"ProcessId":7768,"Orientation":0 supportedPatterns=[]
                                                        properties="ControlType":{"LocalizedControlType":"text","Id":50020,"ProgrammaticName":"ControlType.Text"},"LocalizedControlType":"text","Name":"UsherBot","IsEnabled":true,"BoundingRectangle":{"IsEmpty":false,"Location":{"X":-685,"Y":76},"Size":{"IsEmpty":false,"Width":53,"Height":36},"X":-685,"Y":76,"Width":53,"Height":36,"Left":-685,"Top":76,"Right":-632,"Bottom":112,"TopLeft":{"X":-685,"Y":76},"TopRight":{"X":-632,"Y":76},"BottomLeft":{"X":-685,"Y":112},"BottomRight":{"X":-632,"Y":112}},"IsControlElement":true,"IsContentElement":true,"ProcessId":7768,"Orientation":0 supportedPatterns=[]
                                                        properties="ControlType":{"LocalizedControlType":"text","Id":50020,"ProgrammaticName":"ControlType.Text"},"LocalizedControlType":"text","Name":"(Host)","IsEnabled":true,"BoundingRectangle":{"IsEmpty":false,"Location":{"X":-632,"Y":76},"Size":{"IsEmpty":false,"Width":39,"Height":36},"X":-632,"Y":76,"Width":39,"Height":36,"Left":-632,"Top":76,"Right":-593,"Bottom":112,"TopLeft":{"X":-632,"Y":76},"TopRight":{"X":-593,"Y":76},"BottomLeft":{"X":-632,"Y":112},"BottomRight":{"X":-593,"Y":112}},"IsControlElement":true,"IsContentElement":true,"ProcessId":7768,"Orientation":0 supportedPatterns=[]
                                                    properties="ControlType":{"LocalizedControlType":"text","Id":50020,"ProgrammaticName":"ControlType.Text"},"LocalizedControlType":"text","IsEnabled":true,"BoundingRectangle":{"IsEmpty":false,"Location":{"X":-389,"Y":86},"Size":{"IsEmpty":false,"Width":16,"Height":16},"X":-389,"Y":86,"Width":16,"Height":16,"Left":-389,"Top":86,"Right":-373,"Bottom":102,"TopLeft":{"X":-389,"Y":86},"TopRight":{"X":-373,"Y":86},"BottomLeft":{"X":-389,"Y":102},"BottomRight":{"X":-373,"Y":102}},"IsControlElement":true,"IsContentElement":true,"ProcessId":7768,"Orientation":0 supportedPatterns=[]
                                                properties="ControlType":{"LocalizedControlType":"pane","Id":50033,"ProgrammaticName":"ControlType.Pane"},"LocalizedControlType":"pane","IsEnabled":true,"BoundingRectangle":{"IsEmpty":false,"Location":{"X":-369,"Y":76},"Size":{"IsEmpty":false,"Width":10,"Height":36},"X":-369,"Y":76,"Width":10,"Height":36,"Left":-369,"Top":76,"Right":-359,"Bottom":112,"TopLeft":{"X":-369,"Y":76},"TopRight":{"X":-359,"Y":76},"BottomLeft":{"X":-369,"Y":112},"BottomRight":{"X":-359,"Y":112}},"IsControlElement":true,"IsContentElement":true,"ProcessId":7768,"Orientation":0 supportedPatterns=[]
                                            properties="ControlType":{"LocalizedControlType":"pane","Id":50033,"ProgrammaticName":"ControlType.Pane"},"LocalizedControlType":"pane","IsEnabled":true,"BoundingRectangle":{"IsEmpty":false,"Location":{"X":-739,"Y":112},"Size":{"IsEmpty":false,"Width":380,"Height":2},"X":-739,"Y":112,"Width":380,"Height":2,"Left":-739,"Top":112,"Right":-359,"Bottom":114,"TopLeft":{"X":-739,"Y":112},"TopRight":{"X":-359,"Y":112},"BottomLeft":{"X":-739,"Y":114},"BottomRight":{"X":-359,"Y":114}},"IsControlElement":true,"IsContentElement":true,"ProcessId":7768,"Orientation":0 supportedPatterns=[]
                                            properties="ControlType":{"LocalizedControlType":"pane","Id":50033,"ProgrammaticName":"ControlType.Pane"},"LocalizedControlType":"pane","IsEnabled":true,"BoundingRectangle":{"IsEmpty":false,"Location":{"X":-739,"Y":114},"Size":{"IsEmpty":false,"Width":380,"Height":1},"X":-739,"Y":114,"Width":380,"Height":1,"Left":-739,"Top":114,"Right":-359,"Bottom":115,"TopLeft":{"X":-739,"Y":114},"TopRight":{"X":-359,"Y":114},"BottomLeft":{"X":-739,"Y":115},"BottomRight":{"X":-359,"Y":115}},"IsControlElement":true,"IsContentElement":true,"ProcessId":7768,"Orientation":0 supportedPatterns=[]
                                    properties="ControlType":{"LocalizedControlType":"list item","Id":50007,"ProgrammaticName":"ControlType.ListItem"},"LocalizedControlType":"list item","Name":"Laura (usher),(Co-host), Telephone muted,Video off","IsEnabled":true,"BoundingRectangle":{"IsEmpty":false,"Location":{"X":-739,"Y":115},"Size":{"IsEmpty":false,"Width":380,"Height":41},"X":-739,"Y":115,"Width":380,"Height":41,"Left":-739,"Top":115,"Right":-359,"Bottom":156,"TopLeft":{"X":-739,"Y":115},"TopRight":{"X":-359,"Y":115},"BottomLeft":{"X":-739,"Y":156},"BottomRight":{"X":-359,"Y":156}},"IsControlElement":true,"IsContentElement":true,"ProcessId":7768,"Orientation":0 supportedPatterns=["SelectionItemPatternIdentifiers.Pattern"]
                                        properties="ControlType":{"LocalizedControlType":"pane","Id":50033,"ProgrammaticName":"ControlType.Pane"},"LocalizedControlType":"pane","IsEnabled":true,"BoundingRectangle":{"IsEmpty":false,"Location":{"X":-739,"Y":115},"Size":{"IsEmpty":false,"Width":380,"Height":41},"X":-739,"Y":115,"Width":380,"Height":41,"Left":-739,"Top":115,"Right":-359,"Bottom":156,"TopLeft":{"X":-739,"Y":115},"TopRight":{"X":-359,"Y":115},"BottomLeft":{"X":-739,"Y":156},"BottomRight":{"X":-359,"Y":156}},"IsControlElement":true,"IsContentElement":true,"ProcessId":7768,"Orientation":0 supportedPatterns=[]
                                            properties="ControlType":{"LocalizedControlType":"pane","Id":50033,"ProgrammaticName":"ControlType.Pane"},"LocalizedControlType":"pane","IsEnabled":true,"BoundingRectangle":{"IsEmpty":false,"Location":{"X":-739,"Y":115},"Size":{"IsEmpty":false,"Width":380,"Height":2},"X":-739,"Y":115,"Width":380,"Height":2,"Left":-739,"Top":115,"Right":-359,"Bottom":117,"TopLeft":{"X":-739,"Y":115},"TopRight":{"X":-359,"Y":115},"BottomLeft":{"X":-739,"Y":117},"BottomRight":{"X":-359,"Y":117}},"IsControlElement":true,"IsContentElement":true,"ProcessId":7768,"Orientation":0 supportedPatterns=[]
                                            properties="ControlType":{"LocalizedControlType":"pane","Id":50033,"ProgrammaticName":"ControlType.Pane"},"LocalizedControlType":"pane","IsEnabled":true,"BoundingRectangle":{"IsEmpty":false,"Location":{"X":-739,"Y":117},"Size":{"IsEmpty":false,"Width":380,"Height":36},"X":-739,"Y":117,"Width":380,"Height":36,"Left":-739,"Top":117,"Right":-359,"Bottom":153,"TopLeft":{"X":-739,"Y":117},"TopRight":{"X":-359,"Y":117},"BottomLeft":{"X":-739,"Y":153},"BottomRight":{"X":-359,"Y":153}},"IsControlElement":true,"IsContentElement":true,"ProcessId":7768,"Orientation":0 supportedPatterns=[]
                                                properties="ControlType":{"LocalizedControlType":"pane","Id":50033,"ProgrammaticName":"ControlType.Pane"},"LocalizedControlType":"pane","IsEnabled":true,"BoundingRectangle":{"IsEmpty":false,"Location":{"X":-739,"Y":117},"Size":{"IsEmpty":false,"Width":10,"Height":36},"X":-739,"Y":117,"Width":10,"Height":36,"Left":-739,"Top":117,"Right":-729,"Bottom":153,"TopLeft":{"X":-739,"Y":117},"TopRight":{"X":-729,"Y":117},"BottomLeft":{"X":-739,"Y":153},"BottomRight":{"X":-729,"Y":153}},"IsControlElement":true,"IsContentElement":true,"ProcessId":7768,"Orientation":0 supportedPatterns=[]
                                                properties="ControlType":{"LocalizedControlType":"pane","Id":50033,"ProgrammaticName":"ControlType.Pane"},"LocalizedControlType":"pane","IsEnabled":true,"BoundingRectangle":{"IsEmpty":false,"Location":{"X":-726,"Y":119},"Size":{"IsEmpty":false,"Width":32,"Height":32},"X":-726,"Y":119,"Width":32,"Height":32,"Left":-726,"Top":119,"Right":-694,"Bottom":151,"TopLeft":{"X":-726,"Y":119},"TopRight":{"X":-694,"Y":119},"BottomLeft":{"X":-726,"Y":151},"BottomRight":{"X":-694,"Y":151}},"IsControlElement":true,"IsContentElement":true,"ProcessId":7768,"Orientation":0 supportedPatterns=[]
                                                    properties="ControlType":{"LocalizedControlType":"pane","Id":50033,"ProgrammaticName":"ControlType.Pane"},"LocalizedControlType":"pane","IsEnabled":true,"BoundingRectangle":{"IsEmpty":false,"Location":{"X":-726,"Y":119},"Size":{"IsEmpty":false,"Width":32,"Height":32},"X":-726,"Y":119,"Width":32,"Height":32,"Left":-726,"Top":119,"Right":-694,"Bottom":151,"TopLeft":{"X":-726,"Y":119},"TopRight":{"X":-694,"Y":119},"BottomLeft":{"X":-726,"Y":151},"BottomRight":{"X":-694,"Y":151}},"IsControlElement":true,"IsContentElement":true,"ProcessId":7768,"Orientation":0 supportedPatterns=[]
                                                properties="ControlType":{"LocalizedControlType":"pane","Id":50033,"ProgrammaticName":"ControlType.Pane"},"LocalizedControlType":"pane","IsEnabled":true,"BoundingRectangle":{"IsEmpty":false,"Location":{"X":-693,"Y":117},"Size":{"IsEmpty":false,"Width":8,"Height":36},"X":-693,"Y":117,"Width":8,"Height":36,"Left":-693,"Top":117,"Right":-685,"Bottom":153,"TopLeft":{"X":-693,"Y":117},"TopRight":{"X":-685,"Y":117},"BottomLeft":{"X":-693,"Y":153},"BottomRight":{"X":-685,"Y":153}},"IsControlElement":true,"IsContentElement":true,"ProcessId":7768,"Orientation":0 supportedPatterns=[]
                                                properties="ControlType":{"LocalizedControlType":"pane","Id":50033,"ProgrammaticName":"ControlType.Pane"},"LocalizedControlType":"pane","IsEnabled":true,"BoundingRectangle":{"IsEmpty":false,"Location":{"X":-685,"Y":117},"Size":{"IsEmpty":false,"Width":316,"Height":36},"X":-685,"Y":117,"Width":316,"Height":36,"Left":-685,"Top":117,"Right":-369,"Bottom":153,"TopLeft":{"X":-685,"Y":117},"TopRight":{"X":-369,"Y":117},"BottomLeft":{"X":-685,"Y":153},"BottomRight":{"X":-369,"Y":153}},"IsControlElement":true,"IsContentElement":true,"ProcessId":7768,"Orientation":0 supportedPatterns=[]
                                                    properties="ControlType":{"LocalizedControlType":"pane","Id":50033,"ProgrammaticName":"ControlType.Pane"},"LocalizedControlType":"pane","IsEnabled":true,"BoundingRectangle":{"IsEmpty":false,"Location":{"X":-685,"Y":117},"Size":{"IsEmpty":false,"Width":268,"Height":36},"X":-685,"Y":117,"Width":268,"Height":36,"Left":-685,"Top":117,"Right":-417,"Bottom":153,"TopLeft":{"X":-685,"Y":117},"TopRight":{"X":-417,"Y":117},"BottomLeft":{"X":-685,"Y":153},"BottomRight":{"X":-417,"Y":153}},"IsControlElement":true,"IsContentElement":true,"ProcessId":7768,"Orientation":0 supportedPatterns=[]
                                                        properties="ControlType":{"LocalizedControlType":"text","Id":50020,"ProgrammaticName":"ControlType.Text"},"LocalizedControlType":"text","Name":"Laura (usher)","IsEnabled":true,"BoundingRectangle":{"IsEmpty":false,"Location":{"X":-685,"Y":117},"Size":{"IsEmpty":false,"Width":76,"Height":36},"X":-685,"Y":117,"Width":76,"Height":36,"Left":-685,"Top":117,"Right":-609,"Bottom":153,"TopLeft":{"X":-685,"Y":117},"TopRight":{"X":-609,"Y":117},"BottomLeft":{"X":-685,"Y":153},"BottomRight":{"X":-609,"Y":153}},"IsControlElement":true,"IsContentElement":true,"ProcessId":7768,"Orientation":0 supportedPatterns=[]
                                                        properties="ControlType":{"LocalizedControlType":"text","Id":50020,"ProgrammaticName":"ControlType.Text"},"LocalizedControlType":"text","Name":"(Co-host)","IsEnabled":true,"BoundingRectangle":{"IsEmpty":false,"Location":{"X":-609,"Y":117},"Size":{"IsEmpty":false,"Width":58,"Height":36},"X":-609,"Y":117,"Width":58,"Height":36,"Left":-609,"Top":117,"Right":-551,"Bottom":153,"TopLeft":{"X":-609,"Y":117},"TopRight":{"X":-551,"Y":117},"BottomLeft":{"X":-609,"Y":153},"BottomRight":{"X":-551,"Y":153}},"IsControlElement":true,"IsContentElement":true,"ProcessId":7768,"Orientation":0 supportedPatterns=[]
                                                    properties="ControlType":{"LocalizedControlType":"text","Id":50020,"ProgrammaticName":"ControlType.Text"},"LocalizedControlType":"text","IsEnabled":true,"BoundingRectangle":{"IsEmpty":false,"Location":{"X":-413,"Y":127},"Size":{"IsEmpty":false,"Width":16,"Height":16},"X":-413,"Y":127,"Width":16,"Height":16,"Left":-413,"Top":127,"Right":-397,"Bottom":143,"TopLeft":{"X":-413,"Y":127},"TopRight":{"X":-397,"Y":127},"BottomLeft":{"X":-413,"Y":143},"BottomRight":{"X":-397,"Y":143}},"IsControlElement":true,"IsContentElement":true,"ProcessId":7768,"Orientation":0 supportedPatterns=[]
                                                    properties="ControlType":{"LocalizedControlType":"text","Id":50020,"ProgrammaticName":"ControlType.Text"},"LocalizedControlType":"text","IsEnabled":true,"BoundingRectangle":{"IsEmpty":false,"Location":{"X":-389,"Y":127},"Size":{"IsEmpty":false,"Width":16,"Height":16},"X":-389,"Y":127,"Width":16,"Height":16,"Left":-389,"Top":127,"Right":-373,"Bottom":143,"TopLeft":{"X":-389,"Y":127},"TopRight":{"X":-373,"Y":127},"BottomLeft":{"X":-389,"Y":143},"BottomRight":{"X":-373,"Y":143}},"IsControlElement":true,"IsContentElement":true,"ProcessId":7768,"Orientation":0 supportedPatterns=[]
                                                properties="ControlType":{"LocalizedControlType":"pane","Id":50033,"ProgrammaticName":"ControlType.Pane"},"LocalizedControlType":"pane","IsEnabled":true,"BoundingRectangle":{"IsEmpty":false,"Location":{"X":-369,"Y":117},"Size":{"IsEmpty":false,"Width":10,"Height":36},"X":-369,"Y":117,"Width":10,"Height":36,"Left":-369,"Top":117,"Right":-359,"Bottom":153,"TopLeft":{"X":-369,"Y":117},"TopRight":{"X":-359,"Y":117},"BottomLeft":{"X":-369,"Y":153},"BottomRight":{"X":-359,"Y":153}},"IsControlElement":true,"IsContentElement":true,"ProcessId":7768,"Orientation":0 supportedPatterns=[]
                                            properties="ControlType":{"LocalizedControlType":"pane","Id":50033,"ProgrammaticName":"ControlType.Pane"},"LocalizedControlType":"pane","IsEnabled":true,"BoundingRectangle":{"IsEmpty":false,"Location":{"X":-739,"Y":153},"Size":{"IsEmpty":false,"Width":380,"Height":2},"X":-739,"Y":153,"Width":380,"Height":2,"Left":-739,"Top":153,"Right":-359,"Bottom":155,"TopLeft":{"X":-739,"Y":153},"TopRight":{"X":-359,"Y":153},"BottomLeft":{"X":-739,"Y":155},"BottomRight":{"X":-359,"Y":155}},"IsControlElement":true,"IsContentElement":true,"ProcessId":7768,"Orientation":0 supportedPatterns=[]
                                            properties="ControlType":{"LocalizedControlType":"pane","Id":50033,"ProgrammaticName":"ControlType.Pane"},"LocalizedControlType":"pane","IsEnabled":true,"BoundingRectangle":{"IsEmpty":false,"Location":{"X":-739,"Y":155},"Size":{"IsEmpty":false,"Width":380,"Height":1},"X":-739,"Y":155,"Width":380,"Height":1,"Left":-739,"Top":155,"Right":-359,"Bottom":156,"TopLeft":{"X":-739,"Y":155},"TopRight":{"X":-359,"Y":155},"BottomLeft":{"X":-739,"Y":156},"BottomRight":{"X":-359,"Y":156}},"IsControlElement":true,"IsContentElement":true,"ProcessId":7768,"Orientation":0 supportedPatterns=[]
                                    properties="ControlType":{"LocalizedControlType":"list item","Id":50007,"ProgrammaticName":"ControlType.ListItem"},"LocalizedControlType":"list item","Name":"Terri H, Computer audio unmuted,Video on","IsEnabled":true,"BoundingRectangle":{"IsEmpty":false,"Location":{"X":-739,"Y":156},"Size":{"IsEmpty":false,"Width":380,"Height":41},"X":-739,"Y":156,"Width":380,"Height":41,"Left":-739,"Top":156,"Right":-359,"Bottom":197,"TopLeft":{"X":-739,"Y":156},"TopRight":{"X":-359,"Y":156},"BottomLeft":{"X":-739,"Y":197},"BottomRight":{"X":-359,"Y":197}},"IsControlElement":true,"IsContentElement":true,"ProcessId":7768,"Orientation":0 supportedPatterns=["SelectionItemPatternIdentifiers.Pattern"]
                                        properties="ControlType":{"LocalizedControlType":"pane","Id":50033,"ProgrammaticName":"ControlType.Pane"},"LocalizedControlType":"pane","IsEnabled":true,"BoundingRectangle":{"IsEmpty":false,"Location":{"X":-739,"Y":156},"Size":{"IsEmpty":false,"Width":380,"Height":41},"X":-739,"Y":156,"Width":380,"Height":41,"Left":-739,"Top":156,"Right":-359,"Bottom":197,"TopLeft":{"X":-739,"Y":156},"TopRight":{"X":-359,"Y":156},"BottomLeft":{"X":-739,"Y":197},"BottomRight":{"X":-359,"Y":197}},"IsControlElement":true,"IsContentElement":true,"ProcessId":7768,"Orientation":0 supportedPatterns=[]
                                            properties="ControlType":{"LocalizedControlType":"pane","Id":50033,"ProgrammaticName":"ControlType.Pane"},"LocalizedControlType":"pane","IsEnabled":true,"BoundingRectangle":{"IsEmpty":false,"Location":{"X":-739,"Y":156},"Size":{"IsEmpty":false,"Width":380,"Height":2},"X":-739,"Y":156,"Width":380,"Height":2,"Left":-739,"Top":156,"Right":-359,"Bottom":158,"TopLeft":{"X":-739,"Y":156},"TopRight":{"X":-359,"Y":156},"BottomLeft":{"X":-739,"Y":158},"BottomRight":{"X":-359,"Y":158}},"IsControlElement":true,"IsContentElement":true,"ProcessId":7768,"Orientation":0 supportedPatterns=[]
                                            properties="ControlType":{"LocalizedControlType":"pane","Id":50033,"ProgrammaticName":"ControlType.Pane"},"LocalizedControlType":"pane","IsEnabled":true,"BoundingRectangle":{"IsEmpty":false,"Location":{"X":-739,"Y":158},"Size":{"IsEmpty":false,"Width":380,"Height":36},"X":-739,"Y":158,"Width":380,"Height":36,"Left":-739,"Top":158,"Right":-359,"Bottom":194,"TopLeft":{"X":-739,"Y":158},"TopRight":{"X":-359,"Y":158},"BottomLeft":{"X":-739,"Y":194},"BottomRight":{"X":-359,"Y":194}},"IsControlElement":true,"IsContentElement":true,"ProcessId":7768,"Orientation":0 supportedPatterns=[]
                                                properties="ControlType":{"LocalizedControlType":"pane","Id":50033,"ProgrammaticName":"ControlType.Pane"},"LocalizedControlType":"pane","IsEnabled":true,"BoundingRectangle":{"IsEmpty":false,"Location":{"X":-739,"Y":158},"Size":{"IsEmpty":false,"Width":10,"Height":36},"X":-739,"Y":158,"Width":10,"Height":36,"Left":-739,"Top":158,"Right":-729,"Bottom":194,"TopLeft":{"X":-739,"Y":158},"TopRight":{"X":-729,"Y":158},"BottomLeft":{"X":-739,"Y":194},"BottomRight":{"X":-729,"Y":194}},"IsControlElement":true,"IsContentElement":true,"ProcessId":7768,"Orientation":0 supportedPatterns=[]
                                                properties="ControlType":{"LocalizedControlType":"pane","Id":50033,"ProgrammaticName":"ControlType.Pane"},"LocalizedControlType":"pane","IsEnabled":true,"BoundingRectangle":{"IsEmpty":false,"Location":{"X":-726,"Y":160},"Size":{"IsEmpty":false,"Width":32,"Height":32},"X":-726,"Y":160,"Width":32,"Height":32,"Left":-726,"Top":160,"Right":-694,"Bottom":192,"TopLeft":{"X":-726,"Y":160},"TopRight":{"X":-694,"Y":160},"BottomLeft":{"X":-726,"Y":192},"BottomRight":{"X":-694,"Y":192}},"IsControlElement":true,"IsContentElement":true,"ProcessId":7768,"Orientation":0 supportedPatterns=[]
                                                    properties="ControlType":{"LocalizedControlType":"pane","Id":50033,"ProgrammaticName":"ControlType.Pane"},"LocalizedControlType":"pane","IsEnabled":true,"BoundingRectangle":{"IsEmpty":false,"Location":{"X":-726,"Y":160},"Size":{"IsEmpty":false,"Width":32,"Height":32},"X":-726,"Y":160,"Width":32,"Height":32,"Left":-726,"Top":160,"Right":-694,"Bottom":192,"TopLeft":{"X":-726,"Y":160},"TopRight":{"X":-694,"Y":160},"BottomLeft":{"X":-726,"Y":192},"BottomRight":{"X":-694,"Y":192}},"IsControlElement":true,"IsContentElement":true,"ProcessId":7768,"Orientation":0 supportedPatterns=[]
                                                properties="ControlType":{"LocalizedControlType":"pane","Id":50033,"ProgrammaticName":"ControlType.Pane"},"LocalizedControlType":"pane","IsEnabled":true,"BoundingRectangle":{"IsEmpty":false,"Location":{"X":-693,"Y":158},"Size":{"IsEmpty":false,"Width":8,"Height":36},"X":-693,"Y":158,"Width":8,"Height":36,"Left":-693,"Top":158,"Right":-685,"Bottom":194,"TopLeft":{"X":-693,"Y":158},"TopRight":{"X":-685,"Y":158},"BottomLeft":{"X":-693,"Y":194},"BottomRight":{"X":-685,"Y":194}},"IsControlElement":true,"IsContentElement":true,"ProcessId":7768,"Orientation":0 supportedPatterns=[]
                                                properties="ControlType":{"LocalizedControlType":"pane","Id":50033,"ProgrammaticName":"ControlType.Pane"},"LocalizedControlType":"pane","IsEnabled":true,"BoundingRectangle":{"IsEmpty":false,"Location":{"X":-685,"Y":158},"Size":{"IsEmpty":false,"Width":316,"Height":36},"X":-685,"Y":158,"Width":316,"Height":36,"Left":-685,"Top":158,"Right":-369,"Bottom":194,"TopLeft":{"X":-685,"Y":158},"TopRight":{"X":-369,"Y":158},"BottomLeft":{"X":-685,"Y":194},"BottomRight":{"X":-369,"Y":194}},"IsControlElement":true,"IsContentElement":true,"ProcessId":7768,"Orientation":0 supportedPatterns=[]
                                                    properties="ControlType":{"LocalizedControlType":"pane","Id":50033,"ProgrammaticName":"ControlType.Pane"},"LocalizedControlType":"pane","IsEnabled":true,"BoundingRectangle":{"IsEmpty":false,"Location":{"X":-685,"Y":158},"Size":{"IsEmpty":false,"Width":268,"Height":36},"X":-685,"Y":158,"Width":268,"Height":36,"Left":-685,"Top":158,"Right":-417,"Bottom":194,"TopLeft":{"X":-685,"Y":158},"TopRight":{"X":-417,"Y":158},"BottomLeft":{"X":-685,"Y":194},"BottomRight":{"X":-417,"Y":194}},"IsControlElement":true,"IsContentElement":true,"ProcessId":7768,"Orientation":0 supportedPatterns=[]
                                                        properties="ControlType":{"LocalizedControlType":"text","Id":50020,"ProgrammaticName":"ControlType.Text"},"LocalizedControlType":"text","Name":"Terri H","IsEnabled":true,"BoundingRectangle":{"IsEmpty":false,"Location":{"X":-685,"Y":158},"Size":{"IsEmpty":false,"Width":268,"Height":36},"X":-685,"Y":158,"Width":268,"Height":36,"Left":-685,"Top":158,"Right":-417,"Bottom":194,"TopLeft":{"X":-685,"Y":158},"TopRight":{"X":-417,"Y":158},"BottomLeft":{"X":-685,"Y":194},"BottomRight":{"X":-417,"Y":194}},"IsControlElement":true,"IsContentElement":true,"ProcessId":7768,"Orientation":0 supportedPatterns=[]
                                                    properties="ControlType":{"LocalizedControlType":"text","Id":50020,"ProgrammaticName":"ControlType.Text"},"LocalizedControlType":"text","IsEnabled":true,"BoundingRectangle":{"IsEmpty":false,"Location":{"X":-413,"Y":168},"Size":{"IsEmpty":false,"Width":16,"Height":16},"X":-413,"Y":168,"Width":16,"Height":16,"Left":-413,"Top":168,"Right":-397,"Bottom":184,"TopLeft":{"X":-413,"Y":168},"TopRight":{"X":-397,"Y":168},"BottomLeft":{"X":-413,"Y":184},"BottomRight":{"X":-397,"Y":184}},"IsControlElement":true,"IsContentElement":true,"ProcessId":7768,"Orientation":0 supportedPatterns=[]
                                                    properties="ControlType":{"LocalizedControlType":"text","Id":50020,"ProgrammaticName":"ControlType.Text"},"LocalizedControlType":"text","IsEnabled":true,"BoundingRectangle":{"IsEmpty":false,"Location":{"X":-389,"Y":168},"Size":{"IsEmpty":false,"Width":16,"Height":16},"X":-389,"Y":168,"Width":16,"Height":16,"Left":-389,"Top":168,"Right":-373,"Bottom":184,"TopLeft":{"X":-389,"Y":168},"TopRight":{"X":-373,"Y":168},"BottomLeft":{"X":-389,"Y":184},"BottomRight":{"X":-373,"Y":184}},"IsControlElement":true,"IsContentElement":true,"ProcessId":7768,"Orientation":0 supportedPatterns=[]
                                                properties="ControlType":{"LocalizedControlType":"pane","Id":50033,"ProgrammaticName":"ControlType.Pane"},"LocalizedControlType":"pane","IsEnabled":true,"BoundingRectangle":{"IsEmpty":false,"Location":{"X":-369,"Y":158},"Size":{"IsEmpty":false,"Width":10,"Height":36},"X":-369,"Y":158,"Width":10,"Height":36,"Left":-369,"Top":158,"Right":-359,"Bottom":194,"TopLeft":{"X":-369,"Y":158},"TopRight":{"X":-359,"Y":158},"BottomLeft":{"X":-369,"Y":194},"BottomRight":{"X":-359,"Y":194}},"IsControlElement":true,"IsContentElement":true,"ProcessId":7768,"Orientation":0 supportedPatterns=[]
                                            properties="ControlType":{"LocalizedControlType":"pane","Id":50033,"ProgrammaticName":"ControlType.Pane"},"LocalizedControlType":"pane","IsEnabled":true,"BoundingRectangle":{"IsEmpty":false,"Location":{"X":-739,"Y":194},"Size":{"IsEmpty":false,"Width":380,"Height":2},"X":-739,"Y":194,"Width":380,"Height":2,"Left":-739,"Top":194,"Right":-359,"Bottom":196,"TopLeft":{"X":-739,"Y":194},"TopRight":{"X":-359,"Y":194},"BottomLeft":{"X":-739,"Y":196},"BottomRight":{"X":-359,"Y":196}},"IsControlElement":true,"IsContentElement":true,"ProcessId":7768,"Orientation":0 supportedPatterns=[]
                                            properties="ControlType":{"LocalizedControlType":"pane","Id":50033,"ProgrammaticName":"ControlType.Pane"},"LocalizedControlType":"pane","IsEnabled":true,"BoundingRectangle":{"IsEmpty":false,"Location":{"X":-739,"Y":196},"Size":{"IsEmpty":false,"Width":380,"Height":1},"X":-739,"Y":196,"Width":380,"Height":1,"Left":-739,"Top":196,"Right":-359,"Bottom":197,"TopLeft":{"X":-739,"Y":196},"TopRight":{"X":-359,"Y":196},"BottomLeft":{"X":-739,"Y":197},"BottomRight":{"X":-359,"Y":197}},"IsControlElement":true,"IsContentElement":true,"ProcessId":7768,"Orientation":0 supportedPatterns=[]
                                properties="ControlType":{"LocalizedControlType":"pane","Id":50033,"ProgrammaticName":"ControlType.Pane"},"LocalizedControlType":"pane","IsEnabled":true,"BoundingRectangle":{"IsEmpty":false,"Location":{"X":-741,"Y":912},"Size":{"IsEmpty":false,"Width":384,"Height":1},"X":-741,"Y":912,"Width":384,"Height":1,"Left":-741,"Top":912,"Right":-357,"Bottom":913,"TopLeft":{"X":-741,"Y":912},"TopRight":{"X":-357,"Y":912},"BottomLeft":{"X":-741,"Y":913},"BottomRight":{"X":-357,"Y":913}},"IsControlElement":true,"IsContentElement":true,"ProcessId":7768,"Orientation":0 supportedPatterns=[]
                                properties="ControlType":{"LocalizedControlType":"pane","Id":50033,"ProgrammaticName":"ControlType.Pane"},"LocalizedControlType":"pane","IsEnabled":true,"BoundingRectangle":{"IsEmpty":false,"Location":{"X":-741,"Y":913},"Size":{"IsEmpty":false,"Width":384,"Height":74},"X":-741,"Y":913,"Width":384,"Height":74,"Left":-741,"Top":913,"Right":-357,"Bottom":987,"TopLeft":{"X":-741,"Y":913},"TopRight":{"X":-357,"Y":913},"BottomLeft":{"X":-741,"Y":987},"BottomRight":{"X":-357,"Y":987}},"IsControlElement":true,"IsContentElement":true,"ProcessId":7768,"Orientation":0 supportedPatterns=[]
                                    properties="ControlType":{"LocalizedControlType":"pane","Id":50033,"ProgrammaticName":"ControlType.Pane"},"LocalizedControlType":"pane","IsEnabled":true,"BoundingRectangle":{"IsEmpty":false,"Location":{"X":-741,"Y":913},"Size":{"IsEmpty":false,"Width":384,"Height":74},"X":-741,"Y":913,"Width":384,"Height":74,"Left":-741,"Top":913,"Right":-357,"Bottom":987,"TopLeft":{"X":-741,"Y":913},"TopRight":{"X":-357,"Y":913},"BottomLeft":{"X":-741,"Y":987},"BottomRight":{"X":-357,"Y":987}},"IsControlElement":true,"IsContentElement":true,"ProcessId":7768,"Orientation":0 supportedPatterns=[]
                                        properties="ControlType":{"LocalizedControlType":"pane","Id":50033,"ProgrammaticName":"ControlType.Pane"},"LocalizedControlType":"pane","IsEnabled":true,"BoundingRectangle":{"IsEmpty":false,"Location":{"X":-731,"Y":922},"Size":{"IsEmpty":false,"Width":364,"Height":59},"X":-731,"Y":922,"Width":364,"Height":59,"Left":-731,"Top":922,"Right":-367,"Bottom":981,"TopLeft":{"X":-731,"Y":922},"TopRight":{"X":-367,"Y":922},"BottomLeft":{"X":-731,"Y":981},"BottomRight":{"X":-367,"Y":981}},"IsControlElement":true,"IsContentElement":true,"ProcessId":7768,"Orientation":0 supportedPatterns=[]
                                            properties="ControlType":{"LocalizedControlType":"pane","Id":50033,"ProgrammaticName":"ControlType.Pane"},"LocalizedControlType":"pane","IsEnabled":true,"BoundingRectangle":{"IsEmpty":false,"Location":{"X":-731,"Y":922},"Size":{"IsEmpty":false,"Width":60,"Height":59},"X":-731,"Y":922,"Width":60,"Height":59,"Left":-731,"Top":922,"Right":-671,"Bottom":981,"TopLeft":{"X":-731,"Y":922},"TopRight":{"X":-671,"Y":922},"BottomLeft":{"X":-731,"Y":981},"BottomRight":{"X":-671,"Y":981}},"IsControlElement":true,"IsContentElement":true,"ProcessId":7768,"Orientation":0 supportedPatterns=[]
                                                properties="ControlType":{"LocalizedControlType":"pane","Id":50033,"ProgrammaticName":"ControlType.Pane"},"LocalizedControlType":"pane","IsEnabled":true,"BoundingRectangle":{"IsEmpty":true,"Location":{"X":"\u221E","Y":"\u221E"},"Size":{"IsEmpty":true,"Width":"-\u221E","Height":"-\u221E"},"X":"\u221E","Y":"\u221E","Width":"-\u221E","Height":"-\u221E","Left":"\u221E","Top":"\u221E","Right":"-\u221E","Bottom":"-\u221E","TopLeft":{"X":"\u221E","Y":"\u221E"},"TopRight":{"X":"-\u221E","Y":"\u221E"},"BottomLeft":{"X":"\u221E","Y":"-\u221E"},"BottomRight":{"X":"-\u221E","Y":"-\u221E"}},"IsControlElement":true,"IsContentElement":true,"ProcessId":7768,"Orientation":0 supportedPatterns=[]
                                                properties="ControlType":{"LocalizedControlType":"text","Id":50020,"ProgrammaticName":"ControlType.Text"},"LocalizedControlType":"text","IsEnabled":true,"BoundingRectangle":{"IsEmpty":false,"Location":{"X":-703,"Y":922},"Size":{"IsEmpty":false,"Width":4,"Height":13},"X":-703,"Y":922,"Width":4,"Height":13,"Left":-703,"Top":922,"Right":-699,"Bottom":935,"TopLeft":{"X":-703,"Y":922},"TopRight":{"X":-699,"Y":922},"BottomLeft":{"X":-703,"Y":935},"BottomRight":{"X":-699,"Y":935}},"IsControlElement":true,"IsContentElement":true,"ProcessId":7768,"Orientation":0 supportedPatterns=[]
                                                properties="ControlType":{"LocalizedControlType":"button","Id":50000,"ProgrammaticName":"ControlType.Button"},"LocalizedControlType":"button","Name":"yes","IsEnabled":true,"BoundingRectangle":{"IsEmpty":false,"Location":{"X":-716,"Y":935},"Size":{"IsEmpty":false,"Width":30,"Height":30},"X":-716,"Y":935,"Width":30,"Height":30,"Left":-716,"Top":935,"Right":-686,"Bottom":965,"TopLeft":{"X":-716,"Y":935},"TopRight":{"X":-686,"Y":935},"BottomLeft":{"X":-716,"Y":965},"BottomRight":{"X":-686,"Y":965}},"IsControlElement":true,"IsContentElement":true,"ProcessId":7768,"Orientation":0 supportedPatterns=["InvokePatternIdentifiers.Pattern"]
                                                properties="ControlType":{"LocalizedControlType":"text","Id":50020,"ProgrammaticName":"ControlType.Text"},"LocalizedControlType":"text","Name":"yes","IsEnabled":true,"BoundingRectangle":{"IsEmpty":false,"Location":{"X":-711,"Y":965},"Size":{"IsEmpty":false,"Width":20,"Height":15},"X":-711,"Y":965,"Width":20,"Height":15,"Left":-711,"Top":965,"Right":-691,"Bottom":980,"TopLeft":{"X":-711,"Y":965},"TopRight":{"X":-691,"Y":965},"BottomLeft":{"X":-711,"Y":980},"BottomRight":{"X":-691,"Y":980}},"IsControlElement":true,"IsContentElement":true,"ProcessId":7768,"Orientation":0 supportedPatterns=[]
                                                properties="ControlType":{"LocalizedControlType":"pane","Id":50033,"ProgrammaticName":"ControlType.Pane"},"LocalizedControlType":"pane","IsEnabled":true,"BoundingRectangle":{"IsEmpty":false,"Location":{"X":-731,"Y":980},"Size":{"IsEmpty":false,"Width":60,"Height":7},"X":-731,"Y":980,"Width":60,"Height":7,"Left":-731,"Top":980,"Right":-671,"Bottom":987,"TopLeft":{"X":-731,"Y":980},"TopRight":{"X":-671,"Y":980},"BottomLeft":{"X":-731,"Y":987},"BottomRight":{"X":-671,"Y":987}},"IsControlElement":true,"IsContentElement":true,"ProcessId":7768,"Orientation":0 supportedPatterns=[]
                                            properties="ControlType":{"LocalizedControlType":"pane","Id":50033,"ProgrammaticName":"ControlType.Pane"},"LocalizedControlType":"pane","IsEnabled":true,"BoundingRectangle":{"IsEmpty":false,"Location":{"X":-671,"Y":922},"Size":{"IsEmpty":false,"Width":60,"Height":59},"X":-671,"Y":922,"Width":60,"Height":59,"Left":-671,"Top":922,"Right":-611,"Bottom":981,"TopLeft":{"X":-671,"Y":922},"TopRight":{"X":-611,"Y":922},"BottomLeft":{"X":-671,"Y":981},"BottomRight":{"X":-611,"Y":981}},"IsControlElement":true,"IsContentElement":true,"ProcessId":7768,"Orientation":0 supportedPatterns=[]
                                                properties="ControlType":{"LocalizedControlType":"pane","Id":50033,"ProgrammaticName":"ControlType.Pane"},"LocalizedControlType":"pane","IsEnabled":true,"BoundingRectangle":{"IsEmpty":true,"Location":{"X":"\u221E","Y":"\u221E"},"Size":{"IsEmpty":true,"Width":"-\u221E","Height":"-\u221E"},"X":"\u221E","Y":"\u221E","Width":"-\u221E","Height":"-\u221E","Left":"\u221E","Top":"\u221E","Right":"-\u221E","Bottom":"-\u221E","TopLeft":{"X":"\u221E","Y":"\u221E"},"TopRight":{"X":"-\u221E","Y":"\u221E"},"BottomLeft":{"X":"\u221E","Y":"-\u221E"},"BottomRight":{"X":"-\u221E","Y":"-\u221E"}},"IsControlElement":true,"IsContentElement":true,"ProcessId":7768,"Orientation":0 supportedPatterns=[]
                                                properties="ControlType":{"LocalizedControlType":"text","Id":50020,"ProgrammaticName":"ControlType.Text"},"LocalizedControlType":"text","IsEnabled":true,"BoundingRectangle":{"IsEmpty":false,"Location":{"X":-643,"Y":922},"Size":{"IsEmpty":false,"Width":4,"Height":13},"X":-643,"Y":922,"Width":4,"Height":13,"Left":-643,"Top":922,"Right":-639,"Bottom":935,"TopLeft":{"X":-643,"Y":922},"TopRight":{"X":-639,"Y":922},"BottomLeft":{"X":-643,"Y":935},"BottomRight":{"X":-639,"Y":935}},"IsControlElement":true,"IsContentElement":true,"ProcessId":7768,"Orientation":0 supportedPatterns=[]
                                                properties="ControlType":{"LocalizedControlType":"button","Id":50000,"ProgrammaticName":"ControlType.Button"},"LocalizedControlType":"button","Name":"no","IsEnabled":true,"BoundingRectangle":{"IsEmpty":false,"Location":{"X":-656,"Y":935},"Size":{"IsEmpty":false,"Width":30,"Height":30},"X":-656,"Y":935,"Width":30,"Height":30,"Left":-656,"Top":935,"Right":-626,"Bottom":965,"TopLeft":{"X":-656,"Y":935},"TopRight":{"X":-626,"Y":935},"BottomLeft":{"X":-656,"Y":965},"BottomRight":{"X":-626,"Y":965}},"IsControlElement":true,"IsContentElement":true,"ProcessId":7768,"Orientation":0 supportedPatterns=["InvokePatternIdentifiers.Pattern"]
                                                properties="ControlType":{"LocalizedControlType":"text","Id":50020,"ProgrammaticName":"ControlType.Text"},"LocalizedControlType":"text","Name":"no","IsEnabled":true,"BoundingRectangle":{"IsEmpty":false,"Location":{"X":-650,"Y":965},"Size":{"IsEmpty":false,"Width":18,"Height":15},"X":-650,"Y":965,"Width":18,"Height":15,"Left":-650,"Top":965,"Right":-632,"Bottom":980,"TopLeft":{"X":-650,"Y":965},"TopRight":{"X":-632,"Y":965},"BottomLeft":{"X":-650,"Y":980},"BottomRight":{"X":-632,"Y":980}},"IsControlElement":true,"IsContentElement":true,"ProcessId":7768,"Orientation":0 supportedPatterns=[]
                                                properties="ControlType":{"LocalizedControlType":"pane","Id":50033,"ProgrammaticName":"ControlType.Pane"},"LocalizedControlType":"pane","IsEnabled":true,"BoundingRectangle":{"IsEmpty":false,"Location":{"X":-671,"Y":980},"Size":{"IsEmpty":false,"Width":60,"Height":7},"X":-671,"Y":980,"Width":60,"Height":7,"Left":-671,"Top":980,"Right":-611,"Bottom":987,"TopLeft":{"X":-671,"Y":980},"TopRight":{"X":-611,"Y":980},"BottomLeft":{"X":-671,"Y":987},"BottomRight":{"X":-611,"Y":987}},"IsControlElement":true,"IsContentElement":true,"ProcessId":7768,"Orientation":0 supportedPatterns=[]
                                            properties="ControlType":{"LocalizedControlType":"pane","Id":50033,"ProgrammaticName":"ControlType.Pane"},"LocalizedControlType":"pane","IsEnabled":true,"BoundingRectangle":{"IsEmpty":false,"Location":{"X":-611,"Y":922},"Size":{"IsEmpty":false,"Width":60,"Height":59},"X":-611,"Y":922,"Width":60,"Height":59,"Left":-611,"Top":922,"Right":-551,"Bottom":981,"TopLeft":{"X":-611,"Y":922},"TopRight":{"X":-551,"Y":922},"BottomLeft":{"X":-611,"Y":981},"BottomRight":{"X":-551,"Y":981}},"IsControlElement":true,"IsContentElement":true,"ProcessId":7768,"Orientation":0 supportedPatterns=[]
                                                properties="ControlType":{"LocalizedControlType":"pane","Id":50033,"ProgrammaticName":"ControlType.Pane"},"LocalizedControlType":"pane","IsEnabled":true,"BoundingRectangle":{"IsEmpty":true,"Location":{"X":"\u221E","Y":"\u221E"},"Size":{"IsEmpty":true,"Width":"-\u221E","Height":"-\u221E"},"X":"\u221E","Y":"\u221E","Width":"-\u221E","Height":"-\u221E","Left":"\u221E","Top":"\u221E","Right":"-\u221E","Bottom":"-\u221E","TopLeft":{"X":"\u221E","Y":"\u221E"},"TopRight":{"X":"-\u221E","Y":"\u221E"},"BottomLeft":{"X":"\u221E","Y":"-\u221E"},"BottomRight":{"X":"-\u221E","Y":"-\u221E"}},"IsControlElement":true,"IsContentElement":true,"ProcessId":7768,"Orientation":0 supportedPatterns=[]
                                                properties="ControlType":{"LocalizedControlType":"text","Id":50020,"ProgrammaticName":"ControlType.Text"},"LocalizedControlType":"text","IsEnabled":true,"BoundingRectangle":{"IsEmpty":false,"Location":{"X":-583,"Y":922},"Size":{"IsEmpty":false,"Width":4,"Height":13},"X":-583,"Y":922,"Width":4,"Height":13,"Left":-583,"Top":922,"Right":-579,"Bottom":935,"TopLeft":{"X":-583,"Y":922},"TopRight":{"X":-579,"Y":922},"BottomLeft":{"X":-583,"Y":935},"BottomRight":{"X":-579,"Y":935}},"IsControlElement":true,"IsContentElement":true,"ProcessId":7768,"Orientation":0 supportedPatterns=[]
                                                properties="ControlType":{"LocalizedControlType":"button","Id":50000,"ProgrammaticName":"ControlType.Button"},"LocalizedControlType":"button","Name":"go slower","IsEnabled":true,"BoundingRectangle":{"IsEmpty":false,"Location":{"X":-596,"Y":935},"Size":{"IsEmpty":false,"Width":30,"Height":30},"X":-596,"Y":935,"Width":30,"Height":30,"Left":-596,"Top":935,"Right":-566,"Bottom":965,"TopLeft":{"X":-596,"Y":935},"TopRight":{"X":-566,"Y":935},"BottomLeft":{"X":-596,"Y":965},"BottomRight":{"X":-566,"Y":965}},"IsControlElement":true,"IsContentElement":true,"ProcessId":7768,"Orientation":0 supportedPatterns=["InvokePatternIdentifiers.Pattern"]
                                                properties="ControlType":{"LocalizedControlType":"text","Id":50020,"ProgrammaticName":"ControlType.Text"},"LocalizedControlType":"text","Name":"go slower","IsEnabled":true,"BoundingRectangle":{"IsEmpty":false,"Location":{"X":-609,"Y":965},"Size":{"IsEmpty":false,"Width":55,"Height":15},"X":-609,"Y":965,"Width":55,"Height":15,"Left":-609,"Top":965,"Right":-554,"Bottom":980,"TopLeft":{"X":-609,"Y":965},"TopRight":{"X":-554,"Y":965},"BottomLeft":{"X":-609,"Y":980},"BottomRight":{"X":-554,"Y":980}},"IsControlElement":true,"IsContentElement":true,"ProcessId":7768,"Orientation":0 supportedPatterns=[]
                                                properties="ControlType":{"LocalizedControlType":"pane","Id":50033,"ProgrammaticName":"ControlType.Pane"},"LocalizedControlType":"pane","IsEnabled":true,"BoundingRectangle":{"IsEmpty":false,"Location":{"X":-611,"Y":980},"Size":{"IsEmpty":false,"Width":60,"Height":7},"X":-611,"Y":980,"Width":60,"Height":7,"Left":-611,"Top":980,"Right":-551,"Bottom":987,"TopLeft":{"X":-611,"Y":980},"TopRight":{"X":-551,"Y":980},"BottomLeft":{"X":-611,"Y":987},"BottomRight":{"X":-551,"Y":987}},"IsControlElement":true,"IsContentElement":true,"ProcessId":7768,"Orientation":0 supportedPatterns=[]
                                            properties="ControlType":{"LocalizedControlType":"pane","Id":50033,"ProgrammaticName":"ControlType.Pane"},"LocalizedControlType":"pane","IsEnabled":true,"BoundingRectangle":{"IsEmpty":false,"Location":{"X":-551,"Y":922},"Size":{"IsEmpty":false,"Width":60,"Height":59},"X":-551,"Y":922,"Width":60,"Height":59,"Left":-551,"Top":922,"Right":-491,"Bottom":981,"TopLeft":{"X":-551,"Y":922},"TopRight":{"X":-491,"Y":922},"BottomLeft":{"X":-551,"Y":981},"BottomRight":{"X":-491,"Y":981}},"IsControlElement":true,"IsContentElement":true,"ProcessId":7768,"Orientation":0 supportedPatterns=[]
                                                properties="ControlType":{"LocalizedControlType":"pane","Id":50033,"ProgrammaticName":"ControlType.Pane"},"LocalizedControlType":"pane","IsEnabled":true,"BoundingRectangle":{"IsEmpty":true,"Location":{"X":"\u221E","Y":"\u221E"},"Size":{"IsEmpty":true,"Width":"-\u221E","Height":"-\u221E"},"X":"\u221E","Y":"\u221E","Width":"-\u221E","Height":"-\u221E","Left":"\u221E","Top":"\u221E","Right":"-\u221E","Bottom":"-\u221E","TopLeft":{"X":"\u221E","Y":"\u221E"},"TopRight":{"X":"-\u221E","Y":"\u221E"},"BottomLeft":{"X":"\u221E","Y":"-\u221E"},"BottomRight":{"X":"-\u221E","Y":"-\u221E"}},"IsControlElement":true,"IsContentElement":true,"ProcessId":7768,"Orientation":0 supportedPatterns=[]
                                                properties="ControlType":{"LocalizedControlType":"text","Id":50020,"ProgrammaticName":"ControlType.Text"},"LocalizedControlType":"text","IsEnabled":true,"BoundingRectangle":{"IsEmpty":false,"Location":{"X":-523,"Y":922},"Size":{"IsEmpty":false,"Width":4,"Height":13},"X":-523,"Y":922,"Width":4,"Height":13,"Left":-523,"Top":922,"Right":-519,"Bottom":935,"TopLeft":{"X":-523,"Y":922},"TopRight":{"X":-519,"Y":922},"BottomLeft":{"X":-523,"Y":935},"BottomRight":{"X":-519,"Y":935}},"IsControlElement":true,"IsContentElement":true,"ProcessId":7768,"Orientation":0 supportedPatterns=[]
                                                properties="ControlType":{"LocalizedControlType":"button","Id":50000,"ProgrammaticName":"ControlType.Button"},"LocalizedControlType":"button","Name":"go faster","IsEnabled":true,"BoundingRectangle":{"IsEmpty":false,"Location":{"X":-536,"Y":935},"Size":{"IsEmpty":false,"Width":30,"Height":30},"X":-536,"Y":935,"Width":30,"Height":30,"Left":-536,"Top":935,"Right":-506,"Bottom":965,"TopLeft":{"X":-536,"Y":935},"TopRight":{"X":-506,"Y":935},"BottomLeft":{"X":-536,"Y":965},"BottomRight":{"X":-506,"Y":965}},"IsControlElement":true,"IsContentElement":true,"ProcessId":7768,"Orientation":0 supportedPatterns=["InvokePatternIdentifiers.Pattern"]
                                                properties="ControlType":{"LocalizedControlType":"text","Id":50020,"ProgrammaticName":"ControlType.Text"},"LocalizedControlType":"text","Name":"go faster","IsEnabled":true,"BoundingRectangle":{"IsEmpty":false,"Location":{"X":-546,"Y":965},"Size":{"IsEmpty":false,"Width":50,"Height":15},"X":-546,"Y":965,"Width":50,"Height":15,"Left":-546,"Top":965,"Right":-496,"Bottom":980,"TopLeft":{"X":-546,"Y":965},"TopRight":{"X":-496,"Y":965},"BottomLeft":{"X":-546,"Y":980},"BottomRight":{"X":-496,"Y":980}},"IsControlElement":true,"IsContentElement":true,"ProcessId":7768,"Orientation":0 supportedPatterns=[]
                                                properties="ControlType":{"LocalizedControlType":"pane","Id":50033,"ProgrammaticName":"ControlType.Pane"},"LocalizedControlType":"pane","IsEnabled":true,"BoundingRectangle":{"IsEmpty":false,"Location":{"X":-551,"Y":980},"Size":{"IsEmpty":false,"Width":60,"Height":7},"X":-551,"Y":980,"Width":60,"Height":7,"Left":-551,"Top":980,"Right":-491,"Bottom":987,"TopLeft":{"X":-551,"Y":980},"TopRight":{"X":-491,"Y":980},"BottomLeft":{"X":-551,"Y":987},"BottomRight":{"X":-491,"Y":987}},"IsControlElement":true,"IsContentElement":true,"ProcessId":7768,"Orientation":0 supportedPatterns=[]
                                            properties="ControlType":{"LocalizedControlType":"pane","Id":50033,"ProgrammaticName":"ControlType.Pane"},"LocalizedControlType":"pane","IsEnabled":true,"BoundingRectangle":{"IsEmpty":false,"Location":{"X":-491,"Y":922},"Size":{"IsEmpty":false,"Width":60,"Height":59},"X":-491,"Y":922,"Width":60,"Height":59,"Left":-491,"Top":922,"Right":-431,"Bottom":981,"TopLeft":{"X":-491,"Y":922},"TopRight":{"X":-431,"Y":922},"BottomLeft":{"X":-491,"Y":981},"BottomRight":{"X":-431,"Y":981}},"IsControlElement":true,"IsContentElement":true,"ProcessId":7768,"Orientation":0 supportedPatterns=[]
                                                properties="ControlType":{"LocalizedControlType":"pane","Id":50033,"ProgrammaticName":"ControlType.Pane"},"LocalizedControlType":"pane","IsEnabled":true,"BoundingRectangle":{"IsEmpty":true,"Location":{"X":"\u221E","Y":"\u221E"},"Size":{"IsEmpty":true,"Width":"-\u221E","Height":"-\u221E"},"X":"\u221E","Y":"\u221E","Width":"-\u221E","Height":"-\u221E","Left":"\u221E","Top":"\u221E","Right":"-\u221E","Bottom":"-\u221E","TopLeft":{"X":"\u221E","Y":"\u221E"},"TopRight":{"X":"-\u221E","Y":"\u221E"},"BottomLeft":{"X":"\u221E","Y":"-\u221E"},"BottomRight":{"X":"-\u221E","Y":"-\u221E"}},"IsControlElement":true,"IsContentElement":true,"ProcessId":7768,"Orientation":0 supportedPatterns=[]
                                                properties="ControlType":{"LocalizedControlType":"text","Id":50020,"ProgrammaticName":"ControlType.Text"},"LocalizedControlType":"text","IsEnabled":true,"BoundingRectangle":{"IsEmpty":false,"Location":{"X":-463,"Y":922},"Size":{"IsEmpty":false,"Width":4,"Height":13},"X":-463,"Y":922,"Width":4,"Height":13,"Left":-463,"Top":922,"Right":-459,"Bottom":935,"TopLeft":{"X":-463,"Y":922},"TopRight":{"X":-459,"Y":922},"BottomLeft":{"X":-463,"Y":935},"BottomRight":{"X":-459,"Y":935}},"IsControlElement":true,"IsContentElement":true,"ProcessId":7768,"Orientation":0 supportedPatterns=[]
                                                properties="ControlType":{"LocalizedControlType":"split button","Id":50031,"ProgrammaticName":"ControlType.SplitButton"},"LocalizedControlType":"split button","Name":"More nonverbal feedbacks","IsEnabled":true,"BoundingRectangle":{"IsEmpty":false,"Location":{"X":-476,"Y":935},"Size":{"IsEmpty":false,"Width":30,"Height":30},"X":-476,"Y":935,"Width":30,"Height":30,"Left":-476,"Top":935,"Right":-446,"Bottom":965,"TopLeft":{"X":-476,"Y":935},"TopRight":{"X":-446,"Y":935},"BottomLeft":{"X":-476,"Y":965},"BottomRight":{"X":-446,"Y":965}},"IsControlElement":true,"IsContentElement":true,"ProcessId":7768,"Orientation":0 supportedPatterns=["InvokePatternIdentifiers.Pattern"]
                                                properties="ControlType":{"LocalizedControlType":"text","Id":50020,"ProgrammaticName":"ControlType.Text"},"LocalizedControlType":"text","Name":"more","IsEnabled":true,"BoundingRectangle":{"IsEmpty":false,"Location":{"X":-476,"Y":965},"Size":{"IsEmpty":false,"Width":30,"Height":15},"X":-476,"Y":965,"Width":30,"Height":15,"Left":-476,"Top":965,"Right":-446,"Bottom":980,"TopLeft":{"X":-476,"Y":965},"TopRight":{"X":-446,"Y":965},"BottomLeft":{"X":-476,"Y":980},"BottomRight":{"X":-446,"Y":980}},"IsControlElement":true,"IsContentElement":true,"ProcessId":7768,"Orientation":0 supportedPatterns=[]
                                                properties="ControlType":{"LocalizedControlType":"pane","Id":50033,"ProgrammaticName":"ControlType.Pane"},"LocalizedControlType":"pane","IsEnabled":true,"BoundingRectangle":{"IsEmpty":false,"Location":{"X":-491,"Y":980},"Size":{"IsEmpty":false,"Width":60,"Height":7},"X":-491,"Y":980,"Width":60,"Height":7,"Left":-491,"Top":980,"Right":-431,"Bottom":987,"TopLeft":{"X":-491,"Y":980},"TopRight":{"X":-431,"Y":980},"BottomLeft":{"X":-491,"Y":987},"BottomRight":{"X":-431,"Y":987}},"IsControlElement":true,"IsContentElement":true,"ProcessId":7768,"Orientation":0 supportedPatterns=[]
                                            properties="ControlType":{"LocalizedControlType":"pane","Id":50033,"ProgrammaticName":"ControlType.Pane"},"LocalizedControlType":"pane","IsEnabled":true,"BoundingRectangle":{"IsEmpty":false,"Location":{"X":-431,"Y":922},"Size":{"IsEmpty":false,"Width":64,"Height":59},"X":-431,"Y":922,"Width":64,"Height":59,"Left":-431,"Top":922,"Right":-367,"Bottom":981,"TopLeft":{"X":-431,"Y":922},"TopRight":{"X":-367,"Y":922},"BottomLeft":{"X":-431,"Y":981},"BottomRight":{"X":-367,"Y":981}},"IsControlElement":true,"IsContentElement":true,"ProcessId":7768,"Orientation":0 supportedPatterns=[]
                                                properties="ControlType":{"LocalizedControlType":"pane","Id":50033,"ProgrammaticName":"ControlType.Pane"},"LocalizedControlType":"pane","IsEnabled":true,"BoundingRectangle":{"IsEmpty":true,"Location":{"X":"\u221E","Y":"\u221E"},"Size":{"IsEmpty":true,"Width":"-\u221E","Height":"-\u221E"},"X":"\u221E","Y":"\u221E","Width":"-\u221E","Height":"-\u221E","Left":"\u221E","Top":"\u221E","Right":"-\u221E","Bottom":"-\u221E","TopLeft":{"X":"\u221E","Y":"\u221E"},"TopRight":{"X":"-\u221E","Y":"\u221E"},"BottomLeft":{"X":"\u221E","Y":"-\u221E"},"BottomRight":{"X":"-\u221E","Y":"-\u221E"}},"IsControlElement":true,"IsContentElement":true,"ProcessId":7768,"Orientation":0 supportedPatterns=[]
                                                properties="ControlType":{"LocalizedControlType":"text","Id":50020,"ProgrammaticName":"ControlType.Text"},"LocalizedControlType":"text","IsEnabled":true,"BoundingRectangle":{"IsEmpty":false,"Location":{"X":-401,"Y":922},"Size":{"IsEmpty":false,"Width":4,"Height":13},"X":-401,"Y":922,"Width":4,"Height":13,"Left":-401,"Top":922,"Right":-397,"Bottom":935,"TopLeft":{"X":-401,"Y":922},"TopRight":{"X":-397,"Y":922},"BottomLeft":{"X":-401,"Y":935},"BottomRight":{"X":-397,"Y":935}},"IsControlElement":true,"IsContentElement":true,"ProcessId":7768,"Orientation":0 supportedPatterns=[]
                                                properties="ControlType":{"LocalizedControlType":"button","Id":50000,"ProgrammaticName":"ControlType.Button"},"LocalizedControlType":"button","Name":"clear all","IsEnabled":true,"BoundingRectangle":{"IsEmpty":false,"Location":{"X":-414,"Y":935},"Size":{"IsEmpty":false,"Width":30,"Height":30},"X":-414,"Y":935,"Width":30,"Height":30,"Left":-414,"Top":935,"Right":-384,"Bottom":965,"TopLeft":{"X":-414,"Y":935},"TopRight":{"X":-384,"Y":935},"BottomLeft":{"X":-414,"Y":965},"BottomRight":{"X":-384,"Y":965}},"IsControlElement":true,"IsContentElement":true,"ProcessId":7768,"Orientation":0 supportedPatterns=["InvokePatternIdentifiers.Pattern"]
                                                properties="ControlType":{"LocalizedControlType":"text","Id":50020,"ProgrammaticName":"ControlType.Text"},"LocalizedControlType":"text","Name":"clear all","IsEnabled":true,"BoundingRectangle":{"IsEmpty":false,"Location":{"X":-421,"Y":965},"Size":{"IsEmpty":false,"Width":43,"Height":15},"X":-421,"Y":965,"Width":43,"Height":15,"Left":-421,"Top":965,"Right":-378,"Bottom":980,"TopLeft":{"X":-421,"Y":965},"TopRight":{"X":-378,"Y":965},"BottomLeft":{"X":-421,"Y":980},"BottomRight":{"X":-378,"Y":980}},"IsControlElement":true,"IsContentElement":true,"ProcessId":7768,"Orientation":0 supportedPatterns=[]
                                                properties="ControlType":{"LocalizedControlType":"pane","Id":50033,"ProgrammaticName":"ControlType.Pane"},"LocalizedControlType":"pane","IsEnabled":true,"BoundingRectangle":{"IsEmpty":false,"Location":{"X":-431,"Y":980},"Size":{"IsEmpty":false,"Width":64,"Height":7},"X":-431,"Y":980,"Width":64,"Height":7,"Left":-431,"Top":980,"Right":-367,"Bottom":987,"TopLeft":{"X":-431,"Y":980},"TopRight":{"X":-367,"Y":980},"BottomLeft":{"X":-431,"Y":987},"BottomRight":{"X":-367,"Y":987}},"IsControlElement":true,"IsContentElement":true,"ProcessId":7768,"Orientation":0 supportedPatterns=[]
                                        properties="ControlType":{"LocalizedControlType":"pane","Id":50033,"ProgrammaticName":"ControlType.Pane"},"LocalizedControlType":"pane","IsEnabled":true,"BoundingRectangle":{"IsEmpty":false,"Location":{"X":-731,"Y":986},"Size":{"IsEmpty":false,"Width":364,"Height":1},"X":-731,"Y":986,"Width":364,"Height":1,"Left":-731,"Top":986,"Right":-367,"Bottom":987,"TopLeft":{"X":-731,"Y":986},"TopRight":{"X":-367,"Y":986},"BottomLeft":{"X":-731,"Y":987},"BottomRight":{"X":-367,"Y":987}},"IsControlElement":true,"IsContentElement":true,"ProcessId":7768,"Orientation":0 supportedPatterns=[]
                                properties="ControlType":{"LocalizedControlType":"pane","Id":50033,"ProgrammaticName":"ControlType.Pane"},"LocalizedControlType":"pane","IsEnabled":true,"BoundingRectangle":{"IsEmpty":false,"Location":{"X":-741,"Y":987},"Size":{"IsEmpty":false,"Width":384,"Height":52},"X":-741,"Y":987,"Width":384,"Height":52,"Left":-741,"Top":987,"Right":-357,"Bottom":1039,"TopLeft":{"X":-741,"Y":987},"TopRight":{"X":-357,"Y":987},"BottomLeft":{"X":-741,"Y":1039},"BottomRight":{"X":-357,"Y":1039}},"IsControlElement":true,"IsContentElement":true,"ProcessId":7768,"Orientation":0 supportedPatterns=[]
                                    properties="ControlType":{"LocalizedControlType":"pane","Id":50033,"ProgrammaticName":"ControlType.Pane"},"LocalizedControlType":"pane","IsEnabled":true,"BoundingRectangle":{"IsEmpty":false,"Location":{"X":-741,"Y":987},"Size":{"IsEmpty":false,"Width":384,"Height":52},"X":-741,"Y":987,"Width":384,"Height":52,"Left":-741,"Top":987,"Right":-357,"Bottom":1039,"TopLeft":{"X":-741,"Y":987},"TopRight":{"X":-357,"Y":987},"BottomLeft":{"X":-741,"Y":1039},"BottomRight":{"X":-357,"Y":1039}},"IsControlElement":true,"IsContentElement":true,"ProcessId":7768,"Orientation":0 supportedPatterns=[]
                                        properties="ControlType":{"LocalizedControlType":"pane","Id":50033,"ProgrammaticName":"ControlType.Pane"},"LocalizedControlType":"pane","IsEnabled":true,"BoundingRectangle":{"IsEmpty":false,"Location":{"X":-741,"Y":987},"Size":{"IsEmpty":false,"Width":63,"Height":52},"X":-741,"Y":987,"Width":63,"Height":52,"Left":-741,"Top":987,"Right":-678,"Bottom":1039,"TopLeft":{"X":-741,"Y":987},"TopRight":{"X":-678,"Y":987},"BottomLeft":{"X":-741,"Y":1039},"BottomRight":{"X":-678,"Y":1039}},"IsControlElement":true,"IsContentElement":true,"ProcessId":7768,"Orientation":0 supportedPatterns=[]
                                        properties="ControlType":{"LocalizedControlType":"button","Id":50000,"ProgrammaticName":"ControlType.Button"},"LocalizedControlType":"button","Name":"Invite (Alt\u002BI)","IsEnabled":true,"BoundingRectangle":{"IsEmpty":false,"Location":{"X":-678,"Y":1001},"Size":{"IsEmpty":false,"Width":45,"Height":24},"X":-678,"Y":1001,"Width":45,"Height":24,"Left":-678,"Top":1001,"Right":-633,"Bottom":1025,"TopLeft":{"X":-678,"Y":1001},"TopRight":{"X":-633,"Y":1001},"BottomLeft":{"X":-678,"Y":1025},"BottomRight":{"X":-633,"Y":1025}},"IsControlElement":true,"IsContentElement":true,"ProcessId":7768,"Orientation":0 supportedPatterns=["InvokePatternIdentifiers.Pattern"]
                                        properties="ControlType":{"LocalizedControlType":"pane","Id":50033,"ProgrammaticName":"ControlType.Pane"},"LocalizedControlType":"pane","IsEnabled":true,"BoundingRectangle":{"IsEmpty":false,"Location":{"X":-633,"Y":987},"Size":{"IsEmpty":false,"Width":63,"Height":52},"X":-633,"Y":987,"Width":63,"Height":52,"Left":-633,"Top":987,"Right":-570,"Bottom":1039,"TopLeft":{"X":-633,"Y":987},"TopRight":{"X":-570,"Y":987},"BottomLeft":{"X":-633,"Y":1039},"BottomRight":{"X":-570,"Y":1039}},"IsControlElement":true,"IsContentElement":true,"ProcessId":7768,"Orientation":0 supportedPatterns=[]
                                        properties="ControlType":{"LocalizedControlType":"button","Id":50000,"ProgrammaticName":"ControlType.Button"},"LocalizedControlType":"button","Name":"Mute All (Alt\u002BM)","IsEnabled":true,"BoundingRectangle":{"IsEmpty":false,"Location":{"X":-570,"Y":1001},"Size":{"IsEmpty":false,"Width":61,"Height":24},"X":-570,"Y":1001,"Width":61,"Height":24,"Left":-570,"Top":1001,"Right":-509,"Bottom":1025,"TopLeft":{"X":-570,"Y":1001},"TopRight":{"X":-509,"Y":1001},"BottomLeft":{"X":-570,"Y":1025},"BottomRight":{"X":-509,"Y":1025}},"IsControlElement":true,"IsContentElement":true,"ProcessId":7768,"Orientation":0 supportedPatterns=["InvokePatternIdentifiers.Pattern"]
                                        properties="ControlType":{"LocalizedControlType":"pane","Id":50033,"ProgrammaticName":"ControlType.Pane"},"LocalizedControlType":"pane","IsEnabled":true,"BoundingRectangle":{"IsEmpty":false,"Location":{"X":-509,"Y":987},"Size":{"IsEmpty":false,"Width":63,"Height":52},"X":-509,"Y":987,"Width":63,"Height":52,"Left":-509,"Top":987,"Right":-446,"Bottom":1039,"TopLeft":{"X":-509,"Y":987},"TopRight":{"X":-446,"Y":987},"BottomLeft":{"X":-509,"Y":1039},"BottomRight":{"X":-446,"Y":1039}},"IsControlElement":true,"IsContentElement":true,"ProcessId":7768,"Orientation":0 supportedPatterns=[]
                                        properties="ControlType":{"LocalizedControlType":"split button","Id":50031,"ProgrammaticName":"ControlType.SplitButton"},"LocalizedControlType":"split button","Name":"More options to manage all participants","IsEnabled":true,"BoundingRectangle":{"IsEmpty":false,"Location":{"X":-446,"Y":1001},"Size":{"IsEmpty":false,"Width":25,"Height":24},"X":-446,"Y":1001,"Width":25,"Height":24,"Left":-446,"Top":1001,"Right":-421,"Bottom":1025,"TopLeft":{"X":-446,"Y":1001},"TopRight":{"X":-421,"Y":1001},"BottomLeft":{"X":-446,"Y":1025},"BottomRight":{"X":-421,"Y":1025}},"IsControlElement":true,"IsContentElement":true,"ProcessId":7768,"Orientation":0 supportedPatterns=["InvokePatternIdentifiers.Pattern"]
                                        properties="ControlType":{"LocalizedControlType":"pane","Id":50033,"ProgrammaticName":"ControlType.Pane"},"LocalizedControlType":"pane","IsEnabled":true,"BoundingRectangle":{"IsEmpty":false,"Location":{"X":-421,"Y":987},"Size":{"IsEmpty":false,"Width":64,"Height":52},"X":-421,"Y":987,"Width":64,"Height":52,"Left":-421,"Top":987,"Right":-357,"Bottom":1039,"TopLeft":{"X":-421,"Y":987},"TopRight":{"X":-357,"Y":987},"BottomLeft":{"X":-421,"Y":1039},"BottomRight":{"X":-357,"Y":1039}},"IsControlElement":true,"IsContentElement":true,"ProcessId":7768,"Orientation":0 supportedPatterns=[]
                 */

                aeParticipantsWindow = AutomationElement.FromHandle(hWndParticipants);

                // Invalidate previous AE objects
                aeParticipantsList = null;
            }

            int nRefsUpdated = 0;
            if (aeParticipantsList == null)
            {
                var sw = Stopwatch.StartNew();

                /*
                 * The relevant UITree here is:
                 *
                 * == Zoom v5.2.0 ==
                 * Window - Participants
                 *   Pane
                 *     Pane
                 *       List - Waiting AND Attending
                 *         "Waiting Room (1), expanded"
                 *         WaitingParticpant
                 *         "In the Meeting (1), collapsed"
                 *         AttendingParticipant
                 *
                 * == Zoom v5.1.1 ==
                 * Window - Participants
                 *   Pane
                 *     Pane
                 *       List - Waiting
                 *   Pane
                 *     Pane
                 *       List - Attending
                 *
                 * If there is nobody waiting, it is simply:
                 * Window - Participants
                 *   Pane
                 *     Pane
                 *       List - Attending
                 *
                 */

                var aePanes = aeParticipantsWindow.FindAll(TreeScope.Children, new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Pane));
                if (aePanes.Count == 1)
                {
                    if (aeParticipantsList == null)
                    {
                        aeParticipantsList = aePanes[0]
                            .FindFirst(TreeScope.Children, new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Pane))
                                    .FindFirst(TreeScope.Children, new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.List));
                        nRefsUpdated++;
                    }
                }
                else
                {
                    throw new Exception($"Found {aePanes.Count} Participant panes when 1 was expected");
                }

                sw.Stop();
                hostApp.Log(LogType.DBG, "UpdateParticipants - Updated ref for {0}/{1} list(s) in {2:0.000}s", nRefsUpdated, aePanes.Count, sw.Elapsed.TotalSeconds);
            }

            // Sanity checks
            if (aeParticipantsList == null)
            {
                hostApp.Log(LogType.DBG, "SKIPPED - Do not have attending list object yet");
                return false;
            }
            else if (aeParticipantsList.Current.Name != "participant list, use arrow key to navigate, and press tab for more options")
            {
                throw new Exception($"Participants Attending List has unexpected name {aeParticipantsList.Current.Name}");
            }

            bool bChanged = false;

            Dictionary<string, Controller.Participant> oldList = Controller.participants;

            // === PARSE MEETING PARTICIPANTS ===

            DateTime dtStart;
            int nActualParticipants = -1;
            int nParsedParticipants = 0;
            int nAttempt = 0;
            int nPage;
            double nSecs;
            double nRate;

            participants = new Dictionary<string, Participant>();

            dtStart = DateTime.UtcNow;

            hostApp.Log(LogType.DBG, ".===== GetParticipantsFromList BEGIN =====");

            // Only list items shown on the screen are returned when we query the participant list, so we:
            //    1. Start at the end of the list (press {END})
            //    2. Hit PgUp and parse each page until
            //   3a. Either we parse the "Me" participant (Always at the top of the list) - or -
            //   3b. We didn't parse any new paritipants for two pages (It's possible we lost focus or something else strange happened)

            // The window title has the # of participants in it; Parse out the number
            var sWindowTitle = WindowTools.GetWindowText(hWndParticipants);
            var m = ReZoomParticipantsWindowTitle.Match(sWindowTitle);
            if (!m.Success)
            {
                throw new FormatException(string.Format("Unable to parse Participants window text {0}", repr(sWindowTitle)));
            }
            nActualParticipants = int.Parse(m.Groups[1].Value);

            while (true)
            {
                // Attempt to parse all of the participant pages.  If we don't get the expected count, try again - up to 3 times

                nAttempt++;

                bool foundMe = false;
                nPage = 0;
                int nNoNewParseCount = 0;
                while (true)
                {
                    // Parse this page of participants

                    nPage++;

                    var nLastParsedCount = participants.Count;

                    WindowTools.FocusWindow(hWndParticipants);

                    if (!cfg.DisableParticipantPaging)
                    {
                        WindowTools.SendKeys(IntPtr.Zero, nPage == 1 ? "{HOME}" : "{PGDN}");
                    }

                    // TBD: If the selected item is the same after we PGDN, then we're done, right?

                    WalkParticipantList(aeParticipantsList, out foundMe);
                    nParsedParticipants = participants.Count;

                    var nNewParsedCount = nParsedParticipants - nLastParsedCount;
                    hostApp.Log(LogType.DBG, ">===== GetParticipantsFromList : Page {0}, {1} new item(s) parsed", nPage, nNewParsedCount);


                    // This no longer works
                    /*
                    if (foundMe)
                    {
                        // We're searching from the bottom to the top, and my attendee list item will always be at the top, so we've hit the top of the list
                        hostApp.Log(LogType.DBG, "End of List : Found my own attendee entry");
                        break;
                    }
                    */
                    if (nParsedParticipants >= nActualParticipants)
                    {
                        // We got the number of attendees we were looking for.  If we got more than expected, that's fine -- the list changes in place, unfortunately, so we can expect new attendees can join in the middle of our parsing
                        hostApp.Log(LogType.DBG, "End of List : Got expected number of attendees");
                        break;
                    }
                    else if (nNewParsedCount == 0)
                    {
                        // TBD: Only need one of these?

                        // We didn't parse anything new, so we're probably on the same page.  If we got here, we don't have "me", which should not happen.
                        //   Try to make the best of it anyway
                        if (++nNoNewParseCount == 2)
                        {
                            hostApp.Log(LogType.WRN, "End of List : No new attendees parsed");
                            break;
                        }

                        hostApp.Log(LogType.WRN, "No new attendees parsed this page; Will try one more page");
                    }
                }

                // If for some reason I could not parse out my own Participant list item, then restore the old "me" object
                /*
                if ((me == null) && (oldMe != null))
                {
                    hostApp.Log(LogType.WRN, "I didn't find my own attendee entry this pass; Restoring my old participant object");
                    me = oldMe;
                }
                */

                // The counts match; We're good to go.  TBD: If there are multiple participants with the exact same name, this can throw us off... how to handle?
                //   If we parsed more than there are attending, no biggie... a participant probably left while we were parsing
                if (nParsedParticipants >= nActualParticipants)
                {
                    break;
                }

                // Check counts within tolerance (10%)
                if (nParsedParticipants >= (int)(0.9 * nActualParticipants))
                {
                    hostApp.Log(LogType.WRN, $"Parsed participant count {nParsedParticipants} does not match expected count {nActualParticipants}, but within tolerance (10%); Rollin' with it...");
                    break;
                }

                if (!cfg.DisableParticipantPaging)
                {
                    if (nParsedParticipants < nActualParticipants)
                    {
                        hostApp.Log(LogType.WRN, $"Parsed participant count {nParsedParticipants} does not match expected count {nActualParticipants}, but DisableParticipantPaging=true; Rollin' with it...");
                    }

                    break;
                }

                // Sometimes items move around while we're trying to parse them, resulting in only a partial list of participants.  Use the count of participants
                //   in the window title as a sanity check.  If the counts do not match, then scan again until they do, or until we time out
                var nMaxAttempts = cfg.ParticipantCountMismatchRetries > 1 ? cfg.ParticipantCountMismatchRetries : 1;
                var sLogMsg = $"UpdateParticipants: Attempt #{nAttempt}/{nMaxAttempts} - Parsed participant count {nParsedParticipants} does not match expected count {nActualParticipants}";
                if (nAttempt >= nMaxAttempts)
                {
                    hostApp.Log(LogType.WRN, "{0}; Giving up", sLogMsg);

                    // Just roll with what we have ...
                    break;
                }

                hostApp.Log(LogType.WRN, "{0}; Trying again", sLogMsg);
            }

            nSecs = (DateTime.UtcNow - dtStart).TotalSeconds;
            nRate = (nSecs == 0) ? 0 : nParsedParticipants / nSecs;
            hostApp.Log(LogType.DBG, "`===== GetParticipantsFromList END =====' {0}/{1} attendee(s) in {2:0.000}s ({3:0.00}/s); {4} attempt(s)", nParsedParticipants, nActualParticipants, nSecs, nRate, nAttempt);

            // ====== WAITING LIST ======

            foreach (Participant p in oldList.Values)
            {
                // Leaving events
                if (!participants.ContainsKey(p.name))
                {
                    ParticipantEventArgs e = new ParticipantEventArgs
                    {
                        participant = p,
                    };
                    p.status = ParticipantStatus.Leaving;
                    ParticipantAttendanceStatusChange(null, e);
                    bChanged = true;
                }
            }

            foreach (Participant p in participants.Values)
            {
                bool bNew = !oldList.TryGetValue(p.name, out Participant oldParticipant);
                //if (bNew || (!bNew && (p.status != oldParticipant.status)))
                if (bNew || (p.status != oldParticipant.status))
                {
                    // New participant or status changed
                    if (p.status == ParticipantStatus.Waiting)
                    {
                        p.dtWaiting = dtNow;
                    }
                    else if (p.status == ParticipantStatus.Attending)
                    {
                        p.dtAttending = dtNow;
                    }

                    ParticipantEventArgs e = new ParticipantEventArgs
                    {
                        participant = p,
                    };
                    ParticipantAttendanceStatusChange(null, e);
                    bChanged = true;
                }
            }

            //Dictionary<string, ZoomMeetingBotSDK.Participant> newList = new Dictionary<string, ZoomMeetingBotSDK.Participant>(ZoomMeetingBotSDK.participants);
            hostApp.Log(LogType.DBG, "UpdateParticipants - Exit");
            return bChanged;
        }

        /// <summary>
        /// Waits for a pop-up menu to appear, then returns an AutomationElement containing it.  Since searching the UIA tree structure is expensive, we use a little trick
        /// here: All pop-up menus have a specific title and class name, so we can just wait for the window to appear, then use it's handle to retrieve the AE.
        /// </summary>
        private static AutomationElement WaitPopupMenu(int timeout = 2000)
        {
            // TBD: Could add ae as input, search child windows only
            //return ae.FindFirst(TreeScope.Children, new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Menu));
            return AutomationElement.FromHandle(WindowTools.WaitWindow(SZoomPopUpMenuWindowClass, ReZoomPopUpMenuTitle, out _, timeout));
        }

        private static bool ClickMoreOptionsToManageAllParticipants(bool required = true)
        {
            return ClickParticipantWindowControl("More options to manage all participants", required, ControlType.SplitButton);
        }

        private static bool ClickParticipantItemControl(Participant p, ControlType type, string target, bool required = true)
        {
            var aeParticipant = p._ae;
            try
            {
                // Select the user.  This populates the buttons
                ((SelectionItemPattern)aeParticipant.GetCurrentPattern(SelectionItemPattern.Pattern)).Select();

                // Get the button AE
                // TBD: We know the structure; could optimize
                Stopwatch sw = new Stopwatch();
                sw.Start();
                var aeButton = aeParticipant.FindFirst(
                    TreeScope.Subtree,
                    new AndCondition(
                        new PropertyCondition(AutomationElement.ControlTypeProperty, type),
                        new PropertyCondition(AutomationElement.NameProperty, target)));
                sw.Stop();
                hostApp.Log(LogType.DBG, "ClickParticipantItemControl - Subtree search for {0} completed in {1:0.000}s; Found={2}", repr(target), sw.ElapsedMilliseconds / 1000.0, aeButton != null);

                if ((aeButton == null) && (!required))
                {
                    return false;
                }

                // Click it

                // Sometimes patt returns null, so just return false if click is not required
                var patt = (InvokePattern)aeButton.GetCurrentPattern(InvokePattern.Pattern);
                if ((patt == null) && (!required))
                {
                    return false;
                }

                patt.Invoke();

                // Give UI a chance to process it
                Thread.Sleep(cfg.UIActionDelayMilliseconds);

                return true;
            }
            catch (Exception ex)
            {
                //hostApp.Log(LogType.WRN, "Unable to invoke {0} for participant {1} : {2}", repr(target), repr(p.name), repr(ex.Message));
                hostApp.Log(LogType.WRN, "Unable to invoke {0} for participant {1} : {2}", repr(target), repr(p.name), repr(ex.ToString()));
                hostApp.Log(LogType.WRN, "TREE {0}", UIATools.WalkRawElementsToString(p._ae));
                return false;
            }
        }

        private static readonly Dictionary<string, AutomationElement> CachedAE = new Dictionary<string, AutomationElement>();

        private static bool ClickParticipantWindowControl(string target, bool required = true, ControlType controlType = null)
        {
            if (controlType == null)
            {
                controlType = ControlType.Button;
            }

            WindowTools.FocusWindow((IntPtr)aeParticipantsWindow.Current.NativeWindowHandle);
            try
            {
                // A subtree search is *very* expensive, especially in the Participant Window.  Cache it
                var bNeedSearch = false;
                if (CachedAE.TryGetValue(target, out AutomationElement ae))
                {
                    // We got the cached AE.  Make sure it's still available, and hasn't changed to another name
                    try
                    {
                        var currentName = ae.Current.Name;
                        if (currentName != target)
                        {
                            bNeedSearch = true;
                            hostApp.Log(LogType.WRN, "ClickParticipantWindowControl - Got cached AE for {0}, but Name is {1}; Forcing search", repr(target), repr(currentName));
                        }
                    }
                    catch (ElementNotAvailableException)
                    {
                        bNeedSearch = true;
                        hostApp.Log(LogType.WRN, "ClickParticipantWindowControl - Got cached AE for {0}, but ElementNotAvailableException was thrown; Forcing search", repr(target));
                    }
                }
                else
                {
                    bNeedSearch = true;
                    hostApp.Log(LogType.WRN, "ClickParticipantWindowControl - Doing Subtree search for {0}", repr(target));
                }

                if (bNeedSearch)
                {
                    var sw = new Stopwatch();
                    sw.Start();

                    // It's a big pain to find this pane :p
                    if (aeAllParticipantsPane == null)
                    {
                        // The Participants are in the last child Pane of the window
                        var aePanes = aeParticipantsWindow.FindAll(TreeScope.Children, new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Pane));
                        var aeParticipantsPane = aePanes[aePanes.Count - 1];

                        // The Participant List pane and All Participants pane are contained within a child pane of the Participants pane (@_@)
                        var aeList = aeParticipantsPane
                            .FindFirst(TreeScope.Children, new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Pane))
                                .FindAll(TreeScope.Children, new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Pane));

                        // The All Participants pane is the last pane
                        aeAllParticipantsPane = aeList[aeList.Count - 1]
                            .FindFirst(TreeScope.Children, new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Pane));
                    }

                    /*
                    ae = aeParticipantsWindow.FindFirst(
                        TreeScope.Subtree,
                        new AndCondition(
                            new PropertyCondition(AutomationElement.ControlTypeProperty, controlType),
                            new PropertyCondition(AutomationElement.NameProperty, target)));
                    */
                    ae = aeAllParticipantsPane.FindFirst(
                        TreeScope.Children,
                        new AndCondition(
                            new PropertyCondition(AutomationElement.ControlTypeProperty, controlType),
                            new PropertyCondition(AutomationElement.NameProperty, target)));

                    sw.Stop();
                    hostApp.Log(LogType.DBG, "ClickParticipantWindowControl - Subtree search for {0} completed in {1:0.000}s; Found={2}", repr(target), sw.ElapsedMilliseconds / 1000.0, ae != null);
                    if (ae == null)
                    {
                        if (CachedAE.ContainsKey(target))
                        {
                            CachedAE.Remove(target);
                        }
                    }
                    else
                    {
                        CachedAE[target] = ae;
                    }
                }

                if ((ae == null) && (!required))
                {
                    return false;
                }

                // Click it
                ((InvokePattern)ae.GetCurrentPattern(InvokePattern.Pattern)).Invoke();

                // Give UI a chance to process it
                Thread.Sleep(cfg.UIActionDelayMilliseconds);
                //WindowTools.ClickMiddle(IntPtr.Zero, ae.Current.BoundingRectangle);

                return true;
            }
            catch (Exception ex)
            {
                hostApp.Log(LogType.WRN, "Unable to invoke {0} in participants window : {1}", repr(target), repr(ex.Message));
                hostApp.Log(LogType.WRN, "TREE {0}", UIATools.WalkRawElementsToString(aeParticipantsWindow));
                return false;
            }
        }

        /// <summary>
        /// Invokes an item in the menu.  The Menu can contain any combination of MenuItems, ListItems or CheckBoxes.
        /// </summary>
        private static void InvokeMenuItem(AutomationElement ae, string sName)
        {
            /*
            Window "Participants (3)" {1171,0,400,1047} Content KeyboardFocusable ClassName:"zPlistWndClass" hWnd:1912408254 pid:191188 patterns:[WindowPatternIdentifiers,TransformPatternIdentifiers] rids:[42,1912408254]
                === "More Options for {0}" Menu (Participant in Attending List) ===
                Menu "" {1546,106,182,226} Content ClassName:"WCN_ModelessWnd" hWnd:361367270 pid:191188 rids:[42,361367270]
                    MenuItem "Chat" {1555,119,164,22} Content pid:191188 patterns:[InvokePatternIdentifiers,ValuePatternIdentifiers] rids:[42,361367270,4,0]
                    MenuItem "" {1555,141,164,8} Content pid:191188 patterns:[InvokePatternIdentifiers] rids:[42,361367270,4,1]
                    MenuItem "Stop Video" {1555,149,164,22} Content pid:191188 patterns:[InvokePatternIdentifiers,ValuePatternIdentifiers] rids:[42,361367270,4,2]
                    MenuItem "Spotlight Video" {1555,171,164,22} Content pid:191188 patterns:[InvokePatternIdentifiers,ValuePatternIdentifiers] rids:[42,361367270,4,3]
                    MenuItem "" {1555,193,164,8} Content pid:191188 patterns:[InvokePatternIdentifiers] rids:[42,361367270,4,4]
                    MenuItem "Make Host" {1555,201,164,22} Content pid:191188 patterns:[InvokePatternIdentifiers,ValuePatternIdentifiers] rids:[42,361367270,4,5]
                    MenuItem "Make Co-Host" {1555,223,164,22} Content pid:191188 patterns:[InvokePatternIdentifiers,ValuePatternIdentifiers] rids:[42,361367270,4,6]
                    MenuItem "Rename" {1555,245,164,22} Content pid:191188 patterns:[InvokePatternIdentifiers,ValuePatternIdentifiers] rids:[42,361367270,4,7]
                    MenuItem "Put in Waiting Room" {1555,267,164,22} Content pid:191188 patterns:[InvokePatternIdentifiers,ValuePatternIdentifiers] rids:[42,361367270,4,8]
                    MenuItem "" {1555,289,164,8} Content pid:191188 patterns:[InvokePatternIdentifiers] rids:[42,361367270,4,9]
                    MenuItem "Remove" {1555,297,164,22} Content pid:191188 patterns:[InvokePatternIdentifiers,ValuePatternIdentifiers] rids:[42,361367270,4,10]
                === Participants Window "..." Menu ===
                Menu "" {1468,830,306,210} Content ClassName:"WCN_ModelessWnd" hWnd:505874422 pid:724556 rids:[42,505874422]
                    MenuItem "Reclaim host" {1477,843,288,22} Content pid:724556 patterns:[InvokePatternIdentifiers,ValuePatternIdentifiers] rids:[42,505874422,4,0]
                    CheckBox "Mute Participants upon Entry" {1477,865,288,22} Content pid:724556 patterns:[ValuePatternIdentifiers,TogglePatternIdentifiers] rids:[42,505874422,4,1]
                    CheckBox "Allow Participants to Unmute Themselves" {1477,887,288,22} Content pid:724556 patterns:[ValuePatternIdentifiers,TogglePatternIdentifiers] rids:[42,505874422,4,2]
                    CheckBox "Allow Participants to Rename Themselves" {1477,909,288,22} Content pid:724556 patterns:[ValuePatternIdentifiers,TogglePatternIdentifiers] rids:[42,505874422,4,3]
                    CheckBox "Play sound when someone joins or leaves" {1477,931,288,22} Content pid:724556 patterns:[ValuePatternIdentifiers,TogglePatternIdentifiers] rids:[42,505874422,4,4]
                    CheckBox "Enable Waiting Room" {1477,953,288,22} Content pid:724556 patterns:[ValuePatternIdentifiers,TogglePatternIdentifiers] rids:[42,505874422,4,5]
                    CheckBox "Lock Meeting" {1477,975,288,22} Content pid:724556 patterns:[ValuePatternIdentifiers,TogglePatternIdentifiers] rids:[42,505874422,4,6]
                    MenuItem "" {1477,997,288,8} Content pid:724556 patterns:[InvokePatternIdentifiers] rids:[42,505874422,4,7]
                    MenuItem "Merge to Meeting Window" {1477,1005,288,22} Content pid:724556 patterns:[InvokePatternIdentifiers,ValuePatternIdentifiers] rids:[42,505874422,4,8]
            */

            AutomationElement menu = null;

            try
            {
                hostApp.Log(LogType.DBG, "InvokeMenuItem - WaitPopupMenu");
                menu = WaitPopupMenu();

                hostApp.Log(LogType.DBG, "InvokeMenuItem - Finding MenuItem {0}", repr(sName));
                // 2020.11.11 - v5.4.1 moved menu items under a Pane object, so we have to search descendants, not just children
                //var menuItem = menu.FindFirst(TreeScope.Children, new PropertyCondition(AutomationElement.NameProperty, sName));
                var menuItem = menu.FindFirst(TreeScope.Descendants, new PropertyCondition(AutomationElement.NameProperty, sName));
                if (menuItem == null)
                {
                    throw new KeyNotFoundException("MenuItem does not exist");
                }

                hostApp.Log(LogType.DBG, "InvokeMenuItem - Clicking MenuItem");
                WindowTools.ClickMiddle(IntPtr.Zero, menuItem.Current.BoundingRectangle);

                /* Doesn't work :(
                hostApp.Log(LogType.DBG, "InvokeMenuItem - Invoking MenuItem");
                WindowTools.FocusWindow(new IntPtr(menu.Current.NativeWindowHandle));
                ((InvokePattern)menuItem.GetCurrentPattern(InvokePattern.Pattern)).Invoke();
                */

                hostApp.Log(LogType.DBG, "InvokeMenuItem - Clicked {0}", repr(menuItem.Current.Name));
            }
            catch (Exception ex)
            {
                throw new Exception(string.Format("Failed to invoke MenuItem {0}; Exception {2}; AETree {1}", repr(sName), repr(UIATools.WalkRawElementsToString(menu == null ? ae : menu)), repr(ex.ToString())), ex);
            }
        }

        public static bool MuteAll(bool bAllowUnmute)
        {
            var ret = false;

            VerifyMyRole("MuteAll", ParticipantRole.CoHost);

            try
            {
                bWaitingForChangeNameDialog = true;

                ret = ClickParticipantWindowControl("Mute All (Alt+M)", false);
                if (!ret)
                {
                    return ret;
                }

                hostApp.Log(LogType.DBG, "Waiting for Mute All confirmation dialog");
                IntPtr hDialog = WindowTools.WaitWindow(SZoomConfirmMuteAllWindowClass, ReZoomConfirmMuteAllWindowTitle, out _); // Why zChangeNameWndClass?  That's weird...
                AutomationElement aeDialog = AutomationElement.FromHandle(hDialog);

                // Get AE for "Allow Participants to Unmute Themselves" and toggle it if needed
                var aeCheckBox = aeDialog.FindFirst(
                    TreeScope.Subtree,
                    new AndCondition(
                        new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.CheckBox),
                        new PropertyCondition(AutomationElement.NameProperty, "Allow Participants to Unmute Themselves")));
                var togglePattern = (TogglePattern)aeCheckBox.GetCurrentPattern(TogglePattern.Pattern);
                var newState = bAllowUnmute ? ToggleState.On : ToggleState.Off;
                if (togglePattern.Current.ToggleState != newState)
                {
                    togglePattern.Toggle();
                }

                // Get AE for Yes button
                var aeButton = aeDialog.FindFirst(
                    TreeScope.Subtree,
                    new AndCondition(
                        new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Button),
                        new PropertyCondition(AutomationElement.NameProperty, "Yes")));

                // Click it
                ((InvokePattern)aeButton.GetCurrentPattern(InvokePattern.Pattern)).Invoke();

                // Give UI a chance to process it
                Thread.Sleep(cfg.UIActionDelayMilliseconds);

                hostApp.Log(LogType.DBG, "Successfully clicked Yes in MuteAll confirmation dialog");

                ret = true;
            }
            finally
            {
                bWaitingForChangeNameDialog = false;
            }

            return ret;
        }

        private static bool UnmuteParticipantAction(IndividualParticipantActionArgs args)
        {
            if (!VerifyMyRole("UnmuteParticipant", ParticipantRole.CoHost, false))
            {
                return false;
            }

            return ClickParticipantItemControl(args.Target, ControlType.Button, args.Target.isMe ? "Unmute" : "Ask to Unmute", false);
        }

        public static void UnmuteParticipant(Participant p)
        {
            if (!VerifyMyRole("UnmuteParticipant", ParticipantRole.CoHost, false))
            {
                return;
            }

            QueueIndividualParticipantAction(new IndividualParticipantActionArgs(
                p,
                UnmuteParticipantAction,
                null));
        }

        private static bool MuteParticipantAction(IndividualParticipantActionArgs args)
        {
            if (!VerifyMyRole("MuteParticipant", ParticipantRole.CoHost, false))
            {
                return false;
            }

            return ClickParticipantItemControl(args.Target, ControlType.Button, "Mute", false);
        }

        public static void MuteParticipant(Participant p)
        {
            if (!VerifyMyRole("MuteParticipant", ParticipantRole.CoHost, false))
            {
                return;
            }

            QueueIndividualParticipantAction(new IndividualParticipantActionArgs(
                p,
                MuteParticipantAction,
                null));
        }

        public static bool ReclaimHost()
        {
            try
            {
                if (me == null)
                {
                    throw new Exception("I don't have a Participant object for myself");
                }

                if (me.role == ParticipantRole.Host)
                {
                    throw new Exception("I'm already Host");
                }

                // If I started the meeting, the "Reclaim Host" button shows up directly in the Participants pane.  In all other scenarios, it's under the more options menu.
                //  I don't know of a way to figure out if we've started the meeting or not, so we'll just try to find and click the button first and fall back on the menu
                //  if needed.

                if (!ClickParticipantWindowControl("Reclaim Host", false))
                {
                    //if (!ClickButton(aeParticipantsWindow, "More options to manage all participants")) throw new Exception("Could not click on More options menu");
                    //if (!ClickParticipantWindowControl("More options to manage all participants", true, ControlType.SplitButton)) throw new Exception("Could not click on More options menu");
                    if (!ClickMoreOptionsToManageAllParticipants())
                    {
                        throw new Exception("Could not click on More options menu");
                    }

                    //WindowTools.SendKeys((IntPtr)aeParticipantsWindow.Current.NativeWindowHandle, "{DOWN}{ENTER}");
                    //InvokeMenuItem(hWndParticipants, "Reclaim host", "{DOWN}");
                    InvokeMenuItem(aeParticipantsWindow, "Reclaim host");
                }

                return true;
            }
            catch (Exception ex)
            {
                hostApp.Log(LogType.WRN, "Unable to reclaim host: {0}", repr(ex.ToString()));
                return false;
            }
        }

        private static bool RenameParticipantAction(IndividualParticipantActionArgs args)
        {
            var p = args.Target;
            string name = (string)args.Args;

            try
            {
                // I can always rename myself
                if (!p.isMe)
                {
                    if (p.role == ParticipantRole.Host)
                    {
                        throw new Exception(String.Format("Cannot rename Host {0}", repr(name)));
                    }
                    else if (p.role == ParticipantRole.CoHost)
                    {
                        // I have to be Host to rename a CoHost
                        VerifyMyRole("RenameParticipant", ParticipantRole.Host);
                    }
                    else
                    {
                        // I have to be CoHost to rename anyone else
                        VerifyMyRole("RenameParticipant", ParticipantRole.CoHost);
                    }
                }

                if (p.status != ParticipantStatus.Attending)
                {
                    throw new Exception("Participant is not attending");
                }

                if (!ClickParticipantItemControl(p, ControlType.SplitButton, string.Format("More options for {0}", p.name)))
                {
                    throw new Exception("Could not click on participant more options button");
                }

                // If participant is me or a co-host, "Rename" is the last menu item, otherwise it's 3rd from the bottom
                InvokeMenuItem(aeParticipantsWindow, "Rename");

                // TBD: Move into a reusable function of some sort?
                hostApp.Log(LogType.DBG, "Waiting for Rename dialog");
                IntPtr hDialog = WindowTools.WaitWindow(SZoomRenameWindowClass, ReZoomRenameWindowTitle, out _);
                WindowTools.SendText(hDialog, name);
                WindowTools.SendKeys(hDialog, "{ENTER}");

                return true;
            }
            catch (Exception ex)
            {
                hostApp.Log(LogType.WRN, "Unable to rename participant {0}: {1}", repr(p.name), repr(ex.Message));
                return false;
            }
        }

        public static bool RenameParticipant(Participant p, string name)
        {
            try
            {
                // I can always rename myself
                if (!p.isMe)
                {
                    if (p.role == ParticipantRole.Host)
                    {
                        throw new Exception(String.Format("Cannot rename Host {0}", repr(name)));
                    }
                    else if (p.role == ParticipantRole.CoHost)
                    {
                        // I have to be Host to rename a CoHost
                        VerifyMyRole("RenameParticipant", ParticipantRole.Host);
                    }
                    else
                    {
                        // I have to be CoHost to rename anyone else
                        VerifyMyRole("RenameParticipant", ParticipantRole.CoHost);
                    }
                }

                if (p.status != ParticipantStatus.Attending)
                {
                    throw new Exception("Participant is not attending");
                }

                QueueIndividualParticipantAction(new IndividualParticipantActionArgs(
                    p,
                    RenameParticipantAction,
                    name));

                return true;
            }
            catch (Exception ex)
            {
                hostApp.Log(LogType.WRN, "Unable to rename participant {0}: {1}", repr(p.name), repr(ex.Message));
                return false;
            }
        }

        private static bool VerifyMyRole(string caller, ParticipantRole desiredRole, bool isRequired = true)
        {
            string err = null;

            if (me == null)
            {
                err = $"{caller}: I don't have a Participant object for myself";
            }
            else if (!((me.role == desiredRole) || ((desiredRole == ParticipantRole.CoHost) && (me.role == ParticipantRole.Host))))
            {
                err = $"{caller}: I am {me.role}, not {desiredRole.ToString()}";
            }

            if (err == null)
            {
                return true;
            }

            if (isRequired)
            {
                throw new Exception(err);
            }

            hostApp.Log(LogType.ERR, err);
            return false;
        }

        private static bool PromoteParticipantAction(IndividualParticipantActionArgs args)
        {
            var p = args.Target;
            var newRole = ((PromoteParticipantActionArgs)args.Args).NewRole;
            var promotionOption = (newRole == ParticipantRole.Host) ? "Make Host" : "Make Co-Host";

            try
            {
                VerifyMyRole("PromoteParticipant", ParticipantRole.Host);

                if (p.status != ParticipantStatus.Attending)
                {
                    throw new Exception("Participant is not attending");
                }

                if (p.role == newRole)
                {
                    hostApp.Log(LogType.WRN, "PromoteParticipant: Participant {0} is already {1}", repr(p.name), newRole.ToString());
                    return true;
                }

                if (!ClickParticipantItemControl(p, ControlType.SplitButton, string.Format("More options for {0}", p.name)))
                {
                    throw new Exception("Could not click on participant more options button");
                }
                
                // TBD: TEMP DEBUG
                //hostApp.Log(LogType.DBG, "CoHostParticipant AETree {0}", UIATools.WalkRawElementsToString(aeParticipantsWindow));

                // Make sure event handler knows not to squash this dialog
                bWaitingForChangeNameDialog = true;
                InvokeMenuItem(aeParticipantsWindow, promotionOption);

                // TBD: Move into a reusable function of some sort?
                hostApp.Log(LogType.DBG, "Waiting for promotion confirmation dialog");
                IntPtr hDialog = WindowTools.WaitWindow(SZoomConfirmCoHostWindowClass, ReZoomConfirmCoHostWindowTitle, out _); // Why zChangeNameWndClass?  That's weird...
                AutomationElement aeDialog = AutomationElement.FromHandle(hDialog);
                //hostApp.Log(LogType.DBG, "Dialog: {0}", UIATools.WalkRawElementsToString(aeDialog));

                // Make sure we got the dialog we were expecting
                // TBD: Could squash the dialog if it's not what we expect and loop back to wait again
                var re = (newRole == ParticipantRole.CoHost) ? ReZoomConfirmCoHostPrompt : ReZoomConfirmHostPrompt;
                var m = re.Match(aeDialog.Current.Name);
                if (m.Success == false)
                {
                    throw new Exception(string.Format("Got unexpected dialog when waiting for promotion confirmation dialog: {0}", p.name));
                }
                else if (m.Groups[1].Value != p.name)
                {
                    throw new Exception(string.Format("Got unexpected participant in promotion confirmation dialog: {0}", p.name));
                }

                // Get AE for Yes button
                var aeButton = aeDialog.FindFirst(
                    TreeScope.Subtree,
                    new AndCondition(
                        new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Button),
                        new PropertyCondition(AutomationElement.NameProperty, "Yes")));

                // Click it
                ((InvokePattern)aeButton.GetCurrentPattern(InvokePattern.Pattern)).Invoke();

                // Give UI a chance to process it
                Thread.Sleep(cfg.UIActionDelayMilliseconds);

                hostApp.Log(LogType.DBG, "Successfully clicked Yes in Promote confirmation dialog");

                return true;
            }
            catch (Exception ex)
            {
                hostApp.Log(LogType.WRN, "Unable to promote {0} to {1}: {2}", p.name, newRole.ToString(), repr(ex.Message));
                return false;
            }
            finally
            {
                bWaitingForChangeNameDialog = false;
            }
        }

        public static bool PromoteParticipant(Participant p, ParticipantRole newRole = ParticipantRole.CoHost)
        {
            try
            {
                switch (newRole)
                {
                    case ParticipantRole.CoHost:
                        break;
                    case ParticipantRole.Host:
                        break;
                    default:
                        throw new Exception("New role has to be Host or CoHost");
                }

                VerifyMyRole("PromoteParticipant", ParticipantRole.Host);

                if (p.status != ParticipantStatus.Attending)
                {
                    throw new Exception("Participant is not attending");
                }

                if (p.role == newRole)
                {
                    hostApp.Log(LogType.WRN, "PromoteParticipant: Participant {0} is already {1}", repr(p.name), newRole.ToString());
                    return true;
                }

                QueueIndividualParticipantAction(new IndividualParticipantActionArgs(
                    p,
                    PromoteParticipantAction,
                    new PromoteParticipantActionArgs()
                    {
                        NewRole = newRole,
                    }));

                return true;
            }
            catch (Exception ex)
            {
                hostApp.Log(LogType.WRN, "Unable to promote {0} to {1}: {2}", p.name, newRole.ToString(), repr(ex.Message));
                return false;
            }
            finally
            {
                bWaitingForChangeNameDialog = false;
            }
        }

        private static bool DemoteParticipantAction(IndividualParticipantActionArgs args)
        {
            var p = args.Target;

            try
            {
                VerifyMyRole("DemoteParticipant", ParticipantRole.Host);

                if (p.status != ParticipantStatus.Attending)
                {
                    throw new Exception("Participant is not attending");
                }

                if (p.role != ParticipantRole.CoHost)
                {
                    //throw new Exception("Participant is not Co-Host");
                    hostApp.Log(LogType.WRN, "DemoteParticipant: Participant {0} is {1}, not Co-Host", repr(p.name), p.role.ToString());
                    return true;
                }

                if (!ClickParticipantItemControl(p, ControlType.SplitButton, string.Format("More options for {0}", p.name)))
                {
                    throw new Exception("Could not click on participant more options button");
                }

                // Make sure event handler knows not to squash this dialog
                bWaitingForChangeNameDialog = true;
                InvokeMenuItem(aeParticipantsWindow, "Withdraw Co-Host Permission");

                return true;
            }
            catch (Exception ex)
            {
                hostApp.Log(LogType.WRN, "Unable to demote {0}: {1}", p.name, repr(ex.Message));
                return false;
            }
            finally
            {
                bWaitingForChangeNameDialog = false;
            }
        }

        public static bool DemoteParticipant(Participant p)
        {
            try
            {
                VerifyMyRole("DemoteParticipant", ParticipantRole.Host);

                if (p.status != ParticipantStatus.Attending)
                {
                    throw new Exception("Participant is not attending");
                }

                if (p.role != ParticipantRole.CoHost)
                {
                    hostApp.Log(LogType.WRN, "DemoteParticipant: Participant {0} is {1}, not Co-Host", repr(p.name), p.role.ToString());
                    return true;
                }

                QueueIndividualParticipantAction(new IndividualParticipantActionArgs(
                    p,
                    DemoteParticipantAction,
                    null));

                return true;
            }
            catch (Exception ex)
            {
                hostApp.Log(LogType.WRN, "Unable to co-host {0}: {1}", p.name, repr(ex.Message));
                return false;
            }
            finally
            {
                bWaitingForChangeNameDialog = false;
            }
        }

        public static bool AdmitParticipant(Participant p)
        {
            VerifyMyRole("AdmitParticipant", ParticipantRole.CoHost);

            if (p.isMe)
            {
                throw new Exception("I can't admit myself");
            }

            if (p.status != ParticipantStatus.Waiting)
            {
                throw new Exception("Participant is not waiting");
            }

            return ClickParticipantItemControl(p, ControlType.Button, "Admit");
        }

        private static EventWatcher _eventWatcher = null;

        private static void StartListeningForEvents()
        {
            _eventWatcher = new EventWatcher(nZoomPID);
        }

        private static void StopListeningForEvents()
        {
            _eventWatcher.Dispose();
        }

        public static void CalcWindowLayout()
        {
            /*
             * Example layout
             *
             * === Screen Bounds ===
             * {X=-1920,Y=0,Width=1920,Height=1080}
             *
             * === Working Area ===
             * {X=-1920,Y=0,Width=1920,Height=1040}
             *
             * === Zoom Group Chat ===
             * Rect: (-363, 0)-(7, 1047), 370x1047
             * Client: (8, 31)-(362, 1039), 354x1008
             *
             * === Participants ===
             * Rect: (-749, 0)-(-349, 1047) 400x1047
             * Client: (8, 31)-(392, 1039), 384x1008
             *
             * === Zoom Window ===
             * (-1927, 0)-(-735, 527), 1192x527
             * Client: (8, 31)-(1184, 519), 1176x488
             *
             * === Console Window ===
             * Rect: (-1927, 520)-(-735, 1047), 1192x527
             * Client: (7, 31)-(1183, 519), 1177x488
             */

            //Screen screen = Screen.AllScreens[0]; // Left-most screen

            Screen screen = (cfg.Screen == null) ? Screen.PrimaryScreen : Screen.AllScreens[int.Parse(cfg.Screen)];

            Rectangle area = screen.WorkingArea;

            hostApp.Log(LogType.DBG, "=== CalcWindowLayout === ");
            hostApp.Log(LogType.DBG, "Screen: {0}", repr(screen));
            hostApp.Log(LogType.DBG, "Device Name: {0}", screen.DeviceName);
            hostApp.Log(LogType.DBG, "Bounds: {0}", screen.Bounds.ToString());
            hostApp.Log(LogType.DBG, "Area: {0}", area.ToString());

            int nBorderOffset = 7; // Windows have some invisible border around them; This compensates

            int nChatWndWidth = 370; // 354
            int nChatWndWidthOfs = nBorderOffset;
            int nPartWndHeightOfs = nBorderOffset;
            int nPartWndWidth = 400; // 384
            int nPartWndWidthOfs = nBorderOffset;
            int nChatWndHeightOfs = nBorderOffset;
            int nZoomWndWidthOfs = nBorderOffset;
            int nZoomWndHeightOfs = nBorderOffset;
            int nThisWndWidthOfs = nBorderOffset;
            int nThisWndHeightOfs = nBorderOffset;

            int nAreaLeft = area.X; // 0
            int nAreaTop = area.Y; // 0
            int nAreaRight = area.X + area.Width; // 0
            int nAreaHeight = area.Height - area.Y; // 1040
            int nAreaWidth = area.Width; // 1920
            int nAreaHalfHeight = nAreaHeight / 2; // 520
            int nAreaMiddle = nAreaTop + nAreaHalfHeight; // 520

            chatRect.X = nAreaRight - nChatWndWidth + nChatWndWidthOfs; // -363
            chatRect.Y = nAreaTop; // 0
            chatRect.Width = nChatWndWidth; // 370
            chatRect.Height = nAreaHeight + nChatWndHeightOfs; // 1047
            hostApp.Log(LogType.DBG, "ChatRect {0}", chatRect.ToString());

            partRect.X = chatRect.X - nPartWndWidth + (nPartWndWidthOfs * 2); // -749 -- Compensating for both sides
            partRect.Y = nAreaTop; // 0
            partRect.Width = nPartWndWidth; // 400
            partRect.Height = nAreaHeight + nPartWndHeightOfs; // 1047
            hostApp.Log(LogType.DBG, "PartRect {0}", partRect.ToString());

            zoomRect.X = nAreaLeft - nZoomWndWidthOfs; // 0
            zoomRect.Y = nAreaTop; // 0
            zoomRect.Width = nAreaWidth - Math.Abs(nAreaRight - partRect.X - nPartWndWidthOfs) + (nZoomWndWidthOfs * 2); // 1192
            zoomRect.Height = nAreaHalfHeight + nZoomWndHeightOfs; // 527
            hostApp.Log(LogType.DBG, "ZoomRect {0}", zoomRect.ToString());

            AppRect.X = nAreaLeft - nZoomWndWidthOfs; // 0
            AppRect.Y = nAreaMiddle; // 520
            AppRect.Width = nAreaWidth - Math.Abs(nAreaRight - partRect.X - nPartWndWidthOfs) + (nThisWndWidthOfs * 2); // 1192
            AppRect.Height = nAreaHeight - zoomRect.Height + nZoomWndHeightOfs + nThisWndHeightOfs; // 527
            hostApp.Log(LogType.DBG, "ThisRect {0}", AppRect.ToString());
        }

        private static IntPtr StartMeeting()
        {
            IntPtr ret;

            if (cfg.ZoomUsername != null)
            {
                ret = StartMeetingDirect();
            }
            else
            {
                ret = StartMeetingFromChrome();
            }

            return ret;
        }

        private static IntPtr StartMeetingDirect()
        {
            IntPtr hZoom = IntPtr.Zero;
            Process p = null;
            string exePath = Environment.ExpandEnvironmentVariables(cfg.ZoomExecutable);

            for (int nAttempt = 1; nAttempt <= 5; nAttempt++)
            {
                try
                {
                    // If we're trying again, sleep 15s first to allow things to settle
                    if (nAttempt > 1)
                    {
                        hostApp.Log(LogType.INF, "Waiting for a bit before trying again");
                        Thread.Sleep(15000);
                    }

                    hostApp.Log(LogType.INF, "Trying to start meeting via Zoom Launcher (Attempt #{0})", nAttempt);

                    hostApp.Log(LogType.INF, "Starting: {0}", repr(exePath));

                    p = Process.Start(exePath);

                    try
                    {
                        hZoom = WindowTools.WaitWindow(SZoomMenuWindowClass, ReZoomMenuWindowTitle, out _, 5000, 1000);

                        // If this window is found, then we're already logged in, the main window will already be up.  This is a problem since we might be logged
                        //   in under the wrong account.  Also, sometimes we're logged in but the token is expired, so when we try to join the meeting we'll get
                        //   prompted for the passcode.  Go ahead and log out.

                        hostApp.Log(LogType.WRN, "Main Window is already up; Must already be logged in.  Logging out");
                        WindowTools.SendKeys(hZoom, new StringBuilder().Insert(0, "{TAB}", 2).Append(" {UP}{ENTER}").ToString());

                        Thread.Sleep(2500);
                    }
                    catch (TimeoutException)
                    {
                        // Window wasn't found; we're good!
                    }

                    hostApp.Log(LogType.INF, "Waiting for Login window");
                    hZoom = WindowTools.WaitWindow(SZoomLoginWindowClass, ReZoomLoginWindowTitle, out _, 15000, 1000);

                    // Select "Sign In"
                    WindowTools.SendKeys(hZoom, new StringBuilder().Insert(0, "{TAB}", 2).Append("{ENTER}").ToString());

                    // Send username
                    WindowTools.SendKeys(hZoom, cfg.ZoomUsername + "{TAB}");

                    // Send password
                    WindowTools.SendKeys(hZoom, ProtectedString.Unprotect(cfg.ZoomPassword), true);

                    // Login
                    WindowTools.SendKeys(hZoom, "{ENTER}");

                    // Wait for main window
                    hostApp.Log(LogType.INF, "Waiting for Menu window");
                    hZoom = WindowTools.WaitWindow(SZoomMenuWindowClass, ReZoomMenuWindowTitle, out _, 15000, 1000);

                    Thread.Sleep(2500);

                    // Select "Join"

                    // For some reason, rapid-firing tabs here doesn't work, so let's send them with some delay
                    for (int i = 0; i < 7; i++)
                    {
                        WindowTools.SendKeys("{TAB}");
                        Thread.Sleep(250);
                    }

                    WindowTools.SendKeys(hZoom, "{ENTER}");
                    Thread.Sleep(2500);

                    // Wait for join window
                    hostApp.Log(LogType.INF, "Waiting for Join dialog");
                    hZoom = WindowTools.WaitWindow(SZoomJoinWindowClass, ReZoomJoinWindowTitle, out _, 15000, 1000);

                    // Enter Meeting ID
                    WindowTools.SendKeys(hZoom, cfg.MeetingID);

                    // Set name (if needed)
                    if (cfg.MyParticipantName != null)
                    {
                        WindowTools.SendKeys(hZoom, new StringBuilder().Insert(0, "{TAB}", 2).Append("+{HOME}").Append(cfg.MyParticipantName).ToString());
                    }

                    // Join meeting
                    WindowTools.SendKeys(hZoom, "{ENTER}");

                    // Wait for zoom meeting window
                    return Controller.WaitZoomMeetingWindow(out _);
                }
                catch (Exception ex)
                {
                    hostApp.Log(LogType.ERR, "Failed to open Zoom Launcher: {0}", repr(ex.ToString()));
                }
                finally
                {
                    if (p != null)
                    {
                        hostApp.Log(LogType.INF, "Killing Zoom Launcher process");
                        try
                        {
                            p.Kill();
                        }
                        catch (Exception ex)
                        {
                            hostApp.Log(LogType.ERR, "Failed to kill Zoom Launcher process: {0}", repr(ex.ToString()));
                        }
                        p = null;
                    }
                    if (hZoom != IntPtr.Zero)
                    {
                        hostApp.Log(LogType.INF, "Closing Zoom Launcher");
                        WindowTools.CloseWindow(hZoom);
                        hZoom = IntPtr.Zero;
                    }
                }
            }

            // Should never get here
            return IntPtr.Zero;
        }

        private static IntPtr StartMeetingFromChrome()
        {
            IntPtr hChrome = IntPtr.Zero;
            Process p = null;
            string exePath = Environment.ExpandEnvironmentVariables(cfg.BrowserExecutable);

            for (int nAttempt = 1; nAttempt <= 5; nAttempt++)
            {
                try
                {
                    // If we're trying again, sleep 15s first to allow things to settle
                    if (nAttempt > 1)
                    {
                        hostApp.Log(LogType.INF, "Waiting for a bit before trying again");
                        Thread.Sleep(15000);
                    }

                    hostApp.Log(LogType.INF, "Trying to open Zoom with Chrome (Attempt #{0})", nAttempt);

                    hostApp.Log(LogType.INF, "Starting: {0}", repr(exePath));

                    p = Process.Start(exePath, cfg.BrowserArguments);

                    hostApp.Log(LogType.INF, "Chrome Started; Finding Window");
                    hChrome = WindowTools.WaitWindow(SChromeZoomWindowClass, ReChromeZoomWindowTitle, out string sChromeTitle, 30000, 1000);
                    p = null;

                    hostApp.Log(LogType.INF, "Chrome Window Title: {0}", sChromeTitle);

                    hostApp.Log(LogType.INF, "Restoring Chrome Window");
                    WindowTools.RestoreWindow(hChrome);

                    Thread.Sleep(1000);
                    //SnapLeft(hChrome);
                    WindowTools.SetWindowSize(hChrome, zoomRect);
                    Thread.Sleep(5000);

                    if (sChromeTitle.StartsWith("Sign In "))
                    {
                        hostApp.Log(LogType.INF, "Logging in with Google");
                        hChrome = WindowTools.WaitWindow(SChromeZoomWindowClass, "My Meetings - Zoom - Google Chrome");
                        WindowTools.SendKeys(hChrome, new StringBuilder().Insert(0, "{TAB}", 9).Append("{ENTER}").ToString());
                    }

                    hostApp.Log(LogType.INF, "Starting meeting");
                    WindowTools.SendKeys(hChrome, @"{F6}");
                    Thread.Sleep(1000);

                    WindowTools.SendText("https://zoom.us/s/" + cfg.MeetingID);
                    WindowTools.SendKeys(hChrome, "{ENTER}");

                    hChrome = WindowTools.WaitWindow(SChromeZoomWindowClass, "Launch Meeting - Zoom - Google Chrome");

                    hostApp.Log(LogType.INF, "Waiting for open prompt");
                    Thread.Sleep(3000);

                    hostApp.Log(LogType.INF, "Launching Zoom");

                    // The dialog used to have [ Open ] [ Cancel ], and [ Cancel ] is selected by default.  In a recent update,
                    //   a checkbox was added before these buttons, throwing our sequencing off.  Since we start with [ Cancel ],
                    //   let's just do +{TAB} to go to [ Open ].  This should be more resilient in that anything that might get
                    //   added in the future will likely be added before the [ Open ] button.
                    //WindowTools.SendKeys(hChrome, "{TAB}{ENTER}");
                    WindowTools.SendKeys(hChrome, "+{TAB}{ENTER}");

                    return Controller.WaitZoomMeetingWindow(out _);
                }
                catch (Exception ex)
                {
                    hostApp.Log(LogType.ERR, "Failed to open Zoom: {0}", repr(ex.ToString()));
                }
                finally
                {
                    if (p != null)
                    {
                        hostApp.Log(LogType.INF, "Killing Chrome process");
                        try
                        {
                            p.Kill();
                        }
                        catch (Exception ex)
                        {
                            hostApp.Log(LogType.ERR, "Failed to kill Chrome process: {0}", repr(ex.ToString()));
                        }
                        p = null;
                    }
                    if (hChrome != IntPtr.Zero)
                    {
                        hostApp.Log(LogType.INF, "Closing Chrome");
                        WindowTools.CloseWindow(hChrome);
                        hChrome = IntPtr.Zero;
                    }
                }
            }

            // Should never get here
            return IntPtr.Zero;
        }

        public static void Start()
        {
            var bNeedStart = true;

            WindowTools.WakeScreen();

            Controller.CalcWindowLayout();
            WindowTools.SetWindowSize(Process.GetCurrentProcess().MainWindowHandle, Controller.AppRect);

            hZoomMainWindow = IntPtr.Zero;
            try
            {
                hZoomMainWindow = Controller.GetZoomMeetingWindowHandle();
            }
            catch
            {
                // pass
            }

            if (hZoomMainWindow != IntPtr.Zero)
            {
                aeZoomMainWindow = AutomationElement.FromHandle(hZoomMainWindow);
                bNeedStart = false;
                try
                {
                    aeZoomMainWindow.SetFocus();
                }
                catch
                {
                    // An error is thrown when we try to SetFocus() if the Zoom UI is running, but it's not a meeting
                }
            }

            if (bNeedStart)
            {
                hostApp.Log(LogType.INF, "Starting Zoom Meeting app");
                hZoomMainWindow = StartMeeting();
            }
            else
            {
                hostApp.Log(LogType.INF, "Found running Zoom Meeting app");
            }
            ZoomAlreadyRunning = !bNeedStart;

            if (hZoomMainWindow == IntPtr.Zero)
            {
                hostApp.Log(LogType.CRT, "Failed to start Zoom; Bailing");
                return;
            }

            aeZoomMainWindow = AutomationElement.FromHandle(hZoomMainWindow);
            if (aeZoomMainWindow == null)
            {
                hostApp.Log(LogType.CRT, "Cannot get AutomationElement for Zoom Meeting Window; Bailing");
                return;
            }

            nZoomPID = aeZoomMainWindow.Current.ProcessId;
            if (nZoomPID == 0)
            {
                hostApp.Log(LogType.CRT, "Cannot get Process ID for Zoom Meeting Window; Bailing");
                return;
            }

            // Close audio prompt if it exists
            IntPtr hAudioPrompt = IntPtr.Zero;
            try
            {
                hAudioPrompt = WindowTools.WaitWindow(SJoinAudioWindowClass, ReAnyWindowTitle, out _, 5000);
            }
            catch
            {
                // pass
            }

            if (hAudioPrompt != IntPtr.Zero)
            {
                hostApp.Log(LogType.INF, "Audio Prompt Dialog found; Closing");
                WindowTools.CloseWindow(hAudioPrompt);
                //bAudioDisabled = true;

                // Check for "Do you want to continue without audio?" prompt - happens if we join after another attendee
                try
                {
                    hAudioPrompt = WindowTools.WaitWindow(SChangeNameWindowClass, ReChangeNameWindowTitle, out _, 1000);
                }
                catch
                {
                    hAudioPrompt = IntPtr.Zero;
                }
                if (hAudioPrompt != IntPtr.Zero)
                {
                    hostApp.Log(LogType.INF, "Audio Confirm Prompt Dialog found; Closing");
                    WindowTools.CloseWindow(hAudioPrompt);
                    Thread.Sleep(500);
                }
            }

            LayoutWindows();

            StartListeningForEvents();
        }

        private void Stop()
        {
            StopListeningForEvents();
        }

        /// <summary>
        /// Leaves the meeting, optionally passing off host to another participant.
        /// </summary>
        /// <param name="newHost">Name of participant to pass off host to.  If no name given or it is null, ends the meeting for all participants.</param>
        public static void LeaveMeeting(bool endForAll = false)
        {
            try
            {
                if (hZoomMainWindow == IntPtr.Zero)
                {
                    throw new Exception("Zoom is not running");
                }

                hostApp.Log(LogType.INF, "LeaveMeeting - Sending close window message to Zoom");
                WindowTools.CloseWindow(hZoomMainWindow);

                hostApp.Log(LogType.DBG, "LeaveMeeting - Waiting for confirmation dialog");
                IntPtr hDialog = WindowTools.WaitWindow("zLeaveWndClass", "End Meeting or Leave Meeting?");
                AutomationElement aeDialog = AutomationElement.FromHandle(hDialog);
                //hostApp.Log(LogType.DBG, "LeaveMeeting - Dialog: {0}", UIATools.WalkRawElementsToString(aeDialog));

                // Get AE for buttons: Pick "End Meeting for All" if it exists and endForAll is true, otherwise pick "Leave Meeting"
                AutomationElement aeButton = null;
                string sName = "End Meeting for All";
                if (endForAll)
                {
                    aeButton = aeDialog.FindFirst(
                        TreeScope.Subtree,
                        new AndCondition(
                            new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Button),
                            new PropertyCondition(AutomationElement.NameProperty, sName)));
                }

                if (aeButton == null)
                {
                    sName = "Leave Meeting";
                    // Either the "End Meeting for All" button was not found, or endForAll == false; Pick the "Leave Meeting" button
                    aeButton = aeDialog.FindFirst(
                        TreeScope.Subtree,
                        new AndCondition(
                            new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Button),
                            new PropertyCondition(AutomationElement.NameProperty, sName)));
                }

                // Click it
                hostApp.Log(LogType.INF, "Leave Meeting - Clicking {0}", repr(sName));
                ((InvokePattern)aeButton.GetCurrentPattern(InvokePattern.Pattern)).Invoke();

                // Give Zoom a bit to close down
                hostApp.Log(LogType.INF, "Waiting for a few seconds for Zoom to close down");
                Thread.Sleep(10000);

                // Close the "Zoom Cloud Meetings" dialog
                /*
                var dialogTitle = "Zoom Cloud Meetings";
                hostApp.Log(LogType.DBG, "Leave Meeting - Waiting for {0} dialog", repr(dialogTitle));
                var hWnd = WindowTools.WaitWindow("ZPFTEWndClass", dialogTitle);
                */

                /* Doesn't work!
                hostApp.Log(LogType.DBG, "Leave Meeting - Closing Zoom dialog 0x{0:X8}", hWnd);
                WindowTools.QuitWindow(hWnd);
                WindowTools.CloseWindow(hWnd);
                */

                /*
                hostApp.Log(LogType.DBG, "Leave Meeting - Finding PID for window handle 0x{0:X8}", hWnd);
                _ = WindowTools.GetWindowThreadProcessId(hWnd, out uint pid);
                if (pid == 0)
                {
                    throw new Exception("Could not get PID for window handle");
                }

                hostApp.Log(LogType.INF, "Leave Meeting - Killing Zoom process 0x{0:X8}", pid);
                var p = System.Diagnostics.Process.GetProcessById((int)pid);
                p.Kill();
                if (!p.WaitForExit(10000))
                {
                    throw new TimeoutException("Timeout waiting for process to exit");
                }
                */

                Kill();

                hostApp.Log(LogType.INF, "Leave Meeting - Done!");
            }
            catch (Exception ex)
            {
                hostApp.Log(LogType.ERR, "Leave Meeting - Failed; Exception: {0}", repr(ex.ToString()));
            }
        }

        /*
        public static void Quit()
        {
            if (hZoomMainWindow == IntPtr.Zero)
            {
                hostApp.Log(LogType.WRN, "Quit - Zoom is not running");
                return;
            }

            hostApp.Log(LogType.INF, "Quit - Sending close window message to Zoom");
            WindowTools.CloseWindow(hZoomMainWindow);

            // TBD: Handle prompt for who to hand off control to, or end meeting for all
        }
        */

        private static void Kill()
        {
            if (nZoomPID != 0)
            {
                hostApp.Log(LogType.INF, "Kill - Killing Zoom by PID 0x{0:X8}", nZoomPID);
                try
                {
                    var p = System.Diagnostics.Process.GetProcessById(nZoomPID);
                    p.Kill();
                    if (!p.WaitForExit(10000))
                    {
                        throw new TimeoutException("Timeout waiting for process to exit");
                    }
                }
                catch (Exception ex)
                {
                    hostApp.Log(LogType.ERR, "Kill - Kill Zoom by PID failed; Exception: {0}", repr(ex.ToString()));
                }
            }

            hostApp.Log(LogType.INF, "Kill - Killing Zoom by process name");
            foreach (var p in Process.GetProcessesByName("Zoom"))
            {
                hostApp.Log(LogType.INF, "Kill - Killing Zoom by process name - Killing PID 0x{0:X8}", p.Id);
                try
                {
                    p.Kill();
                    if (!p.WaitForExit(10000))
                    {
                        throw new TimeoutException("Timeout waiting for process to exit");
                    }
                }
                catch (Exception ex)
                {
                    hostApp.Log(LogType.ERR, "Kill - Kill Zoom by process name", repr(ex.ToString()));
                }
            }

            hostApp.Log(LogType.INF, "Kill - Done!");
        }

        public static void LayoutWindows()
        {
            if (hZoomMainWindow == IntPtr.Zero)
            {
                return;
            }

            IntPtr hWnd;

            //hostApp.Log(LogType.INF, "Sizing & Moving Zoom Meeting Window");
            WindowTools.SetWindowSize(hZoomMainWindow, zoomRect);
            //Thread.Sleep(500);

            //hostApp.Log(LogType.INF, "Sizing & Moving Participant Window");
            hWnd = GetParticipantsPanelWindowHandle();
            WindowTools.SetWindowSize(hWnd, partRect);
            //Thread.Sleep(500);

            //hostApp.Log(LogType.INF, "Sizing & Moving Chat Window");
            hWnd = GetChatPanelWindowHandle();
            WindowTools.SetWindowSize(hWnd, chatRect);
            //Thread.Sleep(500);
        }

        private static void TestChat()
        {
            var hChatWnd = Controller.GetChatPanelWindowHandle();
            var aeChatWnd = AutomationElement.FromHandle(hChatWnd);

            var aeChatEdit = aeChatWnd.FindFirst(
                TreeScope.Subtree,
                new AndCondition(
                    new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Edit),
                    new PropertyCondition(AutomationElement.NameProperty, "Input chat text Type message here\u2026")));

            UIATools.LogSupportedPatterns(aeChatEdit); // InvokePatternIdentifiers,ValuePatternIdentifiers

            //var invokePattern = ((InvokePattern)aeChatEdit.GetCurrentPattern(InvokePatternIdentifiers.Pattern));

            var valuePattern = (ValuePattern)aeChatEdit.GetCurrentPattern(ValuePatternIdentifiers.Pattern);
            hostApp.Log(LogType.DBG, "valuePattern.Current={0}", repr(valuePattern.Current)); // Good news: We can read what's there!
                                                                                                           //valuePattern.SetValue("Test!"); // ERROR: The method or operation is not implemented.

            var aeChatMessageList = aeChatWnd.FindFirst(
                TreeScope.Subtree,
                new AndCondition(
                    new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.List),
                    new PropertyCondition(AutomationElement.NameProperty, "chat text list")));

            UIATools.LogSupportedPatterns(aeChatMessageList); // SelectionPatternIdentifiers

            // TBD: Get RIDs & match that way - ,-1,0  {zero-based}

            _eventWatcher.Start();
            //invokePattern.Invoke();
            //WindowTools.FocusWindow(hChatWnd);

            Rect rect = aeChatMessageList.Current.BoundingRectangle;
            //WindowTools.ClickOnPoint(IntPtr.Zero, new System.Windows.Point(rect.X + 2, rect.Y + 2));

            // If we click on rect.X,rect.Y when nothing is selected, no List Item focus event is fired
            // Once we've selected a List Item, the event is always fired
            // The first item in the list can be selected by clicking rect.X+2,rect.Y+2

            WindowTools.ClickOnPoint(IntPtr.Zero, new System.Windows.Point(rect.X, rect.Y));

            // Always returns a blank list  :'(
            //var p = ((SelectionPattern)aeChatMessageList.GetCurrentPattern(SelectionPatternIdentifiers.Pattern));
            //hostApp.Log(LogType.DBG, "Current Selection (List): {0}", repr(p.Current.GetSelection()));

            //WindowTools.SendKeys("{TAB}{END}");
            while (true)
            {
                var evt = _eventWatcher.WaitEvent(5000);
                if (evt == null)
                {
                    hostApp.Log(LogType.DBG, "Timed out");
                    break;
                }
                var aei = evt.aei;
                if (aei.ControlType == ControlType.ListItem)
                {
                    hostApp.Log(LogType.DBG, "Got {0}", UIATools.AEToString(evt.ae));
                    //var sibling = TreeWalker.RawViewWalker.GetNextSibling(evt.ae);
                    var sibling = TreeWalker.RawViewWalker.GetNextSibling(evt.ae);
                    hostApp.Log(LogType.DBG, "Got sibling {0}", UIATools.AEToString(sibling));

                    //var selectionItemPattern = ((SelectionItemPattern)evt.ae.GetCurrentPattern(SelectionItemPatternIdentifiers.Pattern));
                    //selectionItemPattern.Select(); // Does nothing -- Useless

                    // This is just a Pane with no patterns
                    //hostApp.Log(LogType.DBG, "Container for (List Item): {0}", UIATools.AEToString(selectionItemPattern.Current.SelectionContainer));

                    // Again, nothing
                    // hostApp.Log(LogType.DBG, "Current Selection (List): {0}", repr(p.Current.GetSelection()));
                }
            }
            _eventWatcher.Stop();

            hostApp.Log(LogType.DBG, "Chat Window UITree: {0}", UIATools.WalkRawElementsToString(aeChatWnd));
        }

        public static void LogAETree()
        {
            var handles = new List<(IntPtr h, string name)>
            {
                (GetZoomMeetingWindowHandle(), "Main"),
                (GetParticipantsPanelWindowHandle(), "Paticipants"),
                (GetChatPanelWindowHandle(), "Chat"),
            };

            AutomationElement ae;

            hostApp.Log(LogType.INF, "LogAETree : BEGIN");
            foreach (var (h, name) in handles)
            {
                try
                {
                    ae = AutomationElement.FromHandle(h);
                    hostApp.Log(LogType.INF, "LogAETree : {0}", UIATools.WalkRawElementsToString(ae));
                }
                catch
                {
                    continue;
                }
            }
            hostApp.Log(LogType.INF, "LogAETree : END");
        }

        public void Init(IHostApp app)
        {
            hostApp = app;
            LoadSettings();

            hostApp.SettingsChanged += HostApp_SettingsChanged;
        }

        private void HostApp_SettingsChanged(object sender, EventArgs e)
        {
            LoadSettings();
        }

        private void LoadSettings()
        {
            cfg = DeserializeJson<ControllerConfigurationSettings>(hostApp.GetSettingsAsJSON());
        }
    }
}