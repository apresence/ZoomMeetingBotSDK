// TBD:
// - Figure out how to send messages to people in the waiting room

using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using ZOOM_SDK_DOTNET_WRAP;
using static ZoomMeetingBotSDK.Utils;
using static ZoomJWT.CZoomJWT;
using System.Runtime.CompilerServices;
using System.Diagnostics;
using System.ComponentModel;

namespace ZoomMeetingBotSDK
{
    public static class Controller
    {
        /*
         * .Start()
         * .SendChatMessage(recipient, response)
         * .ReclaimHost
         * .RenameParticipant
         * .UnmuteParticipant
         * .AdmitParticipant
         * .PromoteParticipant
         * .DemoteParticipant
         * .LeaveMeeting(t/f)
         * OnChatMessageReceive(ChatEventArgs e)
         * OnParticipantAttendanceStatusChange(ParticipantEventArgs e)
         * video status
         * audio status
         ...
         */

        public class ControllerConfigurationSettings
        {
            public ControllerConfigurationSettings()
            {
                ZoomWebDomain = "https://zoom.us";
                ActionTimerRateInMS = 1000;
                ActiveAudioDebounceTimeInMS = 7000;
            }

            /// <summary>
            /// Name to use when joining the Zoom meeting.  If the default name does not match, a rename is done after joining the meeting.
            /// </summary>
            public string MyParticipantName { get; set; }

            /// <summary>
            /// ID of the meeting to join.
            /// </summary>
            public string MeetingID { get; set; }

            /// <summary>
            /// User name for Zoom account
            /// </summary>
            public string ZoomUsername { get; set; }

            /// <summary>
            /// Password for Zoom account (encrypted - can only be decrypted by the current user on the current machine)
            /// </summary>
            public string ZoomPassword { get; set; }

            /// <summary>
            /// Zoom Web Domain required by the SDK.
            /// </summary>
            public string ZoomWebDomain { get; set; }

            /// <summary>
            /// The Zoom Client SDK Key.
            /// </summary>
            public string ZoomClientSDKKey { get; set; }

            /// <summary>
            /// The Zoom Client SDK Secret.
            /// </summary>
            public string ZoomClientSDKSecret { get; set; }

            /// <summary>
            /// Rate at which to fire the action timer.
            /// </summary>
            public int ActionTimerRateInMS { get; set; }

            /// <summary>
            /// Number of seconds of guard time to use before signalling an audio change event.  This cuts down on unnecessary events.  For example instead of:
            /// activeaudio: A; None; B; None; A B; None... it will look more like talkers: A; talkers: AB; talkers: None.
            /// </summary>
            public int ActiveAudioDebounceTimeInMS { get; set; }
        }
        public static ControllerConfigurationSettings cfg = new ControllerConfigurationSettings();

        private static IHostApp hostApp = null;

        private static volatile bool StartingMeeting = false;
        private static volatile bool ShouldExit = false;

        private static CZoomSDKeDotNetWrap zoom = null;
        private static IAuthServiceDotNetWrap authService = null;
        private static IMeetingServiceDotNetWrap mtgService = null;
        private static IMeetingParticipantsControllerDotNetWrap participantController = null;
        private static IMeetingAudioControllerDotNetWrap audioController = null;
        private static IMeetingVideoControllerDotNetWrap videoController = null;
        private static IMeetingShareControllerDotNetWrap shareController = null;
        private static IMeetingWaitingRoomControllerDotNetWrap waitController = null;
        private static IMeetingChatControllerDotNetWrap chatController = null;
        private static IMeetingUIControllerDotNetWrap uiController = null;

        // External event handlers

        public static event EventHandler OnActionTimerTick = (sender, e) => { };

        public static event EventHandler OnExit = (sender, e) => { };

        public class OnChatMessageReceiveArgs : EventArgs
        {
            public Participant from;
            public Participant to;
            public string text = null;
            public bool isToEveryone = false;
        }

        public static event EventHandler<OnChatMessageReceiveArgs> OnChatMessageReceive = (sender, e) => { };

        public class OnParticipantJoinWaitingRoomArgs : EventArgs
        {
            public Participant participant;
        }

        public static event EventHandler<OnParticipantJoinWaitingRoomArgs> OnParticipantJoinWaitingRoom = (sender, e) => { };

        public class OnParticipantLeaveWaitingRoomArgs : EventArgs
        {
            public Participant participant;
        }

        public static event EventHandler<OnParticipantLeaveWaitingRoomArgs> OnParticipantLeaveWaitingRoom = (sender, e) => { };

        public class OnParticipantJoinMeetingArgs : EventArgs
        {
            public Participant participant;
        }

        public static event EventHandler<OnParticipantJoinMeetingArgs> OnParticipantJoinMeeting = (sender, e) => { };

        public class OnParticipantLeaveMeetingArgs : EventArgs
        {
            public Participant participant;
        }

        public static event EventHandler<OnParticipantLeaveMeetingArgs> OnParticipantLeaveMeeting = (sender, e) => { };

        public class OnParticipantActiveAudioChangeArgs : EventArgs
        {
            public List<Participant> activeAudioParticipants;
        }

        public static event EventHandler<OnParticipantActiveAudioChangeArgs> OnParticipantActiveAudioChange = (sender, e) => { };

        public class OnParticipantRaisedHandsChangeArgs : EventArgs
        {
            public List<Participant> raisedHandParticipants;
        }

        public static event EventHandler<OnParticipantRaisedHandsChangeArgs> OnParticipantRaisedHandsChange = (sender, e) => { };


        /* ===== Callbacks ===== */

        public static void Zoom_OnAuthenticationReturn(AuthResult ret)
        {
            hostApp.Log(LogType.DBG, $"authRet ret={ret}");
            if (ret != ZOOM_SDK_DOTNET_WRAP.AuthResult.AUTHRET_SUCCESS)
            {
                throw new Exception($"Zoom Authentication Failed; ret={ret}");
            }

            hostApp.Log(LogType.DBG, "Logging in");
            var loginParam = new LoginParam4Email()
            {
                bRememberMe = false,
                userName = cfg.ZoomUsername,
                password = ProtectedString.Unprotect(cfg.ZoomPassword),
            };
            var authParam = new LoginParam()
            {
                loginType = LoginType.LoginType_Email,
                emailLogin = loginParam,
            };

            var sdkErr = authService.Login(authParam);
            if (sdkErr != SDKError.SDKERR_SUCCESS)
            {
                throw new Exception($"authService.Login failed; rc={sdkErr}");
            }
        }

        public static void Zoom_OnLoginRet(LOGINSTATUS ret, IAccountInfo pAccountInfo)
        {
            if (pAccountInfo == null)
            {
                hostApp.Log(LogType.DBG, $"login ret={ret}");
            }
            else
            {
                hostApp.Log(LogType.DBG, $"login ret={ret} account={repr(pAccountInfo.GetDisplayName())} type={repr(pAccountInfo.GetLoginType())}");
            }

            if ((ret == LOGINSTATUS.LOGIN_PROCESSING) || (ret == LOGINSTATUS.LOGIN_IDLE))
            {
                // todo
            }
            else if (ret == LOGINSTATUS.LOGIN_SUCCESS)
            {
                hostApp.Log(LogType.DBG, "Registering callbacks");

                mtgService = zoom.GetMeetingServiceWrap();
                participantController = mtgService.GetMeetingParticipantsController();
                audioController = mtgService.GetMeetingAudioController();
                videoController = mtgService.GetMeetingVideoController();
                shareController = mtgService.GetMeetingShareController();
                waitController = mtgService.GetMeetingWaitingRoomController();
                chatController = mtgService.GetMeetingChatController();
                uiController = mtgService.GetUIController();

                // TBD: Implement all controllers, at least for logging events

                RegisterCallBacks();

                hostApp.Log(LogType.DBG, "Starting meeting");

                /*
                var startParamArgs = new StartParam4WithoutLogin()
                {
                    meetingNumber = UInt64.Parse(meetingNumber),
                    userId = userId,
                    userZAK = userToken,
                    userName = userName,
                    zoomuserType = ZoomUserType.ZoomUserType_APIUSER
                };
                var startParam = new StartParam()
                {
                    userType = SDKUserType.SDK_UT_WITHOUT_LOGIN,
                    withoutloginStart = startParamArgs
                };
                */

                var startParamArgs = new StartParam4NormalUser()
                {
                    meetingNumber = UInt64.Parse(cfg.MeetingID),
                    isAudioOff = false,
                    isVideoOff = false,
                };
                var startParam = new StartParam()
                {
                    userType = SDKUserType.SDK_UT_NORMALUSER,
                    normaluserStart = startParamArgs
                };

                var sdkErr = mtgService.Start(startParam);
                if (sdkErr != SDKError.SDKERR_SUCCESS)
                {
                    throw new Exception($"mtgService.Start failed; rc={sdkErr}");
                }
            }
            else
            {
                throw new Exception($"Zoom Login Failed; ret={ret}");
            }
        }

        public static void Zoom_OnLogout()
        {
            hostApp.Log(LogType.DBG, $"logout");
        }

        public static void Zoom_OnMeetingStatusChanged(MeetingStatus status, int iResult)
        {
            hostApp.Log(LogType.DBG, $"{MethodBase.GetCurrentMethod().Name} {status}; rc={iResult}");

            switch (status)
            {
                case MeetingStatus.MEETING_STATUS_INMEETING:
                    /*
                    var userInfo = participantController.GetUserByUserID(0);
                    var myUserID = userInfo.GetUserID();
                    hostApp.Log(LogType.DBG, $"My User ID: {myUserID}");

                    participantController.ChangeUserName(myUserID, userName, false);

                    chatController.SendChatTo(0, $"Hello from {userName}!");
                    */

                    StartingMeeting = false;
                    break;
                case MeetingStatus.MEETING_STATUS_ENDED:
                    OnExit(null, null);
                    break;
            }
        }

        [MethodImplAttribute(MethodImplOptions.NoInlining)]
        private static void HandleParticipantRoleChanged(UInt32 userId)
        {
            lock (participants)
            {
                if (!GetParticipantById(userId, out Participant p))
                {
                    return;
                }

                var role = participantController.GetUserByUserID(userId).GetUserRole(); // TBD: This may throw an exception?
                p.IsHost = role == UserRole.USERROLE_HOST;
                p.IsCoHost = role == UserRole.USERROLE_COHOST;

                hostApp.Log(LogType.INF, $"{new StackFrame(1).GetMethod().Name} p={p} newRole={role}");
            }
        }

        public static void Zoom_OnHostChangeNotification(UInt32 userId)
        {
            //hostApp.Log(LogType.DBG, $"hostChange {userId}");
            HandleParticipantRoleChanged(userId);
        }

        public static void Zoom_OnCoHostChangeNotification(UInt32 userId)
        {
            //hostApp.Log(LogType.DBG, $"coHostChange {userId}");
            HandleParticipantRoleChanged(userId);
        }

        public static void Zoom_OnChatMsgNotification(ZOOM_SDK_DOTNET_WRAP.IChatMsgInfoDotNetWrap chatMsg)
        {
            var e = new OnChatMessageReceiveArgs()
            {
                from = GetOrCreateParticipant(chatMsg.GetSenderUserId(), chatMsg.GetSenderDisplayName()),
                to = GetOrCreateParticipant(chatMsg.GetReceiverUserId(), chatMsg.GetReceiverDisplayName()),
                text = chatMsg.GetContent(),
            };

            e.isToEveryone = SpecialParticipant.TryGetValue(e.to.UserId, out _);
            hostApp.Log(LogType.DBG, $"chatMsgNotification from={e.from} to={e.to} isToEveryone={e.isToEveryone} text={repr(e.text)}");

            OnChatMessageReceive(null, e);
        }

        public static void Zoom_OnUserJoin(uint[] lstUserID)
        {
            hostApp.Log(LogType.DBG, $"{MethodBase.GetCurrentMethod().Name} ids={repr(lstUserID)}");

            foreach (var userId in lstUserID)
            {
                var p = UpdateParticipant(userId, false);
                //hostApp.Log(LogType.DBG, $"userJoin p={p}");
                OnParticipantJoinMeeting(null, new OnParticipantJoinMeetingArgs()
                {
                    participant = p,
                });
            }
        }

        public static void Zoom_OnUserLeft(uint[] lstUserID)
        {
            hostApp.Log(LogType.DBG, $"{MethodBase.GetCurrentMethod().Name} ids={repr(lstUserID)}");

            lock (participants)
            {
                foreach (var userId in lstUserID)
                {
                    if (!GetParticipantById(userId, out Participant p))
                    {
                        continue;
                    }

                    if (p.Status != ParticipantStatus.Attending)
                    {
                        hostApp.Log(LogType.WRN, $"{MethodBase.GetCurrentMethod().Name} Participant {p} status is {p.Status}, not Attending");
                        continue;
                    }

                    OnParticipantLeaveMeeting(null, new OnParticipantLeaveMeetingArgs()
                    {
                        participant = p,
                    });
                    RemoveParticipant(p);

                    hostApp.Log(LogType.INF, $"Participant {p} left the meeting");
                }
            }
        }

        public static void Zoom_OnUserNameChanged(uint userId, string userName)
        {
            lock (participants)
            {
                if (GetParticipantById(userId, out Participant p))
                {
                    hostApp.Log(LogType.DBG, $"userNameChanged p={p} newName={repr(userName)}");
                    p.Name = userName;
                }
            }
        }

        public static void Zoom_OnUserAudioStatusChange(IUserAudioStatusDotNetWrap[] lstAudioStatusChange)
        {
            //hostApp.Log(LogType.DBG, $"userAudioStatusChange count={lstAudioStatusChange.Length}");

            foreach (var audioStatusChange in lstAudioStatusChange)
            {
                var userId = audioStatusChange.GetUserId();
                var audioType = audioStatusChange.GetAudioType();
                var audioStatus = audioStatusChange.GetStatus();

                lock (participants)
                {
                    if (!GetParticipantById(userId, out Participant p))
                    {
                        continue;
                    }

                    //hostApp.Log(LogType.DBG, $"{MethodBase.GetCurrentMethod().Name} {string.Format("0x{0:X8}", Thread.CurrentThread.ManagedThreadId)} p={p} type={audioType} status={audioStatus}");

                    var currentAudioDevice = (ControllerAudioType)audioType;
                    var currentIsAudioMuted = (audioStatus == AudioStatus.Audio_Muted) || (audioStatus == AudioStatus.Audio_MutedAll_ByHost) || (audioStatus == AudioStatus.Audio_Muted_ByHost);

                    if (p.audioDevice != currentAudioDevice)
                    {
                        hostApp.Log(LogType.DBG, $"{MethodBase.GetCurrentMethod().Name} p={p} audioDevice={currentAudioDevice}");
                        p.audioDevice = currentAudioDevice;
                    }

                    if (p.IsAudioMuted != currentIsAudioMuted)
                    {
                        hostApp.Log(LogType.DBG, $"{MethodBase.GetCurrentMethod().Name} p={p} audioMuted={currentIsAudioMuted}");
                        p.IsAudioMuted = currentIsAudioMuted;
                    }

                    // TBD: Assuming this makes sense?
                    p.IsPurePhoneUser = participantController.GetUserByUserID(p.UserId).IsPurePhoneUser();
                }
            }
        }

        public static void Zoom_OnUserVideoStatusChange(uint userId, VideoStatus status)
        {
            lock (participants)
            {
                if (!GetParticipantById(userId, out Participant p))
                {
                    return;
                }

                //hostApp.Log(LogType.DBG, $"userVideoStatusChange user={p} status={status}");

                var currentVideoOn = status == VideoStatus.Video_ON;

                if (currentVideoOn != p.IsVideoOn)
                {
                    hostApp.Log(LogType.DBG, $"{MethodBase.GetCurrentMethod().Name} p={p} videoOn={currentVideoOn}");
                    p.IsVideoOn = currentVideoOn;
                }
            }
        }

        public static void Zoom_OnSharingStatus(SharingStatus status, uint userId)
        {
            hostApp.Log(LogType.DBG, $"sharingStatus id={userId} status={status}");

            /*
            lock (participants)
            {
                if (participants.TryGetValue(userId, out var p))
                {
                    p.isSharing = status == SharingStatus.
                }
            }
            */
        }

        /// <summary>
        /// Handle waiting room join event from SDK.
        /// NOTE:
        /// - When a user leaves the waiting room and joins the meeting, their userId changes.
        /// - When a user leaves the meeting (due to being put in the waiting room), their userId stays the same.  Weird.
        /// </summary>
        public static void Zoom_OnWatingRoomUserJoin(uint userId)
        {
            hostApp.Log(LogType.DBG, $"{MethodBase.GetCurrentMethod().Name} id={userId}");
            var p = UpdateParticipant(userId, true);
            OnParticipantJoinWaitingRoom(null, new OnParticipantJoinWaitingRoomArgs()
            {
                participant = p,
            });
        }

        public static void Zoom_OnWatingRoomUserLeft(uint userId)
        {
            // NOTE: Interestingly, when a user leaves the waiting room and joins the meeting room, the userId changes!
            hostApp.Log(LogType.DBG, $"{MethodBase.GetCurrentMethod().Name} id={userId}");

            if (!GetParticipantById(userId, out Participant p))
            {
                return;
            }

            if ((p.Status != ParticipantStatus.Waiting) && (p.Status != ParticipantStatus.Joining)) {
                hostApp.Log(LogType.WRN, $"{MethodBase.GetCurrentMethod().Name} Participant {p} is {p.Status}, not Waiting or Joining");
                return;
            }

            OnParticipantLeaveWaitingRoom(null, new OnParticipantLeaveWaitingRoomArgs()
            {
                participant = p,
            });
            RemoveParticipant(p);

            hostApp.Log(LogType.INF, $"Participant {p} left the waiting room");
        }

        public static void Zoom_OnLowOrRaiseHandStatusChanged(bool bLow, uint userId)
        {
            // NOTE: This does not get called when one of the hosts or co-hosts lowers ALL hands

            hostApp.Log(LogType.DBG, $"{MethodBase.GetCurrentMethod().Name} id={userId} low={bLow}");

            var list = new List<Participant>();

            lock (raisedHands)
            {
                var hasBeenRaised = !bLow;
                var wasRaisedBefore = raisedHands.Contains(userId);

                if (hasBeenRaised == wasRaisedBefore)
                {
                    // No change
                    return;
                }

                if (hasBeenRaised)
                {
                    if (!GetParticipantById(userId, out Participant p))
                    {
                        // Ignore event if we don't have the participant in our participant list
                        return;
                    }
                    raisedHands.Add(userId);
                }
                else
                {
                    raisedHands.Remove(userId);
                }

                foreach (uint x in raisedHands)
                {
                    if (GetParticipantById(x, out Participant p))
                    {
                        list.Add(p);
                    }
                }
            }

            hostApp.Log(LogType.DBG, $"{MethodBase.GetCurrentMethod().Name} list: {((list.Count == 0) ? "None" : list.ToDelimString())}");

            // This list is maintained in the order the hands were raised
            OnParticipantRaisedHandsChange(null, new OnParticipantRaisedHandsChangeArgs()
            {
                raisedHandParticipants = list,
            });
        }

        /// <summary>
        /// Applies debounce to active audio participants and fires event when the active audio list changes.
        /// </summary>
        private static void UpdateParticipantActiveAudio()
        {
            // TBD: Maybe add a throttle, like... no more frequently than once per second?

            var activityCutOffDT = DateTime.UtcNow.AddMilliseconds(-cfg.ActiveAudioDebounceTimeInMS);
            var currentActiveAudio = new HashSet<uint>();

            lock (participants)
            {
                foreach (var p in participants.Values)
                {
                    if (p.LastAudioActivityDT > activityCutOffDT)
                    {
                        currentActiveAudio.Add(p.UserId);
                    }
                }
            }

            lock (activeAudio)
            {
                if (currentActiveAudio.SetEquals(activeAudio))
                {
                    // List hasn't changed
                    return;
                }

                var list = new List<Participant>();
                activeAudio = currentActiveAudio;
                foreach (uint userId in activeAudio)
                {
                    if (GetParticipantById(userId, out Participant p))
                    {
                        list.Add(p);
                    }
                }

                // Sort by name so it's easier to eyeball changes in the logs
                list.Sort(delegate (Participant p1, Participant p2) { return p1.ToString().CompareTo(p2.ToString()); });

                hostApp.Log(LogType.DBG, $"{MethodBase.GetCurrentMethod().Name} list: {((list.Count == 0) ? "None" : list.ToDelimString())}");

                OnParticipantActiveAudioChange(null, new OnParticipantActiveAudioChangeArgs()
                {
                    activeAudioParticipants = list,
                });
            }
        }

        /// <summary>
        /// Fired when the list of participants making sound (or talking) "changes".  I put "changes" in quotes because it seems the event is fired every second or two as long as
        /// someone is talking, even if the list of activeAudio have not changed.
        /// </summary>
        public static void Zoom_OnUserActiveAudioChange(uint[] lstActiveAudio)
        {
            // This event fires too often (spammy)...
            //hostApp.Log(LogType.DBG, $"{MethodBase.GetCurrentMethod().Name} ids={repr(lstActiveAudio)}");

            var nowDT = DateTime.UtcNow;
            if (lstActiveAudio != null)
            {
                foreach (var userId in lstActiveAudio)
                {
                    if (GetParticipantById(userId, out Participant p))
                    {
                        p.LastAudioActivityDT = nowDT;
                    }
                }
            }

            UpdateParticipantActiveAudio();
        }

        public static void Zoom_OnSpotlightVideoChangeNotification(bool bSpotlight, uint userid)
        {
            hostApp.Log(LogType.DBG, $"{MethodBase.GetCurrentMethod().Name} id={userid} spotlight={bSpotlight}");
        }

        public static void Zoom_OnLockShareStatus(bool bLocked)
        {
            hostApp.Log(LogType.DBG, $"{MethodBase.GetCurrentMethod().Name} locked={bLocked}");
        }

        public static void Zoom_OnShareContentNotification(ValueType shareInfo)
        {
            hostApp.Log(LogType.DBG, $"{MethodBase.GetCurrentMethod().Name} info={repr(shareInfo)}");
        }

        public static void Zoom_OnChatStatusChangedNotification(ValueType status)
        {
            hostApp.Log(LogType.DBG, $"{MethodBase.GetCurrentMethod().Name} status={repr(status)}");
        }

        public static void Zoom_OnMeetingSecureKeyNotification(byte[] key, int len, IMeetingExternalSecureKeyHandler pHandler)
        {
            hostApp.Log(LogType.DBG, $"{MethodBase.GetCurrentMethod().Name} len={len} key={repr(key)}");
        }

        public static void Zoom_OnMeetingStatisticsWarningNotification(StatisticsWarningType type)
        {
            hostApp.Log(LogType.DBG, $"{MethodBase.GetCurrentMethod().Name} type={type}");
        }

        public static void Zoom_OnEndMeetingBtnClicked()
        {
            hostApp.Log(LogType.DBG, $"{MethodBase.GetCurrentMethod().Name}");
        }

        /*
        unsafe public static void Zoom_OnInviteBtnClicked(bool* handled)
        {
            hostApp.Log(LogType.DBG, $"{MethodBase.GetCurrentMethod().Name}");
        }
        */

        public static void Zoom_OnParticipantListBtnClicked()
        {
            hostApp.Log(LogType.DBG, $"{MethodBase.GetCurrentMethod().Name}");
        }

        public static void Zoom_OnStartShareBtnClicked()
        {
            hostApp.Log(LogType.DBG, $"{MethodBase.GetCurrentMethod().Name}");
        }

        public static void Zoom_OnZoomInviteDialogFailed()
        {
            hostApp.Log(LogType.DBG, $"{MethodBase.GetCurrentMethod().Name}");
        }

        private static void RegisterCallBacks()
        {
            mtgService.Add_CB_onMeetingSecureKeyNotification(Zoom_OnMeetingSecureKeyNotification);
            mtgService.Add_CB_onMeetingStatisticsWarningNotification(Zoom_OnMeetingStatisticsWarningNotification);
            mtgService.Add_CB_onMeetingStatusChanged(Zoom_OnMeetingStatusChanged);
            participantController.Add_CB_onHostChangeNotification(Zoom_OnHostChangeNotification);
            participantController.Add_CB_onCoHostChangeNotification(Zoom_OnCoHostChangeNotification);
            participantController.Add_CB_onLowOrRaiseHandStatusChanged(Zoom_OnLowOrRaiseHandStatusChanged);
            participantController.Add_CB_onUserJoin(Zoom_OnUserJoin);
            participantController.Add_CB_onUserLeft(Zoom_OnUserLeft);
            participantController.Add_CB_onUserNameChanged(Zoom_OnUserNameChanged);
            audioController.Add_CB_onUserAudioStatusChange(Zoom_OnUserAudioStatusChange);
            audioController.Add_CB_onUserActiveAudioChange(Zoom_OnUserActiveAudioChange);
            videoController.Add_CB_onUserVideoStatusChange(Zoom_OnUserVideoStatusChange);
            videoController.Add_CB_onSpotlightVideoChangeNotification(Zoom_OnSpotlightVideoChangeNotification);
            shareController.Add_CB_onLockShareStatus(Zoom_OnLockShareStatus);
            shareController.Add_CB_onShareContentNotification(Zoom_OnShareContentNotification);
            shareController.Add_CB_onSharingStatus(Zoom_OnSharingStatus);
            waitController.Add_CB_onWatingRoomUserJoin(Zoom_OnWatingRoomUserJoin);
            waitController.Add_CB_onWatingRoomUserLeft(Zoom_OnWatingRoomUserLeft);
            chatController.Add_CB_onChatMsgNotifcation(Zoom_OnChatMsgNotification);
            chatController.Add_CB_onChatStatusChangedNotification(Zoom_OnChatStatusChangedNotification);
            uiController.Add_CB_onEndMeetingBtnClicked(Zoom_OnEndMeetingBtnClicked);

            /*
            unsafe
            {
                uiController.Add_CB_onInviteBtnClicked(Zoom_OnInviteBtnClicked);
            }
            */

            uiController.Add_CB_onParticipantListBtnClicked(Zoom_OnParticipantListBtnClicked);
            uiController.Add_CB_onStartShareBtnClicked(Zoom_OnStartShareBtnClicked);
            uiController.Add_CB_onZoomInviteDialogFailed(Zoom_OnZoomInviteDialogFailed);
        }

        private static void UnregisterCallBacks()
        {
            mtgService.Remove_CB_onMeetingSecureKeyNotification(Zoom_OnMeetingSecureKeyNotification);
            mtgService.Remove_CB_onMeetingStatisticsWarningNotification(Zoom_OnMeetingStatisticsWarningNotification);
            mtgService.Remove_CB_onMeetingStatusChanged(Zoom_OnMeetingStatusChanged);
            participantController.Remove_CB_onHostChangeNotification(Zoom_OnHostChangeNotification);
            participantController.Remove_CB_onCoHostChangeNotification(Zoom_OnCoHostChangeNotification);
            participantController.Remove_CB_onLowOrRaiseHandStatusChanged(Zoom_OnLowOrRaiseHandStatusChanged);
            participantController.Remove_CB_onUserJoin(Zoom_OnUserJoin);
            participantController.Remove_CB_onUserLeft(Zoom_OnUserLeft);
            participantController.Remove_CB_onUserNameChanged(Zoom_OnUserNameChanged);
            audioController.Remove_CB_onUserAudioStatusChange(Zoom_OnUserAudioStatusChange);
            audioController.Remove_CB_onUserActiveAudioChange(Zoom_OnUserActiveAudioChange);
            videoController.Remove_CB_onUserVideoStatusChange(Zoom_OnUserVideoStatusChange);
            videoController.Remove_CB_onSpotlightVideoChangeNotification(Zoom_OnSpotlightVideoChangeNotification);
            shareController.Remove_CB_onLockShareStatus(Zoom_OnLockShareStatus);
            shareController.Remove_CB_onShareContentNotification(Zoom_OnShareContentNotification);
            shareController.Remove_CB_onSharingStatus(Zoom_OnSharingStatus);
            waitController.Remove_CB_onWatingRoomUserJoin(Zoom_OnWatingRoomUserJoin);
            waitController.Remove_CB_onWatingRoomUserLeft(Zoom_OnWatingRoomUserLeft);
            chatController.Remove_CB_onChatMsgNotifcation(Zoom_OnChatMsgNotification);
            chatController.Remove_CB_onChatStatusChangedNotification(Zoom_OnChatStatusChangedNotification);
            uiController.Remove_CB_onEndMeetingBtnClicked(Zoom_OnEndMeetingBtnClicked);

            /*
            unsafe
            {
                uiController.Remove_CB_onInviteBtnClicked(Zoom_OnInviteBtnClicked);
            }
            */

            uiController.Remove_CB_onParticipantListBtnClicked(Zoom_OnParticipantListBtnClicked);
            uiController.Remove_CB_onStartShareBtnClicked(Zoom_OnStartShareBtnClicked);
            uiController.Remove_CB_onZoomInviteDialogFailed(Zoom_OnZoomInviteDialogFailed);
        }

        /* ================ MAIN CODE ================ */

        private static string getJwt()
        {
            return CreateClientSDKToken(cfg.ZoomClientSDKKey, ProtectedString.Unprotect(cfg.ZoomClientSDKSecret));
        }

        private static void BeginLoginSequence()
        {
            SDKError sdkErr;

            hostApp.Log(LogType.DBG, "Creating Zoom instance");

            zoom = CZoomSDKeDotNetWrap.Instance;

            hostApp.Log(LogType.DBG, "Initializing");

            var initParam = new InitParam()
            {
                web_domain = cfg.ZoomWebDomain,
            };
            sdkErr = zoom.Initialize(initParam);
            if (sdkErr != SDKError.SDKERR_SUCCESS)
            {
                throw new Exception($"zoom.Initialize failed; rc={sdkErr}");
            }

            //Thread.Sleep(2000);

            hostApp.Log(LogType.DBG, "Authorizing");
            authService = zoom.GetAuthServiceWrap();
            authService.Add_CB_onAuthenticationReturn(Zoom_OnAuthenticationReturn);
            authService.Add_CB_onLoginRet(Zoom_OnLoginRet);
            authService.Add_CB_onLogout(Zoom_OnLogout);

            /*
            var authParam = new AuthParam()
            {
                appKey = apiKey,
                appSecret = apiSecret,
            };
            sdkErr = authService.SDKAuth(authParam);
            */
            var authCtx = new AuthContext()
            {
                jwt_token = getJwt(),
            };
            sdkErr = authService.SDKAuth(authCtx);
            if (sdkErr != SDKError.SDKERR_SUCCESS)
            {
                throw new Exception($"authService.SDKAuth failed; rc={sdkErr}");
            }
        }

        private static Participant UpdateParticipant(uint userId, bool waiting)
        {
            // TBD: Error handling
            DateTime dtNow = DateTime.UtcNow;

            IUserInfoDotNetWrap user = waiting ? waitController.GetWaitingRoomUserInfoByID(userId) : participantController.GetUserByUserID(userId);

            Participant p = GetOrCreateParticipant(userId, user.GetUserNameW(), false);

            p.audioDevice = (ControllerAudioType)user.GetAudioJoinType();
            p.IsAudioMuted = user.IsAudioMuted();
            p.IsHost = user.IsHost();
            p.IsCoHost = user.GetUserRole() == UserRole.USERROLE_COHOST;
            p.Status = waiting ? ParticipantStatus.Waiting : ParticipantStatus.Attending;
            p.IsMe = user.IsMySelf();
            p.IsPurePhoneUser = user.IsPurePhoneUser();
            p.IsVideoOn = user.IsVideoOn();
            p.IsHandRaised = user.IsRaiseHand();

            if (p.IsMe)
            {
                me = p;
            }

            if (waiting)
            {
                if (p.WaitingRoomEntryDT < dtNow)
                {
                    p.WaitingRoomEntryDT = dtNow;
                }

                p.MeetingEntryDT = DateTime.MinValue;
            }
            else
            {
                if (p.MeetingEntryDT < dtNow)
                {
                    p.MeetingEntryDT = dtNow;
                }
            }

            lock (participants) {
                if (participants.TryGetValue(userId, out Participant existing))
                {
                    if (p.Status != existing.Status)
                    {
                        hostApp.Log(LogType.INF, $"Participant {p} changed from {existing.Status} to {p.Status}");
                    }
                }
                else
                {
                    hostApp.Log(LogType.INF, $"Participant {p} joined the {(waiting ? "waiting room" : "meeting")}");
                }

                participants[userId] = p;
            }

            return p;
        }

        private static void RemoveParticipant(Participant p)
        {
            lock (participants)
            {
                if (!participants.ContainsKey(p.UserId))
                {
                    return;
                }

                participants.Remove(p.UserId);
            }
        }


        // ========== PUBLIC STUFF =============

        public static class SpecialParticipant
        {
            public static readonly Participant everyoneInMeeting = new Participant()
            {
                UserId = 0,
                Name = "Everyone (in Meeting)",
            };

            public static readonly Participant everyoneInWaitingRoom = new Participant()
            {
                UserId = 4294967295, // TBD: This doesn't actually work...
                Name = "Everyone (in Waiting Room)",
            };

            /// <summary>This method returns true if the given recipient is one of the special Everyone options; false otherwise.</summary>
            public static bool IsEveryone(string name)
            {
                return name.StartsWith("Everyone");
            }

            public static bool IsEveryone(uint userId)
            {
                return (userId == everyoneInMeeting.UserId || userId == everyoneInWaitingRoom.UserId);
            }

            public static bool IsEveryone(Participant p)
            {
                return IsEveryone(p.UserId);
            }

            /// <summary>When there is nobody in the waiting room, the "Everyone (in Meeting)" selection item is renamed to "Everyone".  This
            /// method normalizes the value.</summary>
            public static string Normalize(string name)
            {
                return name == "Everyone" ? everyoneInMeeting.Name : name;
            }

            public static bool TryGetValue(uint id, out Participant participant)
            {
                if (id == everyoneInMeeting.UserId)
                {
                    participant = everyoneInMeeting;
                    return true;
                }
                else if (id == everyoneInWaitingRoom.UserId)
                {
                    participant = everyoneInWaitingRoom;
                    return true;
                }

                participant = null;
                return false;
            }

            public static bool TryGet(string name, out Participant participant)
            {
                if (name == everyoneInMeeting.Name)
                {
                    participant = everyoneInMeeting;
                    return true;
                }
                else if (name == everyoneInWaitingRoom.Name)
                {
                    participant = everyoneInWaitingRoom;
                    return true;
                }

                participant = null;
                return false;
            }
        }

        public enum ParticipantStatus
        {
            Unknown = -1,
            Waiting = 0,
            Joining = 1,
            Attending = 2,
        }

        public enum ParticipantRole
        {
            None = UserRole.USERROLE_NONE,
            Host = UserRole.USERROLE_HOST,
            CoHost = UserRole.USERROLE_COHOST
        }

        public enum ParticipantAudioDevice
        {
            Unknown = AudioType.AUDIOTYPE_UNKNOW,
            Computer = AudioType.AUDIOTYPE_VOIP,
            Telephone = AudioType.AUDIOTYPE_PHONE
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

        public enum ControllerAudioType // From SDK AudioType.  We redefine it here so the calling assembly does not have to reference the SDK library directly
        {
            AUDIOTYPE_NONE = 0,
            AUDIOTYPE_VOIP = 1,
            AUDIOTYPE_PHONE = 2,
            AUDIOTYPE_UNKNOWN = 3
        }

        public class Participant
        {
            public uint UserId = 0;
            public string Name = null;
            public ControllerAudioType audioDevice = ControllerAudioType.AUDIOTYPE_UNKNOWN;
            //public UserRole role = UserRole.USERROLE_NONE; // TBD: No events other than onHostChange and onCoHostChange
            public bool IsAudioMuted = false;
            public bool IsHost = false;
            public bool IsCoHost = false;
            //public bool isInWaitingRoom = false;
            public ParticipantStatus Status = ParticipantStatus.Unknown;
            public bool IsMe = false;
            public bool IsPurePhoneUser = false; // TBD: No event for this (?)
            public bool IsHandRaised = false;
            public bool IsVideoOn = false;

            public DateTime LastAudioActivityDT = DateTime.MinValue;
            public DateTime WaitingRoomEntryDT = DateTime.MinValue;
            public DateTime MeetingEntryDT = DateTime.MinValue;

            public override string ToString()
            {
                return $"{repr(Name)}#{UserId}";
            }
        }

        public static Dictionary<uint, Participant> participants = new Dictionary<uint, Participant>();

        public static List<uint> raisedHands = new List<uint>();

        public static HashSet<uint> activeAudio = new HashSet<uint>();

        // Special Participants
        public static Participant me = null;

        /// <summary>
        /// Did we join a meeting in progress, or start the meeting?
        /// </summary>
        public static bool ZoomAlreadyRunning = false;

        public static void Init(IHostApp app)
        {
            hostApp = app;
            cfg = DeserializeJson<ControllerConfigurationSettings>(app.GetSettingsAsJSON());
        }

        public static void Start()
        {
            StartingMeeting = true;
            ShouldExit = false;

            var StartCompleteEvent = new EventWaitHandle(false, EventResetMode.AutoReset);

            hostApp.Log(LogType.DBG, $"Controller main thread={string.Format("0x{0:X8}", Thread.CurrentThread.ManagedThreadId)}");

            Task.Run(() =>
            {
                hostApp.Log(LogType.DBG, $"Controller task thread={string.Format("0x{0:X8}", Thread.CurrentThread.ManagedThreadId)}");

                try
                {
                    BeginLoginSequence();

                    // TBD: There has to be a better way to do this ...
                    hostApp.Log(LogType.DBG, $"Waiting for meeting to start");
                    while ((StartingMeeting) && (!ShouldExit))
                    {
                        Application.DoEvents();
                        Thread.Sleep(250);
                    }

                    hostApp.Log(LogType.DBG, $"Determining if meeting is already in progress or not");
                    lock (participants)
                    {
                        // Some control bots need to know if we started the meeting or joined a meeting already in progress.
                        //   TBD: I don't immediately see a way to access this info via the API, but if we're the only participant
                        //   attending the meeting (vs in the waiting room), then it's probably safe to assume we started it.
                        int numAttending = 0;
                        foreach (var p in participants.Values)
                        {
                            if (p.Status == ParticipantStatus.Attending)
                            {
                                numAttending++;
                            }
                        }

                        ZoomAlreadyRunning = numAttending > 1;
                    }

                    hostApp.Log(LogType.DBG, $"Signalling meeting started");
                    StartCompleteEvent.Set();
                }
                catch (Exception ex)
                {
                    hostApp.Log(LogType.ERR, $"[Controller] Failed to start: {ex}");
                    StartingMeeting = false;
                    OnExit(null, null);
                    return;
                }

                try
                {
                    hostApp.Log(LogType.DBG, $"Entering action loop");
                    var nextActionDT = DateTime.MinValue;
                    while (!ShouldExit)
                    {
                        try
                        {
                            Application.DoEvents();
                            Thread.Sleep(100);
                        }
                        catch (Exception ex)
                        {
                            hostApp.Log(LogType.ERR, $"[Controller] Exception in inner action loop message pump: {ex}");
                        }

                        try
                        {
                            // This has to fire from the same thread for the SDK to work
                            var nowDT = DateTime.UtcNow;
                            if (nowDT >= nextActionDT)
                            {
                                var rateInMS = cfg.ActionTimerRateInMS;
                                nextActionDT = nowDT.AddMilliseconds(rateInMS);

                                OnActionTimerTick(null, null);

                                var execTimeInMS = DateTime.UtcNow.Subtract(nowDT).TotalMilliseconds;
                                if (execTimeInMS > rateInMS)
                                {
                                    hostApp.Log(LogType.WRN, $"OnActionTimerTick lagging {(int)(execTimeInMS - rateInMS)}ms");
                                }

                                UpdateParticipantActiveAudio();
                            }
                        }
                        catch (Exception ex)
                        {
                            hostApp.Log(LogType.ERR, $"[Controller] Exception in inner action loop: {ex}");
                        }
                    }
                }
                catch (Exception ex)
                {
                    hostApp.Log(LogType.ERR, $"[Controller] Exception in outer action loop: {ex}");
                    OnExit(null, null);
                }
            });

            hostApp.Log(LogType.DBG, $"Waiting for meeting started signal");
            StartCompleteEvent.WaitOne();

            hostApp.Log(LogType.INF, "Controller start successful");
        }

        /// <summary>
        /// Processes application events for the given number of milliseconds using the given polling interval.  This is like the native
        /// Thread.Sleep() but it allows background tasks to execute during the sleep time.
        /// </summary>
        public static void SleepWithDoEvents(int millisecondsTimeout, int millisecondsPollingInterval = 100)
        {
            if (millisecondsTimeout < 0)
            {
                throw new ArgumentOutOfRangeException("millisecondsTimeout must be >= 0");
            }

            if (millisecondsPollingInterval < 0)
            {
                throw new ArgumentOutOfRangeException("millisecondsPollingInterval must be >= 0");
            }

            var nowDT = DateTime.UtcNow;
            var timeoutDT = nowDT.AddMilliseconds(millisecondsTimeout);
            while (nowDT <= timeoutDT)
            {
                Application.DoEvents();
                if (millisecondsPollingInterval > 0)
                {
                    Thread.Sleep(millisecondsPollingInterval);
                }

                nowDT = DateTime.UtcNow;
            }
        }

        public static void Stop()
        {
            ShouldExit = true;
        }

        public static List<Participant> GetParticipantsByName(string name)
        {
            List<Participant> ret = new List<Participant>();

            lock (participants)
            {
                foreach (var p in participants.Values)
                {
                    if (p.Name == name)
                    {
                        ret.Add(p);
                    }
                }
            }

            return ret;
        }

        public static Participant GetParticipantByName(string name)
        {
            List<Participant> list = GetParticipantsByName(name);

            if (list.Count == 0)
            {
                _ = SpecialParticipant.TryGet(name, out Participant p);
                return p;
            }

            if (list.Count == 1)
            {
                return list[0];
            }

            throw new ArgumentException($"User name {name} is not unique");
        }

        [MethodImplAttribute(MethodImplOptions.NoInlining)]
        public static bool GetParticipantById(uint userId, out Participant p, bool logWarningIfNotFound = true)
        {
            if (SpecialParticipant.TryGetValue(userId, out p))
            {
                return true;
            }

            lock (participants)
            {
                if (participants.TryGetValue(userId, out p))
                {
                    return true;
                }
            }

            if (logWarningIfNotFound)
            {
                hostApp.Log(LogType.WRN, $"{new StackFrame(1).GetMethod().Name} No participant object for {userId}");
            }

            p = null;
            return false;
        }

        [MethodImplAttribute(MethodImplOptions.NoInlining)]
        private static Participant GetOrCreateParticipant(uint userId, string name, bool logWarningIfNotFound = true)
        {
            if (!GetParticipantById(userId, out Participant p, false))
            {
                p = new Participant()
                {
                    UserId = userId,
                    Name = name,
                };

                if (logWarningIfNotFound)
                {
                    hostApp.Log(LogType.WRN, $"{new StackFrame(1).GetMethod().Name} Proceeding with non-existent participant {p}");
                }
            }

            return p;
        }

        [MethodImplAttribute(MethodImplOptions.NoInlining)]
        public static bool SendChatMessage(Participant to, string text)
        {
            try
            {
                var sdkErr = chatController.SendChatTo(to.UserId, text);
                if (sdkErr != SDKError.SDKERR_SUCCESS)
                {
                    throw new Exception(sdkErr.ToString());
                }
                return true;
            }
            catch (Exception ex)
            {
                hostApp.Log(LogType.ERR, $"{new StackFrame(1).GetMethod().Name} Failed to send chat message {repr(text)} to {to}: {ex}");
            }
            return false;
        }

        [MethodImplAttribute(MethodImplOptions.NoInlining)]
        public static bool ReclaimHost(string hostKey = null)
        {
            try
            {
                var sdkErr = (hostKey == null) ? participantController.ReclaimHost() : participantController.ReclaimHostByHostKey(hostKey);
                if (sdkErr != SDKError.SDKERR_SUCCESS)
                {
                    throw new Exception(sdkErr.ToString());
                }
                return true;
            }
            catch (Exception ex)
            {
                hostApp.Log(LogType.ERR, $"{new StackFrame(1).GetMethod().Name} Failed to reclaim host: {ex}");
            }
            return false;
        }

        [MethodImplAttribute(MethodImplOptions.NoInlining)]
        public static bool RenameParticipant(Participant p, string newName)
        {
            try
            {
                var sdkErr = participantController.ChangeUserName(p.UserId, newName, false);
                if (sdkErr != SDKError.SDKERR_SUCCESS)
                {
                    throw new Exception(sdkErr.ToString());
                }
                return true;
            }
            catch (Exception ex)
            {
                hostApp.Log(LogType.ERR, $"{new StackFrame(1).GetMethod().Name} Failed to rename {p} to {repr(newName)}: {ex}");
            }
            return false;
        }

        [MethodImplAttribute(MethodImplOptions.NoInlining)]
        public static bool UnmuteParticipant(Participant p)
        {
            try
            {
                var sdkErr = audioController.UnMuteAudio(p.UserId);
                if (sdkErr != SDKError.SDKERR_SUCCESS)
                {
                    throw new Exception(sdkErr.ToString());
                }
                return true;
            }
            catch (Exception ex)
            {
                hostApp.Log(LogType.ERR, $"{new StackFrame(1).GetMethod().Name} Failed to unmute {p}: {ex}");
            }
            return false;
        }

        [MethodImplAttribute(MethodImplOptions.NoInlining)]
        public static bool MuteParticipant(Participant p, bool allowUnmute = true)
        {
            try
            {
                var sdkErr = audioController.MuteAudio(p.UserId, allowUnmute);
                if (sdkErr != SDKError.SDKERR_SUCCESS)
                {
                    throw new Exception(sdkErr.ToString());
                }
                return true;
            }
            catch (Exception ex)
            {
                hostApp.Log(LogType.ERR, $"{new StackFrame(1).GetMethod().Name} Failed to mute {p}: {ex}");
            }
            return false;
        }

        [MethodImplAttribute(MethodImplOptions.NoInlining)]
        public static bool AdmitParticipant(Participant p)
        {
            try
            {
                var sdkErr = waitController.AdmitToMeeting(p.UserId);
                if (sdkErr != SDKError.SDKERR_SUCCESS)
                {
                    throw new Exception(sdkErr.ToString());
                }

                lock (participants)
                {
                    // Admitting a participant takes several seconds; Mark them as Joining so that we don't try to admit them again
                    hostApp.Log(LogType.INF, $"Participant {p} is joining the meeting");
                    p.Status = ParticipantStatus.Joining;
                }

                return true;
            }
            catch (Exception ex)
            {
                hostApp.Log(LogType.ERR, $"{new StackFrame(1).GetMethod().Name} Failed to admit {p}: {ex}");
            }
            return false;
        }

        [MethodImplAttribute(MethodImplOptions.NoInlining)]
        public static bool PromoteParticipant(Participant p, ParticipantRole newRole)
        {
            try
            {
                SDKError sdkErr;
                if (newRole == ParticipantRole.CoHost)
                {
                    sdkErr = participantController.AssignCoHost(p.UserId);
                }
                else if (newRole == ParticipantRole.Host)
                {
                    sdkErr = participantController.MakeHost(p.UserId);
                }
                else
                {
                    throw new InvalidOperationException($"Requested role not valid");
                }

                if (sdkErr != SDKError.SDKERR_SUCCESS)
                {
                    throw new Exception(sdkErr.ToString());
                }
                return true;
            }
            catch (Exception ex)
            {
                hostApp.Log(LogType.ERR, $"{new StackFrame(1).GetMethod().Name} Failed to promote {p} to {newRole}: {ex}");
            }
            return false;
        }

        [MethodImplAttribute(MethodImplOptions.NoInlining)]
        public static bool DemoteParticipant(Participant p)
        {
            try
            {
                var sdkErr = participantController.RevokeCoHost(p.UserId);
                if (sdkErr != SDKError.SDKERR_SUCCESS)
                {
                    throw new Exception(sdkErr.ToString());
                }
                return true;
            }
            catch (Exception ex)
            {
                hostApp.Log(LogType.ERR, $"{new StackFrame(1).GetMethod().Name} Failed to demote {p}: {ex}");
            }
            return false;
        }

        [MethodImplAttribute(MethodImplOptions.NoInlining)]
        public static bool ExpelParticipant(Participant p)
        {
            try
            {
                var sdkErr = participantController.ExpelUser(p.UserId);
                if (sdkErr != SDKError.SDKERR_SUCCESS)
                {
                    throw new Exception(sdkErr.ToString());
                }
                return true;
            }
            catch (Exception ex)
            {
                hostApp.Log(LogType.ERR, $"{new StackFrame(1).GetMethod().Name} Failed to expel {p}: {ex}");
            }
            return false;
        }

        [MethodImplAttribute(MethodImplOptions.NoInlining)]
        public static bool PutParticipantInWaitingRoom(Participant p)
        {
            try
            {
                var sdkErr = waitController.PutInWaitingRoom(p.UserId);
                if (sdkErr != SDKError.SDKERR_SUCCESS)
                {
                    throw new Exception(sdkErr.ToString());
                }
                return true;
            }
            catch (Exception ex)
            {
                hostApp.Log(LogType.ERR, $"{new StackFrame(1).GetMethod().Name} Failed to put {p} in waiting room: {ex}");
            }
            return false;
        }

        [MethodImplAttribute(MethodImplOptions.NoInlining)]
        public static bool SetMeetingTopic(string newTopic)
        {
            try
            {
                var sdkErr = uiController.SetMeetingTopic(newTopic);
                if (sdkErr != SDKError.SDKERR_SUCCESS)
                {
                    throw new Exception(sdkErr.ToString());
                }
                return true;
            }
            catch (Exception ex)
            {
                hostApp.Log(LogType.ERR, $"{new StackFrame(1).GetMethod().Name} Failed to set meeting topic to {repr(newTopic)}: {ex}");
            }
            return false;
        }

        [MethodImplAttribute(MethodImplOptions.NoInlining)]
        public static bool LeaveMeeting(bool endForAll = false)
        {
            try
            {
                if (mtgService != null)
                {
                    var sdkErr = mtgService.Leave(endForAll ? LeaveMeetingCmd.END_MEETING : LeaveMeetingCmd.LEAVE_MEETING);
                    if (sdkErr != SDKError.SDKERR_SUCCESS)
                    {
                        throw new Exception(sdkErr.ToString());
                    }
                }

                return true;
            }
            catch (Exception ex)
            {
                hostApp.Log(LogType.ERR, $"{new StackFrame(1).GetMethod().Name} Failed to leave meeting: {ex}");
            }

            return false;
        }

        [MethodImplAttribute(MethodImplOptions.NoInlining)]
        public static bool LockMeeting(bool enableLock)
        {
            try
            {
                if (mtgService != null)
                {
                    var sdkErr = enableLock ? mtgService.LockMeeting() : mtgService.UnlockMeeting();
                    if (sdkErr != SDKError.SDKERR_SUCCESS)
                    {
                        throw new Exception(sdkErr.ToString());
                    }
                }

                return true;
            }
            catch (Exception ex)
            {
                hostApp.Log(LogType.ERR, $"{new StackFrame(1).GetMethod().Name} Failed to {(enableLock ? "lock" : "unlock")} meeting: {ex}");
            }

            return false;
        }

        [MethodImplAttribute(MethodImplOptions.NoInlining)]
        public static bool SetAllowParticipantsToUnmuteSelf(bool allow)
        {
            try
            {
                if (mtgService != null)
                {
                    var sdkErr = participantController.AllowParticipantsToUnmuteSelf(allow);
                    if (sdkErr != SDKError.SDKERR_SUCCESS)
                    {
                        throw new Exception(sdkErr.ToString());
                    }
                }

                return true;
            }
            catch (Exception ex)
            {
                hostApp.Log(LogType.ERR, $"{new StackFrame(1).GetMethod().Name} Failed to set allowParticipantsToUnmuteSelf={allow}: {ex}");
            }

            return false;
        }

        private static bool loggedLayoutWindowsWarning = false;
        public static void LayoutWindows()
        {
            if (!loggedLayoutWindowsWarning)
            {
                hostApp.Log(LogType.WRN, $"{MethodBase.GetCurrentMethod().Name} Not Yet Implemented");
                loggedLayoutWindowsWarning = true;
            }
        }
    }
}