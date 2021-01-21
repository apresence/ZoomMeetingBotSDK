/// <summary>
/// This is a simple ChatBot that picks up on keywords and, if a response is defined for a given keyword, returns the configured response.  It has a few
/// other tricks too such as greeting someone the first time they converse with it.
/// </summary>

namespace ZoomMeetingBotSDK.ChatBot.SimpleBot
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using static Utils;

    public class ChatBot : IChatBot
    {
        internal class SimpleBotConfigurationSettings
        {
            internal SimpleBotConfigurationSettings()
            {
                OneTimeHiSequences = new Dictionary<string, string>();
                OneTimeMsg = "I'm just a Bot. If you need something, please chat with one of the Co-Hosts.";
                SmallTalkSequences = new Dictionary<string, string[]>();
            }

            /// <summary>
            /// A list of pre-defined one-time greetings based on received private chat messages.  Use pipes to allow more than one query to produce the same response.  TBD: Move into RemedialBot
            /// </summary>
            public Dictionary<string, string> OneTimeHiSequences { get; set; }

            /// <summary>
            /// One-time message send after one-time hi (and optionally topic).
            /// </summary>
            public string OneTimeMsg { get; set; }

            /// <summary>
            /// A list of pre-defined small talk responses based on received private chat messages.  Use pipes to allow more than one query to produce the same response.  TBD: Move into RemedialBot
            /// </summary>
            public Dictionary<string, string[]> SmallTalkSequences { get; set; }
        }

        private static IHostApp hostApp = null;
        private static readonly object SettingsLock = new object();

        private static readonly Dictionary<uint, HashSet<string>> ThingsISaidToUserId = new Dictionary<uint, HashSet<string>>();
        private static readonly Dictionary<string, HashSet<string>> ThingsISaidToName = new Dictionary<string, HashSet<string>>();

        private static ChatBotInfo chatBotInfo = new ChatBotInfo()
        {
             DefaultOrder = 10,
             Name = "SimpleBot",
        };

        private static SimpleBotConfigurationSettings cfg = null;

        private static bool DidISay(string text, IChatBotUser to)
        {
            HashSet<string> thingsSaid;

            if ((to.UserId != 0) && ThingsISaidToUserId.TryGetValue(to.UserId, out thingsSaid))
            {
                if (thingsSaid.Contains(text))
                {
                    return true;
                }
            }

            if (ThingsISaidToName.TryGetValue(to.Name, out thingsSaid))
            {
                if (thingsSaid.Contains(text))
                {
                    return true;
                }
            }

            return false;
        }

        private static void RememberThatISaid(string text, IChatBotUser to)
        {
            if (text == null)
            {
                return;
            }

            if (text.StartsWith("{") || text.StartsWith("/"))
            {
                // Don't remember these...
                return;
            }

            HashSet<string> thingsISaid = null;

            if ((to.UserId != 0) && (!ThingsISaidToUserId.TryGetValue(to.UserId, out thingsISaid)))
            {
                thingsISaid = new HashSet<string>();
                ThingsISaidToUserId[to.UserId] = thingsISaid;
            }

            thingsISaid.Add(text);

            if (!ThingsISaidToName.TryGetValue(to.Name, out thingsISaid))
            {
                thingsISaid = new HashSet<string>();
                ThingsISaidToName[to.Name] = thingsISaid;
            }

            thingsISaid.Add(text);
        }

        private static string ChooseRandomResponse(string[] responses, IChatBotUser to)
        {
            // TBD: Track things we've said so we don't say them more than once.
            if (responses == null)
            {
                return null;
            }

            // Don't repeat things I've already said
            HashSet<string> thingsISaid = null;
            HashSet<string> allThingsISaid = new HashSet<string>();

            if ((to.UserId != 0) && ThingsISaidToUserId.TryGetValue(to.UserId, out thingsISaid))
            {
                allThingsISaid.UnionWith(thingsISaid);
            }

            if (ThingsISaidToName.TryGetValue(to.Name, out thingsISaid))
            {
                allThingsISaid.UnionWith(thingsISaid);
            }

            HashSet<string> responseHS = new HashSet<string>(responses);
            responseHS.ExceptWith(allThingsISaid);

            if (responseHS.Count == 0)
            {
                return null;
            }

            // Choose something random from somehing I haven't already said
            return responseHS.RandomElement();
        }

        private static string OneTimeHi(string text, IChatBotUser to)
        {
            string response = null;
            string key = "_onetimehi_";

            if (cfg.OneTimeHiSequences == null)
            {
                return null;
            }

            // Try to give a specific response
            foreach (var word in text.GetWordsInSentence())
            {
                if (cfg.OneTimeHiSequences.TryGetValue(word.ToLower(), out response))
                {
                    break;
                }
            }

            if (DidISay(key, to))
            {
                // I already said it
                return null;
            }

            // No specific response? Pick something at random
            if (response == null)
            {
                response = ChooseRandomResponse(cfg.OneTimeHiSequences.Values.ToArray<string>(), to);
            }

            RememberThatISaid(key, to);

            return response;
        }

        private static string ExpandOneTimeMsg(string text, IChatBotUser to)
        {
            string key = "_onetimemsg_";
            string msg = null;

            if (!DidISay(key, to) && (cfg.OneTimeMsg != null))
            {
                if (text == null)
                {
                    text = cfg.OneTimeMsg;
                }
                else
                {
                    text += " " + cfg.OneTimeMsg;
                }

                RememberThatISaid(key, to);
            }

            return text;
        }

        private static string SmallTalk(string text, IChatBotUser to)
        {
            string[] responses = null;
            bool isToEveryone = to.Name.ToLower().StartsWith("everyone") || to.UserId == 0;

            var smallTalk = cfg.SmallTalkSequences;
            if (smallTalk != null)
            {
                foreach (var word in text.GetWordsInSentence())
                {
                    if (smallTalk.TryGetValue(word.ToLower(), out responses))
                    {
                        break;
                    }
                }
            }

            var response = ChooseRandomResponse(responses, to);
            RememberThatISaid(response, to);

            // Add one-time message if needed
            if ((!isToEveryone) && ((response == null) || ((response != null) && (!response.StartsWith("{")))))
            {
                response = ExpandOneTimeMsg(response, to);
            }

            return response;
        }

        public string Converse(string text, IChatBotUser from)
        {
            string response = null;

            lock (SettingsLock)
            {
                var justSayHi = text == "_onetimehi_";

                if (justSayHi)
                {
                    // Use time-appropriate greeting
                    text = "morning";
                }

                // Say hi if we haven't already
                var hi = OneTimeHi(text, from);
                if (justSayHi)
                {
                    return hi;
                }

                var smalltalk = SmallTalk(text, from);

                if ((hi != null) && (smalltalk != null))
                {
                    response = hi + " " + smalltalk;
                }
                else if (smalltalk != null)
                {
                    response = smalltalk;
                }
                else
                {
                    response = hi;
                }
            }

            return response;
        }

        public ChatBotInfo GetChatBotInfo()
        {
            return chatBotInfo;
        }

        public void Init(ChatBotInitParam param)
        {
            if (hostApp != null)
            {
                hostApp.Log(LogType.WRN, $"SimpleBot.Start(): Already initialized");
                return;
            }

            hostApp = param.hostApp;
            hostApp.SettingsChanged += new EventHandler(SettingsChanged);

            LoadSettings();
        }

        public void Start()
        {
            // Nothing for this bot to do
        }

        public void Stop()
        {
            // Nothing for this bot to do
        }

        public void SettingsChanged(object sender, EventArgs e)
        {
            LoadSettings();
        }

        private void LoadSettings()
        {
            lock (SettingsLock)
            {
                cfg = DeserializeJson<SimpleBotConfigurationSettings>(hostApp.GetSettingsAsJSON());
                ExpandDictionaryPipes(cfg.OneTimeHiSequences);
                ExpandDictionaryPipes(cfg.SmallTalkSequences);
            }
        }
    }
}