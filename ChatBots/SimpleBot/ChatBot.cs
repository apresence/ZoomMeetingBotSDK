namespace ZoomMeetingBotSDK.ChatBot.SimpleBot
{
    using System;
    using System.Collections.Generic;
    using static Utils;

    public class ChatBot : IChatBot
    {
        private static IHostApp hostApp = null;
        private static readonly object SettingsLock = new object();

        private static ChatBotInfo chatBotInfo = new ChatBotInfo()
        {
             IntelligenceLevel = 10,
             Name = "SimpleBot",
        };

        private static Dictionary<string, string> smallTalk = new Dictionary<string, string>();
        private static List<string> randomTalk = new List<string>();

        private static string SmallTalk(string text)
        {
            string response = null;

            if (smallTalk != null)
            {
                foreach (var word in text.GetWordsInSentence())
                {
                    if (smallTalk.TryGetValue(word.ToLower(), out response))
                    {
                        break;
                    }
                }
            }

            return response;
        }

        private static string RandomTalk(string text)
        {
            string response = null;

            if (randomTalk != null)
            {
                response = randomTalk.RandomElement();
            }            

            return response;
        }

        public string Converse(string input, string from)
        {
            string response = null;

            lock (SettingsLock)
            {
                response = SmallTalk(input);
                if (response == null)
                {
                    response = RandomTalk(input);
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
                smallTalk = DynToStrDic(hostApp.GetSetting("SmallTalkSequences"));
                ExpandDictionaryPipes(smallTalk);

                randomTalk = DynToStrList(hostApp.GetSetting("RandomTalk"));
            }
        }
    }
}