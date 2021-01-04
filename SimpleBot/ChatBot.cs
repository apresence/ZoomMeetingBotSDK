namespace ZoomMeetingBotSDK.ChatBot.SimpleBot
{
    using System.Collections.Generic;
    using ZoomMeetingBotSDK.Interop.HostApp;
    using ZoomMeetingBotSDK.Interop.ChatBot;
    using ZoomMeetingBotSDK.Utils;

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

        public void Start(ChatBotInitParam param)
        {
            if (hostApp != null)
            {
                hostApp.Log(LogType.WRN, $"SimpleBot.Start(): Already started");
                return;
            }

            hostApp = param.hostApp;

            SettingsUpdated();
        }

        void IChatBot.Stop()
        {
            // todo
        }

        public void SettingsUpdated()
        {
            lock (SettingsLock)
            {
                smallTalk = hostApp.GetSetting("SmallTalkSequences");
                randomTalk = hostApp.GetSetting("RandomTalk");
            }
        }
    }
}
    