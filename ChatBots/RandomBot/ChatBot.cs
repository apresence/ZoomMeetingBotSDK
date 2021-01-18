/// <summary>
/// This is a minimal example of a ChatBot.  It mindlessly returns a random response from "RandomTalk" in the config file.
/// </summary>

namespace ZoomMeetingBotSDK.ChatBot.RandomBot
{
    using System;
    using System.Collections.Generic;
    using static Utils;

    public class ChatBot : IChatBot
    {
        internal class RandomBotConfigurationSettings
        {
            internal RandomBotConfigurationSettings()
            {
                RandomTalk = null;
            }

            /// <summary>
            /// Random bot responses.  TBD: Move into RemedialBot
            /// </summary>
            public string[] RandomTalk { get; set; }
        }

        private static IHostApp hostApp = null;
        private static readonly object SettingsLock = new object();

        private static ChatBotInfo chatBotInfo = new ChatBotInfo()
        {
             DefaultOrder = 1000,
             Name = "RandomBot",
        };

        private static RandomBotConfigurationSettings randomBotConfigurationSettings = null;

        private static string RandomTalk(string text)
        {
            string response = null;

            var randomTalk = randomBotConfigurationSettings.RandomTalk;
            if (randomTalk != null)
            {
                response = randomTalk.RandomElement();
            }

            return response;
        }

        public string Converse(string input, IChatBotUser from)
        {
            string response = null;

            lock (SettingsLock)
            {
                response = RandomTalk(input);
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
                hostApp.Log(LogType.WRN, $"RandomBot.Start(): Already initialized");
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
                randomBotConfigurationSettings = DeserializeJson<RandomBotConfigurationSettings>(hostApp.GetSettingsAsJSON());
            }
        }
    }
}