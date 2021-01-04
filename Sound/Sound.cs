using System;
using System.Collections.Generic;
using System.Linq;
using System.Speech.Synthesis;
using System.Text;
using System.Threading.Tasks;
using System.Collections.Concurrent;
using System.Security.AccessControl;
using System.IO;
using System.Media;
using ZoomMeetingBotSDK.Interop.HostApp;
using global::ZoomMeetingBotSDK.Utils;

namespace ZoomMeetingBotSDK
{
    public static class Sound
    {
        private static IHostApp hostApp = null;

        private static SpeechSynthesizer tts;
        private static ConcurrentQueue<SoundItem> q = new ConcurrentQueue<SoundItem>();
        private static System.Threading.Timer timer;

        private class SoundItem
        {
            public string Action;
            public string Parameters;

            public SoundItem(string action, string parameters)
            {
                Action = action;
                Parameters = parameters;
            }
        }

        public static void Init(IHostApp app)
        {
            if (hostApp != null)
            {
                hostApp.Log(LogType.WRN, "ZoomMeetingBotSDK.Sound: Already Initialized");
                return;
            }

            hostApp.Log(LogType.DBG, "Initializing TTS");
            tts = new SpeechSynthesizer();
            var voice = hostApp.GetSetting("tts.voice");
            if ((voice != null) && (voice.Length > 0))
            {
                try
                {
                    tts.SelectVoice(voice);
                }
                catch (Exception e)
                {
                    hostApp.Log(LogType.ERR, "TTS failed to load voice {0}; Falling back on default; Err={1}", voice, ZMBUtils.repr(e));
                }
            }
            hostApp.Log(LogType.INF, "TTS Loaded voice {0}", ZMBUtils.repr(tts.Voice.Name));
            tts.SetOutputToDefaultAudioDevice();
            timer = new System.Threading.Timer(TimerHandler, null, 0, 250);
        }

        private static void PlaySoundFile(string soundFilePath)
        {
            soundFilePath = "sounds\\" + soundFilePath;
            if (Path.GetExtension(soundFilePath) == String.Empty) {
                if (Directory.Exists(soundFilePath))
                {
                    // If it's a directory, pick one of the .wav files in it at random
                    soundFilePath = Directory.EnumerateFiles(soundFilePath, "*.wav").RandomElement();
                }
                else
                {
                    soundFilePath += ".wav";
                }
            }
            hostApp.Log(LogType.INF, "Playing random sound file: {0}", ZMBUtils.repr(soundFilePath));
            new SoundPlayer(soundFilePath).PlaySync();
        }

        private static void TimerHandler(object state)
        {            
            while (q.TryDequeue(out SoundItem si))
            {
                try
                {
                    if (si.Action == "speak")
                    {
                        hostApp.Log(LogType.INF, "TTS Speaking: {0}", ZMBUtils.repr(si.Parameters));
                        tts.Speak(si.Parameters);
                    }
                    else if (si.Action == "play")
                    {
                        PlaySoundFile(si.Parameters);
                    }
                    else
                    {
                        throw new Exception("Unsupported sound action: " + si.Action);
                    }
                }
                catch (Exception ex)
                {
                    hostApp.Log(LogType.ERR, "Exception while processing sound item [{0}, {1}]: {2}", ZMBUtils.repr(si.Action), ZMBUtils.repr(si.Parameters), ZMBUtils.repr(ex.ToString()));
                }
            }
        }

        /// <summary>
        /// Splits out camel case words into individual words.  For example: "This is a CamelCase word" => "This is a Camel Case word".
        /// Based on: https://stackoverflow.com/a/7599674
        /// </summary>
        static private string SplitCamelCase(string input)
        {
            return string.Join(" ", System.Text.RegularExpressions.Regex.Split(input, @"(?<=[a-z])(?=[A-Z])|(?<=[A-Z])(?=[A-Z][a-z])"));
        }

        /// <summary>
        /// Pre-processes text so the TTS engine speaks it properly
        /// </summary>
        static private string TTSPreprocess(string input)
        {
            // Breaks up CamelCase words longer than 4 chars into multiple words, for example: "CamelCase" => "Camel Case".
            var ret = SplitCamelCase(input);

            // Makes TTS say A.I. instead of like the Japanese word for love (ai)
            ret = FastRegex.Replace(ret, @"\bAI\b", "<say-as interpret-as='characters'>AI</say-as>", System.Text.RegularExpressions.RegexOptions.Compiled);

            return ret;
        }

        static public void Speak(string text)
        {
            q.Enqueue(new SoundItem("speak", TTSPreprocess(text)));
        }

        static public void Play(string soundFilePath)
        {
            // TBD: Could pre-load the sound before playing in the queue
            q.Enqueue(new SoundItem("play", soundFilePath));
        }
    }
}
