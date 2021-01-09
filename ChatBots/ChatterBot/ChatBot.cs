namespace ZoomMeetingBotSDK.ChatBot.ChatterBot
{
    using System;
    using System.Diagnostics;
    using ZoomMeetingBotSDK;
    using static Utils;

    /// <summary>
    /// This is an extremly simple/naive wrapper around chatbot.py which is a wrapper for chatterbot in Python.  Feel free to re-write it.
    /// https://chatterbot.readthedocs.io/en/stable/
    /// </summary>
    public class ChatBot : IChatBot
    {
        private static readonly ChatBotInfo chatBotInfo = new ChatBotInfo()
        {
            Name = "ChatterBot",
            IntelligenceLevel = 100,
        };

        private IHostApp hostApp;
        private Process p = null;
        private bool bStarted = false;

        
        public void Init(ChatBotInitParam param)
        {
            if (hostApp != null)
            {
                hostApp.Log(LogType.WRN, "chatBot already initialized");
            }

            hostApp = param.hostApp;
        }

        public void Start()
        {
            if (bStarted)
            {
                hostApp.Log(LogType.WRN, "chatBot already started");
                return;
            }

            hostApp.Log(LogType.DBG, "chatBot starting");

            try
            {
                p = new Process()
                {
                    StartInfo = new ProcessStartInfo
                    {
                        UseShellExecute = false,
                        RedirectStandardInput = true,
                        RedirectStandardOutput = true,
                        FileName = @"C:\Windows\system32\cmd.exe", // TBD: Load from env
                        Arguments = "/c ChatterBot.cmd",
                    },
                };
                p.Start();

                string line;
                while ((line = p.StandardOutput.ReadLine()) != null)
                {
                    if (line == "Chatbot started")
                    {
                        hostApp.Log(LogType.INF, "chatBot started");
                        bStarted = true;
                        return;
                    }
                }
            }
            catch (Exception ex)
            {
                hostApp.Log(LogType.ERR, "chatBot failed to start: {0}", repr(ex.ToString()));
            }
        }

        public void Stop()
        {
            if ((!bStarted) && (p == null))
            {
                hostApp.Log(LogType.WRN, "chatBot is not started");
                return;
            }

            try
            {
                if (bStarted)
                {
                    hostApp.Log(LogType.DBG, "chatBot stopping");
                    p.StandardInput.WriteLine("quit");
                    if (p.WaitForExit(2000))
                    {
                        p = null;
                    }
                }
            }
            catch (Exception ex)
            {
                hostApp.Log(LogType.DBG, "chatBot exception {0}", ex.ToString());
            }
            finally
            {
                if (p != null)
                {
                    try
                    {
                        hostApp.Log(LogType.DBG, "chatBot killing");
                        p.Kill();
                    }
                    catch (Exception ex)
                    {
                        hostApp.Log(LogType.WRN, "chatBot exception during kill: {0}", ex.ToString());
                    }
                    p = null;
                }
                bStarted = false;
            }
            hostApp.Log(LogType.INF, "chatBot stopped");
        }

        /// <summary>
        /// Takes input and returns conversational output from the chatbot.
        /// It is assumed that the chatbot will prompt for input with "] ".
        /// </summary>
        public string Converse(string input, string from)
        {
            if (!bStarted)
            {
                return null;
            }

            hostApp.Log(LogType.DBG, "chatBot > {0}", repr(input));
            p.StandardInput.WriteLine(input);

            // TBD: Implement some sort of timeout or something so we don't hang forever if things go wonky
            string line = p.StandardOutput.ReadLine();
            hostApp.Log(LogType.DBG, "chatBot < {0}", repr(line));

            return line;
        }

        public ChatBotInfo GetChatBotInfo()
        {
            return chatBotInfo;
        }

        public void SettingsUpdated(object sender, EventArgs e)
        {
            // We don't use any config settings, so there is nothing to do
        }

        ~ChatBot()
        {
            ((IChatBot)this).Stop();
        }
    }
}
