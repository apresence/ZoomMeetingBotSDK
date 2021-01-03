using global::ZoomController.Interop.Bot;

namespace ZoomController.Bot
{
    using System;
    using System.Collections.Generic;
    using System.Diagnostics;
    using System.Net.Http;
    using System.Text.RegularExpressions;

    namespace ChatterBot
    {
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

            private Process p = null;
            private bool bStarted = false;

            /// <summary>Starts up ChatBot and prepares it to converse.</summary>
            void IChatBot.Start(ChatBotInitParam param)
            {
                if (bStarted)
                {
                    Global.Log(Global.LogType.WRN, "chatBot already started");
                    return;
                }

                Global.Log(Global.LogType.DBG, "chatBot starting");

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
                            Arguments = "/c chatbot.cmd",
                        },
                    };
                    p.Start();

                    string line;
                    while ((line = p.StandardOutput.ReadLine()) != null)
                    {
                        if (line == "Chatbot started")
                        {
                            Global.Log(Global.LogType.INF, "chatBot started");
                            bStarted = true;
                            return;
                        }
                    }
                }
                catch (Exception ex)
                {
                    Global.Log(Global.LogType.ERR, "chatBot failed to start: {0}", Global.repr(ex.ToString()));
                }
            }

            void IChatBot.Stop()
            {
                if ((!bStarted) && (p == null))
                {
                    Global.Log(Global.LogType.WRN, "chatBot is not started");
                    return;
                }

                try
                {
                    if (bStarted)
                    {
                        Global.Log(Global.LogType.DBG, "chatBot stopping");
                        p.StandardInput.WriteLine("quit");
                        if (p.WaitForExit(2000))
                        {
                            p = null;
                        }
                    }
                }
                catch (Exception ex)
                {
                    Global.Log(Global.LogType.DBG, "chatBot exception {0}", ex.ToString());
                }
                finally
                {
                    if (p != null)
                    {
                        try
                        {
                            Global.Log(Global.LogType.DBG, "chatBot killing");
                            p.Kill();
                        }
                        catch (Exception ex)
                        {
                            Global.Log(Global.LogType.WRN, "chatBot exception during kill: {0}", ex.ToString());
                        }
                        p = null;
                    }
                    bStarted = false;
                }
                Global.Log(Global.LogType.INF, "chatBot stopped");
            }

            /// <summary>
            /// Takes input and returns conversational output from the chatbot.
            /// It is assumed that the chatbot will prompt for input with "] ".
            /// </summary>
            string IChatBot.Converse(string input, string from)
            {
                if (!bStarted)
                {
                    return null;
                }

                Global.Log(Global.LogType.DBG, "chatBot > {0}", Global.repr(input));
                p.StandardInput.WriteLine(input);

                // TBD: Implement some sort of timeout or something so we don't hang forever if things go wonky
                string line = p.StandardOutput.ReadLine();
                Global.Log(Global.LogType.DBG, "chatBot < {0}", Global.repr(line));

                return line;
            }

            ChatBotInfo IChatBot.GetChatBotInfo()
            {
                return chatBotInfo;
            }

            ~ChatBot()
            {
                ((IChatBot)this).Stop();
            }
        }
    }
}
