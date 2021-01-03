﻿using System;
using System.Diagnostics;
using ZoomController.Utils;
using ZoomController.Interop.HostApp;
using ZoomController.Interop.ChatBot;

namespace ZoomController.ChatBot
{
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

            private IHostApp hostApp;
            private Process p = null;
            private bool bStarted = false;

            /// <summary>Starts up ChatBot and prepares it to converse.</summary>
            void IChatBot.Start(ChatBotInitParam param)
            {
                hostApp = param.hostApp;

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
                            Arguments = "/c chatbot.cmd",
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
                    hostApp.Log(LogType.ERR, "chatBot failed to start: {0}", ZCUtils.repr(ex.ToString()));
                }
            }

            void IChatBot.Stop()
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
            string IChatBot.Converse(string input, string from)
            {
                if (!bStarted)
                {
                    return null;
                }

                hostApp.Log(LogType.DBG, "chatBot > {0}", ZCUtils.repr(input));
                p.StandardInput.WriteLine(input);

                // TBD: Implement some sort of timeout or something so we don't hang forever if things go wonky
                string line = p.StandardOutput.ReadLine();
                hostApp.Log(LogType.DBG, "chatBot < {0}", ZCUtils.repr(line));

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
