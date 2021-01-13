// Christopher Mapes - apresence@hotmail.com
// Created: 2020.06

// GitHub ignore file for Visual Studio:
// https://raw.githubusercontent.com/github/gitignore/master/VisualStudio.gitignore

// W10 SDK - https://developer.microsoft.com/en-us/windows/downloads/windows-10-sdk/
// https://social.msdn.microsoft.com/Forums/windowsdesktop/en-US/674b4951-1c6e-400a-838e-dc72c672a12c/uac-automation-success?forum=windowssecurity
// https://docs.microsoft.com/en-us/windows/win32/winauto/uiauto-securityoverview?f1url=https%3A%2F%2Fmsdn.microsoft.com%2Fquery%2Fdev16.query%3FappId%3DDev16IDEF1%26l%3DEN-US%26k%3Dk(vs.debug.error.launch_elevation_requirements)%3Bk(TargetFrameworkMoniker-.NETFramework%2CVersion%3Dv4.8)%26rd%3Dtrue
// https://docs.microsoft.com/en-us/windows/security/identity-protection/user-account-control/user-account-control-security-policy-settings
// https://ironsoftware.com/csharp/ocr/docs/
// https://docs.microsoft.com/en-us/windows/win32/winauto/uiauto-clientportal
// https://docs.microsoft.com/en-us/windows/win32/winauto/uiauto-howto-query-a-virtualized-item-in-itemsview
// TO DEBUG -- secpol.msc: User Account Control: Only elevate UIAccess applications that are installed in secure locations = False

// https://github.com/tesseract-ocr/tesseract
// https://github.com/UB-Mannheim/tesseract/wiki

// ChatBot:
// - Fix issue with embedded stuff
// - Strip html tags from responses
// - Strip emojis that are not recognized by Zoom from responses
// TBD:
// - Any way to handle users with duplicate names? Idea: Rename dups to #2, #3, etc. Could also use hash of avatar to distinguish
// - Create & move IChatBot and ILogging interfaces out into separate class DLL
// - Move specific ChatBot implementations out into their own DLLs
// - Enumerate ChatBots and allow dynamically loading them
// ? Fix issue where paging through participants gets stuck if # of participants is exactly equal to one page (21 on my screen)
//   - Still not working!!!
// - Implement log file rolling
// - Check/fix "More options to manage all participants" settings when joining the meeting
// - When leaving meeting, add option to pass off to a known co-host
// - Move chat processing into it's own thread
// - Fix () in name, esp rename... recognize users with (Usher), (Speaker) etc. in name
// - Change participant scan retry to do one per pass rather than up to 3 loops per pass
// - Test sending chat when recipient search (edit) box is enabled
// - Set up GitHub or other project
// - Modify Participant menu invocation to check  and update before attempting to execute
// - Add /lookup command (dictionaryapi.com)
// - Add /remember option
// - Fix admission throttling - seems to be broken; Also, make configurable
// - Update participants based on some kind of reliable event instead of polling?
// - Add people to good list if they've attended a certain # of meetings w/ video on
// - Time shares
// - Handle hand raise order?
// - Handle it when the participants panel is embedded in the Zoom window
// - How to prevent 40m idle timeout?
// - Move Chrome stuff into it's own class
// - Add option to re-open Zoom if it is closed
// - Remove Co-Host from unapproved attendees
// - Don't open more than once
// - There's got to be some kind of way to track attendee renames and whatnot.  Maybe rids?
// - Maybe use same API being used for mouse events for kbd events too?
// - Quit meeting when time is up
// - Keep track of first seen, last seen, number of times seen for each user
// - Normalize command line args for remote vs. local runs
// - AIML: https://www.instructables.com/id/How-to-Make-a-Robot-That-Talks-Back-Using-AIML-in-/
// ! Try direct click/type method with screen reco & without UIAutomation
// x Do something about walking participant listItems while in flight
//   - Seems like there is one or more duplicates when this happens.  Could detect that and scan the list again
//   - Could also do refcount - ie: only mark participant as leaving if the name does not show up twice
//   - Could pull participant count from window title and use that to determine if we have a complete list or not
//   - Is there some way to get the list without these issues?
// x Refactor to move all interaction/events to ZoomMeetingBotSDK and add OnParticipantWaiting, OnParticipantJoining, OnParticipantJoin, OnParticipantLeave, etc.
// x Retry to open Zoom if it fails
// x Could improve reliability of selecting menu items by using Focus events to detect which menu item we're on and select the right one
// x Re-write for UIA support
// x Enable UI caching
// x Handling of dialogs is still susceptible to timing issues - Find a better way of handling (Ex: List of handlers, etc.)
// x Support chat!
// x Deal with participant list scrolling - 21 max per page at current resolution.  Uggggh!  -- Maybe try {END} then {UP} until we get to "me" (or up to # participants), or try PgDn?
// x Add /rename X to Y option
// x Add /demote option
// x Deal with audio prompt: zJoinAudioWndClass
// x Citadel mode
// x Config file with dynamic reload, including Global.cfg.bDebug
// x Fix lock screen issue
// x Reclaim host
// x Parse video, audio states properly
// x Fix issue with screen sharing.  Probaby need to detect & restore Zoom window back to smaller size
// x Why so much delay on stand-alone PC?
// x Move all of the ZoomMeetingBotSDK stuff into its own class
// x Track attendee state changes like video on/off
// x Track attendees who leave
// x Immediately join numbers?
// x Fix new console issue.  Need to be able to see errors!
// x One-time "hi"
// x Add /citadel, /lockdown and /passive to chat commands
// x Add /waitmsg support (Waiting room announcement)
// x Fix issue when bot has the same name as someone else
// x Fix issue with scanning chat messages.  Sometimes we get confused and parse messages we've already parsed all over again.  Idea: Use timestamps?
// x Add chatterbot support
// x Re-write InvokeMenuItem with new Hunt logic
// x Rewrite Chat with RID & GetNextSibling logic
// x Rewrite Chat To hunting w/ RID & GetNextSibling logic
// x Rewrite Participant menu item hunting w/ RID & GetNextSibling logic?
// x Move initial chat list item hunting to EH; then we don't have to deal with the timeout
// x Optimize event handler: Only enable when waiting for events -- Could probably shave 2-4 seconds off of chat parsing with this
// x Add /speaker name|off option
// x Optimize all Subtree searches, esp. ClickParticipantWindowControl
// x Optimize typing by copying to clipboard and then pasting
// x Fix issue with accessing sales screen when joining meeting via Chrome
// x Fix "More options to manage all participants" subtree search being so slow
// x Add option to leave meeting
// x Fix issue with actions for attendees not on the first page not working
//   Maybe queue up actions and process them when parsing the list items?
//   This would avoid the issue with trying to do something on an item that is not in view
// x Fix issue with Zoom meeting timeout causing Zoom exe to hang around
// x Login directly to Zoom app with Zoom credentials, bypassing browser.  Add encryption support for password
namespace ZoomMeetingBotSDK
{
    using System;
    using System.Collections.Generic;
    using System.Diagnostics;
    using System.Text;
    using System.Threading;
    using System.Threading.Tasks;
    using System.Windows.Forms;
    using static Utils;

    internal class Program
    {
        public class ProgramSettings
        {
            public ProgramSettings()
            {
                WaitForDebuggerAttach = false;
                PromptOnStartup = false;
            }

            public bool WaitForDebuggerAttach { get; set; }

            public bool PromptOnStartup { get; set; }
        }

        private static HostApp hostApp = null;
        private static UsherBot usherBot = null;

        private static ProgramSettings programSettings = null;

        [STAThread]
        private static int Main(string[] args)
        {
            hostApp = new HostApp();
            hostApp.Init();

            programSettings = DeserializeJson<ProgramSettings>(hostApp.GetSettingsAsJSON());
            WaitDebuggerAttach();

            if (Array.IndexOf(args, "/protect") != -1)
            {
                Console.WriteLine("Unprotected Value:");

                StringBuilder pass = new StringBuilder();
                while (true)
                {
                    ConsoleKeyInfo i = Console.ReadKey(true);
                    if (i.Key == ConsoleKey.Enter)
                    {
                        break;
                    }
                    else if (i.Key == ConsoleKey.Backspace)
                    {
                        if (pass.Length > 0)
                        {
                            pass.Remove(pass.Length - 1, 1);
                            Console.Write("\b \b");
                        }
                    }
                    else if (i.KeyChar != '\u0000') // KeyChar == '\u0000' if the key pressed does not correspond to a printable character, e.g. F1, Pause-Break, etc
                    {
                        pass.Append(i.KeyChar);
                        Console.Write("*");
                    }
                }
                Console.WriteLine();

                string passString = pass.ToString();

                string prot = ProtectedString.Protect(passString);
                string unprot = ProtectedString.Unprotect(prot);

                if (unprot == passString)
                {
                    Console.WriteLine($"\nProtected Value:\n{prot}");
                }
                else
                {
                    Console.WriteLine("ERROR: Decryption test failed");
                }

                /*
                Console.WriteLine();
                Console.WriteLine("Press ENTER to exit");
                Console.ReadLine();
                */

                return 0;
            }

            if (Array.IndexOf(args, "/command") != -1)
            {
                List<string> commands = new List<string>();

                Console.WriteLine("Remote command mode");
                foreach (var arg in args)
                {
                    var arg_l = arg.ToLower();
                    if (arg_l == "/command")
                    {
                        continue;
                    }

                    if (arg_l.StartsWith("/debug:") || arg_l.StartsWith("/citadel:") || (arg_l == "/exit") || (arg_l == "/leave") || (arg_l == "/kill") || (arg_l == "/end") || arg_l.StartsWith("/pause:"))
                    {
                        commands.Add(arg_l.Substring(1));
                    }
                    else
                    {
                        Console.WriteLine("Unsupported command: {0}", arg_l);
                        return 1;
                    }
                }
                if (commands.Count == 0)
                {
                    Console.WriteLine("Command required");
                    return 1;
                }
                UsherBot.WriteRemoteCommands(commands.ToArray());
                return 0;
            }
            UsherBot.ClearRemoteCommands();

            for (int i = 0; i < args.Length; i++)
            {
                var arg = args[i];
                var arg_l = arg.ToLower();
                if (arg_l.Equals("/pause"))
                {
                    UsherBot.SetMode("pause", true);
                }
                else if (arg_l.Equals("/passive"))
                {
                    UsherBot.SetMode("passive", true);
                }
                else if (arg_l.Equals("/prompt"))
                {
                    programSettings.PromptOnStartup = true;
                }
                else if (arg_l.Equals("/waitattach"))
                {
                    programSettings.WaitForDebuggerAttach = true;
                }
                else if (arg_l.Equals("/debug"))
                {
                    UsherBot.SetMode("debug", true);
                }
                else if (arg_l.Equals("/citadel"))
                {
                    UsherBot.SetMode("citadel", true);
                }
                else if (arg_l.Equals("/lockdown"))
                {
                    UsherBot.SetMode("lockdown", true);
                }
                else if (arg_l.Equals("/waitmsg"))
                {
                    UsherBot.cfg.WaitingRoomAnnouncementMessage = args[++i];
                }
                else
                {
                    Console.WriteLine("Unsupported option: {0}", arg_l);
                    return 1;
                }
            }
            WaitDebuggerAttach();

            hostApp.Log(LogType.DBG, "Debug logging enabled");

            //Global.hostApp.Log(LogType.DBG, "Main thread_id=0x{0:X8}", Thread.CurrentThread.ManagedThreadId);

            // TBD: Exit when Zoom app exits
            Task.Factory.StartNew(() =>
                {
                    try
                    {
                        usherBot = new UsherBot();
                        usherBot.Init(new ControlBot.ControlBotInitParam()
                        {
                            hostApp = hostApp,
                        });

                        if (programSettings.PromptOnStartup)
                        {
                            Console.WriteLine("Press ENTER to proceed");
                            Console.ReadLine();
                        }

                        hostApp.Start();
                        usherBot.Start();
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex.ToString());
                    }

                    Console.WriteLine("CONSOLE : === Listening for keystrokes ===");

                    bool paused = false;

                    while (true)
                    {
                        var keyInfo = Console.ReadKey();

                        Console.WriteLine("CONSOLE : === ReadKey {0}, {1}, {2} ===", repr(keyInfo.KeyChar), keyInfo.Key.ToString(), "{" + keyInfo.Modifiers.ToString() + "}");
                        if (keyInfo.Modifiers == 0)
                        {
                            var ch = keyInfo.KeyChar;
                            /*
                            if (ch == 'a')
                            {
                                Console.WriteLine("CONSOLE : LogAETree Requested");
                                Controller.LogAETree();
                            }
                            else
                            */
                            if (ch == 'p')
                            {
                                if (paused)
                                {
                                    paused = false;
                                    hostApp.Log(LogType.INF, "[Console] Pause");
                                    UsherBot.SetMode("pause", paused);
                                }
                                else
                                {
                                    paused = true;
                                    hostApp.Log(LogType.INF, "[Console] Unpause");
                                    UsherBot.SetMode("pause", paused);
                                }
                            }
                            else if (ch == 'k')
                            {
                                hostApp.Log(LogType.INF, "[Console] Kill");

                                UsherBot.LeaveMeeting(true);

                                break;
                            }
                            else if (ch == 'q')
                            {
                                hostApp.Log(LogType.INF, "[Console] Quit");

                                UsherBot.LeaveMeeting(false);
                                usherBot.Stop();

                                break;
                            }
                        }
                    }
                });

            while (!UsherBot.shouldExit)
            {
                Application.DoEvents();
                Thread.Sleep(250);
            }

            hostApp.Log(LogType.INF, "[Program] Cleaning Up");
            if (usherBot != null)
            {
                usherBot.Stop();
            }

            if (hostApp != null)
            {
                hostApp.Log(LogType.INF, "[Program] Stopping HostApp");
                hostApp.Stop();
            }

            Console.WriteLine("Exiting");
            return 0;
        }

        private static void WaitDebuggerAttach(bool force = false)
        {
            if (Debugger.IsAttached || ((!force) && ((programSettings == null) || (!programSettings.WaitForDebuggerAttach))))
            {
                return;
            }

            Console.WriteLine("Waiting for debugger to attach");
            while (!Debugger.IsAttached)
            {
                Thread.Sleep(250);
            }

            Console.WriteLine("Debugger attached");
        }
    }
}
