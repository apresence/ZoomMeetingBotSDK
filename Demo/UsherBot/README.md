# ZoomController
Programmatic C# interface to control the [Zoom Client for Meetings](https://zoom.us/client/latest/ZoomInstaller.exe) on Windows.

Operations include starting/ending meetings, admitting participants in the waiting room, co-hosting/muting/unmuting/renaming participants, sending/receiving public and private chat messages, changing current meeting settings, etc.

Features pluggable chat bot system.

Example project includes implementation of a basic bot that recognizes and automatically admits and/or co-hosts known participants, as well as a chat bot provider that wraps [ChatterBot](https://chatterbot.readthedocs.io/en/stable/).

**Don't bother trying to clone the code just yet!**  I'd hate for someone to clone the code and get frustrated trying to get it to work.  To prevent that, I'm working on an out-of-the-box setup that can be used to start/join a meeting with minimal effort.  I hope to have that done and committed by early October 2020.  Before that, I need to:
- [ ] Create a sample configuration that's ready to go with minimal tweaking
- [ ] Do not commit private files/information such as zoom username, password, etc. and other content that is not needed
- [ ] Make chat bots programatically loadable (Currently they're included in the same project, but I'm not ready to share one of the bots I'm working on just yet)
- [ ] Move RemedialBot out into it's own DLL
- [ ] Fix ChatterBot wrapper; I think I broke it recently
- [ ] Add support for direct Zoom account login (Currently I'm using Google SSO)
- [ ] Create a Quick Start Guide

# Requirements

ZoomController works by controlling the Zoom Client for Meetings on Windows, so a Windows system is required.  The controller hooks into events from the client, sending keystrokes and/or clicks to invoke actions.  As such, it's not the best idea to run ZoomController on a computer system that a human is also trying to use.  To work around that issue, I recommend running ZoomController on a dedicated Windows VM.

If you wish to modify the code, [Visual Studio Community 2019](https://visualstudio.microsoft.com/vs/community/) and the [.NET Framework 4.8](https://devblogs.microsoft.com/dotnet/announcing-the-net-framework-4-8/) SDK are required.

# Caveats

Since ZoomController manipulates the Zoom user interface, changes to the Zoom Client for Meetings can cause it to stop working.  For example, if it's looking for a dialog named "Rename User" that gets changed at some point to "Change User Name", ZoomController will not be able to find it.  Some steps have been taken to detect and compensate for these changes, but from time to time you can expect ZoomController to stop working when Zoom releases an incompatible update.

In other words: This is a highly experimental project that should not be used in production.

# Background

Since the [COVID-19 pandemic](https://en.wikipedia.org/wiki/COVID-19_pandemic) started in early 2020, many people moved their previously physical meetings online to [Zoom](https://zoom.us/).  Suddenly, teachers, soccer coaches, therapists, self-help groups, etc. who didn't know anything about setting up or hosting a virtual meeting got thrust into doing so.  There were lots of problems and lots of confusion, and [Zoombombing](https://en.wikipedia.org/wiki/Zoombombing) reigned supreme.

Partly because I needed a "Covid Project" to keep me sane during shelter-in-place quarantine, and partly because I was curious if I could help to alleviate some of these issues, I looked into the various available Zoom APIs to see if I could build some automation around it.  Unfortunately, none of the APIs gave access to all of the in-meeting functionality I thought I'd need, so I decided to create something that controlled the Zoom Client for Meetings directly.  This gives the controller complete control over the meeting, just as a person sitting at the computer would have.

I explored many different ways to build this, including stitching together the various requisite [Zoom APIs](https://marketplace.zoom.us/docs/api-reference/zoom-api), using the [Zoom SDK](https://marketplace.zoom.us/docs/sdk/native-sdks/windows/mastering-sdk/windows-sdk-functions) and the [Zoom C# wrapper](https://marketplace.zoom.us/docs/sdk/native-sdks/windows/c-sharp-wrapper) to build my own client, controlling the existing Zoom Client for Meetings using image recognition, and creating a wrapper DLL that intercepted calls between different parts of the client.  I had significant show-stopping challenges with all of them.  Ultimately, I settled on using a combination of [UIAutomation](https://docs.microsoft.com/en-us/dotnet/framework/ui-automation/ui-automation-overview), which is basically a way to hook into events from the client and invoke user interface elements, and sending mouse/keyboard events to the client via the Windows [SendInput](https://docs.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-sendinput) API Function.
