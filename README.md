# ZoomMeetingBotSDK

This project provides a C# SDK which aims to make implementing Bots to manage Zoom meetings easier.

The goal is to relieve meeting hosts of much of the burden involved with running meetings so that they can focus on actually *meeting* in their meetings.

A reference Bot called "UsherBot" is included, which provides the following features:
* Start and stop meetings on-demand
* Automatically manage waiting room by admitting known particpants and throttling entry of others
* Automatically manage co-hosts so that each one does not have to be set up in the Zoom account.  This is especially useful for meetings that have rotating co-hosts
* In-meeting chat commands to perform various functions including co-hosting/demoting and muting/unmuting participants, and other such functions
* A pluggable system to experiment with various ChatBots so the Bot can engage in banter with participants.  A wrapper around the [ChatterBot](https://pypi.org/project/ChatterBot/) python package is used for this purpose
* It can "speak" using text-to-speech
* It can automatically send emails via Gmail on demand

**Don't bother trying to clone the code just yet!**  While it is already live and managing several meetings, the code is not yet in a state where it can be easily deployed.  I'd hate for someone to clone the code and get frustrated trying to get it to work.  If you're interested in getting it working, shoot me an email and I'll help get you set up.

I'm working on the following next:
- [ ] Create a release that's easy to install, configure and use
- [ ] Create a sample configuration that's ready to go with minimal tweaking
- [ ] Create a Quick Start Guide

# Background

Since the lock down that started in early 2020 caused by the [COVID-19 pandemic](https://en.wikipedia.org/wiki/COVID-19_pandemic), many people moved their previously physical meetings online to [Zoom](https://zoom.us/).  Suddenly, teachers, soccer coaches, therapists and self-help groups that didn't know anything about setting up or hosting a virtual meeting were thrust blindly into doing so.  There was lots of frustration and confusion, and [Zoombombing](https://en.wikipedia.org/wiki/Zoombombing) reigned supreme.

Partly because I needed a "Covid Project" to keep me sane during shelter-in-place quarantine, and partly because I was curious if I could help to make managing Zoom meetings easier, I embarked upon this project.

I explored many different ways to build this, including stitching together the various requisite [Zoom APIs](https://marketplace.zoom.us/docs/api-reference/zoom-api), using the [Zoom Client SDK for Windows](https://marketplace.zoom.us/docs/sdk/native-sdks/windows/mastering-sdk/windows-sdk-functions) and the [Zoom Client C# Wrapper for Windows](https://marketplace.zoom.us/docs/sdk/native-sdks/windows/c-sharp-wrapper).  At the time (March 2020), none of the SDKs provided all of the functionality I needed.  The closest match was the Zoom Client C# Wrapper for Windows, but it was an incomplete implementation.  I tried the native C++ Windows Client SDK, but I ran into several serious issues, including crashes.

I tried creating a wrapper DLL that intercepted calls between different parts of the client, but that became too complicated and in any case would likely break with new Zoom releases.

I have worked on several projects in the past which involved automating GUI applications that did not have SDKs.  I was able to do so by simulating inputs (mouse movements, clicks, key presses), looking for windows by title or class, reading text in them, etc.  So I decided to go that route here as well.

It turned out to not be so easy to pull text from Zoom windows due to the way it was implemented, so I started looking into several image recognition libraries.  None of the free ones I could get my hand on had any reliable accuracy level.  I did find some paid ones, but they were prohibitively expensive, and even then they still weren't 100% accurate.

I thought I might have to give up on the whole idea until I discovered that Zoom implements a [UIAutomation](https://docs.microsoft.com/en-us/dotnet/framework/ui-automation/ui-automation-overview) interface.  UIAutomation is primarily used for accessibilty reasons, but is also sometimes used for automated testing.  Basically, it's an interface that allows a client to hook into events and invoke user interface elements in another application.  Alas, UIAutomation is inherently slow, and Zoom's implementation was neither complete nor entirely stable.

Ultimately, I ended going up with a hybrid approach, using Windows' [SendInput](https://docs.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-sendinput) API Function to control the app where I could, and UIAutomation where I couldn't.  This worked quite well, but was very slow, and sometimes the timing was off causing some operation to fail, so retries had to be implemented, etc.  I've archived this code into the [uiautomation](/apresence/ZoomMeetingBotSDK/tree/uiautomation) branch.  If you are working on a project that uses UIAutomation and/or the SendInput APIs, you may find referencing this code useful.

In January of 2021 or so, I looked into the [Zoom Client C# Wrapper for Windows](https://marketplace.zoom.us/docs/sdk/native-sdks/windows/c-sharp-wrapper) and found that it now had (mostly) what I needed.  After adding a few code fixes of my own, I was able to refactor everything to use this instead of the previous UIAutomation/hybrid method.  The resulting product is _much_ faster, and _much_ more stable.  The one thing that is missing from this SDK is the ability to send messages to participants in the waiting room.  I have [posted a message](https://devforum.zoom.us/t/how-to-send-chat-messages-to-everyone-in-waiting-room/39538/3) on Zoom's support board requesting for this to be added.

# Requirements

This project must be run on a Windows system.  This latest version can be run along side a regular Zoom participant session.  In other words, you can join a meeting as a regular participant on the same system the bot is running on.  However, if you are going to use the bot to automatically manage a meeting for you, I reccomend setting up a dedicated system for this purpose.  Any old laptop/PC with good enough hardware to run Zoom should be fine.  You can also run it on a VM in AWS, Azure or the like.

The SDK has been tested on Windows 10 and Windows Server 2019.  Since it requires user interface support, Windows Server Core systems are not supported.

If you wish to modify the code, [Visual Studio Community 2019](https://visualstudio.microsoft.com/vs/community/) and the [.NET Framework 4.8](https://devblogs.microsoft.com/dotnet/announcing-the-net-framework-4-8/) SDK are required.

# Caveats

This project uses the [Zoom Client C# Wrapper for Windows](https://marketplace.zoom.us/docs/sdk/native-sdks/windows/c-sharp-wrapper) which isn't actively maintained or supported by Zoom.  While I have been maintaining my own copy and updating it as needed, there is no guarantee that a new release of Zoom won't render this SDK completely unusable.

In other words: **This is an experimental, for-fun project.**
