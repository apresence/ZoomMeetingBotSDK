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

**Don't bother trying to clone the code just yet!**  While it is already live and managing several meetings, the code is not yet in a state where it can be easily deployed.  I'd hate for someone to clone the code and get frustrated trying to get it to work.

I'm working on cleaning up the code now, including the following:
- [ ] Create a release that's easy to install, configure and use
- [ ] Create a sample configuration that's ready to go with minimal tweaking
- [x] Do not commit private files/information such as zoom username, password, etc. and other content that is not needed
- [X] Make chat bots pluggable
- [X] Move SimpleBot chatbot out into it's own DLL
- [X] Clean up ChatterBot wrapper; the bits are all over the place
- [X] Add support for direct Zoom account login (Was previously depending on Google SSO)
- [ ] Create a Quick Start Guide
- [ ] Re-factor to use the native Windows Zoom SDK which appears to be stable as of Jan 2021

# Background

Since the lock down that started in early 2020 caused by the [COVID-19 pandemic](https://en.wikipedia.org/wiki/COVID-19_pandemic), many people moved their previously physical meetings online to [Zoom](https://zoom.us/).  Suddenly, teachers, soccer coaches, therapists and self-help groups that didn't know anything about setting up or hosting a virtual meeting were thrust blindly into doing so.  There was lots of frustration and confusion, and [Zoombombing](https://en.wikipedia.org/wiki/Zoombombing) reigned supreme.

Partly because I needed a "Covid Project" to keep me sane during shelter-in-place quarantine, and partly because I was curious if I could help to make managing Zoom meetings easier, I embarked upon this project.

I explored many different ways to build this, including stitching together the various requisite [Zoom APIs](https://marketplace.zoom.us/docs/api-reference/zoom-api), using the [Zoom Client SDK for Windows](https://marketplace.zoom.us/docs/sdk/native-sdks/windows/mastering-sdk/windows-sdk-functions) and the [Zoom Client C# Wrapper for Windows](https://marketplace.zoom.us/docs/sdk/native-sdks/windows/c-sharp-wrapper).  At the time (March 2020), none of the SDKs provided all of the functionality I needed.  The closest match was the Zoom Client C# Wrapper for Windows, but it was an incomplete implementation.  I tried the native C++ Windows Client SDK, but I ran into several serious issues, including crashes.

I tried creating a wrapper DLL that intercepted calls between different parts of the client, but that became too complicated and in any case would likely break with new Zoom releases.

I have worked on several projects in the past which involved automating GUI applications that did not have SDKs.  I was able to do so by simulating inputs (mouse movements, clicks, key presses), looking for windows by title or class, reading text in them, etc.  So I decided to go that route here as well.

It turned out to not be so easy to pull text from Zoom windows due to the way it was implemented, so I started looking into several image recognition libraries.  None of the free ones I could get my hand on had any reliable accuracy level.  I did find some paid ones, but they were prohibitively expensive, and even then they still weren't 100% accurate.

I thought I might have to give up on the whole idea until I discovered that Zoom implements a [UIAutomation](https://docs.microsoft.com/en-us/dotnet/framework/ui-automation/ui-automation-overview) interface.  UIAutomation is primarily used for accessibilty reasons, but is also sometimes used for automated testing.  Basically, it's an interface that allows a client to hook into events and invoke user interface elements in another application.

Alas, UIAutomation is inherently slow, and Zoom's implementation was neither complete nor entirely stable.

Ultimately, I ended going up with a hybrid approach, using Windows' [SendInput](https://docs.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-sendinput) API Function to control the app where I could, and UIAutomation where I couldn't.

# Requirements

A Windows system with an interactive desktop session and a user with administrative rights are required.  The SDK has been tested on Windows 10 and Windows Server 2019.

The SDK works by controlling the [Zoom Meeting Desktop Client for Windows](https://zoom.us/client/latest/ZoomInstaller.exe), which has to be installed.

If you wish to modify the code, [Visual Studio Community 2019](https://visualstudio.microsoft.com/vs/community/) and the [.NET Framework 4.8](https://devblogs.microsoft.com/dotnet/announcing-the-net-framework-4-8/) SDK are required.

# Caveats

Since the SDK currently works by manipulating the Zoom user interface, changes to the Zoom Meeting Client can cause it to stop working.  For example, if it's looking for a dialog named "Rename User" that gets changed at some point to "Change User Name", the SDK will not be able to find it.  Some steps have been taken to detect and compensate for these changes, but from time to time you can expect it to stop working when Zoom releases an incompatible update.

In other words: **This is a highly experimental project that should not be used in production.**