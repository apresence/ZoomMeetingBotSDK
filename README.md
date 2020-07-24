# ZoomController
Programmatic C# interface to control the Zoom Client for Meetings on Windows.

Operations include starting/ending meetings, admitting participants in the waiting room, co-hosting/muting/unmuting/renaming participants, sending/receiving public and private chat messages, changing current meeting settings, etc.

Example project includes implementation of a basic bot that recognizes and automatically admits and/or co-hosts known participants, as well as provides basic chat bot functionality.

# Requirements

ZoomController works by controlling the Zoom Client for Meetings on Windows, so a Windows system is required.  The controller hooks into events from the client, sending keystrokes and/or clicks to invoke actions.  As such, it's not the best idea to run ZoomController on a computer system that a human is also trying to use.  To work around that issue, I recommend running ZoomController on a dedicated Windows VM.

# Caveats

Since ZoomController manipulates the Zoom user interface, changes to the Zoom Client for Meetings can cause it to stop working.  For example, if it's looking for a dialog named "Rename User" that gets changed at some point to "Change User Name", ZoomController will not be able to find it.  Some steps have been taken to detect and compensate for these changes, but from time to time you can expect ZoomController to stop working when the Zoom releases an incompatible update.

In other words: This is a highly experimental project that should not be used in production.

# Background

Since the Covid-19 pandemic started in early 2020, many people moved their previously physical meetings online to Zoom.  Suddenly, people who didn't know anything about hosting a virtual meeting got thrust into the role.  There were lots of problems and lots of confusion, and Zoom Bombers reigned supreme.  Partly because I needed a "Covid Project" to keep me sane, and partly because I was curious if I could help to alleviate some of these issues, I looked into the various available Zoom APIs to see if I could build some automation around it.  Unfortunately, none of the APIs gave access to all of the in-meeting functionality I thought I'd need, so I decided to create something that controlled the Zoom Client for Meetings directly.  This gives the controller complete control over the meeting, just as a person sitting at the computer would have.

I explored many different ways to build this, including stitching together the various requisite Zoom APIs, using the Zoom SDK to build my own client, controlling the existing Zoom Client for Meetings using image recognition, and creating a wrapper DLL that intercepted calls between different parts of the client.  Ultimately, I settled on using [UIAutomation](https://docs.microsoft.com/en-us/dotnet/framework/ui-automation/ui-automation-overview).
