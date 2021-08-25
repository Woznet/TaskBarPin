# TaskBarPin
pin.exe is a simple tool that allows you to programmatically pin an item to the Windows 10 Taskbar.

Usage:
Pin a Shortcut
pin "C:\path\to\file.lnk"

Unpin a Shortcut
pin "C:\path\to\file.lnk" u

Exit codes:
-1 = printed usage
0 = succeeded
1 = CoInitialize failed
2 = ILCreateFromPathW failed
3 = CoCreateInstance failed
4 = IPinnedList3::Modify failed


Source: https://geelaw.blog/entries/msedge-pins/assets/pin.cc
