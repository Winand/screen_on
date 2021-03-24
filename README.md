# Screen On
Screen On keeps your screen on while running. [SetThreadExecutionState](https://docs.microsoft.com/en-us/windows/win32/api/winbase/nf-winbase-setthreadexecutionstate) system call is used to achieve that.

![](assets/screenshot.png)

Active state is indicated by yellow icon in Windows taskbar. The program turns off when it gets focus e. g. the icon is clicked. A shortcut is supposed to be pinned to taskbar so Screen On can be easily activated.
