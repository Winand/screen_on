from tkinter import Tk
from itaskbarlist3 import ITaskBarList3, ctypes


if __name__ == '__main__':
    root = Tk()
    root.title("Screen On")
    root.attributes('-alpha', 0)
    root.state('iconic')
    root.bind('<Map>', lambda event: root.destroy())

    # don't group (called before `update_idletasks`)
    ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID('winand.screen_on')
    root.update_idletasks()  # https://stackoverflow.com/a/29159152
    root.iconbitmap('assets/sun_white.ico') # called after `update_idletasks`

    bar = ITaskBarList3()
    top_level_hwnd = int(root.wm_frame(), 16)
    bar.SetProgressValue(top_level_hwnd, 1, 1)
    bar.SetProgressState(top_level_hwnd, 8)

    # https://stackoverflow.com/q/49045701 Prevent screen from sleeping with C#
    # https://docs.microsoft.com/en-us/windows/win32/api/winbase/nf-winbase-setthreadexecutionstate
    # https://github.com/pedrolcl/screensaver-disabler
    # https://stackoverflow.com/q/64870484 prevent display from turning off
    ES_CONTINUOUS = 0x80000000
    ES_SYSTEM_REQUIRED = 0x00000001
    ES_DISPLAY_REQUIRED = 0x00000002
    ctypes.windll.kernel32.SetThreadExecutionState(ES_CONTINUOUS | ES_DISPLAY_REQUIRED)

    root.mainloop() 
