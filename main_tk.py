from tkinter import Tk
from screen_on.itaskbarlist3 import ITaskBarList3, ctypes


if __name__ == '__main__':
    root = Tk()
    root.title("Screen On")
    root.state('iconic')
    root.bind('<Map>', lambda event: root.destroy())

    # don't group (called before `update`)
    ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID('winand.screen_on')
    root.update()
    root.iconbitmap('assets/sun_white.ico') # called after `update`

    bar = ITaskBarList3()
    top_level_hwnd = ctypes.windll.user32.GetParent(root.winfo_id())
    bar.SetProgressValue(top_level_hwnd, 1, 1)
    bar.SetProgressState(top_level_hwnd, 8)

    # https://stackoverflow.com/questions/49045701/prevent-screen-from-sleeping-with-c-sharp
    # https://docs.microsoft.com/en-us/windows/win32/api/winbase/nf-winbase-setthreadexecutionstate
    # https://github.com/pedrolcl/screensaver-disabler
    # https://stackoverflow.com/questions/64870484
    ES_CONTINUOUS = 0x80000000
    ES_SYSTEM_REQUIRED = 0x00000001
    ES_DISPLAY_REQUIRED = 0x00000002
    ctypes.windll.kernel32.SetThreadExecutionState(ES_CONTINUOUS | ES_DISPLAY_REQUIRED)

    root.mainloop() 
