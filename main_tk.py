from tkinter import Tk
from itaskbarlist3 import ITaskBarList3, ctypes, TBPF_PAUSED, TBPF_ERROR

ES_CONTINUOUS = 0x80000000
ES_SYSTEM_REQUIRED = 0x00000001
ES_DISPLAY_REQUIRED = 0x00000002

current_state = -1
states = [  # progress state, execution state
    (TBPF_PAUSED, ES_CONTINUOUS | ES_DISPLAY_REQUIRED),
    (TBPF_ERROR, ES_CONTINUOUS | ES_SYSTEM_REQUIRED),
]


def set_next_state(event=None):
    global current_state
    current_state += 1
    if current_state == len(states):
        root.destroy()  # exit
        return
    # root.withdraw()  # hide window to skip minimize animation
    root.state('iconic')

    bar.SetProgressValue(top_level_hwnd, 1, 1)
    bar.SetProgressState(top_level_hwnd, states[current_state][0])
    # https://stackoverflow.com/questions/49045701/prevent-screen-from-sleeping-with-c-sharp
    # https://docs.microsoft.com/en-us/windows/win32/api/winbase/nf-winbase-setthreadexecutionstate
    # https://github.com/pedrolcl/screensaver-disabler
    # https://stackoverflow.com/questions/64870484
    ctypes.windll.kernel32.SetThreadExecutionState(states[current_state][1])


if __name__ == '__main__':
    root = Tk()
    root.title("Screen On")
    # root.overrideredirect(True)  # no borders but window cannot be minimized
    root.attributes('-alpha', 0)  # hide window animations
    root.state('iconic')
    root.bind('<Map>', set_next_state)

    # don't group (called before `update_idletasks`)
    ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID('winand.screen_on')
    root.update_idletasks()  # https://stackoverflow.com/a/29159152
    root.iconbitmap('assets/sun_white.ico') # called after `update_idletasks`

    bar = ITaskBarList3()
    top_level_hwnd = int(root.wm_frame(), 16)

    set_next_state()

    root.mainloop() 
