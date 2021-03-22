from screen_on.itaskbarlist3 import ITaskBarList3, ctypes
from PyQt5 import QtWidgets, QtGui, QtCore


class MinWidget(QtWidgets.QWidget):
    def __init__(self):
        super().__init__()
        self.showMinimized()

    def changeEvent(self, e):
        if isinstance(e, QtGui.QWindowStateChangeEvent):
            if self.windowState() == QtCore.Qt.WindowNoState:
                self.close()
        

if __name__ == '__main__':
    app = QtWidgets.QApplication([])
    w = MinWidget()
    bar = ITaskBarList3()
    bar.SetProgressValue(int(w.winId()), 1, 1)
    bar.SetProgressState(int(w.winId()), 8)

    # https://stackoverflow.com/questions/49045701/prevent-screen-from-sleeping-with-c-sharp
    # https://docs.microsoft.com/en-us/windows/win32/api/winbase/nf-winbase-setthreadexecutionstate
    # https://github.com/pedrolcl/screensaver-disabler
    # https://stackoverflow.com/questions/64870484
    ES_CONTINUOUS = 0x80000000
    ES_SYSTEM_REQUIRED = 0x00000001
    ES_DISPLAY_REQUIRED = 0x00000002
    ctypes.windll.kernel32.SetThreadExecutionState(ES_CONTINUOUS | ES_DISPLAY_REQUIRED)
    app.exec()
