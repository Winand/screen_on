import ctypes
import logging
from ctypes import (HRESULT, POINTER, Structure, byref, c_short, c_ubyte,
                    c_uint, c_void_p, oledll)
from ctypes.wintypes import HWND, INT, ULARGE_INTEGER
from types import FunctionType
from typing import Union as U
from typing import get_args, get_type_hints

ole32 = oledll.ole32
CLSCTX_INPROC_SERVER = 0x1
CLSID_TaskbarList = "{56FDF344-FD6D-11d0-958A-006097C9A090}"
IID_ITaskbarList3 = "{EA1AFB91-9E28-4B86-90E9-9E9F8A5EEFAF}"

class TBPF:  # TBPFLAG
    NOPROGRESS = 0
    INDETERMINATE = 1
    NORMAL = 2  # green
    ERROR = 4  # red
    PAUSED = 8  # yellow

class DT:
    "Additional data types"
    TBPFLAG = U[INT, int]
    HWND = U[HWND, int]
    INT = U[INT, int]
    ULONGLONG = U[ULARGE_INTEGER, int]

ctypes.windll.ole32.CoInitialize(None)  # instead of `import pythoncom`


class Guid(Structure):
    _fields_ = [("Data1", c_uint),
                ("Data2", c_short),
                ("Data3", c_short),
                ("Data4", c_ubyte*8)]
                
    def __init__(self, name):
        ole32.CLSIDFromString(name, byref(self))


def gen_method(ptr, method_index, *arg_types):
    # https://stackoverflow.com/a/49053176
    # https://stackoverflow.com/a/12638860
    vtable = ctypes.cast(ptr, POINTER(c_void_p))
    wk = c_void_p(vtable[0])
    function = ctypes.cast(wk, POINTER(c_void_p))
    WFC = ctypes.WINFUNCTYPE(HRESULT, c_void_p, *arg_types)
    METH = WFC(function[method_index])
    return lambda *args: METH(ptr, *args)


def create_instance(clsid, iid):
    ptr = c_void_p()
    ole32.CoCreateInstance(byref(Guid(clsid)), 0, CLSCTX_INPROC_SERVER,
                           byref(Guid(iid)), byref(ptr))
    return ptr


class IUnknown:
    clsid, iid = None, None
    _methods_ = {}

    def __init__(self):
        "Creates an instance and generates methods"
        self.ptr = create_instance(self.clsid, self.iid)
        self.__generate_methods_from_class(IUnknown)
        self.__generate_methods_from_class(self.__class__)
        self.__generate_methods_from_dict(self._methods_)

    def Release(self):
        "index: 2"

    def __del__(self):
        if self.ptr:
            self.Release()

    def isAccessible(self):
        return bool(self.ptr)

    def __generate_methods_from_dict(self, methods: dict):
        """
        Methods are described in a `_methods_` dict of a class:
        _methods_ = {
            "Method1": {'index': 1, 'args': {"hwnd": DT.HWND}},
            "Method2": {'index': 5, 'args': (HWND, INT)},
            "Method3": {'index': 6},
        }
        """
        for name, info in methods.items():
            if hasattr(self, name):
                logging.warning("Overriding existing method %s", name)
            args = info.get('args', ())
            if isinstance(args, dict):
                args = tuple(args.values())
            args = tuple((get_args(i) or [i])[0] for i in args)
            setattr(self, name, gen_method(self.ptr, info['index'], *args))

    def __generate_methods_from_class(self, cls):
        """
        Method argument types are specified in type hints of a Python method.
        If a type hint is a Union then its first element is used.
        Method index is specified on the 1st line of a doc string like this "index: 1"
        """
        for func_name, func in cls.__dict__.items():
            if not isinstance(func, FunctionType) or not func.__doc__:
                continue
            check_idx = func.__doc__.partition('\n')[0].split(":")
            if len(check_idx) != 2:
                continue
            s_index, index = (i.strip() for i in check_idx)
            if s_index.lower() != "index" or not index.isdecimal():
                raise ValueError("Specify `index:<int>` on the first line of doc string")
            setattr(self, func_name,
                gen_method(
                    self.ptr, int(index),
                    *((get_args(i) or [i])[0] for i in get_type_hints(func).values())
                )
            )


class ITaskBarList3(IUnknown):
    """
    https://learn.microsoft.com/en-us/windows/win32/api/shobjidl_core/nn-shobjidl_core-itaskbarlist3
    """
    clsid, iid = CLSID_TaskbarList, IID_ITaskbarList3

    def SetProgressValue(self, hwnd: DT.HWND, ullCompleted: DT.ULONGLONG,
                         ullTotal: DT.ULONGLONG):
        """ index: 9
        Displays or updates a progress bar hosted in a taskbar button to show
        the specific percentage completed of the full operation.
        """

    def SetProgressState(self, hwnd: DT.HWND, tbpFlags: DT.TBPFLAG):
        """ index: 10
        Sets the type and state of the progress indicator displayed on
        a taskbar button.
        """
