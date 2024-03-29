"""
See https://github.com/Winand/com_interfaces
"""

import logging
from ctypes import (HRESULT, POINTER, WINFUNCTYPE, Structure, byref, c_short,
                    c_ubyte, c_uint, c_void_p, cast, oledll)
from ctypes.wintypes import HWND, INT, ULARGE_INTEGER
from types import FunctionType
from typing import Optional, TypeVar
from typing import Union as U
from typing import get_args, get_type_hints

ole32 = oledll.ole32
CLSCTX_INPROC_SERVER = 0x1

ole32.CoInitialize(None)  # instead of `import pythoncom`

class TBPF:  # TBPFLAG
    NOPROGRESS = 0
    INDETERMINATE = 1
    NORMAL = 2  # green
    ERROR = 4  # red
    PAUSED = 8  # yellow

class Guid(Structure):
    _fields_ = [("Data1", c_uint),
                ("Data2", c_short),
                ("Data3", c_short),
                ("Data4", c_ubyte*8)]
                
    def __init__(self, name):
        ole32.CLSIDFromString(name, byref(self))


def interface(cls):
    # if "clsid" not in cls.__dict__ or "iid" not in cls.__dict__:
    #     raise ValueError(f"{cls.__name__}: clsid / iid class variables not found")
    if len(cls.__bases__) != 1:
        # https://stackoverflow.com/questions/70222391/com-multiple-interface-inheritance
        raise TypeError('Multiple inheritance is not supported')
    # if cls.__bases__[0] is object:
    #     if cls.__name__ != 'IUnknown':
    #         logging.warning(f"COM interfaces should be derived from IUnknown, not {cls.__name__}")
    __func_table__ = getattr(cls.__bases__[0], '__func_table__', {}).copy()
    for member_name, member in cls.__dict__.items():
        if not isinstance(member, FunctionType):
            continue
        __com_func__ = getattr(member, '__com_func__', None)
        if not __com_func__:
            continue
        __func_table__[member_name] = __com_func__
    __methods__ = cls.__dict__.get('__methods__')
    if isinstance(__methods__, dict):
        # Collect COM methods from __methods__ dict:
        # __methods__ = {
        #     "Method1": {'index': 1, 'args': {"hwnd": DT.HWND}},
        #     "Method2": {'index': 5, 'args': (HWND, INT)},
        #     "Method3": {'index': 6},
        # }
        for member_name, info in __methods__.items():
            if member_name in __func_table__:
                logging.warning("Overriding existing method %s.%s", cls.__name__, member_name)
            args = info.get('args', ())
            if isinstance(args, dict):
                args = tuple(args.values())
            __func_table__[member_name] = {
                'index': info['index'],
                'args': WINFUNCTYPE(HRESULT, c_void_p,
                    *((get_args(i) or [i])[0] for i in args)
                )
            }
    setattr(cls, '__func_table__', __func_table__)
    return cls

def method(index):
    # https://stackoverflow.com/a/2367605
    def func_decorator(func):
        type_hints = get_type_hints(func)
        # Type of return value is not used.
        # Return type is HRESULT https://stackoverflow.com/a/20733034
        type_hints.pop('return', None)
        func.__com_func__ = {
            'index': index,
            'args': WINFUNCTYPE(HRESULT, c_void_p,
                *((get_args(i) or [i])[0] for i in type_hints.values())
            )
        }
        return func
    return func_decorator

def create_instance(clsid, iid):
    ptr = c_void_p()
    ole32.CoCreateInstance(byref(Guid(clsid)), 0, CLSCTX_INPROC_SERVER,
                           byref(Guid(iid)), byref(ptr))
    return ptr


CArgObject = type(byref(c_void_p()))

class DT:
    "Additional data types for type hints"
    REFIID = U[POINTER(Guid), CArgObject]
    void_pp = U[POINTER(c_void_p), CArgObject]
    TBPFLAG = U[INT, int]
    HWND = U[HWND, int]
    ULONGLONG = U[ULARGE_INTEGER, int]


@interface
class IUnknown:
    """
    The IUnknown interface enables clients to retrieve pointers to other interfaces
    on a given object through the QueryInterface method, and to manage the existence
    of the object through the AddRef and Release methods. All other COM interfaces are
    inherited, directly or indirectly, from IUnknown.
    IUnknown https://learn.microsoft.com/en-us/windows/win32/api/unknwn/nn-unknwn-iunknown
    IUnknown (DCOM) https://learn.microsoft.com/en-us/openspecs/windows_protocols/ms-dcom/2b4db106-fb79-4a67-b45f-63654f19c54c
    IUnknown (source) https://github.com/tpn/winsdk-10/blob/9b69fd26ac0c7d0b83d378dba01080e93349c2ed/Include/10.0.14393.0/um/Unknwn.h#L108
    """
    clsid, iid, __func_table__ = None, "{00000000-0000-0000-C000-000000000046}", {}
    T = TypeVar('T', bound="IUnknown")

    def __init__(self, ptr: Optional[c_void_p]=None):
        "Creates an instance and generates methods"
        self.ptr = ptr or create_instance(self.clsid, self.iid)
        # Access COM methods from Python https://stackoverflow.com/a/49053176
        # ctypes + COM access https://stackoverflow.com/a/12638860
        vtable = cast(self.ptr, POINTER(c_void_p))
        wk = c_void_p(vtable[0])
        functions = cast(wk, POINTER(c_void_p))  # method list
        for func_name, __com_opts__ in self.__func_table__.items():
            # Variable in a loop https://www.coursera.org/learn/golang-webservices-1/discussions/threads/0i1G0HswEemBSQpvxxG8fA/replies/m_pdt1kPQqS6XbdZD6Kkiw
            win_func = __com_opts__['args'](functions[__com_opts__['index']])
            setattr(self, func_name,
                lambda *args, f=win_func: f(self.ptr, *args)
            )
    
    def query_interface(self, IID: type[T]) -> T:
        "Helper method for QueryInterface"
        ptr = c_void_p()
        self.QueryInterface(byref(Guid(IID.iid)), byref(ptr))
        return IID(ptr)

    @method(index=0)
    def QueryInterface(self, riid: DT.REFIID, ppvObject: DT.void_pp) -> HRESULT:
        "Retrieves pointers to the supported interfaces on an object."
        raise NotImplementedError
    
    @method(index=1)
    def AddRef(self) -> HRESULT:
        "Increments the reference count for an interface pointer to a COM object"
        raise NotImplementedError

    @method(index=2)
    def Release(self) -> HRESULT:
        "Decrements the reference count for an interface on a COM object"
        raise NotImplementedError

    def __del__(self):
        if self.ptr:
            self.Release()

    def isAccessible(self):
        return bool(self.ptr)


@interface
class ITaskBarList3(IUnknown):
    """
    Exposes methods that control the taskbar
    https://learn.microsoft.com/en-us/windows/win32/api/shobjidl_core/nn-shobjidl_core-itaskbarlist3
    ITaskBarList https://github.com/tpn/winsdk-10/blob/9b69fd26ac0c7d0b83d378dba01080e93349c2ed/Include/10.0.16299.0/um/ShObjIdl_core.h#L14087
    ITaskBarList2 https://github.com/tpn/winsdk-10/blob/9b69fd26ac0c7d0b83d378dba01080e93349c2ed/Include/10.0.16299.0/um/ShObjIdl_core.h#L14205
    ITaskBarList3 https://github.com/tpn/winsdk-10/blob/9b69fd26ac0c7d0b83d378dba01080e93349c2ed/Include/10.0.16299.0/um/ShObjIdl_core.h#L14382
    """
    clsid = CLSID_TaskbarList = "{56FDF344-FD6D-11d0-958A-006097C9A090}"
    iid = IID_ITaskbarList3 = "{EA1AFB91-9E28-4B86-90E9-9E9F8A5EEFAF}"

    @method(index=9)
    def SetProgressValue(self, hwnd: DT.HWND, ullCompleted: DT.ULONGLONG,
                         ullTotal: DT.ULONGLONG):
        """
        Displays or updates a progress bar hosted in a taskbar button to show
        the specific percentage completed of the full operation.
        """

    @method(index=10)
    def SetProgressState(self, hwnd: DT.HWND, tbpFlags: DT.TBPFLAG):
        """
        Sets the type and state of the progress indicator displayed on
        a taskbar button.
        """
