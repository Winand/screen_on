import pythoncom
import ctypes
from ctypes import c_uint,c_short,c_ubyte,byref,Structure,oledll, \
                    POINTER, HRESULT, c_void_p
from ctypes.wintypes import HWND, ULARGE_INTEGER, INT
ole32 = oledll.ole32
CLSCTX_INPROC_SERVER = 0x1
CLSID_TaskbarList = "{56FDF344-FD6D-11d0-958A-006097C9A090}"
IID_ITaskbarList3 = "{EA1AFB91-9E28-4B86-90E9-9E9F8A5EEFAF}"


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


def create_instance_ex(obj, clsid, iid):
    "Creates an instance and generates methods"
    ptr = create_instance(clsid, iid)
    for i in obj._methods_:
        setattr(obj, i[0], gen_method(ptr, i[1], *i[2]))
    obj.Release = gen_method(ptr, 2)
    return ptr


class ITaskBarList3:
    ptr = None
    _methods_ = (
        ("SetProgressValue", 9, (HWND,ULARGE_INTEGER,ULARGE_INTEGER)),
        ("SetProgressState", 10, (HWND,INT)),
    )
                
    def __init__(self):
        self.ptr = create_instance_ex(self, CLSID_TaskbarList, 
                                      IID_ITaskbarList3)
                               
    def __del__(self):
        if self.ptr:
            self.Release()
    
    def isAccessible(self):
        return bool(self.ptr)
