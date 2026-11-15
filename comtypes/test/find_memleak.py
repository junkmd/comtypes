import gc
from collections.abc import Callable
from ctypes import (
    Structure,
    WinDLL,
    WinError,
    byref,
    c_size_t,
    c_void_p,
    sizeof,
)
from ctypes import c_ulonglong as ULONGLONG
from ctypes.wintypes import BOOL, DWORD, HANDLE
from typing import TYPE_CHECKING, TypeVar

if TYPE_CHECKING:
    from _ctypes import _CField
    from ctypes import _SimpleCData

    _CIntLikeT = TypeVar("_CIntLikeT", bound=_SimpleCData[int])
    _CIntLikeField = _CField[_CIntLikeT, int, _CIntLikeT | int]

################################################################


class PROCESS_MEMORY_COUNTERS(Structure):
    if TYPE_CHECKING:
        cb: _CIntLikeField[DWORD]
        PageFaultCount: _CIntLikeField[DWORD]
        PeakWorkingSetSize: _CIntLikeField[c_size_t]
        WorkingSetSize: _CIntLikeField[c_size_t]
        QuotaPeakPagedPoolUsage: _CIntLikeField[c_size_t]
        QuotaPagedPoolUsage: _CIntLikeField[c_size_t]
        QuotaPeakNonPagedPoolUsage: _CIntLikeField[c_size_t]
        QuotaNonPagedPoolUsage: _CIntLikeField[c_size_t]
        PagefileUsage: _CIntLikeField[c_size_t]
        PeakPagefileUsage: _CIntLikeField[c_size_t]

    _fields_ = [
        ("cb", DWORD),
        ("PageFaultCount", DWORD),
        ("PeakWorkingSetSize", c_size_t),
        ("WorkingSetSize", c_size_t),
        ("QuotaPeakPagedPoolUsage", c_size_t),
        ("QuotaPagedPoolUsage", c_size_t),
        ("QuotaPeakNonPagedPoolUsage", c_size_t),
        ("QuotaNonPagedPoolUsage", c_size_t),
        ("PagefileUsage", c_size_t),
        ("PeakPagefileUsage", c_size_t),
    ]

    def __init__(self):
        self.cb = sizeof(self)

    def dump(self):
        for n, _ in self._fields_[2:]:
            print(n, getattr(self, n) / 1e6)


class PROCESS_MEMORY_COUNTERS_EX(Structure):
    if TYPE_CHECKING:
        cb: _CIntLikeField[DWORD]
        PageFaultCount: _CIntLikeField[DWORD]
        PeakWorkingSetSize: _CIntLikeField[c_size_t]
        WorkingSetSize: _CIntLikeField[c_size_t]
        QuotaPeakPagedPoolUsage: _CIntLikeField[c_size_t]
        QuotaPagedPoolUsage: _CIntLikeField[c_size_t]
        QuotaPeakNonPagedPoolUsage: _CIntLikeField[c_size_t]
        QuotaNonPagedPoolUsage: _CIntLikeField[c_size_t]
        PagefileUsage: _CIntLikeField[c_size_t]
        PeakPagefileUsage: _CIntLikeField[c_size_t]
        PrivateUsage: _CIntLikeField[c_size_t]

    _fields_ = [
        ("cb", DWORD),
        ("PageFaultCount", DWORD),
        ("PeakWorkingSetSize", c_size_t),
        ("WorkingSetSize", c_size_t),
        ("QuotaPeakPagedPoolUsage", c_size_t),
        ("QuotaPagedPoolUsage", c_size_t),
        ("QuotaPeakNonPagedPoolUsage", c_size_t),
        ("QuotaNonPagedPoolUsage", c_size_t),
        ("PagefileUsage", c_size_t),
        ("PeakPagefileUsage", c_size_t),
        ("PrivateUsage", c_size_t),
    ]

    def __init__(self):
        self.cb = sizeof(self)


class PROCESS_MEMORY_COUNTERS_EX2(Structure):
    if TYPE_CHECKING:
        cb: _CIntLikeField[DWORD]
        PageFaultCount: _CIntLikeField[DWORD]
        PeakWorkingSetSize: _CIntLikeField[c_size_t]
        WorkingSetSize: _CIntLikeField[c_size_t]
        QuotaPeakPagedPoolUsage: _CIntLikeField[c_size_t]
        QuotaPagedPoolUsage: _CIntLikeField[c_size_t]
        QuotaPeakNonPagedPoolUsage: _CIntLikeField[c_size_t]
        QuotaNonPagedPoolUsage: _CIntLikeField[c_size_t]
        PagefileUsage: _CIntLikeField[c_size_t]
        PeakPagefileUsage: _CIntLikeField[c_size_t]
        PrivateUsage: _CIntLikeField[c_size_t]
        PrivateWorkingSetSize: _CIntLikeField[c_size_t]
        SharedCommitUsage: _CIntLikeField[ULONGLONG]

    _fields_ = [
        ("cb", DWORD),
        ("PageFaultCount", DWORD),
        ("PeakWorkingSetSize", c_size_t),
        ("WorkingSetSize", c_size_t),
        ("QuotaPeakPagedPoolUsage", c_size_t),
        ("QuotaPagedPoolUsage", c_size_t),
        ("QuotaPeakNonPagedPoolUsage", c_size_t),
        ("QuotaNonPagedPoolUsage", c_size_t),
        ("PagefileUsage", c_size_t),
        ("PeakPagefileUsage", c_size_t),
        ("PrivateUsage", c_size_t),
        ("PrivateWorkingSetSize", c_size_t),
        ("SharedCommitUsage", ULONGLONG),
    ]

    def __init__(self):
        self.cb = sizeof(self)


_psapi = WinDLL("psapi")

_GetProcessMemoryInfo = _psapi.GetProcessMemoryInfo
# Use `c_void_p` to accept any of the `PROCESS_MEMORY_COUNTERS*` structures.
_GetProcessMemoryInfo.argtypes = [HANDLE, c_void_p, DWORD]
_GetProcessMemoryInfo.restype = BOOL

_MemInfoT = TypeVar(
    "_MemInfoT",
    bound=PROCESS_MEMORY_COUNTERS
    | PROCESS_MEMORY_COUNTERS_EX
    | PROCESS_MEMORY_COUNTERS_EX2,
)


def _get_current_process_memory_info(typ: type[_MemInfoT]) -> _MemInfoT:
    # -1 is a pseudo-handle representing the current process.
    pmi = typ()
    if not _GetProcessMemoryInfo(-1, byref(pmi), sizeof(pmi)):
        raise WinError()
    return pmi


def get_working_set_size() -> int:
    """Return the working set size (memory used by process)."""
    pmi = _get_current_process_memory_info(PROCESS_MEMORY_COUNTERS)
    return pmi.WorkingSetSize


class FunctionDetector:
    """Detect memory leaks using an injected function to get current memory usage."""

    def __init__(self, get_memory_func: Callable[[], int]):
        self._get_memory_func = get_memory_func
        self._baseline = 0
        self._diff = 0

    def __enter__(self) -> "FunctionDetector":
        # Measure baseline
        self._baseline = self._get_memory_func()
        return self

    def __exit__(self, exc_type, exc_val, exc_tb) -> None:
        current = self._get_memory_func()
        self._diff = current - self._baseline

    def get_diff(self) -> int:
        return self._diff


LOOPS = 10, 1000


def find_memleak(func, loops=LOOPS):
    # call 'func' several times, so that memory consumption
    # stabilizes:
    # Warmup phase
    for j in range(loops[0]):
        for k in range(loops[1]):
            func()
    gc.collect()
    gc.collect()
    gc.collect()
    # call 'func' several times, recording the difference in
    # memory consumption before and after the call.  Repeat this a
    # few times, and return a list containing the memory
    # consumption differences.
    with FunctionDetector(get_working_set_size) as d:
        # Measurement phase
        for j in range(loops[0]):
            for k in range(loops[1]):
                func()
        gc.collect()
        gc.collect()
        gc.collect()
    # Sometimes the process size did decrease, we do not report leaks
    # in this case:
    return max(d.get_diff(), 0)
