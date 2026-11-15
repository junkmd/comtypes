import gc
from collections.abc import Callable
from ctypes import POINTER, Structure, WinDLL, WinError, byref, c_size_t, sizeof
from ctypes.wintypes import BOOL, DWORD, HANDLE

################################################################


class PROCESS_MEMORY_COUNTERS(Structure):
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


_psapi = WinDLL("psapi")

_GetProcessMemoryInfo = _psapi.GetProcessMemoryInfo
_GetProcessMemoryInfo.argtypes = [HANDLE, POINTER(PROCESS_MEMORY_COUNTERS), DWORD]
_GetProcessMemoryInfo.restype = BOOL


def get_working_set_size() -> int:
    """Return the working set size (memory used by process)."""
    pmi = PROCESS_MEMORY_COUNTERS()
    if not _GetProcessMemoryInfo(-1, byref(pmi), sizeof(pmi)):
        raise WinError()
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
