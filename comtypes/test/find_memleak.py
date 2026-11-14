import gc
import sys
import tracemalloc
from ctypes import (
    POINTER,
    Structure,
    WinDLL,
    WinError,
    byref,
    c_size_t,
    c_ulonglong,
    sizeof,
)
from ctypes.wintypes import BOOL, DWORD, HANDLE

# SIZE_T is not in ctypes.wintypes on all Python versions; use c_size_t instead
# ULONGLONG is not in ctypes.wintypes; use c_ulonglong instead
SIZE_T = c_size_t
ULONGLONG = c_ulonglong

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


class PROCESS_MEMORY_COUNTERS_EX(Structure):
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
# Use c_void_p for the buffer pointer to accept any of the PROCESS_MEMORY_COUNTERS* structures
from ctypes import c_void_p

_GetProcessMemoryInfo.argtypes = [HANDLE, c_void_p, DWORD]
_GetProcessMemoryInfo.restype = BOOL


def _get_memory_info():
    """Get process memory info using the most detailed structure available.

    Returns a tuple (structure_instance, structure_type) where structure_type
    is one of 'EX2', 'EX', or 'basic'.
    """
    # Try EX2 first (Windows 10 22H2+ / Windows 11 22H2+)
    pmi_ex2 = PROCESS_MEMORY_COUNTERS_EX2()
    if _GetProcessMemoryInfo(-1, byref(pmi_ex2), sizeof(pmi_ex2)):
        return pmi_ex2, "EX2"

    # Try EX (Windows Vista+)
    pmi_ex = PROCESS_MEMORY_COUNTERS_EX()
    if _GetProcessMemoryInfo(-1, byref(pmi_ex), sizeof(pmi_ex)):
        return pmi_ex, "EX"

    # Fall back to basic
    pmi = PROCESS_MEMORY_COUNTERS()
    if _GetProcessMemoryInfo(-1, byref(pmi), sizeof(pmi)):
        return pmi, "basic"

    raise WinError()


def get_private_usage():
    """Return the private memory usage (Commit Charge) in bytes.

    This is a better indicator of actual memory leaks than Working Set Size,
    because it represents the total amount of private memory committed for
    the process, regardless of whether pages are currently resident in RAM.
    """
    pmi, pmi_type = _get_memory_info()
    if pmi_type in ("EX2", "EX"):
        return pmi.PrivateUsage
    # Fall back to PagefileUsage for basic structure (same as PrivateUsage on older Windows)
    return pmi.PagefileUsage


def get_working_set_size():
    """Return the working set size in bytes (for backward compatibility)."""
    pmi, _ = _get_memory_info()
    return pmi.WorkingSetSize


def get_memory_stats():
    """Return a dict with all available memory statistics."""
    pmi, pmi_type = _get_memory_info()
    stats = {
        "type": pmi_type,
        "WorkingSetSize": pmi.WorkingSetSize,
        "PeakWorkingSetSize": pmi.PeakWorkingSetSize,
        "PagefileUsage": pmi.PagefileUsage,
        "PeakPagefileUsage": pmi.PeakPagefileUsage,
        "QuotaPagedPoolUsage": pmi.QuotaPagedPoolUsage,
        "QuotaNonPagedPoolUsage": pmi.QuotaNonPagedPoolUsage,
    }
    if pmi_type in ("EX2", "EX"):
        stats["PrivateUsage"] = pmi.PrivateUsage
    if pmi_type == "EX2":
        stats["PrivateWorkingSetSize"] = pmi.PrivateWorkingSetSize
        stats["SharedCommitUsage"] = pmi.SharedCommitUsage
    return stats


LOOPS = 10, 1000


def find_memleak(func, loops=LOOPS, use_private_usage=True):
    """Detect memory leaks by measuring memory usage before and after repeated calls.

    Args:
        func: Callable to test for memory leaks.
        loops: Tuple of (outer_loops, inner_loops) for warmup and measurement phases.
        use_private_usage: If True, use PrivateUsage (Commit Charge) which is a better
            indicator of actual memory leaks. If False, use WorkingSetSize (legacy behavior).

    Returns:
        The increase in memory usage in bytes (0 if no increase or decrease).

    Note:
        This function uses PrivateUsage (Commit Charge) by default, which tracks
        the total private memory committed for the process. This is more reliable
        than WorkingSetSize for detecting actual memory leaks because:
        - WorkingSetSize only measures pages currently in physical RAM
        - PrivateUsage measures all committed private memory regardless of residency
        - An increase in PrivateUsage directly indicates unreleased allocations
    """
    get_mem = get_private_usage if use_private_usage else get_working_set_size

    # Warmup phase: call func several times to stabilize memory consumption
    for j in range(loops[0]):
        for k in range(loops[1]):
            func()
    gc.collect()
    gc.collect()
    gc.collect()

    # Measure baseline
    baseline = get_mem()

    # Measurement phase: call func repeatedly
    for j in range(loops[0]):
        for k in range(loops[1]):
            func()
    gc.collect()
    gc.collect()
    gc.collect()

    # Measure after
    after = get_mem()

    # Return the increase (0 if decreased or unchanged)
    result = after - baseline
    return max(result, 0)


def find_memleak_with_tracemalloc(func, loops=LOOPS):
    """Detect memory leaks using both native memory (PrivateUsage) and Python tracemalloc.

    This provides a more comprehensive leak detection by checking:
    1. Native memory leaks via PrivateUsage (Commit Charge)
    2. Python object leaks via tracemalloc

    Args:
        func: Callable to test for memory leaks.
        loops: Tuple of (outer_loops, inner_loops) for warmup and measurement phases.

    Returns:
        A tuple (native_leak_bytes, python_leak_stats) where:
        - native_leak_bytes: Increase in PrivateUsage (0 if no increase)
        - python_leak_stats: Dict with tracemalloc statistics or None if not available
    """
    # Start tracemalloc if not already started
    was_tracing = tracemalloc.is_tracing()
    if not was_tracing:
        tracemalloc.start()

    try:
        # Warmup phase
        for j in range(loops[0]):
            for k in range(loops[1]):
                func()
        gc.collect()
        gc.collect()
        gc.collect()

        # Take baseline snapshots
        baseline_native = get_private_usage()
        baseline_snapshot = tracemalloc.take_snapshot()

        # Measurement phase
        for j in range(loops[0]):
            for k in range(loops[1]):
                func()
        gc.collect()
        gc.collect()
        gc.collect()

        # Take after snapshots
        after_native = get_private_usage()
        after_snapshot = tracemalloc.take_snapshot()

        # Calculate native memory leak
        native_leak = max(after_native - baseline_native, 0)

        # Calculate Python memory leak
        top_stats = after_snapshot.compare_to(baseline_snapshot, "lineno")
        python_leak_stats = {
            "total_size_diff": sum(stat.size_diff for stat in top_stats),
            "top_diffs": [
                {
                    "file": str(stat.traceback[0].filename)
                    if stat.traceback
                    else "unknown",
                    "lineno": stat.traceback[0].lineno if stat.traceback else 0,
                    "size_diff": stat.size_diff,
                    "count_diff": stat.count_diff,
                }
                for stat in top_stats[:10]  # Top 10 differences
            ],
        }

        return native_leak, python_leak_stats
    finally:
        if not was_tracing:
            tracemalloc.stop()
