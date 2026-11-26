"""Helpers for running a function in a separate process with fault logging.

These helpers are shared by tests that must isolate a call which can crash
the interpreter (for example, a segfault caused by a COM lifetime bug).
Running the call in a spawned child process keeps such a crash from taking
down the main test process, and enabling `faulthandler` in the child makes
the crash information available to the parent.
"""

import faulthandler
import io
import multiprocessing
import tempfile
from collections.abc import Callable
from pathlib import Path
from typing import Any


def enable_fault_log(tmpfile: Path) -> io.TextIOWrapper:
    """Send faulthandler output to `tmpfile` until the process exits."""
    f = tmpfile.open(mode="w", encoding="utf-8")
    # We keep `faulthandler` enabled and the file open until the process
    # exits so that crash information can be captured.
    # It is only enabled in the child process and does not affect the main
    # process, so there is no need to worry about forgetting to close and
    # disable.
    faulthandler.enable(file=f)
    return f


def run_in_mp(fn: Callable[[Path], Any]) -> tuple[int | None, str]:
    """Run a function in a separate process and return exit code and output.

    Using spawn isolates the process so crashes like segfaults do not affect
    the main test process.
    """
    with tempfile.TemporaryDirectory() as td:
        tmpdir = Path(td)
        tmpfile = tmpdir / "fault.txt"
        ctx = multiprocessing.get_context("spawn")
        process = ctx.Process(target=fn, args=(tmpfile,))
        process.start()
        process.join()
        output = tmpfile.read_text()
    return process.exitcode, output
