# comtypes.test package.

from __future__ import print_function
import ctypes
import getopt
import os
import sys
import time
import unittest

use_resources = []

def get_numpy():
    '''Get numpy if it is available.'''
    try:
        import numpy
        return numpy
    except ImportError:
        return None

def register_server(source_dir):
    """ Register testing server appropriate for the python architecture.

    ``source_dir`` gives the absolute path to the comtype source in which the
    32- and 64-bit testing server, "AvmcIfc.dll" is defined.

    If the server is already registered, do nothing.

    """
    # The 64-bitness of the python interpreter determines the testing dll to
    # use.
    dll_name = "AvmcIfc_x64.dll" if sys.maxsize > 2**32 else "AvmcIfc.dll"
    dll_path = os.path.join(source_dir, "Debug", dll_name)
    # Register our ATL COM tester dll
    dll = ctypes.OleDLL(dll_path)
    dll.DllRegisterServer()
    return

class ResourceDenied(Exception):
    """Test skipped because it requested a disallowed resource.

    This is raised when a test calls requires() for a resource that
    has not be enabled.  Resources are defined by test modules.
    """

def is_resource_enabled(resource):
    """Test whether a resource is enabled.

    If the caller's module is __main__ then automatically return True."""
    if sys._getframe().f_back.f_globals.get("__name__") == "__main__":
        return True
    result = use_resources is not None and \
           (resource in use_resources or "*" in use_resources)
    if not result:
        _unavail[resource] = None
    return result

_unavail = {}
def requires(resource, msg=None):
    """Raise ResourceDenied if the specified resource is not available.

    If the caller's module is __main__ then automatically return True."""
    # see if the caller's module is __main__ - if so, treat as if
    # the resource was set
    if sys._getframe().f_back.f_globals.get("__name__") == "__main__":
        return
    if not is_resource_enabled(resource):
        if msg is None:
            msg = "Use of the `%s' resource not enabled" % resource
        raise ResourceDenied(msg)

def find_package_modules(package, mask):
    import fnmatch
    if hasattr(package, "__loader__"):
        path = package.__name__.replace(".", os.path.sep)
        mask = os.path.join(path, mask)
        for fnm in package.__loader__._files.keys():
            if fnmatch.fnmatchcase(fnm, mask):
                yield os.path.splitext(fnm)[0].replace(os.path.sep, ".")
    else:
        path = package.__path__[0]
        for fnm in os.listdir(path):
            if fnmatch.fnmatchcase(fnm, mask):
                yield "%s.%s" % (package.__name__, os.path.splitext(fnm)[0])
