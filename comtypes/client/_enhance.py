from ctypes import POINTER
import sys

from comtypes import IUnknown


# for type hinting without function annotations or stubs
if sys.version_info >= (3, 5):
	from typing import cast as _t_cast
else:
	def _t_cast(typ, val):
		return val


_PUnk = POINTER(IUnknown)


class EnhancedImplementation(_PUnk):
	@classmethod
	def from_punk(cls, punk):
		interface = cls.__mro__[1]
		_punk = punk.QueryInterface(interface)
		_punk.__class__ = cls
		return _t_cast(cls, _punk)
