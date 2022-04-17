from comtypes.client import CreateObject

import pytest


class Test_Excel_Application:
	def test_CreateObject(self):
		xlapp = CreateObject("Excel.Application")
