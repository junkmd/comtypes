from comtypes.client import CreateObject

import pytest


@pytest.mark.skip(reason="WORKAROUND! Excel is not installed on Environment.")
class Test_Excel_Application:
	def test_CreateObject(self):
		xlapp = CreateObject("Excel.Application")
		xlapp.Quit()
