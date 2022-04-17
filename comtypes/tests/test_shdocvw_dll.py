from dataclasses import dataclass, field
import gc
import sys
from typing import List

from comtypes import IUnknown
from comtypes.automation import IDispatch
from comtypes.client.lazybind import Dispatch
from comtypes.client import CreateObject, GetEvents, GetModule

import pytest


@pytest.fixture(scope="module", autouse=True)
def _setup_module(cleanup_gen_import):
	GetModule("SHDocVw.dll")
	cleanup_gen_import()


@dataclass
class _Sink:
	events: List[str] = field(init=False, default_factory=list)

	def OnVisible(self, this, Visible):
		self.events.append(f"Onvisible({Visible})")


class Test_WebBrowser:
	def test_CreateObject(self):
		from comtypes.gen import SHDocVw as shdocvw
		browser = CreateObject(shdocvw.WebBrowser)
		sink = _Sink()
		conn = GetEvents(browser, sink)
		browser.Visible = True
		assert sink.events == ["Onvisible(True)"]
		browser.Visible = False
		del conn
		gc.collect()
		assert sink.events == ["Onvisible(True)", "Onvisible(False)"]
