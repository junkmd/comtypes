from dataclasses import dataclass, field
import os
import subprocess as sp
import time
from typing import Any, List

import comtypes
from comtypes.client import CreateObject, GetEvents, GetModule

import pytest


@pytest.fixture(scope="module", autouse=True)
def _setup_module(cleanup_gen_import):
	GetModule("UIAutomationCore.dll")
	cleanup_gen_import()


class Test_IUIAutomation:
	def test_CoCreateInstanceIUIAutomation(self):
		from comtypes.gen import UIAutomationClient as uia_core

		def FindAllChildren(
			elm: uia_core.IUIAutomationElement,
			cond: uia_core.IUIAutomationCondition,
		) -> List[uia_core.IUIAutomationElement]:
			found = elm.FindAll(uia_core.TreeScope_Children, cond)
			return [found.GetElement(i) for i in range(0, found.Length)]

		iuia = comtypes.CoCreateInstance(
            uia_core.CUIAutomation().IPersist_GetClassID(),
            interface=uia_core.IUIAutomation,
            clsctx=comtypes.CLSCTX_INPROC_SERVER
        )
		popen = sp.Popen("notepad")
		found = []
		while not found:
			cond = iuia.CreateTrueCondition()
			root = iuia.GetRootElement()
			found = [
				elm for elm in FindAllChildren(root, cond)
				if elm.CurrentProcessId == popen.pid
			]
		wnd = found[0]
		assert wnd.CurrentClassName == "Notepad"
		ptn = wnd.GetCurrentPattern(uia_core.UIA_WindowPatternId)
		iface = ptn.QueryInterface(uia_core.IUIAutomationWindowPattern)
		iface.Close()
