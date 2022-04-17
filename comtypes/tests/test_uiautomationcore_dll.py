from dataclasses import dataclass, field
import gc
from typing import List

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
		iuia = comtypes.CoCreateInstance(
            uia_core.CUIAutomation().IPersist_GetClassID(),
            interface=uia_core.IUIAutomation,
            clsctx=comtypes.CLSCTX_INPROC_SERVER
        )
		true_cond = iuia.CreateTrueCondition()
		root = iuia.GetRootElement()
