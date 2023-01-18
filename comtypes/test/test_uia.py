import subprocess
import time
import unittest as ut

import comtypes
from comtypes.client import GetModule

GetModule("UIAutomationCore.dll")
from comtypes.gen import UIAutomationClient as uiac


class Test_IUIAutomation(ut.TestCase):
    def test_gets_and_closes_wnd(self):
        # Using `Notepad`(which is existed in any environment) to see
        # UIA can get and close windows.
        # Anything more complex than this would be environment-dependent.
        # So tests using `UIAutomationCore.dll` should be limited to simple
        # like this.
        iuia = comtypes.CoCreateInstance(
            uiac.CUIAutomation().IPersist_GetClassID(),
            interface=uiac.IUIAutomation,
            clsctx=comtypes.CLSCTX_INPROC_SERVER,
        )
        popen = subprocess.Popen("Notepad")
        found = []
        start = time.time()
        cond = iuia.CreateTrueCondition()
        root = iuia.GetRootElement()
        while not found:
            arr = root.FindAll(uiac.TreeScope_Children, cond)
            elements = [arr.GetElement(i) for i in range(0, arr.Length)]
            found = [elm for elm in elements if elm.CurrentProcessId == popen.pid]
            if time.time() > start + 60:  # pragma: no cover
                raise OSError  # TimeoutError is new in Python3.3
        wnd = found[0]
        self.assertEqual(wnd.CurrentClassName, "Notepad")
        ptn = wnd.GetCurrentPattern(uiac.UIA_WindowPatternId)
        iface = ptn.QueryInterface(uiac.IUIAutomationWindowPattern)
        iface.Close()
        self.assertEqual(popen.wait(), 0)
