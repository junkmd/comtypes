import time
import unittest

import comtypes.client
import comtypes.test

# XXX leaks references.

comtypes.test.requires("ui")

def setUpModule():
    raise unittest.SkipTest("External test dependencies like this seem bad.  Find a different "
                            "built-in win32 API to use.")


class _Sink(object):
    def __init__(self):
        self.events = []

    # Word Application Event
    def DocumentChange(self, this, *args):
        self.events.append("DocumentChange")


class Test(unittest.TestCase):
    def test(self):
        # create a word instance
        word = comtypes.client.CreateObject("Word.Application")
        from comtypes.gen import Word

        # Get the instance again, and receive events from that
        w2 = comtypes.client.GetActiveObject("Word.Application")
        sink = _Sink()
        conn = comtypes.client.GetEvents(w2, sink=sink)

        word.Visible = 1

        doc = word.Documents.Add()
        wrange = doc.Range()
        for i in range(10):
            wrange.InsertAfter("Hello from comtypes %d\n" % i)

        for i, para in enumerate(doc.Paragraphs):
            f = para.Range.Font
            f.ColorIndex = i+1
            f.Size = 12 + (2 * i)

        time.sleep(0.5)

        doc.Close(SaveChanges = Word.wdDoNotSaveChanges)

        word.Quit()
        del word, w2

        time.sleep(0.5)

        self.assertEqual(sink.events, ["DocumentChange", "DocumentChange"])

    def test_commandbar(self):
        word = comtypes.client.CreateObject("Word.Application")
        word.Visible = 1
        tb = word.CommandBars("Standard")
        btn = tb.Controls[1]

        if 0: # word does not allow programmatic access, so this does fail
            evt = word.VBE.Events.CommandBarEvents(btn)
            from comtypes.gen import Word, VBIDE
            comtypes.client.ShowEvents(evt, interface=VBIDE._dispCommandBarControlEvents)
            comtypes.client.ShowEvents(evt)

        word.Quit()

if __name__ == "__main__":
    unittest.main()
