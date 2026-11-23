"""
Test for the ``comtypes.clear_cache`` module.
"""

import contextlib
import runpy
from unittest import TestCase
from unittest.mock import call, patch

from comtypes.client import _find_gen_dir


class ClearCacheTestCase(TestCase):
    # we patch sys.stdout so unittest doesn't show the print statements

    @patch("sys.argv", ["clear_cache.py", "-y"])
    @patch("shutil.rmtree")
    def test_clear_cache(self, mock_rmtree):
        with contextlib.redirect_stdout(None):
            runpy.run_module("comtypes.clear_cache", {}, "__main__")

        # because we don't actually delete anything, _find_gen_dir() will
        # give the same answer every time we call it
        self.assertEqual(
            mock_rmtree.call_args_list, [call(_find_gen_dir()) for _ in range(2)]
        )


class ClearCacheDeclineTestCase(TestCase):
    @patch("sys.argv", ["clear_cache.py"])
    @patch("builtins.input", return_value="n")
    def test_declining_confirmation_does_not_remove(self, mock_input):
        # When the user does not answer "y", ``main`` must print a notice and
        # return without deleting any cache directory.  ``shutil.rmtree`` is
        # deliberately not patched: if the confirmation were ignored, this
        # test would delete real directories.
        with contextlib.redirect_stdout(None):
            runpy.run_module("comtypes.clear_cache", {}, "__main__")
        mock_input.assert_called_once()
