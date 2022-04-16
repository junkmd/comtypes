from pathlib import Path
import shutil
import sys
from typing import Callable, Iterator

import comtypes
import pytest


@pytest.fixture(scope="session")
def gen_dir() -> Path:
	comtypes_dir = Path(comtypes.__file__).parent
	return comtypes_dir / "gen"


@pytest.fixture(scope="session")
def cleanup_gen_import() -> Callable[[], None]:
	def _cleanup():
		gen_mod_names = [k for k in sys.modules if k.startswith("comtypes.gen")]
		for k in gen_mod_names:
			sys.modules.pop(k)
	return _cleanup


@pytest.fixture(autouse=True, scope="module")
def cleanup_gen_dir(gen_dir: Path, cleanup_gen_import: Callable[[], None]) -> Iterator[None]:
	def _cleanup():
		for p in gen_dir.iterdir():
			if p.is_dir():
				shutil.rmtree(p, ignore_errors=True)
			if p.is_file() and p.name != "__init__.py" and p.suffix == ".py":
				p.unlink()
	
	_cleanup()
	cleanup_gen_import()
	yield
	_cleanup()
	cleanup_gen_import()
