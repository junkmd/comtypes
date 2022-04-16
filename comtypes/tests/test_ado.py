from pathlib import Path
import shutil
import sys
import tempfile
from typing import Iterator

from comtypes import IUnknown
from comtypes.automation import IDispatch
from comtypes.client.lazybind import Dispatch
from comtypes.client import CreateObject, GetModule

import pytest


PROVIDER = "Microsoft.ACE.OLEDB.12.0"
MDB_NAME = "test.mdb"


@pytest.fixture(scope="module")
def tmp_root_dir(tmp_path_factory) -> Path:
	return tmp_path_factory.mktemp("root")


@pytest.fixture(scope="module", autouse=True)
def _setup_module(cleanup_gen_import, tmp_root_dir):
	catalog = CreateObject("ADOX.Catalog")
	mdb_file = tmp_root_dir / MDB_NAME
	conn = catalog.Create(f"Provider={PROVIDER};Data Source={mdb_file}")
	conn.Close()
	cleanup_gen_import()


@pytest.fixture
def tmp_mdb(tmp_root_dir: Path) -> Iterator[Path]:
	with tempfile.TemporaryDirectory(dir=tmp_root_dir) as t:
		tmp = Path(t)
		dest = tmp / MDB_NAME
		shutil.copyfile(tmp_root_dir / MDB_NAME, dest)
		yield dest


class Test_ADODB_Connection:
	def test_CreateObject_TakesCoClass(self, tmp_mdb):
		from comtypes.gen import ADODB as adodb
		conn = CreateObject(adodb.Connection)
		conn.ConnectionString = f"Provider={PROVIDER};Data Source={tmp_mdb}"
		conn.Open()
		conn.Close()
