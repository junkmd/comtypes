from decimal import Decimal
from pathlib import Path
import shutil
import tempfile
from typing import Iterator

from comtypes.client import CreateObject, GetModule

import pytest


PROVIDER = "Microsoft.ACE.OLEDB.12.0"
MDB_NAME = "test.mdb"


@pytest.fixture(scope="module")
def tmp_root_dir(tmp_path_factory) -> Path:
	return tmp_path_factory.mktemp("root")


@pytest.fixture(scope="module", autouse=True)
def _setup_module(cleanup_gen_import, tmp_root_dir) -> None:
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
	def test_CreateObject(self, tmp_mdb):
		from comtypes.gen import ADODB as adodb
		conn = CreateObject(adodb.Connection)
		conn.ConnectionString = f"Provider={PROVIDER};Data Source={tmp_mdb}"
		conn.Open()
		fld_cmd = "Id IDENTITY(1,1) PRIMARY KEY, Name TEXT(9), Price CURRENCY"
		conn.Execute(f"CREATE TABLE Merc({fld_cmd})")
		ins_into_cmd = f"INSERT INTO Merc(Name, price)"
		conn.Execute(f"{ins_into_cmd} VALUES('spam', 50)")
		conn.Execute(f"{ins_into_cmd} VALUES('ham', 30)")
		_, rs = conn.Execute("SELECT Id, Name, Price FROM Merc")
		records = []
		while not (rs.BOF or rs.EOF):
			records.append({f.Name: f.Value for f in rs.Fields})
			rs.MoveNext()
		D = Decimal
		assert records == [
			{"Id": 1, "Name": "spam", "Price": D("50")},
			{"Id": 2, "Name": "ham", "Price": D("30")},
		]
		conn.Close()
