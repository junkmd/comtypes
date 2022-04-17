# `comtypes`

## Features
### A lightweight Python COM package
- Based on the `ctypes` FFI library.
- Depends on only standard library (excluding tests and type stubs).
- Allows to define, call, and implement custom and dispatch-based COM interfaces in pure Python.
- Works on Windows and 64-bit Windows.

## Developing
### Requirements in environment
- Python third party packages
    - [`pytest`](https://pypi.org/project/pytest/)
    - [`pytest-cov`](https://pypi.org/project/pytest-cov/)
    - [`pytest-mock`](https://pypi.org/project/pytest-mock/)
    - [`typing-extensions`](https://pypi.org/project/typing-extensions/)
- Others
    - [`Microsoft Access Database Engine 2016 Redistributable`](https://www.microsoft.com/en-US/download/details.aspx?id=54920)
        - For tests using `ADO` and `Access`.

### Testing Command
```python -m pytest --cov -p no:faulthandler comtypes\ -vv```

### CI/CD
WIP
