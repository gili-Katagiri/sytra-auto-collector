# sytra-auto-collector
## Auto collect daily stock data in cooperation with Sytra
### Install and Setup
**This tool only compatible with WindowsOS.**
#### Prepare
First, you might have to install below tools.
- Python
- WinAppDriver
- MS and RSS

---
Install `robotframework-appiumlibrary` from PyPI.
```bash
    pip install robotframework-appiumlibrary
```

---
Fill in the blanks `mksummary.py` and `daily.bat`.
- `mksumary.py`
    - msapp=
    - pafrase=
- `daily.bat`
    - stockpath=
    - wadpath=
    - mspath=
    - rsspath=

---
At last, create `transer.xlsm` with `vba.txt` VBA script. Add macro [Ctrl+i]: 'read\_csv', [Ctrl+q]: 'save\_csv'.
