# VBA Macro Writer

A basic VBA macro writer with Python.

## Requirements

* Python 3+
* Python package: pywin32

  ```bash
  $ pip3 install pywin32
  ```

## Usage

```python
from vba_macro_writer import VbaMacroWriter


INPUT_FILE   = 'The ABSOLITE file path of input xlsm file'
STARTUP_MACRO = 'The VBA startup macro'
MODULE_MACRO  = 'The VBA module macro'


if __name__ == '__main__':
    m_writer = VbaWriter(INPUT_FILE)
    if not m_writer.check_reg_accessable():
        m_writer.write_reg_accessable()
    m_writer.open_file()
    m_writer.write_macro_workbook_from_text(STARTUP_MACRO)
    m_writer.write_macro_module_from_text(MODULE_MACRO)
    m_writer.save_file()
```

## References

* [Python Inject VBA Macro into XLSM File](https://blog.holey.cc/2022/09/28/python-inject-vba-into-xlsm-file/)
