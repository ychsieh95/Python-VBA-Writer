import os
import sys
import winreg
from win32com.client import Dispatch


class VbaMacroWriter():
    def __init__(self, input_xlsm_filepath: str, output_xlsm_filepath: str=None) -> None:
        # comtypes.COINIT_MULTITHREADED
        sys.coinit_flags = 0

        # Variables
        self.input_file   = os.path.abspath(input_xlsm_filepath)
        self.output_file  = os.path.abspath(output_xlsm_filepath) if output_xlsm_filepath else None
        self.reg_path     = r'Software\Microsoft\Office\16.0\Excel\Security'
        self.reg_name     =  'AccessVBOM'
        self.com_instance = Dispatch("Excel.Application") # USING WIN32COM
        self.objworkbook  = None

        # Initialize
        self.com_instance.Visible = False
        self.com_instance.DisplayAlerts = False
        pass


    def check_reg_accessable(self) -> bool:
        return (self.__get_reg(self.reg_name, self.reg_path) == 1)


    def write_reg_accessable(self) -> bool:
        return self.__set_reg(self.reg_name, 1, self.reg_path)


    def open_file(self) -> bool:
        self.objworkbook = self.com_instance.Workbooks.Open(self.input_file)
        return True


    def save_file(self) -> bool:
        if not self.objworkbook:
            return False

        self.objworkbook.SaveAs(self.output_file if self.output_file else self.input_file)
        self.com_instance.Quit()


    def check_file_is_open(self) -> bool:
        return self.objworkbook is not None


    def write_macro_workbook_from_file(self, macro_filepath: str) -> bool:
        if not self.objworkbook:
            return False

        xlmodule = self.objworkbook.VBProject.VBComponents('ThisWorkbook')
        xlmodule.CodeModule.AddFromString(''.join(self.read_file(macro_filepath)).strip())
        return True


    def write_macro_workbook_from_text(self, macro_code: str) -> bool:
        if not self.objworkbook:
            return False

        xlmodule = self.objworkbook.VBProject.VBComponents('ThisWorkbook')
        xlmodule.CodeModule.AddFromString(macro_code.strip())
        return True


    def write_macro_module_from_file(self, macro_filepath: str) -> bool:
        if not self.objworkbook:
            return False

        xlmodule = self.objworkbook.VBProject.VBComponents.Add(1)
        xlmodule.CodeModule.AddFromString(''.join(self.read_file(macro_filepath)).strip())
        return True


    def write_macro_module_from_text(self, macro_code: str) -> bool:
        if not self.objworkbook:
            return False

        xlmodule = self.objworkbook.VBProject.VBComponents.Add(1)
        xlmodule.CodeModule.AddFromString(macro_code.strip())
        return True


    def __get_reg(self, name: str, reg_path: str) -> int:
        try:
            registry_key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, reg_path, 0, winreg.KEY_READ)
            value, regtype = winreg.QueryValueEx(registry_key, name)
            winreg.CloseKey(registry_key)
            return value
        except WindowsError:
            return None


    def __set_reg(self, name: str, value: int, reg_path: str):
        try:
            winreg.CreateKey(winreg.HKEY_CURRENT_USER, reg_path)
            registry_key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, reg_path, 0, winreg.KEY_WRITE)
            winreg.SetValueEx(registry_key, name, 0, winreg.REG_DWORD, value)
            winreg.CloseKey(registry_key)
            return True
        except WindowsError:
            return False


    def read_file(self, filepath: str) -> str:
        content = ''
        with open(filepath, 'r') as f:
            content = f.readlines()
        return ''.join(content)
