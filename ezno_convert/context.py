import ctypes
import sys
import winreg

from ezno_convert.common import here

reg_classes = (
    'Word.Document.8',
    'Excel.Sheet.8',
    'PowerPoint.Show.8',
    'Word.Document.12',
    'Excel.Sheet.12',
    'PowerPoint.Show.12',
    'Directory',
    # 'Directory\\Background',
)

gui_exe = here.with_name('eznoc-gui.exe').absolute()
gui_with_options_cmd = str(gui_exe) + ' -p %1'
gui_no_options_cmd = str(gui_exe) + ' --pdf -p %1'


def is_admin():
    result = False
    try:
        result = ctypes.windll.shell32.IsUserAnAdmin()
    finally:
        return result


def install():
    for key in reg_classes:
        root = winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, sub_key=rf'{key}\shell', access=winreg.KEY_CREATE_SUB_KEY)
        command_pdf = winreg.CreateKeyEx(root, r'EZNO Convert to PDF\command', access=winreg.KEY_SET_VALUE)
        winreg.SetValueEx(command_pdf, '', 0, winreg.REG_SZ, gui_no_options_cmd)
        command_gui = winreg.CreateKeyEx(root, r'EZNO Convert to...\command', access=winreg.KEY_SET_VALUE)
        winreg.SetValueEx(command_gui, '', 0, winreg.REG_SZ, gui_with_options_cmd)


def uninstall():
    for key in reg_classes:
        root = winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, sub_key=rf'{key}\shell')
        winreg.DeleteKeyEx(root, r'EZNO Convert to PDF\command')
        winreg.DeleteKeyEx(root, r'EZNO Convert to...\command')
        winreg.DeleteKeyEx(root, r'EZNO Convert to PDF')
        winreg.DeleteKeyEx(root, r'EZNO Convert to...')


if __name__ == '__main__':
    if is_admin():
        if sys.argv[-1] == 'install':
            install()
        elif sys.argv[-1] == 'uninstall':
            uninstall()
    else:
        args = sys.argv[1:] if getattr(sys, 'frozen', False) else sys.argv
        # Re-run the program with admin rights
        ctypes.windll.shell32.ShellExecuteW(None, "runas", sys.executable, ' '.join(args), None, 1)
