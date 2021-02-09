from cx_Freeze import setup, Executable

from ezno_convert.common import VERSION
from ezno_convert.context import gui_exe

cli = Executable(
    base=None,
    script='ezno_convert\\cli.py',
    target_name='eznoc.exe',
)

gui = Executable(
    base='Win32GUI',
    script='ezno_convert\\gui.py',
    target_name=gui_exe.name,
    icon='images\\ezno-icon.png'
)

context = Executable(
    base='Win32GUI',
    script='ezno_convert\\context.py',
)

build_exe = dict(
    include_files=['images/'],
    replace_paths=[('*', '')],  # Obfuscate paths from traceback
    # include_msvcr=True,  # Not sure if needed
    excludes=['tkinter', 'numpy'],
    optimize=2,
    zip_include_packages='*',
    zip_exclude_packages='',
)

setup(
    version=VERSION,
    executables=[cli, gui, context],
    options={'build_exe': build_exe}
)
