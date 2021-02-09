from cx_Freeze import setup, Executable

from ezno_convert.common import VERSION

cli = Executable(
    base=None,
    script='ezno_convert\\cli.py',
    target_name='eznoc.exe',
)
gui = Executable(
    base='Win32GUI',
    script='ezno_convert\\gui.py',
    target_name='eznoc-gui.exe',
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
    optimize=1,  # 2 fails because of numpy bug
)

setup(
    name='ezno_convert',
    version=VERSION,
    executables=[cli, gui, context],
    options={'build_exe': build_exe}
)
