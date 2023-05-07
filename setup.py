import sys
from cx_Freeze import setup, Executable

# Set the path to your Python script
script_path = 'py.py'

# Set the path to your icon file
icon_path = 'icon.ico'

# Dependencies are automatically detected, but it might need
# fine tuning.
build_exe_options = {
    'packages': [],
    'excludes': [],
    'include_files': [icon_path],
}

# GUI applications require a different base on Windows (the default is for a
# console application).
base = None
if sys.platform == 'win32':
    base = 'Win32GUI'

# Create the executable
executables = [
    Executable(script_path, base=base, icon=icon_path)
]

setup(name='py',
      version='0.1',
      description='Tenster',
      options={'build_exe': build_exe_options},
      executables=executables)
