from cx_Freeze import setup, Executable
import sys

EXE = 'TestTriagingApp'
filename = EXE+'.pyw'

base = None
if (sys.platform == "win32"):
    base = "Win32GUI"    # Tells the build script to hide the console.


setup(
    name = EXE ,
    version = "0.1" ,
    description = "first release" ,
    executables = [Executable(filename, base=base, icon="M.ico")]
    )							 