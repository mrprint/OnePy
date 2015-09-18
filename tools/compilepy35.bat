rem set "PYTHONHOME=%PROGRAMFILES(x86)%\IronPython 2.7"
rem set "PYTHONPATH=%PYTHONHOME%Lib:%PYTHONHOME%Dll"
..\..\CPython\python.exe ..\..\tools\source_prepare.py ..\..\connector.py ..\..\tools\11.py %2 %3 %4 %5
..\..\CPython\python.exe ..\..\tools\source_prepare.py ..\..\interfacing.py ..\..\tools\12.py %2 %3 %4 %5
..\..\CPython\python.exe ..\..\tools\pyobfuscate.py -r ..\..\tools\11.py > ..\..\tools\21.py
..\..\CPython\python.exe ..\..\tools\pyobfuscate.py -r ..\..\tools\12.py > ..\..\tools\22.py
..\..\CPython\python.exe ..\..\tools\source_prepare.py ..\..\tools\21.py ..\..\#1
..\..\CPython\python.exe ..\..\tools\source_prepare.py ..\..\tools\22.py ..\..\#2
rem del ..\..\tools\1.py ..\..\tools\2.py
copy /Y ..\..\App35.config ..\..\App.config
