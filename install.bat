@echo off
rem INSTALL.BAT - Easy installer for Python modules on Windows

rem version 0.04 2014-02-24 Philippe Lagadec - http://www.decalage.info

rem License: 
rem This file install.bat can freely used, modified and redistributed, as 
rem long as credit to the author is kept intact. Please send any feedback,
rem issues or improvements to decalage at laposte.net.

rem CHANGELOG:
rem 2007-09-04 v0.01 PL: - first version, for Python 2.3 to 2.5
rem 2009-02-27 v0.02 PL: - added support for Python 2.6
rem 2013-05-07 v0.03 PL: - added support for Python 2.7
rem 2014-02-24 v0.04 PL: - added support for py.exe

rem 1) test if py.exe or python.exe is in the path:
rem (py.exe is better because it can select python 2 or 3 according to shebang lines)

py.exe --version >NUL 2>&1
if errorlevel 1 goto notpy
echo py.exe found in the path.
py.exe setup.py install
if errorlevel 1 goto error
goto end
:NOTPY

python.exe --version >NUL 2>&1
if errorlevel 1 goto notpath
echo Python.exe found in the path.
python setup.py install
if errorlevel 1 goto error
goto end
:NOTPATH

rem 2) test for usual python.exe paths:

REM Python 2.7: 
c:\python27\python.exe --version >NUL 2>&1
if errorlevel 1 goto notpy27
echo Python.exe found in C:\Python27
c:\python27\python.exe setup.py install
if errorlevel 1 goto error
goto end 
:NOTPY27

REM Python 2.6: 
c:\python26\python.exe --version >NUL 2>&1
if errorlevel 1 goto notpy26
echo Python.exe found in C:\Python26
c:\python26\python.exe setup.py install
if errorlevel 1 goto error
goto end 
:NOTPY26

c:\python25\python.exe --version >NUL 2>&1
if errorlevel 1 goto notpy25
echo Python.exe found in C:\Python25
c:\python25\python.exe setup.py install
if errorlevel 1 goto error
goto end 
:NOTPY25

c:\python24\python.exe --version >NUL 2>&1
if errorlevel 1 goto notpy24
echo Python.exe found in C:\Python24
c:\python24\python.exe setup.py install
if errorlevel 1 goto error
goto end 
:NOTPY24

c:\python23\python.exe --version >NUL 2>&1
if errorlevel 1 goto notpy23
echo Python.exe found in C:\Python23
c:\python23\python.exe setup.py install
if errorlevel 1 goto error
goto end 
:NOTPY23

"c:\program files\python\python.exe" --version >NUL 2>&1
if errorlevel 1 goto notpf
echo Python.exe found in C:\Program Files\Python
"c:\program files\python\python.exe" setup.py install
if errorlevel 1 goto error
goto end 
:NOTPF

rem 3) last we just try to launch the script, if .py is associated to python.exe
echo Python.exe not found, trying to launch setup.py directly.
setup.py install
if errorlevel 1 goto error
goto end

:ERROR
echo.
echo If the installation is not successful, try to run "python setup.py install"
echo or simply "setup.py install" in the script directory.
echo You can also copy files by hand in the site-package directory of your
echo Python directory.
REM pause

:END
pause
