@echo off

setlocal

echo Recoding 'source' files from UTF-8 to UCS-2-LE...

set PATH=%~d0%~p0;%PATH%

cd /d "%~d0%~p0"
cd ..\source
SET sourcePath=%CD%\

:: process forms, macros, queries, reports, and modules

echo %sourcePath%tables
cd %sourcePath%tables >NUL 2>&1
del *.data >NUL 2>&1
if exist *.txt (
   forfiles /m *.txt /c "cmd /c iconv -f UTF-8 -t UCS-2LE @file>@fname.data" >NUL 2>&1
)

echo %sourcePath%forms
cd %sourcePath%forms >NUL 2>&1
del *.data >NUL 2>&1
if exist *.txt (
   forfiles /m *.txt /c "cmd /c iconv -f UTF-8 -t UCS-2LE @file>@fname.data" >NUL 2>&1
)

echo %sourcePath%macros
cd %sourcePath%macros >NUL 2>&1
del *.data >NUL 2>&1
if exist *.txt (
   forfiles /m *.txt /c "cmd /c iconv -f UTF-8 -t UCS-2LE @file>@fname.data" >NUL 2>&1
)

echo %sourcePath%queries
cd %sourcePath%queries >NUL 2>&1
del *.data >NUL 2>&1
if exist *.txt (
   forfiles /m *.txt /c "cmd /c iconv -f UTF-8 -t UCS-2LE @file>@fname.data" >NUL 2>&1
)

echo %sourcePath%reports
cd %sourcePath%reports >NUL 2>&1
del *.data >NUL 2>&1
if exist *.txt (
   forfiles /m *.txt /c "cmd /c iconv -f UTF-8 -t UCS-2LE @file>@fname.data" >NUL 2>&1
)

echo %sourcePath%modules
cd %sourcePath%modules >NUL 2>&1
:: Exported text from Access for 'modules' is not UCS-2; don't convert.
del *.data >NUL 2>&1
if exist *.txt (
   forfiles /m *.txt /c "cmd /c copy @file @fname.data" >NUL 2>&1
)

echo Done.

endlocal

pause
