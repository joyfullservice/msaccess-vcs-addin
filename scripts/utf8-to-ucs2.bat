@echo off

setlocal

echo Recoding 'source' files from UTF-8 to UCS-2-LE...

set PATH=%~d0%~p0;%PATH%

cd "%~d0%~p0"
cd ..\source

:: process forms, macros, queries, reports, and modules

echo forms
cd forms
del *.data >NUL 2>NUL
forfiles /m *.txt /c "cmd /c iconv -f UTF-8 -t UCS-2LE @file>@fname.data">NUL

echo macros
cd ..\macros
del *.data >NUL 2>NUL
forfiles /m *.txt /c "cmd /c iconv -f UTF-8 -t UCS-2LE @file>@fname.data">NUL

echo queries
cd ..\queries
del *.data >NUL 2>NUL
forfiles /m *.txt /c "cmd /c iconv -f UTF-8 -t UCS-2LE @file>@fname.data">NUL

echo reports
cd ..\reports
del *.data >NUL 2>NUL
forfiles /m *.txt /c "cmd /c iconv -f UTF-8 -t UCS-2LE @file>@fname.data">NUL

echo modules
cd ..\modules
:: Exported text from Access for 'modules' is not UCS-2; don't convert.
del *.data >NUL 2>NUL
forfiles /m *.txt /c "cmd /c copy @file @fname.data">NUL

echo Done.

endlocal

pause
