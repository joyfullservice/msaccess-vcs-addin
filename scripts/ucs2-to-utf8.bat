@echo off

setlocal

echo Recoding 'source' files from UCS-2-LE to UTF-8...

set PATH=%~d0%~p0;%PATH%

cd "%~d0%~p0"
cd ..\source

:: process forms, macros, queries, reports, and modules

echo forms
cd forms
forfiles /m *.data /c "cmd /c iconv -f UCS-2LE -t UTF-8 @file>@fname.txt">NUL
del *.data

echo macros
cd ..\macros
forfiles /m *.data /c "cmd /c iconv -f UCS-2LE -t UTF-8 @file>@fname.txt">NUL
del *.data

echo queries
cd ..\queries
forfiles /m *.data /c "cmd /c iconv -f UCS-2LE -t UTF-8 @file>@fname.txt">NUL
del *.data

echo reports
cd ..\reports
forfiles /m *.data /c "cmd /c iconv -f UCS-2LE -t UTF-8 @file>@fname.txt">NUL
del *.data

echo modules
cd ..\modules
:: Exported text from Access for 'modules' is not UCS-2; don't convert.
forfiles /m *.data /c "cmd /c copy @file @fname.txt">NUL
del *.data

echo Done.

endlocal

pause
