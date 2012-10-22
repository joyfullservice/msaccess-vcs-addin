@echo off

setlocal

echo Recoding 'source' files from UCS-2-LE to UTF-8...

set PATH=%~d0%~p0;%PATH%

cd /d "%~d0%~p0"
cd ..\source

:: process forms, macros, queries, reports, and modules

echo tables
cd tables
if exist *.data (
   forfiles /m *.data /c "cmd /c iconv -f UCS-2LE -t UTF-8 @file>@fname.txt">NUL
   del *.data
)

echo forms
cd ..\forms
if exist *.data (
   forfiles /m *.data /c "cmd /c iconv -f UCS-2LE -t UTF-8 @file>@fname.txt">NUL
   del *.data
)

echo macros
cd ..\macros
if exist *.data (
   forfiles /m *.data /c "cmd /c iconv -f UCS-2LE -t UTF-8 @file>@fname.txt">NUL
   del *.data
)

echo queries
cd ..\queries
if exist *.data (
   forfiles /m *.data /c "cmd /c iconv -f UCS-2LE -t UTF-8 @file>@fname.txt">NUL
   del *.data
)

echo reports
cd ..\reports
if exist *.data (
   forfiles /m *.data /c "cmd /c iconv -f UCS-2LE -t UTF-8 @file>@fname.txt">NUL
   del *.data
)

echo modules
cd ..\modules
:: Exported text from Access for 'modules' is not UCS-2; don't convert.
if exist *.data (
   forfiles /m *.data /c "cmd /c copy @file @fname.txt">NUL
   del *.data
)

echo Done.

endlocal

pause
