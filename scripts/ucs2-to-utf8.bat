@echo off

setlocal

echo Recoding 'source' files from UCS-2-LE to UTF-8...

set PATH=%~d0%~p0;%PATH%

cd /d "%~d0%~p0"
cd ..\source
SET sourcePath=%CD%\

:: process forms, macros, queries, reports, and modules

echo %sourcePath%tables
cd %sourcePath%tables >NUL 2>&1
if exist *.data (
   forfiles /m *.data /c "cmd /c iconv -f UCS-2LE -t UTF-8 @file>@fname.txt" >NUL 2>&1
   del *.data
)

echo %sourcePath%forms
cd %sourcePath%forms >NUL 2>&1
if exist *.data (
   forfiles /m *.data /c "cmd /c iconv -f UCS-2LE -t UTF-8 @file>@fname.txt" >NUL 2>&1
   del *.data
)

echo %sourcePath%macros
cd %sourcePath%macros >NUL 2>&1
if exist *.data (
   forfiles /m *.data /c "cmd /c iconv -f UCS-2LE -t UTF-8 @file>@fname.txt" >NUL 2>&1
   del *.data
)

echo %sourcePath%queries
cd %sourcePath%queries >NUL 2>&1
if exist *.data (
   forfiles /m *.data /c "cmd /c iconv -f UCS-2LE -t UTF-8 @file>@fname.txt" >NUL 2>&1
   del *.data
)

echo %sourcePath%reports
cd %sourcePath%reports >NUL 2>&1
if exist *.data (
   forfiles /m *.data /c "cmd /c iconv -f UCS-2LE -t UTF-8 @file>@fname.txt" >NUL 2>&1
   del *.data
)

echo %sourcePath%modules
cd %sourcePath%modules >NUL 2>&1
:: Exported text from Access for 'modules' is not UCS-2; don't convert.
if exist *.data (
   forfiles /m *.data /c "cmd /c move @file @fname.txt"  >NUL 2>&1
)

echo Done.

endlocal

pause
