REM ********************************************************************************************
REM *                             --== Export.bat file ==--                                    *
REM *                                                                                          *
REM * Extracts a twinBASIC project container file into version-control friendly text documents *
REM * Script Author: Mike Wolfe (mike@nolongerset.com)                                         *
REM * Script Source: https://nolongerset.com/version-control-with-twinbasic/                   *
REM ********************************************************************************************

REM Get the folder name of the latest version of the twinBASIC extension
REM   - VS Code normally uninstalls old versions but not always
REM   - https://stackoverflow.com/a/6362922/154439

FOR /F "tokens=* USEBACKQ" %%F IN (`dir %userprofile%\.vscode\extensions\twinbasic* /B`) DO (
SET tb_ver=%%F
)


REM Build the full path to the twinBASIC executable

SET tb_exe=%userprofile%\.vscode\extensions\%tb_ver%\out\bin\twinBASIC_win32.exe


REM Get the name of the .twinproj file

FOR /F "tokens=* USEBACKQ" %%F IN (`dir %~dp0*.twinproj /B`) DO (
SET twinproj=%%F
)


REM Set the destination folder to a \Source\ subfolder in current folder
REM   - you could override this with a different destination path

SET outfolder=%~dp0Source


REM Export to \Source\ subfolder in current folder
REM   - %~dp0 refers to script folder
REM     o see: https://stackoverflow.com/a/4420078/154439
REM   - REM twinBASIC export command uses the following format:
REM     o twinBASIC_win32.exe export <twinproj_path> <export_folder_path>   --overwrite
REM     o see: https://github.com/WaynePhillipsEA/twinbasic/issues/232#issuecomment-866840797

"%tb_exe%" export "%~dp0%twinproj%" "%outfolder%" --overwrite


pause