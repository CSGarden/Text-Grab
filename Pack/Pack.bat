@echo off
cd /d %~dp0 
set INNO_SETUP_EXE6="Inno Setup 6\ISCC.exe"
set OUTPUT_SETUP_TEMP_DIR=PackedFiles

set ProjectFile=..\Text-Grab\Text-Grab.csproj
set PROGRAM_FILES_DIR=ProgramFiles


rd /s /q "%PROGRAM_FILES_DIR%"
md "%PROGRAM_FILES_DIR%"

if not exist %OUTPUT_SETUP_TEMP_DIR%\ md %OUTPUT_SETUP_TEMP_DIR%

dotnet publish %ProjectFile%  --runtime win-x64  --self-contained true -c Release -v minimal -o %PROGRAM_FILES_DIR% -p:PublishReadyToRun=true -p:PublishSingleFile=true -p:CopyOutputSymbolsToPublishDirectory=false -p:Version=1.0 --nologo

echo Packaging...
%INNO_SETUP_EXE6% "InnoSetup.iss"
echo Finished!
pause