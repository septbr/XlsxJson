@echo off

call dotnet publish -c release --self-contained=false -r win10-x64 /p:PublishSingleFile=true

if errorlevel 0 (echo successful.) else (echo failed.)
pause > nul