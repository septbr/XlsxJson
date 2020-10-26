@echo off

call dotnet publish -c Release -r win-x64 --self-contained=false /p:PublishSingleFile=true

if errorlevel 0 (echo successful.) else (echo failed.)
pause > nul