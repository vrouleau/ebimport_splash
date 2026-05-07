@echo off
REM Audit results PDF against source xlsx
REM Usage: audit.bat resultats.pdf input.xlsx
REM        audit.bat resultats.pdf  (uses test_attendees.xlsx)

setlocal
set PDF=%1
set XLSX=%2

if "%PDF%"=="" (
    echo Usage: audit.bat resultats.pdf [input.xlsx]
    exit /b 1
)
if "%XLSX%"=="" set XLSX=..\..\test_attendees.xlsx

curl -sS -X POST http://localhost:5000/api/audit -F "pdf=@%PDF%" -F "xlsx=@%XLSX%" -o "%TEMP%\audit_result.json"
python "%~dp0format_audit.py" "%TEMP%\audit_result.json"
