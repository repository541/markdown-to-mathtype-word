@echo off
setlocal DisableDelayedExpansion

if "%~1"=="" (
  echo Usage: %~nx0 "input.md" ["output.docx"] [batch-size]
  echo.
  echo The conversion requires Microsoft Word and MathType 7 to be closed.
  exit /b 64
)

set "TOOLROOT=%~dp0"
set "PS=%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe"
set "SOURCE=%~f1"
set "OUTPUT=%~f2"
set "BATCH=%~3"

if "%OUTPUT%"=="" (
  if "%BATCH%"=="" (
    "%PS%" -NoLogo -NoProfile -ExecutionPolicy Bypass -File "%TOOLROOT%Convert-MarkdownToMathTypeWord.ps1" -MarkdownPath "%SOURCE%"
  ) else (
    "%PS%" -NoLogo -NoProfile -ExecutionPolicy Bypass -File "%TOOLROOT%Convert-MarkdownToMathTypeWord.ps1" -MarkdownPath "%SOURCE%" -BatchSize %BATCH%
  )
) else (
  if "%BATCH%"=="" (
    "%PS%" -NoLogo -NoProfile -ExecutionPolicy Bypass -File "%TOOLROOT%Convert-MarkdownToMathTypeWord.ps1" -MarkdownPath "%SOURCE%" -OutputDocx "%OUTPUT%"
  ) else (
    "%PS%" -NoLogo -NoProfile -ExecutionPolicy Bypass -File "%TOOLROOT%Convert-MarkdownToMathTypeWord.ps1" -MarkdownPath "%SOURCE%" -OutputDocx "%OUTPUT%" -BatchSize %BATCH%
  )
)

set "RESULT=%ERRORLEVEL%"
if not "%RESULT%"=="0" echo Conversion failed. Read the message above; existing source files and validated checkpoints were preserved.
endlocal & exit /b %RESULT%
