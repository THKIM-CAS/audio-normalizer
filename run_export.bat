@echo off
REM Batch file to run the PowerPoint to Video exporter
REM Usage: run_export.bat [input_directory] [output_directory]

echo PowerPoint to Video Batch Exporter
echo ===================================
echo.

if "%~1"=="" (
    REM If no arguments, run with GUI prompts
    cscript //NoLogo export_pptx_to_video.vbs
) else if "%~2"=="" (
    REM If only input directory provided
    cscript //NoLogo export_pptx_to_video.vbs "%~1"
) else (
    REM If both input and output directories provided
    cscript //NoLogo export_pptx_to_video.vbs "%~1" "%~2"
)

pause
