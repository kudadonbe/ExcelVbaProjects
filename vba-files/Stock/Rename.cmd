@REM rename all files in folder 
@REM by taking the current file name
@REM and check if it contains request number
@REM take the request number and change it to "GRF-2023-" + number in "000" format
@REM 'example: 116_Goods Requisition.xlsx -> GRF-2023-116.xlsx
@REM rename all files in folder according to the new name
@REM '===============================================================
@REM @echo off
setlocal enabledelayedexpansion
set "folder=%~1"
set "request=%~2"
set "prefix=GRF-2023-"
set "suffix=.xlsx"
set "count=0"
for /f "delims=" %%i in ('dir /b /a-d "%folder%\*%request%*.xlsx"') do (
    set "file=%%~i"
    set "file=!file:%request%=!"
    set "file=!file:~0,-5!"
    set "file=!prefix!!file!!suffix!"
    ren "%folder%\%%~i" "!file!"
    set /a count+=1
)
echo %count% files renamed
pause
```


