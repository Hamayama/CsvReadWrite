set EXCEL_FILE1=..\CsvReadWrite.xlsm
set EXPORT_FOLDER1=source_101

pause Excelファイルのマクロのソースをエクスポートします。

cd /d "%~dp0"

1001_ExportExcelMacro.vbs "%EXCEL_FILE1%" "%EXPORT_FOLDER1%"

@echo.
@echo ret = %ERRORLEVEL%

pause 完了しました。
