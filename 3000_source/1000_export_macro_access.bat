set EXCEL_FILE1=データベース.mdb
set EXPORT_FOLDER1=source_access

pause Accessファイルのマクロのソースをエクスポートします。

cd /d "%~dp0"

1001_ExportExcelMacro.vbs "%EXCEL_FILE1%" "%EXPORT_FOLDER1%"

@echo.
@echo ret = %ERRORLEVEL%

pause 完了しました。
