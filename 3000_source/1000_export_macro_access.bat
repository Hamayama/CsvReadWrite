set EXCEL_FILE1=�f�[�^�x�[�X.mdb
set EXPORT_FOLDER1=source_access

pause Access�t�@�C���̃}�N���̃\�[�X���G�N�X�|�[�g���܂��B

cd /d "%~dp0"

1001_ExportExcelMacro.vbs "%EXCEL_FILE1%" "%EXPORT_FOLDER1%"

@echo.
@echo ret = %ERRORLEVEL%

pause �������܂����B
