'
' 1001_ExportExcelMacro.vbs
' 2025-3-26 v1.05
'
' ���T�v��
'   Excel�t�@�C���̃}�N���̃\�[�X���G�N�X�|�[�g���܂��B
'
' ���g������
'   1001_ExportExcelMacro.vbs excelFile exportPath
'
' �����ӎ�����
'   �E�{�c�[�������s����ꍇ�A
'     Excel�̃I�v�V�����ݒ�ŁA�Z�L�����e�B�Z���^�[(�܂��̓g���X�g�Z���^�[)�̐ݒ���J���A
'     �u�}�N���̐ݒ�v-�uVBA �v���W�F�N�g �I�u�W�F�N�g ���f���ւ̃A�N�Z�X��M������v
'     �Ƀ`�F�b�N������K�v������܂��B
'     (���̐ݒ�����Ȃ��ƁA�A�N�Z�X���̃G���[���������܂�)
'   �E�G�N�X�|�[�g��̃t�H���_�����݂��Ȃ��ꍇ�A�쐬���܂��B
'   �E�{�c�[���̖߂�l�́A���Ăɂ��Ȃ��ł��������B
'     (�G���[���������Ă��A0 (����) ���Ԃ�P�[�X������܂�)
'   �Ev1.05 ����AAccess�t�@�C�� (�g���q�� .mdb �̃t�@�C��) �̃}�N���̃\�[�X��
'     �G�N�X�|�[�g�\�ɂȂ�܂����B
'
' ���Q�lURL��
'   https://gist.github.com/aimoriu/7718005
'   https://taka-2.hatenablog.jp/entry/20090907/p2
'
Option Explicit

Dim objFSO
Dim objShell
Dim strFilePath
Dim strExportPath
Dim intRet

' �ϐ��̏�����
strFilePath = ""
strExportPath = ""
intRet = 1

' �����̎擾
Set objFSO = CreateObject("Scripting.FileSystemObject")
If WScript.Arguments.Count = 2 Then
    strFilePath   = objFSO.GetAbsolutePathName(WScript.Arguments.Item(0))
    strExportPath = objFSO.GetAbsolutePathName(WScript.Arguments.Item(1))
Else
    WScript.Echo "�����̐����s���ł��B"
    WScript.Quit intRet
End If

' �G�N�X�|�[�g��̃t�H���_�����݂��Ȃ���΍쐬����
Set objShell = CreateObject("WScript.Shell")
If Not objFSO.FolderExists(strExportPath) Then
    Call objShell.Run("cmd.exe /c mkdir """ & strExportPath & """", 0, True)
End If

' �t�@�C���̊g���q���`�F�b�N
If LCase(objFSO.GetExtensionName(strFilePath)) = "mdb" Then
    ' Access�t�@�C���̃}�N���̃\�[�X���G�N�X�|�[�g
    intRet = ExportAccessSource(strFilePath, strExportPath)
Else
    ' Excel�t�@�C���̃}�N���̃\�[�X���G�N�X�|�[�g
    intRet = ExportExcelSource(strFilePath, strExportPath)
End If

' �I�u�W�F�N�g�̉��
Set objShell = Nothing
Set objFSO = Nothing

' �I��
WScript.Quit intRet

' == �ȉ��͊֐� ==

' Excel�t�@�C���̃}�N���̃\�[�X���G�N�X�|�[�g
Private Function ExportExcelSource(strFilePath, strExportPath)
    Dim objExcel
    Dim objWorkbook
    Dim objVBProject
    Dim objComponent
    Dim intAutomation
    Dim intRet

    ' �߂�l�̏�����
    intRet = 0

    ' Excel�t�@�C���̃I�[�v��
    Set objExcel = CreateObject("Excel.Application")
    objExcel.Visible = False
    objExcel.DisplayAlerts = False
    objExcel.EnableEvents = False
    objExcel.ScreenUpdating = False
    objExcel.Cursor = 2             ' (=xlWait)
    intAutomation = objExcel.AutomationSecurity
    objExcel.AutomationSecurity = 3 ' (=msoAutomationSecurityForceDisable)
    Set objWorkbook = objExcel.Workbooks.Open(strFilePath)
    Set objVBProject = objWorkbook.VBProject

    ' �}�N���ւ̃A�N�Z�X�����`�F�b�N
    ' (Excel�t�@�C�����J�����ςȂ��ɂȂ邱�Ƃ�h�����߁A���O�Ƀ`�F�b�N����)
    On Error Resume Next
    For Each objComponent In objVBProject.VBComponents
        ' NOP
    Next
    If Err.Number <> 0 Then
        WScript.Echo "VBA �v���W�F�N�g �I�u�W�F�N�g ���f���ւ̃A�N�Z�X��������܂���B" & vbCrLf & _
                     "(Excel �̐ݒ肪�K�v�ł�)"
        intRet = 1
    End If
    On Error Goto 0

    ' �\�[�X�t�@�C���̃G�N�X�|�[�g
    If intRet = 0 Then
        Call ExportSourceFile(objVBProject, strExportPath)
    End If

    ' Excel�t�@�C���̃N���[�Y
    objWorkbook.Close False
    objExcel.AutomationSecurity = intAutomation
    objExcel.Cursor = -4143         ' (=xlDefault)
    objExcel.ScreenUpdating = True
    objExcel.EnableEvents = True
    objExcel.DisplayAlerts = True
    objExcel.Quit

    ' �I�u�W�F�N�g�̉��
    Set objComponent = Nothing
    Set objVBProject = Nothing
    Set objWorkbook = Nothing
    Set objExcel = Nothing

    ' �߂�l��ݒ�
    ExportExcelSource = intRet

End Function

' Access�t�@�C���̃}�N���̃\�[�X���G�N�X�|�[�g
Private Function ExportAccessSource(strFilePath, strExportPath)
    Dim objAccess
    Dim objVBProject
    Dim intRet

    ' �߂�l�̏�����
    intRet = 0

    ' Access�t�@�C���̃I�[�v��
    Set objAccess = CreateObject("Access.Application")
    objAccess.OpenCurrentDatabase(strFilePath)
    Set objVBProject = objAccess.VBE.ActiveVBProject

    ' �\�[�X�t�@�C���̃G�N�X�|�[�g
    Call ExportSourceFile(objVBProject, strExportPath)

    ' Access�t�@�C���̃N���[�Y
    objAccess.Quit

    ' �I�u�W�F�N�g�̉��
    Set objVBProject = Nothing
    Set objAccess = Nothing

    ' �߂�l��ݒ�
    ExportAccessSource = intRet

End Function

' �\�[�X�t�@�C���̃G�N�X�|�[�g
Private Sub ExportSourceFile(objVBProject, strExportPath)
    Dim objComponent
    Dim strExportFilePath

    ' �R���|�[�l���g�̌���
    For Each objComponent In objVBProject.VBComponents
        Select Case objComponent.Type
            ' �W�����W���[��
            Case 1
                strExportFilePath = strExportPath & "\" & objComponent.Name & ".bas"
            ' �N���X���W���[��
            Case 2
                strExportFilePath = strExportPath & "\" & objComponent.Name & ".cls"
            ' Microsoft Form (���[�U�[�t�H�[��)
            Case 3
                strExportFilePath = strExportPath & "\" & objComponent.Name & ".frm"
            ' ActiveX �f�U�C�i
            Case 11
                strExportFilePath = strExportPath & "\" & objComponent.Name & ".cls"
            ' Document ���W���[�� (�V�[�g��ThisWorkBook)
            Case 100
                strExportFilePath = strExportPath & "\" & objComponent.Name & ".cls"
            ' ���̑�
            Case Else
                strExportFilePath = ""
        End Select

        ' (�f�o�b�O�p)
        'If strExportFilePath <> "" Then
        '    WScript.Echo objComponent.Name & " : " & objComponent.CodeModule.CountOfLines
        'End If

        ' �G�N�X�|�[�g (��̃\�[�X�t�@�C���͏��O)
        If strExportFilePath <> "" And objComponent.CodeModule.CountOfLines > 0 Then
            objComponent.Export strExportFilePath
        End If
    Next

    ' �I�u�W�F�N�g�̉��
    Set objComponent = Nothing

End Sub
