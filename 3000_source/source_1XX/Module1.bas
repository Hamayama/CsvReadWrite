Attribute VB_Name = "Module1"
Option Explicit

' == CSV�t�@�C���̓ǂݏ��� ==

' �t�@�C�����̒萔
Const DEFAULT_CSV_FILE_NAME_1 As String = "SampleData.csv"  ' (�ǂݍ��ݗp)
Const DEFAULT_CSV_FILE_NAME_2 As String = "SampleData1.csv" ' (�����o���p)

' �����R�[�h�̒萔
Const CHAR_CODE_1 As String = "utf-8"       ' (�ǂݍ��ݗp)
'Const CHAR_CODE_1 As String = "shift_jis"   ' (�ǂݍ��ݗp)
Const CHAR_CODE_2 As String = "utf-8"       ' (�����o���p)
'Const CHAR_CODE_2 As String = "shift_jis"   ' (�����o���p)
Const UTF_8_BOM_1 As Boolean = False        ' (�ǂݍ��ݗp)(���g�p)
Const UTF_8_BOM_2 As Boolean = False        ' (�����o���p)

' �G���[�̒萔
Const READ_ERROR As Long = 1                ' (�ǂݍ��ݗp)
Const WRITE_ERROR As Long = 2               ' (�����o���p)

' �s�̒萔
Const ROW_HEADER As Long = 6
Const ROW_DATA_BEGIN As Long = 7

' ��̒萔
Const COL_NO As Long = 1
Const COL_SKIP As Long = 2
Const COL_DATA_BEGIN As Long = 3
Const COL_DATA_CHECK As Long = 3

' �G���[���b�Z�[�W�\��
Private Sub DispErrorMessage(msg As String, lineNo As Long, errType As Long)
    Dim msg1 As String

    ' �G���[�̃^�C�v�ŏꍇ����
    msg1 = ""
    Select Case errType
        Case READ_ERROR
            If lineNo > 0 Then
                msg1 = "(" & lineNo & "�s�ڂ̓ǂݍ��݂ŃG���[���������܂���)"
            End If
        Case WRITE_ERROR
            If lineNo > 0 Then
                msg1 = "(" & lineNo & "�s�ڂ̏����o���ŃG���[���������܂���)"
            End If
    End Select

    ' �G���[���b�Z�[�W�\��
    Call MsgBox(msg & vbCrLf & msg1, vbOKOnly + vbExclamation)

End Sub

' �_�u���N�H�[�g�폜
Private Function RemoveDoubleQuote(str As String) As String
    Dim ret As String
    Dim ch As String
    Dim quoteFlag As Boolean
    Dim i As Long

    ' �߂�l�̏�����
    ret = ""

    ' �󕶎���Ȃ甲����
    If str = "" Then
        RemoveDoubleQuote = ret
        Exit Function
    End If

    ' 1���������ׂă_�u���N�H�[�g���폜����
    quoteFlag = False
    For i = 1 To Len(str)
        ch = Mid(str, i, 1)
        ' �_�u���N�H�[�g�̂Ƃ�
        If ch = """" Then
            ' �_�u���N�H�[�g��2�A�������Ƃ�
            If quoteFlag = True Then
                ' �_�u���N�H�[�g��1�����ɂ���
                ret = ret & ch
            End If
            quoteFlag = Not quoteFlag
        Else
            ret = ret & ch
            quoteFlag = False
        End If
    Next

    ' �߂�l��Ԃ�
    RemoveDoubleQuote = ret

End Function

' �_�u���N�H�[�g�ň͂�
' (special�ɕ�������w�肷��ƁA�����̕���������Ƃ������_�u���N�H�[�g�ň͂�)
Private Function AddDoubleQuote(str As String, Optional special As String = "") As String
    Dim ret As String
    Dim convFlag As Boolean
    Dim i As Long

    ' �߂�l�̏�����
    ret = ""

    ' �󕶎���Ȃ甲����
    If str = "" Then
        AddDoubleQuote = ret
        Exit Function
    End If

    ' ���ꕶ���`�F�b�N
    convFlag = False
    If special <> "" Then
        For i = 1 To Len(special)
            If InStr(str, Mid(special, i, 1)) > 0 Then
                convFlag = True
                Exit For
            End If
        Next
    Else
        convFlag = True
    End If

    ' �ϊ�����̂Ƃ�
    If convFlag = True Then
        ' �_�u���N�H�[�g1���A�_�u���N�H�[�g2�ɒu��
        ret = Replace(str, """", """""")
        ' �_�u���N�H�[�g�ň͂�
        ret = """" & ret & """"
    Else
        ret = str
    End If

    ' �߂�l��Ԃ�
    AddDoubleQuote = ret

End Function

' �_�u���N�H�[�g�Ή���Split
' (��؂蕶��(delim)��1�����̂ݑΉ�)
Private Function SplitWithDoubleQuote(str As String, delim As String) As String()
    Dim ret() As String
    Dim retCount As Long
    Dim ch As String
    Dim quoteFlag As Boolean
    Dim i As Long
    Dim i1 As Long

    ' 1���������ׂĕ�������
    retCount = 0
    quoteFlag = False
    i1 = 1
    For i = 1 To Len(str)
        ch = Mid(str, i, 1)
        ' �_�u���N�H�[�g�̂Ƃ�
        If ch = """" Then
            quoteFlag = Not quoteFlag
        ' ��؂蕶���̂Ƃ�
        ElseIf ch = delim Then
            If quoteFlag = False Then
                ReDim Preserve ret(retCount)
                ret(retCount) = Mid(str, i1, i - i1)
                retCount = retCount + 1
                i1 = i + 1
            End If
        End If
    Next
    ReDim Preserve ret(retCount)
    If i1 <= Len(str) Then
        ret(retCount) = Mid(str, i1)
    Else
        ret(retCount) = ""
    End If
    SplitWithDoubleQuote = ret

End Function

' CSV�t�@�C���̓ǂݍ���
Sub ReadCSVFile()
    Dim defaultFileName As String
    'Dim fnameVariant As Variant
    Dim fname As String
    Dim fd As FileDialog
    'Dim wso As Object
    'Dim curDirOld As String
    Dim delim As String

    ' �t�@�C�����̎擾
    ' (�����t�H���_�̓��[�N�u�b�N�̂���t�H���_�Ƃ���)
    ' (GetOpenFilename() �ł͏����t�@�C�������w��ł��Ȃ��������߁A
    '  FileDialog(msoFileDialogOpen) ���g�p����悤�ɂ���)
    '
    'Set wso = CreateObject("WScript.Shell")
    'curDirOld = wso.CurrentDirectory
    'wso.CurrentDirectory = ThisWorkbook.Path
    ''defaultFileName = DEFAULT_CSV_FILE_NAME_1
    'fnameVariant = Application.GetOpenFilename("CSV�t�@�C�� (*.csv),*.csv,TSV�t�@�C�� (*.tsv),*.tsv,���ׂẴt�@�C�� (*.*),*.*")
    'wso.CurrentDirectory = curDirOld
    'If fnameVariant = False Then
    '    Exit Sub
    'End If
    'fname = CStr(fnameVariant)
    '
    Set fd = Application.FileDialog(msoFileDialogOpen)
    defaultFileName = DEFAULT_CSV_FILE_NAME_1
    fd.InitialFileName = ThisWorkbook.Path & "\" & defaultFileName
    Call fd.Filters.Clear
    Call fd.Filters.Add("CSV�t�@�C��", "*.csv")
    Call fd.Filters.Add("TSV�t�@�C��", "*.tsv")
    Call fd.Filters.Add("���ׂẴt�@�C��", "*.*")
    'fd.FilterIndex = fd.Filters.count
    fd.FilterIndex = 0
    If fd.Show = False Then
        Exit Sub
    End If
    fname = fd.SelectedItems(1)

    ' ��؂蕶����ݒ�
    If LCase(Right(fname, 4)) = ".tsv" Then
        delim = vbTab
    Else
        delim = ","
    End If
    
    ' CSV�t�@�C���̓ǂݍ���(�T�u)
    Call ReadCSVFileSub(fname, ActiveSheet, delim)

    ' �I�u�W�F�N�g�̉��
    'Set wso = Nothing
    Set fd = Nothing

End Sub

' CSV�t�@�C���̓ǂݍ���(�T�u)
Private Sub ReadCSVFileSub(fname As String, sheet As Worksheet, delim As String)
    Dim stream As Object
    Dim lineNo As Long
    Dim row As Long
    Dim col As Long
    Dim line As String
    Dim header() As String
    Dim headerNum As Long
    Dim data() As String
    Dim dataNum As Long
    Dim i As Long

    ' �t�@�C���̃I�[�v��
    Set stream = CreateObject("ADODB.Stream")
    stream.Charset = CHAR_CODE_1
    Call stream.Open

    ' �t�@�C���̓ǂݍ���
    Call stream.LoadFromFile(fname)

    ' �擪�s��ǂݍ���
    lineNo = 1
    line = stream.ReadText(-2)
    lineNo = lineNo + 1

    ' �w�b�_���
    header = SplitWithDoubleQuote(line, delim)
    headerNum = UBound(header)

    ' �w�b�_��\��
    row = ROW_HEADER
    For i = 0 To headerNum
        col = COL_DATA_BEGIN + i
        sheet.Cells(row, col) = RemoveDoubleQuote(header(i))
    Next

    ' �f�[�^�̖{�̂�ǂݍ���
    row = ROW_DATA_BEGIN
    Do Until stream.EOS

        ' 1�s�ǂݍ���
        line = stream.ReadText(-2)

        ' �f�[�^���
        data = SplitWithDoubleQuote(line, delim)
        dataNum = UBound(data)
        If dataNum < headerNum Then
            Call DispErrorMessage("1�s�̃f�[�^�����Ȃ����܂��B", lineNo, READ_ERROR)
            GoTo Label_Exit
        ElseIf dataNum > headerNum Then
            Call DispErrorMessage("1�s�̃f�[�^���������܂��B", lineNo, READ_ERROR)
            GoTo Label_Exit
        End If

        ' ��ʍX�V���~
        Application.ScreenUpdating = False
        Application.Cursor = xlWait

        ' No. ��\��
        sheet.Cells(row, COL_NO) = CStr(lineNo - 1)

        ' �f�[�^��\��
        For i = 0 To headerNum
            col = COL_DATA_BEGIN + i
            If i <= dataNum Then
                sheet.Cells(row, col) = RemoveDoubleQuote(data(i))
            Else
                sheet.Cells(row, col) = ""
            End If
        Next

        ' ��ʍX�V���ĊJ
        Application.Cursor = xlDefault
        Application.ScreenUpdating = True

        ' ��ʂ��X�V
        If (lineNo Mod 10) = 0 Then
            DoEvents
        End If

        ' ���̍s��
        lineNo = lineNo + 1
        row = row + 1
    Loop

Label_Exit:

    ' �t�@�C���̃N���[�Y
    If Not stream Is Nothing Then
        Call stream.Close
    End If

    ' �I�u�W�F�N�g�̉��
    Set stream = Nothing

End Sub

' CSV�t�@�C���̏����o��
Sub WriteCSVFile()
    Dim defaultFileName As String
    Dim fnameVariant As Variant
    Dim fname As String
    'Dim fd As FileDialog
    Dim wso As Object
    Dim curDirOld As String
    Dim delim As String

    ' �t�@�C�����̎擾
    ' (�����t�H���_�̓��[�N�u�b�N�̂���t�H���_�Ƃ���)
    ' (FileDialog(msoFileDialogSaveAs) �ł͊g���q���w��ł��Ȃ��������߁A
    '  GetSaveAsFilename() ���g�p����悤�ɂ���)
    '
    Set wso = CreateObject("WScript.Shell")
    curDirOld = wso.CurrentDirectory
    wso.CurrentDirectory = ThisWorkbook.Path
    defaultFileName = DEFAULT_CSV_FILE_NAME_2
    fnameVariant = Application.GetSaveAsFilename(defaultFileName, "CSV�t�@�C�� (*.csv),*.csv,TSV�t�@�C�� (*.tsv),*.tsv,���ׂẴt�@�C�� (*.*),*.*")
    wso.CurrentDirectory = curDirOld
    If fnameVariant = False Then
        Exit Sub
    End If
    fname = CStr(fnameVariant)
    '
    'Set fd = Application.FileDialog(msoFileDialogSaveAs)
    'defaultFileName = DEFAULT_CSV_FILE_NAME_2
    'fd.InitialFileName = ThisWorkbook.Path & "\" & defaultFileName
    ''Call fd.Filters.Clear
    ''Call fd.Filters.Add("CSV�t�@�C��", "*.csv")
    ''Call fd.Filters.Add("TSV�t�@�C��", "*.tsv")
    ''Call fd.Filters.Add("���ׂẴt�@�C��", "*.*")
    ''fd.FilterIndex = fd.Filters.Count
    'fd.FilterIndex = 0
    'If fd.Show = False Then
    '    Exit Sub
    'End If
    'fname = fd.SelectedItems(1)

    ' ��؂蕶����ݒ�
    If LCase(Right(fname, 4)) = ".tsv" Then
        delim = vbTab
    Else
        delim = ","
    End If

    ' CSV�t�@�C���̏����o��(�T�u)
    Call WriteCSVFileSub(fname, ActiveSheet, delim)

    ' �I�u�W�F�N�g�̉��
    Set wso = Nothing
    'Set fd = Nothing

End Sub

' CSV�t�@�C���̏����o��(�T�u)
Private Sub WriteCSVFileSub(fname As String, sheet As Worksheet, delim As String)
    Dim stream As Object
    Dim streamNoBom As Object
    Dim lineNo As Long
    Dim row As Long
    Dim col As Long
    Dim colLast As Long
    Dim header As String
    Dim data As String
    Dim line As String

    ' �ŏI��̗�ԍ����擾
    col = COL_DATA_BEGIN
    colLast = 0
    Do Until Trim(sheet.Cells(ROW_HEADER, col)) = ""
        colLast = col
        col = col + 1
    Loop

    ' �t�@�C���̃I�[�v��
    Set stream = CreateObject("ADODB.Stream")
    stream.Charset = CHAR_CODE_2
    Call stream.Open

    ' ���[�N�V�[�g����w�b�_�s��ǂݍ���
    line = ""
    For col = COL_DATA_BEGIN To colLast
        If col > COL_DATA_BEGIN Then
            line = line & delim
        End If
        header = AddDoubleQuote(sheet.Cells(ROW_HEADER, col), """" & delim & vbLf)
        line = line & header
    Next

    ' 1�s��������
    Call stream.WriteText(line, 1)
    lineNo = lineNo + 1

    ' �f�[�^�o��
    lineNo = 1
    row = ROW_DATA_BEGIN
    Do Until Trim(sheet.Cells(row, COL_NO)) = ""

        ' �o�̓X�L�b�v���f�[�^���Ȃ���΁A���̍s��
        If Trim(sheet.Cells(row, COL_SKIP)) <> "" Or Trim(sheet.Cells(row, COL_DATA_CHECK)) = "" Then
            GoTo Label_Next_1:
        End If

        ' ���[�N�V�[�g����1�s�ǂݍ���
        line = ""
        For col = COL_DATA_BEGIN To colLast
            If col > COL_DATA_BEGIN Then
                line = line & delim
            End If
            data = AddDoubleQuote(sheet.Cells(row, col), """" & delim & vbLf)
            line = line & data
        Next

        ' 1�s��������
        Call stream.WriteText(line, 1)
        lineNo = lineNo + 1

Label_Next_1:

        ' ���̍s��
        row = row + 1
    Loop

    ' �t�@�C���̕ۑ�
    ' UTF-8 �� BOM �Ȃ��̂Ƃ�
    If LCase(CHAR_CODE_2) = "utf-8" And UTF_8_BOM_2 = False Then
        ' BOM �Ȃ��ɕϊ�����
        Set streamNoBom = CreateObject("ADODB.Stream")
        streamNoBom.Type = 1                    ' �o�C�i�����[�h
        Call streamNoBom.Open
        stream.Position = 3                     ' BOM�̃T�C�Y
        stream.CopyTo streamNoBom
        Call streamNoBom.SaveToFile(fname, 2)   ' �㏑��
    Else
        Call stream.SaveToFile(fname, 2)        ' �㏑��
    End If

Label_Exit:

    ' �t�@�C���̃N���[�Y
    If Not streamNoBom Is Nothing Then
        Call streamNoBom.Close
    End If
    If Not stream Is Nothing Then
        Call stream.Close
    End If

    ' �I�u�W�F�N�g�̉��
    Set streamNoBom = Nothing
    Set stream = Nothing

End Sub

' �V�[�g�̃f�[�^���N���A
Sub ClearSheetData()
    Dim ret As Long

    ' �m�F���b�Z�[�W�̕\��
    ret = MsgBox("�V�[�g�̃f�[�^��S�ăN���A���܂��B��낵���ł���?", _
                 vbYesNoCancel + vbQuestion + vbDefaultButton3)
    If ret <> vbYes Then
        ' �L�����Z��
        Exit Sub
    End If

    ' �V�[�g�̃f�[�^���N���A(�T�u)
    Call ClearSheetDataSub(ActiveSheet)

End Sub

' �V�[�g�̃f�[�^���N���A(�T�u)
Private Sub ClearSheetDataSub(sheet As Worksheet)
    Dim row As Long
    Dim col As Long
    Dim rowLast As Long
    Dim colLast As Long

    ' �ŏI�s�̍s�ԍ����擾
    row = ROW_DATA_BEGIN
    rowLast = 0
    Do Until Trim(sheet.Cells(row, COL_NO)) = ""
        rowLast = row
        row = row + 1
    Loop

    ' �ŏI��̗�ԍ����擾
    col = COL_DATA_BEGIN
    colLast = 0
    Do Until Trim(sheet.Cells(ROW_HEADER, col)) = ""
        colLast = col
        col = col + 1
    Loop

    ' �w�b�_�s���N���A(�Z���̒l�݂̂��N���A)
    If colLast >= COL_DATA_BEGIN Then
        Call sheet.Range(sheet.Cells(ROW_HEADER, COL_DATA_BEGIN), _
                         sheet.Cells(ROW_HEADER, colLast)).ClearContents
    End If

    ' �f�[�^���N���A(�Z���̒l�݂̂��N���A)
    If colLast >= COL_DATA_BEGIN And rowLast >= ROW_DATA_BEGIN Then
        Call sheet.Range(sheet.Cells(ROW_DATA_BEGIN, COL_NO), _
                         sheet.Cells(rowLast, colLast)).ClearContents
    End If

End Sub

' Long�ɕϊ�(�ϊ��G���[����errVal��Ԃ�)
Private Function CLngErrVal(val, errVal) As Long

    On Error GoTo Label_Exit
    CLngErrVal = CLng(val)
    Exit Function

Label_Exit:

    CLngErrVal = errVal

End Function

' �A�Ԑ���
Sub MakeSeqNumber()
    Dim cell As Range
    Dim count As Long
    Dim str As String

    ' �I��͈͂̃Z������������
    count = 1
    str = ""
    For Each cell In Selection
        If count = 1 Then
            ' �擪�̃Z�����x�[�X�ɂ���
            str = cell.Value
        Else
            ' �A�Ԑ���(�T�u)
            str = MakeSeqNumberSub(str)
            cell.Value = str
        End If
        count = count + 1
    Next

End Sub

' �A�Ԑ���(�T�u)
Private Function MakeSeqNumberSub(str As String) As String
    Dim ret As String
    Dim ch As String
    Dim chVal As Long
    Dim convFlag As Boolean
    Dim i As Long

    ' ���A�Ԃ̐�����
    ' �E�����̕�������A�E����1�������������Ă����B
    ' �E��ԉE�̕����� "0" �` "9" �Ȃ� "1" �` "0" �ɒu������B
    ' �E�J��オ�肪����΍��ׂ�̕��������l�ɒu�����Ă����B
    ' �E�r���� "0" �` "9" �ȊO�̕���������΁A�ϊ����I������B
    ' �E��ԍ��̕����́A����ȏ�J��オ������Ȃ��B(�����𑝂₳�Ȃ�)
    ret = ""
    convFlag = True
    For i = Len(str) To 1 Step -1
        ch = Mid(str, i, 1)
        If convFlag = True Then
            chVal = CLngErrVal(ch, -1) + 1
            If chVal = 10 Then
                ch = "0"
            ElseIf chVal > 0 Then
                ch = CStr(chVal)
            End If
            If chVal < 10 Then
                convFlag = False
            End If
        End If
        ret = ch & ret
    Next

    ' �߂�l��Ԃ�
    MakeSeqNumberSub = ret

End Function
