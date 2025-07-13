Attribute VB_Name = "Module1"
Option Explicit

' == CSVファイルの読み書き ==

' ファイル名の定数
Const DEFAULT_CSV_FILE_NAME_1 As String = "SampleData.csv"  ' (読み込み用)
Const DEFAULT_CSV_FILE_NAME_2 As String = "SampleData1.csv" ' (書き出し用)

' 文字コードの定数
Const CHAR_CODE_1 As String = "utf-8"       ' (読み込み用)
'Const CHAR_CODE_1 As String = "shift_jis"   ' (読み込み用)
Const CHAR_CODE_2 As String = "utf-8"       ' (書き出し用)
'Const CHAR_CODE_2 As String = "shift_jis"   ' (書き出し用)
Const UTF_8_BOM_1 As Boolean = False        ' (読み込み用)(未使用)
Const UTF_8_BOM_2 As Boolean = False        ' (書き出し用)

' エラーの定数
Const READ_ERROR As Long = 1                ' (読み込み用)
Const WRITE_ERROR As Long = 2               ' (書き出し用)

' 行の定数
Const ROW_HEADER As Long = 6
Const ROW_DATA_BEGIN As Long = 7

' 列の定数
Const COL_NO As Long = 1
Const COL_SKIP As Long = 2
Const COL_DATA_BEGIN As Long = 3
Const COL_DATA_CHECK As Long = 3

' エラーメッセージ表示
Private Sub DispErrorMessage(msg As String, lineNo As Long, errType As Long)
    Dim msg1 As String

    ' エラーのタイプで場合分け
    msg1 = ""
    Select Case errType
        Case READ_ERROR
            If lineNo > 0 Then
                msg1 = "(" & lineNo & "行目の読み込みでエラーが発生しました)"
            End If
        Case WRITE_ERROR
            If lineNo > 0 Then
                msg1 = "(" & lineNo & "行目の書き出しでエラーが発生しました)"
            End If
    End Select

    ' エラーメッセージ表示
    Call MsgBox(msg & vbCrLf & msg1, vbOKOnly + vbExclamation)

End Sub

' ダブルクォート削除
Private Function RemoveDoubleQuote(str As String) As String
    Dim ret As String
    Dim ch As String
    Dim quoteFlag As Boolean
    Dim i As Long

    ' 戻り値の初期化
    ret = ""

    ' 空文字列なら抜ける
    If str = "" Then
        RemoveDoubleQuote = ret
        Exit Function
    End If

    ' 1文字ずつ調べてダブルクォートを削除する
    quoteFlag = False
    For i = 1 To Len(str)
        ch = Mid(str, i, 1)
        ' ダブルクォートのとき
        If ch = """" Then
            ' ダブルクォートが2個連続したとき
            If quoteFlag = True Then
                ' ダブルクォートを1個だけにする
                ret = ret & ch
            End If
            quoteFlag = Not quoteFlag
        Else
            ret = ret & ch
            quoteFlag = False
        End If
    Next

    ' 戻り値を返す
    RemoveDoubleQuote = ret

End Function

' ダブルクォートで囲う
' (specialに文字列を指定すると、それらの文字があるときだけダブルクォートで囲う)
Private Function AddDoubleQuote(str As String, Optional special As String = "") As String
    Dim ret As String
    Dim convFlag As Boolean
    Dim i As Long

    ' 戻り値の初期化
    ret = ""

    ' 空文字列なら抜ける
    If str = "" Then
        AddDoubleQuote = ret
        Exit Function
    End If

    ' 特殊文字チェック
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

    ' 変換ありのとき
    If convFlag = True Then
        ' ダブルクォート1個を、ダブルクォート2個に置換
        ret = Replace(str, """", """""")
        ' ダブルクォートで囲う
        ret = """" & ret & """"
    Else
        ret = str
    End If

    ' 戻り値を返す
    AddDoubleQuote = ret

End Function

' ダブルクォート対応のSplit
' (区切り文字(delim)は1文字のみ対応)
Private Function SplitWithDoubleQuote(str As String, delim As String) As String()
    Dim ret() As String
    Dim retCount As Long
    Dim ch As String
    Dim quoteFlag As Boolean
    Dim i As Long
    Dim i1 As Long

    ' 1文字ずつ調べて分割する
    retCount = 0
    quoteFlag = False
    i1 = 1
    For i = 1 To Len(str)
        ch = Mid(str, i, 1)
        ' ダブルクォートのとき
        If ch = """" Then
            quoteFlag = Not quoteFlag
        ' 区切り文字のとき
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

' CSVファイルの読み込み
Sub ReadCSVFile()
    Dim defaultFileName As String
    'Dim fnameVariant As Variant
    Dim fname As String
    Dim fd As FileDialog
    'Dim wso As Object
    'Dim curDirOld As String
    Dim delim As String

    ' ファイル名の取得
    ' (初期フォルダはワークブックのあるフォルダとする)
    ' (GetOpenFilename() では初期ファイル名を指定できなかったため、
    '  FileDialog(msoFileDialogOpen) を使用するようにした)
    '
    'Set wso = CreateObject("WScript.Shell")
    'curDirOld = wso.CurrentDirectory
    'wso.CurrentDirectory = ThisWorkbook.Path
    ''defaultFileName = DEFAULT_CSV_FILE_NAME_1
    'fnameVariant = Application.GetOpenFilename("CSVファイル (*.csv),*.csv,TSVファイル (*.tsv),*.tsv,すべてのファイル (*.*),*.*")
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
    Call fd.Filters.Add("CSVファイル", "*.csv")
    Call fd.Filters.Add("TSVファイル", "*.tsv")
    Call fd.Filters.Add("すべてのファイル", "*.*")
    'fd.FilterIndex = fd.Filters.count
    fd.FilterIndex = 0
    If fd.Show = False Then
        Exit Sub
    End If
    fname = fd.SelectedItems(1)

    ' 区切り文字を設定
    If LCase(Right(fname, 4)) = ".tsv" Then
        delim = vbTab
    Else
        delim = ","
    End If
    
    ' CSVファイルの読み込み(サブ)
    Call ReadCSVFileSub(fname, ActiveSheet, delim)

    ' オブジェクトの解放
    'Set wso = Nothing
    Set fd = Nothing

End Sub

' CSVファイルの読み込み(サブ)
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

    ' ファイルのオープン
    Set stream = CreateObject("ADODB.Stream")
    stream.Charset = CHAR_CODE_1
    Call stream.Open

    ' ファイルの読み込み
    Call stream.LoadFromFile(fname)

    ' 先頭行を読み込む
    lineNo = 1
    line = stream.ReadText(-2)
    lineNo = lineNo + 1

    ' ヘッダ解析
    header = SplitWithDoubleQuote(line, delim)
    headerNum = UBound(header)

    ' ヘッダを表示
    row = ROW_HEADER
    For i = 0 To headerNum
        col = COL_DATA_BEGIN + i
        sheet.Cells(row, col) = RemoveDoubleQuote(header(i))
    Next

    ' データの本体を読み込む
    row = ROW_DATA_BEGIN
    Do Until stream.EOS

        ' 1行読み込む
        line = stream.ReadText(-2)

        ' データ解析
        data = SplitWithDoubleQuote(line, delim)
        dataNum = UBound(data)
        If dataNum < headerNum Then
            Call DispErrorMessage("1行のデータが少なすぎます。", lineNo, READ_ERROR)
            GoTo Label_Exit
        ElseIf dataNum > headerNum Then
            Call DispErrorMessage("1行のデータが多すぎます。", lineNo, READ_ERROR)
            GoTo Label_Exit
        End If

        ' 画面更新を停止
        Application.ScreenUpdating = False
        Application.Cursor = xlWait

        ' No. を表示
        sheet.Cells(row, COL_NO) = CStr(lineNo - 1)

        ' データを表示
        For i = 0 To headerNum
            col = COL_DATA_BEGIN + i
            If i <= dataNum Then
                sheet.Cells(row, col) = RemoveDoubleQuote(data(i))
            Else
                sheet.Cells(row, col) = ""
            End If
        Next

        ' 画面更新を再開
        Application.Cursor = xlDefault
        Application.ScreenUpdating = True

        ' 画面を更新
        If (lineNo Mod 10) = 0 Then
            DoEvents
        End If

        ' 次の行へ
        lineNo = lineNo + 1
        row = row + 1
    Loop

Label_Exit:

    ' ファイルのクローズ
    If Not stream Is Nothing Then
        Call stream.Close
    End If

    ' オブジェクトの解放
    Set stream = Nothing

End Sub

' CSVファイルの書き出し
Sub WriteCSVFile()
    Dim defaultFileName As String
    Dim fnameVariant As Variant
    Dim fname As String
    'Dim fd As FileDialog
    Dim wso As Object
    Dim curDirOld As String
    Dim delim As String

    ' ファイル名の取得
    ' (初期フォルダはワークブックのあるフォルダとする)
    ' (FileDialog(msoFileDialogSaveAs) では拡張子を指定できなかったため、
    '  GetSaveAsFilename() を使用するようにした)
    '
    Set wso = CreateObject("WScript.Shell")
    curDirOld = wso.CurrentDirectory
    wso.CurrentDirectory = ThisWorkbook.Path
    defaultFileName = DEFAULT_CSV_FILE_NAME_2
    fnameVariant = Application.GetSaveAsFilename(defaultFileName, "CSVファイル (*.csv),*.csv,TSVファイル (*.tsv),*.tsv,すべてのファイル (*.*),*.*")
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
    ''Call fd.Filters.Add("CSVファイル", "*.csv")
    ''Call fd.Filters.Add("TSVファイル", "*.tsv")
    ''Call fd.Filters.Add("すべてのファイル", "*.*")
    ''fd.FilterIndex = fd.Filters.Count
    'fd.FilterIndex = 0
    'If fd.Show = False Then
    '    Exit Sub
    'End If
    'fname = fd.SelectedItems(1)

    ' 区切り文字を設定
    If LCase(Right(fname, 4)) = ".tsv" Then
        delim = vbTab
    Else
        delim = ","
    End If

    ' CSVファイルの書き出し(サブ)
    Call WriteCSVFileSub(fname, ActiveSheet, delim)

    ' オブジェクトの解放
    Set wso = Nothing
    'Set fd = Nothing

End Sub

' CSVファイルの書き出し(サブ)
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

    ' 最終列の列番号を取得
    col = COL_DATA_BEGIN
    colLast = 0
    Do Until Trim(sheet.Cells(ROW_HEADER, col)) = ""
        colLast = col
        col = col + 1
    Loop

    ' ファイルのオープン
    Set stream = CreateObject("ADODB.Stream")
    stream.Charset = CHAR_CODE_2
    Call stream.Open

    ' ワークシートからヘッダ行を読み込む
    line = ""
    For col = COL_DATA_BEGIN To colLast
        If col > COL_DATA_BEGIN Then
            line = line & delim
        End If
        header = AddDoubleQuote(sheet.Cells(ROW_HEADER, col), """" & delim & vbLf)
        line = line & header
    Next

    ' 1行書き込む
    Call stream.WriteText(line, 1)
    lineNo = lineNo + 1

    ' データ出力
    lineNo = 1
    row = ROW_DATA_BEGIN
    Do Until Trim(sheet.Cells(row, COL_NO)) = ""

        ' 出力スキップかデータがなければ、次の行へ
        If Trim(sheet.Cells(row, COL_SKIP)) <> "" Or Trim(sheet.Cells(row, COL_DATA_CHECK)) = "" Then
            GoTo Label_Next_1:
        End If

        ' ワークシートから1行読み込む
        line = ""
        For col = COL_DATA_BEGIN To colLast
            If col > COL_DATA_BEGIN Then
                line = line & delim
            End If
            data = AddDoubleQuote(sheet.Cells(row, col), """" & delim & vbLf)
            line = line & data
        Next

        ' 1行書き込む
        Call stream.WriteText(line, 1)
        lineNo = lineNo + 1

Label_Next_1:

        ' 次の行へ
        row = row + 1
    Loop

    ' ファイルの保存
    ' UTF-8 の BOM なしのとき
    If LCase(CHAR_CODE_2) = "utf-8" And UTF_8_BOM_2 = False Then
        ' BOM なしに変換する
        Set streamNoBom = CreateObject("ADODB.Stream")
        streamNoBom.Type = 1                    ' バイナリモード
        Call streamNoBom.Open
        stream.Position = 3                     ' BOMのサイズ
        stream.CopyTo streamNoBom
        Call streamNoBom.SaveToFile(fname, 2)   ' 上書き
    Else
        Call stream.SaveToFile(fname, 2)        ' 上書き
    End If

Label_Exit:

    ' ファイルのクローズ
    If Not streamNoBom Is Nothing Then
        Call streamNoBom.Close
    End If
    If Not stream Is Nothing Then
        Call stream.Close
    End If

    ' オブジェクトの解放
    Set streamNoBom = Nothing
    Set stream = Nothing

End Sub

' シートのデータをクリア
Sub ClearSheetData()
    Dim ret As Long

    ' 確認メッセージの表示
    ret = MsgBox("シートのデータを全てクリアします。よろしいですか?", _
                 vbYesNoCancel + vbQuestion + vbDefaultButton3)
    If ret <> vbYes Then
        ' キャンセル
        Exit Sub
    End If

    ' シートのデータをクリア(サブ)
    Call ClearSheetDataSub(ActiveSheet)

End Sub

' シートのデータをクリア(サブ)
Private Sub ClearSheetDataSub(sheet As Worksheet)
    Dim row As Long
    Dim col As Long
    Dim rowLast As Long
    Dim colLast As Long

    ' 最終行の行番号を取得
    row = ROW_DATA_BEGIN
    rowLast = 0
    Do Until Trim(sheet.Cells(row, COL_NO)) = ""
        rowLast = row
        row = row + 1
    Loop

    ' 最終列の列番号を取得
    col = COL_DATA_BEGIN
    colLast = 0
    Do Until Trim(sheet.Cells(ROW_HEADER, col)) = ""
        colLast = col
        col = col + 1
    Loop

    ' ヘッダ行をクリア(セルの値のみをクリア)
    If colLast >= COL_DATA_BEGIN Then
        Call sheet.Range(sheet.Cells(ROW_HEADER, COL_DATA_BEGIN), _
                         sheet.Cells(ROW_HEADER, colLast)).ClearContents
    End If

    ' データをクリア(セルの値のみをクリア)
    If colLast >= COL_DATA_BEGIN And rowLast >= ROW_DATA_BEGIN Then
        Call sheet.Range(sheet.Cells(ROW_DATA_BEGIN, COL_NO), _
                         sheet.Cells(rowLast, colLast)).ClearContents
    End If

End Sub

' Longに変換(変換エラー時はerrValを返す)
Private Function CLngErrVal(val, errVal) As Long

    On Error GoTo Label_Exit
    CLngErrVal = CLng(val)
    Exit Function

Label_Exit:

    CLngErrVal = errVal

End Function

' 連番生成
Sub MakeSeqNumber()
    Dim cell As Range
    Dim count As Long
    Dim str As String

    ' 選択範囲のセルを処理する
    count = 1
    str = ""
    For Each cell In Selection
        If count = 1 Then
            ' 先頭のセルをベースにする
            str = cell.Value
        Else
            ' 連番生成(サブ)
            str = MakeSeqNumberSub(str)
            cell.Value = str
        End If
        count = count + 1
    Next

End Sub

' 連番生成(サブ)
Private Function MakeSeqNumberSub(str As String) As String
    Dim ret As String
    Dim ch As String
    Dim chVal As Long
    Dim convFlag As Boolean
    Dim i As Long

    ' ＜連番の生成＞
    ' ・引数の文字列を、右から1文字ずつ処理していく。
    ' ・一番右の文字が "0" ～ "9" なら "1" ～ "0" に置換する。
    ' ・繰り上がりがあれば左隣りの文字も同様に置換していく。
    ' ・途中に "0" ～ "9" 以外の文字があれば、変換を終了する。
    ' ・一番左の文字は、それ以上繰り上がりをしない。(桁数を増やさない)
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

    ' 戻り値を返す
    MakeSeqNumberSub = ret

End Function
