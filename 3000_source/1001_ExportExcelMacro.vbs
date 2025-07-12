'
' 1001_ExportExcelMacro.vbs
' 2025-3-26 v1.05
'
' ＜概要＞
'   Excelファイルのマクロのソースをエクスポートします。
'
' ＜使い方＞
'   1001_ExportExcelMacro.vbs excelFile exportPath
'
' ＜注意事項＞
'   ・本ツールを実行する場合、
'     Excelのオプション設定で、セキュリティセンター(またはトラストセンター)の設定を開き、
'     「マクロの設定」-「VBA プロジェクト オブジェクト モデルへのアクセスを信頼する」
'     にチェックを入れる必要があります。
'     (この設定をしないと、アクセス権のエラーが発生します)
'   ・エクスポート先のフォルダが存在しない場合、作成します。
'   ・本ツールの戻り値は、あてにしないでください。
'     (エラーが発生しても、0 (正常) が返るケースがあります)
'   ・v1.05 から、Accessファイル (拡張子が .mdb のファイル) のマクロのソースも
'     エクスポート可能になりました。
'
' ＜参考URL＞
'   https://gist.github.com/aimoriu/7718005
'   https://taka-2.hatenablog.jp/entry/20090907/p2
'
Option Explicit

Dim objFSO
Dim objShell
Dim strFilePath
Dim strExportPath
Dim intRet

' 変数の初期化
strFilePath = ""
strExportPath = ""
intRet = 1

' 引数の取得
Set objFSO = CreateObject("Scripting.FileSystemObject")
If WScript.Arguments.Count = 2 Then
    strFilePath   = objFSO.GetAbsolutePathName(WScript.Arguments.Item(0))
    strExportPath = objFSO.GetAbsolutePathName(WScript.Arguments.Item(1))
Else
    WScript.Echo "引数の数が不正です。"
    WScript.Quit intRet
End If

' エクスポート先のフォルダが存在しなければ作成する
Set objShell = CreateObject("WScript.Shell")
If Not objFSO.FolderExists(strExportPath) Then
    Call objShell.Run("cmd.exe /c mkdir """ & strExportPath & """", 0, True)
End If

' ファイルの拡張子をチェック
If LCase(objFSO.GetExtensionName(strFilePath)) = "mdb" Then
    ' Accessファイルのマクロのソースをエクスポート
    intRet = ExportAccessSource(strFilePath, strExportPath)
Else
    ' Excelファイルのマクロのソースをエクスポート
    intRet = ExportExcelSource(strFilePath, strExportPath)
End If

' オブジェクトの解放
Set objShell = Nothing
Set objFSO = Nothing

' 終了
WScript.Quit intRet

' == 以下は関数 ==

' Excelファイルのマクロのソースをエクスポート
Private Function ExportExcelSource(strFilePath, strExportPath)
    Dim objExcel
    Dim objWorkbook
    Dim objVBProject
    Dim objComponent
    Dim intAutomation
    Dim intRet

    ' 戻り値の初期化
    intRet = 0

    ' Excelファイルのオープン
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

    ' マクロへのアクセス権をチェック
    ' (Excelファイルが開きっぱなしになることを防ぐため、事前にチェックする)
    On Error Resume Next
    For Each objComponent In objVBProject.VBComponents
        ' NOP
    Next
    If Err.Number <> 0 Then
        WScript.Echo "VBA プロジェクト オブジェクト モデルへのアクセス権がありません。" & vbCrLf & _
                     "(Excel の設定が必要です)"
        intRet = 1
    End If
    On Error Goto 0

    ' ソースファイルのエクスポート
    If intRet = 0 Then
        Call ExportSourceFile(objVBProject, strExportPath)
    End If

    ' Excelファイルのクローズ
    objWorkbook.Close False
    objExcel.AutomationSecurity = intAutomation
    objExcel.Cursor = -4143         ' (=xlDefault)
    objExcel.ScreenUpdating = True
    objExcel.EnableEvents = True
    objExcel.DisplayAlerts = True
    objExcel.Quit

    ' オブジェクトの解放
    Set objComponent = Nothing
    Set objVBProject = Nothing
    Set objWorkbook = Nothing
    Set objExcel = Nothing

    ' 戻り値を設定
    ExportExcelSource = intRet

End Function

' Accessファイルのマクロのソースをエクスポート
Private Function ExportAccessSource(strFilePath, strExportPath)
    Dim objAccess
    Dim objVBProject
    Dim intRet

    ' 戻り値の初期化
    intRet = 0

    ' Accessファイルのオープン
    Set objAccess = CreateObject("Access.Application")
    objAccess.OpenCurrentDatabase(strFilePath)
    Set objVBProject = objAccess.VBE.ActiveVBProject

    ' ソースファイルのエクスポート
    Call ExportSourceFile(objVBProject, strExportPath)

    ' Accessファイルのクローズ
    objAccess.Quit

    ' オブジェクトの解放
    Set objVBProject = Nothing
    Set objAccess = Nothing

    ' 戻り値を設定
    ExportAccessSource = intRet

End Function

' ソースファイルのエクスポート
Private Sub ExportSourceFile(objVBProject, strExportPath)
    Dim objComponent
    Dim strExportFilePath

    ' コンポーネントの検索
    For Each objComponent In objVBProject.VBComponents
        Select Case objComponent.Type
            ' 標準モジュール
            Case 1
                strExportFilePath = strExportPath & "\" & objComponent.Name & ".bas"
            ' クラスモジュール
            Case 2
                strExportFilePath = strExportPath & "\" & objComponent.Name & ".cls"
            ' Microsoft Form (ユーザーフォーム)
            Case 3
                strExportFilePath = strExportPath & "\" & objComponent.Name & ".frm"
            ' ActiveX デザイナ
            Case 11
                strExportFilePath = strExportPath & "\" & objComponent.Name & ".cls"
            ' Document モジュール (シートとThisWorkBook)
            Case 100
                strExportFilePath = strExportPath & "\" & objComponent.Name & ".cls"
            ' その他
            Case Else
                strExportFilePath = ""
        End Select

        ' (デバッグ用)
        'If strExportFilePath <> "" Then
        '    WScript.Echo objComponent.Name & " : " & objComponent.CodeModule.CountOfLines
        'End If

        ' エクスポート (空のソースファイルは除外)
        If strExportFilePath <> "" And objComponent.CodeModule.CountOfLines > 0 Then
            objComponent.Export strExportFilePath
        End If
    Next

    ' オブジェクトの解放
    Set objComponent = Nothing

End Sub
