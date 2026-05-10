Sub ImportSheetsByCopying()
    Dim wbDest As Workbook
    Dim wbSrc As Workbook
    Dim wb As Workbook
    Dim srcFolder As String
    Dim fso As Object
    Dim folder As Object
    Dim file As Object
    Dim ext As String
    Dim count As Integer

    ' 取り込み先ブック（「本日の作業メモ_」で始まる、開いているブック）を検索
    For Each wb In Workbooks
        If Left(wb.Name, Len("本日の作業メモ_")) = "本日の作業メモ_" Then
            Set wbDest = wb
            Exit For
        End If
    Next wb

    If wbDest Is Nothing Then
        MsgBox "「本日の作業メモ_」で始まるブックが開かれていません。", vbExclamation
        Exit Sub
    End If

    ' 取込元フォルダのパス（ダウンロード\CSV\excel_out）
    srcFolder = Environ("USERPROFILE") & "\Downloads\CSV\excel_out\"

    If Dir(srcFolder, vbDirectory) = "" Then
        MsgBox "フォルダが見つかりません:" & vbCrLf & srcFolder, vbExclamation
        Exit Sub
    End If

    Set fso = CreateObject("Scripting.FileSystemObject")
    Set folder = fso.GetFolder(srcFolder)

    count = 0

    For Each file In folder.Files
        ext = LCase(fso.GetExtensionName(file.Name))
        If ext = "xlsx" Or ext = "xls" Or ext = "xlsm" Then
            ' 取込元ブックを開く
            Set wbSrc = Workbooks.Open(file.Path)
            ' 1枚しかないシートを取り込み先ブックの末尾にコピー
            wbSrc.Sheets(1).Copy After:=wbDest.Sheets(wbDest.Sheets.Count)
            ' コピーしたシート名をファイル名（拡張子なし）に変更（使用不可文字を除去・31文字以内に切り詰め）
            wbDest.Sheets(wbDest.Sheets.Count).Name = Left(SanitizeSheetName(fso.GetBaseName(file.Name)), 31)
            ' 取込元ブックを保存せずに閉じる
            wbSrc.Close SaveChanges:=False
            count = count + 1
        End If
    Next file

    MsgBox count & " 件のシートを取り込みました。", vbInformation
End Sub

Private Function SanitizeSheetName(name As String) As String
    Dim result As String
    Dim i As Integer
    Dim c As String
    Const INVALID_CHARS As String = ":\/?*[]"
    result = ""
    For i = 1 To Len(name)
        c = Mid(name, i, 1)
        If InStr(INVALID_CHARS, c) = 0 Then
            result = result & c
        End If
    Next i
    SanitizeSheetName = result
End Function
