' =============================================================================
' 本日の作業メモ シート並べ替えマクロ
' =============================================================================
'
' 概要:
'   開いているExcelブックの中から、ファイル名が
'   「本日の作業メモ_」で始まるブックを探し、そのブック内の全シートを
'   シート名の昇順に並べ替えます。
'
' 対象ブックの例:
'   本日の作業メモ_20260510.xlsx
'
' 実行方法:
'   対象ブックをExcelで開いた状態にして、VBAエディターから
'   SortSheets を実行します。
'
' 注意:
'   - 対象ブックが開かれていない場合は処理を中止します。
'   - ワークシートだけでなく、対象ブックの全Sheetsを並べ替えます。
'   - 同じ条件に該当するブックが複数ある場合は、最初に見つけたブックが対象です。
' =============================================================================

Sub SortSheets()
    Dim wb As Workbook
    Dim targetWb As Workbook
    Dim i As Long, j As Long
    Dim swapped As Boolean
    Dim sheetNames() As String
    Dim n As Long

    ' 「本日の作業メモ_」で始まる開いているブックを検索
    For Each wb In Application.Workbooks
        If wb.Name Like "本日の作業メモ_*" Then
            Set targetWb = wb
            Exit For
        End If
    Next wb

    If targetWb Is Nothing Then
        MsgBox "「本日の作業メモ_」で始まるブックが見つかりません。" & vbCrLf & _
               "対象ブックを開いてから実行してください。", vbExclamation, "対象ブック未検出"
        Exit Sub
    End If

    n = targetWb.Sheets.Count
    ReDim sheetNames(1 To n)

    For i = 1 To n
        sheetNames(i) = targetWb.Sheets(i).Name
    Next i

    ' バブルソート（昇順）
    Do
        swapped = False
        For i = 1 To n - 1
            If sheetNames(i) > sheetNames(i + 1) Then
                Dim tmp As String
                tmp = sheetNames(i)
                sheetNames(i) = sheetNames(i + 1)
                sheetNames(i + 1) = tmp
                swapped = True
            End If
        Next i
    Loop While swapped

    ' シートを並び替え
    For i = 1 To n
        targetWb.Sheets(sheetNames(i)).Move After:=targetWb.Sheets(n)
    Next i

    MsgBox "「" & targetWb.Name & "」のシートをソートしました。", vbInformation, "完了"
End Sub
