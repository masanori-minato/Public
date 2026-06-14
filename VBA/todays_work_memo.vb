' =============================================================================
' 本日の作業メモ作成マクロ
' =============================================================================
'
' 概要:
'   実行日の日付が入った空のExcelブックを新規作成し、
'   Windowsのダウンロードフォルダーへ保存します。
'
' 作成されるファイル:
'   本日の作業メモ_YYYYMMDD.xlsx
'
' 保存先:
'   C:\Users\<ユーザー名>\Downloads
'
' 実行方法:
'   ExcelのVBAエディターから CreateTodaysWorkMemo を実行します。
'
' 注意:
'   同じ日付のファイルがすでに存在する場合、上書き確認が表示されます。
' =============================================================================

Sub CreateTodaysWorkMemo()
    Dim wb As Workbook
    Dim fileName As String
    Dim savePath As String

    fileName = "本日の作業メモ_" & Format(Date, "YYYYMMDD") & ".xlsx"
    savePath = Environ("USERPROFILE") & "\Downloads\" & fileName

    Set wb = Workbooks.Add
    wb.SaveAs fileName:=savePath, FileFormat:=xlOpenXMLWorkbook

    MsgBox "作成しました：" & fileName
End Sub
