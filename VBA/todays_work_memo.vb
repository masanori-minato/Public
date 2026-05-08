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
