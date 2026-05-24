Option Explicit

' Qlik Sense の Set分析結果を、Excelで再集計した値と突合するための標準モジュールです。
'
' 使い方:
' 1. GenerateQlikTestDataCsv を実行し、Excel側で検証用明細CSVを作成する
' 2. 作成したCSVをQlikにロードする
' 3. QlikからSet分析の集計結果をCSVで出力する、または QlikResult シートに貼り付ける
'    必須列: テストID, Qlik結果
' 4. RunSetAnalysisEvidence を実行する
' 5. Evidence シートをテスト証跡として保存する

Private Const SHEET_DATA As String = "SalesData"
Private Const SHEET_QLIK As String = "QlikResult"
Private Const SHEET_TEST As String = "TestCases"
Private Const SHEET_EVIDENCE As String = "Evidence"
Private Const EPSILON As Double = 0.000001

Public Sub RunSetAnalysisEvidence()
    Dim qlikPath As Variant

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    PrepareWorkbook
    If LastRow(Worksheets(SHEET_DATA), 1) < 2 Then
        GenerateSalesDataSheet
    End If
    If LastRow(Worksheets(SHEET_TEST), 1) < 2 Then
        CreateDefaultTestCases
    End If
    If LastRow(Worksheets(SHEET_QLIK), 1) < 2 Then
        qlikPath = Application.GetOpenFilename("CSV Files (*.csv),*.csv", , "Qlikから出力したSet分析結果CSVを選択")
        If VarType(qlikPath) = vbBoolean Then GoTo Finally
        ImportCsvToSheet CStr(qlikPath), SHEET_QLIK
    End If

    CalculateEvidence

    Worksheets(SHEET_EVIDENCE).Activate
    MsgBox "Set分析の突合が完了しました。Evidence シートを確認してください。", vbInformation

Finally:
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
End Sub

Public Sub GenerateQlikTestDataCsv()
    Dim outputPath As Variant

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    PrepareWorkbook
    GenerateSalesDataSheet
    CreateDefaultTestCases

    outputPath = Application.GetSaveAsFilename( _
        InitialFileName:="SalesData_For_Qlik.csv", _
        FileFilter:="CSV Files (*.csv),*.csv", _
        Title:="Qlikにロードする明細CSVの保存先を指定")
    If VarType(outputPath) = vbBoolean Then GoTo Finally

    ExportSheetToCsv Worksheets(SHEET_DATA), CStr(outputPath)
    Worksheets(SHEET_DATA).Activate
    MsgBox "Qlikロード用CSVを作成しました。" & vbCrLf & CStr(outputPath), vbInformation

Finally:
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
End Sub

Public Sub CreateQlikResultTemplate()
    Dim ws As Worksheet

    Set ws = EnsureSheet(SHEET_QLIK)
    ws.Cells.Clear
    ws.Range("A1:B1").Value = Array("テストID", "Qlik結果")
    ws.Range("A2:A11").Value = Application.Transpose(Array("T001", "T002", "T003", "T004", "T005", "T006", "T007", "T008", "T009", "T010"))
    ws.Columns.AutoFit
End Sub

Public Sub CreateDefaultTestCases()
    Dim ws As Worksheet
    Set ws = EnsureSheet(SHEET_TEST)

    ws.Cells.Clear
    ws.Range("A1:J1").Value = Array("テストID", "説明", "集計項目", "年", "地域", "都道府県", "顧客業種", "商品カテゴリ", "年月From", "年月To")

    AddTestCase ws, 2, "T001", "売上実績合計", "売上実績"
    AddTestCase ws, 3, "T002", "売上予測合計", "売上予測"
    AddTestCase ws, 4, "T003", "2024年の売上実績", "売上実績", "2024"
    AddTestCase ws, 5, "T004", "関東の売上実績", "売上実績", , "関東"
    AddTestCase ws, 6, "T005", "関西の売上実績", "売上実績", , "関西"
    AddTestCase ws, 7, "T006", "ソフトウェアの売上実績", "売上実績", , , , , "ソフトウェア"
    AddTestCase ws, 8, "T007", "ハードウェアの売上実績", "売上実績", , , , , "ハードウェア"
    AddTestCase ws, 9, "T008", "2024年・関東・ソフトウェアの売上実績", "売上実績", "2024", "関東", , , "ソフトウェア"
    AddTestCase ws, 10, "T009", "製造業の売上実績", "売上実績", , , , "製造業"
    AddTestCase ws, 11, "T010", "2024/01-2024/12の売上実績", "売上実績", , , , , , "2024/01", "2024/12"

    ws.Columns.AutoFit
End Sub

Public Sub CalculateEvidence()
    Dim wsData As Worksheet
    Dim wsQlik As Worksheet
    Dim wsTest As Worksheet
    Dim wsEvidence As Worksheet
    Dim data As Variant
    Dim testCases As Variant
    Dim qlikResults As Object
    Dim dataHeader As Object
    Dim testHeader As Object
    Dim lastTestRow As Long
    Dim rowIndex As Long
    Dim outRow As Long
    Dim excelValue As Double
    Dim qlikValue As Double
    Dim diffValue As Double
    Dim testId As String

    Set wsData = Worksheets(SHEET_DATA)
    Set wsQlik = Worksheets(SHEET_QLIK)
    Set wsTest = Worksheets(SHEET_TEST)
    Set wsEvidence = EnsureSheet(SHEET_EVIDENCE)

    data = GetUsedRangeValues(wsData)
    If IsEmpty(data) Then Err.Raise vbObjectError + 1, , SHEET_DATA & " シートにデータがありません。"

    lastTestRow = LastRow(wsTest, 1)
    If lastTestRow < 2 Then Err.Raise vbObjectError + 2, , SHEET_TEST & " シートにテストケースがありません。"

    testCases = wsTest.Range("A1:J" & lastTestRow).Value
    Set dataHeader = BuildHeaderMap(data)
    Set testHeader = BuildHeaderMap(testCases)
    Set qlikResults = LoadQlikResults(wsQlik)

    ValidateDataHeaders dataHeader

    wsEvidence.Cells.Clear
    wsEvidence.Range("A1:H1").Value = Array("テストID", "説明", "Set分析イメージ", "Excel算出結果", "Qlik結果", "差異", "判定", "実行日時")

    outRow = 2
    For rowIndex = 2 To UBound(testCases, 1)
        testId = CStr(testCases(rowIndex, testHeader("テストID")))
        If Len(testId) > 0 Then
            excelValue = CalculateByTestCase(data, dataHeader, testCases, testHeader, rowIndex)
            qlikValue = GetQlikValue(qlikResults, testId)
            diffValue = excelValue - qlikValue

            wsEvidence.Cells(outRow, 1).Value = testId
            wsEvidence.Cells(outRow, 2).Value = testCases(rowIndex, testHeader("説明"))
            wsEvidence.Cells(outRow, 3).Value = BuildSetAnalysisImage(testCases, testHeader, rowIndex)
            wsEvidence.Cells(outRow, 4).Value = excelValue
            wsEvidence.Cells(outRow, 5).Value = qlikValue
            wsEvidence.Cells(outRow, 6).Value = diffValue
            wsEvidence.Cells(outRow, 7).Value = IIf(Abs(diffValue) <= EPSILON, "OK", "NG")
            wsEvidence.Cells(outRow, 8).Value = Now
            outRow = outRow + 1
        End If
    Next rowIndex

    FormatEvidence wsEvidence
End Sub

Private Sub PrepareWorkbook()
    EnsureSheet SHEET_DATA
    EnsureSheet SHEET_QLIK
    EnsureSheet SHEET_TEST
    EnsureSheet SHEET_EVIDENCE
End Sub

Private Sub GenerateSalesDataSheet()
    Dim ws As Worksheet
    Dim rowIndex As Long
    Dim txnIndex As Long
    Dim prefectures As Variant
    Dim regions As Variant
    Dim industries As Variant
    Dim categoryByProduct As Variant
    Dim nameByProduct As Variant
    Dim saleDate As Date
    Dim purchaseDate As Date
    Dim productId As Long
    Dim quantity As Long
    Dim unitCost As Double
    Dim purchaseAmount As Double
    Dim forecastAmount As Double
    Dim actualAmount As Double

    Set ws = EnsureSheet(SHEET_DATA)
    ws.Cells.Clear
    ws.Range("A1:T1").Value = Array("トランザクションID", "仕入日", "売上日", "都道府県", "顧客業種", "商品カテゴリ", "商品名", "数量", "仕入金額", "売上予測", "売上実績", "地域", "年", "四半期", "年四半期", "月", "年月", "年月キー", "週番号", "曜日")

    prefectures = Array("北海道", "東京都", "神奈川県", "埼玉県", "千葉県", "愛知県", "静岡県", "大阪府", "京都府", "兵庫県", "広島県", "福岡県", "熊本県", "沖縄県")
    regions = Array("北海道", "関東", "関東", "関東", "関東", "中部", "中部", "関西", "関西", "関西", "中四国", "九州", "九州", "九州")
    industries = Array("製造業", "小売業", "金融業", "情報通信業")
    categoryByProduct = Array("ハードウェア", "ハードウェア", "ソフトウェア", "ソフトウェア", "サービス")
    nameByProduct = Array("エンタープライズサーバー", "ハイエンドPC", "ERPライセンス", "BIツールライセンス", "クラウド導入支援")

    rowIndex = 2
    For txnIndex = 1 To 240
        saleDate = DateSerial(2024 + ((txnIndex - 1) Mod 3), ((txnIndex - 1) Mod 12) + 1, ((txnIndex - 1) Mod 27) + 1)
        purchaseDate = DateAdd("d", -(((txnIndex - 1) Mod 20) + 1), saleDate)
        productId = ((txnIndex - 1) Mod 5) + 1
        quantity = ((txnIndex * 7) Mod 15) + 1
        unitCost = Choose(((txnIndex - 1) Mod 4) + 1, 100000, 300000, 500000, 1000000)
        purchaseAmount = unitCost * quantity
        forecastAmount = RoundToUnit(purchaseAmount * (1.1 + (((txnIndex - 1) Mod 5) * 0.05)), 1000)
        actualAmount = RoundToUnit(purchaseAmount * (1.05 + (((txnIndex - 1) Mod 6) * 0.07)), 1000)

        ws.Cells(rowIndex, 1).Value = "T" & Format$(txnIndex, "000")
        ws.Cells(rowIndex, 2).Value = purchaseDate
        ws.Cells(rowIndex, 3).Value = saleDate
        ws.Cells(rowIndex, 4).Value = prefectures((txnIndex - 1) Mod (UBound(prefectures) + 1))
        ws.Cells(rowIndex, 5).Value = industries((txnIndex - 1) Mod (UBound(industries) + 1))
        ws.Cells(rowIndex, 6).Value = categoryByProduct(productId - 1)
        ws.Cells(rowIndex, 7).Value = nameByProduct(productId - 1)
        ws.Cells(rowIndex, 8).Value = quantity
        ws.Cells(rowIndex, 9).Value = purchaseAmount
        ws.Cells(rowIndex, 10).Value = forecastAmount
        ws.Cells(rowIndex, 11).Value = actualAmount
        ws.Cells(rowIndex, 12).Value = regions((txnIndex - 1) Mod (UBound(regions) + 1))
        ws.Cells(rowIndex, 13).Value = Year(saleDate)
        ws.Cells(rowIndex, 14).Value = "Q" & DatePart("q", saleDate)
        ws.Cells(rowIndex, 15).Value = Year(saleDate) & "/Q" & DatePart("q", saleDate)
        ws.Cells(rowIndex, 16).Value = Month(saleDate) & "月"
        ws.Cells(rowIndex, 17).Value = Format$(saleDate, "yyyy/mm")
        ws.Cells(rowIndex, 18).Value = Year(saleDate) * 100 + Month(saleDate)
        ws.Cells(rowIndex, 19).Value = DatePart("ww", saleDate, vbMonday, vbFirstFourDays)
        ws.Cells(rowIndex, 20).Value = Format$(saleDate, "aaa")

        rowIndex = rowIndex + 1
    Next txnIndex

    ws.Range("B:C").NumberFormatLocal = "yyyy/mm/dd"
    ws.Columns.AutoFit
End Sub

Private Function RoundToUnit(ByVal value As Double, ByVal unitValue As Long) As Double
    RoundToUnit = WorksheetFunction.Round(value / unitValue, 0) * unitValue
End Function

Private Sub ExportSheetToCsv(ByVal ws As Worksheet, ByVal csvPath As String)
    Dim tempBook As Workbook

    ws.Copy
    Set tempBook = ActiveWorkbook
    tempBook.SaveAs Filename:=csvPath, FileFormat:=xlCSVUTF8, CreateBackup:=False
    tempBook.Close SaveChanges:=False
End Sub

Private Sub ImportCsvToSheet(ByVal csvPath As String, ByVal sheetName As String)
    Dim ws As Worksheet
    Set ws = EnsureSheet(sheetName)
    ws.Cells.Clear

    With ws.QueryTables.Add(Connection:="TEXT;" & csvPath, Destination:=ws.Range("A1"))
        .TextFilePlatform = 65001
        .TextFileParseType = xlDelimited
        .TextFileCommaDelimiter = True
        .TextFileTextQualifier = xlTextQualifierDoubleQuote
        .Refresh BackgroundQuery:=False
        .Delete
    End With

    ws.Columns.AutoFit
End Sub

Private Sub AddTestCase( _
    ByVal ws As Worksheet, _
    ByVal targetRow As Long, _
    ByVal testId As String, _
    ByVal description As String, _
    ByVal measureName As String, _
    Optional ByVal yearValue As String = "", _
    Optional ByVal regionValue As String = "", _
    Optional ByVal prefValue As String = "", _
    Optional ByVal industryValue As String = "", _
    Optional ByVal categoryValue As String = "", _
    Optional ByVal ymFrom As String = "", _
    Optional ByVal ymTo As String = "")

    ws.Cells(targetRow, 1).Value = testId
    ws.Cells(targetRow, 2).Value = description
    ws.Cells(targetRow, 3).Value = measureName
    ws.Cells(targetRow, 4).Value = yearValue
    ws.Cells(targetRow, 5).Value = regionValue
    ws.Cells(targetRow, 6).Value = prefValue
    ws.Cells(targetRow, 7).Value = industryValue
    ws.Cells(targetRow, 8).Value = categoryValue
    ws.Cells(targetRow, 9).Value = ymFrom
    ws.Cells(targetRow, 10).Value = ymTo
End Sub

Private Function CalculateByTestCase( _
    ByRef data As Variant, _
    ByVal dataHeader As Object, _
    ByRef testCases As Variant, _
    ByVal testHeader As Object, _
    ByVal testRow As Long) As Double

    Dim dataRow As Long
    Dim measureName As String
    Dim totalValue As Double

    measureName = CStr(testCases(testRow, testHeader("集計項目")))
    If Not dataHeader.Exists(measureName) Then Err.Raise vbObjectError + 3, , "明細データに集計項目 [" & measureName & "] がありません。"

    For dataRow = 2 To UBound(data, 1)
        If MatchesTestCase(data, dataHeader, testCases, testHeader, testRow, dataRow) Then
            totalValue = totalValue + ToDouble(data(dataRow, dataHeader(measureName)))
        End If
    Next dataRow

    CalculateByTestCase = totalValue
End Function

Private Function MatchesTestCase( _
    ByRef data As Variant, _
    ByVal dataHeader As Object, _
    ByRef testCases As Variant, _
    ByVal testHeader As Object, _
    ByVal testRow As Long, _
    ByVal dataRow As Long) As Boolean

    If Not MatchEquals(data, dataHeader, "年", testCases(testRow, testHeader("年")), dataRow) Then Exit Function
    If Not MatchEquals(data, dataHeader, "地域", testCases(testRow, testHeader("地域")), dataRow) Then Exit Function
    If Not MatchEquals(data, dataHeader, "都道府県", testCases(testRow, testHeader("都道府県")), dataRow) Then Exit Function
    If Not MatchEquals(data, dataHeader, "顧客業種", testCases(testRow, testHeader("顧客業種")), dataRow) Then Exit Function
    If Not MatchEquals(data, dataHeader, "商品カテゴリ", testCases(testRow, testHeader("商品カテゴリ")), dataRow) Then Exit Function
    If Not MatchYearMonthRange(data, dataHeader, testCases, testHeader, testRow, dataRow) Then Exit Function

    MatchesTestCase = True
End Function

Private Function MatchEquals( _
    ByRef data As Variant, _
    ByVal dataHeader As Object, _
    ByVal fieldName As String, _
    ByVal expectedValue As Variant, _
    ByVal dataRow As Long) As Boolean

    Dim expectedText As String
    expectedText = Trim$(CStr(expectedValue))

    If Len(expectedText) = 0 Then
        MatchEquals = True
    Else
        MatchEquals = (CStr(data(dataRow, dataHeader(fieldName))) = expectedText)
    End If
End Function

Private Function MatchYearMonthRange( _
    ByRef data As Variant, _
    ByVal dataHeader As Object, _
    ByRef testCases As Variant, _
    ByVal testHeader As Object, _
    ByVal testRow As Long, _
    ByVal dataRow As Long) As Boolean

    Dim ym As Long
    Dim ymFrom As Long
    Dim ymTo As Long
    Dim fromText As String
    Dim toText As String

    fromText = Trim$(CStr(testCases(testRow, testHeader("年月From"))))
    toText = Trim$(CStr(testCases(testRow, testHeader("年月To"))))

    If Len(fromText) = 0 And Len(toText) = 0 Then
        MatchYearMonthRange = True
        Exit Function
    End If

    ym = YearMonthToNumber(CStr(data(dataRow, dataHeader("年月"))))
    If Len(fromText) > 0 Then
        ymFrom = YearMonthToNumber(fromText)
        If ym < ymFrom Then Exit Function
    End If
    If Len(toText) > 0 Then
        ymTo = YearMonthToNumber(toText)
        If ym > ymTo Then Exit Function
    End If

    MatchYearMonthRange = True
End Function

Private Function BuildSetAnalysisImage(ByRef testCases As Variant, ByVal testHeader As Object, ByVal testRow As Long) As String
    Dim filters As Collection
    Dim measureName As String
    Dim filterText As String

    Set filters = New Collection
    measureName = CStr(testCases(testRow, testHeader("集計項目")))

    AddFilterText filters, "年", testCases(testRow, testHeader("年"))
    AddFilterText filters, "地域", testCases(testRow, testHeader("地域"))
    AddFilterText filters, "都道府県", testCases(testRow, testHeader("都道府県"))
    AddFilterText filters, "顧客業種", testCases(testRow, testHeader("顧客業種"))
    AddFilterText filters, "商品カテゴリ", testCases(testRow, testHeader("商品カテゴリ"))

    filterText = JoinCollection(filters, ", ")
    If Len(Trim$(CStr(testCases(testRow, testHeader("年月From"))))) > 0 Or Len(Trim$(CStr(testCases(testRow, testHeader("年月To"))))) > 0 Then
        If Len(filterText) > 0 Then filterText = filterText & ", "
        filterText = filterText & "年月={""" & BuildYearMonthCondition(testCases, testHeader, testRow) & """}"
    End If

    If Len(filterText) = 0 Then
        BuildSetAnalysisImage = "Sum(" & measureName & ")"
    Else
        BuildSetAnalysisImage = "Sum({<" & filterText & ">} " & measureName & ")"
    End If
End Function

Private Sub AddFilterText(ByVal filters As Collection, ByVal fieldName As String, ByVal value As Variant)
    Dim textValue As String
    textValue = Trim$(CStr(value))
    If Len(textValue) > 0 Then filters.Add fieldName & "={'" & textValue & "'}"
End Sub

Private Function BuildYearMonthCondition(ByRef testCases As Variant, ByVal testHeader As Object, ByVal testRow As Long) As String
    Dim conditionText As String
    If Len(Trim$(CStr(testCases(testRow, testHeader("年月From"))))) > 0 Then
        conditionText = ">=" & testCases(testRow, testHeader("年月From"))
    End If
    If Len(Trim$(CStr(testCases(testRow, testHeader("年月To"))))) > 0 Then
        conditionText = conditionText & "<=" & testCases(testRow, testHeader("年月To"))
    End If
    BuildYearMonthCondition = conditionText
End Function

Private Function LoadQlikResults(ByVal ws As Worksheet) As Object
    Dim values As Variant
    Dim header As Object
    Dim results As Object
    Dim rowIndex As Long
    Dim testId As String

    values = GetUsedRangeValues(ws)
    If IsEmpty(values) Then Err.Raise vbObjectError + 4, , SHEET_QLIK & " シートにデータがありません。"

    Set header = BuildHeaderMap(values)
    If Not header.Exists("テストID") Then Err.Raise vbObjectError + 5, , SHEET_QLIK & " に [テストID] 列がありません。"
    If Not header.Exists("Qlik結果") Then Err.Raise vbObjectError + 6, , SHEET_QLIK & " に [Qlik結果] 列がありません。"

    Set results = CreateObject("Scripting.Dictionary")
    For rowIndex = 2 To UBound(values, 1)
        testId = CStr(values(rowIndex, header("テストID")))
        If Len(testId) > 0 Then results(testId) = ToDouble(values(rowIndex, header("Qlik結果")))
    Next rowIndex

    Set LoadQlikResults = results
End Function

Private Function GetQlikValue(ByVal qlikResults As Object, ByVal testId As String) As Double
    If Not qlikResults.Exists(testId) Then Err.Raise vbObjectError + 7, , "Qlik結果にテストID [" & testId & "] がありません。"
    GetQlikValue = qlikResults(testId)
End Function

Private Sub ValidateDataHeaders(ByVal header As Object)
    Dim requiredHeaders As Variant
    Dim index As Long

    requiredHeaders = Array("年", "地域", "都道府県", "顧客業種", "商品カテゴリ", "年月", "売上実績", "売上予測", "仕入金額")
    For index = LBound(requiredHeaders) To UBound(requiredHeaders)
        If Not header.Exists(CStr(requiredHeaders(index))) Then
            Err.Raise vbObjectError + 8, , "明細データに [" & requiredHeaders(index) & "] 列がありません。"
        End If
    Next index
End Sub

Private Function BuildHeaderMap(ByRef values As Variant) As Object
    Dim header As Object
    Dim colIndex As Long
    Dim key As String

    Set header = CreateObject("Scripting.Dictionary")
    For colIndex = 1 To UBound(values, 2)
        key = Trim$(CStr(values(1, colIndex)))
        If Len(key) > 0 Then header(key) = colIndex
    Next colIndex

    Set BuildHeaderMap = header
End Function

Private Function GetUsedRangeValues(ByVal ws As Worksheet) As Variant
    If Application.WorksheetFunction.CountA(ws.Cells) = 0 Then
        GetUsedRangeValues = Empty
    Else
        GetUsedRangeValues = ws.UsedRange.Value
    End If
End Function

Private Function EnsureSheet(ByVal sheetName As String) As Worksheet
    On Error Resume Next
    Set EnsureSheet = Worksheets(sheetName)
    On Error GoTo 0

    If EnsureSheet Is Nothing Then
        Set EnsureSheet = Worksheets.Add(After:=Worksheets(Worksheets.Count))
        EnsureSheet.Name = sheetName
    End If
End Function

Private Function LastRow(ByVal ws As Worksheet, ByVal colIndex As Long) As Long
    LastRow = ws.Cells(ws.Rows.Count, colIndex).End(xlUp).Row
End Function

Private Function ToDouble(ByVal value As Variant) As Double
    Dim textValue As String

    If IsNumeric(value) Then
        ToDouble = CDbl(value)
    Else
        textValue = Replace(CStr(value), ",", "")
        If Len(Trim$(textValue)) = 0 Then
            ToDouble = 0
        Else
            ToDouble = CDbl(textValue)
        End If
    End If
End Function

Private Function YearMonthToNumber(ByVal value As String) As Long
    Dim textValue As String

    textValue = Replace(Trim$(value), "/", "")
    textValue = Replace(textValue, "-", "")
    textValue = Replace(textValue, "年", "")
    textValue = Replace(textValue, "月", "")

    If Len(textValue) = 6 And IsNumeric(textValue) Then
        YearMonthToNumber = CLng(textValue)
    Else
        Err.Raise vbObjectError + 9, , "年月 [" & value & "] をYYYY/MM形式として解釈できません。"
    End If
End Function

Private Function JoinCollection(ByVal values As Collection, ByVal delimiter As String) As String
    Dim parts() As String
    Dim index As Long

    If values.Count = 0 Then
        JoinCollection = ""
        Exit Function
    End If

    ReDim parts(1 To values.Count)
    For index = 1 To values.Count
        parts(index) = CStr(values(index))
    Next index

    JoinCollection = Join(parts, delimiter)
End Function

Private Sub FormatEvidence(ByVal ws As Worksheet)
    Dim lastEvidenceRow As Long

    lastEvidenceRow = LastRow(ws, 1)
    ws.Columns.AutoFit
    ws.Range("A1:H1").Font.Bold = True
    ws.Range("D:F").NumberFormatLocal = "#,##0.00"
    ws.Range("H:H").NumberFormatLocal = "yyyy/mm/dd hh:mm:ss"

    If lastEvidenceRow >= 2 Then
        With ws.Range("G2:G" & lastEvidenceRow)
            .FormatConditions.Delete
            .FormatConditions.Add Type:=xlCellValue, Operator:=xlEqual, Formula1:="=""OK"""
            .FormatConditions(1).Interior.Color = RGB(198, 239, 206)
            .FormatConditions.Add Type:=xlCellValue, Operator:=xlEqual, Formula1:="=""NG"""
            .FormatConditions(2).Interior.Color = RGB(255, 199, 206)
        End With
    End If
End Sub
