Attribute VB_Name = "Xtra_SupplierAgreements"
Option Explicit

Public Function ModuleExists() As Boolean
    ModuleExists = True
End Function

Public Function GeneratePPS100UploadSheets() As Boolean
    Dim rngSelect As Range, wsSource As Worksheet, wbTarget As Workbook
    Dim outStartRow As Long: outStartRow = 8
    Dim outStartCol As Long: outStartCol = 2 ' Column B
    
    ' 1. SELECT SOURCE
    Dim answer As VbMsgBoxResult
    answer = MsgBox("Is the source workbook already open?" & vbCrLf & vbCrLf & _
                    "Click YES to select a cell from it." & vbCrLf & _
                    "Click NO to open it first, then re-run.", _
                    vbYesNo + vbQuestion, "Source Workbook")
    If answer = vbNo Then
        GeneratePPS100UploadSheets = False
        Exit Function
    End If

    On Error Resume Next
    Set rngSelect = Application.InputBox("Select any cell in the SOURCE sheet", "Select Master Data", Type:=8)
    On Error GoTo 0

    If rngSelect Is Nothing Then
        GeneratePPS100UploadSheets = False
        Exit Function
    End If
    
    Set wsSource = rngSelect.Parent
    Set wbTarget = ThisWorkbook
    
    ' 2. SETUP SHEETS & HEADERS
    Dim wsH As Worksheet, wsL As Worksheet, wsS As Worksheet
    Set wsH = SetupSheet(wbTarget, "PPS100MI AddAgrHead", outStartRow, outStartCol, _
              Array("SUNO", "AGTP", "FVDT", "UVDT", "AGNB", "TX30", "CUCD", "MODL", "TEDL", "TEPY", "PAST", "BUYE", "AGPT"))
    Set wsL = SetupSheet(wbTarget, "PPS100MI AddAgrLine", outStartRow, outStartCol, _
              Array("SUNO", "AGNB", "GRPI", "OBV1", "FVDT", "PUPR", "UVDT", "SAGL"))
    Set wsS = SetupSheet(wbTarget, "PPS100MI AddStgPrice", outStartRow, outStartCol, _
              Array("SUNO", "AGNB", "GRPI", "OBV1", "FVDT", "FRQT", "PUPR"))

    ' 3. VARIABLES & DATA PREP
    Dim lastRow As Long: lastRow = wsSource.Cells(wsSource.Rows.count, "I").End(xlUp).row
    Dim suno As String, agnb As String, fvdt As String, uvdt As String
    Dim i As Long, j As Long
    
    suno = wsSource.Range("G1").value
    agnb = wsSource.Range("G3").value
    uvdt = Format(wsSource.Range("G8").value, "yyyymmdd")
    fvdt = Year(wsSource.Range("G8").value) & "0101"
    
    ' --- SECTION 1: HEADER ---
    wsH.Cells(outStartRow + 1, outStartCol).Resize(1, 13).value = Array( _
        suno, "001", fvdt, uvdt, agnb, wsSource.Range("G2").value, _
        wsSource.Range("G4").value, wsSource.Range("G5").value, wsSource.Range("G6").value, _
        wsSource.Range("G7").value, 40, "USAPALI", 50)

    ' --- SECTION 2 & 3: LINES & STAGED PRICES ---
    Dim lineCount As Long: lineCount = 0
    Dim stageCount As Long: stageCount = 0
    
    Const MAX_STAGE_PAIRS As Integer = 30
    Dim qCols(29) As Long, pCols(29) As Long
    Dim k As Integer
    For k = 0 To MAX_STAGE_PAIRS - 1
        qCols(k) = 11 + k * 3
        pCols(k) = 13 + k * 3
    Next k

    Const MAX_LINES As Long = 30

    ' Load source data into memory once - avoids repeated cell reads inside the nested loop.
    ' Also pre-allocate output arrays so rows can be written in a single bulk operation.
    Dim srcData As Variant
    Dim lastSrcCol As Long
    lastSrcCol = wsSource.Cells(1, wsSource.Columns.count).End(xlToLeft).Column
    srcData = wsSource.Range(wsSource.Cells(1, 1), wsSource.Cells(lastRow, lastSrcCol)).value

    Dim arrLines() As Variant
    ReDim arrLines(1 To MAX_LINES, 1 To 8)
    Dim arrStages() As Variant
    ReDim arrStages(1 To MAX_LINES * MAX_STAGE_PAIRS, 1 To 7)

    For i = 2 To lastRow
        If lineCount >= MAX_LINES Then Exit For
        Dim itno As String: itno = CStr(srcData(i, 9))               ' Column I = 9
        Dim basePrice As Variant: basePrice = CleanPrice(srcData(i, 13)) ' Column M = 13

        If itno <> "" Then
            lineCount = lineCount + 1
            arrLines(lineCount, 1) = suno:  arrLines(lineCount, 2) = agnb
            arrLines(lineCount, 3) = 50:    arrLines(lineCount, 4) = itno
            arrLines(lineCount, 5) = fvdt:  arrLines(lineCount, 6) = basePrice
            arrLines(lineCount, 7) = uvdt:  arrLines(lineCount, 8) = 20

            For j = 0 To UBound(qCols)
                ' Guard against source data being narrower than the price tier columns
                If qCols(j) <= lastSrcCol And pCols(j) <= lastSrcCol Then
                    Dim qty As Variant: qty = srcData(i, qCols(j))
                    Dim price As Variant: price = CleanPrice(srcData(i, pCols(j)))

                    If IsNumeric(qty) And qty <> 0 And price <> 0 Then
                        stageCount = stageCount + 1
                        arrStages(stageCount, 1) = suno:  arrStages(stageCount, 2) = agnb
                        arrStages(stageCount, 3) = 50:    arrStages(stageCount, 4) = itno
                        arrStages(stageCount, 5) = fvdt:  arrStages(stageCount, 6) = qty
                        arrStages(stageCount, 7) = price
                    End If
                End If
            Next j
        End If
    Next i

    ' Write outputs in bulk rather than row-by-row
    If lineCount > 0 Then
        wsL.Cells(outStartRow + 1, outStartCol).Resize(lineCount, 8).value = arrLines
    End If
    If stageCount > 0 Then
        wsS.Cells(outStartRow + 1, outStartCol).Resize(stageCount, 7).value = arrStages
    End If

    ' 4. FINAL CLEANUP & SORTING
    ' Clear Row 7 as requested
    wsH.Rows(7).ClearContents
    wsL.Rows(7).ClearContents
    wsS.Rows(7).ClearContents

    ' Format PUPR columns
    wsL.Columns(outStartCol + 5).NumberFormat = "0.0000"
    wsS.Columns(outStartCol + 6).NumberFormat = "0.0000"
    
    ' Sort Staged Prices by Item (OBV1) then Qty (FRQT)
    If stageCount > 0 Then
        Dim sortRng As Range
        Set sortRng = wsS.Range(wsS.Cells(outStartRow, outStartCol), wsS.Cells(outStartRow + stageCount, outStartCol + 6))
        sortRng.Sort Key1:=wsS.Cells(outStartRow, outStartCol + 3), Order1:=xlAscending, _
                     Key2:=wsS.Cells(outStartRow, outStartCol + 5), Order2:=xlAscending, _
                     header:=xlYes
    End If

    ' Force full workbook calculation now that all sheets are populated
    Application.Calculate

    ' Autofit and Finish
    wsH.Activate: Call AutoFit_Click
    wsL.Activate: Call AutoFit_Click
    wsS.Activate: Call AutoFit_Click
    wsH.Activate

    MsgBox "PPS100MI Sheets Generated Successfully.", vbInformation, "Done"
    GeneratePPS100UploadSheets = True
End Function

' --- Private Helpers ---

Private Function CleanPrice(val As Variant) As Double
    Dim strVal As String
    strVal = Replace(CStr(val), "$", "")
    strVal = Replace(strVal, ",", "")
    If IsNumeric(strVal) Then CleanPrice = CDbl(strVal) Else CleanPrice = 0
End Function

Private Function SetupSheet(wb As Workbook, name As String, startRow As Long, startCol As Long, headers As Variant) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(name)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = wb.Worksheets.Add(After:=wb.Sheets(wb.Sheets.count))
        ws.name = name
    End If
    ws.Rows(startRow & ":" & ws.Rows.count).ClearContents
    ws.Cells(startRow, startCol).Resize(1, UBound(headers) + 1).value = headers
    Set SetupSheet = ws
End Function

Public Function RunAgreementProcess() As Boolean
    Dim isSuccess As Boolean

    Application.Calculation = xlCalculationManual

    Dim sheetsExist As Boolean
    sheetsExist = Settings_SheetExists("PPS100MI AddAgrHead") And _
                  Settings_SheetExists("PPS100MI AddAgrLine") And _
                  Settings_SheetExists("PPS100MI AddStgPrice")

    If Not sheetsExist Then
        ' Rename the trigger sheet to the first agreement sheet
        ActiveSheet.name = "PPS100MI AddAgrHead"
        With ActiveSheet
            .Range("A2").value = "PPS100MI": .Range("B2").value = "API": .Range("G4").value = "AddAgrHead"
        End With
        GetLayoutAll_Click

        ' Sheet 2
        Settings_NewSheet
        With ActiveSheet
            .Range("A2").value = "PPS100MI": .Range("B2").value = "API": .Range("G4").value = "AddAgrLine"
        End With
        GetLayoutAll_Click

        ' Sheet 3
        Settings_NewSheet
        With ActiveSheet
            .Range("A2").value = "PPS100MI": .Range("B2").value = "API": .Range("G4").value = "AddStgPrice"
        End With
        GetLayoutAll_Click
    End If

    isSuccess = GeneratePPS100UploadSheets()
    RunAgreementProcess = isSuccess
End Function


