Attribute VB_Name = "Xtra_PriceList"
Public Function ModuleExists() As Boolean
    ModuleExists = True
End Function

Public Function GeneratePriceUploadSheetsInteractive() As Boolean
    Dim rngSelect As Range
    Dim wsSource As Worksheet
    Dim wbTarget As Workbook
    Dim outStartRow As Long
    Dim outStartCol As Long
    
    ' =====================================================================
    ' 1. INTERACTIVE PROMPTS
    ' =====================================================================
    
    ' Prompt user to click a cell to identify the Workbook and Sheet
    On Error Resume Next
    Set rngSelect = Application.InputBox( _
        Prompt:="Please click on ANY CELL within the SOURCE WORKBOOK and SHEET you want to use." & vbCrLf & vbCrLf & _
                "(You can use your mouse to navigate to the correct sheet and click a cell)", _
        title:="Select Source Data", _
        Type:=8)
    On Error GoTo 0
    
    ' Exit if user clicked Cancel
    If rngSelect Is Nothing Then
        MsgBox "Operation cancelled.", vbExclamation, "Cancelled"
        GeneratePriceUploadSheetsInteractive = False ' <-- Tell the caller we cancelled
        Exit Function ' <-- Changed from Exit Sub
    End If
    
    ' Extract the Worksheet from the user's click
    Set wsSource = rngSelect.Parent
    
    ' Set the target workbook to the CURRENTLY ACTIVE WORKBOOK
    Set wbTarget = ActiveWorkbook
    
    ' --- HARDCODED OUTPUT LOCATION ---
    outStartRow = 8    ' Start on Row 8
    outStartCol = 2    ' Start on Column B (2)
    
    ' =====================================================================
    ' 2. PREPARE DESTINATION SHEETS (USE EXISTING)
    ' =====================================================================
    Dim wsPriceList As Worksheet
    Dim wsBase As Worksheet
    Dim wsTiered As Worksheet
    Dim lastRow As Long

    lastRow = wsSource.Cells(wsSource.Rows.count, "A").End(xlUp).row

    ' Try to link to the existing AddPriceList sheet, create if it doesn't exist
    On Error Resume Next
    Set wsPriceList = wbTarget.Worksheets("OIS017MI AddPriceList")
    On Error GoTo 0
    If wsPriceList Is Nothing Then
        Set wsPriceList = wbTarget.Worksheets.Add(After:=wbTarget.Sheets(wbTarget.Sheets.count))
        wsPriceList.name = "OIS017MI AddPriceList"
    End If

    ' Try to link to the existing Base Price sheet, create if it doesn't exist
    On Error Resume Next
    Set wsBase = wbTarget.Worksheets("OIS017MI AddBasePrice")
    On Error GoTo 0
    If wsBase Is Nothing Then
        Set wsBase = wbTarget.Worksheets.Add(After:=wbTarget.Sheets(wbTarget.Sheets.count))
        wsBase.name = "OIS017MI AddBasePrice"
    End If

    ' Try to link to the existing Tiered Price sheet, create if it doesn't exist
    On Error Resume Next
    Set wsTiered = wbTarget.Worksheets("OIS017MI AddGradSlsPrc")
    On Error GoTo 0
    If wsTiered Is Nothing Then
        Set wsTiered = wbTarget.Worksheets.Add(After:=wbTarget.Sheets(wbTarget.Sheets.count))
        wsTiered.name = "OIS017MI AddGradSlsPrc"
    End If

    ' Clear out any old data from row 8 downwards so leftover rows don't mix with new data
    wsPriceList.Rows(outStartRow & ":" & wsPriceList.Rows.count).ClearContents
    wsBase.Rows(outStartRow & ":" & wsBase.Rows.count).ClearContents
    wsTiered.Rows(outStartRow & ":" & wsTiered.Rows.count).ClearContents

    ' Set up headers at the hardcoded locations (Row 8, Column B)
    wsPriceList.Cells(outStartRow, outStartCol).Resize(1, 11).value = Array("PRRF", "CUCD", "FVDT", "LVDT", "TX40", "TX15", "SCMO", "SCMU", "CRTP", "WHLO", "PCTP")
    wsBase.Cells(outStartRow, outStartCol).Resize(1, 6).value = Array("PRRF", "CUCD", "FVDT", "ITNO", "SAPR", "VFDT")
    wsTiered.Cells(outStartRow, outStartCol).Resize(1, 7).value = Array("PRRF", "CUCD", "FVDT", "ITNO", "QTYL", "SAPR", "VFDT")
    
    ' =====================================================================
    ' 3. MAP COLUMNS DYNAMICALLY
    ' =====================================================================
    Dim colPRRF As Integer, colITNO As Integer, colFallbackPrice As Integer
    Dim colValidFrom As Integer ' Added variable for the new date column
    Dim qlCols(1 To 15) As Integer
    Dim drCols(1 To 15) As Integer
    Dim c As Integer
    Dim header As String, num As Integer
    
    For c = 1 To wsSource.Cells(1, wsSource.Columns.count).End(xlToLeft).Column
        header = UCase(Trim(Replace(wsSource.Cells(1, c).value, "_", " ")))
        
        If header = "PRICE LIST" Then colPRRF = c
        If header = "ITEM" Then colITNO = c
        If header = "LINE VALID FROM" Then colValidFrom = c ' Map the date column
        
        If header = "SALES PRICE" Then
            colFallbackPrice = c
        ElseIf header = "COST" And colFallbackPrice = 0 Then
            colFallbackPrice = c
        End If
        
        If header Like "QUALIFYING LIMIT *" Then
            num = val(Replace(header, "QUALIFYING LIMIT ", ""))
            If num >= 1 And num <= 15 Then qlCols(num) = c
        End If
        
        If header Like "DISC RATE *" Then
            num = val(Replace(header, "DISC RATE ", ""))
            If num >= 1 And num <= 15 Then drCols(num) = c
        End If
    Next c
    
    ' Error check updated to ensure LINE VALID FROM is found
    If colPRRF = 0 Or colITNO = 0 Or colValidFrom = 0 Then
        MsgBox "Error: Could not find 'PRICE LIST', 'ITEM', or 'LINE VALID FROM' columns on the selected sheet.", vbCritical
        Exit Function
    End If

    ' =====================================================================
    ' 4. PROCESS DATA ARRAYS
    ' =====================================================================
    Dim arrPriceList() As Variant
    ReDim arrPriceList(1 To lastRow, 1 To 11)
    Dim priceListCount As Long: priceListCount = 0
    Dim seenPriceLists As New Collection   ' tracks unique PRRF|FVDT combos

    Dim arrBase() As Variant
    ReDim arrBase(1 To lastRow, 1 To 6)
    Dim baseCount As Long: baseCount = 0

    Dim arrTiered() As Variant
    ReDim arrTiered(1 To lastRow * 15, 1 To 7)
    Dim tieredCount As Long: tieredCount = 0

    Dim prrf As String, itno As String, validFromDate As String
    Dim ql As String, dr As String
    Dim i As Long, j As Long
    Dim srcData As Variant
    Dim lastSrcCol As Long

    ' Load entire source range into memory once - avoids repeated cell-by-cell reads
    ' inside the loop, which is the primary performance bottleneck for large sheets.
    lastSrcCol = wsSource.Cells(1, wsSource.Columns.count).End(xlToLeft).Column
    srcData = wsSource.Range(wsSource.Cells(1, 1), wsSource.Cells(lastRow, lastSrcCol)).value

    For i = 2 To lastRow
        prrf = CStr(srcData(i, colPRRF))
        itno = CStr(srcData(i, colITNO))

        ' Read the date from the column and format to YYYYMMDD
        validFromDate = Format(srcData(i, colValidFrom), "yyyymmdd")

        ' AddPriceList Logic — one row per unique (PRRF, FVDT) combination
        Dim plKey As String
        plKey = prrf & "|" & validFromDate
        On Error Resume Next
        seenPriceLists.Add plKey, plKey   ' errors if key already exists
        If Err.Number = 0 Then
            Dim lvdt As String
            lvdt = Left(validFromDate, 4) & "1231"
            priceListCount = priceListCount + 1
            arrPriceList(priceListCount, 1) = prrf
            arrPriceList(priceListCount, 2) = "USD"
            arrPriceList(priceListCount, 3) = validFromDate
            arrPriceList(priceListCount, 4) = lvdt
            arrPriceList(priceListCount, 5) = prrf & " Price List"
            arrPriceList(priceListCount, 6) = prrf & " Price List"
            arrPriceList(priceListCount, 7) = "COST"
            arrPriceList(priceListCount, 8) = "1"
            arrPriceList(priceListCount, 9) = "1"
            arrPriceList(priceListCount, 10) = "US1"
            arrPriceList(priceListCount, 11) = "3"
        End If
        On Error GoTo 0

        ' AddBasePrice Logic
        If qlCols(1) > 0 And drCols(1) > 0 Then
            ql = Trim(CStr(srcData(i, qlCols(1))))
            dr = Trim(CStr(srcData(i, drCols(1))))

            If ql <> "" Then
                baseCount = baseCount + 1
                arrBase(baseCount, 1) = prrf
                arrBase(baseCount, 2) = "USD"
                arrBase(baseCount, 3) = validFromDate
                arrBase(baseCount, 4) = itno

                If val(dr) = 0 And colFallbackPrice > 0 Then
                    arrBase(baseCount, 5) = srcData(i, colFallbackPrice)
                Else
                    arrBase(baseCount, 5) = CDbl(dr)
                End If
                arrBase(baseCount, 6) = validFromDate
            End If
        End If

        ' AddGradSlsPrc (Tiered Price) Logic
        For j = 1 To 15
            If qlCols(j) > 0 And drCols(j) > 0 Then
                ql = Trim(CStr(srcData(i, qlCols(j))))
                dr = Trim(CStr(srcData(i, drCols(j))))

                If ql <> "" Then
                    If IsNumeric(dr) Then
                        If CDbl(dr) > 0 Then
                            tieredCount = tieredCount + 1
                            arrTiered(tieredCount, 1) = prrf
                            arrTiered(tieredCount, 2) = "USD"
                            arrTiered(tieredCount, 3) = validFromDate
                            arrTiered(tieredCount, 4) = itno
                            arrTiered(tieredCount, 5) = CLng(ql)
                            arrTiered(tieredCount, 6) = CDbl(dr)
                            arrTiered(tieredCount, 7) = validFromDate
                        End If
                    End If
                End If
            End If
        Next j
    Next i
    
    ' =====================================================================
    ' 5. OUTPUT AND SORT DATA
    ' =====================================================================
    If priceListCount > 0 Then
        wsPriceList.Cells(outStartRow + 1, outStartCol).Resize(priceListCount, 11).value = arrPriceList
    End If

    If baseCount > 0 Then
        wsBase.Cells(outStartRow + 1, outStartCol).Resize(baseCount, 6).value = arrBase
    End If
    
    If tieredCount > 0 Then
        wsTiered.Cells(outStartRow + 1, outStartCol).Resize(tieredCount, 7).value = arrTiered
        
        Dim sortRng As Range
        Set sortRng = wsTiered.Range(wsTiered.Cells(outStartRow, outStartCol), _
                                     wsTiered.Cells(outStartRow + tieredCount, outStartCol + 6))
        
        sortRng.Sort _
            Key1:=wsTiered.Cells(outStartRow, outStartCol), Order1:=xlAscending, _
            Key2:=wsTiered.Cells(outStartRow, outStartCol + 3), Order2:=xlAscending, _
            Key3:=wsTiered.Cells(outStartRow, outStartCol + 4), Order3:=xlAscending, _
            header:=xlYes, MatchCase:=False, Orientation:=xlTopToBottom
    End If
    
    ' =====================================================================
    ' 6. FINAL CLEANUP & FORMATTING
    ' =====================================================================
    ' Clear row 7 on all three sheets
    wsPriceList.Rows(7).ClearContents
    wsBase.Rows(7).ClearContents
    wsTiered.Rows(7).ClearContents

    ' Autofit all three sheets (activate each in turn so AutoFit_Click works)
    wsPriceList.Activate
    Call AutoFit_Click

    wsBase.Activate
    Call AutoFit_Click

    wsTiered.Activate
    Call AutoFit_Click

    ' Return user to AddPriceList — it's first in the API call sequence
    wsPriceList.Activate
    
    MsgBox "Price lists updated successfully based on the source dates.", vbInformation, "Done"
    
    GeneratePriceUploadSheetsInteractive = True
End Function

''
' Ensures the three standard OIS017MI sheets exist (AddPriceList,
' AddBasePrice, AddGradSlsPrc), then launches
' GeneratePriceUploadSheetsInteractive to populate them.
'
' Sheet creation: only creates sheets that are missing — existing sheets
' keep their layout and just have their data replaced.
'
' Deletion: the original sheet is deleted only when (a) the user completed
' the interactive prompt AND (b) the active sheet is NOT one of the three
' price list sheets (so re-running from an existing price list tab is safe).
''
Public Sub SetupPriceListSheets()
    Dim wsOriginal As Worksheet
    Dim isSuccess As Boolean
    Dim wbThis As Workbook

    Set wbThis = ThisWorkbook
    Set wsOriginal = ActiveSheet

    Const SHEET_PL As String = "OIS017MI AddPriceList"
    Const SHEET_BP As String = "OIS017MI AddBasePrice"
    Const SHEET_GS As String = "OIS017MI AddGradSlsPrc"

    ' ----------------------------------------------------------------
    ' Detect which sheets are already present
    ' ----------------------------------------------------------------
    Dim wsPL As Worksheet, wsBP As Worksheet, wsGS As Worksheet
    On Error Resume Next
    Set wsPL = wbThis.Worksheets(SHEET_PL)
    Set wsBP = wbThis.Worksheets(SHEET_BP)
    Set wsGS = wbThis.Worksheets(SHEET_GS)
    On Error GoTo 0

    ' ----------------------------------------------------------------
    ' Force naming = 0 (API + Transaction) so sheets get the correct
    ' names regardless of the user's current Settings naming value
    ' ----------------------------------------------------------------
    Dim savedNaming As Integer
    savedNaming = sheetNaming
    sheetNaming = 0

    ' ----------------------------------------------------------------
    ' Create any that are missing (preserve existing ones as-is)
    ' ----------------------------------------------------------------
    If wsPL Is Nothing Then
        Settings_NewSheet
        With ActiveSheet
            .Range("A2").value = "OIS017MI"
            .Range("B2").value = "API"
            .Range("G4").value = "AddPriceList"
        End With
        GetLayoutAll_Click
    End If

    If wsBP Is Nothing Then
        Settings_NewSheet
        With ActiveSheet
            .Range("A2").value = "OIS017MI"
            .Range("B2").value = "API"
            .Range("G4").value = "AddBasePrice"
        End With
        GetLayoutAll_Click
    End If

    If wsGS Is Nothing Then
        Settings_NewSheet
        With ActiveSheet
            .Range("A2").value = "OIS017MI"
            .Range("B2").value = "API"
            .Range("G4").value = "AddGradSlsPrc"
        End With
        GetLayoutAll_Click
    End If

    ' Restore original naming setting
    sheetNaming = savedNaming

    ' ----------------------------------------------------------------
    ' Populate data
    ' ----------------------------------------------------------------
    isSuccess = GeneratePriceUploadSheetsInteractive

    ' ----------------------------------------------------------------
    ' Delete the originating sheet only when the user finished the
    ' prompt AND it isn't one of the three price list sheets itself
    ' ----------------------------------------------------------------
    If isSuccess Then
        Dim origName As String
        origName = wsOriginal.name
        If origName <> SHEET_PL And origName <> SHEET_BP And origName <> SHEET_GS Then
            Application.DisplayAlerts = False
            wsOriginal.Delete
            Application.DisplayAlerts = True
        End If
    End If
End Sub


