Attribute VB_Name = "Xtra_PrelimJournal"
Public Function ModuleExists() As Boolean
    ModuleExists = True
End Function

Sub Confirm_Prelim_Journal_Click()
    Dim ws As Worksheet
    Dim startTime As Single
    Dim settings As ApiSettings
    Dim response As httpResponse
    Dim json As Object
    Dim cancelled As Boolean
    
    On Error GoTo ErrorHandler
    
    Application.EnableCancelKey = xlErrorHandler
    cancelled = False
    
    Set ws = ActiveSheet
    
    ' =========================================================================
    ' SETUP
    ' =========================================================================
    
    UI_UpdateVersion
    If ws.Range("Environment").value = "" Then Exit Sub
    
    SetFormulasAndFormatting_New ws
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    ' Freeze panes
    ws.Cells(1, 1).Select
    ActiveWindow.FreezePanes = False
    ws.Range("C9").Select
    ActiveWindow.FreezePanes = True
    Application.GoTo Reference:="R9C2", Scroll:=True
    
    startTime = Timer
    UI_ClearStatus
'    ClearLogSheet
    
    ' =========================================================================
    ' CUAM VALIDATION
    ' =========================================================================
    
    If Not ValidateCUAMTotal(ws) Then
        Application.Calculation = xlCalculationManual
        Exit Sub
    End If
    
    ' =========================================================================
    ' AUTHENTICATION (Doppio_Process pattern)
    ' =========================================================================
    
    settings = Config_ApiSettings
    If settings.maxbulk <= 1 Or settings.refreshSeconds = 0 Then
        Config_LoadSettingsFromSheet
        settings = Config_ApiSettings
    End If
    
    If activeEnvironment = ws.Range("I2").value Then
        ws.Range("J3").value = 2
    End If
    
    If m_s_AccessToken = "" Or _
       ws.Range("I3").value = "" Or _
       ws.Range("I4").value = "" Or _
       activeEnvironment <> ws.Range("I2").value Then
        Tenant_Token
    End If
    
    If ws.Range("I2").value = "Access requested" Or m_s_M3user = "" Then Exit Sub
    
    m_s_Company = ws.Range("Company").value
    m_s_Division = ws.Range("Division").value
    
    ' =========================================================================
    ' BUILD BATCH KEYS
    ' =========================================================================
    
    Dim key As String
    Dim BMIN As String
    Dim ms As String
    Dim interfaceName As String
    
    ms = Right("000" & CStr(CLng((Timer - Int(Timer)) * 1000)), 3)
    key = Format(Now, "JVyyyymmddhh") & ms
    interfaceName = ws.Range("interface").value
    
    ' Setup AutoFilter based on interface type
    SetupJournalFilter ws, interfaceName
    
    ' =========================================================================
    ' STEP 1: Create Batch Header (GLS840MI/AddBatchHead)
    ' =========================================================================
    
    UI_ShowPleaseWait "Creating Batch Header..."
    
    Dim batchBody As String
    batchBody = BuildBatchHeadBody(key, interfaceName, ws)
    
    response = ExecuteApiPost(m_s_MainUrl, m_s_MiPath, batchBody)
    
    If Not response.success Then
        MsgBox "Failed to create batch header: " & response.errorMessage, vbCritical
        GoTo Cleanup
    End If
    
    ' =========================================================================
    ' STEP 2: Process Batch Lines
    ' =========================================================================
    
    Dim strServiceName As String
    Dim strMethod As String
    Dim lastColumn As Integer
    Dim inputColumns() As Integer
    Dim inputFields() As String
    Dim inputValues() As String
    Dim fullBody As String
    Dim body As String
    Dim counter As Long
    Dim i As Long, j As Long
    
    strServiceName = ws.Range("API").value
    strMethod = ws.Range("Transaction").value
    
    If strServiceName = "" Or strMethod = "" Then
        MsgBox "Please enter API and Transaction.", vbExclamation
        GoTo Cleanup
    End If
    
    lastColumn = ws.Cells(8, ws.Columns.count).End(xlToLeft).Column
    m_l_Row = 9
    
    ' Initialize field arrays
    ReDim inputColumns(1 To 400)
    ReDim inputFields(1 To 400)
    ReDim inputValues(1 To 400)
    
    j = 0
    For i = 1 To lastColumn
        If ws.Cells(m_l_Row - 1, i + 1) <> "" Then
            j = j + 1
            If j > 400 Then Exit For
            inputColumns(j) = i + 1
            inputFields(j) = ws.Cells(m_l_Row - 1, i + 1)
        End If
    Next i
    
    ws.Calculate
    
    If IsEmpty(ws.Range("B9").value) Then
        ws.Range("B9").value = "?"
    End If
    
    fullBody = ""
    counter = 0
    
    While ws.Cells(m_l_Row, 2).value <> ""
        DoEvents  ' Allow ESC key detection
        
        ' Collect values with field overrides
        For i = 1 To lastColumn
            If inputColumns(i) <> 0 Then
                Dim cellValue As Variant
                cellValue = ws.Cells(m_l_Row, inputColumns(i))
                
                If isError(cellValue) Then cellValue = ""
                If cellValue = "?" And m_l_Row = 9 And i = 1 Then cellValue = ""
                
                ' Journal-specific field overrides
                cellValue = ApplyFieldOverride(inputFields(i), cellValue, key, m_s_Division)
                
                inputValues(i) = CStr(cellValue)
            End If
        Next i
        
'        body = BuildMITransactionBody(strMethod, inputFields, inputValues)
        
        If fullBody = "" Then
            fullBody = body
        Else
            fullBody = fullBody & "," & body
        End If
        
        m_l_Row = m_l_Row + 1
        counter = counter + 1
        
        If counter >= settings.maxbulk Then
            body = "{""program"":""" & strServiceName & """,""transactions"":[" & fullBody & "]}"
            fullBody = ""
            
            UI_ShowPleaseWait "Please Wait... Calling API (batch " & _
                Int(m_l_Row / settings.maxbulk) & ") - Press ESC to cancel"
            
            response = ExecuteApiPost(m_s_MainUrl, m_s_MiPath, body)
            ParseAndWriteResults response, ws, m_l_Row - counter
            
            counter = 0
        End If
    Wend
    
    ' Flush remaining rows
    If counter > 0 Then
        m_l_Row = m_l_Row + (settings.maxbulk - counter)
        body = "{""program"":""" & strServiceName & """,""transactions"":[" & fullBody & "]}"
        
        UI_ShowPleaseWait "Please Wait... Calling API (final batch)"
        
        response = ExecuteApiPost(m_s_MainUrl, m_s_MiPath, body)
        ParseAndWriteResults response, ws, m_l_Row - settings.maxbulk
    End If
    
    UI_DisplayElapsedTime startTime, ws
    
    ' =========================================================================
    ' STEP 3: Initialize MBM (MNS260MI/AddMBMInit)
    ' =========================================================================
    
    UI_ShowPleaseWait "Initializing MBM..."
    
    Dim mbmBody As String
    mbmBody = BuildAddMBMInitBody(key, interfaceName)
    
    response = ExecuteApiPost(m_s_MainUrl, m_s_MiPath, mbmBody)
    
    If response.success Then
        BMIN = ExtractBMIN(response)
    End If
    
    If BMIN = "" Then
        MsgBox "Failed to get BMIN from MNS260MI/AddMBMInit.", vbCritical
        GoTo Cleanup
    End If
    
    ' =========================================================================
    ' STEP 4: Process MBM (MNS260MI/PrcMBMInit)
    ' =========================================================================
    
    UI_ShowPleaseWait "Processing MBM..."
    
    Dim prcBody As String
    prcBody = "{""program"":""MNS260MI"",""transactions"":[{""transaction"":""PrcMBMInit"",""record"":{""BMIN"":""" & BMIN & """}}]}"
    
    response = ExecuteApiPost(m_s_MainUrl, m_s_MiPath, prcBody)
    
    If Not response.success Then
        MsgBox "Failed to process MBM: " & response.errorMessage, vbCritical
    End If

Cleanup:
    Application.EnableCancelKey = xlInterrupt
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    UI_KillPleaseWait
    DoEvents
    m_l_Row = 9
    
    If cancelled Then
        ws.Range("A" & m_l_Row).value = "CANCELLED"
        ws.Range("A" & m_l_Row).Font.Color = COLOR_ERROR
        MsgBox "Process cancelled by user at row " & m_l_Row & "." & vbCrLf & _
               "Rows processed before cancellation are valid.", vbExclamation, "Cancelled"
    End If
    Exit Sub
    
ErrorHandler:
    If Err.Number = 18 Then
        cancelled = True
        
        ' Flush pending batch
        If counter > 0 Then
            body = "{""program"":""" & strServiceName & """,""transactions"":[" & fullBody & "]}"
            On Error Resume Next
            response = ExecuteApiPost(m_s_MainUrl, m_s_MiPath, body)
            ParseAndWriteResults response, ws, m_l_Row - counter
            On Error GoTo 0
        End If
        
        Resume Cleanup
    End If
    
    Application.EnableCancelKey = xlInterrupt
    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = True
    UI_KillPleaseWait
    MsgBox "Error in Confirm_Prelim_Journal: " & Err.description, vbCritical
    m_l_Row = 9
End Sub


' =============================================================================
' VALIDATION
' =============================================================================

Private Function ValidateCUAMTotal(ws As Worksheet) As Boolean
    Dim lastColumn As Long
    Dim lastRow As Long
    Dim cuamColumn As Long
    Dim cuamTotal As Double
    Dim i As Long
    
    ValidateCUAMTotal = True
    
    lastColumn = ws.Cells(8, ws.Columns.count).End(xlToLeft).Column
    
    ' Find CUAM column
    cuamColumn = 0
    For i = 1 To lastColumn
        If ws.Cells(8, i).value = "CUAM" Then
            cuamColumn = i
            Exit For
        End If
    Next i
    
    If cuamColumn = 0 Then Exit Function
    
    lastRow = ws.Cells(ws.Rows.count, 2).End(xlUp).row
    
    cuamTotal = 0
    For i = 9 To lastRow
        If ws.Cells(i, 2).value <> "" Then
            If IsNumeric(ws.Cells(i, cuamColumn).value) Then
                cuamTotal = cuamTotal + CDbl(ws.Cells(i, cuamColumn).value)
            End If
        End If
    Next i
    
    If Abs(cuamTotal) > 0.01 Then
        MsgBox "CUAM column must total to 0. Current total: " & Format(cuamTotal, "#,##0.00"), _
               vbCritical, "Validation Error"
        ValidateCUAMTotal = False
    End If
End Function


' =============================================================================
' JSON BODY BUILDERS
' =============================================================================

Private Function BuildBatchHeadBody(key As String, interfaceName As String, ws As Worksheet) As String
    Dim desc As String
    
    If interfaceName = "PRELJOURNAL" Then
        desc = "Journal Excel - " & Format(Now, "yyyymmdd")
    Else
        desc = "Auto Reversal Excel Jrnl-" & Format(Now, "yyyymmdd")
    End If
    
    BuildBatchHeadBody = "{""program"":""GLS840MI"",""transactions"":[{" & _
        """transaction"":""AddBatchHead"",""record"":{" & _
        """CONO"":""" & m_s_Company & """," & _
        """KEY1"":""" & key & """," & _
        """INTN"":""" & interfaceName & """," & _
        """DESC"":""" & desc & """," & _
        """USID"":""" & m_s_M3user & """," & _
        """DIVI"":""" & m_s_Division & """" & _
        "}}]}"
End Function


Private Function BuildAddMBMInitBody(key As String, interfaceName As String) As String
    BuildAddMBMInitBody = "{""program"":""MNS260MI"",""transactions"":[{" & _
        """transaction"":""AddMBMInit"",""record"":{" & _
        """DONR"":""MBM""," & _
        """PRF1"":""JATOOL"",""PRTF"":""JATOOL""," & _
        """OBJC"":""KEY1"",""DONO"":""" & key & """," & _
        """ARSD"":""0"",""PRF2"":""JATOOL"",""CPPL"":""0""," & _
        """MKF4"":""USID"",""MKV4"":""" & m_s_M3user & """," & _
        """MKF5"":""INTN"",""MKV5"":""" & interfaceName & """," & _
        """MKF6"":""CTROPT"",""MKV6"":""False""" & _
        "}}]}"
End Function


' =============================================================================
' FIELD OVERRIDES
' =============================================================================

Private Function ApplyFieldOverride(fieldName As String, cellValue As Variant, _
                                     key As String, division As String) As Variant
    Select Case fieldName
        Case "KEY1":  ApplyFieldOverride = key
        Case "INRI":  ApplyFieldOverride = "I1"
        Case "DIVI":  ApplyFieldOverride = division
        Case "GRNR":  ApplyFieldOverride = "123456"
        Case "EICD":  ApplyFieldOverride = "0"
        Case "EXN1":  ApplyFieldOverride = "001"
        Case "EXI1":  ApplyFieldOverride = key
        Case Else:    ApplyFieldOverride = cellValue
    End Select
End Function


' =============================================================================
' AUTOFILTER SETUP
' =============================================================================

Private Sub SetupJournalFilter(ws As Worksheet, interfaceName As String)
    If interfaceName = "PRELJOURNAL" Then
        If Not ws.AutoFilterMode Then
            ws.Range("A8:U8").AutoFilter
        ElseIf ws.FilterMode Then
            ws.ShowAllData
            ws.Range("A8:U8").AutoFilter
        End If
        ws.Range("V8").value = ""
    Else
        If Not ws.AutoFilterMode Then
            ws.Range("A8:V8").AutoFilter
        ElseIf ws.FilterMode Then
            ws.ShowAllData
            ws.Range("A8:V8").AutoFilter
        End If
        ws.Range("V8").value = "SHDT"
    End If
End Sub


' =============================================================================
' RESPONSE PARSING
' =============================================================================

Private Function ExtractBMIN(response As httpResponse) As String
    Dim json As Object
    Dim records As Object
    Dim record As Object
    
    On Error Resume Next
    
    If Len(response.body) = 0 Then Exit Function
    
    Set json = ParseJson(response.body)
    If json Is Nothing Then Exit Function
    
    Set records = json.item("results")(1).item("records")
    If records Is Nothing Then Exit Function
    
    For Each record In records
        ExtractBMIN = record.item("BMIN")
        If ExtractBMIN <> "" Then Exit Function
    Next record
    
    On Error GoTo 0
End Function


Private Sub ParseAndWriteResults(response As httpResponse, ws As Worksheet, startRow As Long)
    Dim json As Object
    Dim results As Object
    Dim resultItem As Object
    Dim record As Object
    Dim rowNum As Long
    Dim statusMsg As String
    
    On Error Resume Next
    
    If Not response.success Or Len(response.body) = 0 Then Exit Sub
    
    Set json = ParseJson(response.body)
    If json Is Nothing Then Exit Sub
    
    Set results = json.item("results")
    If results Is Nothing Then Exit Sub
    
    rowNum = startRow
    For Each resultItem In results
        statusMsg = ""
        If Not resultItem Is Nothing Then
            statusMsg = resultItem.item("errorMessage")
            If statusMsg = "" Then
                ws.Cells(rowNum, 1).value = "OK"
            Else
                ws.Cells(rowNum, 1).value = "NOK " & statusMsg
                ws.Cells(rowNum, 1).Font.Color = COLOR_ERROR
            End If
        End If
        rowNum = rowNum + 1
    Next resultItem
    
    On Error GoTo 0
End Sub

Sub PreliminaryPrep()
    Dim ws As Worksheet
    Dim shp As Shape
    Dim btn1 As Button
    Dim arrRow7 As Variant
    Dim arrRow8 As Variant
    
    Settings_NewSheet
    
    ' Set the target to the currently active sheet
    Set ws = ActiveSheet
    
    Application.ScreenUpdating = False
    
    ' =====================================================================
    ' 1. SET SPECIFIC CELL VALUES
    ' =====================================================================
    ws.Range("A2").value = "GLS840MI"
    ws.Range("C4").value = "Interface:  "
    ws.Range("C4").HorizontalAlignment = xlRight
    ws.Range("G4").value = "AddBatchLineFld"
    GetLayout_Click False
    
    Application.ScreenUpdating = False
    
    ' =====================================================================
    ' 2. CREATE COMBOBOX (DATA VALIDATION) IN D4 AND NAME THE RANGE
    ' =====================================================================
    With ws.Range("D4").Validation
        .Delete ' Clear any existing validation
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:="PRELJOURNAL,ARPLJOURNAL"
        .IgnoreBlank = True
        .InCellDropdown = True
    End With
    
    ws.Range("D4").value = "PRELJOURNAL" ' Set default value
    
    ' Change the background color of D4 to Rose
    ws.Range("D4").Interior.Color = RGB(208, 153, 150)
    
    ' Name the cell "interface" (Creates a Named Range)
    ws.names.Add name:="interface", RefersTo:=ws.Range("D4")
    
    ' =====================================================================
    ' 3. REMOVE STANDARD BUTTONS
    ' =====================================================================
    ' Loops through all shapes and removes any Form Control Buttons
    For Each shp In ws.Shapes
        If shp.Type = msoFormControl Then
            If shp.FormControlType = xlButtonControl Then
                shp.Delete
            End If
        End If
    Next shp
    
    ' =====================================================================
    ' 4. ADD NEW "CONFIRM JOURNAL" BUTTON
    ' =====================================================================
    Set btn1 = ws.Buttons.Add(10, 85, 131, 45)
    btn1.Caption = "Confirm Journal"
    btn1.OnAction = "Confirm_Prelim_Journal_Click"
    
    ' =====================================================================
    ' 5. POPULATE ROWS 7 & 8 STARTING IN COLUMN 2 (COLUMN B)
    ' =====================================================================
    ws.Rows("7:8").ClearContents
    arrRow8 = Array("LINE", "KEY1", "AIT1", "AIT2", "AIT3", "AIT4", "AIT5", "AIT6", "AIT7", "ACQT", _
                    "CUCD", "CUAM", "DBCR", "ACDT", "VTXT", "EXN1", "EXI1", "EXN2", "EXI2", "RERF", _
                    "", "INRI", "DIVI", "GRNR", "EICD")
        
    ' Output arrays starting at B7 and B8 (which is Column 2)
    ws.Range("B8").Resize(1, 25).value = arrRow8
    
    AutoFit_Click
    AutoFit_ColumnsAndRows False, False
    Application.ScreenUpdating = True
End Sub

