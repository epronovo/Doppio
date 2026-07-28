Attribute VB_Name = "Xtra_NPI"
''
' Xtra_NPI - New Product Introduction Process Module
' Executes a hardcoded sequence of MI transactions for NPI workflows.
'
' Auth pattern: Tenant_Token -> SyncConfigFromDoppioModule (same as Doppio_Process)
' API pattern:  ExecuteRequest via m_s_* global auth state
'
' @module Xtra_NPI
' @author Doppio Group - eric@doppiogroup.com
' @version 2.0
''
Option Explicit

' Module-level state
Private m_CurrentRow As Long
Private m_Worksheet As Worksheet

' NPI API/Transaction sequence (0-based)
Private m_ApiList(0 To 3) As String
Private m_TrxList(0 To 3) As String

' IPS steps (last steps � SOAP/XML, executed after all MI steps)
Private m_ips_ServiceList() As String
Private m_ips_MethodList() As String

Public Function ModuleExists() As Boolean
    ModuleExists = True
End Function

' =============================================================================
' HEADER SETUP
' =============================================================================

''
' Sets up rows 7 and 8 on the active sheet to match the NPI Template layout.
'
' Row 7: field description + length constraint (two lines, wrapped)
' Row 8: M3 MI field codes (ITNO, STAT, CITN, ...)
'
' Colors use theme-based fills matching the source template:
'   Row 7 B-N : Accent1, tint  0.0  (medium blue)
'   Row 8 A-N : Accent1, tint -0.5  (darker blue)
'   Row 7 A   : Accent3, tint  0.8  (light gray)
''
Public Sub SetupNPISheet()
    Dim ws As Worksheet
    Dim wsOriginal As Worksheet
    Dim i As Integer
    Dim col As Integer
    Dim LF As String
    Dim r7Vals(0 To 35) As String
    Dim r8Vals(0 To 35) As String
    Dim isRefresh As Boolean

    ' If already on the NPI Load Sheet reuse it (values-only refresh).
    ' Otherwise create a fresh sheet from Master.
    isRefresh = (ActiveSheet.name = "NPI Load Sheet")
    Set wsOriginal = ActiveSheet
    UI_ShowPleaseWait "Building NPI Sheet..."

    If isRefresh Then
        Set ws = ActiveSheet
    Else
        ' Create a new sheet from Master (sets up rows 1-6, named ranges, etc.)
        Settings_NewSheet
        Set ws = ActiveSheet
    End If
    LF = Chr(10)

    ' -----------------------------------------------------------------------
    ' Field data 36 fields, columns B (2) through AK (37)
    ' Row 7: "Description" & LF & "Constraint"   Row 8: MI field code
    ' Array() cannot hold expressions, so individual assignment is used.
    ' -----------------------------------------------------------------------
    r7Vals(0) = "Item number" & LF & "A15"
    r7Vals(1) = "Status" & LF & "A2"
    r7Vals(2) = "Copy Item number" & LF & "A15"
    r7Vals(3) = "Name" & LF & "A60"
    r7Vals(4) = "Description 2" & LF & "A120"
    r7Vals(5) = "Item group" & LF & "A8"
    r7Vals(6) = "Product group" & LF & "A5"
    r7Vals(7) = "Net weight" & LF & "N17"
    r7Vals(8) = "User-defined field 1 - item" & LF & "A10"
    r7Vals(9) = "Warehouse" & LF & "A3"
    r7Vals(10) = "Copy Warehouse" & LF & "A3"
    r7Vals(11) = "Supplier number" & LF & "A10"
    r7Vals(12) = "Safety stock" & LF & "N17"
    r7Vals(13) = "Minimum order quantity" & LF & "N17"
    r7Vals(14) = "Acquisition code" & LF & "N1"
    r7Vals(15) = "Order type" & LF & "A3"
    r7Vals(16) = "ABC class - manual" & LF & "A1"
    r7Vals(17) = "Physical inventory cycle" & LF & "A3"
    r7Vals(18) = "Economical order quantity days" & LF & "N3"
    r7Vals(19) = "Annual demand" & LF & "N17"
    r7Vals(20) = "Order multiple" & LF & "N17"
    r7Vals(21) = "Product line" & LF & "A10"
    r7Vals(22) = "Issue multiple" & LF & "N17"
    r7Vals(23) = "Facility" & LF & "A3"
    r7Vals(24) = "Average cost" & LF & "N19"
    r7Vals(25) = "On-hand balance method - facility" & LF & "N1"
    r7Vals(26) = "Customs statistical number" & LF & "A16"
    r7Vals(27) = "Sales price" & LF & "N19"
    r7Vals(28) = "Currency - sales price" & LF & "A3"
    r7Vals(29) = "Goods receiving method" & LF & "A3"
    r7Vals(30) = "Order multiple" & LF & "N17"
    r7Vals(31) = "Lowest quality inspection method" & LF & "A1"
    r7Vals(32) = "Country of origin" & LF & "A3"
    r7Vals(33) = "Responsible" & LF & "A10"
    r7Vals(34) = "Buyer" & LF & "A10"
    r7Vals(35) = "Supply lead time" & LF & "N3"

    r8Vals(0) = "ITNO": r8Vals(1) = "STAT": r8Vals(2) = "CITN": r8Vals(3) = "ITDS"
    r8Vals(4) = "FUDS": r8Vals(5) = "ITGR": r8Vals(6) = "ITCL": r8Vals(7) = "NEWE"
    r8Vals(8) = "CFI1": r8Vals(9) = "WHLO": r8Vals(10) = "CWHL": r8Vals(11) = "SUNO"
    r8Vals(12) = "SSQT": r8Vals(13) = "LOQT": r8Vals(14) = "PUIT": r8Vals(15) = "ORTY"
    r8Vals(16) = "MABC": r8Vals(17) = "INCD": r8Vals(18) = "EQDA": r8Vals(19) = "YEQT"
    r8Vals(20) = "UNMU": r8Vals(21) = "PDLN": r8Vals(22) = "TOMU": r8Vals(23) = "FACI"
    r8Vals(24) = "APPR": r8Vals(25) = "FATM": r8Vals(26) = "CSNO": r8Vals(27) = "SAPR"
    r8Vals(28) = "CUCS": r8Vals(29) = "GRMT": r8Vals(30) = "UNMU": r8Vals(31) = "LCLV"
    r8Vals(32) = "ORCO": r8Vals(33) = "RESP": r8Vals(34) = "BUYE": r8Vals(35) = "LEA1"

    ' -----------------------------------------------------------------------
    ' Row 7 field descriptions
    ' -----------------------------------------------------------------------

    ws.Rows(7).RowHeight = 54

    ' Apply A7 keyword-cell format (matches style 30 on all other sheets)
    With ws.Cells(7, 1)
        .value = ""
        .Font.name = "Avenir Book"
        .Font.Size = 12
        .Font.Bold = True
        .Font.Italic = True
        .Interior.ThemeColor = xlThemeColorAccent3
        .Interior.TintAndShade = 0.799981688894314
        .HorizontalAlignment = xlLeft
        .WrapText = True
    End With

    ' Columns B-AK descriptions, Accent1 tint 0 (medium blue)
    For i = 0 To UBound(r7Vals)
        col = i + 2   ' B = column 2
        With ws.Cells(7, col)
            .value = r7Vals(i)
            .Font.name = "Avenir Book"
            .Font.Size = 12
            .Font.Bold = False
            .Font.Italic = False
            .Font.ThemeColor = xlThemeColorDark1
            .Font.TintAndShade = 0
            .Interior.ThemeColor = xlThemeColorAccent1
            .Interior.TintAndShade = 0
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlVAlignCenter
            .WrapText = True
        End With
    Next i

    ' -----------------------------------------------------------------------
    ' Row 8 MI field codes
    ' -----------------------------------------------------------------------

    ' Column A empty, Accent1 tint -0.5 (darker blue), thin borders
    With ws.Cells(8, 1)
        .value = ""
        .Interior.ThemeColor = xlThemeColorAccent1
        .Interior.TintAndShade = -0.499984740745262
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlVAlignCenter
        .WrapText = False
        .Borders(xlEdgeTop).LineStyle = xlContinuous
        .Borders(xlEdgeTop).Weight = xlThin
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).Weight = xlThin
        .Borders(xlEdgeLeft).LineStyle = xlContinuous
        .Borders(xlEdgeLeft).Weight = xlThin
    End With

    ' Columns B-AK field codes, Accent1 tint -0.5 (darker blue)
    For i = 0 To UBound(r8Vals)
        col = i + 2   ' B = column 2
        With ws.Cells(8, col)
            .value = r8Vals(i)
            .Font.name = "Avenir Book"
            .Font.Size = 12
            .Font.Bold = False
            .Font.Italic = False
            .Font.ThemeColor = xlThemeColorDark1
            .Font.TintAndShade = 0
            .Interior.ThemeColor = xlThemeColorAccent1
            .Interior.TintAndShade = -0.499984740745262
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlVAlignCenter
            .WrapText = False
        End With
    Next i

    ' -----------------------------------------------------------------------
    ' Extend fill colour from AL (col 38) to end of sheet (both paths)
    ' Replaces default cell colours so the NPI palette covers the full row
    ' -----------------------------------------------------------------------
    Dim npiExtEnd As Long
    npiExtEnd = ws.Columns.count

    ' Row 7 extension: Accent1 tint 0 (same as B-AK)
    With ws.Range(ws.Cells(7, 38), ws.Cells(7, npiExtEnd))
        .Interior.ThemeColor = xlThemeColorAccent1
        .Interior.TintAndShade = 0
    End With

    ' Row 8 extension: Accent1 tint -0.5 (same as B-AK)
    With ws.Range(ws.Cells(8, 38), ws.Cells(8, npiExtEnd))
        .Interior.ThemeColor = xlThemeColorAccent1
        .Interior.TintAndShade = -0.499984740745262
    End With

    ' -----------------------------------------------------------------------
    ' AutoFilter on row 8, columns A through AK
    ' -----------------------------------------------------------------------
    If ws.AutoFilterMode Then ws.AutoFilterMode = False
    If isRefresh Then
        ' Apply autofilter directly - skip FilterRow8BasedOnPopulatedColumns because
        ' that sub clears rows 7-8 when Transaction is blank.
        Dim npiLastCol As Long
        npiLastCol = ws.Cells(8, ws.Columns.count).End(xlToLeft).Column
        If npiLastCol > 1 Then
            ws.Range(ws.Cells(8, 1), ws.Cells(8, npiLastCol)).AutoFilter field:=1
        End If
    Else
        FilterRow8BasedOnPopulatedColumns ws
    End If

    ' -----------------------------------------------------------------------
    ' Sheet rename (skip if already named correctly), button, and cleanup
    ' -----------------------------------------------------------------------
    If Not isRefresh Then UI_RenameSheet "NPI Load Sheet"
    UI_RemoveButtons ws

    ' Add "Load" button (matches Confirm Journal pattern)
    Dim btn As Button
    Set btn = ws.Buttons.Add(10, 85, 131, 45)
    btn.Caption = "Load"
    btn.OnAction = "Process_Click_NPI"

    ws.Range("F4:G5").ClearContents
    ws.Range("A2:B2").ClearContents
    
    ' Remove dropdown validation from G4 and restore normal format
    With ws.Range("G4")
        .Validation.Delete
        .NumberFormat = "General"
        .Interior.ColorIndex = xlNone
        .Font.ColorIndex = xlAutomatic
    End With

    ' Set sheet tab colour to blue
    ws.Tab.Color = RGB(0, 112, 192)

    ws.Rows(7).WrapText = True
    ws.Columns.AutoFit

    ' Remove the "Wait" shape from the original sheet where it was placed
    On Error Resume Next
    wsOriginal.Shapes("Wait").Delete
    On Error GoTo 0
    Application.ScreenUpdating = True
    DoEvents

End Sub

' =============================================================================
' PUBLIC ENTRY POINT
' =============================================================================

''
' Main entry point for the NPI process.
' Loops through a fixed sequence of MI API/transaction pairs and processes
' all data rows on the active sheet for each step.
''
Public Sub Process_Click_NPI()
    Dim startTime As Single
    Dim lastColumn As Long
    Dim ws As Worksheet
    Dim inputColumns() As Long
    Dim inputFields() As String
    Dim inputValues() As String
    Dim iAPI As Long
    Dim i As Long, j As Long

    On Error GoTo ErrorHandler

    Set ws = ActiveSheet
    Set m_Worksheet = ws
    m_CurrentRow = 9

    ' --- UI / version setup ---
    UI_UpdateVersion
    SetFormulasAndFormatting_New ws
    ResetCountFormat ws
    Auth_ValidateSelectedEnvironment

    ws.Calculate
    DoEvents

    If ws.Range("Environment").value = "" Or ws.Range("Environment").value = "Access requested" Then
        MsgBox "Please select a valid environment.", vbExclamation
        Exit Sub
    End If

    startTime = Timer
    UI_ClearStatus

    ' --- Freeze panes / scroll ---
    ws.Cells(1, 1).Select
    ActiveWindow.FreezePanes = False
    ws.Range("C9").Select
    ActiveWindow.FreezePanes = True
    Application.Calculation = xlCalculationManual
    Application.GoTo Reference:="R9C2", Scroll:=True

    ' --- Ensure settings are loaded ---
    If Config_ApiSettings.maxbulk <= 1 Then Config_LoadSettingsFromSheet
    If Config_ApiSettings.refreshSeconds = 0 Then Config_LoadSettingsFromSheet

    ' --- Multi-output: clear rows 10+ if needed ---
    If ws.Range("G5").value = "M" Then
        ws.Rows("10:" & ws.Rows.count).ClearContents
    End If

    ' --- Load environment from cache if available ---
    If manager Is Nothing Then Set manager = New EnvironmentManager
    If manager.LoadEnvironment(ws.Range("I2").value, ws) Then
        ws.Range("J3").value = 2
    End If

    ' --- Authenticate ---
    Config_Company = ws.Range("Company").value
    Config_Division = ws.Range("Division").value
    Tenant_Token
    Doppio_Process.SyncConfigFromDoppioModule

    If ws.Range("I2").value = "Access requested" Or m_s_M3user = "" Then Exit Sub

    ws.Range("User").value = m_s_M3user

    ' --- Ensure B9 has a placeholder value ---
    If IsEmpty(ws.Range("B9").value) Then ws.Range("B9").value = "?"

    ' --- Read input column headers from row 8 ---
    lastColumn = ws.Cells(8, ws.Columns.count).End(xlToLeft).Column

    ReDim inputColumns(1 To 400)
    ReDim inputFields(1 To 400)
    ReDim inputValues(1 To 400)

    j = 0
    For i = 1 To lastColumn
        If ws.Cells(8, i + 1).value <> "" Then
            j = j + 1
            If j > 400 Then Exit For
            inputColumns(j) = i + 1
            inputFields(j) = ws.Cells(8, i + 1).value
        End If
    Next i

    ws.Calculate

    ' --- Execute each NPI MI step ---
    InitNpiSequence

    Dim stopProcess As Boolean
    stopProcess = False

    For iAPI = 0 To UBound(m_ApiList)
        If MsgBox("Step " & (iAPI + 1) & " of " & (UBound(m_ApiList) + 1) & ":" & vbNewLine & _
                  m_ApiList(iAPI) & " / " & m_TrxList(iAPI) & vbNewLine & vbNewLine & _
                  "Run this step?", vbYesNo + vbQuestion, "NPI Process") = vbNo Then
            stopProcess = True
            Exit For
        End If
        UI_ShowPleaseWait "Processing " & m_ApiList(iAPI) & " / " & m_TrxList(iAPI) & "..."
        ProcessNpiStep ws, m_ApiList(iAPI), m_TrxList(iAPI), lastColumn, inputFields, inputColumns, inputValues
        ws.Calculate
        DoEvents
    Next iAPI

    ' --- IPS steps (last) - loops through all configured IPS service/method pairs ---
    Dim iIPS As Long
    If Not stopProcess Then
        For iIPS = 0 To UBound(m_ips_ServiceList)
            If m_ips_ServiceList(iIPS) <> "" And m_ips_MethodList(iIPS) <> "" Then
                If MsgBox("IPS Step " & (iIPS + 1) & ":" & vbNewLine & _
                          m_ips_ServiceList(iIPS) & " / " & m_ips_MethodList(iIPS) & vbNewLine & vbNewLine & _
                          "Run this step?", vbYesNo + vbQuestion, "NPI Process") = vbNo Then
                    Exit For
                End If
                UI_ShowPleaseWait "Processing IPS " & m_ips_ServiceList(iIPS) & " / " & m_ips_MethodList(iIPS) & "..."
                ProcessNpiIPSStep ws, startTime, m_ips_ServiceList(iIPS), m_ips_MethodList(iIPS)
                ws.Calculate
                DoEvents
            End If
        Next iIPS
    End If

    ' --- Cleanup ---
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    UI_DisplayElapsedTime startTime, ws
    UI_KillPleaseWait
    m_CurrentRow = 9
    Exit Sub

ErrorHandler:
    Application.Calculation = xlCalculationManual
    UI_KillPleaseWait
    MsgBox "Error in Process_Click_NPI: " & Err.description, vbCritical
    m_CurrentRow = 9
End Sub

' =============================================================================
' PRIVATE HELPERS
' =============================================================================

''
' Initialise the hardcoded NPI API/Transaction sequence.
' Edit this sub to add, remove, or reorder NPI steps.
''
Private Sub InitNpiSequence()
    m_ApiList(0) = "MMS200MI": m_TrxList(0) = "CpyItmBasic"
    m_ApiList(1) = "MMS200MI": m_TrxList(1) = "CpyItmWhs"
    m_ApiList(2) = "MMS200MI": m_TrxList(2) = "UpdItmPrice"
    m_ApiList(3) = "PPS040MI": m_TrxList(3) = "AddItemSupplier"

    ' IPS steps� add as many as needed, same pattern as the MI steps above.
    ' Leave the array empty (ReDim to -1) to skip all IPS steps.
    ReDim m_ips_ServiceList(0 To 1)
    ReDim m_ips_MethodList(0 To 1)
    m_ips_ServiceList(0) = "MPD_PPS044_Add": m_ips_MethodList(0) = "Add"
    m_ips_ServiceList(1) = "MPD_PPS044_Upd": m_ips_MethodList(1) = "Upd"
End Sub

''
' Process all data rows for a single NPI step (one API/transaction pair).
' Batches rows up to maxbulk, flushes to the API, then writes results back.
''
Private Sub ProcessNpiStep(ws As Worksheet, serviceName As String, method As String, _
                            lastColumn As Long, inputFields() As String, _
                            inputColumns() As Long, inputValues() As String)
    Dim fullBody As String
    Dim body As String
    Dim counter As Long
    Dim settings As ApiSettings
    Dim response As apiResponse
    Dim cellValue As Variant
    Dim i As Long

    settings = Config_ApiSettings
    fullBody = ""
    counter = 0
    m_CurrentRow = 9

    While ws.Cells(m_CurrentRow, 2).value <> ""

        ' Read values from current row
        For i = 1 To UBound(inputColumns)
            If inputColumns(i) <> 0 Then
                cellValue = ws.Cells(m_CurrentRow, inputColumns(i)).value
                If isError(cellValue) Then cellValue = ""
                inputValues(i) = CStr(cellValue)
            End If
        Next i

        ' Accumulate transaction body
        body = BuildNpiTransactionBody(method, inputFields, inputValues)
        If fullBody = "" Then
            fullBody = body
        Else
            fullBody = fullBody & "," & body
        End If

        m_CurrentRow = m_CurrentRow + 1
        counter = counter + 1

        ' Flush when batch is full
        If counter >= settings.maxbulk Then
            body = "{""program"":""" & serviceName & """,""transactions"":[" & fullBody & "]}"
            fullBody = ""
            response = ExecuteNpiApiCall(serviceName, body)
            ProcessNpiResults response, ws, m_CurrentRow - counter
            ws.Calculate
            counter = 0
        End If

    Wend

    ' Flush remaining rows
    If counter > 0 Then
        body = "{""program"":""" & serviceName & """,""transactions"":[" & fullBody & "]}"
        response = ExecuteNpiApiCall(serviceName, body)
        ProcessNpiResults response, ws, m_CurrentRow - counter
    End If
End Sub

''
' Execute a single MI bulk POST for NPI.
' Uses m_s_* global auth state (populated by Tenant_Token).
' Mirrors Doppio_Process.ExecuteMIBulkCall.
''
Private Function ExecuteNpiApiCall(program As String, body As String) As apiResponse
    Dim config As httpConfig
    Dim httpResp As httpResponse
    Dim response As apiResponse
    Dim apiUrl As String
    Dim settings As ApiSettings
    Dim json As Object
    Dim results As Object
    Dim records As Object
    Dim resultItem As Object

    On Error GoTo ErrorHandler

    settings = Config_ApiSettings

    ' Build URL from Doppio global state (same as Doppio_Process.ExecuteMIBulkCall)
    apiUrl = m_s_MainUrl & "/M3/m3api-rest/v2/execute"
    apiUrl = apiUrl & "?maxrecs=" & settings.MaxRecords & "&extendedresult=true"
    If m_s_M3user <> "" Then apiUrl = apiUrl & "&m3user=" & m_s_M3user
    If settings.righttrim Then
        apiUrl = apiUrl & "&righttrim=true"
    Else
        apiUrl = apiUrl & "&righttrim=false"
    End If
    If m_s_Company <> "" Then apiUrl = apiUrl & "&cono=" & m_s_Company
    If m_s_Division <> "" Then apiUrl = apiUrl & "&divi=" & m_s_Division

    With config
        .url = apiUrl
        .method = HttpMethod_POST
        .contentType = "application/json; charset=UTF-8"
        .AcceptType = "application/json; charset=UTF-8"
        .authHeader = m_s_TokenType & " " & m_s_AccessToken
        .timeoutSeconds = settings.MaxTimeout
        .body = body
    End With

    Debug.Print "ExecuteNpiApiCall: " & apiUrl
    httpResp = ExecuteRequest(config)
    Debug.Print "ExecuteNpiApiCall: HTTP " & httpResp.statusCode & " (" & program & ")"

    ' Handle 401 refresh token and retry once
    If httpResp.IsUnauthorized Then
        Debug.Print "ExecuteNpiApiCall: Unauthorized refreshing token..."
        If HandleUnauthorized() Then
            config.authHeader = m_s_TokenType & " " & m_s_AccessToken
            httpResp = ExecuteRequest(config)
        End If
    End If

    response.success = httpResp.success
    response.data = httpResp.body
    response.errorMessage = httpResp.errorMessage

    ' Parse JSON and attach results collection
    If httpResp.success And Len(httpResp.body) > 0 Then
        Set json = ParseJson(httpResp.body)
        If Not json Is Nothing Then
            Set results = json.item("results")
            Set response.results = results
            response.recordCount = 0
            If Not results Is Nothing Then
                For Each resultItem In results
                    On Error Resume Next
                    Set records = resultItem.item("records")
                    If Not records Is Nothing Then
                        response.recordCount = response.recordCount + records.count
                    End If
                    On Error GoTo ErrorHandler
                Next resultItem
            End If
        End If
    End If

    ExecuteNpiApiCall = response
    Exit Function

ErrorHandler:
    Debug.Print "ExecuteNpiApiCall: ERROR - " & Err.description
    response.success = False
    response.errorMessage = Err.description
    ExecuteNpiApiCall = response
End Function

''
' Write NPI API results back to the worksheet.
' Mirrors Doppio_Process.ProcessMIResults.
''
Private Sub ProcessNpiResults(response As apiResponse, ws As Worksheet, startRow As Long)
    Dim results As Object
    Dim resultItem As Object
    Dim records As Object
    Dim record As Object
    Dim rowNum As Long
    Dim colNum As Long
    Dim lastColumn As Long
    Dim fieldName As String
    Dim fieldValue As Variant
    Dim statusCell As Range
    Dim errorMessage As String
    Dim zzusidError As String
    Dim recordCount As Long

    On Error GoTo ErrorHandler

    lastColumn = ws.Cells(8, ws.Columns.count).End(xlToLeft).Column
    UI_ShowPleaseWait "Please Wait... Parsing Results"

    ' HTTP-level failure mark all rows in range as NOK
    If Not response.success Then
        rowNum = startRow
        While ws.Cells(rowNum, 2).value <> "" And rowNum < m_CurrentRow
            Set statusCell = ws.Cells(rowNum, 1)
            statusCell.value = "NOK " & response.errorMessage
            statusCell.Font.Color = COLOR_ERROR
            rowNum = rowNum + 1
        Wend
        Exit Sub
    End If

    If response.results Is Nothing Then Exit Sub

    Set results = response.results
    rowNum = startRow

    For Each resultItem In results
        Set statusCell = ws.Cells(rowNum, 1)

        ' Check for transaction-level error
        errorMessage = ""
        Dim errorField As String
        Dim errorCode As String
        Dim errorCfg As String
        errorField = ""
        errorCode = ""
        errorCfg = ""
        On Error Resume Next
        errorMessage = resultItem.item("errorMessage")
        errorField = resultItem.item("errorField")
        errorCode = resultItem.item("errorCode")
        errorCfg = resultItem.item("errorCfg")
        On Error GoTo ErrorHandler

        If errorMessage <> "" Then
            Dim errDetail As String
            errDetail = "NOK " & errorMessage
            If Trim(errorField) <> "" Then errDetail = errDetail & " [" & Trim(errorField) & "]"
            If Trim(errorCode) <> "" Then errDetail = errDetail & " (" & Trim(errorCode) & ")"
            If Trim(errorCfg) <> "" Then errDetail = errDetail & " {" & Trim(errorCfg) & "}"
            statusCell.value = errDetail
            statusCell.Font.Color = COLOR_ERROR
            rowNum = rowNum + 1
        Else
            On Error Resume Next
            Set records = resultItem.item("records")
            recordCount = 0
            If Not records Is Nothing Then recordCount = records.count
            On Error GoTo ErrorHandler

            If records Is Nothing Or recordCount = 0 Then
                ' No records and no error treat as OK (e.g. update/delete with no output)
                statusCell.value = "OK"
                statusCell.Font.Color = COLOR_SUCCESS
                rowNum = rowNum + 1
            Else
                For Each record In records
                    Set statusCell = ws.Cells(rowNum, 1)

                    ' ZZUSID non-blank (and not whitespace) signals an M3 error
                    zzusidError = ""
                    On Error Resume Next
                    zzusidError = record.item("ZZUSID")
                    On Error GoTo ErrorHandler

                    If zzusidError <> "" And Left(zzusidError, 3) <> "   " Then
                        statusCell.value = "NOK " & Trim(zzusidError)
                        statusCell.Font.Color = COLOR_ERROR
                    Else
                        statusCell.value = "OK"
                        statusCell.Font.Color = COLOR_SUCCESS

                        ' Write output fields back to the row
                        For colNum = 2 To lastColumn
                            fieldName = ws.Cells(8, colNum).value
                            If fieldName <> "" Then
                                On Error Resume Next
                                fieldValue = record.item(fieldName)
                                If Err.Number = 0 Then
                                    If Not IsEmpty(fieldValue) And Not IsNull(fieldValue) Then
                                        ws.Cells(rowNum, colNum).value = fieldValue
                                    End If
                                End If
                                Err.Clear
                                On Error GoTo ErrorHandler
                            End If
                        Next colNum
                    End If

                    rowNum = rowNum + 1
                Next record
            End If
        End If
    Next resultItem

    Exit Sub

ErrorHandler:
    Debug.Print "ProcessNpiResults: ERROR - " & Err.description
    Resume Next
End Sub

''
' Execute the IPS (SOAP/XML) step as the final NPI step.
'
' The NPI sheet stores M3 MI field codes in row 8 (e.g. ITNO, SUNO, LEA1).
' Doppio_IPS.ProcessIPSTransactions expects IPS alias names there instead
' (e.g. ItemNumber, Supplier, PPS044:SupplyLeadTime).
'
' This sub:
'   1. Patches the API / Transaction named ranges so ProcessIPSTransactions
'      knows which service and method to call.
'   2. Loads the IPS field layout via GetLayoutWS to obtain the M3->alias map.
'   3. Temporarily replaces row 8 with IPS aliases derived from that map.
'   4. Calls ProcessIPSTransactions.
'   5. Restores row 8 and the named ranges to their original values.
''
Private Sub ProcessNpiIPSStep(ws As Worksheet, startTime As Single, _
                               serviceName As String, method As String)
    Dim savedApi As String
    Dim savedTrx As String
    Dim lastColumn As Long
    Dim i As Long, k As Long
    Dim savedRow8() As String
    Dim m3Codes() As String
    Dim ipsAliases() As String
    Dim mapCount As Long
    Dim m3Code As String
    Dim ipsAlias As String

    ' Patch named ranges so ProcessIPSTransactions reads the right service/method
    savedApi = ws.Range("API").value
    savedTrx = ws.Range("Transaction").value
    ws.Range("API").value = serviceName
    ws.Range("Transaction").value = method

    ' Reset row counter - each IPS step processes all data rows from the top
    m_CurrentRow = 9

    ' Load the IPS layout for this service/method into the global column collections
    Doppio_IPS.GetLayoutWS serviceName, ws

    ' Build parallel arrays: m3Codes(k) -> ipsAliases(k)
    mapCount = Doppio_IPS.BuildM3ToAliasMap(m3Codes, ipsAliases)

    If mapCount > 0 Then
        ' Save row 8 and replace M3 codes with IPS aliases
        lastColumn = ws.Cells(8, ws.Columns.count).End(xlToLeft).Column
        ReDim savedRow8(1 To lastColumn)

        For i = 1 To lastColumn
            savedRow8(i) = ws.Cells(8, i).value
            m3Code = ws.Cells(8, i).value
            ipsAlias = ""
            For k = 1 To mapCount
                If m3Codes(k) = m3Code Then
                    ipsAlias = ipsAliases(k)
                    Exit For
                End If
            Next k
            ws.Cells(8, i).value = ipsAlias
        Next i

        Doppio_IPS.ProcessIPSTransactions ws, startTime, m_CurrentRow

        ' Restore original row 8 values
        For i = 1 To lastColumn
            ws.Cells(8, i).value = savedRow8(i)
        Next i
    Else
        ' No layout loaded - fall back to calling as-is (M3 codes in row 8)
        Doppio_IPS.ProcessIPSTransactions ws, startTime, m_CurrentRow
    End If

    ' Restore original named range values
    ws.Range("API").value = savedApi
    ws.Range("Transaction").value = savedTrx
End Sub

''
' Build a single MI transaction JSON body.
' Escapes special characters in field values.
''
Private Function BuildNpiTransactionBody(method As String, inputFields() As String, _
                                         inputValues() As String) As String
    Dim body As String
    Dim i As Long
    Dim fieldValue As String
    Dim firstField As Boolean

    body = "{""transaction"":""" & method & """,""record"":{"
    firstField = True

    For i = LBound(inputFields) To UBound(inputFields)
        If inputFields(i) <> "" Then
            fieldValue = inputValues(i)
            fieldValue = Replace(fieldValue, "\", "\\")
            fieldValue = Replace(fieldValue, """", "\""")
            fieldValue = Replace(fieldValue, vbCr, "")
            fieldValue = Replace(fieldValue, vbLf, "")

            If Not firstField Then body = body & ","
            body = body & """" & inputFields(i) & """:""" & fieldValue & """"
            firstField = False
        End If
    Next i

    body = body & "}}"
    BuildNpiTransactionBody = body
End Function


