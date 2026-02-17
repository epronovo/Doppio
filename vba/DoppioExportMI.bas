Attribute VB_Name = "DoppioExportMI"
''
' Doppio Group - EXPORTMI Module (CONSOLIDATED)
' (c) Eric Pronovost - eric@doppiogroup.com
'
' All EXPORTMI SQL query builder and execution routines.
''
Option Explicit

' ============================================================================
' UI SETUP AND BUTTON MANAGEMENT
' ============================================================================

Sub EXPORTMI_AddButtons()
    Dim btn1 As Button, btn2 As Button, btn3 As Button, btn4 As Button, btn5 As Button
    Dim ws As Worksheet
    Set ws = ThisWorkbook.ActiveSheet
    
    ' Platform-specific button positioning
    #If Mac Then
        Set btn1 = ws.Buttons.Add(8, 80, 69, 29)
        Set btn2 = ws.Buttons.Add(8, 113, 69, 29)
        Set btn3 = ws.Buttons.Add(80, 80, 69, 19)
        Set btn4 = ws.Buttons.Add(80, 102, 69, 18)
        Set btn5 = ws.Buttons.Add(80, 123, 69, 19)
    #Else
        Set btn1 = ws.Buttons.Add(8, 78, 69, 27)
        Set btn2 = ws.Buttons.Add(8, 107, 69, 27)
        Set btn3 = ws.Buttons.Add(80, 78, 69, 17)
        Set btn4 = ws.Buttons.Add(80, 98, 69, 17)
        Set btn5 = ws.Buttons.Add(80, 117, 69, 17)
    #End If
    
    ' Configure buttons
    With btn1
        .Caption = "Transactions"
        .OnAction = "GetTransactions_Click"
    End With
    
    With btn2
        .Caption = "Run"
        .OnAction = "EXPORTMI_BuildSQLQuery"
    End With
    
    With btn3
        .Caption = "Layout"
        .OnAction = "GetLayoutAll_Click"
    End With
    
    With btn4
        .Caption = "Parse Qry"
        .OnAction = "EXPORTMI_ParseSQLQuery"
    End With
    
    With btn5
        .Caption = "Autofit"
        .OnAction = "AutoFit_Click"
    End With
    
    ' Initialize sheet layout
    Call InitializeSheetLayout(ws)
End Sub

Sub EXPORTMI_RemoveButtons()
    Dim btn As Button
    Dim ws As Worksheet
    Set ws = ThisWorkbook.ActiveSheet
    
    On Error Resume Next
    For Each btn In ws.Buttons
        btn.Delete
    Next btn
    On Error GoTo 0
End Sub

Private Sub InitializeSheetLayout(ws As Worksheet)
    With ws
        ' Labels
        .Range("A3").value = "table:  "
        .Range("A4").value = "fields:  "
        .Range("A5").value = "where:  "
        .Range("A6").value = "query:  "
        
        ' Label formatting
        With .Range("A3:A6")
            .Font.Bold = True
            .HorizontalAlignment = xlRight
            .VerticalAlignment = xlCenter
        End With
        
        .Range("A3:A5").Font.Color = RGB(47, 117, 181)
        .Range("A6").Font.Color = RGB(84, 130, 53)
        
        ' Input cells
        .Range("B4:B6").NumberFormat = "General"
        .Range("B3").value = ""
        .Range("B4").value = ""
        .Range("B5").value = ""
        
        ' Preserve existing query if valid
        If Left(LCase(.Range("B6").value), 6) <> "select" Then
            .Range("B6").value = ""
        End If
    End With
End Sub

' ============================================================================
' HELPER FUNCTIONS
' ============================================================================

''
' Get separator character from Settings sheet
' Falls back to "^" if not found
''
Private Function GetSeparatorFromSettings() As String
    Dim separator As String
    
    On Error Resume Next
    
    ' Try D12 first
    separator = Sheets("Settings").Range("D12").value
    
    ' Try E12 if D12 is empty
    If separator = "" Then
        separator = Sheets("Settings").Range("E12").value
    End If
    
    On Error GoTo 0
    
    ' Default to ^
    If separator = "" Then
        separator = "^"
    End If
    
    GetSeparatorFromSettings = separator
End Function

' ============================================================================
' SQL QUERY BUILDING AND PARSING
' ============================================================================

Sub EXPORTMI_BuildSQLQuery()
    Dim ws As Worksheet
    Dim tableName As String
    Dim fields As String
    Dim whereclause As String
    Dim sqlQuery As String
    Dim strServiceName As String
    
    Set ws = ThisWorkbook.ActiveSheet
    ws.Rows("7:" & ws.Rows.count).ClearContents
    
    If ws.Range("B4").value = "" Then
        ws.Range("B4").value = "*"
    End If
    
    ' Get input values
    tableName = UCase(Trim(ws.Range("B3").value))
    ws.Range("B3").value = tableName
    
    fields = UCase(ws.Range("B4").value)
    fields = Replace(fields, vbTab, ",")
    ws.Range("B4").value = fields
    fields = Trim(ws.Range("B4").value)
    whereclause = Trim(ws.Range("B5").value)
    
    ' Validate table name
    If tableName = "" Then
        MsgBox "Table must be entered!", vbExclamation, "Missing Table"
        Exit Sub
    End If
    
    ' Get transaction name
    strServiceName = ws.Range("G4").value
    If strServiceName = "" Then strServiceName = "EXPORTMI"
    
    ' Set default fields if empty
    If fields = "" Then fields = "*"
    
    ' Build the SQL query
    If whereclause <> "" Then
        whereclause = EXPORTMI_ParseCondition(whereclause)
        sqlQuery = "select " & fields & " from " & tableName & " where " & whereclause
    Else
        sqlQuery = "select " & fields & " from " & tableName
    End If
    
    ' Write query to cell
    ws.Range("B6").value = sqlQuery
    
    ' Validate and process
    ValidateTransaction strServiceName
    
    ' Parse fields to headers (IMPORTANT: This now puts headers in Col 2+)
    EXPORTMI_ParseFieldsIntoColumns
    
    ' Run process
    EXPORTMI_Process
End Sub

Private Sub ValidateTransaction(strServiceName As String)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.ActiveSheet
    
    If strServiceName = "" Then
        ws.Range("G4").value = "EXPORTMI"
    End If
End Sub

Function EXPORTMI_ParseCondition(ByVal inputStr As String) As String
    Dim parsedStr As String
    Dim i As Long
    Dim currentChar As String
    Dim nextWord As String
    Dim specialChars As String
    Dim keywords As String
    Dim words As Variant
    
    parsedStr = inputStr
    
    ' Clean up WHERE clause
    parsedStr = Replace(parsedStr, "WHERE ", "")
    parsedStr = Replace(parsedStr, "where ", "")
    parsedStr = Replace(parsedStr, " AND ", " and ")
    parsedStr = Replace(parsedStr, " OR ", " or ")
    parsedStr = Replace(parsedStr, " !=", "<>")
    
    ' Normalize operators
    parsedStr = Replace(parsedStr, " = ", "=")
    parsedStr = Replace(parsedStr, " < ", "<")
    parsedStr = Replace(parsedStr, " > ", ">")
    
    ' Protect >= and <=
    parsedStr = Replace(parsedStr, ">=", "[GTE]")
    parsedStr = Replace(parsedStr, "<=", "[LTE]")
    
    ' Add spaces around operators
    specialChars = "=><"
    For i = 1 To Len(specialChars)
        currentChar = Mid(specialChars, i, 1)
        parsedStr = Replace(parsedStr, currentChar, " " & currentChar & " ")
    Next i
    
    ' Restore >= and <=
    parsedStr = Replace(parsedStr, "[GTE]", ">=")
    parsedStr = Replace(parsedStr, "[LTE]", "<=")
    parsedStr = Replace(parsedStr, "<  >", "<>")
    
    ' Uppercase 6-character field names
    keywords = "= < > and or"
    words = Split(parsedStr, " ")
    For i = LBound(words) To UBound(words)
        If i + 1 <= UBound(words) Then
            nextWord = Trim(words(i + 1))
            If Len(words(i)) = 6 And InStr(keywords, nextWord) > 0 Then
                words(i) = UCase(words(i))
            End If
        End If
    Next i
    parsedStr = Join(words, " ")
    
    ' Remove extra spaces
    Do While InStr(parsedStr, "  ") > 0
        parsedStr = Replace(parsedStr, "  ", " ")
    Loop
    
    EXPORTMI_ParseCondition = parsedStr
End Function

Sub EXPORTMI_ParseFieldsIntoColumns()
    Dim ws As Worksheet
    Dim fieldsCell As Range
    Dim fieldsValue As String
    Dim fieldArray() As String
    Dim i As Long
    Dim startCol As Long
    
    Set ws = ThisWorkbook.ActiveSheet
    Set fieldsCell = ws.Range("B4")
    fieldsValue = Trim(fieldsCell.value)
    
    ' Clear existing headers in Row 7 and 8
    ws.Rows("7:8").ClearContents
    
    If fieldsValue = "" Or fieldsValue = "*" Then Exit Sub
    
    ' Split by comma
    fieldArray = Split(fieldsValue, ",")
    
    ' FIXED: Start writing at column 2 (B) for Row 8
    ' Column 1 (A) is reserved for Status
    startCol = 2
    
    For i = LBound(fieldArray) To UBound(fieldArray)
        ws.Cells(8, startCol + i).value = Trim(fieldArray(i))
    Next i
End Sub

Sub EXPORTMI_ParseSQLQuery()
    Dim ws As Worksheet
    Dim query As String
    Dim selectField As String
    Dim fromTable As String
    Dim whereCondition As String
    Dim selectPos As Long
    Dim fromPos As Long
    Dim wherePos As Long
    
    DoppioUI.UI_UpdateVersion
    
    If Range("I2").value = "" Or Range("I2").value = "Access requested" Then
        MsgBox "Please select an environment from the dropdown in I2.", vbExclamation
        Exit Sub
    End If
    
    Set ws = ThisWorkbook.ActiveSheet
    
    If ws.Range("B6").value = "" Then
        MsgBox "Query must be provided!", vbExclamation, "Missing Query"
        Exit Sub
    End If
    
    query = ws.Range("B6").value
    
    ' Normalize query keywords
    query = Replace(query, "SELECT ", "select ")
    query = Replace(query, " FROM ", " from ")
    query = Replace(query, " WHERE ", " where ")
    ws.Range("B6").value = query
    
    ' Parse query components
    selectPos = InStr(1, LCase(query), "select") + 6
    fromPos = InStr(1, LCase(query), "from")
    wherePos = InStr(1, LCase(query), "where")
    
    selectField = Trim(Mid(query, selectPos, fromPos - selectPos))
    
    If wherePos > 0 Then
        fromTable = UCase(Trim(Mid(query, fromPos + 4, wherePos - fromPos - 4)))
        whereCondition = Trim(Mid(query, wherePos + 5))
    Else
        fromTable = UCase(Trim(Mid(query, fromPos + 4)))
        whereCondition = ""
    End If
    
    whereCondition = EXPORTMI_ParseCondition(whereCondition)
    
    ' Update cells
    ws.Range("B4").value = selectField
    ws.Range("B3").value = fromTable
    ws.Range("B5").value = whereCondition
    
    ' Rebuild query
    If whereCondition <> "" Then
        whereCondition = " where " & whereCondition
    End If
    
    query = "select " & selectField & " from " & fromTable & whereCondition
    ws.Range("B6").value = query
    
    Call RenameSheet("EXPORTMI for " & Left(ws.Range("B3").value, 6))
    Call EXPORTMI_ParseFieldsIntoColumns
    
    ' FIXED: Ensure headers are fetched if * is used or needed
    Call EXPORTMI_GetTableColumns(ws)
End Sub


Function EXPORTMI_CreateJsonBody() As String
    Dim ws As Worksheet
    Dim separator As String
    Dim query As String
    
    Set ws = ActiveSheet
    
    ' Get separator from Settings sheet
    separator = GetSeparatorFromSettings()
    If separator = "" Then separator = "^"
    
    ' Get query
    query = ws.Range("B6").value
    
    EXPORTMI_CreateJsonBody = "{""program"":""EXPORTMI"",""transactions"":[{""transaction"":""Select"",""record"":{""SEPC"":""" & separator & """,""HDRS"":""0"",""QERY"":""" & query & """},""selectedColumns"":[""REPL""]}]}"
End Function

Private Function EXPORTMI_CreateJsonBodyWithQuery(ws As Worksheet) As String
    Dim program As String
    Dim transaction As String
    Dim record As String
    Dim selectedColumns As String
    Dim query As String
    Dim body As String
    Dim splitChar As String
    
    program = "EXPORTMI"
    transaction = "Select"
    
    ' Get split character from DoppioConfig or Settings
    splitChar = GetSeparatorFromSettings()
    If splitChar = "" Then splitChar = "^"
    
    ' Get query from sheet
    query = ws.Range("B6").value
    query = Replace(query, """", "\""")
    query = Replace(query, "select ", "")
    
    record = """SEPC"":""" & splitChar & """,""HDRS"":""0"",""QERY"":""" & query & """"
    selectedColumns = """REPL"""
    
    ' Build JSON
    body = "{""program"":""" & program & """," _
         & """transactions"":[{" _
         & """transaction"":""" & transaction & """," _
         & """record"":{" & record & "}," _
         & """selectedColumns"":[" & selectedColumns & "]" _
         & "}]}"
    
    EXPORTMI_CreateJsonBodyWithQuery = body
End Function

' ============================================================================
' MAIN PROCESS FUNCTION (FIXED AUTHENTICATION)
' Matches Process_Click pattern exactly
' ============================================================================

Public Sub EXPORTMI_Process()
    Dim ws As Worksheet
    Dim startTime As Single
    Dim body As String
    Dim response As apiResponse
    Dim currentEnv As String
    
    On Error GoTo ErrorHandler
    
    Set ws = ActiveSheet
        
    ' Force calculation
    ws.Calculate
    DoEvents
    DoppioUI.UI_ShowPleaseWait "Please Wait... Calling EXPORTMI API"

    ' Get the current environment from worksheet
    currentEnv = ws.Range("I2").value
    
    ' Check if environment is selected
    If currentEnv = "" Or currentEnv = "Access requested" Then
        MsgBox "Please select a valid environment.", vbExclamation
        Exit Sub
    End If
    
    ' =========================================================================
    ' AUTHENTICATION
    ' =========================================================================
    If Not EXPORTMI_EnsureAuthenticated(ws) Then
        MsgBox "Authentication failed. Please check your environment settings.", vbCritical
        Exit Sub
    End If
    
    #If DEBUG_MODE Then
        Debug.Print "EXPORTMI_Process: Using URL = " & Doppio.m_s_MainUrl
    #End If
    
    ' =========================================================================
    ' END AUTHENTICATION
    ' =========================================================================
    
    ' Parse headers (ensures Row 8 is correct based on B4)
    EXPORTMI_ParseFieldsIntoColumns
    
    ' If fields are wildcard, fetch columns from API to fill headers
    'If Doppio.m_obj_ColumnNames.count <= 0 Or ws.Range("B4").value = "*" Then
        EXPORTMI_GetTableColumns ws
    'End If
    
    startTime = Timer
    
    ' Build the JSON body
    body = EXPORTMI_CreateJsonBodyWithQuery(ws)
    
    #If DEBUG_MODE Then
        Debug.Print "EXPORTMI_Process: Body = " & Left(body, 200)
    #End If
        
    ' Execute the API call (uses Doppio.m_s_* values)
    response = EXPORTMI_ExecuteCall(body, ws)
    
    #If DEBUG_MODE Then
        Debug.Print "EXPORTMI_Process: Success = " & response.success
        Debug.Print "EXPORTMI_Process: RecordCount = " & response.recordCount
    #End If
    
    ' Process results
    If response.success Then
        EXPORTMI_ProcessResults response, ws
    Else
        MsgBox "EXPORTMI failed: " & response.errorMessage, vbCritical
    End If
    
    ' Update timer
    DoppioUI.UI_DisplayElapsedTime startTime, ws
    Doppio.AutoFit_Click
    
    Exit Sub
    
ErrorHandler:
    DoppioUI.UI_KillPleaseWait
    MsgBox "Error in EXPORTMI_Process: " & Err.description, vbCritical
End Sub

' ============================================================================
' AUTHENTICATION HELPER
' ============================================================================

Private Function EXPORTMI_EnsureAuthenticated(ws As Worksheet) As Boolean
    Dim currentEnv As String
    currentEnv = ws.Range("I2").value
    
    ' Check if we need a new token:
    ' 1. No token exists
    ' 2. User info not loaded
    ' 3. Environment has CHANGED
    If Doppio.m_s_AccessToken = "" Or _
       ws.Range("I3").value = "" Or _
       ws.Range("I4").value = "" Or _
       Doppio.activeEnvironment <> currentEnv Then
        
        #If DEBUG_MODE Then
            Debug.Print "EXPORTMI_EnsureAuthenticated: Getting new token for " & currentEnv
        #End If
        
        ' Get new token using Doppio.Tenant_Token
        Doppio.Tenant_Token
        
        ' Verify we got a token
        If Doppio.m_s_AccessToken = "" Then
            EXPORTMI_EnsureAuthenticated = False
            Exit Function
        End If
    End If
    
    EXPORTMI_EnsureAuthenticated = True
End Function

' ============================================================================
' TABLE COLUMNS (Uses Doppio.m_s_* values directly)
' ============================================================================

Private Sub EXPORTMI_GetTableColumns(ws As Worksheet)
    Dim response As apiResponse
    Dim tableName As String
    Dim record As Object
    Dim flds As String, typ As String, leng As String, description As String
    
    On Error GoTo ErrorHandler
    
    tableName = Left(ws.Range("B3").value, 6)
    
    If tableName = "" Then Exit Sub
    
    #If DEBUG_MODE Then
        Debug.Print "EXPORTMI_GetTableColumns: Table = " & tableName
    #End If
    
    ' FIXED: Ensure we are authenticated before fetching columns
    If Not EXPORTMI_EnsureAuthenticated(ws) Then
        MsgBox "Authentication failed. Cannot fetch table columns.", vbCritical
        Exit Sub
    End If
    
    ' Call MRS001MI/LstFieldInfo using Doppio values
    response = ExecuteMIGetCall("MRS001MI", "LstFieldInfo", "FILE=" & tableName, ws)
    
    ' Initialize column collections
    Doppio.m_obj_ColumnNames.Initialize
    Doppio.m_obj_ColumnDescriptions.Initialize
    Doppio.m_obj_ColumnTypes.Initialize
    Doppio.m_obj_ColumnConditions.Initialize
    Doppio.m_obj_ColumnDirections.Initialize
    
    ' Process records
    If response.success And Not response.records Is Nothing Then
        For Each record In response.records
            Doppio.m_obj_ColumnNames.Add record.item("FLNA")
            
            flds = record.item("FRL1") & vbCrLf
            typ = record.item("FLTY")
            leng = record.item("FLLE")
            description = flds & typ & leng
            
            Doppio.m_obj_ColumnDescriptions.Add description
            Doppio.m_obj_ColumnTypes.Add record.item("FLTY")
            Doppio.m_obj_ColumnConditions.Add False
            Doppio.m_obj_ColumnDirections.Add "O"
        Next record
    End If
    
    ' Autofit columns
    If ws.Range("B4").value = "*" Then
        Doppio.AutoFit_ColumnsAndRows True, False
        Doppio.AutoFit_ColumnsAndRows False, False
    Else
        Doppio.AutoFit_ColumnsAndRows False, False
    End If
    
    Exit Sub
    
ErrorHandler:
    #If DEBUG_MODE Then
        Debug.Print "EXPORTMI_GetTableColumns: ERROR - " & Err.description
    #End If
End Sub

' ============================================================================
' API CALL HELPERS (Use Doppio.m_s_* values directly)
' ============================================================================

Private Function ExecuteMIGetCall(program As String, transaction As String, parameters As String, ws As Worksheet) As apiResponse
    Dim config As httpConfig
    Dim httpResponse As httpResponse
    Dim response As apiResponse
    Dim apiUrl As String
    Dim settings As ApiSettings
    Dim json As Object
    Dim wsCompany As String
    
    On Error GoTo ErrorHandler
    
    settings = DoppioConfig.Config_ApiSettings
    wsCompany = ws.Range("Company").value
    
    ' Build URL using Doppio.m_s_MainUrl
    apiUrl = Doppio.m_s_MainUrl & "/M3/m3api-rest/v2/execute/" & program & "/" & transaction & "?"
    apiUrl = apiUrl & parameters
    apiUrl = apiUrl & "&maxrecs=" & settings.MaxRecords
    
    If wsCompany <> "" Then
        apiUrl = apiUrl & "&cono=" & wsCompany
    End If
    
    ' Configure HTTP request
    config.url = apiUrl
    config.method = HttpMethod_GET
    config.contentType = "application/json; charset=UTF-8"
    config.AcceptType = "application/json; charset=UTF-8"
    config.authHeader = Doppio.m_s_TokenType & " " & Doppio.m_s_AccessToken
    config.TimeoutSeconds = settings.MaxTimeout
    config.body = ""
    
    #If DEBUG_MODE Then
        Debug.Print "ExecuteMIGetCall: URL = " & apiUrl
    #End If
    
    ' Execute request
    httpResponse = DoppioHttp.ExecuteRequest(config)
    
    #If DEBUG_MODE Then
        Debug.Print "ExecuteMIGetCall: HTTP Status = " & httpResponse.statusCode
    #End If
    
    ' Build response
    response.success = httpResponse.success
    response.data = httpResponse.body
    response.errorMessage = httpResponse.errorMessage
    
    ' Parse JSON response
    If httpResponse.success And Len(httpResponse.body) > 0 Then
        Set json = JsonConverter.ParseJson(httpResponse.body)
        
        If Not json Is Nothing Then
            Set response.results = json.item("results")
            
            ' Get records from first result
            If Not response.results Is Nothing Then
                If response.results.count > 0 Then
                    Set response.records = response.results(1).item("records")
                    If Not response.records Is Nothing Then
                        response.recordCount = response.records.count
                    End If
                End If
            End If
        End If
    End If
    
    ExecuteMIGetCall = response
    Exit Function
    
ErrorHandler:
    #If DEBUG_MODE Then
        Debug.Print "ExecuteMIGetCall: ERROR - " & Err.description
    #End If
    response.success = False
    response.errorMessage = Err.description
    ExecuteMIGetCall = response
End Function

Private Function EXPORTMI_ExecuteCall(body As String, ws As Worksheet) As apiResponse
    Dim config As httpConfig
    Dim httpResponse As httpResponse
    Dim response As apiResponse
    Dim apiUrl As String
    Dim settings As ApiSettings
    Dim json As Object
    Dim results As Object
    Dim resultItem As Object
    Dim records As Object
    Dim wsCompany As String
    Dim wsDivision As String
    
    On Error GoTo ErrorHandler
    
    settings = DoppioConfig.Config_ApiSettings
    
    ' Get company/division from worksheet
    wsCompany = ws.Range("Company").value
    wsDivision = ws.Range("Division").value
    
    ' Build URL using Doppio.m_s_MainUrl
    apiUrl = Doppio.m_s_MainUrl & "/M3/m3api-rest/v2/execute"
    apiUrl = apiUrl & "?maxrecs=" & settings.MaxRecords & "&extendedresult=true"
    
    If Doppio.m_s_M3user <> "" Then
        apiUrl = apiUrl & "&m3user=" & Doppio.m_s_M3user
    End If
    
    If settings.righttrim Then
        apiUrl = apiUrl & "&righttrim=true"
    Else
        apiUrl = apiUrl & "&righttrim=false"
    End If
    
    ' Use WORKSHEET values for company/division
    If wsCompany <> "" Then
        apiUrl = apiUrl & "&cono=" & wsCompany
    End If
    
    If wsDivision <> "" Then
        apiUrl = apiUrl & "&divi=" & wsDivision
    End If
    
    #If DEBUG_MODE Then
        Debug.Print "EXPORTMI_ExecuteCall: URL = " & apiUrl
    #End If
    
    ' Configure HTTP request
    config.url = apiUrl
    config.method = HttpMethod_POST
    config.contentType = "application/json; charset=UTF-8"
    config.AcceptType = "application/json; charset=UTF-8"
    config.authHeader = Doppio.m_s_TokenType & " " & Doppio.m_s_AccessToken
    config.TimeoutSeconds = settings.MaxTimeout
    config.body = body
    
    ' Execute request
    httpResponse = DoppioHttp.ExecuteRequest(config)
    
    ' Handle unauthorized - try to get fresh token
    If httpResponse.statusCode = 401 Then
        #If DEBUG_MODE Then
            Debug.Print "EXPORTMI_ExecuteCall: Unauthorized, refreshing token..."
        #End If
        Doppio.m_s_AccessToken = ""
        Doppio.Tenant_Token
        
        If Doppio.m_s_AccessToken <> "" Then
            config.authHeader = Doppio.m_s_TokenType & " " & Doppio.m_s_AccessToken
            httpResponse = DoppioHttp.ExecuteRequest(config)
        End If
    End If
    
    ' Build response
    response.success = httpResponse.success
    response.data = httpResponse.body
    response.errorMessage = httpResponse.errorMessage
    
    ' Parse JSON
    If httpResponse.success And Len(httpResponse.body) > 0 Then
        Set json = JsonConverter.ParseJson(httpResponse.body)
        
        If Not json Is Nothing Then
            Set results = json.item("results")
            Set response.results = results
            
            ' Count total records
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
    
    EXPORTMI_ExecuteCall = response
    Exit Function
    
ErrorHandler:
    #If DEBUG_MODE Then
        Debug.Print "EXPORTMI_ExecuteCall: ERROR - " & Err.description
    #End If
    response.success = False
    response.errorMessage = Err.description
    EXPORTMI_ExecuteCall = response
End Function

' ============================================================================
' RESULTS PROCESSING
' ============================================================================

Private Sub EXPORTMI_ProcessResults(response As apiResponse, ws As Worksheet)
    Dim results As Object
    Dim resultItem As Object
    Dim records As Object
    Dim record As Object
    Dim replValue As String
    Dim replParts() As String
    Dim separator As String
    Dim row As Long
    Dim i As Long
    Dim dataArray() As Variant
    Dim recordCount As Long
    Dim fieldCount As Long
    
    On Error GoTo ErrorHandler
    
    ' --- Get Separator Logic (Embedded) ---
    separator = ""
    On Error Resume Next
    separator = ThisWorkbook.Sheets("Settings").Range("D12").value
    If separator = "" Then separator = ThisWorkbook.Sheets("Settings").Range("E12").value
    On Error GoTo ErrorHandler
    If separator = "" Then separator = "^"
    ' --------------------------------------
    
    Set results = response.results
    If results Is Nothing Then
        #If DEBUG_MODE Then
            Debug.Print "EXPORTMI_ProcessResults_New: No results"
        #End If
        Exit Sub
    End If
    
    ' 1. Calculate total record count first to dimension array efficiently
    recordCount = 0
    For Each resultItem In results
        On Error Resume Next
        Set records = resultItem.item("records")
        If Not records Is Nothing Then
            recordCount = recordCount + records.count
        End If
        On Error GoTo ErrorHandler
    Next resultItem
    
'    If recordCount = 0 Then Exit Sub
    If recordCount = 0 Then
        ' Check for error message in results
        On Error Resume Next
        Dim errMsg As String
        For Each resultItem In results
            errMsg = resultItem.item("errorMessage")
            If errMsg <> "" Then
                ws.Range("A9").value = "NOK " & errMsg
                ws.Range("A9").Font.Color = COLOR_ERROR
                Exit For
            End If
        Next resultItem
        On Error GoTo ErrorHandler
        Exit Sub
    End If
    ' 2. Determine field count from the first available record
    fieldCount = 0
    For Each resultItem In results
        If Not resultItem.item("records") Is Nothing Then
            If resultItem.item("records").count > 0 Then
                replValue = resultItem.item("records")(1).item("REPL")
                replParts = Split(replValue, separator)
                fieldCount = UBound(replParts) + 1
                Exit For
            End If
        End If
    Next resultItem
    
    If fieldCount = 0 Then Exit Sub
    
    ' 3. Create Array (1-based for easy Excel write)
    ReDim dataArray(1 To recordCount, 1 To fieldCount)
    
    ' 4. Populate Array
    row = 1
    For Each resultItem In results
        On Error Resume Next
        Set records = resultItem.item("records")
        On Error GoTo ErrorHandler
        
        If Not records Is Nothing Then
            For Each record In records
                replValue = record.item("REPL")
                
                ' Clean specific M3 characters if needed
                replValue = Replace(replValue, "_Ü_", "-")
                replValue = Replace(replValue, Chr(194) & Chr(160), "")
                
                replParts = Split(replValue, separator)
                
                ' Fill row
                For i = 0 To UBound(replParts)
                    If i < fieldCount Then
                        dataArray(row, i + 1) = replParts(i)
                    End If
                Next i
                
                row = row + 1
            Next record
        End If
    Next resultItem
    
    ' 5. Write to Sheet (FIXED RANGE CALCULATION)
    ' Start at Row 9, Column 2 (B)
    ' End Column = 1 + fieldCount (e.g., if 1 field, end col is 2. If 5 fields, end col is 6)
    ws.Range(ws.Cells(9, 2), ws.Cells(9 + recordCount - 1, 1 + fieldCount)).value = dataArray
    
    #If DEBUG_MODE Then
        Debug.Print "EXPORTMI_ProcessResults_New: Wrote " & recordCount & " rows, " & fieldCount & " columns"
    #End If
    
    ' Autofit columns
    Doppio.AutoFit_Click
    
    Exit Sub
    
ErrorHandler:
    #If DEBUG_MODE Then
        Debug.Print "EXPORTMI_ProcessResults_New: ERROR - " & Err.description
    #End If
End Sub


