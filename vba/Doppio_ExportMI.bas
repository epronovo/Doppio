Attribute VB_Name = "Doppio_ExportMI"
''
' Doppio Group - EXPORTMI Module (CONSOLIDATED)
' (c) Eric Pronovost - eric@doppiogroup.com
'
' All EXPORTMI SQL query builder and execution routines.
'
'
Option Explicit

' ============================================================================
' UI SETUP AND BUTTON MANAGEMENT
' ============================================================================

Sub EXPORTMI_AddButtons()
    Dim btn1 As Button, btn2 As Button, btn3 As Button, btn4 As Button, btn5 As Button
    Dim ws As Worksheet
    Set ws = ThisWorkbook.ActiveSheet
    
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
    
    InitializeSheetLayout ws
End Sub



Private Sub InitializeSheetLayout(ws As Worksheet)
    With ws
        .Range("A3").value = "table:  "
        .Range("A4").value = "fields:  "
        .Range("A5").value = "where:  "
        .Range("A6").value = "query:  "
        
        With .Range("A3:A6")
            .Font.Bold = True
            .HorizontalAlignment = xlRight
            .VerticalAlignment = xlCenter
        End With
        
        .Range("A3:A5").Font.Color = RGB(47, 117, 181)
        .Range("A6").Font.Color = RGB(84, 130, 53)
        
        .Range("B4:B6").NumberFormat = "General"
        .Range("B3").value = ""
        .Range("B3").Font.Color = RGB(0, 0, 0)
        .Range("B4").value = ""
        .Range("B5").value = ""
        
        If Left(LCase(.Range("B6").value), 6) <> "select" Then
            .Range("B6").value = ""
        End If
    End With
End Sub


' ============================================================================
' HELPER FUNCTIONS
' ============================================================================

Private Function GetSeparatorFromSettings() As String
    Dim separator As String
    
    On Error Resume Next
    separator = Sheets("Settings").Range("D12").value
    If separator = "" Then
        separator = Sheets("Settings").Range("E12").value
    End If
    On Error GoTo 0
    
    If separator = "" Then separator = "^"
    
    GetSeparatorFromSettings = separator
End Function


' ============================================================================
' AUTHENTICATION & CONFIG SYNC
' (Follows Doppio_Process pattern: Tenant_Token ? Sync ? Override)
' ============================================================================

Private Function EnsureAuthenticated(ws As Worksheet) As Boolean
    Dim currentEnv As String
    
    On Error GoTo ErrorHandler
    
    currentEnv = ws.Range("I2").value
    
    If currentEnv = "" Or currentEnv = "Access requested" Then
        EnsureAuthenticated = False
        Exit Function
    End If
    
    ' Authenticate if needed
    If m_s_AccessToken = "" Or _
       ws.Range("I3").value = "" Or _
       ws.Range("I4").value = "" Or _
       activeEnvironment <> currentEnv Then
        
        Tenant_Token
        
        If m_s_AccessToken = "" Then
            EnsureAuthenticated = False
            Exit Function
        End If
    End If
    
    ' Sync Doppio globals ? Doppio_Config
    SyncConfigFromDoppio
    
    ' Override company/division from worksheet
    Config_Company = ws.Range("Company").value
    Config_Division = ws.Range("Division").value
    
    ' Ensure settings are loaded
    Dim settings As ApiSettings
    settings = Config_ApiSettings
    If settings.maxbulk <= 1 Or settings.refreshSeconds = 0 Then
        Config_LoadSettingsFromSheet
    End If
    
    EnsureAuthenticated = True
    Exit Function
    
ErrorHandler:
    Debug.Print "EXPORTMI EnsureAuthenticated: ERROR - " & Err.description
    EnsureAuthenticated = False
End Function


Private Sub SyncConfigFromDoppio()
    On Error Resume Next
    Config_AccessToken = m_s_AccessToken
    Config_RefreshToken = m_s_RefreshToken
    Config_TokenType = m_s_TokenType
    Config_SelectedEnvironment = m_s_SelectedEnvironment
    Config_M3User = m_s_M3user
    Config_Company = m_s_Company
    Config_Division = m_s_Division
    
    ' Load tenant config to set mainUrl in Doppio_Config
    Config_LoadTenantConfig m_s_SelectedEnvironment
    On Error GoTo 0
End Sub


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
    
    tableName = UCase(Trim(ws.Range("B3").value))
    ws.Range("B3").value = tableName
    
    If InStr(1, LCase(ws.Range("B4").value), "count") = 0 Then
        fields = UCase(ws.Range("B4").value)
    Else
        fields = ws.Range("B4").value
    End If
    fields = Replace(fields, vbTab, ",")
    ws.Range("B4").value = fields
    fields = Trim(ws.Range("B4").value)
    whereclause = Trim(ws.Range("B5").value)
    
    If tableName = "" Then
        MsgBox "Table must be entered!", vbExclamation, "Missing Table"
        Exit Sub
    End If
    
    strServiceName = ws.Range("G4").value
    If strServiceName = "" Then strServiceName = "EXPORTMI"
    
    If fields = "" Then fields = "*"
    
    If whereclause <> "" Then
        whereclause = EXPORTMI_ParseCondition(whereclause)
        sqlQuery = "select " & fields & " from " & tableName & " where " & whereclause
    Else
        sqlQuery = "select " & fields & " from " & tableName
    End If
    
    ws.Range("B6").value = sqlQuery
    
    Dim tableName1 As String
    tableName1 = Trim(ws.Range("B3").value)
    If tableName1 Like "*[A-Za-z0-9]*" Then
        Call UI_RenameSheet("EXPORTMI for " & tableName1)
    Else
        Call UI_RenameSheet("")
    End If
    ValidateTransaction strServiceName
    EXPORTMI_ParseFieldsIntoColumns
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
    Dim Keywords As String
    Dim words As Variant
    
    parsedStr = inputStr
    
    parsedStr = Replace(parsedStr, "WHERE ", "")
    parsedStr = Replace(parsedStr, "where ", "")
    parsedStr = Replace(parsedStr, " AND ", " and ")
    parsedStr = Replace(parsedStr, " OR ", " or ")
    parsedStr = Replace(parsedStr, " !=", "<>")
    
    parsedStr = Replace(parsedStr, " = ", "=")
    parsedStr = Replace(parsedStr, " < ", "<")
    parsedStr = Replace(parsedStr, " > ", ">")
    
    parsedStr = Replace(parsedStr, ">=", "[GTE]")
    parsedStr = Replace(parsedStr, "<=", "[LTE]")
    
    specialChars = "=><"
    For i = 1 To Len(specialChars)
        currentChar = Mid(specialChars, i, 1)
        parsedStr = Replace(parsedStr, currentChar, " " & currentChar & " ")
    Next i
    
    parsedStr = Replace(parsedStr, "[GTE]", ">=")
    parsedStr = Replace(parsedStr, "[LTE]", "<=")
    parsedStr = Replace(parsedStr, "<  >", "<>")
    
    Keywords = "= < > and or"
    words = Split(parsedStr, " ")
    For i = LBound(words) To UBound(words)
        If i + 1 <= UBound(words) Then
            nextWord = Trim(words(i + 1))
            If Len(words(i)) = 6 And InStr(Keywords, nextWord) > 0 Then
                words(i) = UCase(words(i))
            End If
        End If
    Next i
    parsedStr = Join(words, " ")
    
    Do While InStr(parsedStr, "  ") > 0
        parsedStr = Replace(parsedStr, "  ", " ")
    Loop
    
    EXPORTMI_ParseCondition = parsedStr
End Function


Sub EXPORTMI_ParseFieldsIntoColumns()
    Dim ws As Worksheet
    Dim fieldsValue As String
    Dim fieldArray() As String
    Dim i As Long
    Dim startCol As Long
    
    Set ws = ThisWorkbook.ActiveSheet
    fieldsValue = Trim(ws.Range("B4").value)
    
    ws.Rows("7:8").ClearContents
    
    If fieldsValue = "" Or fieldsValue = "*" Then Exit Sub
    
    fieldArray = Split(fieldsValue, ",")
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
    
    UI_UpdateVersion
    
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
    
    query = Replace(query, "SELECT ", "select ")
    query = Replace(query, " FROM ", " from ")
    query = Replace(query, " WHERE ", " where ")
    ws.Range("B6").value = query
    
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
    
    ws.Range("B4").value = selectField
    ws.Range("B3").value = fromTable
    ws.Range("B5").value = whereCondition
    
    If whereCondition <> "" Then
        whereCondition = " where " & whereCondition
    End If
    
    query = "select " & selectField & " from " & fromTable & whereCondition
    ws.Range("B6").value = query
    
    Dim tableName2 As String
    tableName2 = Trim(ws.Range("B3").value)
    If tableName2 Like "*[A-Za-z0-9]*" Then
        Call UI_RenameSheet("EXPORTMI for " & tableName2)
    Else
        Call UI_RenameSheet("")
    End If
    Call EXPORTMI_ParseFieldsIntoColumns
    Call EXPORTMI_GetTableColumns(ws)
End Sub


' ============================================================================
' JSON BODY BUILDER
' ============================================================================

Private Function EXPORTMI_BuildTransactions(ws As Worksheet) As String
    Dim query As String
    Dim separator As String
    
    separator = GetSeparatorFromSettings()
    
    query = ws.Range("B6").value
    query = Replace(query, """", "\""")
    query = Replace(query, "select ", "")
    
    EXPORTMI_BuildTransactions = "[{" & _
        """transaction"":""Select""," & _
        """record"":{""SEPC"":""" & separator & """,""HDRS"":""0"",""QERY"":""" & query & """}," & _
        """selectedColumns"":[""REPL""]" & _
        "}]"
End Function


' ============================================================================
' TABLE COLUMNS (Uses ExecuteMiGet)
' ============================================================================

Private Sub EXPORTMI_GetTableColumns(ws As Worksheet)
    Dim response As apiResponse
    Dim tableName As String
    Dim record As Object
    Dim flds As String, typ As String, leng As String
    
    On Error GoTo ErrorHandler
    
    tableName = Left(ws.Range("B3").value, 6)
    If tableName = "" Then Exit Sub
    
    If Not EnsureAuthenticated(ws) Then
        MsgBox "Authentication failed. Cannot fetch table columns.", vbCritical
        Exit Sub
    End If
    
    ' Use Doppio_Api for the GET call
    response = ExecuteMiGet("MRS001MI", "LstFieldInfo", "FILE=" & tableName)
    
    ' Initialize column collections
    m_obj_ColumnNames.Initialize
    m_obj_ColumnDescriptions.Initialize
    m_obj_ColumnTypes.Initialize
    m_obj_ColumnConditions.Initialize
    m_obj_ColumnDirections.Initialize
    
    If response.success And Not response.records Is Nothing Then
        For Each record In response.records
            m_obj_ColumnNames.Add record.item("FLNA")
            
            flds = record.item("FRL1") & vbCrLf
            typ = record.item("FLTY")
            leng = record.item("FLLE")
            
            m_obj_ColumnDescriptions.Add flds & typ & leng
            m_obj_ColumnTypes.Add typ
            m_obj_ColumnConditions.Add False
            m_obj_ColumnDirections.Add "O"
        Next record
    End If
    
    If ws.Range("B4").value = "*" Then
        AutoFit_ColumnsAndRows True, False
        AutoFit_ColumnsAndRows False, False
    Else
        AutoFit_ColumnsAndRows False, False
    End If
    
    Exit Sub
    
ErrorHandler:
    Debug.Print "EXPORTMI_GetTableColumns: ERROR - " & Err.description
End Sub


' ============================================================================
' MAIN PROCESS
' ============================================================================

Public Sub EXPORTMI_Process()
    Dim ws As Worksheet
    Dim startTime As Single
    Dim transactions As String
    Dim response As apiResponse
    
    On Error GoTo ErrorHandler
    
    Set ws = ActiveSheet
            
    ws.Calculate
    DoEvents
    ResetCountFormat ws
    UI_ShowPleaseWait "Please Wait... Calling API"
    
    ' Authenticate and sync config
    If Not EnsureAuthenticated(ws) Then
        UI_KillPleaseWait
        MsgBox "Authentication failed. Please check your environment settings.", vbCritical
        Exit Sub
    End If
    
    ' Parse headers
    EXPORTMI_ParseFieldsIntoColumns
    
    ' Fetch column metadata
    EXPORTMI_GetTableColumns ws
    
    startTime = Timer
    
    ' Build transactions JSON and execute via Doppio_Api
    transactions = EXPORTMI_BuildTransactions(ws)
    response = ExecuteMiBulk("EXPORTMI", transactions)
    
    ' Process results
    If response.success Then
        EXPORTMI_ProcessResults response, ws
        CheckMaxRecordsWarning ws
    Else
        MsgBox "EXPORTMI failed: " & response.errorMessage, vbCritical
    End If
    
    UI_DisplayElapsedTime startTime, ws
    AutoFit_Click
    
    Exit Sub
    
ErrorHandler:
    UI_KillPleaseWait
    MsgBox "Error in EXPORTMI_Process: " & Err.description, vbCritical
End Sub


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
    
    separator = GetSeparatorFromSettings()
    
    Set results = response.results
    If results Is Nothing Then Exit Sub
    
    ' Count total records across all result items
    recordCount = 0
    For Each resultItem In results
        On Error Resume Next
        Set records = resultItem.item("records")
        If Not records Is Nothing Then
            recordCount = recordCount + records.count
        End If
        On Error GoTo ErrorHandler
    Next resultItem
    
    If recordCount = 0 Then
        ' Check for error message
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
    
    ' Determine field count from first record
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
    
    ' Build data array
    ReDim dataArray(1 To recordCount, 1 To fieldCount)
    
    row = 1
    For Each resultItem In results
        On Error Resume Next
        Set records = resultItem.item("records")
        On Error GoTo ErrorHandler
        
        If Not records Is Nothing Then
            For Each record In records
                replValue = record.item("REPL")
                replValue = Replace(replValue, "_�_", "-")
                replValue = Replace(replValue, Chr(194) & Chr(160), "")
                
                replParts = Split(replValue, separator)
                
                For i = 0 To UBound(replParts)
                    If i < fieldCount Then
                        dataArray(row, i + 1) = replParts(i)
                    End If
                Next i
                
                row = row + 1
            Next record
        End If
    Next resultItem
    
    ' Write to sheet (Row 9, Column B onwards)
    ws.Range(ws.Cells(9, 2), ws.Cells(9 + recordCount - 1, 1 + fieldCount)).value = dataArray
    
    AutoFit_Click
    
    Exit Sub
    
ErrorHandler:
    Debug.Print "EXPORTMI_ProcessResults: ERROR - " & Err.description
End Sub


