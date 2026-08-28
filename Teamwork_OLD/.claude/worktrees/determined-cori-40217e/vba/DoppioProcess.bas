Attribute VB_Name = "DoppioProcess"
''
' DoppioProcess - Modernized Process Functions
' Replaces the old Process_Click with new modular methods
'
' @module DoppioProcess
' @version 2.0
'
' Note: Uses COLOR_ERROR and COLOR_SUCCESS constants from DoppioCore
''
Option Explicit

' Module-level variables
Private m_CurrentRow As Long
Private m_Worksheet As Worksheet

''
' Main Process Click - Modernized version
' Uses new DoppioAuth, DoppioConfig, DoppioHttp, DoppioApi modules
''
Public Sub Process_Click()
    Dim apiType As String
    Dim startTime As Single
    Dim lastColumn As Integer
    Dim ws As Worksheet
    
    On Error GoTo ErrorHandler
    
    Set ws = ActiveSheet
    Set m_Worksheet = ws
    m_CurrentRow = 9
    
    ' Update version display
    DoppioUI.UI_UpdateVersion
    SetFormulasAndFormatting_New ws
    
    ' Validate environment selection
    DoppioAuth.Auth_ValidateSelectedEnvironment
    
    ' Force calculation to ensure all formulas are up to date
    ws.Calculate
    DoEvents
    
    ' Check if environment is selected
    If ws.Range("Environment").value = "" Or ws.Range("Environment").value = "Access requested" Then
        MsgBox "Please select a valid environment.", vbExclamation
        Exit Sub
    End If
    
    ' Get API type
    apiType = ws.Cells(2, 2).value
    If apiType = "" Then apiType = "API"
    
    ' Handle IDM separately
    If apiType = "IDM" Then
'        Doppio.IDM_Process  ' Use existing IDM_Process for now
        Exit Sub
    End If
    
    startTime = Timer
    
    ' Clear status column
    DoppioUI.UI_ClearStatus
    
    ' Setup freeze panes
    SetupFreezePanes ws
    
    ' Ensure settings are loaded
    EnsureSettingsLoaded
    
    ' Handle multi-output API (sets G5 formula and calculates)
    HandleMultiOutput ws
    
    ' IMPORTANT: Sync from Doppio module to ensure we use correct environment
    SyncConfigFromDoppioModule
    
    ' Get company/division from sheet
    DoppioConfig.Config_Company = ws.Range("Company").value
    DoppioConfig.Config_Division = ws.Range("Division").value
    
    UpdateEnvironmentCache DoppioConfig.Config_SelectedEnvironment, _
                           "", _
                           "", _
                           DoppioConfig.Config_Company, _
                           DoppioConfig.Config_Division
    
    ' Authenticate using OLD Doppio method to ensure correct environment
    Doppio.Tenant_Token
    
    ' Process based on API type
    Select Case apiType
        Case "API"
            ProcessMITransactions ws, startTime
        Case "IPS"
            ProcessIPSTransactions ws, startTime
        Case "XtendM3"
            ProcessXtendM3Transactions ws, startTime
        Case Else
            MsgBox "Unknown API type: " & apiType, vbExclamation
    End Select
    
    ' Cleanup
    Application.Calculation = xlCalculationAutomatic
    DoppioUI.UI_DisplayElapsedTime startTime, ws
    DoppioUI.UI_KillPleaseWait
    
    m_CurrentRow = 9
    Exit Sub
    
ErrorHandler:
    Application.Calculation = xlCalculationAutomatic
    DoppioUI.UI_KillPleaseWait
    MsgBox "Error in Process_Click: " & Err.description, vbCritical
    m_CurrentRow = 9
End Sub

''
' Sync configuration from old Doppio module to new DoppioConfig
' This ensures we always use the current environment settings
''
Private Sub SyncConfigFromDoppioModule()
    On Error Resume Next
    
    ' Sync tokens
    DoppioConfig.Config_AccessToken = Doppio.m_s_AccessToken
    DoppioConfig.Config_RefreshToken = Doppio.m_s_RefreshToken
    DoppioConfig.Config_TokenType = Doppio.m_s_TokenType
    
    ' Sync user info
    DoppioConfig.Config_SelectedEnvironment = Doppio.m_s_SelectedEnvironment
    DoppioConfig.Config_M3User = Doppio.m_s_M3user
    DoppioConfig.Config_Company = Doppio.m_s_Company
    DoppioConfig.Config_Division = Doppio.m_s_Division
    
    On Error GoTo 0
    
    #If DEBUG_MODE Then
        Debug.Print "SyncConfigFromDoppioModule:"
        Debug.Print "  Doppio.m_s_MainUrl = " & Doppio.m_s_MainUrl
        Debug.Print "  Doppio.m_s_AccessToken = " & IIf(Len(Doppio.m_s_AccessToken) > 0, "Yes (" & Len(Doppio.m_s_AccessToken) & " chars)", "No")
        Debug.Print "  Doppio.m_s_M3user = " & Doppio.m_s_M3user
    #End If
End Sub

''
' Ensure user is authenticated
''
Private Function EnsureAuthenticated(ws As Worksheet) As Boolean
    ' Try to load cached environment
    Dim manager As EnvironmentManager
    Set manager = DoppioConfig.Config_EnvironmentManager
    
    If manager.LoadEnvironment(ws.Range("I2").value, ws) Then
        ws.Range("J3").value = 2
    End If
    
    ' Check if we need to get a token
    If DoppioConfig.Config_AccessToken = "" Or _
       ws.Range("I3").value = "" Or _
       ws.Range("I4").value = "" Then
        
        ' Use new authentication
        If Not DoppioAuth.InitializeTenantToken() Then
            EnsureAuthenticated = False
            Exit Function
        End If
    End If
    
    ' Verify we have what we need
    If ws.Range("I2").value = "Access requested" Or DoppioConfig.Config_M3User = "" Then
        EnsureAuthenticated = False
        Exit Function
    End If
    
    EnsureAuthenticated = True
End Function

''
' Process MI API transactions using new methods
''
Private Sub ProcessMITransactions(ws As Worksheet, startTime As Single)
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
    Dim response As apiResponse
    Dim settings As ApiSettings
    
    On Error GoTo ErrorHandler
    
    ' Get API settings
    settings = DoppioConfig.Config_ApiSettings
    
    strServiceName = ws.Range("API").value
    strMethod = ws.Range("Transaction").value
    
    If strServiceName = "" Or strMethod = "" Then
        MsgBox "Please enter API and Transaction.", vbExclamation
        Exit Sub
    End If
    
    lastColumn = ws.Cells(8, ws.columns.count).End(xlToLeft).column
    
    ' Initialize arrays
    ReDim inputColumns(1 To 400)
    ReDim inputFields(1 To 400)
    ReDim inputValues(1 To 400)
    
    ' Populate the key fields from row 8
    j = 0
    For i = 1 To lastColumn
        If ws.Cells(m_CurrentRow - 1, i + 1) <> "" Then
            j = j + 1
            If j > 400 Then Exit For
            inputColumns(j) = i + 1
            inputFields(j) = ws.Cells(m_CurrentRow - 1, i + 1)
        End If
    Next i
    
    ws.Calculate
    
    ' Ensure B9 has a value
    If IsEmpty(ws.Range("B9").value) Then
        ws.Range("B9").value = "?"
    End If
    
    fullBody = ""
    counter = 0
    
    ' Process each row
    While ws.Cells(m_CurrentRow, 2).value <> ""
        ' Get values from current row
        For i = 1 To lastColumn
            If inputColumns(i) <> 0 Then
                Dim cellValue As Variant
                cellValue = ws.Cells(m_CurrentRow, inputColumns(i))
                
                If isError(cellValue) Then
                    cellValue = ""
                End If
                
                ' Handle the "?" placeholder
                If cellValue = "?" And m_CurrentRow = 9 And i = 1 Then
                    cellValue = ""
                End If
                
                inputValues(i) = CStr(cellValue)
            End If
        Next i
        
        ' Build transaction body
        body = BuildMITransactionBody(strMethod, inputFields, inputValues)
        
        If fullBody = "" Then
            fullBody = body
        Else
            fullBody = fullBody & "," & body
        End If
        
        m_CurrentRow = m_CurrentRow + 1
        counter = counter + 1
        
        ' Check if it's time to make the API call (batch size reached)
        If counter >= settings.maxbulk Then
            body = "{""program"":""" & strServiceName & """,""transactions"":[" & fullBody & "]}"
            fullBody = ""
            
            DoppioUI.UI_ShowPleaseWait "Please Wait... Calling API (batch " & Int(m_CurrentRow / settings.maxbulk) & ")"
            
            ' Use new API call
            response = ExecuteMIBulkCall(strServiceName, body)
            
            ' Process results
            ProcessMIResults response, ws, m_CurrentRow - counter
            
            counter = 0
        End If
    Wend
    
    ' Process remaining transactions
    If counter > 0 Then
        m_CurrentRow = m_CurrentRow + (settings.maxbulk - counter)
        body = "{""program"":""" & strServiceName & """,""transactions"":[" & fullBody & "]}"
        
        DoppioUI.UI_ShowPleaseWait "Please Wait... Calling API (final batch)"
        
        response = ExecuteMIBulkCall(strServiceName, body)
        ProcessMIResults response, ws, m_CurrentRow - settings.maxbulk
    End If
    
    Exit Sub
    
ErrorHandler:
    DoppioUI.UI_KillPleaseWait
    MsgBox "Error processing MI transactions: " & Err.description, vbCritical
End Sub

''
' Execute MI bulk API call using new HTTP module
''
Private Function ExecuteMIBulkCall(program As String, body As String) As apiResponse
    Dim config As httpConfig
    Dim httpResponse As httpResponse
    Dim response As apiResponse
    Dim apiUrl As String
    Dim settings As ApiSettings
    Dim json As Object
    Dim results As Object
    Dim firstResult As Object
    Dim records As Object
    Dim wsCompany As String
    Dim wsDivision As String
    
    On Error GoTo ErrorHandler
    
    settings = DoppioConfig.Config_ApiSettings
    
    ' Get company/division from worksheet (these override module-level values)
    wsCompany = m_Worksheet.Range("Company").value
    wsDivision = m_Worksheet.Range("Division").value
    
    ' Build URL - USE DOPPIO.M_S_MAINURL for correct environment
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
    
    ' Configure HTTP request
    config.url = apiUrl
    config.method = HttpMethod_POST
    config.contentType = "application/json; charset=UTF-8"
    config.AcceptType = "application/json; charset=UTF-8"
    config.authHeader = Doppio.m_s_TokenType & " " & Doppio.m_s_AccessToken
    config.TimeoutSeconds = settings.MaxTimeout
    config.body = body
    
    #If DEBUG_MODE Then
        Debug.Print "ExecuteMIBulkCall: URL = " & apiUrl
    #End If
    
    ' Execute request
    httpResponse = DoppioHttp.ExecuteRequest(config)
    
    #If DEBUG_MODE Then
        Debug.Print "ExecuteMIBulkCall: HTTP Status = " & httpResponse.statusCode
        Debug.Print "ExecuteMIBulkCall: Response length = " & Len(httpResponse.body)
    #End If
    
    ' Handle unauthorized - try to refresh token
    If httpResponse.IsUnauthorized Then
        #If DEBUG_MODE Then
            Debug.Print "ExecuteMIBulkCall: Unauthorized, attempting refresh..."
        #End If
        If DoppioAuth.HandleUnauthorized() Then
            ' Retry with new token
            config.authHeader = Doppio.m_s_TokenType & " " & Doppio.m_s_AccessToken
            httpResponse = DoppioHttp.ExecuteRequest(config)
        End If
    End If
    
    ' Build response
    response.success = httpResponse.success
    response.data = httpResponse.body
    response.errorMessage = httpResponse.errorMessage
    
    ' Parse JSON response
    If httpResponse.success And Len(httpResponse.body) > 0 Then
        Set json = JsonConverter.ParseJson(httpResponse.body)
        
        If Not json Is Nothing Then
            ' Get the results array
            Set results = json.item("results")
            Set response.results = results
            
            ' Count total records across all results
            response.recordCount = 0
            If Not results Is Nothing Then
                Dim resultItem As Object
                For Each resultItem In results
                    On Error Resume Next
                    Set records = resultItem.item("records")
                    If Not records Is Nothing Then
                        response.recordCount = response.recordCount + records.count
                    End If
                    On Error GoTo ErrorHandler
                Next resultItem
            End If
            
            #If DEBUG_MODE Then
                Debug.Print "ExecuteMIBulkCall: Parsed " & response.recordCount & " records"
            #End If
        End If
    End If
    
    ExecuteMIBulkCall = response
    Exit Function
    
ErrorHandler:
    #If DEBUG_MODE Then
        Debug.Print "ExecuteMIBulkCall: ERROR - " & Err.description
    #End If
    response.success = False
    response.errorMessage = Err.description
    ExecuteMIBulkCall = response
End Function

''
' Build MI transaction body (same logic as BulkMIBody_create)
''
Private Function BuildMITransactionBody(method As String, inputFields() As String, inputValues() As String) As String
    Dim body As String
    Dim i As Long
    Dim fieldName As String
    Dim fieldValue As String
    
    body = "{""transaction"":""" & method & """,""record"":{"
    
    Dim firstField As Boolean
    firstField = True
    
    For i = LBound(inputFields) To UBound(inputFields)
        If inputFields(i) <> "" Then
            fieldName = inputFields(i)
            fieldValue = inputValues(i)
            
            ' Escape special characters in value
            fieldValue = Replace(fieldValue, "\", "\\")
            fieldValue = Replace(fieldValue, """", "\""")
            fieldValue = Replace(fieldValue, vbCr, "")
            fieldValue = Replace(fieldValue, vbLf, "")
            
            If Not firstField Then
                body = body & ","
            End If
            
            body = body & """" & fieldName & """:""" & fieldValue & """"
            firstField = False
        End If
    Next i
    
    body = body & "}}"
    
    BuildMITransactionBody = body
End Function

''
' Process MI API results and write to worksheet
''
Private Sub ProcessMIResults(response As apiResponse, ws As Worksheet, startRow As Long)
    Dim results As Object
    Dim resultItem As Object
    Dim records As Object
    Dim record As Object
    Dim rowNum As Long
    Dim colNum As Long
    Dim lastColumn As Integer
    Dim fieldName As String
    Dim fieldValue As Variant
    Dim headerRow As Long
    Dim key As Variant
    Dim statusCell As Range
    Dim errorMessage As String
    Dim recordCount As Long
    
    On Error GoTo ErrorHandler
    
    headerRow = 8  ' Row with field names
    lastColumn = ws.Cells(headerRow, ws.columns.count).End(xlToLeft).column
    
    #If DEBUG_MODE Then
        Debug.Print "ProcessMIResults: startRow=" & startRow & ", lastColumn=" & lastColumn
    #End If
    
    ' If request failed, mark rows as failed
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
    
    If response.results Is Nothing Then
        #If DEBUG_MODE Then
            Debug.Print "ProcessMIResults: No results object"
        #End If
        Exit Sub
    End If
    
    Set results = response.results
    rowNum = startRow
    
    #If DEBUG_MODE Then
        Debug.Print "ProcessMIResults: Processing " & results.count & " result items"
    #End If
    
    ' Loop through each result (transaction)
    For Each resultItem In results
        Set statusCell = ws.Cells(rowNum, 1)
        
        ' First check if there's an errorMessage at the result level
        errorMessage = ""
        On Error Resume Next
        errorMessage = resultItem.item("errorMessage")
        On Error GoTo ErrorHandler
        
        If errorMessage <> "" Then
            ' Transaction failed - show error
            statusCell.value = "NOK " & errorMessage
            statusCell.Font.Color = COLOR_ERROR
            rowNum = rowNum + 1
        Else
            ' Get records for this transaction
            On Error Resume Next
            Set records = resultItem.item("records")
            recordCount = 0
            If Not records Is Nothing Then
                recordCount = records.count
            End If
            On Error GoTo ErrorHandler
            
            If records Is Nothing Or recordCount = 0 Then
                ' No records and no error - mark as OK (might be an update/delete with no output)
                statusCell.value = "OK"
                statusCell.Font.Color = COLOR_SUCCESS
                rowNum = rowNum + 1
            Else
                #If DEBUG_MODE Then
                    Debug.Print "ProcessMIResults: Transaction has " & recordCount & " records"
                #End If
                
                ' Loop through each record
                For Each record In records
                    Set statusCell = ws.Cells(rowNum, 1)
                    
                    ' Check for error message in ZZUSID field
                    Dim zzusidError As String
                    zzusidError = ""
                    
                    On Error Resume Next
                    zzusidError = record.item("ZZUSID")
                    On Error GoTo ErrorHandler
                    
                    ' If ZZUSID has content and doesn't start with spaces, it's an error
                    If zzusidError <> "" And Left(zzusidError, 3) <> "   " Then
                        statusCell.value = "NOK " & Trim(zzusidError)
                        statusCell.Font.Color = COLOR_ERROR
                    Else
                        statusCell.value = "OK"
                        statusCell.Font.Color = COLOR_SUCCESS
                        
                        ' Write output fields to the row
                        ' Loop through columns and match field names
                        For colNum = 2 To lastColumn
                            fieldName = ws.Cells(headerRow, colNum).value
                            If fieldName <> "" Then
                                On Error Resume Next
                                fieldValue = record.item(fieldName)
                                If Err.Number = 0 Then
                                    ' Only write if field exists in response
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
    
    #If DEBUG_MODE Then
        Debug.Print "ProcessMIResults: Completed, final row = " & rowNum
    #End If
    Exit Sub
    
ErrorHandler:
    #If DEBUG_MODE Then
        Debug.Print "ProcessMIResults: ERROR - " & Err.description
    #End If
    ' Try to continue
    Resume Next
End Sub

''
' Process IPS (Web Services) transactions
''
Private Sub ProcessIPSTransactions(ws As Worksheet, startTime As Single)
    ' For now, delegate to old method
    ' TODO: Implement using new modules
    MsgBox "IPS processing not yet migrated to new modules. Using old method.", vbInformation
    'Doppio.Process_Click
End Sub

''
' Process XtendM3 transactions
''
Private Sub ProcessXtendM3Transactions(ws As Worksheet, startTime As Single)
    ' For now, delegate to old method
    ' TODO: Implement using new modules
    MsgBox "XtendM3 processing not yet migrated to new modules. Using old method.", vbInformation
    'Doppio.Process_Click
End Sub

''
' Setup freeze panes
''
Private Sub SetupFreezePanes(ws As Worksheet)
    ws.Cells(1, 1).Select
    ActiveWindow.FreezePanes = False
    ws.Range("C9").Select
    ActiveWindow.FreezePanes = True
    
    Application.Calculation = xlCalculationManual
    Application.GoTo Reference:="R9C2", Scroll:=True
End Sub

''
' Ensure settings are loaded
''
Private Sub EnsureSettingsLoaded()
    Dim settings As ApiSettings
    settings = DoppioConfig.Config_ApiSettings
    
    If settings.maxbulk <= 1 Then
        DoppioConfig.Config_LoadSettingsFromSheet
    End If
    
    If settings.refreshSeconds = 0 Then
        DoppioConfig.Config_LoadSettingsFromSheet
    End If
End Sub

''
' Handle multi-output API setup
' Sets G5 formula to lookup Single/Multiple from Transactions sheet
''
Private Sub HandleMultiOutput(ws As Worksheet)
    Dim wsTrans As Worksheet
    Dim currentApi As String
    Dim found As Range
    
    On Error Resume Next
    
    ' --- NEW LOGIC START ---
    ' Check if the Transactions sheet is loaded for the current API
    currentApi = ws.Range("API").value
    
    If currentApi <> "" Then
        Set wsTrans = ThisWorkbook.Sheets("Transactions")
        
        ' Look for the API name in Column A of the Transactions sheet
        Set found = wsTrans.Range("A:A").Find(What:=currentApi, LookIn:=xlValues, LookAt:=xlWhole)
        
        ' If the API is not found, reload the transactions
        If found Is Nothing Then
            ' Call the legacy wrapper or the new method directly
            GetTransactions_Click
        End If
    End If
    ' --- NEW LOGIC END ---
    
    ' Always set the formula to ensure it's correct
    With ws.Range("G5")
        .NumberFormat = "General"
        .Formula = "=IFNA(VLOOKUP(Transaction,Transactions!B:C,2,False),"""")"
    End With
    
    ' Force calculation
    Application.Calculate
    ws.Calculate
    DoEvents
    
    ' If Multi-output, clear rows 10 and below
    If ws.Range("G5").value = "M" Then
        ws.Rows("10:" & ws.Rows.count).ClearContents
    End If
    
    On Error GoTo 0
End Sub

' =============================================================================
' Helper function for URL encoding (if not already available)
' =============================================================================
Private Function Core_UrlEncode(text As String) As String
    Dim i As Long
    Dim c As String
    Dim result As String
    
    For i = 1 To Len(text)
        c = Mid(text, i, 1)
        Select Case c
            Case "A" To "Z", "a" To "z", "0" To "9", "-", "_", ".", "~"
                result = result & c
            Case " "
                result = result & "+"
            Case Else
                result = result & "%" & Right("0" & Hex(Asc(c)), 2)
        End Select
    Next i
    
    Core_UrlEncode = result
End Function




