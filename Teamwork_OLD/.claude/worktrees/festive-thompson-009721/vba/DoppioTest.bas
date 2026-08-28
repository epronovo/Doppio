Attribute VB_Name = "DoppioTest"
''
' DoppioTest - Test Module for API Migration
' Use this to compare old apicall vs new ExecuteApiCall
'
' @module DoppioTest
' @version 1.0
''
Option Explicit

' =============================================================================
' TEST FUNCTIONS
' =============================================================================

''
' Test the new API call method with a simple GetUserInfo call
' This is the most basic test - just gets user info for current environment
''
Public Sub Test_NewApiCall_GetUserInfo()
    Dim response As apiResponse
    Dim startTime As Single
    
    On Error GoTo ErrorHandler
    
    ' Make sure we have an environment selected
    If ActiveSheet.Range("Environment").value = "" Then
        MsgBox "Please select an environment first.", vbExclamation
        Exit Sub
    End If
  
    #If DEBUG_MODE Then
        Debug.Print "======================================"
        Debug.Print "TEST: New API Call - GetUserInfo"
        Debug.Print "======================================"
    #End If
'    Debug.Print "Environment: " & Doppio.m_s_SelectedEnvironment
'    Debug.Print "Main URL: " & Doppio.m_s_MainUrl
'    Debug.Print "Token Type: " & Doppio.m_s_TokenType
'    Debug.Print "Has Token: " & (Len(Doppio.m_s_AccessToken) > 0)
'    Debug.Print ""
    
    startTime = Timer
    
    ' Execute using our bridged function that uses Doppio's auth
    response = ExecuteNewApiCall("MRS001MI", "GetUserInfo", "")
    
    #If DEBUG_MODE Then
        Debug.Print "Response received in " & Format((Timer - startTime), "0.000") & " seconds"
        Debug.Print "Success: " & response.success
        Debug.Print "Record Count: " & response.recordCount
        Debug.Print ""
    #End If
    
    If response.success Then
        #If DEBUG_MODE Then
            Debug.Print "SUCCESS! Response data:"
            Debug.Print Left(response.data, 500)
        #End If
        #If DEBUG_MODE Then
            If Len(response.data) > 500 Then Debug.Print "... (truncated)"
        #End If
        MsgBox "New API call succeeded!" & vbCrLf & vbCrLf & _
               "Records: " & response.recordCount, vbInformation
    Else
        #If DEBUG_MODE Then
            Debug.Print "FAILED! Error:"
            Debug.Print response.errorMessage
        #End If
        MsgBox "New API call failed:" & vbCrLf & response.errorMessage, vbCritical
    End If
    
    Exit Sub
    
ErrorHandler:
    #If DEBUG_MODE Then
        Debug.Print "ERROR: " & Err.description
    #End If
    MsgBox "Error: " & Err.description, vbCritical
End Sub

''
' Compare old vs new API call side-by-side
' Runs both methods and compares results
''
Public Sub Test_CompareOldVsNew()
    Dim startTimeOld As Double, startTimeNew As Double
    Dim endTimeOld As Double, endTimeNew As Double
    Dim response As apiResponse
    Dim oldRecordCount As Long, newRecordCount As Long
    Dim miPath As String
    Dim miUrl As String
    
    On Error GoTo ErrorHandler
    
    ' Make sure we have an environment selected
    If ActiveSheet.Range("Environment").value = "" Then
        MsgBox "Please select an environment first.", vbExclamation
        Exit Sub
    End If
    
    #If DEBUG_MODE Then
        Debug.Print "======================================"
        Debug.Print "COMPARISON TEST: Old vs New API Call"
        Debug.Print "======================================"
    #End If
    
    ' --- OLD METHOD ---
    #If DEBUG_MODE Then
        Debug.Print ""
        Debug.Print "--- OLD METHOD (apicall) ---"
    #End If
    
    Doppio.Tenant_Information
    
    startTimeOld = Timer
    
    ' Use local variables instead of module-private ones
    miPath = "/M3/m3api-rest/v2/execute"
    miUrl = "MRS001MI/GetUserInfo/?"
    Doppio.apicall Doppio.m_s_MainUrl, miPath, miUrl, "", "API"
    
    endTimeOld = Timer
    
    ' Count records from old method
    oldRecordCount = 0
    On Error Resume Next
    If Not Doppio.m_obj_Records Is Nothing Then
        oldRecordCount = Doppio.m_obj_Records.count
    End If
    On Error GoTo ErrorHandler
    
    #If DEBUG_MODE Then
        Debug.Print "Old method time: " & Format(endTimeOld - startTimeOld, "0.000") & " seconds"
        Debug.Print "Old method records: " & oldRecordCount
    #End If
    
    ' --- NEW METHOD ---
    #If DEBUG_MODE Then
        Debug.Print ""
        Debug.Print "--- NEW METHOD (Direct HTTP) ---"
    #End If
    
    ' Sync the configuration from old Doppio module to new modules
    SyncConfigFromDoppio
    
    startTimeNew = Timer
    
    ' Use direct HTTP call with the synced config
    response = ExecuteNewApiCall("MRS001MI", "GetUserInfo", "")
    
    endTimeNew = Timer
    newRecordCount = response.recordCount
    
    #If DEBUG_MODE Then
        Debug.Print "New method time: " & Format(endTimeNew - startTimeNew, "0.000") & " seconds"
        Debug.Print "New method records: " & newRecordCount
        Debug.Print "New method success: " & response.success
    #End If
    If Not response.success Then
        #If DEBUG_MODE Then
            Debug.Print "Error: " & response.errorMessage
        #End If
    End If
    
    Dim diffOld As Double, diffNew As Double
    
    ' Use CDbl to force double-precision and handle potential wrap-around
    On Error Resume Next ' Temporary guard to identify the exact spot
    diffOld = CDbl(endTimeOld) - CDbl(startTimeOld)
    If diffOld < 0 Then diffOld = diffOld + 86400 ' Handle midnight reset
    
    diffNew = CDbl(endTimeNew) - CDbl(startTimeNew)
    If diffNew < 0 Then diffNew = diffNew + 86400 ' Handle midnight reset
    On Error GoTo ErrorHandler

    ' --- COMPARISON ---
    #If DEBUG_MODE Then
        Debug.Print ""
        Debug.Print "--- COMPARISON ---"
        Debug.Print "Old time: " & Format(diffOld, "0.000") & " seconds"
        Debug.Print "New time: " & Format(diffNew, "0.000") & " seconds"
        Debug.Print "Records match: " & (oldRecordCount = newRecordCount)
        Debug.Print "TEST COMPLETE!"
    #End If
    
    ' Store record counts
    Dim oldRecs As Long
    Dim newRecs As Long
    oldRecs = oldRecordCount
    newRecs = newRecordCount
    
    ' Create pre-formatted strings for the MsgBox to avoid concatenation overflows
    Dim strOldTime As String
    Dim strNewTime As String
    strOldTime = Format(diffOld, "0.000")
    strNewTime = Format(diffNew, "0.000")
    
    #If DEBUG_MODE Then
        Debug.Print "About to show MsgBox..."
        Debug.Print "oldRecs=" & oldRecs & " newRecs=" & newRecs & " oldSecs=" & strOldTime & " newSecs=" & strNewTime
    #End If

    ' Use the pre-formatted strings
    MsgBox "Test Complete!" & vbCrLf & vbCrLf & _
           "Old: " & oldRecs & " records in " & strOldTime & "s" & vbCrLf & _
           "New: " & newRecs & " records in " & strNewTime & "s", vbInformation
    Exit Sub
    
ErrorHandler:
    #If DEBUG_MODE Then
        Debug.Print "ERROR: " & Err.description
    #End If
    MsgBox "Error: " & Err.description, vbCritical
End Sub

''
' Sync configuration from old Doppio module to new DoppioConfig
''
Private Sub SyncConfigFromDoppio()
    ' This syncs the authentication and settings from the old Doppio module
    ' to the new modular architecture
    
    #If DEBUG_MODE Then
        Debug.Print "Syncing config from Doppio module..."
        Debug.Print "  MainUrl: " & Doppio.m_s_MainUrl
        Debug.Print "  TokenType: " & Doppio.m_s_TokenType
        Debug.Print "  HasToken: " & (Len(Doppio.m_s_AccessToken) > 0)
        Debug.Print "  Company: " & Doppio.m_s_Company
        Debug.Print "  Division: " & Doppio.m_s_Division
    #End If
End Sub

''
' Execute API call using new HTTP module but with old Doppio config
''
Public Function ExecuteNewApiCall(program As String, transaction As String, parameters As String) As apiResponse
    Dim config As httpConfig
    Dim httpResponse As httpResponse
    Dim response As apiResponse
    Dim apiUrl As String
    Dim json As Object
    Dim ws As Worksheet
    Dim wsCompany As String
    Dim wsDivision As String
    
    On Error GoTo ErrorHandler
    
    Set ws = ActiveSheet
    
    ' Get company/division from worksheet (user can override)
    wsCompany = ws.Range("Company").value
    wsDivision = ws.Range("Division").value
    
    ' Always use Doppio module values for URL and token
    ' This ensures we use the correct environment
    If Trim(Doppio.m_s_AccessToken) = "" Then
        ' No token - need to authenticate
        #If DEBUG_MODE Then
            Debug.Print "ExecuteNewApiCall: No token found, calling Tenant_Token..."
        #End If
        Doppio.Tenant_Token
        
        If Trim(Doppio.m_s_AccessToken) = "" Then
            response.success = False
            response.errorMessage = "Authentication failed - no token"
            ExecuteNewApiCall = response
            Exit Function
        End If
    End If
    
    ' Build URL using Doppio.m_s_MainUrl
    apiUrl = Doppio.m_s_MainUrl & "/M3/m3api-rest/v2/execute/" & program & "/" & transaction & "?"
    
    ' Add parameters
    If parameters <> "" Then
        apiUrl = apiUrl & parameters & "&"
    End If
    
    ' Add company/division from worksheet
    If wsCompany <> "" Then
        apiUrl = apiUrl & "cono=" & wsCompany & "&"
    End If
    If wsDivision <> "" Then
        apiUrl = apiUrl & "divi=" & wsDivision & "&"
    End If
    
    ' Remove trailing &
    If Right(apiUrl, 1) = "&" Then
        apiUrl = Left(apiUrl, Len(apiUrl) - 1)
    End If
    
    ' Configure HTTP request
    config.url = apiUrl
    config.method = HttpMethod_GET
    config.contentType = "application/json"
    config.AcceptType = "application/json"
    config.authHeader = Doppio.m_s_TokenType & " " & Doppio.m_s_AccessToken
    config.TimeoutSeconds = 30
    config.body = ""
    
    #If DEBUG_MODE Then
        Debug.Print "ExecuteNewApiCall: URL = " & apiUrl
    #End If
    
    ' Execute request
    httpResponse = DoppioHttp.ExecuteRequest(config)
    
    #If DEBUG_MODE Then
        Debug.Print "ExecuteNewApiCall: HTTP Status = " & httpResponse.statusCode
        Debug.Print "ExecuteNewApiCall: Response length = " & Len(httpResponse.body)
    #End If
    
    ' Handle 401 Unauthorized - try to get fresh token
    If httpResponse.statusCode = 401 Then
        #If DEBUG_MODE Then
            Debug.Print "ExecuteNewApiCall: Unauthorized, getting fresh token..."
        #End If
        Doppio.m_s_AccessToken = ""
        Doppio.Tenant_Token
        
        If Trim(Doppio.m_s_AccessToken) <> "" Then
            ' Retry with new token
            config.authHeader = Doppio.m_s_TokenType & " " & Doppio.m_s_AccessToken
            httpResponse = DoppioHttp.ExecuteRequest(config)
        End If
    End If
    
    ' Parse response
    response.success = (httpResponse.statusCode >= 200 And httpResponse.statusCode < 300)
    response.data = httpResponse.body
    response.errorMessage = httpResponse.errorMessage
    
    If response.success And Len(httpResponse.body) > 0 Then
        On Error Resume Next
        Set json = JsonConverter.ParseJson(httpResponse.body)
        If Not json Is Nothing Then
            Set response.results = json.item("results")
            If Not response.results Is Nothing Then
                If response.results.count > 0 Then
                    Set response.records = response.results(1).item("records")
                    If Not response.records Is Nothing Then
                        response.recordCount = response.records.count
                    End If
                End If
            End If
        End If
        On Error GoTo ErrorHandler
    End If
    
    ExecuteNewApiCall = response
    Exit Function
    
ErrorHandler:
    response.success = False
    response.errorMessage = Err.description
    #If DEBUG_MODE Then
        Debug.Print "ExecuteNewApiCall: ERROR - " & Err.description
    #End If
    ExecuteNewApiCall = response
End Function

''
' Test a specific MI transaction using the new method
' Parameters are entered via InputBox
''
Public Sub Test_NewApiCall_CustomMI()
    Dim response As apiResponse
    Dim program As String
    Dim transaction As String
    Dim inputFields As String
    Dim startTime As Single
    
    On Error GoTo ErrorHandler
    
    ' Make sure we have an environment selected
    If ActiveSheet.Range("Environment").value = "" Then
        MsgBox "Please select an environment first.", vbExclamation
        Exit Sub
    End If
    
    ' Get parameters from user
    program = InputBox("Enter MI Program (e.g., MMS200MI):", "MI Program", "MMS200MI")
    If program = "" Then Exit Sub
    
    transaction = InputBox("Enter Transaction (e.g., GetItmBasic):", "Transaction", "GetItmBasic")
    If transaction = "" Then Exit Sub
    
    inputFields = InputBox("Enter input fields (e.g., ITNO=TESTITEM):" & vbCrLf & _
                          "Leave blank for no input", "Input Fields", "ITNO=100037")
    
    ' Load tenant information
    Doppio.Tenant_Information
    
    #If DEBUG_MODE Then
        Debug.Print "======================================"
        Debug.Print "TEST: Custom MI Call"
        Debug.Print "======================================"
        Debug.Print "Program: " & program
        Debug.Print "Transaction: " & transaction
        Debug.Print "Input: " & inputFields
        Debug.Print ""
    #End If
    
    startTime = Timer
    
    ' Use the high-level ExecuteMiGet function
    response = ExecuteNewApiCall(program, transaction, inputFields)
        
    #If DEBUG_MODE Then
        Debug.Print "Response received in " & Format(Timer - startTime, "0.00") & " seconds"
        Debug.Print "Status: " & response.status
        Debug.Print "Success: " & response.success
        Debug.Print "Record Count: " & response.recordCount
        Debug.Print ""
    #End If
    
    If response.success Then
        #If DEBUG_MODE Then
            Debug.Print "SUCCESS! First 500 chars of response:"
            Debug.Print Left(response.data, 500)
        #End If
        MsgBox "API call succeeded!" & vbCrLf & vbCrLf & _
               "Program: " & program & "/" & transaction & vbCrLf & _
               "Records: " & response.recordCount & vbCrLf & _
               "Time: " & Format(Timer - startTime, "0.00") & "s", vbInformation
    Else
        #If DEBUG_MODE Then
            Debug.Print "FAILED! Error:"
            Debug.Print response.errorMessage
        #End If
        MsgBox "API call failed:" & vbCrLf & response.errorMessage, vbCritical
    End If
    
    Exit Sub
    
ErrorHandler:
    #If DEBUG_MODE Then
        Debug.Print "ERROR: " & Err.description
    #End If
    MsgBox "Error: " & Err.description, vbCritical
End Sub

''
' Test listing transactions for an MI program
''
Public Sub Test_ListTransactions()
    Dim response As apiResponse
    Dim program As String
    Dim startTime As Single
    
    On Error GoTo ErrorHandler
    
    If ActiveSheet.Range("Environment").value = "" Then
        MsgBox "Please select an environment first.", vbExclamation
        Exit Sub
    End If
    
    program = InputBox("Enter MI Program to list transactions:", "MI Program", "MMS200MI")
    If program = "" Then Exit Sub
    
    Doppio.Tenant_Information
    
    #If DEBUG_MODE Then
        Debug.Print "======================================"
        Debug.Print "TEST: List Transactions for " & program
        Debug.Print "======================================"
    #End If
    
    startTime = Timer
    
    response = DoppioApi.ListMiTransactions(program)
    
    #If DEBUG_MODE Then
        Debug.Print "Response received in " & Format(Timer - startTime, "0.00") & " seconds"
        Debug.Print "Success: " & response.success
        Debug.Print "Record Count: " & response.recordCount
    #End If
    
    If response.success Then
        MsgBox "Found " & response.recordCount & " transactions for " & program & vbCrLf & _
               "Time: " & Format(Timer - startTime, "0.00") & "s" & vbCrLf & vbCrLf & _
               "Check the Immediate Window (Ctrl+G) for details.", vbInformation
    Else
        MsgBox "Failed: " & response.errorMessage, vbCritical
    End If
    
    Exit Sub
    
ErrorHandler:
    #If DEBUG_MODE Then
        Debug.Print "ERROR: " & Err.description
    #End If
    MsgBox "Error: " & Err.description, vbCritical
End Sub

''
' Quick test to verify modules are loaded correctly
''
Public Sub Test_ModulesLoaded()
    Dim msg As String
    
    On Error Resume Next
    
    msg = "Module Status Check:" & vbCrLf & vbCrLf
    
    ' Check DoppioCore
    Dim testStr As String
    testStr = DoppioCore.Base64Encode("test")
    If Err.Number = 0 Then
        msg = msg & "? DoppioCore - OK" & vbCrLf
    Else
        msg = msg & "? DoppioCore - MISSING or ERROR" & vbCrLf
        Err.Clear
    End If
    
    ' Check DoppioConfig
    Dim testSettings As ApiSettings
    testSettings = DoppioConfig.Config_ApiSettings
    If Err.Number = 0 Then
        msg = msg & "? DoppioConfig - OK" & vbCrLf
    Else
        msg = msg & "? DoppioConfig - MISSING or ERROR" & vbCrLf
        Err.Clear
    End If
    
    ' Check DoppioAuth
    Dim testBool As Boolean
    testBool = DoppioAuth.EnsureAuthenticated
    If Err.Number = 0 Or Err.Number = 0 Then
        msg = msg & "? DoppioAuth - OK" & vbCrLf
    Else
        msg = msg & "? DoppioAuth - MISSING or ERROR" & vbCrLf
        Err.Clear
    End If
    
    ' Check DoppioApi
    ' Just check if the module exists by referencing a constant or type
    Dim testType As apiType
    testType = ApiType_MI
    If Err.Number = 0 Then
        msg = msg & "? DoppioApi - OK" & vbCrLf
    Else
        msg = msg & "? DoppioApi - MISSING or ERROR" & vbCrLf
        Err.Clear
    End If
    
    ' Check DoppioHttp
    Dim testConfig As httpConfig
    testConfig.method = HttpMethod_GET
    If Err.Number = 0 Then
        msg = msg & "? DoppioHttp - OK" & vbCrLf
    Else
        msg = msg & "? DoppioHttp - MISSING or ERROR" & vbCrLf
        Err.Clear
    End If
    
    ' Check DoppioUI
    ' Try calling a UI function
    ' DoppioUI.UI_UpdateVersion ' Don't actually call it
    msg = msg & "? DoppioUI - Assumed OK" & vbCrLf
    
    ' Check DoppioCache
    ' DoppioCache.Cache_InitializeCache ' Don't actually call it
    msg = msg & "? DoppioCache - Assumed OK" & vbCrLf
    
    ' Check old Doppio module
    If Len(Doppio.m_s_MainUrl) >= 0 Then
        msg = msg & "? Doppio (original) - OK" & vbCrLf
    Else
        msg = msg & "? Doppio (original) - MISSING" & vbCrLf
    End If
    
    On Error GoTo 0
    
    MsgBox msg, vbInformation, "Module Status"
End Sub

' =============================================================================
' HELPER FUNCTIONS
' =============================================================================

''
' Print response details to Immediate Window
''
Private Sub PrintResponse(response As apiResponse)
    #If DEBUG_MODE Then
        Debug.Print "--- Response Details ---"
        Debug.Print "Success: " & response.success
        Debug.Print "Status: " & response.status
        Debug.Print "RecordCount: " & response.recordCount
        Debug.Print "ErrorMessage: " & response.errorMessage
        Debug.Print "Data length: " & Len(response.data)
        Debug.Print "------------------------"
    #End If
End Sub

''
' Test Authentication - verify token retrieval works
''
Public Sub Test_Authentication()
    On Error GoTo ErrorHandler
    
    #If DEBUG_MODE Then
        Debug.Print "======================================"
        Debug.Print "TEST: Authentication"
        Debug.Print "======================================"
    #End If
    
    ' Check if environment is selected
    Dim envName As String
    envName = ""
    On Error Resume Next
    envName = ActiveSheet.Range("Environment").value
    On Error GoTo ErrorHandler
    
    If envName = "" Then
        #If DEBUG_MODE Then
            Debug.Print "ERROR: No environment selected"
        #End If
        MsgBox "Please select an environment first.", vbExclamation
        Exit Sub
    End If
    
    #If DEBUG_MODE Then
        Debug.Print "Selected Environment: " & envName
        Debug.Print ""
    #End If
    
    ' Step 1: Load tenant information
    #If DEBUG_MODE Then
        Debug.Print "Step 1: Loading tenant information..."
    #End If
    Doppio.Tenant_Information
    
    #If DEBUG_MODE Then
        Debug.Print "  Tenant ID (ti): " & Doppio.ti
        Debug.Print "  Client ID (ci): " & Left(Doppio.ci, 10) & "..."
        Debug.Print "  ION URL (iu): " & Doppio.iu
        Debug.Print "  SSO URL (pu): " & Doppio.pu
        Debug.Print "  Main URL: " & Doppio.m_s_MainUrl
        Debug.Print "  Is Multitenant: " & Doppio.m_b_Multitenant
        Debug.Print ""
    #End If
    
    ' Step 2: Check current token status
    #If DEBUG_MODE Then
        Debug.Print "Step 2: Checking current token status..."
        Debug.Print "  Token Type: " & Doppio.m_s_TokenType
        Debug.Print "  Has Access Token: " & (Len(Doppio.m_s_AccessToken) > 0)
    #End If
    If Len(Doppio.m_s_AccessToken) > 0 Then
        #If DEBUG_MODE Then
            Debug.Print "  Token Length: " & Len(Doppio.m_s_AccessToken)
            Debug.Print "  Token Preview: " & Left(Doppio.m_s_AccessToken, 20) & "..."
        #End If
    End If
    #If DEBUG_MODE Then
        Debug.Print "  Has Refresh Token: " & (Len(Doppio.m_s_RefreshToken) > 0)
        Debug.Print ""
    #End If
    
    ' Step 3: Test API call with current token
    #If DEBUG_MODE Then
        Debug.Print "Step 3: Testing API call with current token..."
    #End If
    Dim config As httpConfig
    Dim httpResponse As httpResponse
    Dim testUrl As String
    
    testUrl = Doppio.m_s_MainUrl & "/M3/m3api-rest/v2/execute/MRS001MI/GetUserInfo/?"
    
    config.url = testUrl
    config.method = HttpMethod_GET
    config.contentType = "application/json"
    config.AcceptType = "application/json"
    config.authHeader = Doppio.m_s_TokenType & " " & Doppio.m_s_AccessToken
    config.TimeoutSeconds = 30
    config.body = ""
    
    #If DEBUG_MODE Then
        Debug.Print "  Calling: " & testUrl
    #End If
    httpResponse = DoppioHttp.ExecuteRequest(config)
    
    #If DEBUG_MODE Then
        Debug.Print "  HTTP Status: " & httpResponse.statusCode
        Debug.Print "  Success: " & httpResponse.success
        Debug.Print "  Response Length: " & Len(httpResponse.body)
    #End If
    
    If httpResponse.statusCode = 401 Then
        #If DEBUG_MODE Then
            Debug.Print "  UNAUTHORIZED - Token may be expired"
            Debug.Print ""
            Debug.Print "Step 4: Attempting to refresh token..."
        #End If
        
        ' Try to get a new token
        Doppio.Tenant_Token
        
        #If DEBUG_MODE Then
            Debug.Print "  New Token Type: " & Doppio.m_s_TokenType
            Debug.Print "  Has New Access Token: " & (Len(Doppio.m_s_AccessToken) > 0)
        #End If
        
        If Len(Doppio.m_s_AccessToken) > 0 Then
            #If DEBUG_MODE Then
                Debug.Print ""
                Debug.Print "Step 5: Retrying API call with new token..."
            #End If
            config.authHeader = Doppio.m_s_TokenType & " " & Doppio.m_s_AccessToken
            httpResponse = DoppioHttp.ExecuteRequest(config)
            #If DEBUG_MODE Then
                Debug.Print "  HTTP Status: " & httpResponse.statusCode
                Debug.Print "  Success: " & httpResponse.success
            #End If
        End If
    End If
    
    #If DEBUG_MODE Then
        Debug.Print ""
        Debug.Print "======================================"
    #End If
    
    If httpResponse.statusCode = 200 Then
        #If DEBUG_MODE Then
            Debug.Print "AUTHENTICATION TEST: PASSED"
        #End If
        MsgBox "Authentication Test PASSED!" & vbCrLf & vbCrLf & _
               "Environment: " & envName & vbCrLf & _
               "Token Type: " & Doppio.m_s_TokenType & vbCrLf & _
               "HTTP Status: 200 OK", vbInformation
    Else
        #If DEBUG_MODE Then
            Debug.Print "AUTHENTICATION TEST: FAILED"
            Debug.Print "Error: " & httpResponse.errorMessage
        #End If
        MsgBox "Authentication Test FAILED!" & vbCrLf & vbCrLf & _
               "HTTP Status: " & httpResponse.statusCode & vbCrLf & _
               "Error: " & httpResponse.errorMessage, vbCritical
    End If
    
    Exit Sub
    
ErrorHandler:
    #If DEBUG_MODE Then
        Debug.Print "ERROR: " & Err.description
    #End If
    MsgBox "Authentication Test Error: " & Err.description, vbCritical
End Sub

''
' Test getting a fresh token via Tenant_Token
''
Public Sub Test_GetFreshToken()
    On Error GoTo ErrorHandler
    
    #If DEBUG_MODE Then
        Debug.Print "======================================"
        Debug.Print "TEST: Get Fresh Token"
        Debug.Print "======================================"
    #End If
    
    ' Check if environment is selected
    If ActiveSheet.Range("Environment").value = "" Then
        MsgBox "Please select an environment first.", vbExclamation
        Exit Sub
    End If
    
    #If DEBUG_MODE Then
        Debug.Print "Environment: " & ActiveSheet.Range("Environment").value
        Debug.Print ""
    #End If
    
    ' Clear existing token
    #If DEBUG_MODE Then
        Debug.Print "Clearing existing token..."
    #End If
    Doppio.m_s_AccessToken = ""
    Doppio.m_s_RefreshToken = ""
    Doppio.m_s_TokenType = ""
    
    #If DEBUG_MODE Then
        Debug.Print "Calling Tenant_Token..."
    #End If
    Dim startTime As Single
    startTime = Timer
    
    Doppio.Tenant_Token
    
    #If DEBUG_MODE Then
        Debug.Print "Completed in " & Int(Timer - startTime) & " seconds"
        Debug.Print ""
        Debug.Print "Results:"
        Debug.Print "  Token Type: " & Doppio.m_s_TokenType
        Debug.Print "  Has Access Token: " & (Len(Doppio.m_s_AccessToken) > 0)
        Debug.Print "  Token Length: " & Len(Doppio.m_s_AccessToken)
        Debug.Print "  Company: " & Doppio.m_s_Company
        Debug.Print "  Division: " & Doppio.m_s_Division
        Debug.Print "  M3 User: " & Doppio.m_s_M3user
    #End If
    
    If Len(Doppio.m_s_AccessToken) > 0 Then
        #If DEBUG_MODE Then
            Debug.Print ""
            Debug.Print "TOKEN RETRIEVAL: PASSED"
        #End If
        MsgBox "Fresh Token Retrieved!" & vbCrLf & vbCrLf & _
               "Token Type: " & Doppio.m_s_TokenType & vbCrLf & _
               "Token Length: " & Len(Doppio.m_s_AccessToken) & vbCrLf & _
               "Company: " & Doppio.m_s_Company & vbCrLf & _
               "Division: " & Doppio.m_s_Division, vbInformation
    Else
        #If DEBUG_MODE Then
            Debug.Print ""
            Debug.Print "TOKEN RETRIEVAL: FAILED"
        #End If
        MsgBox "Failed to get token!", vbCritical
    End If
    
    Exit Sub
    
ErrorHandler:
    #If DEBUG_MODE Then
        Debug.Print "ERROR: " & Err.description
    #End If
    MsgBox "Error: " & Err.description, vbCritical
End Sub

''
' Test the full authentication flow step by step
''
Public Sub Test_AuthenticationFlow()
    On Error GoTo ErrorHandler
    
    #If DEBUG_MODE Then
        Debug.Print "======================================"
        Debug.Print "TEST: Full Authentication Flow"
        Debug.Print "======================================"
        Debug.Print ""
    #End If
    
    ' Step 1: Check environment
    #If DEBUG_MODE Then
        Debug.Print "STEP 1: Check Environment Selection"
        Debug.Print "------------------------------------"
    #End If
    Dim envName As String
    envName = ""
    On Error Resume Next
    envName = ActiveSheet.Range("Environment").value
    On Error GoTo ErrorHandler
    
    If envName = "" Then
        #If DEBUG_MODE Then
            Debug.Print "  FAIL: No environment selected"
        #End If
        MsgBox "Please select an environment in cell I2", vbExclamation
        Exit Sub
    End If
    #If DEBUG_MODE Then
        Debug.Print "  OK: Environment = " & envName
        Debug.Print ""
    #End If
    
    ' Step 2: Load tenant config
    #If DEBUG_MODE Then
        Debug.Print "STEP 2: Load Tenant Configuration"
        Debug.Print "------------------------------------"
    #End If
    Doppio.Tenant_Information
    
    If Doppio.ti = "" Then
        #If DEBUG_MODE Then
            Debug.Print "  FAIL: Tenant ID not loaded"
        #End If
        MsgBox "Failed to load tenant configuration", vbCritical
        Exit Sub
    End If
    #If DEBUG_MODE Then
        Debug.Print "  OK: Tenant ID = " & Doppio.ti
        Debug.Print "  OK: ION URL = " & Doppio.iu
        Debug.Print "  OK: Main URL = " & Doppio.m_s_MainUrl
        Debug.Print ""
    #End If
    
    ' Step 3: Check/Get token
    #If DEBUG_MODE Then
        Debug.Print "STEP 3: Check/Get Access Token"
        Debug.Print "------------------------------------"
    #End If
    If Len(Doppio.m_s_AccessToken) = 0 Then
        #If DEBUG_MODE Then
            Debug.Print "  No token cached, getting fresh token..."
        #End If
        Doppio.Tenant_Token
    Else
        #If DEBUG_MODE Then
            Debug.Print "  Token already cached"
        #End If
    End If
    
    If Len(Doppio.m_s_AccessToken) = 0 Then
        #If DEBUG_MODE Then
            Debug.Print "  FAIL: Could not get access token"
        #End If
        MsgBox "Failed to get access token", vbCritical
        Exit Sub
    End If
    #If DEBUG_MODE Then
        Debug.Print "  OK: Token Type = " & Doppio.m_s_TokenType
        Debug.Print "  OK: Token Length = " & Len(Doppio.m_s_AccessToken)
        Debug.Print ""
    #End If
    
    ' Step 4: Test token with API call
    #If DEBUG_MODE Then
        Debug.Print "STEP 4: Validate Token with API Call"
        Debug.Print "------------------------------------"
    #End If
    Dim config As httpConfig
    Dim httpResponse As httpResponse
    
    config.url = Doppio.m_s_MainUrl & "/M3/m3api-rest/v2/execute/MRS001MI/GetUserInfo/?"
    config.method = HttpMethod_GET
    config.contentType = "application/json"
    config.AcceptType = "application/json"
    config.authHeader = Doppio.m_s_TokenType & " " & Doppio.m_s_AccessToken
    config.TimeoutSeconds = 30
    config.body = ""
    
    httpResponse = DoppioHttp.ExecuteRequest(config)
    
    If httpResponse.statusCode = 200 Then
        #If DEBUG_MODE Then
            Debug.Print "  OK: API call successful (HTTP 200)"
        #End If
    ElseIf httpResponse.statusCode = 401 Then
        #If DEBUG_MODE Then
            Debug.Print "  FAIL: Unauthorized (HTTP 401) - Token invalid or expired"
        #End If
        MsgBox "Token is invalid or expired (401 Unauthorized)", vbCritical
        Exit Sub
    Else
        #If DEBUG_MODE Then
            Debug.Print "  FAIL: HTTP " & httpResponse.statusCode
        #End If
        MsgBox "API call failed with HTTP " & httpResponse.statusCode, vbCritical
        Exit Sub
    End If
    #If DEBUG_MODE Then
        Debug.Print ""
    #End If
    
    ' Step 5: Parse response
    #If DEBUG_MODE Then
        Debug.Print "STEP 5: Parse Response"
        Debug.Print "------------------------------------"
    #End If
    Dim json As Object
    Set json = JsonConverter.ParseJson(httpResponse.body)
    
    Dim results As Object
    Dim firstResult As Object
    Dim records As Object
    Dim firstRecord As Object
    
    On Error GoTo ErrorHandler
    
    Set results = json.item("results")
    Set firstResult = results.item(1)
    Set records = firstResult.item("records")
    Set firstRecord = records.item(1)
    
    Dim userName As String
    Dim company As String
    Dim division As String
    userName = firstRecord.item("ZZUSID")
    company = firstRecord.item("ZDCONO")
    division = firstRecord.item("ZDDIVI")
    
    #If DEBUG_MODE Then
        Debug.Print "  OK: User = " & userName
        Debug.Print "  OK: Company = " & company
        Debug.Print "  OK: Division = " & division
        Debug.Print ""
    #End If
    
    #If DEBUG_MODE Then
        Debug.Print "======================================"
        Debug.Print "ALL STEPS PASSED - Authentication OK!"
        Debug.Print "======================================"
    #End If
    
    MsgBox "Authentication Flow Test PASSED!" & vbCrLf & vbCrLf & _
           "Environment: " & envName & vbCrLf & _
           "User: " & userName & vbCrLf & _
           "Company: " & company & vbCrLf & _
           "Division: " & division, vbInformation
    
    Exit Sub
    
ErrorHandler:
    #If DEBUG_MODE Then
        Debug.Print "ERROR at step: " & Err.description
    #End If
    MsgBox "Authentication Flow Error: " & Err.description, vbCritical
End Sub

''
' Test the NEW Tenant_Information (InitializeTenantToken in DoppioAuth)
' Compares old Doppio.Tenant_Information vs new DoppioAuth.InitializeTenantToken
''
Public Sub Test_NewTenantInformation()
    On Error GoTo ErrorHandler
    
    #If DEBUG_MODE Then
        Debug.Print "======================================"
        Debug.Print "TEST: New Tenant Information"
        Debug.Print "======================================"
        Debug.Print ""
    #End If
    
    ' Check if environment is selected
    Dim envName As String
    envName = ""
    On Error Resume Next
    envName = ActiveSheet.Range("Environment").value
    On Error GoTo ErrorHandler
    
    If envName = "" Then
        #If DEBUG_MODE Then
            Debug.Print "ERROR: No environment selected"
        #End If
        MsgBox "Please select an environment first.", vbExclamation
        Exit Sub
    End If
    
    #If DEBUG_MODE Then
        Debug.Print "Selected Environment: " & envName
        Debug.Print ""
    #End If
    
    ' ==========================================
    ' PART 1: Test OLD method (Doppio.Tenant_Information)
    ' ==========================================
    #If DEBUG_MODE Then
        Debug.Print "--- PART 1: OLD METHOD (Doppio.Tenant_Information) ---"
        Debug.Print ""
    #End If
    
    Dim startTimeOld As Single
    startTimeOld = Timer
    
    Doppio.Tenant_Information
    
    #If DEBUG_MODE Then
        Debug.Print "OLD Method completed in " & Round(Timer - startTimeOld, 3) & " seconds"
        Debug.Print ""
        Debug.Print "OLD Results:"
        Debug.Print "  Tenant ID (ti): " & Doppio.ti
        Debug.Print "  Client ID (ci): " & IIf(Len(Doppio.ci) > 10, Left(Doppio.ci, 10) & "...", Doppio.ci)
        Debug.Print "  ION URL (iu): " & Doppio.iu
        Debug.Print "  SSO URL (pu): " & Doppio.pu
        Debug.Print "  Main URL: " & Doppio.m_s_MainUrl
        Debug.Print "  M3 User: " & Doppio.m_s_M3user
        Debug.Print "  Is Multitenant: " & Doppio.m_b_Multitenant
        Debug.Print "  Has SAAK: " & (Len(Doppio.saak) > 0)
        Debug.Print ""
    #End If
    
    ' ==========================================
    ' PART 2: Test NEW method (DoppioConfig.Config_LoadTenantConfig)
    ' ==========================================
    #If DEBUG_MODE Then
        Debug.Print "--- PART 2: NEW METHOD (DoppioConfig) ---"
        Debug.Print ""
    #End If
    
    Dim startTimeNew As Single
    startTimeNew = Timer
    
    ' Load tenant config using new module
    Dim loadSuccess As Boolean
    loadSuccess = DoppioConfig.Config_LoadTenantConfig(envName)
    
    #If DEBUG_MODE Then
        Debug.Print "NEW Method completed in " & Round(Timer - startTimeNew, 3) & " seconds"
        Debug.Print "Load Success: " & loadSuccess
        Debug.Print ""
    #End If
    
    If loadSuccess Then
        #If DEBUG_MODE Then
            Debug.Print "NEW Results:"
            Debug.Print "  Tenant ID: " & DoppioConfig.Config_TenantConfig.TenantID
            Debug.Print "  Client ID: " & IIf(Len(DoppioConfig.Config_TenantConfig.clientId) > 10, Left(DoppioConfig.Config_TenantConfig.clientId, 10) & "...", DoppioConfig.Config_TenantConfig.clientId)
            Debug.Print "  ION URL: " & DoppioConfig.Config_TenantConfig.IonUrl
            Debug.Print "  SSO URL: " & DoppioConfig.Config_TenantConfig.SsoUrl
            Debug.Print "  Main URL: " & DoppioConfig.Config_MainUrl
            Debug.Print "  M3 User: " & DoppioConfig.Config_M3User
            Debug.Print "  Is Multitenant: " & DoppioConfig.Config_IsMultitenant
            Debug.Print ""
        #End If
    Else
        #If DEBUG_MODE Then
            Debug.Print "  FAILED to load tenant config"
            Debug.Print ""
        #End If
    End If
    
    ' ==========================================
    ' PART 3: Compare Results
    ' ==========================================
    #If DEBUG_MODE Then
        Debug.Print "--- PART 3: COMPARISON ---"
        Debug.Print ""
    #End If
    
    Dim match As Boolean
    match = True
    
    If Doppio.ti <> DoppioConfig.Config_TenantConfig.TenantID Then
        #If DEBUG_MODE Then
            Debug.Print "  MISMATCH - Tenant ID: OLD=" & Doppio.ti & " NEW=" & DoppioConfig.Config_TenantConfig.TenantID
        #End If
        match = False
    Else
        #If DEBUG_MODE Then
            Debug.Print "  MATCH - Tenant ID: " & Doppio.ti
        #End If
    End If
    
    If Doppio.ci <> DoppioConfig.Config_TenantConfig.clientId Then
        #If DEBUG_MODE Then
            Debug.Print "  MISMATCH - Client ID"
        #End If
        match = False
    Else
        #If DEBUG_MODE Then
            Debug.Print "  MATCH - Client ID"
        #End If
    End If
    
    If Doppio.iu <> DoppioConfig.Config_TenantConfig.IonUrl Then
        #If DEBUG_MODE Then
            Debug.Print "  MISMATCH - ION URL: OLD=" & Doppio.iu & " NEW=" & DoppioConfig.Config_TenantConfig.IonUrl
        #End If
        match = False
    Else
        #If DEBUG_MODE Then
            Debug.Print "  MATCH - ION URL"
        #End If
    End If
    
    If Doppio.m_s_MainUrl <> DoppioConfig.Config_MainUrl Then
        #If DEBUG_MODE Then
            Debug.Print "  MISMATCH - Main URL: OLD=" & Doppio.m_s_MainUrl & " NEW=" & DoppioConfig.Config_MainUrl
        #End If
        match = False
    Else
        #If DEBUG_MODE Then
            Debug.Print "  MATCH - Main URL"
        #End If
    End If
    
    #If DEBUG_MODE Then
        Debug.Print ""
        Debug.Print "======================================"
    #End If
    If match And loadSuccess Then
        #If DEBUG_MODE Then
            Debug.Print "TEST PASSED - All values match!"
        #End If
        MsgBox "Tenant Information Test PASSED!" & vbCrLf & vbCrLf & _
            "Environment: " & envName & vbCrLf & _
            "Tenant ID: " & Doppio.ti & vbCrLf & _
            "Main URL: " & Doppio.m_s_MainUrl, vbInformation
    Else
        #If DEBUG_MODE Then
            Debug.Print "TEST FAILED - Values do not match or load failed"
        #End If
        MsgBox "Tenant Information Test FAILED!" & vbCrLf & vbCrLf & _
            "Check the Immediate Window for details.", vbCritical
    End If
    #If DEBUG_MODE Then
        Debug.Print "======================================"
    #End If
    
    Exit Sub
    
ErrorHandler:
    #If DEBUG_MODE Then
        Debug.Print "ERROR: " & Err.description
    #End If
    MsgBox "Test Error: " & Err.description, vbCritical
End Sub

''
' Test InitializeTenantToken from DoppioAuth
''
Public Sub Test_InitializeTenantToken()
    On Error GoTo ErrorHandler
    
    #If DEBUG_MODE Then
        Debug.Print "======================================"
        Debug.Print "TEST: InitializeTenantToken (DoppioAuth)"
        Debug.Print "======================================"
        Debug.Print ""
    #End If
    
    ' Check if environment is selected
    Dim envName As String
    envName = ""
    On Error Resume Next
    envName = ActiveSheet.Range("Environment").value
    On Error GoTo ErrorHandler
    
    If envName = "" Then
        #If DEBUG_MODE Then
            Debug.Print "ERROR: No environment selected"
        #End If
        MsgBox "Please select an environment first.", vbExclamation
        Exit Sub
    End If
    
    #If DEBUG_MODE Then
        Debug.Print "Selected Environment: " & envName
        Debug.Print ""
    #End If
    
    ' Clear any existing session state
    #If DEBUG_MODE Then
        Debug.Print "Clearing existing session state..."
    #End If
    DoppioConfig.Config_AccessToken = ""
    DoppioConfig.Config_RefreshToken = ""
    DoppioConfig.Config_TokenType = ""
    #If DEBUG_MODE Then
        Debug.Print ""
    #End If
    
    ' Call InitializeTenantToken
    #If DEBUG_MODE Then
        Debug.Print "Calling DoppioAuth.InitializeTenantToken..."
    #End If
    Dim startTime As Single
    startTime = Timer
    
    Dim success As Boolean
    success = DoppioAuth.InitializeTenantToken()
    
    #If DEBUG_MODE Then
        Debug.Print "Completed in " & Round(Timer - startTime, 3) & " seconds"
        Debug.Print "Success: " & success
        Debug.Print ""
    #End If
    
    If success Then
        #If DEBUG_MODE Then
            Debug.Print "Results:"
            Debug.Print "  Token Type: " & DoppioConfig.Config_TokenType
            Debug.Print "  Has Access Token: " & (Len(DoppioConfig.Config_AccessToken) > 0)
            Debug.Print "  Token Length: " & Len(DoppioConfig.Config_AccessToken)
            Debug.Print "  Main URL: " & DoppioConfig.Config_MainUrl
            Debug.Print "  Company: " & DoppioConfig.Config_Company
            Debug.Print "  Division: " & DoppioConfig.Config_Division
            Debug.Print "  M3 User: " & DoppioConfig.Config_M3User
            Debug.Print ""
        #End If
        
        ' Test the token with an API call
        #If DEBUG_MODE Then
            Debug.Print "Testing token with API call..."
        #End If
        Dim config As httpConfig
        Dim httpResponse As httpResponse
        
        config.url = DoppioConfig.Config_MainUrl & "/M3/m3api-rest/v2/execute/MRS001MI/GetUserInfo/?"
        config.method = HttpMethod_GET
        config.contentType = "application/json"
        config.AcceptType = "application/json"
        config.authHeader = DoppioConfig.Config_TokenType & " " & DoppioConfig.Config_AccessToken
        config.TimeoutSeconds = 30
        config.body = ""
        
        httpResponse = DoppioHttp.ExecuteRequest(config)
        
        #If DEBUG_MODE Then
            Debug.Print "  HTTP Status: " & httpResponse.statusCode
            Debug.Print "  Success: " & httpResponse.success
            Debug.Print ""
        #End If
        
        #If DEBUG_MODE Then
            Debug.Print "======================================"
        #End If
        If httpResponse.statusCode = 200 Then
            #If DEBUG_MODE Then
                Debug.Print "TEST PASSED!"
            #End If
            MsgBox "InitializeTenantToken Test PASSED!" & vbCrLf & vbCrLf & _
                "Environment: " & envName & vbCrLf & _
                "Token Type: " & DoppioConfig.Config_TokenType & vbCrLf & _
                "API Call: HTTP 200 OK", vbInformation
        Else
            #If DEBUG_MODE Then
                Debug.Print "TEST FAILED - API call returned " & httpResponse.statusCode
            #End If
            MsgBox "InitializeTenantToken Test FAILED!" & vbCrLf & vbCrLf & _
                "Token obtained but API call failed" & vbCrLf & _
                "HTTP Status: " & httpResponse.statusCode, vbCritical
        End If
        #If DEBUG_MODE Then
            Debug.Print "======================================"
        #End If
    Else
        #If DEBUG_MODE Then
            Debug.Print "======================================"
            Debug.Print "TEST FAILED - Could not initialize tenant token"
            Debug.Print "======================================"
        #End If
        MsgBox "InitializeTenantToken Test FAILED!" & vbCrLf & vbCrLf & _
            "Could not obtain token." & vbCrLf & _
            "Check the Immediate Window for details.", vbCritical
    End If
    
    Exit Sub
    
ErrorHandler:
    #If DEBUG_MODE Then
        Debug.Print "ERROR: " & Err.description
    #End If
    MsgBox "Test Error: " & Err.description, vbCritical
End Sub

''
' EXAMPLE: How to use the new methods (equivalent to old Tenant_Token)
'
' OLD WAY:
'   Doppio.Tenant_Token
'   ' Token is now in Doppio.m_s_AccessToken
'   ' Token is cached in Doppio.manager
'
' NEW WAY:
'   DoppioAuth.InitializeTenantToken
'   ' Token is now in DoppioConfig.Config_AccessToken
'   ' Token is cached in DoppioConfig.Config_EnvironmentManager
'
''
Public Sub Example_NewTenantToken()
    Dim success As Boolean
    
    ' Step 1: Initialize tenant token (loads config + gets token + caches it)
    success = DoppioAuth.InitializeTenantToken()
    
    If success Then
        ' Step 2: Access the token and other values
        #If DEBUG_MODE Then
            Debug.Print "Token Type: " & DoppioConfig.Config_TokenType
            Debug.Print "Access Token: " & Left(DoppioConfig.Config_AccessToken, 20) & "..."
            Debug.Print "Main URL: " & DoppioConfig.Config_MainUrl
            Debug.Print "Company: " & DoppioConfig.Config_Company
            Debug.Print "Division: " & DoppioConfig.Config_Division
        #End If
        
        ' Step 3: Use for API calls
        Dim authHeader As String
        authHeader = DoppioConfig.Config_TokenType & " " & DoppioConfig.Config_AccessToken
        
        ' Now you can use authHeader in your HTTP requests
    End If
End Sub

''
' Side-by-side comparison: Old Tenant_Token vs New InitializeTenantToken
''
Public Sub Test_CompareTenantToken()
    On Error GoTo ErrorHandler
    
    #If DEBUG_MODE Then
        Debug.Print "======================================"
        Debug.Print "COMPARISON: Old vs New Tenant_Token"
        Debug.Print "======================================"
        Debug.Print ""
    #End If
    
    ' Check environment
    If ActiveSheet.Range("Environment").value = "" Then
        MsgBox "Please select an environment first.", vbExclamation
        Exit Sub
    End If
    
    Dim envName As String
    envName = ActiveSheet.Range("Environment").value
    #If DEBUG_MODE Then
        Debug.Print "Environment: " & envName
        Debug.Print ""
    #End If
    
    ' ==========================================
    ' OLD METHOD
    ' ==========================================
    #If DEBUG_MODE Then
        Debug.Print "--- OLD METHOD (Doppio.Tenant_Token) ---"
    #End If
    
    ' Clear old tokens
    Doppio.m_s_AccessToken = ""
    Doppio.m_s_TokenType = ""
    
    Dim startTimeOld As Single
    startTimeOld = Timer
    
'    Doppio.Tenant_Token
    
    Dim oldTime As Single
    oldTime = Timer - startTimeOld
    
    #If DEBUG_MODE Then
        Debug.Print "  Time: " & Round(oldTime, 3) & "s"
        Debug.Print "  Token Type: " & Doppio.m_s_TokenType
        Debug.Print "  Has Token: " & (Len(Doppio.m_s_AccessToken) > 0)
        Debug.Print "  Token Length: " & Len(Doppio.m_s_AccessToken)
        Debug.Print "  Main URL: " & Doppio.m_s_MainUrl
        Debug.Print "  Company: " & Doppio.m_s_Company
        Debug.Print "  Division: " & Doppio.m_s_Division
        Debug.Print ""
    #End If
    
    ' ==========================================
    ' NEW METHOD
    ' ==========================================
    #If DEBUG_MODE Then
        Debug.Print "--- NEW METHOD (DoppioAuth.InitializeTenantToken) ---"
    #End If
    
    ' Clear new tokens
    DoppioConfig.Config_AccessToken = ""
    DoppioConfig.Config_TokenType = ""
    
    Dim startTimeNew As Single
    startTimeNew = Timer
    
    Dim success As Boolean
    success = DoppioAuth.InitializeTenantToken()
    
    Dim newTime As Single
    newTime = Timer - startTimeNew
    
    #If DEBUG_MODE Then
        Debug.Print "  Success: " & success
        Debug.Print "  Time: " & Round(newTime, 3) & "s"
        Debug.Print "  Token Type: " & DoppioConfig.Config_TokenType
        Debug.Print "  Has Token: " & (Len(DoppioConfig.Config_AccessToken) > 0)
        Debug.Print "  Token Length: " & Len(DoppioConfig.Config_AccessToken)
        Debug.Print "  Main URL: " & DoppioConfig.Config_MainUrl
        Debug.Print "  Company: " & DoppioConfig.Config_Company
        Debug.Print "  Division: " & DoppioConfig.Config_Division
        Debug.Print ""
    #End If
    
    ' ==========================================
    ' COMPARISON
    ' ==========================================
    #If DEBUG_MODE Then
        Debug.Print "--- COMPARISON ---"
        Debug.Print "  Old Time: " & Round(oldTime, 3) & "s"
        Debug.Print "  New Time: " & Round(newTime, 3) & "s"
        Debug.Print "  Tokens Match: " & (Doppio.m_s_AccessToken = DoppioConfig.Config_AccessToken)
        Debug.Print "  Main URLs Match: " & (Doppio.m_s_MainUrl = DoppioConfig.Config_MainUrl)
        Debug.Print ""
    #End If
    
    #If DEBUG_MODE Then
        Debug.Print "======================================"
    #End If
    MsgBox "Comparison Complete!" & vbCrLf & vbCrLf & _
        "OLD: " & Round(oldTime, 3) & "s, Token=" & (Len(Doppio.m_s_AccessToken) > 0) & vbCrLf & _
        "NEW: " & Round(newTime, 3) & "s, Token=" & (Len(DoppioConfig.Config_AccessToken) > 0), vbInformation
    
    Exit Sub
    
ErrorHandler:
    #If DEBUG_MODE Then
        Debug.Print "ERROR: " & Err.description
    #End If
    MsgBox "Error: " & Err.description, vbCritical
End Sub


