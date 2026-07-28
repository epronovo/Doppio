Attribute VB_Name = "Doppio_Api"
''
' Doppio API Module
' High-level API operations for M3 (MI, IDM, IPS, XtendM3, etc.)
'
' @module Doppio_Api
' @author Doppio Group - eric@doppiogroup.com
' @version 2.0
''
Option Explicit

' =============================================================================
' PUBLIC API
' =============================================================================

''
' Execute an API call with full error handling and retry logic
' @param request - API request configuration
' @return ApiResponse - Response from the API
''
Public Function ExecuteApiCall(request As ApiRequest) As apiResponse
    Dim response As apiResponse
    Dim httpConfig As httpConfig
    Dim httpResponse As httpResponse
    Dim retryCount As Integer
    Dim maxRetries As Integer
    
    On Error GoTo ErrorHandler
    
    maxRetries = 3
    retryCount = 0
    
    ' Ensure we have authentication
    If Not EnsureAuthenticated() Then
        response.success = False
        response.status = ApiStatus_Unauthorized
        response.errorMessage = "Not authenticated"
        ExecuteApiCall = response
        Exit Function
    End If
    
    ' Check cache first
    If request.UseCache And request.cacheKey <> "" Then
        If Cache_TryGetFromCache(request.cacheKey, response) Then
            ExecuteApiCall = response
            Exit Function
        End If
    End If
    
    ' Build HTTP configuration
    httpConfig = BuildHttpConfigForApi(request)
    
    ' Execute with retry logic
    Do While retryCount < maxRetries
        httpResponse = ExecuteRequest(httpConfig)
        
        If httpResponse.success Then
            Exit Do
        ElseIf httpResponse.IsUnauthorized Then
            ' Token expired - try to refresh
            If HandleUnauthorized() Then
                ' Update auth header with new token
                httpConfig.authHeader = Config_GetAuthorizationHeader()
                retryCount = retryCount + 1
            Else
                ' Refresh failed
                response.success = False
                response.status = ApiStatus_Unauthorized
                response.errorMessage = "Authentication failed"
                ExecuteApiCall = response
                Exit Function
            End If
        Else
            ' Other error - don't retry
            Exit Do
        End If
    Loop
    
    ' Parse the response based on API type
    response = ParseApiResponse(httpResponse, request.apiType)
    
    ' Cache successful responses
    If response.success And request.UseCache And request.cacheKey <> "" Then
        Cache_StoreInCache request.cacheKey, response
    End If
    
    ExecuteApiCall = response
    Exit Function
    
ErrorHandler:
    response.success = False
    response.status = ApiStatus_Error
    response.errorMessage = Err.description
    Core_LogError CreateError("ExecuteApiCall", request.endpoint)
    ExecuteApiCall = response
End Function

''
' Execute a simple MI API GET request
' @param program - MI program name (e.g., "CRS610MI")
' @param transaction - Transaction name (e.g., "LstByNumber")
' @param Optional parameters - URL parameters (e.g., "CUNO=TEST")
' @return ApiResponse - Response from the API
''
Public Function ExecuteMiGet(program As String, transaction As String, _
                              Optional parameters As String = "") As apiResponse
    Dim request As ApiRequest
    Dim endpoint As String
    
    endpoint = program & "/" & transaction
    If parameters <> "" Then
        endpoint = endpoint & "?" & parameters
    End If
    
    With request
        .mainUrl = Config_MainUrl
        .apiPath = MI_API_PATH
        .endpoint = endpoint
        .body = ""
        .apiType = ApiType_MI
        .UseCache = True
        .cacheKey = endpoint
    End With
    
    ExecuteMiGet = ExecuteApiCall(request)
End Function

''
' Execute an MI API bulk operation (POST)
' @param program - MI program name
' @param transactions - JSON body with transactions
' @return ApiResponse - Response from the API
''
Public Function ExecuteMiBulk(program As String, transactions As String) As apiResponse
    Dim request As ApiRequest

    With request
        .mainUrl = Config_MainUrl
        .apiPath = MI_API_PATH
        .endpoint = BuildMiQueryString()
        .body = "{""program"":""" & program & """,""transactions"":" & transactions & "}"
        .apiType = ApiType_MI
        .UseCache = False
        .cacheKey = ""
    End With

    ExecuteMiBulk = ExecuteApiCall(request)
End Function

''
' Execute an IDM API request
' @param endpoint - IDM endpoint
' @param method - HTTP method
' @param Optional body - Request body
' @return ApiResponse - Response from the API
''
Public Function ExecuteIdmRequest(endpoint As String, method As httpMethod, _
                                   Optional body As String = "") As apiResponse
    Dim request As ApiRequest
    
    With request
        .mainUrl = Config_MainUrl
        .apiPath = IDM_API_PATH
        .endpoint = endpoint
        .body = body
        .apiType = ApiType_IDM
        .UseCache = False
        .cacheKey = ""
    End With
    
    ExecuteIdmRequest = ExecuteApiCall(request)
End Function

''
' Execute an XtendM3 API request
' @param endpoint - XtendM3 endpoint
' @param body - Request body
' @return ApiResponse - Response from the API
''
Public Function ExecuteXtendM3Request(endpoint As String, body As String) As apiResponse
    Dim request As ApiRequest
    
    With request
        .mainUrl = Config_MainUrl
        .apiPath = "/M3/xtendm3/api"
        .endpoint = endpoint
        .body = body
        .apiType = ApiType_XtendM3
        .UseCache = False
        .cacheKey = ""
    End With
    
    ExecuteXtendM3Request = ExecuteApiCall(request)
End Function

''
' Upload a file to M3 File Management
' @param fileName - Name of the file
' @param filePath - Local path to the file
' @return ApiResponse - Response from the API
''
Public Function UploadFile(fileName As String, filePath As String) As apiResponse
    Dim request As ApiRequest
    
    With request
        .mainUrl = Config_MainUrl
        .apiPath = FILE_MGT_PATH
        .endpoint = "file/FileImport/" & fileName
        .body = "@" & filePath  ' Special marker for file upload
        .apiType = ApiType_FileMng
        .UseCache = False
        .cacheKey = ""
    End With
    
    UploadFile = ExecuteApiCall(request)
End Function

''
' Get metadata for an MI program
' @param program - MI program name
' @return ApiResponse - Response containing metadata
''
Public Function GetMiMetadata(program As String) As apiResponse
    Dim request As ApiRequest
    
    With request
        .mainUrl = Config_MainUrl
        .apiPath = MI_API_PATH
        .endpoint = program & ".meta"
        .body = ""
        .apiType = ApiType_MI
        .UseCache = True
        .cacheKey = program & ".meta"
    End With
    
    GetMiMetadata = ExecuteApiCall(request)
End Function

''
' List transactions for an MI program
' @param program - MI program name
' @return ApiResponse - Response containing transaction list
''
Public Function ListMiTransactions(program As String) As apiResponse
    Dim endpoint As String
    
    endpoint = "MRS001MI/LstTransactions;maxrecs=0?MINM=" & program & _
               "&returncols=MINM,TRNM,SIMU"
    
    ExecuteMiGet "MRS001MI", "LstTransactions", "MINM=" & program & "&returncols=MINM,TRNM,SIMU"
End Function

''
' List field information for a database table
' @param tableName - Table name (first 6 characters)
' @return ApiResponse - Response containing field info
''
Public Function ListTableFields(tableName As String) As apiResponse
    ListTableFields = ExecuteMiGet("MRS001MI", "LstFieldInfo", "FILE=" & Left(tableName, 6))
End Function

' =============================================================================
' PRIVATE HELPERS
' =============================================================================

''
' Build HTTP configuration for an API request
' @param request - API request
' @return HttpConfig - HTTP configuration
''
Private Function BuildHttpConfigForApi(request As ApiRequest) As httpConfig
    Dim config As httpConfig
    Dim url As String
    Dim method As httpMethod
    Dim contentType As String
    
    ' Build URL
    url = request.mainUrl & request.apiPath
    If request.endpoint <> "" Then
        If Right(url, 1) <> "/" And Left(request.endpoint, 1) <> "/" And Left(request.endpoint, 1) <> "?" Then
            url = url & "/"
        End If
        url = url & request.endpoint
    End If
    
    ' Remove double question marks if present
    url = Replace(url, "?&", "?")
    
    ' Determine method and content type based on API type
    Select Case request.apiType
        Case ApiType_MI
            If request.body = "" Then
                method = HttpMethod_GET
            Else
                method = HttpMethod_POST
            End If
            contentType = "application/json; charset=UTF-8"
            
        Case ApiType_IDM
            method = HttpMethod_GET  ' IDM method determined by caller
            contentType = "application/json; charset=UTF-8"
            
        Case ApiType_IPS
            If request.body = "" Then
                method = HttpMethod_GET
            Else
                method = HttpMethod_POST
            End If
            contentType = "application/xml"
            
        Case ApiType_FileMng
            method = HttpMethod_PUT
            contentType = "application/octet-stream"
            
        Case ApiType_XtendM3
            method = HttpMethod_PUT
            contentType = "application/json; charset=UTF-8"
            
        Case Else
            method = HttpMethod_GET
            contentType = "application/json; charset=UTF-8"
    End Select
    
    ' Build configuration
    With config
        .url = url
        .method = method
        .contentType = contentType
        .AcceptType = IIf(request.apiType = ApiType_IPS, "application/xml", "application/json")
        .authHeader = Config_GetAuthorizationHeader()
        .body = request.body
        .timeoutSeconds = Config_MaxTimeout
        .WriteToFile = True
    End With
    
    BuildHttpConfigForApi = config
End Function

''
' Parse an HTTP response into an API response
' @param httpResponse - Raw HTTP response
' @param apiType - Type of API
' @return ApiResponse - Parsed response
''
Private Function ParseApiResponse(httpResponse As httpResponse, apiType As apiType) As apiResponse
    Dim response As apiResponse
    Dim json As Object
    
    On Error GoTo ErrorHandler
    
    ' ALWAYS preserve the raw data
    response.data = httpResponse.body
    
    If Not httpResponse.success Then
        response.success = False
        response.status = IIf(httpResponse.IsUnauthorized, ApiStatus_Unauthorized, ApiStatus_Error)
        response.errorMessage = httpResponse.errorMessage
        ParseApiResponse = response
        Exit Function
    End If
    
    ' Parse based on API type
    Select Case apiType
        Case ApiType_MI
            response = ParseMiResponse(httpResponse.body)
            response.data = httpResponse.body  ' Restore raw data after parsing
            
        Case ApiType_IDM
            response = ParseIdmResponse(httpResponse.body)
            response.data = httpResponse.body  ' Restore raw data after parsing
            
        Case ApiType_IPS
            response = ParseIpsResponse(httpResponse.body)
            response.data = httpResponse.body  ' Restore raw data after parsing
            
        Case ApiType_XtendM3
            response = ParseXtendM3Response(httpResponse.body)
            response.data = httpResponse.body  ' Restore raw data after parsing
            
        Case ApiType_FileMng
            ' File management returns status code
            response.success = (httpResponse.statusCode = 201)
            response.status = IIf(response.success, ApiStatus_Success, ApiStatus_Error)
            ' Already has response.data from above
            
        Case Else
            ' Generic JSON parsing
            Set json = ParseJson(httpResponse.body)
            Set response.results = json
            response.success = True
            response.status = ApiStatus_Success
            ' Already has response.data from above
    End Select
    
    ParseApiResponse = response
    Exit Function
    
ErrorHandler:
    response.success = False
    response.status = ApiStatus_ParseError
    response.errorMessage = "Failed to parse response: " & Err.description
    ParseApiResponse = response
End Function

''
' Parse MI API response
' @param responseBody - Raw response body
' @return ApiResponse - Parsed response
''
Private Function ParseMiResponse(responseBody As String) As apiResponse
    Dim response As apiResponse
    Dim json As Object
    Dim terminationReason As String
    Dim connectionError As String
    
    On Error GoTo ErrorHandler
    
    Set json = ParseJson(responseBody)
    
    ' Check for termination reason (API error)
    terminationReason = SafeStr(json.item("terminationReason"))
    connectionError = SafeStr(json.item("error"))
    
    If terminationReason <> "" Or connectionError <> "" Then
        response.success = False
        response.status = ApiStatus_Error
        response.errorMessage = IIf(terminationReason <> "", terminationReason, connectionError)
        ParseMiResponse = response
        Exit Function
    End If
    
    ' Extract results and records
    Set response.results = json.item("results")
    
    If Not response.results Is Nothing Then
        If response.results.count > 0 Then
            Set response.records = response.results(1).item("records")
            response.recordCount = response.records.count
        End If
    End If
    
    response.success = True
    response.status = ApiStatus_Success
    ParseMiResponse = response
    Exit Function
    
ErrorHandler:
    response.success = False
    response.status = ApiStatus_ParseError
    response.errorMessage = "Failed to parse MI response"
    ParseMiResponse = response
End Function

''
' Parse IDM API response
' @param responseBody - Raw response body
' @return ApiResponse - Parsed response
''
Private Function ParseIdmResponse(responseBody As String) As apiResponse
    Dim response As apiResponse
    Dim json As Object
    
    On Error GoTo ErrorHandler
    
    Set json = ParseJson(responseBody)
    
    ' Check for error
    If json.exists("error") Then
        response.success = False
        response.status = ApiStatus_Error
        If IsObject(json.item("error")) Then
            response.errorMessage = SafeStr(json.item("error").item("message"))
        Else
            response.errorMessage = SafeStr(json.item("error"))
        End If
        ParseIdmResponse = response
        Exit Function
    End If
    
    ' Check for termination reason
    If json.exists("terminationReason") Then
        response.success = False
        response.status = ApiStatus_Error
        response.errorMessage = SafeStr(json.item("terminationReason"))
        ParseIdmResponse = response
        Exit Function
    End If
    
    Set response.results = json
    response.success = True
    response.status = ApiStatus_Success
    ParseIdmResponse = response
    Exit Function
    
ErrorHandler:
    response.success = False
    response.status = ApiStatus_ParseError
    response.errorMessage = "Failed to parse IDM response"
    ParseIdmResponse = response
End Function

''
' Parse IPS/SOAP API response
' @param responseBody - Raw response body
' @return ApiResponse - Parsed response
''
Private Function ParseIpsResponse(responseBody As String) As apiResponse
    Dim response As apiResponse
    
    ' IPS responses are XML - store raw and let caller parse
    response.data = responseBody
    response.success = True
    response.status = ApiStatus_Success
    
    ParseIpsResponse = response
End Function

''
' Parse XtendM3 API response
' @param responseBody - Raw response body
' @return ApiResponse - Parsed response
''
Private Function ParseXtendM3Response(responseBody As String) As apiResponse
    Dim response As apiResponse
    Dim json As Object
    
    On Error GoTo ErrorHandler
    
    ' XtendM3 may return just a status code
    If IsNumeric(responseBody) Then
        response.success = (CLng(responseBody) >= 200 And CLng(responseBody) < 300)
        response.status = IIf(response.success, ApiStatus_Success, ApiStatus_Error)
        ParseXtendM3Response = response
        Exit Function
    End If
    
    Set json = ParseJson(responseBody)
    Set response.results = json
    response.success = True
    response.status = ApiStatus_Success
    ParseXtendM3Response = response
    Exit Function
    
ErrorHandler:
    response.success = False
    response.status = ApiStatus_ParseError
    response.errorMessage = "Failed to parse XtendM3 response"
    ParseXtendM3Response = response
End Function

' =============================================================================
' BULK OPERATION HELPERS
' =============================================================================

''
' Build a bulk MI transaction JSON body
' @param program - MI program name
' @param method - Transaction method
' @param fields - Array of field names
' @param values - Array of field values
' @return String - JSON transaction body
''
Public Function BuildBulkTransaction(program As String, method As String, _
                                      fields() As String, values() As String) As String
    Dim json As String
    Dim recordJson As String
    Dim selectedJson As String
    Dim i As Long
    Dim fieldCount As Long
    
    fieldCount = 0
    recordJson = ""
    selectedJson = ""
    
    ' Build record and selected columns
    For i = LBound(fields) To UBound(fields)
        If fields(i) <> "" Then
            ' Add to selected columns
            If selectedJson <> "" Then selectedJson = selectedJson & ","
            selectedJson = selectedJson & """" & fields(i) & """"
            
            ' Add to record if value is not empty
            If values(i) <> "" Then
                If recordJson <> "" Then recordJson = recordJson & ","
                
                ' Escape special characters in value
                Dim escapedValue As String
                escapedValue = Replace(values(i), "\", "\\")
                escapedValue = Replace(escapedValue, """", "\""")
                
                ' Handle PAR1 numeric validation
                If fields(i) = "PAR1" Then
                    escapedValue = ReplaceAlphaWithZero(escapedValue)
                End If
                
                recordJson = recordJson & """" & fields(i) & """:""" & escapedValue & """"
            End If
        End If
    Next i
    
    ' Build the transaction JSON
    json = "{""transaction"":""" & method & ""","
    json = json & """record"":{" & recordJson & "},"
    json = json & """selectedColumns"":[" & selectedJson & "]}"
    
    BuildBulkTransaction = json
End Function

''
' Build multiple bulk transactions for batch processing
' @param program - MI program name
' @param method - Transaction method
' @param fieldNames - Array of field names
' @param dataRows - 2D array of values (rows x columns)
' @param startRow - Starting row in dataRows
' @param rowCount - Number of rows to process
' @return String - JSON array of transactions
''
Public Function ZZZ_BuildBulkTransactions(program As String, method As String, _
                                       fieldNames() As String, dataRows As Variant, _
                                       startRow As Long, rowCount As Long) As String
    Dim transactions As String
    Dim i As Long
    Dim j As Long
    Dim values() As String
    Dim fieldCount As Long
    
    On Error GoTo ErrorHandler
    
    fieldCount = UBound(fieldNames) - LBound(fieldNames) + 1
    ReDim values(LBound(fieldNames) To UBound(fieldNames))
    
    transactions = "["
    
    For i = startRow To startRow + rowCount - 1
        ' Extract values for this row
        For j = LBound(fieldNames) To UBound(fieldNames)
            values(j) = SafeStr(dataRows(i, j - LBound(fieldNames) + 1))
        Next j
        
        ' Add transaction
        If i > startRow Then transactions = transactions & ","
        transactions = transactions & BuildBulkTransaction(program, method, fieldNames, values)
    Next i
    
    transactions = transactions & "]"
    
    ZZZ_BuildBulkTransactions = transactions
    Exit Function

ErrorHandler:
    ZZZ_BuildBulkTransactions = "[]"
End Function

''
' Build the standard MI API query string including maxrecs, extendedresult,
' m3user, righttrim, cono, and divi from current config/session state.
' All parameters are optional overrides; omit them to use config values.
' @param maxRecs      - Override for max records (0 = use config)
' @param m3user       - Override for M3 user (empty = use config)
' @param company      - Override for company (empty = use config)
' @param division     - Override for division (empty = use config)
''
Public Function BuildMiQueryString(Optional maxRecs As Long = 0, _
                                   Optional m3user As String = "", _
                                   Optional company As String = "", _
                                   Optional division As String = "") As String
    Dim settings As ApiSettings
    Dim qs As String
    Dim effectiveMaxRecs As Long
    Dim effectiveUser As String
    Dim effectiveCono As String
    Dim effectiveDivi As String

    settings = Config_ApiSettings

    effectiveMaxRecs = IIf(maxRecs > 0, maxRecs, settings.MaxRecords)
    effectiveUser = IIf(m3user <> "", m3user, Config_M3User)
    effectiveCono = IIf(company <> "", company, Config_Company)
    effectiveDivi = IIf(division <> "", division, Config_Division)

    qs = "?maxrecs=" & effectiveMaxRecs & "&extendedresult=true"

    If effectiveUser <> "" Then qs = qs & "&m3user=" & effectiveUser

    If settings.righttrim Then
        qs = qs & "&righttrim=true"
    Else
        qs = qs & "&righttrim=false"
    End If

    If effectiveCono <> "" Then qs = qs & "&cono=" & effectiveCono
    If effectiveDivi <> "" Then qs = qs & "&divi=" & effectiveDivi

    BuildMiQueryString = qs
End Function


