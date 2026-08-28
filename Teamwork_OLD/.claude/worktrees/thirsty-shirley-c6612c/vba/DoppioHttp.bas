Attribute VB_Name = "DoppioHttp"
''
' Doppio HTTP Module
' Platform-independent HTTP request handling
' Abstracts differences between Mac (curl) and Windows (WinHttp)
'
' @module DoppioHttp
' @author Doppio Group - eric@doppiogroup.com
' @version 2.0
''
Option Explicit

' =============================================================================
' PUBLIC API
' =============================================================================

''
' Execute an HTTP request (platform-independent)
' @param config - HTTP request configuration
' @return HttpResponse - Response from the request
''
Public Function ExecuteRequest(config As httpConfig) As httpResponse
    Dim response As httpResponse
    
    #If Mac Then
        response = ExecuteMacRequest(config)
    #Else
        response = ExecuteWindowsRequest(config)
    #End If
    
    ' --- Retry with re-auth on failure ---
    If ShouldRetryWithReauth(response, config) Then
        If PromptRetry() Then
            ' Clear token everywhere
            ClearCachedToken
            
            ' Re-authenticate
            Doppio.Tenant_Token
            
            ' If we got a new token, update the config and retry
            If Doppio.m_s_AccessToken <> "" Then
                config.authHeader = Doppio.m_s_TokenType & " " & Doppio.m_s_AccessToken
                
                #If Mac Then
                    response = ExecuteMacRequest(config)
                #Else
                    response = ExecuteWindowsRequest(config)
                #End If
            End If
        End If
    End If
    
    ExecuteRequest = response
End Function
''' Determine if a failed response warrants a retry with fresh auth
'
Private Function ShouldRetryWithReauth(response As httpResponse, config As httpConfig) As Boolean
    ' Retry on: 401 Unauthorized, timeout (status 0), or explicit unauthorized flag
    If response.success Then
        ShouldRetryWithReauth = False
    ElseIf response.IsUnauthorized Then
        ShouldRetryWithReauth = True
    ElseIf response.statusCode = 0 Then
        ShouldRetryWithReauth = True  ' Timeout or network error
    Else
        ShouldRetryWithReauth = False
    End If
End Function
''' Clear the cached token from memory and from the Environments sheet
'
Private Sub ClearCachedToken()
    ' Clear in-memory
    Doppio.m_s_AccessToken = ""
    Doppio.activeEnvironment = ""
    
    ' Clear on Environments sheet
    Dim wsEnv As Worksheet
    Dim rngFound As Range
    
    On Error Resume Next
    Set wsEnv = ThisWorkbook.Sheets("Environments")
    On Error GoTo 0
    
    If Not wsEnv Is Nothing Then
        Set rngFound = wsEnv.columns("A").Find(What:=Doppio.m_s_SelectedEnvironment, _
                                               LookIn:=xlValues, _
                                               LookAt:=xlWhole)
        If Not rngFound Is Nothing Then
            wsEnv.Cells(rngFound.row, "E").ClearContents
            #If DEBUG_MODE Then
                Debug.Print "ClearCachedToken: Cleared sheet token for " & Doppio.m_s_SelectedEnvironment
            #End If
        End If
    End If
End Sub



''
' Build an HttpConfig for a standard JSON API call
' @param url - Full URL
' @param method - HTTP method
' @param authHeader - Authorization header value
' @param Optional body - Request body
' @param Optional timeout - Timeout in seconds
' @return HttpConfig - Configured request
''
Public Function BuildJsonConfig(url As String, _
                                 method As httpMethod, _
                                 authHeader As String, _
                                 Optional body As String = "", _
                                 Optional timeout As Integer = 30) As httpConfig
    Dim config As httpConfig
    
    With config
        .url = url
        .method = method
        .contentType = "application/json; charset=UTF-8"
        .AcceptType = "application/json; charset=UTF-8"
        .authHeader = authHeader
        .body = body
        .TimeoutSeconds = timeout
        .WriteToFile = True
    End With
    
    BuildJsonConfig = config
End Function

''
' Build an HttpConfig for XML/SOAP API call
' @param url - Full URL
' @param method - HTTP method
' @param authHeader - Authorization header value
' @param Optional body - Request body
' @param Optional timeout - Timeout in seconds
' @return HttpConfig - Configured request
''
Public Function BuildXmlConfig(url As String, _
                                method As httpMethod, _
                                authHeader As String, _
                                Optional body As String = "", _
                                Optional timeout As Integer = 30) As httpConfig
    Dim config As httpConfig
    
    With config
        .url = url
        .method = method
        .contentType = "application/xml"
        .AcceptType = "application/xml"
        .authHeader = authHeader
        .body = body
        .TimeoutSeconds = timeout
        .WriteToFile = True
    End With
    
    BuildXmlConfig = config
End Function

''
' Build an HttpConfig for file upload
' @param url - Full URL
' @param authHeader - Authorization header value
' @param filePath - Path to file to upload
' @param Optional timeout - Timeout in seconds
' @return HttpConfig - Configured request
''
Public Function BuildFileUploadConfig(url As String, _
                                       authHeader As String, _
                                       filePath As String, _
                                       Optional timeout As Integer = 60) As httpConfig
    Dim config As httpConfig
    
    With config
        .url = url
        .method = HttpMethod_PUT
        .contentType = "application/octet-stream"
        .AcceptType = "application/json"
        .authHeader = authHeader
        .body = "@" & filePath  ' Special marker for file upload
        .TimeoutSeconds = timeout
        .WriteToFile = True
    End With
    
    BuildFileUploadConfig = config
End Function

' =============================================================================
' MAC IMPLEMENTATION (CURL)
' =============================================================================

#If Mac Then

''
' Execute HTTP request on Mac using curl
' @param config - Request configuration
' @return HttpResponse - Response
''
Private Function ExecuteMacRequest(config As httpConfig) As httpResponse
    Dim response As httpResponse
    Dim curlCommand As String
    Dim script As String
    Dim tempFilePath As String
    Dim resultText As String
    
    On Error GoTo ErrorHandler
    
    curlCommand = BuildCurlCommand(config)
    tempFilePath = GetTempFilePath(OUTPUT_FILE_NAME)
    script = "Do shell script """ & curlCommand & " > " & tempFilePath & """"
    
    DeleteFileIfExists tempFilePath
    
    ' Single attempt - retries are handled at ExecuteRequest level
    On Error Resume Next
    MacScript (script)
    
    If Err.Number <> 0 Then
        response.success = False
        response.errorMessage = Err.description
        response.statusCode = 0
        ExecuteMacRequest = response
        Exit Function
    End If
    On Error GoTo ErrorHandler
    
    resultText = Core_ReadFileToString(tempFilePath)
    
    response.body = resultText
    response.success = (Len(resultText) > 0)
    response.statusCode = 200
    
    If Left(resultText, 9) = "{""error"":" Then
        response.IsUnauthorized = True
        response.success = False
    End If
    
    ExecuteMacRequest = response
    Exit Function
    
ErrorHandler:
    response.success = False
    response.errorMessage = Err.description
    response.statusCode = 0
    ExecuteMacRequest = response
End Function

''
' Build a curl command string
' @param config - Request configuration
' @return String - Curl command
''
Private Function BuildCurlCommand(config As httpConfig) As String
    Dim cmd As String
    Dim methodStr As String
    Dim escapedBody As String
    
    methodStr = HttpMethodToString(config.method)
    
    ' Start building command
    cmd = "curl --request " & methodStr
    cmd = cmd & " --max-time " & config.TimeoutSeconds
    cmd = cmd & " --location '" & config.url & "'"
    
    ' Add headers
    cmd = cmd & " --header 'Accept: " & config.AcceptType & "'"
    cmd = cmd & " --header 'Content-Type: " & config.contentType & "'"
    
    If config.authHeader <> "" Then
        cmd = cmd & " --header 'Authorization: " & config.authHeader & "'"
    End If
    
    ' Add body if present
    If config.body <> "" Then
        If Left(config.body, 1) = "@" Then
            ' File upload
            cmd = cmd & " --data-binary '" & config.body & "'"
        Else
            ' Regular body - escape special characters
            escapedBody = EscapeForCurl(config.body)
            cmd = cmd & " --data-raw '" & escapedBody & "'"
        End If
    End If
    
    BuildCurlCommand = cmd
End Function

''
' Escape a string for use in curl command
' @param text - Text to escape
' @return String - Escaped text
''
Private Function EscapeForCurl(text As String) As String
    Dim result As String
    
    result = text
    result = Replace(result, "\", "\\\\")
    result = Replace(result, """", "\""")
    result = Replace(result, "'", "'\\''")
    result = Replace(result, "!!", "\\\""")
    
    EscapeForCurl = result
End Function

''
' Prompt user about retry
' @param oldTimeout - Previous timeout
' @param newTimeout - New timeout to try
' @return Boolean - True if user wants to retry
''
Private Function PromptRetry() As Boolean
    Dim result As VbMsgBoxResult
    
    result = MsgBox("The request failed (unauthorized or timed out)." & vbCrLf & _
                    "Would you like to re-authenticate and retry?", _
                    vbYesNo + vbQuestion, "Retry")
    
    PromptRetry = (result = vbYes)
End Function

#Else

' =============================================================================
' WINDOWS IMPLEMENTATION (WINHTTP)
' =============================================================================

''
' Execute HTTP request on Windows using WinHttp
' @param config - Request configuration
' @return HttpResponse - Response
''
Private Function ExecuteWindowsRequest(config As httpConfig) As httpResponse
    Dim response As httpResponse
    Dim http As Object
    Dim methodStr As String
    Dim timeoutMs As Long
    
    On Error GoTo ErrorHandler
    
    methodStr = HttpMethodToString(config.method)
    
    ' Calculate timeout in milliseconds (default 30 seconds if not set)
    If config.TimeoutSeconds <= 0 Then
        timeoutMs = 30000
    ElseIf config.TimeoutSeconds > 300 Then
        timeoutMs = 300000  ' Max 5 minutes
    Else
        timeoutMs = CLng(config.TimeoutSeconds) * 1000
    End If
    
    ' Create HTTP request object
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    
    ' Configure the request
    http.Open methodStr, config.url, False
    http.SetTimeouts timeoutMs, timeoutMs, timeoutMs, timeoutMs
    
    ' Set headers
    If config.AcceptType <> "" Then
        http.setRequestHeader "Accept", config.AcceptType
    Else
        http.setRequestHeader "Accept", "application/json"
    End If
    
    If config.contentType <> "" Then
        http.setRequestHeader "Content-Type", config.contentType
    Else
        http.setRequestHeader "Content-Type", "application/json"
    End If
    
    If config.authHeader <> "" Then
        http.setRequestHeader "Authorization", config.authHeader
    End If
    
    ' Send the request
    If config.body <> "" Then
        If Left(config.body, 1) = "@" Then
            ' File upload
            Dim fileData() As Byte
            fileData = ReadFileAsBinary(Mid(config.body, 2))
            http.Send fileData
        Else
            http.Send config.body
        End If
    Else
        http.Send
    End If
    
    ' Process response
    response.statusCode = http.status
    
    ' Use ADODB.Stream to properly decode UTF-8 response
    On Error Resume Next
    Dim responseBytes() As Byte
    responseBytes = http.responseBody
    
    If Err.Number = 0 And UBound(responseBytes) >= 0 Then
        Dim stream As Object
        Set stream = CreateObject("ADODB.Stream")
        stream.Type = 1 ' Binary
        stream.Open
        stream.Write responseBytes
        stream.position = 0
        stream.Type = 2 ' Text
        stream.Charset = "UTF-8"
        response.body = stream.ReadText
        stream.Close
        Set stream = Nothing
    Else
        ' Fallback to responseText if responseBody fails
        Err.Clear
        response.body = http.responseText
    End If
    On Error GoTo ErrorHandler
    
    If response.body = "" Then
        response.body = CStr(http.status)
    End If
    
    response.success = (http.status >= 200 And http.status < 300)
    response.IsUnauthorized = (http.status = 401)
    
    Set http = Nothing
    ExecuteWindowsRequest = response
    Exit Function
    
ErrorHandler:
    response.success = False
    response.errorMessage = Err.description
    response.statusCode = 0
    
    If Not http Is Nothing Then Set http = Nothing
    
    Core_LogError CreateError("ExecuteWindowsRequest", config.url)
    ExecuteWindowsRequest = response
End Function

#End If

' =============================================================================
' SHARED UTILITIES
' =============================================================================

''
' Parse and execute a curl command string (for Windows compatibility)
' Extracts curl parameters and executes via WinHttp
' @param curlCommand - Curl command string
' @return Long - HTTP status code
''
Public Function ParseAndExecuteCurl(curlCommand As String) As Long
    #If Mac Then
        ' On Mac, just execute the curl directly
        Dim script As String
        Dim tempFilePath As String
        
        tempFilePath = GetTempFilePath(OUTPUT_FILE_NAME)
        script = "Do shell script """ & curlCommand & " > " & tempFilePath & """"
        
        On Error Resume Next
        MacScript (script)
        
        If Err.Number = 0 Then
            ParseAndExecuteCurl = 200
        Else
            ParseAndExecuteCurl = 0
        End If
        On Error GoTo 0
    #Else
        ' On Windows, parse the curl command and use WinHttp
        ParseAndExecuteCurl = ParseCurlForWindows(curlCommand)
    #End If
End Function

#If Not Mac Then
''
' Parse curl command and execute via WinHttp (Windows only)
' @param curlCommand - Curl command string
' @return Long - HTTP status code
''
Private Function ParseCurlForWindows(curlCommand As String) As Long
    Dim regEx As Object
    Dim method As String
    Dim url As String
    Dim bodyFile As String
    Dim bodyRaw As String
    Dim headers As Object
    Dim headerMatches As Object
    Dim match As Object
    Dim http As Object
    Dim pos As Long
    
    On Error GoTo ErrorHandler
    
    Set headers = CreateObject("Scripting.Dictionary")
    Set regEx = CreateObject("VBScript.RegExp")
    
    ' Extract method
    regEx.Pattern = "--request\s+(\w+)"
    regEx.IgnoreCase = True
    regEx.Global = False
    If regEx.Test(curlCommand) Then
        method = regEx.Execute(curlCommand)(0).SubMatches(0)
    Else
        method = "GET"
    End If
    
    ' Extract URL
    regEx.Pattern = "--location\s+'([^']+)'"
    If regEx.Test(curlCommand) Then
        url = regEx.Execute(curlCommand)(0).SubMatches(0)
    End If
    
    ' Extract headers
    regEx.Pattern = "--header\s+'([^']+)'"
    regEx.Global = True
    Set headerMatches = regEx.Execute(curlCommand)
    For Each match In headerMatches
        pos = InStr(match.SubMatches(0), ":")
        If pos > 0 Then
            headers(Trim(Left(match.SubMatches(0), pos - 1))) = Trim(Mid(match.SubMatches(0), pos + 1))
        End If
    Next
    
    ' Extract data-binary file
    regEx.Pattern = "--data-binary\s+'@([^']+)'"
    regEx.Global = False
    If regEx.Test(curlCommand) Then
        bodyFile = regEx.Execute(curlCommand)(0).SubMatches(0)
        bodyFile = Replace(bodyFile, "\\\\", "\")
    End If
    
    ' Extract data-raw body
    regEx.Pattern = "--data-raw\s+'([^']+)'"
    If regEx.Test(curlCommand) Then
        bodyRaw = regEx.Execute(curlCommand)(0).SubMatches(0)
        bodyRaw = Replace(bodyRaw, "\""", """")
    End If
    
    ' Execute HTTP request
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    http.Open method, url, False
    
    ' Apply headers
    Dim hdr As Variant
    For Each hdr In headers.keys
        http.setRequestHeader hdr, headers(hdr)
    Next
    
    ' Send body
    If Len(bodyFile) > 0 Then
        Dim fileData() As Byte
        fileData = ReadFileAsBinary(bodyFile)
        http.Send fileData
    ElseIf Len(bodyRaw) > 0 Then
        http.Send bodyRaw
    Else
        http.Send
    End If
    
    ParseCurlForWindows = http.status
    Set http = Nothing
    Exit Function
    
ErrorHandler:
    ParseCurlForWindows = 0
End Function

''
' Read a file as binary data (Windows only)
' @param filePath - Path to the file
' @return Byte() - File contents as bytes
''
Private Function ReadFileAsBinary(filePath As String) As Byte()
    Dim fileStream As Object
    Dim fileData() As Byte
    
    Set fileStream = CreateObject("ADODB.Stream")
    fileStream.Type = 1 ' Binary
    fileStream.Open
    fileStream.LoadFromFile filePath
    fileData = fileStream.Read
    fileStream.Close
    Set fileStream = Nothing
    
    ReadFileAsBinary = fileData
End Function
#End If

''
' Simple GET request helper
' @param url - URL to fetch
' @param authHeader - Authorization header
' @return String - Response body or empty on error
''
Public Function SimpleGet(url As String, authHeader As String) As String
    Dim config As httpConfig
    Dim response As httpResponse
    
    config = BuildJsonConfig(url, HttpMethod_GET, authHeader)
    response = ExecuteRequest(config)
    
    If response.success Then
        SimpleGet = response.body
    Else
        SimpleGet = ""
    End If
End Function

''
' Simple POST request helper
' @param url - URL to post to
' @param authHeader - Authorization header
' @param body - Request body
' @return String - Response body or empty on error
''
Public Function SimplePost(url As String, authHeader As String, body As String) As String
    Dim config As httpConfig
    Dim response As httpResponse
    
    config = BuildJsonConfig(url, HttpMethod_POST, authHeader, body)
    response = ExecuteRequest(config)
    
    If response.success Then
        SimplePost = response.body
    Else
        SimplePost = ""
    End If
End Function


