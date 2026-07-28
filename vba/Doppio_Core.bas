Attribute VB_Name = "Doppio_Core"
''
' Doppio Core Module
' Core types, constants, enumerations, and utility functions
'
' @module Doppio_Core
' @author Doppio Group - eric@doppiogroup.com
' @version 2.0
''
Option Explicit

' =============================================================================
' VERSION
' =============================================================================
Public Const DOPPIO_VERSION As String = "v2.07"

' =============================================================================
' COMPILER CONSTANTS
' =============================================================================
#Const DEBUG_MODE = True

' =============================================================================
' ENUMERATIONS
' =============================================================================

' API Types supported by Doppio
Public Enum apiType
    ApiType_MI = 1          ' M3 API (MI Programs)
    ApiType_IDM = 2         ' Infor Document Management
    ApiType_IPS = 3         ' IPS/SOAP Web Services
    ApiType_FileMng = 4     ' File Management
    ApiType_XtendM3 = 5     ' XtendM3 Extensions
    ApiType_FNC = 6         ' Function calls
    ApiType_ExportMI = 7    ' Export MI (SQL-like queries)
End Enum

' HTTP Methods
Public Enum httpMethod
    HttpMethod_GET = 1
    HttpMethod_POST = 2
    HttpMethod_PUT = 3
    HttpMethod_DELETE = 4
End Enum

' API Result Status
Public Enum ApiStatus
    ApiStatus_Success = 0
    ApiStatus_Error = 1
    ApiStatus_Unauthorized = 2
    ApiStatus_Timeout = 3
    ApiStatus_NetworkError = 4
    ApiStatus_ParseError = 5
End Enum

' Column Direction (Input/Output)
Public Enum ColumnDirection
    ColumnDirection_Input = 1
    ColumnDirection_Output = 2
    ColumnDirection_Both = 3
End Enum

' Company/Division display options
Public Enum ConoDiviOption
    ConoDivi_Hide = 0
    ConoDivi_ShowFirst = 1
    ConoDivi_Remove = 2
    ConoDivi_ShowLast = 3
End Enum

' =============================================================================
' UI COLORS (RGB values as Long)
' =============================================================================
Public Const COLOR_MANDATORY As Long = 130          ' RGB(130, 0, 0) - Dark red
Public Const COLOR_OPTIONAL As Long = 4210752     ' RGB(64, 64, 64) - Dark gray
Public Const COLOR_OUTPUT As Long = 8421504       ' RGB(128, 128, 128) - Gray
Public Const COLOR_ERROR As Long = 255            ' RGB(255, 0, 0) - Red
Public Const COLOR_SUCCESS As Long = 40000        ' Green
Public Const COLOR_HEADER_BG As Long = 14013909   ' RGB(213, 232, 240) - Light blue

' =============================================================================
' DEFAULT VALUES
' =============================================================================
Public Const DEFAULT_MAX_RECORDS As Long = 1000
Public Const DEFAULT_MAX_BULK As Integer = 1000
Public Const DEFAULT_MAX_TIMEOUT As Integer = 30
Public Const DEFAULT_REFRESH_SECONDS As Integer = 300
Public Const DEFAULT_SPLIT_CHAR As String = ","
Public Const DEFAULT_SHEET_NAMING As Integer = 0

' =============================================================================
' FILE PATHS
' =============================================================================
Public Const OUTPUT_FILE_NAME As String = "curl_output.txt"
Public Const INPUT_FILE_NAME As String = "curl_input.sh"

' =============================================================================
' API PATHS
' =============================================================================
Public Const MI_API_PATH As String = "/M3/m3api-rest/v2/execute"
Public Const IDM_API_PATH As String = "/IDM/api"
Public Const FILE_MGT_PATH As String = "/M3/foundation-rest/file-management/v1"

' =============================================================================
' TYPE DEFINITIONS
' =============================================================================

' Tenant/Environment Configuration
Public Type TenantConfig
    tenantId As String          ' ti - Tenant identifier
    clientId As String          ' ci - OAuth client ID
    clientSecret As String      ' cs - OAuth client secret
    IonUrl As String            ' iu - ION API base URL
    SsoUrl As String            ' pu - SSO base URL
    authEndpoint As String      ' oa - Authorization endpoint
    tokenEndpoint As String     ' ot - Token endpoint
    redirectUri As String       ' ru - OAuth redirect URI
    ServiceAccountKey As String ' saak - Service account access key
    ServiceAccountSecret As String ' sask - Service account secret
    SingleTenantUrl As String   ' url - For single-tenant (non-MT) environments
    SingleTenantUser As String  ' user - Basic auth user
    SingleTenantPassword As String ' password - Basic auth password
End Type

' API Settings
Public Type ApiSettings
    MaxRecords As Long          ' Maximum records to retrieve
    maxbulk As Integer          ' Maximum bulk operations
    MaxTimeout As Integer       ' HTTP timeout in seconds
    refreshSeconds As Integer   ' Token refresh interval
    righttrim As Boolean        ' Trim trailing spaces
    formatting As Boolean       ' Apply cell formatting
    splitChar As String         ' Character for splitting values
    conoDivi As ConoDiviOption  ' Company/Division display option
    sheetNaming As Integer      ' Sheet naming method (0=API+Txn, 1=Txn, 2=API, 3=first 6 of API)
End Type

' Session State (runtime values)
Public Type SessionState
    AccessToken As String       ' Current OAuth access token
    RefreshToken As String      ' OAuth refresh token
    TokenType As String         ' Token type (usually "Bearer")
    selectedEnvironment As String ' Currently selected environment name
    mainUrl As String           ' Constructed main URL (iu/ti)
    m3user As String            ' M3 user for API calls
    company As String           ' Current company
    division As String          ' Current division
    IsMultitenant As Boolean    ' True for MT, False for single-tenant
    SingleTenantToken As String ' Base64 encoded credentials for ST
End Type

' HTTP Request Configuration
Public Type httpConfig
    url As String
    method As httpMethod
    contentType As String
    AcceptType As String
    authHeader As String
    body As String
    timeoutSeconds As Integer
    WriteToFile As Boolean      ' Mac: write response to file
End Type

' HTTP Response
Public Type httpResponse
    statusCode As Long
    body As String
    success As Boolean
    errorMessage As String
    IsUnauthorized As Boolean
End Type

' API Request
Public Type ApiRequest
    mainUrl As String
    apiPath As String
    endpoint As String
    body As String
    apiType As apiType
    UseCache As Boolean
    cacheKey As String
End Type

' API Response
Public Type apiResponse
    success As Boolean
    status As ApiStatus
    data As String              ' Raw response data
    errorMessage As String
    records As Object           ' Parsed records collection
    results As Object           ' Parsed results object
    recordCount As Long
End Type

' Error Information
Public Type DoppioError
    Number As Long
    description As String
    source As String
    Timestamp As Date
    additionalInfo As String
End Type

' Column Metadata
Public Type ColumnInfo
    name As String              ' Field name (e.g., "CUNO")
    description As String       ' Field description
    dataType As String          ' A=Alpha, N=Numeric, D=Date
    direction As ColumnDirection
    isMandatory As Boolean
    Length As Integer
End Type

' Version History Entry
Public Type VersionEntry
    version     As String       ' Version number (e.g., "2.06")
    description As String       ' What changed
    releaseDate As String       ' ISO date string "YYYY-MM-DD", empty if none
    status      As String       ' e.g., "complete"
End Type

' =============================================================================
' UTILITY FUNCTIONS
' =============================================================================

''
' Get the user's home directory path
' @return String - Home directory path
''
Public Function GetHomePath() As String
    #If Mac Then
        GetHomePath = Environ("HOME")
    #Else
        GetHomePath = Environ("USERPROFILE")
    #End If
End Function

''
' Get the full path to a temp file
' @param fileName - Name of the file
' @return String - Full path
''
Public Function GetTempFilePath(fileName As String) As String
    GetTempFilePath = GetHomePath() & "/" & fileName
End Function

''
' Check if we're running on Mac
' @return Boolean - True if Mac
''
Public Function IsMac() As Boolean
    #If Mac Then
        IsMac = True
    #Else
        IsMac = False
    #End If
End Function

''
' Convert ApiType enum to string
' @param apiType - The API type
' @return String - String representation
''
Public Function ZZZ_ApiTypeToString(apiType As apiType) As String
    Select Case apiType
        Case ApiType_MI: ZZZ_ApiTypeToString = "API"
        Case ApiType_IDM: ZZZ_ApiTypeToString = "IDM"
        Case ApiType_IPS: ZZZ_ApiTypeToString = "IPS"
        Case ApiType_FileMng: ZZZ_ApiTypeToString = "FileMng"
        Case ApiType_XtendM3: ZZZ_ApiTypeToString = "XtendM3"
        Case ApiType_FNC: ZZZ_ApiTypeToString = "FNC"
        Case ApiType_ExportMI: ZZZ_ApiTypeToString = "EXPORTMI"
        Case Else: ZZZ_ApiTypeToString = "API"
    End Select
End Function

''
' Convert string to ApiType enum
' @param typeStr - String representation
' @return ApiType - The enum value
''
Public Function StringToApiType(typeStr As String) As apiType
    Select Case UCase(Trim(typeStr))
        Case "API", "MI": StringToApiType = ApiType_MI
        Case "IDM": StringToApiType = ApiType_IDM
        Case "IPS": StringToApiType = ApiType_IPS
        Case "FILEMNG": StringToApiType = ApiType_FileMng
        Case "XTENDM3": StringToApiType = ApiType_XtendM3
        Case "FNC": StringToApiType = ApiType_FNC
        Case "EXPORTMI": StringToApiType = ApiType_ExportMI
        Case Else: StringToApiType = ApiType_MI
    End Select
End Function

''
' Convert HttpMethod enum to string
' @param method - The HTTP method
' @return String - String representation
''
Public Function HttpMethodToString(method As httpMethod) As String
    Select Case method
        Case HttpMethod_GET: HttpMethodToString = "GET"
        Case HttpMethod_POST: HttpMethodToString = "POST"
        Case HttpMethod_PUT: HttpMethodToString = "PUT"
        Case HttpMethod_DELETE: HttpMethodToString = "DELETE"
        Case Else: HttpMethodToString = "GET"
    End Select
End Function

''
' Safe string conversion - handles errors and Nothing
' @param value - Value to convert
' @return String - Converted string or empty
''
Public Function SafeStr(value As Variant) As String
    On Error Resume Next
    If IsNull(value) Or IsEmpty(value) Or isError(value) Then
        SafeStr = ""
    ElseIf IsObject(value) Then
        If value Is Nothing Then
            SafeStr = ""
        Else
            SafeStr = CStr(value)
        End If
    Else
        SafeStr = CStr(value)
    End If
    On Error GoTo 0
End Function

''
' Safe long conversion - handles errors
' @param value - Value to convert
' @param defaultValue - Default if conversion fails
' @return Long - Converted value or default
''
Public Function SafeLong(value As Variant, Optional defaultValue As Long = 0) As Long
    On Error Resume Next
    If IsNumeric(value) Then
        SafeLong = CLng(value)
    Else
        SafeLong = defaultValue
    End If
    On Error GoTo 0
End Function

''
' Check if a worksheet exists
' @param sheetName - Name of the sheet
' @param Optional wb - Workbook to check (defaults to ThisWorkbook)
' @return Boolean - True if sheet exists
''
Public Function Core_SheetExists(sheetName As String, Optional wb As Workbook = Nothing) As Boolean
    Dim ws As Worksheet
    
    If wb Is Nothing Then Set wb = ThisWorkbook
    
    On Error Resume Next
    Set ws = wb.Sheets(sheetName)
    Core_SheetExists = Not ws Is Nothing
    On Error GoTo 0
End Function

''
' Check if a sheet is visible
' @param sheetName - Name of the sheet
' @return Boolean - True if visible
''
Public Function Core_IsSheetVisible(sheetName As String) As Boolean
    On Error Resume Next
    Core_IsSheetVisible = (ThisWorkbook.Sheets(sheetName).Visible = xlSheetVisible)
    On Error GoTo 0
End Function

''
' Base64 encode a string
' @param text - Text to encode
' @return String - Base64 encoded string
''
Public Function Base64Encode(text As String) As String
    Dim bytes() As Byte
    Dim objXML As Object
    Dim objNode As Object
    
    On Error GoTo ErrorHandler
    
    ' Convert the text to a byte array
    bytes = StrConv(text, vbFromUnicode)
    
    ' Create an XML document to handle Base64 encoding
    Set objXML = CreateObject("MSXML2.DOMDocument")
    Set objNode = objXML.createElement("b64")
    
    ' Encode the byte array as Base64
    objNode.dataType = "bin.base64"
    objNode.nodeTypedValue = bytes
    Base64Encode = objNode.text
    
    ' Remove line breaks
    Base64Encode = Replace(Base64Encode, vbLf, "")
    Base64Encode = Replace(Base64Encode, vbCr, "")
    
    ' Clean up
    Set objNode = Nothing
    Set objXML = Nothing
    Exit Function
    
ErrorHandler:
    Base64Encode = ""
End Function

''
' Base64 decode a string
' @param encodedText - Base64 encoded text
' @return String - Decoded string
''
Public Function Base64Decode(encodedText As String) As String
    Dim objXML As Object
    Dim objNode As Object
    Dim bytes() As Byte
    
    On Error GoTo ErrorHandler
    
    ' Restore Base64 padding if it was removed
    Do While Len(encodedText) Mod 4 <> 0
        encodedText = encodedText & "="
    Loop
    
    ' Create an XML document to handle Base64 decoding
    Set objXML = CreateObject("MSXML2.DOMDocument")
    Set objNode = objXML.createElement("b64")
    
    objNode.dataType = "bin.base64"
    objNode.text = encodedText
    bytes = objNode.nodeTypedValue
    
    ' Convert the byte array back to string
    Base64Decode = StrConv(bytes, vbUnicode)
    
    ' Clean up
    Set objNode = Nothing
    Set objXML = Nothing
    Exit Function
    
ErrorHandler:
    Base64Decode = ""
End Function

''
' Log an error to the debug window and optionally to a sheet
' @param errInfo - Error information structure
' @param Optional writeToSheet - Whether to write to Log sheet
''
Public Sub Core_LogError(errInfo As DoppioError, Optional writeToSheet As Boolean = True)
    Dim logMessage As String
    
    logMessage = Format(errInfo.Timestamp, "yyyy-mm-dd hh:nn:ss") & _
                 " [" & errInfo.source & "] " & _
                 "Error " & errInfo.Number & ": " & errInfo.description
    
    If errInfo.additionalInfo <> "" Then
        logMessage = logMessage & " | " & errInfo.additionalInfo
    End If
    
    #If DEBUG_MODE Then
        Debug.Print "Core_LogError: " & logMessage
    #End If
    
    If writeToSheet And Core_IsSheetVisible("Log") Then
        WriteToLogSheet errInfo
    End If
End Sub

''
' Write error to the Log sheet
' @param errInfo - Error information structure
''
Private Sub WriteToLogSheet(errInfo As DoppioError)
    Dim ws As Worksheet
    Dim NextRow As Long
    
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Log")
    If ws Is Nothing Then Exit Sub
    
    NextRow = ws.Cells(ws.Rows.count, 1).End(xlUp).row + 1
    
    ws.Cells(NextRow, 1).value = errInfo.Timestamp
    ws.Cells(NextRow, 2).value = errInfo.source
    ws.Cells(NextRow, 3).value = errInfo.Number
    ws.Cells(NextRow, 4).value = errInfo.description
    ws.Cells(NextRow, 5).value = errInfo.additionalInfo
    
    On Error GoTo 0
End Sub

''
' Create a DoppioError from current Err object
' @param source - Source function/module name
' @param Optional additionalInfo - Extra context
' @return DoppioError - Populated error structure
''
Public Function CreateError(source As String, Optional additionalInfo As String = "") As DoppioError
    Dim errInfo As DoppioError
    
    errInfo.Number = Err.Number
    errInfo.description = Err.description
    errInfo.source = source
    errInfo.Timestamp = Now
    errInfo.additionalInfo = additionalInfo
    
    CreateError = errInfo
End Function

''
' Replace non-alphanumeric characters for safe string handling
' @param inputStr - Input string
' @return String - String with alpha replaced by zero
''
Public Function Core_ReplaceAlphaWithZero(inputStr As String) As String
    Dim i As Long
    Dim char As String
    Dim result As String
    
    result = ""
    For i = 1 To Len(inputStr)
        char = Mid(inputStr, i, 1)
        If char Like "[0-9]" Or char = "." Or char = "-" Then
            result = result & char
        Else
            result = result & "0"
        End If
    Next i
    
    Core_ReplaceAlphaWithZero = result
End Function

''
' URL encode a string
' @param text - Text to encode
' @return String - URL encoded string
''
Public Function Core_UrlEncode(text As String) As String
    Dim i As Long
    Dim char As String
    Dim ascVal As Long
    Dim result As String
    
    result = ""
    For i = 1 To Len(text)
        char = Mid(text, i, 1)
        ascVal = Asc(char)
        
        ' Keep alphanumeric and some safe characters
        If (ascVal >= 48 And ascVal <= 57) Or _
           (ascVal >= 65 And ascVal <= 90) Or _
           (ascVal >= 97 And ascVal <= 122) Or _
           char = "-" Or char = "_" Or char = "." Or char = "~" Then
            result = result & char
        Else
            result = result & "%" & Right("0" & Hex(ascVal), 2)
        End If
    Next i
    
    Core_UrlEncode = result
End Function
Public Function Core_ReadFileToString(filePath As String) As String
    On Error GoTo ErrorHandler
    
    #If Mac Then
        ' Mac: Use "do shell script cat" to read the file.
        ' The shell outputs UTF-8; AppleScript converts it to Unicode before
        ' handing it to MacScript, which returns Mac Roman to VBA.
        ' This correctly maps characters like Í (U+00CD, Mac Roman 0xEA)
        ' whereas the old "read ... as utf8" approach was failing silently
        ' and falling back to a raw byte read, producing garbage like √ç.
        Dim script As String
        Dim result As String

        On Error Resume Next
        script = "do shell script ""cat '" & filePath & "'"""
        result = MacScript(script)

        If Err.Number = 0 And Len(result) > 0 Then
            Core_ReadFileToString = result
        Else
            ' Fallback: raw byte read (no encoding conversion � ASCII-safe only)
            Err.Clear
            Dim fileNumber As Integer
            Dim fileContent As String

            fileNumber = FreeFile
            Open filePath For Input As #fileNumber
            fileContent = Input$(LOF(fileNumber), fileNumber)
            Close #fileNumber

            Core_ReadFileToString = fileContent
        End If
        On Error GoTo 0
    #Else
        ' Windows: Use ADODB.Stream for UTF-8
        Dim stream As Object
        
        Set stream = CreateObject("ADODB.Stream")
        stream.Type = 2 ' Text
        stream.Charset = "UTF-8"
        stream.Open
        stream.LoadFromFile filePath
        Core_ReadFileToString = stream.ReadText
        stream.Close
        Set stream = Nothing
    #End If
    
    Exit Function
    
ErrorHandler:
    ' Fallback to basic file reading if everything else fails
    On Error Resume Next
    Dim fn As Integer
    Dim fc As String
    
    fn = FreeFile
    Open filePath For Input As fn
    fc = Input$(LOF(fn), fn)
    Close fn
    
    Core_ReadFileToString = fc
    On Error GoTo 0
End Function


''
' Write string to file
' @param filePath - Full path to file
' @param content - Content to write
' @return Boolean - True if successful
''
Public Function WriteStringToFile(filePath As String, content As String) As Boolean
    Dim fileNumber As Integer
    
    On Error GoTo ErrorHandler
    
    fileNumber = FreeFile
    Open filePath For Output As #fileNumber
    Print #fileNumber, content
    Close #fileNumber
    
    WriteStringToFile = True
    Exit Function
    
ErrorHandler:
    WriteStringToFile = False
    On Error GoTo 0
End Function

''
' Delete a file if it exists
' @param filePath - Full path to file
''
Public Sub DeleteFileIfExists(filePath As String)
    On Error Resume Next
    Kill filePath
    On Error GoTo 0
End Sub

' =============================================================================
' VERSION HISTORY
' =============================================================================

''
' Writes the embedded version history to the Versions sheet.
' Assign this Sub directly to a button on the Versions sheet.
' Update the entries array here whenever a new version is released.
''
Public Sub Core_GetVersionHistory()
    Dim ws As Worksheet
    Dim e(0 To 6) As VersionEntry
    Dim i As Long

    On Error GoTo ErrorHandler

    If Not Core_SheetExists("Versions") Then
        MsgBox "Versions sheet not found in this workbook.", vbExclamation, "Doppio"
        Exit Sub
    End If

    ' --- Version data: update this block for each release ---
    e(0).version = "1.3"
    e(0).description = "Released to Doppio Caf�"
    e(0).releaseDate = ""
    e(0).status = "complete"

    e(1).version = "2.02"
    e(1).description = "Skip empty fields in MI transaction body"
    e(1).releaseDate = "2026-03-20"
    e(1).status = "complete"

    e(2).version = "2.03"
    e(2).description = "Fix retry"
    e(2).releaseDate = "2026-03-27"
    e(2).status = "complete"

    e(3).version = "2.04"
    e(3).description = "Format lookup moved from Environments tab to Logos sheet"
    e(3).releaseDate = "2026-04-03"
    e(3).status = "complete"

    e(4).version = "2.05"
    e(4).description = "AutoFit toggle - first click fits row 7 headers, second click fits full UsedRange"
    e(4).releaseDate = "2026-04-06"
    e(4).status = "complete"

    e(5).version = "2.06"
    e(5).description = "Added transpose keyword"
    e(5).releaseDate = "2026-04-07"
    e(5).status = "complete"

    e(6).version = "2.07"
    e(6).description = "Added new setting for sheet renaming"
    e(6).releaseDate = "2026-04-08"
    e(6).status = "complete"
    ' --- End version data ---

    Set ws = ActiveWorkbook.Sheets("Versions")

    ' Clear data rows via column A last-used row (avoids UsedRange quirks)
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).row
    If lastRow > 1 Then ws.Range("A2:D" & lastRow).ClearContents

    ' Write / refresh header row
    ws.Cells(1, 1).value = "Version"
    ws.Cells(1, 2).value = "Description"
    ws.Cells(1, 3).value = "Date"
    ws.Cells(1, 4).value = "Status"

    ' Write version rows
    For i = 0 To UBound(e)
        ws.Cells(i + 2, 1).value = e(i).version
        ws.Cells(i + 2, 2).value = e(i).description
        If e(i).releaseDate <> "" Then
            ws.Cells(i + 2, 3).value = CDate(e(i).releaseDate)
            ws.Cells(i + 2, 3).NumberFormat = "yyyy-mm-dd"
        End If
        ws.Cells(i + 2, 4).value = e(i).status
    Next i

    #If DEBUG_MODE Then
        Debug.Print "Core_GetVersionHistory: wrote " & (UBound(e) + 1) & " entries."
    #End If

    Exit Sub

ErrorHandler:
    Debug.Print "Core_GetVersionHistory error " & Err.Number & ": " & Err.description
End Sub

