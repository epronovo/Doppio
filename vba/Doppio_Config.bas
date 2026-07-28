Attribute VB_Name = "Doppio_Config"
''
' Doppio Configuration Module
' Handles environment configuration, settings, and tenant management
'
' @module Doppio_Config
' @author Doppio Group - eric@doppiogroup.com
' @version 2.0
''
Option Explicit

' =============================================================================
' MODULE-LEVEL STATE
' =============================================================================

' Current configuration (loaded from environment)
Private m_TenantConfig As TenantConfig
Private m_ApiSettings As ApiSettings
Private m_SessionState As SessionState

' Environment manager instance
Private m_EnvironmentManager As EnvironmentManager

' Flag to track if settings are initialized
Private m_Initialized As Boolean

' =============================================================================
' INITIALIZATION
' =============================================================================

''
' Initialize configuration with default values
' Call this on workbook open or when resetting
''
Public Sub InitializeConfig()
    ' Set default API settings
    With m_ApiSettings
        .MaxRecords = DEFAULT_MAX_RECORDS
        .maxbulk = DEFAULT_MAX_BULK
        .MaxTimeout = DEFAULT_MAX_TIMEOUT
        .refreshSeconds = DEFAULT_REFRESH_SECONDS
        .righttrim = True
        .formatting = True
        .splitChar = DEFAULT_SPLIT_CHAR
        .conoDivi = ConoDivi_Hide
    End With
    
    ' Clear session state
    ClearSessionState
    
    ' Initialize environment manager
    If m_EnvironmentManager Is Nothing Then
        Set m_EnvironmentManager = New EnvironmentManager
    End If
    
    ' Try to load settings from Settings sheet
    Config_LoadSettingsFromSheet
    
    m_Initialized = True
End Sub

''
' Clear the current session state
''
Public Sub ClearSessionState()
    With m_SessionState
        .AccessToken = ""
        .RefreshToken = ""
        .TokenType = ""
        .selectedEnvironment = ""
        .mainUrl = ""
        .m3user = ""
        .company = ""
        .division = ""
        .IsMultitenant = True
        .SingleTenantToken = ""
    End With
End Sub

' =============================================================================
' PROPERTY ACCESSORS
' =============================================================================

''
' Get the current tenant configuration
' @return TenantConfig - Current tenant config
''
Public Property Get Config_TenantConfig() As TenantConfig
    Config_TenantConfig = m_TenantConfig
End Property

''
' Get the current API settings
' @return ApiSettings - Current settings
''
Public Property Get Config_ApiSettings() As ApiSettings
    If Not m_Initialized Then InitializeConfig
    Config_ApiSettings = m_ApiSettings
End Property

''
' Get the current session state
' @return SessionState - Current session
''
Public Property Get Config_SessionState() As SessionState
    Config_SessionState = m_SessionState
End Property

''
' Get the environment manager
' @return EnvironmentManager - Manager instance
''
Public Property Get Config_EnvironmentManager() As EnvironmentManager
    If m_EnvironmentManager Is Nothing Then
        Set m_EnvironmentManager = New EnvironmentManager
    End If
    Set Config_EnvironmentManager = m_EnvironmentManager
End Property

' Individual setting accessors for convenience
Public Property Get Config_MaxRecords() As Long
    Config_MaxRecords = m_ApiSettings.MaxRecords
End Property

Public Property Let Config_MaxRecords(value As Long)
    m_ApiSettings.MaxRecords = value
End Property

Public Property Get Config_MaxBulk() As Integer
    Config_MaxBulk = m_ApiSettings.maxbulk
End Property

Public Property Let Config_MaxBulk(value As Integer)
    m_ApiSettings.maxbulk = value
End Property

Public Property Get Config_MaxTimeout() As Integer
    Config_MaxTimeout = m_ApiSettings.MaxTimeout
End Property

Public Property Let Config_MaxTimeout(value As Integer)
    m_ApiSettings.MaxTimeout = value
End Property

Public Property Get Config_Developer() As Boolean
    If Not m_Initialized Then InitializeConfig
    Config_Developer = m_ApiSettings.developer
End Property

Public Property Let Config_Developer(value As Boolean)
    m_ApiSettings.developer = value
End Property

Public Property Get Config_AccessToken() As String
    Config_AccessToken = m_SessionState.AccessToken
End Property

Public Property Let Config_AccessToken(value As String)
    m_SessionState.AccessToken = value
End Property

Public Property Get Config_RefreshToken() As String
    Config_RefreshToken = m_SessionState.RefreshToken
End Property

Public Property Let Config_RefreshToken(value As String)
    m_SessionState.RefreshToken = value
End Property

Public Property Get Config_TokenType() As String
    Config_TokenType = m_SessionState.TokenType
End Property

Public Property Let Config_TokenType(value As String)
    m_SessionState.TokenType = value
End Property

Public Property Get Config_SelectedEnvironment() As String
    Config_SelectedEnvironment = m_SessionState.selectedEnvironment
End Property

Public Property Let Config_SelectedEnvironment(value As String)
    m_SessionState.selectedEnvironment = value
End Property

Public Property Get Config_MainUrl() As String
    Config_MainUrl = m_SessionState.mainUrl
End Property

Public Property Get Config_M3User() As String
    Config_M3User = m_SessionState.m3user
End Property

Public Property Let Config_M3User(value As String)
    m_SessionState.m3user = value
End Property

Public Property Get Config_Company() As String
    Config_Company = m_SessionState.company
End Property

Public Property Let Config_Company(value As String)
    m_SessionState.company = value
End Property

Public Property Get Config_Division() As String
    Config_Division = m_SessionState.division
End Property

Public Property Let Config_Division(value As String)
    m_SessionState.division = value
End Property

Public Property Get Config_IsMultitenant() As Boolean
    Config_IsMultitenant = m_SessionState.IsMultitenant
End Property

' =============================================================================
' ENVIRONMENT MANAGEMENT
' =============================================================================

''
' Load environments from .ionapi files in the workbook directory
''
Public Sub Config_LoadEnvironmentsFromFiles()
    Dim ws As Worksheet
    Dim folderPath As String
    Dim ionApiFile As String
    Dim rowNum As Long
    Dim jsonData As String
    Dim json As Object
    Dim dirList As ArrayList
    Dim i As Long
    Dim fileName As String
    
    On Error GoTo ErrorHandler
    
    ' Get or create the Environments sheet
    Set ws = GetOrCreateSheet("Environments", xlSheetHidden)
    ws.Columns("A:G").ClearContents
    
    ' Get the directory of the current workbook
    folderPath = ThisWorkbook.path
    If folderPath = "" Then
        folderPath = GetHomePath() & "/Documents"
    End If
    
    ' Collect all .ionapi files
    Set dirList = New ArrayList
    dirList.Initialize
    
    ionApiFile = Dir(folderPath & "/*.ionapi")
    Do While ionApiFile <> ""
        dirList.Add ionApiFile
        ionApiFile = Dir
    Loop
    dirList.Sort
    
    ' Process each file
    rowNum = 1
    For i = 1 To dirList.count
        fileName = dirList.item(i)
        jsonData = Core_ReadFileToString(folderPath & "/" & fileName)
        
        On Error Resume Next
        Set json = ParseJson(jsonData)
        
        ' Only add if it has required fields
        If json.item("saak") <> "" Or json.item("ci") <> "" Then
            ws.Cells(rowNum, 1).value = Left(fileName, InStrRev(fileName, ".") - 1)
            ws.Cells(rowNum, 2).value = jsonData
            ws.Cells(rowNum, 3).value = " "
            ws.Cells(rowNum, 4).value = " "
            ws.Cells(rowNum, 5).value = " "
            ws.Cells(rowNum, 6).value = " "
            rowNum = rowNum + 1
        End If
        On Error GoTo ErrorHandler
    Next i
    
    ws.Columns("B").WrapText = False
    Exit Sub
    
ErrorHandler:
    Core_LogError CreateError("LoadEnvironmentsFromFiles", Err.description)
End Sub

''
' Load tenant configuration for the selected environment
' @param environmentName - Name of the environment to load
' @return Boolean - True if successful
''
Public Function Config_LoadTenantConfig(EnvironmentName As String) As Boolean
    Dim ws As Worksheet
    Dim environmentRange As Range
    Dim targetCell As Range
    Dim jsonString As String
    Dim json As Object
    
    On Error GoTo ErrorHandler
    
    If EnvironmentName = "" Then
        Config_LoadTenantConfig = False
        Exit Function
    End If
    
    ' Find the environment in the Environments sheet
    Set ws = ThisWorkbook.Sheets("Environments")
    Set environmentRange = ws.Range("A:A")
    Set targetCell = environmentRange.Find(What:=EnvironmentName, LookIn:=xlValues, LookAt:=xlWhole)
    
    If targetCell Is Nothing Then
        Config_LoadTenantConfig = False
        Exit Function
    End If
    
    ' Parse the JSON configuration
    jsonString = targetCell.Offset(0, 1).value
    Set json = ParseJson(jsonString)
    
    ' Populate TenantConfig
    With m_TenantConfig
        .tenantId = SafeStr(json.item("ti"))
        .clientId = SafeStr(json.item("ci"))
        .clientSecret = SafeStr(json.item("cs"))
        .IonUrl = SafeStr(json.item("iu"))
        .SsoUrl = SafeStr(json.item("pu"))
        .authEndpoint = SafeStr(json.item("oa"))
        .tokenEndpoint = SafeStr(json.item("ot"))
        .redirectUri = SafeStr(json.item("ru"))
        .ServiceAccountKey = SafeStr(json.item("saak"))
        .ServiceAccountSecret = SafeStr(json.item("sask"))
        .SingleTenantUrl = SafeStr(json.item("url"))
        .SingleTenantUser = SafeStr(json.item("user"))
        .SingleTenantPassword = SafeStr(json.item("password"))
    End With
    
    ' Update session state
    m_SessionState.selectedEnvironment = EnvironmentName
    m_SessionState.m3user = targetCell.Offset(0, 3).value

    ' Load cached token from sheet if available (column C)
    Dim cachedToken As String
    cachedToken = Trim(CStr(targetCell.Offset(0, 2).value))
    If cachedToken <> "" Then
        m_SessionState.AccessToken = cachedToken
        m_SessionState.TokenType = "Bearer"
    End If

    ' Determine if multi-tenant or single-tenant
    If m_TenantConfig.SingleTenantUrl <> "" Then
        m_SessionState.IsMultitenant = False
        m_SessionState.SingleTenantToken = Base64Encode( _
            m_TenantConfig.SingleTenantUser & ":" & m_TenantConfig.SingleTenantPassword)
        m_SessionState.mainUrl = m_TenantConfig.SingleTenantUrl
    Else
        m_SessionState.IsMultitenant = True
        m_SessionState.mainUrl = m_TenantConfig.IonUrl & "/" & m_TenantConfig.tenantId
    End If
    
    ' Update environment manager
    UpdateEnvironmentManager EnvironmentName, jsonString
    
    Config_LoadTenantConfig = True
    Exit Function
    
ErrorHandler:
    Core_LogError CreateError("LoadTenantConfig", EnvironmentName)
    Config_LoadTenantConfig = False
End Function

''
' Update the environment manager with current configuration
''
Private Sub UpdateEnvironmentManager(envName As String, jsonString As String)
    Dim env As Environment
    
    If m_EnvironmentManager Is Nothing Then
        Set m_EnvironmentManager = New EnvironmentManager
    End If
    
    Set env = m_EnvironmentManager.GetEnvironment(envName)
    
    If Not m_EnvironmentManager.HasEnvironment(envName) Then
        m_EnvironmentManager.AddEnvironment envName, envName, jsonString, "", _
            m_SessionState.mainUrl, m_SessionState.m3user, "", ""
    ElseIf env.User = "" Then
        m_EnvironmentManager.AddEnvironment envName, envName, jsonString, "", _
            m_SessionState.mainUrl, m_SessionState.m3user, "", ""
    End If
End Sub

' =============================================================================
' SETTINGS PERSISTENCE
' =============================================================================

''
' Load settings from the Settings sheet
''
Public Sub Config_LoadSettingsFromSheet()
    Dim ws As Worksheet
    
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Settings")
    If ws Is Nothing Then Exit Sub
    
    With m_ApiSettings
        .MaxRecords = SafeLong(ws.Range("maxrecs").value, DEFAULT_MAX_RECORDS)
        .maxbulk = CInt(SafeLong(ws.Range("maxbulk").value, DEFAULT_MAX_BULK))
        .MaxTimeout = CInt(SafeLong(ws.Range("maxtime").value, DEFAULT_MAX_TIMEOUT))
        .refreshSeconds = CInt(SafeLong(ws.Range("refreshSeconds").value, DEFAULT_REFRESH_SECONDS))
        .righttrim = CBool(ws.Range("righttrim").value)
        .formatting = CBool(ws.Range("formatting").value)
        .splitChar = SafeStr(ws.Range("splitChar").value)
        If .splitChar = "" Then .splitChar = DEFAULT_SPLIT_CHAR
        .conoDivi = SafeLong(ws.Range("conoDivi").value, 0)
        .sheetNaming = CInt(SafeLong(ws.Range("naming").value, DEFAULT_SHEET_NAMING))
        .developer = CBool(ws.Range("developer").value)
    End With

    ' Keep legacy globals in sync so SettingsSheet() doesn't overwrite with stale values
    maxRecs = m_ApiSettings.MaxRecords
    maxbulk = m_ApiSettings.maxbulk
    refreshSeconds = m_ApiSettings.refreshSeconds
    righttrim = m_ApiSettings.righttrim
    formatting = m_ApiSettings.formatting
    splitChar = m_ApiSettings.splitChar
    maxtime = m_ApiSettings.MaxTimeout
    conoDivi = m_ApiSettings.conoDivi
    sheetNaming = m_ApiSettings.sheetNaming

    On Error GoTo 0
End Sub

''
' Save current settings to the Settings sheet
''
Public Sub Config_SaveSettingsToSheet()
    Dim ws As Worksheet
    Dim wasHidden As Boolean
    
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Settings")
    If ws Is Nothing Then Exit Sub
    
    wasHidden = (ws.Visible <> xlSheetVisible)
    If wasHidden Then ws.Visible = xlSheetVisible
    
    With m_ApiSettings
        ws.Range("maxrecs").value = .MaxRecords
        ws.Range("maxbulk").value = .maxbulk
        ws.Range("maxtime").value = .MaxTimeout
        ws.Range("refreshSeconds").value = .refreshSeconds
        ws.Range("righttrim").value = .righttrim
        ws.Range("formatting").value = .formatting
        ws.Range("splitChar").value = .splitChar
        ws.Range("conoDivi").value = .conoDivi
        ws.Range("naming").value = .sheetNaming
        ws.Range("developer").value = .developer
    End With
    
    If wasHidden Then ws.Visible = xlSheetVeryHidden
    On Error GoTo 0
End Sub

''
' Reset settings to defaults
''
Public Sub Config_ResetSettingsToDefaults()
    With m_ApiSettings
        .MaxRecords = DEFAULT_MAX_RECORDS
        .maxbulk = DEFAULT_MAX_BULK
        .MaxTimeout = DEFAULT_MAX_TIMEOUT
        .refreshSeconds = DEFAULT_REFRESH_SECONDS
        .righttrim = True
        .formatting = True
        .splitChar = DEFAULT_SPLIT_CHAR
        .conoDivi = ConoDivi_Hide
        .sheetNaming = DEFAULT_SHEET_NAMING
    End With

    Config_SaveSettingsToSheet
End Sub

' =============================================================================
' HELPER FUNCTIONS
' =============================================================================

''
' Get or create a worksheet
' @param sheetName - Name of the sheet
' @param Optional visibility - Sheet visibility
' @return Worksheet - The sheet
''
Private Function GetOrCreateSheet(sheetName As String, _
                                   Optional visibility As XlSheetVisibility = xlSheetVisible) As Worksheet
    Dim ws As Worksheet
    
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(sheetName)
    On Error GoTo 0
    
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add
        ws.name = sheetName
    End If
    
    ws.Visible = visibility
    Set GetOrCreateSheet = ws
End Function

''
' Get list of available environment names
' @return Collection - Collection of environment names
''
Public Function Config_GetEnvironmentList() As Collection
    Dim ws As Worksheet
    Dim envList As New Collection
    Dim row As Long
    
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Environments")
    If ws Is Nothing Then
        Set Config_GetEnvironmentList = envList
        Exit Function
    End If
    
    row = 1
    Do While ws.Cells(row, 1).value <> ""
        envList.Add ws.Cells(row, 1).value
        row = row + 1
    Loop
    
    Set Config_GetEnvironmentList = envList
    On Error GoTo 0
End Function

''
' Check if configuration is valid for API calls
' @return Boolean - True if ready to make API calls
''
Public Function Config_IsConfigValid() As Boolean
    ' Check we have an environment selected
    If m_SessionState.selectedEnvironment = "" Then
        Config_IsConfigValid = False
        Exit Function
    End If
    
    ' Check we have credentials
    If m_SessionState.IsMultitenant Then
        Config_IsConfigValid = (m_TenantConfig.clientId <> "" And m_TenantConfig.clientSecret <> "")
    Else
        Config_IsConfigValid = (m_SessionState.SingleTenantToken <> "")
    End If
End Function

''
' Check if we have a valid access token
' @return Boolean - True if token is present
''
Public Function Config_HasValidToken() As Boolean
    Config_HasValidToken = (m_SessionState.AccessToken <> "" Or _
                     (Not m_SessionState.IsMultitenant And m_SessionState.SingleTenantToken <> ""))
End Function

''
' Get the authorization header value
' @return String - Authorization header (e.g., "Bearer xyz...")
''
Public Function Config_GetAuthorizationHeader() As String
    If m_SessionState.IsMultitenant Then
        Config_GetAuthorizationHeader = m_SessionState.TokenType & " " & m_SessionState.AccessToken
    Else
        Config_GetAuthorizationHeader = "Basic " & m_SessionState.SingleTenantToken
    End If
End Function

''
' Build the full API URL
' @param apiPath - API path (e.g., "/M3/m3api-rest/v2/execute")
' @param endpoint - Specific endpoint
' @return String - Full URL
''
Public Function Config_BuildApiUrl(apiPath As String, endpoint As String) As String
    Dim url As String
    
    url = m_SessionState.mainUrl & apiPath
    
    If endpoint <> "" Then
        If Right(url, 1) <> "/" And Left(endpoint, 1) <> "/" Then
            url = url & "/"
        End If
        url = url & endpoint
    End If
    
    Config_BuildApiUrl = url
End Function

''
' Process command entered in a cell (like "settings", "clear cache", etc.)
' @param command - The command string
''
Public Sub Config_ProcessCommand(command As String)
    Dim cmd As String
    Dim startIndex As Long
    Dim numericValue As Long
    
    cmd = LCase(Trim(command))
    
    Select Case cmd
        Case "settings"
            ShowSettingsSheet
            
        Case "defaults"
            Config_ResetSettingsToDefaults
            ShowSettingsSheet
            
        Case "environments"
            ToggleSheetVisibility "Environments"
            
        Case "master"
            ToggleSheetVisibility "Master"
            
        Case "versions"
            ToggleSheetVisibility "versions"
            
        Case "cache"
            Cache_DisplayCache
            
        Case "clear cache", "cc"
            Cache_ClearCache
            
        Case "apis"
            ToggleSheetVisibility "AvailableMIs"
            
        Case "transactions"
            ToggleSheetVisibility "Transactions"
            
        Case "unhide"
            ThisWorkbook.Sheets("Environments").Visible = True
            ThisWorkbook.Sheets("AvailableMIs").Visible = True
            
        Case "hide"
            HideSystemSheets
            
        Case "help"
            ShowHelpSheet
            
        Case "clear", "clr"
            ClearActiveSheet
            
        Case "clearstatus", "clrsts"
            ClearStatusColumn
            
        Case Else
            ' Check for parameterized commands
            If Left(cmd, 7) = "maxrecs" Then
                startIndex = InStr(cmd, "=")
                If startIndex > 0 Then
                    numericValue = SafeLong(Mid(cmd, startIndex + 1))
                    m_ApiSettings.MaxRecords = numericValue
                End If
                
            ElseIf Left(cmd, 7) = "maxbulk" Then
                startIndex = InStr(cmd, "=")
                If startIndex > 0 Then
                    numericValue = SafeLong(Mid(cmd, startIndex + 1))
                    m_ApiSettings.maxbulk = CInt(numericValue)
                End If
                
            ElseIf Left(cmd, 7) = "refresh" Then
                startIndex = InStr(cmd, "=")
                If startIndex > 0 Then
                    numericValue = SafeLong(Mid(cmd, startIndex + 1))
                    m_ApiSettings.refreshSeconds = CInt(numericValue)
                End If
            End If
    End Select
End Sub

''
' Toggle visibility of a sheet
''
Private Sub ToggleSheetVisibility(sheetName As String)
    On Error Resume Next
    With ThisWorkbook.Sheets(sheetName)
        .Visible = Not (.Visible = xlSheetVisible)
    End With
    On Error GoTo 0
End Sub

''
' Hide all system sheets
''
Private Sub HideSystemSheets()
    Dim systemSheets As Variant
    Dim i As Long
    
    systemSheets = Array("Master", "Log", "Cache", "Settings", "Environments", _
                         "AvailableMIs", "Transactions", "Help", "Versions")
    
    On Error Resume Next
    For i = LBound(systemSheets) To UBound(systemSheets)
        ThisWorkbook.Sheets(systemSheets(i)).Visible = False
    Next i
    On Error GoTo 0
End Sub

''
' Show the Settings sheet
''
Private Sub ShowSettingsSheet()
    Dim ws As Worksheet
    
    On Error GoTo ErrorHandler
    
    Set ws = ThisWorkbook.Sheets("Settings")
    
    With ws
        .Visible = True
        .Range("maxrecs").value = m_ApiSettings.MaxRecords
        .Range("maxbulk").value = m_ApiSettings.maxbulk
        .Range("refreshSeconds").value = m_ApiSettings.refreshSeconds
        .Range("formatting").value = m_ApiSettings.formatting
        .Range("righttrim").value = m_ApiSettings.righttrim
        .Range("splitChar").value = m_ApiSettings.splitChar
        .Range("maxtime").value = m_ApiSettings.MaxTimeout
        .Range("conoDivi").value = m_ApiSettings.conoDivi
        .Range("naming").value = m_ApiSettings.sheetNaming
    End With

    ws.Activate
    Exit Sub
    
ErrorHandler:
    MsgBox "Error showing settings: " & Err.description, vbExclamation
End Sub

''
' Show the Help sheet
''
Private Sub ShowHelpSheet()
    On Error Resume Next
    ThisWorkbook.Sheets("Help").Visible = True
    ThisWorkbook.Sheets("Help").Activate
    On Error GoTo 0
End Sub

''
' Clear the active sheet data area
''
Private Sub ClearActiveSheet()
    On Error Resume Next
    With ActiveSheet
        .Rows("9:" & .Rows.count).ClearContents
        .Columns("A:" & .Columns.count).ClearContents
    End With
    On Error GoTo 0
End Sub

''
' Clear just the status column (A)
''
Private Sub ClearStatusColumn()
    On Error Resume Next
    ActiveSheet.Range("A9:A" & ActiveSheet.Rows.count).ClearContents
    On Error GoTo 0
End Sub


