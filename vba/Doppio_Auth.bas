Attribute VB_Name = "Doppio_Auth"
''
' Doppio Authentication Module
' Handles OAuth token management, refresh, and authorization
'
' @module Doppio_Auth
' @author Doppio Group - eric@doppiogroup.com
' @version 2.0
''
Option Explicit

' =============================================================================
' MODULE-LEVEL STATE
' =============================================================================

' Tracks when the current access token expires.
' Initialises to epoch (0) so any cached token loaded from the sheet is treated
' as expired until a fresh token has been obtained this session.
Private m_TokenExpiry As Date

' =============================================================================
' PRIVATE HELPERS
' =============================================================================

''
' Returns True when the in-memory token has passed its expiry timestamp.
' An uninitialised expiry (epoch) is always considered expired.
''
Private Function IsTokenExpired() As Boolean
    IsTokenExpired = (Now >= m_TokenExpiry)
End Function

''
' Set the token expiry from the expires_in value returned by the auth server.
' @param expiresInSeconds - Lifetime of the token in seconds (0 = use default)
''
Private Sub SetTokenExpiry(expiresInSeconds As Long)
    Dim ttl As Long
    ttl = expiresInSeconds
    If ttl <= 0 Then ttl = 3300  ' Default: 55 minutes (conservative for 1-hour tokens)
    m_TokenExpiry = Now + (ttl / 86400#)  ' Convert seconds to fractional days
End Sub

' =============================================================================
' PUBLIC API
' =============================================================================

''
' Get or refresh the access token for the current environment
' This is the main entry point for authentication
' @return Boolean - True if we have a valid token
''
Public Function EnsureAuthenticated() As Boolean
    ' Check if we have configuration
    If Not Config_IsConfigValid() Then
        EnsureAuthenticated = False
        Exit Function
    End If
    
    ' Check if we already have a valid, unexpired token
    If Config_HasValidToken() Then
        If Config_IsMultitenant Then
            ' Multi-tenant: honour the expiry clock
            If Not IsTokenExpired() Then
                EnsureAuthenticated = True
                Exit Function
            End If
            ' Token string is present but expired – clear it so we fall through
            ' to re-authentication below
            Config_AccessToken = ""
        Else
            ' Single-tenant basic auth does not expire
            EnsureAuthenticated = True
            Exit Function
        End If
    End If
    
    ' Try to get a new token
    If Config_IsMultitenant Then
        EnsureAuthenticated = GetServiceAccountToken()
    Else
        ' Single-tenant uses basic auth - already configured
        EnsureAuthenticated = True
    End If
End Function

''
' Get token using service account credentials
' @return Boolean - True if successful
''
Public Function GetServiceAccountToken() As Boolean
    Dim tenant As TenantConfig
    Dim tokenUrl As String
    Dim response As httpResponse
    Dim config As httpConfig
    Dim body As String
    Dim json As Object
    
    On Error GoTo ErrorHandler
    
    tenant = Config_TenantConfig
    
    ' Build token URL
    tokenUrl = tenant.SsoUrl & tenant.tokenEndpoint
    
    Debug.Print "GetServiceAccountToken: URL = " & tokenUrl
    Debug.Print "GetServiceAccountToken: ClientId = " & Left(tenant.clientId, 10) & "..."
    Debug.Print "GetServiceAccountToken: Has SAAK = " & (tenant.ServiceAccountKey <> "")
    
    ' Build request body (URL-encoded form data)
    body = "client_id=" & Core_UrlEncode(tenant.clientId)
    body = body & "&client_secret=" & Core_UrlEncode(tenant.clientSecret)
    body = body & "&grant_type=password"
    body = body & "&username=" & Core_UrlEncode(tenant.ServiceAccountKey)
    body = body & "&password=" & Core_UrlEncode(tenant.ServiceAccountSecret)
    
    ' Configure the request
    With config
        .url = tokenUrl
        .method = HttpMethod_POST
        .contentType = "application/x-www-form-urlencoded"
        .AcceptType = "application/json"
        .authHeader = ""  ' No auth for token endpoint
        .body = body
        .timeoutSeconds = 30
        .WriteToFile = True
    End With
    
    ' Execute the request
    response = ExecuteRequest(config)
    
    Debug.Print "GetServiceAccountToken: HTTP Status = " & response.statusCode
    Debug.Print "GetServiceAccountToken: Success = " & response.success
    Debug.Print "GetServiceAccountToken: Response length = " & Len(response.body)
    
    If Not response.success Then
        Debug.Print "GetServiceAccountToken: Error = " & response.errorMessage
        Core_LogError CreateError("GetServiceAccountToken", "HTTP request failed: " & response.errorMessage)
        GetServiceAccountToken = False
        Exit Function
    End If
    
    ' Parse the token response
    Set json = ParseJson(response.body)
    
    ' Extract tokens
    Config_AccessToken = SafeStr(json.item("access_token"))
    Config_RefreshToken = SafeStr(json.item("refresh_token"))
    Config_TokenType = SafeStr(json.item("token_type"))

    If Config_TokenType = "" Then
        Config_TokenType = "Bearer"
    End If

    ' Record when this token expires so EnsureAuthenticated can auto-refresh
    Dim expiresIn As Long
    expiresIn = 0
    On Error Resume Next
    expiresIn = CLng(json.item("expires_in"))
    On Error GoTo 0
    SetTokenExpiry expiresIn

    GetServiceAccountToken = (Config_AccessToken <> "")

    Debug.Print "GetServiceAccountToken: Got token = " & GetServiceAccountToken
    Debug.Print "GetServiceAccountToken: Token expires at " & m_TokenExpiry
    
    Exit Function
    
ErrorHandler:
    Debug.Print "GetServiceAccountToken: ERROR - " & Err.description
    Core_LogError CreateError("GetServiceAccountToken", Err.description)
    GetServiceAccountToken = False
End Function

''
' Refresh the access token using the refresh token
' @return Boolean - True if successful
''
Public Function RefreshAccessToken() As Boolean
    Dim tenant As TenantConfig
    Dim tokenUrl As String
    Dim response As httpResponse
    Dim config As httpConfig
    Dim body As String
    Dim basicAuth As String
    Dim json As Object
    
    On Error GoTo ErrorHandler
    
    ' Check if we have a refresh token
    If Config_RefreshToken = "" Then
        ' No refresh token, try full authentication
        RefreshAccessToken = GetServiceAccountToken()
        Exit Function
    End If
    
    tenant = Config_TenantConfig
    
    ' Build token URL
    tokenUrl = tenant.SsoUrl & tenant.tokenEndpoint
    
    ' Build Basic auth header
    basicAuth = "Basic " & Base64Encode(tenant.clientId & ":" & tenant.clientSecret)
    
    ' Build request body
    body = "client_id=" & Core_UrlEncode(tenant.clientId)
    body = body & "&client_secret=" & Core_UrlEncode(tenant.clientSecret)
    body = body & "&grant_type=refresh_token"
    body = body & "&refresh_token=" & Core_UrlEncode(Config_RefreshToken)
    
    ' Configure the request
    With config
        .url = tokenUrl
        .method = HttpMethod_POST
        .contentType = "application/x-www-form-urlencoded"
        .AcceptType = "application/json"
        .authHeader = basicAuth
        .body = body
        .timeoutSeconds = 30
        .WriteToFile = True
    End With
    
    ' Execute the request
    response = ExecuteRequest(config)
    
    If Not response.success Then
        ' Refresh failed, try full authentication
        Core_LogError CreateError("RefreshAccessToken", "Refresh failed, trying full auth")
        RefreshAccessToken = GetServiceAccountToken()
        Exit Function
    End If
    
    ' Parse the token response
    Set json = ParseJson(response.body)
    
    ' Extract tokens
    Config_AccessToken = SafeStr(json.item("access_token"))
    Config_TokenType = SafeStr(json.item("token_type"))

    ' Update refresh token if a new one was provided
    Dim newRefresh As String
    newRefresh = SafeStr(json.item("refresh_token"))
    If newRefresh <> "" Then
        Config_RefreshToken = newRefresh
    End If

    If Config_TokenType = "" Then
        Config_TokenType = "Bearer"
    End If

    ' Record expiry so future calls to EnsureAuthenticated can auto-refresh
    Dim expiresIn As Long
    expiresIn = 0
    On Error Resume Next
    expiresIn = CLng(json.item("expires_in"))
    On Error GoTo 0
    SetTokenExpiry expiresIn

    RefreshAccessToken = (Config_AccessToken <> "")
    
    Exit Function
    
ErrorHandler:
    Core_LogError CreateError("RefreshAccessToken", "")
    ' Try full authentication as fallback
    RefreshAccessToken = GetServiceAccountToken()
End Function

''
' Clear all authentication tokens (logout)
''
Public Sub ClearAuthentication()
    Config_AccessToken = ""
    Config_RefreshToken = ""
    Config_TokenType = ""
    m_TokenExpiry = 0  ' Reset expiry so the next EnsureAuthenticated forces a fresh login
End Sub

''
' Handle unauthorized response - attempt to refresh and retry
' @return Boolean - True if successfully refreshed
''
Public Function HandleUnauthorized() As Boolean
    #If DEBUG_MODE Then
        Debug.Print "HandleUnauthorized: Handling unauthorized response - attempting token refresh"
    #End If
    
    ' Clear current token
    Config_AccessToken = ""
    
    ' Try to refresh
    HandleUnauthorized = RefreshAccessToken()
End Function

' =============================================================================
' TENANT TOKEN MANAGEMENT
' =============================================================================

''
' Full tenant token flow - load config and get token
' Called when environment changes or on initial load
' @return Boolean - True if successful
''
Public Function InitializeTenantToken() As Boolean
    Dim envName As String
    
    On Error GoTo ErrorHandler
    
    ' Get selected environment from active sheet
    envName = GetSelectedEnvironment()
    
    If envName = "" Then
        ' Clear everything if no environment selected
        ClearAuthentication
        UI_ClearEnvironmentFields
        InitializeTenantToken = False
        Exit Function
    End If
    
    ' Update UI colors based on environment
    UI_UpdateEnvironmentColors
    
    ' Load the tenant configuration
    If Not Config_LoadTenantConfig(envName) Then
        MsgBox "Failed to load configuration for environment: " & envName, vbExclamation
        InitializeTenantToken = False
        Exit Function
    End If
    
    ' For multi-tenant, get a token
    If Config_IsMultitenant Then
        ' Check if we have a cached token
        If Not TryGetCachedToken(envName) Then
            ' Get a new token
            If Not GetServiceAccountToken() Then
                MsgBox "Failed to authenticate with environment: " & envName, vbExclamation
                InitializeTenantToken = False
                Exit Function
            End If
            
            ' Cache the token
            CacheToken envName
        End If
    End If
    
    InitializeTenantToken = True
    Exit Function
    
ErrorHandler:
    Core_LogError CreateError("InitializeTenantToken", envName)
    InitializeTenantToken = False
End Function

''
' Get the selected environment from the active sheet
' @return String - Environment name or empty
''
Private Function GetSelectedEnvironment() As String
    On Error Resume Next
    GetSelectedEnvironment = ActiveSheet.Range("Environment").value
    On Error GoTo 0
End Function

''
' Try to get a cached token for the environment
' @param envName - Environment name
' @return Boolean - True if valid cached token found
''
Private Function TryGetCachedToken(envName As String) As Boolean
    Dim env As Environment
    Dim manager As EnvironmentManager
    
    On Error GoTo ErrorHandler
    
    Set manager = Config_EnvironmentManager
    
    If Not manager.HasEnvironment(envName) Then
        TryGetCachedToken = False
        Exit Function
    End If
    
    Set env = manager.GetEnvironment(envName)
    
    ' Check if we have a token and it's not empty
    If env.token <> "" Then
        Config_AccessToken = env.token
        Config_TokenType = "Bearer"
        TryGetCachedToken = True
    Else
        TryGetCachedToken = False
    End If
    
    Exit Function
    
ErrorHandler:
    TryGetCachedToken = False
End Function

''
' Cache the current token for the environment
' @param envName - Environment name
''
Private Sub CacheToken(envName As String)
    Dim env As Environment
    Dim manager As EnvironmentManager

    On Error Resume Next

    Set manager = Config_EnvironmentManager
    Set env = manager.GetEnvironment(envName)

    If Not env Is Nothing Then
        ' Update the environment with the new token
        manager.AddEnvironment env.name, env.tenant, env.Details, _
            Config_AccessToken, env.url, env.User, env.company, env.division
    End If

    ' Persist token to Environments sheet column C so it survives session restarts
    Dim ws As Worksheet
    Dim envRange As Range
    Dim tokenCell As Range

    Set ws = ThisWorkbook.Sheets("Environments")
    If Not ws Is Nothing Then
        Set envRange = ws.Range("A:A")
        Set tokenCell = envRange.Find(What:=envName, LookIn:=xlValues, LookAt:=xlWhole)
        If Not tokenCell Is Nothing Then
            tokenCell.Offset(0, 2).value = Config_AccessToken
        End If
    End If

    On Error GoTo 0
End Sub

' =============================================================================
' VALIDATION
' =============================================================================

''
' Validate that we're ready to make API calls
' @return Boolean - True if ready
''
Public Function ValidateAuthState() As Boolean
    ' Check environment is selected
    If Config_SelectedEnvironment = "" Then
        ValidateAuthState = False
        Exit Function
    End If
    
    ' Check we have a token (or basic auth for single-tenant)
    If Config_IsMultitenant Then
        ValidateAuthState = (Config_AccessToken <> "")
    Else
        ValidateAuthState = (Config_SessionState.SingleTenantToken <> "")
    End If
End Function

''
' Validate environment selection and ensure we have a token
' Shows messages to user if validation fails
' @return Boolean - True if valid and ready
''
Public Function Auth_ValidateSelectedEnvironment() As Boolean
    Dim envName As String
    
    envName = GetSelectedEnvironment()
    
    If envName = "" Or envName = "Access requested" Then
        MsgBox "Please select a valid environment.", vbExclamation
        Auth_ValidateSelectedEnvironment = False
        Exit Function
    End If
    
    ' Ensure we have authentication
    If Not EnsureAuthenticated() Then
        'MsgBox "Failed to authenticate. Please check your credentials.", vbExclamation
        Auth_ValidateSelectedEnvironment = False
        Exit Function
    End If
    
    Auth_ValidateSelectedEnvironment = True
End Function


