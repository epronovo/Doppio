Attribute VB_Name = "DoppioBridge"
''
' Doppio Bridge Module
' Provides backward compatibility with existing code
' Maps old function signatures to new modular architecture
'
' This module allows gradual migration - you can import all new modules
' and this bridge, then your existing code will continue to work while
' you migrate piece by piece.
'
' @module DoppioBridge
' @author Doppio Group - eric@doppiogroup.com
' @version 2.0
''
Option Explicit

' =============================================================================
' LEGACY GLOBAL VARIABLES
' Note: These are commented out because they conflict with DoppioGroup.bas
' Uncomment them only AFTER removing DoppioGroup.bas from your project
' =============================================================================

' Tenant configuration - delegates to DoppioConfig
' These properties are disabled to avoid conflicts with existing globals in DoppioGroup.bas
' When you fully migrate away from DoppioGroup, you can uncomment these

'Public Property Get ci() As String
'    ci = DoppioConfig.TenantConfig.ClientId
'End Property
'
'Public Property Get cs() As String
'    cs = DoppioConfig.TenantConfig.ClientSecret
'End Property
'
' ... etc (all other properties commented out for now)

' =============================================================================
' LEGACY FUNCTION MAPPINGS
' =============================================================================

''
' Legacy: Initialize tenant information
' Maps to: DoppioConfig.Config_LoadTenantConfig
''
Public Sub Bridge_Tenant_Information()
    Dim envName As String
    
    On Error Resume Next
    envName = ActiveSheet.Range("Environment").value
    On Error GoTo 0
    
    If envName = "" Then
        DoppioUI.UI_ClearEnvironmentFields
    Else
        DoppioUI.UI_UpdateEnvironmentColors
        DoppioConfig.Config_LoadTenantConfig envName
    End If
End Sub

''
' Legacy: Get token for tenant
' Maps to: DoppioAuth.InitializeTenantToken
''
Public Sub Bridge_Tenant_Token()
    DoppioAuth.InitializeTenantToken
End Sub

''
' Legacy: Get new token
' Maps to: DoppioAuth.GetServiceAccountToken
''
Public Sub Bridge_GetToken()
    DoppioAuth.GetServiceAccountToken
End Sub

''
' Legacy: Refresh existing token
' Maps to: DoppioAuth.RefreshAccessToken
''
Public Sub Bridge_RefreshToken()
    DoppioAuth.RefreshAccessToken
End Sub

''
' Legacy: Validate selected environment
' Maps to: DoppioAuth.Auth_ValidateSelectedEnvironment
''
Public Sub Bridge_ValidateSelectedEnvironment()
    DoppioAuth.Auth_ValidateSelectedEnvironment
End Sub

''
' Legacy: Show please wait dialog
' Maps to: DoppioUI.UI_ShowPleaseWait
''
Public Sub Bridge_ShowPleaseWait(message As String)
    DoppioUI.UI_ShowPleaseWait message
End Sub

''
' Legacy: Kill please wait dialog
' Maps to: DoppioUI.UI_KillPleaseWait
''
Public Sub Bridge_KillPleaseWait()
    DoppioUI.UI_KillPleaseWait
End Sub

''
' Legacy: Update environment colors
' Maps to: DoppioUI.UI_UpdateEnvironmentColors
''
Public Sub Bridge_ChangeCellColorBasedOnEnvironment()
    DoppioUI.UI_UpdateEnvironmentColors
End Sub

''
' Legacy: Clear status column
' Maps to: DoppioUI.UI_ClearStatus
''
Public Sub Bridge_ClearStatus()
    DoppioUI.UI_ClearStatus
End Sub

''
' Legacy: Clear fields when no environment selected
' Maps to: DoppioUI.UI_ClearEnvironmentFields
''
Public Sub Bridge_ClearFields()
    DoppioUI.UI_ClearEnvironmentFields
End Sub

''
' Legacy: Auto-fit columns
' Maps to: DoppioUI.UI_AutoFitColumns
''
Public Sub Bridge_AutoFit_ColumnsAndRows(reload As Boolean, mandatory As Boolean)
    DoppioUI.UI_AutoFitColumns ActiveSheet, reload, mandatory
End Sub

''
' Legacy: Display elapsed time
' Maps to: DoppioUI.UI_DisplayElapsedTime
''
Public Sub Bridge_DisplayElapsedTime(startTime As Single, ws As Worksheet)
    DoppioUI.UI_DisplayElapsedTime startTime, ws
End Sub

''
' Legacy: Update version display
' Maps to: DoppioUI.UI_UpdateVersion
''
Public Sub Bridge_UpdateVersion()
    DoppioUI.UI_UpdateVersion
End Sub

''
' Legacy: Prompt user with yes/no
' Maps to: DoppioUI.UI_PromptUser
''
Public Function Bridge_PromptUser(message As String) As Boolean
    Bridge_PromptUser = DoppioUI.UI_PromptUser(message)
End Function

''
' Legacy: Exit program prompt
' Maps to: DoppioUI.UI_PromptExitProgram
''
Public Sub Bridge_ExitProgram()
    DoppioUI.UI_PromptExitProgram
End Sub

''
' Legacy: Log error to sheet
' Maps to: DoppioUI.UI_LogError
''
Public Sub Bridge_LogError(data As String)
    DoppioUI.UI_LogError data
End Sub

''
' Legacy: Base64 encode
' Maps to: DoppioCore.Base64Encode
''
Public Function Bridge_Base64EncodeVBA(text As String) As String
    Bridge_Base64EncodeVBA = DoppioCore.Base64Encode(text)
End Function

''
' Legacy: Base64 decode
' Maps to: DoppioCore.Base64Decode
''
Public Function Bridge_Base64DecodeVBA(encodedText As String) As String
    Bridge_Base64DecodeVBA = DoppioCore.Base64Decode(encodedText)
End Function

''
' Legacy: Check if sheet is visible
' Maps to: DoppioCore.Core_IsSheetVisible
''
Public Function Bridge_IsSheetVisible(sheetName As String) As Boolean
    Bridge_IsSheetVisible = DoppioCore.Core_IsSheetVisible(sheetName)
End Function

''
' Legacy: Read file to string
' Maps to: DoppioCore.Core_ReadFileToString
''
Public Function Bridge_ReadFile(filePath As String) As String
    Bridge_ReadFile = DoppioCore.Core_ReadFileToString(filePath)
End Function

''
' Legacy: Write string to file
' Maps to: DoppioCore.WriteStringToFile
''
Public Sub Bridge_WriteFile(filePath As String, content As String)
    DoppioCore.WriteStringToFile filePath, content
End Sub

''
' Legacy: URL encode
' Maps to: DoppioCore.Core_UrlEncode
''
Public Function Bridge_IDM_URLEncode(text As String) As String
    Bridge_IDM_URLEncode = DoppioCore.Core_UrlEncode(text)
End Function

''
' Legacy: Replace alpha with zero
' Note: Uses original function from DoppioGroup if available
' Maps to: DoppioCore.Core_ReplaceAlphaWithZero
''
Public Function Bridge_ReplaceAlphaWithZero(inputStr As String) As String
    Bridge_ReplaceAlphaWithZero = DoppioCore.Core_ReplaceAlphaWithZero(inputStr)
End Function

' =============================================================================
' CACHE COMPATIBILITY
' =============================================================================

Public Sub Bridge_RecordCache_Initialize()
    DoppioCache.Cache_InitializeCache
End Sub

Public Sub Bridge_RecordCache_Store(cacheKey As String)
    ' In the old code, this stored m_obj_Records
    ' With new architecture, use StoreInCache directly
End Sub

Public Sub Bridge_RecordCache_Retreive(cacheKey As String, ByRef found As Boolean)
    Dim response As apiResponse
    found = DoppioCache.Cache_TryGetFromCache(cacheKey, response)
End Sub

'Public Function Bridge_RecordCache_Find(cacheKey As String) As Long
'    Bridge_RecordCache_Find = DoppioCache.RecordCache_Find(cacheKey)
'End Function

Public Sub Bridge_RecordCache_Reset()
    DoppioCache.Cache_ClearCache
End Sub

Public Sub Bridge_RecordCache_Display()
    DoppioCache.Cache_DisplayCache
End Sub

Public Sub Bridge_RecordCache_Load()
    DoppioCache.Cache_LoadCacheFromSheet
End Sub

' =============================================================================
' SETTINGS COMPATIBILITY
' =============================================================================

''
' Legacy: Copy defaults from Settings sheet
' Maps to: DoppioConfig.Config_LoadSettingsFromSheet
''
Public Sub Bridge_Settings_CopyDefaults()
    DoppioConfig.Config_LoadSettingsFromSheet
End Sub

''
' Legacy: Process command from cell
' Maps to: DoppioConfig.Config_ProcessCommand
''
Public Sub Settings_ProcessCommand(command As String)
    DoppioConfig.Config_ProcessCommand command
End Sub

''
' Legacy: Create new sheet from Master
' Maps to: DoppioUI.UI_CreateNewSheet
''
Public Sub Bridge_Settings_NewSheet()
    DoppioUI.UI_CreateNewSheet
End Sub

''
' Legacy: Check if sheet exists
' Maps to: DoppioCore.Core_SheetExists
''
Public Function Bridge_Settings_SheetExists(sheetName As String) As Boolean
    Bridge_Settings_SheetExists = DoppioCore.Core_SheetExists(sheetName)
End Function

' =============================================================================
' ENVIRONMENT COMPATIBILITY
' =============================================================================

''
' Legacy: Load environments from .ionapi files
' Maps to: DoppioConfig.Config_LoadEnvironmentsFromFiles
''
Public Sub Bridge_Environments_Load()
    DoppioConfig.Config_LoadEnvironmentsFromFiles
End Sub

' =============================================================================
' HTTP/CURL COMPATIBILITY
' =============================================================================

''
' Legacy: Execute AppleScript with retry
' This is now handled internally by DoppioHttp
' Provided here for any direct calls
''
Public Sub Bridge_ExecuteScriptWithRetry(ByRef script As String)
    ' This functionality is now built into DoppioHttp.ExecuteRequest
    ' For backward compatibility, we execute directly
    #If Mac Then
        Dim retryCount As Integer
        Dim success As Boolean
        
        retryCount = 0
        success = False
        
        Do While retryCount < 3 And Not success
            On Error Resume Next
            MacScript (script)
            If Err.Number = 0 Then
                success = True
            Else
                retryCount = retryCount + 1
                Err.Clear
            End If
            On Error GoTo 0
        Loop
    #End If
End Sub

''
' Legacy: Parse and execute curl (Windows)
' Maps to: DoppioHttp.ParseAndExecuteCurl
''
Public Function Bridge_ParseAndExecuteCurl_Regex(script As String) As Long
    Bridge_ParseAndExecuteCurl_Regex = DoppioHttp.ParseAndExecuteCurl(script)
End Function

' =============================================================================
' API CALL COMPATIBILITY
' =============================================================================

''
' This is a simplified bridge for the main apicall function
' For full functionality, migrate to DoppioApi.ExecuteApiCall
''
Public Sub ZZZ_apicall_Bridge(main_url As String, mi_path As String, mi_url As String, _
                          body As String, apiType As String)
    Dim request As ApiRequest
    Dim response As apiResponse
    
    With request
        .mainUrl = main_url
        .apiPath = mi_path
        .endpoint = mi_url
        .body = body
        .apiType = StringToApiType(apiType)
        .UseCache = (body = "")
        .cacheKey = mi_url
    End With
    
    response = DoppioApi.ExecuteApiCall(request)
    
    ' The old code stored results in module-level variables
    ' This bridge doesn't fully replicate that behavior
    ' For full compatibility, continue using the original apicall
    ' or migrate to using ApiResponse directly

    If response.success Then
        ' Store raw JSON result
        Doppio.m_s_CurlResult = response.data
        
        ' Store parsed JSON objects
        If Not response.records Is Nothing Then
            Set Doppio.m_obj_Records = response.records
            Set Doppio.m_obj_Results = response.results
            
            ' Get the first result object
            If response.results.count > 0 Then
                Set Doppio.m_obj_JsonResponse = response.results(1)
            End If
            #If DEBUG_MODE Then
                Debug.Print "DoppioBridge: Loaded " & response.recordCount & " records"
            #End If
        Else
            #If DEBUG_MODE Then
                Debug.Print "DoppioBridge: No records returned"
            #End If
        End If
    Else
        #If DEBUG_MODE Then
            Debug.Print "DoppioBridge: Error: " & response.data
        #End If
    End If

End Sub

Public Sub apicall_Bridge(main_url As String, mi_path As String, mi_url As String, _
                          body As String, apiType As String)
    Dim request As ApiRequest
    Dim response As apiResponse
    Dim url As String
    Dim curlMethod As String
    Dim curlFormat As String
    
    ' Handle API type logic and build URL
    If apiType = "" Then apiType = m_obj_ws.Range("Type").value
    If apiType = "" Then apiType = "API"
    
    Select Case apiType
    Case "API"
        If body <> "" Then
            mi_url = mi_url & "?maxrecs=" & maxrecs & "&extendedresult=true"
            If m_s_M3user <> "" Then mi_url = mi_url & "&m3user=" & m_s_M3user
            If righttrim Then
                mi_url = mi_url & "&righttrim=true"
            Else
                mi_url = mi_url & "&righttrim=false"
            End If
            If m_s_Company <> "" Then mi_url = mi_url & "&cono=" & m_s_Company
            If m_s_Division <> "" Then mi_url = mi_url & "&divi=" & m_s_Division
        End If
    End Select
    
    ' Clean up URL
    mi_url = Replace(mi_url, "?&", "?")
    
    ' Build the API request
    With request
        .mainUrl = main_url
        .apiPath = mi_path
        .endpoint = mi_url
        .body = body
        .apiType = StringToApiType(apiType)
        .UseCache = (body = "")
        .cacheKey = mi_url
    End With
    
    response = DoppioApi.ExecuteApiCall(request)
    
    If response.success And response.data <> "" Then
        ' Store raw JSON result
        Doppio.m_s_CurlResult = response.data
        
        ' Parse and store JSON objects (replicating old behavior)
        Set Doppio.m_obj_JsonResponse = JsonConverter.ParseJson(response.data)
        Set Doppio.m_obj_Results = Doppio.m_obj_JsonResponse.item("results")
        
        If Doppio.m_obj_Results.count > 0 Then
            Set Doppio.m_obj_Records = Doppio.m_obj_Results(1).item("records")
            #If DEBUG_MODE Then
                Debug.Print "apicall_Bridge: Loaded " & Doppio.m_obj_Records.count & " records"
            #End If
        Else
            #If DEBUG_MODE Then
                Debug.Print "apicall_Bridge: No results returned"
            #End If
        End If
    Else
        #If DEBUG_MODE Then
            Debug.Print "apicall_Bridge: Error or no data: " & response.errorMessage
        #End If
    End If
End Sub

