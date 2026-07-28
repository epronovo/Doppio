VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "EnvironmentManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Class Module: EnvironmentManager
Option Explicit

Private environmentDetails As Collection

Private Sub Class_Initialize()
    Set environmentDetails = New Collection
End Sub

Public Sub AddEnvironment(tenant As String, name As String, Details As String, token As String, url As String, User As String, company As String, division As String)
    Dim detail As Environment
    ' Check if environment already exists
    On Error Resume Next
    Set detail = environmentDetails(tenant)
    On Error GoTo 0
    If detail Is Nothing Then
        ' Create new if it doesn't exist
        Set detail = New Environment
        detail.tenant = tenant
        environmentDetails.Add detail, key:=tenant
    End If
    ' Update values
    detail.name = name
    detail.Details = Details
    detail.token = token
    detail.url = url
    detail.User = User
    detail.company = company
    detail.division = division
End Sub

Public Function GetEnvironment(key As String) As Environment
    Dim detail As Environment
    On Error Resume Next
    Set detail = environmentDetails(key)
    On Error GoTo 0
    If detail Is Nothing Then
        'MsgBox "Environment '" & key & "' not found.", vbExclamation
    End If
    Set GetEnvironment = detail
End Function

Public Function HasEnvironment(key As String) As Boolean
    Dim dummy As Object
    On Error Resume Next
    Set dummy = environmentDetails(key)
    HasEnvironment = (Err.Number = 0)
    Err.Clear
    On Error GoTo 0
End Function

Public Function LoadEnvironment(key As String, ws As Worksheet) As Boolean
    Dim env As Environment
    Set env = Me.GetEnvironment(key)
    
    If env Is Nothing Then
        LoadEnvironment = False
        Exit Function
    End If
    
    ' Check if token exists
    LoadEnvironment = (env.token <> "")
End Function

Public Sub ResetToken(tenant As String)
    Dim env As Environment
    Set env = Me.GetEnvironment(tenant)
    
    If Not env Is Nothing Then
        env.token = ""
    End If
End Sub

' FIXED: ClearEnvironment now uses the Collection properly
Public Sub ClearEnvironment(envName As String)
    Dim env As Environment
    
    ' Try to get the environment by name (which is the key)
    On Error Resume Next
    Set env = environmentDetails(envName)
    On Error GoTo 0
    
    If Not env Is Nothing Then
        ' Clear the token
        env.token = ""
    End If
End Sub

' NEW: Clear all environments
Public Sub ZZZ_ClearAll()
    Set environmentDetails = New Collection
End Sub

' NEW: Remove a specific environment from the collection
Public Sub RemoveEnvironment(envName As String)
    On Error Resume Next
    environmentDetails.Remove envName
    On Error GoTo 0
End Sub

' NEW: Get count of environments
Public Property Get count() As Long
    count = environmentDetails.count
End Property

