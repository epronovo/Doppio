Attribute VB_Name = "Doppio_Logos"
Sub InsertLogoCrossPlatform(EnvironmentName As String, wsTarget As Worksheet)
    Dim wsLogos As Worksheet
    Dim logoName As String
    Dim pastedShape As Shape
    Dim shp As Shape ' Added declaration so the cleanup loop works
    
    ' =========================================================
    ' EXTRACT LOGO PREFIX (Up to first space or underscore)
    ' =========================================================
    Dim envPrefix As String
    envPrefix = EnvironmentName
    If InStr(envPrefix, " ") > 0 Then
        envPrefix = Split(envPrefix, " ")(0)
    End If
    If InStr(envPrefix, "_") > 0 Then
        envPrefix = Split(envPrefix, "_")(0)
    End If
    
    ' Extract the last 3 characters for the type check (TRN, TST, DEM)
    Dim envSuffix As String
    envSuffix = UCase(Right(Trim(EnvironmentName), 3))
    
    ' =====================================================================
    ' 0. CLEANUP: REMOVE EXISTING LOGO
    ' =====================================================================
    ' Loop through all shapes on the sheet. If it finds one named "MyCompanyLogo", delete it.
    For Each shp In wsTarget.Shapes
        If shp.name = "MyCompanyLogo" Or shp.name = "EnvSuffix" Then
            shp.Delete
        End If
    Next shp
    
    ' =====================================================================
    ' 1. DETERMINE WHICH LOGO TO GRAB
    ' =====================================================================
    logoName = envPrefix
    
    ' =====================================================================
    ' 2. LOCATE THE HIDDEN SHEET AND COPY THE SHAPE
    ' =====================================================================
    On Error Resume Next
    Set wsLogos = ThisWorkbook.Sheets("Logos")
    
    If wsLogos Is Nothing Then
        ' Silently exit if the Logos sheet doesn't exist
        Err.Clear
        On Error GoTo 0
        Exit Sub
    End If
    
    ' First attempt: Copy the specific environment logo
    wsLogos.Shapes(logoName).Copy
    
    ' If it fails (logo doesn't exist), try the fallback
    If Err.Number <> 0 Then
        Err.Clear ' Clear the first error
        
        ' Fallback to DOPPIO logo
        logoName = "DOPPIO"
        wsLogos.Shapes(logoName).Copy
        
        ' If DOPPIO ALSO doesn't exist, then finally exit
        If Err.Number <> 0 Then
            Err.Clear
            On Error GoTo 0
            Exit Sub
        End If
    End If
    
    On Error GoTo 0 ' Turn normal error reporting back on
    
    ' =====================================================================
    ' 3. PASTE THE LOGO AND CENTER IT IN A1
    ' =====================================================================
    wsTarget.Activate
    wsTarget.Range("A1").Select
    wsTarget.Paste
    
    Set pastedShape = Selection.ShapeRange(1)
    pastedShape.name = "MyCompanyLogo"
    pastedShape.Left = wsTarget.Range("A1").Left + _
                      (wsTarget.Range("A1").Width - pastedShape.Width) / 2
    pastedShape.Top = wsTarget.Range("A1").Top + _
                     (wsTarget.Range("A1").Height - pastedShape.Height) / 2
    Application.CutCopyMode = False
    
    If envSuffix = "TRN" Or envSuffix = "TST" Or envSuffix = "DEM" Then
        wsTarget.Range("I1").Select
        wsLogos.Shapes(envSuffix).Copy
        wsTarget.Paste
        
        Set pastedShape = Selection.ShapeRange(1)
        pastedShape.name = "EnvSuffix"
        pastedShape.Left = wsTarget.Range("H2").Left + _
                          (wsTarget.Range("H2").Width - pastedShape.Width) / 2
        pastedShape.Top = wsTarget.Range("H2").Top + _
                         (wsTarget.Range("H2").Height - pastedShape.Height) / 2
        Application.CutCopyMode = False
    End If
    
    wsTarget.Range("A7").Select
    
End Sub


