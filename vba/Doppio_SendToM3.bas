Attribute VB_Name = "Doppio_SendToM3"
Sub ExportToEVS100()
    Dim wbSource As Workbook
    Dim wbNew As Workbook
    Dim wsSource As Worksheet
    Dim wsNew As Worksheet
    Dim wsControl As Worksheet
    Dim sheetName As String
    Dim wbName As String
    Dim lastCol As Long
    Dim i As Long, r As Long
    Dim rawVal As String, lines() As String
    Dim label As String, ref As String
    Dim colLetter As String, rowNum As String
    Dim filePath As Variant
    Dim firstAPIName As String
    
    On Error GoTo ErrorHandler
    
    Tenant_Token
    ClearStatus
    KillPleaseWait
    UI_HideSystemSheets
    
    Set wbSource = ThisWorkbook
    
    ' Create new workbook with Control sheet
    Set wbNew = Workbooks.Add(xlWBATWorksheet)
    Set wsControl = wbNew.Sheets(1)
    wsControl.name = "Control"
    wsControl.Range("A1").value = "Worksheet"
    wsControl.Range("B1").value = "Description"
    wsControl.Range("C1").value = "Data"
    r = 2
    firstAPIName = ""
    
    ' Loop through all visible sheets in source workbook
    For Each wsSource In wbSource.Sheets
        If wsSource.Visible <> xlSheetVisible Then GoTo NextSheet
        
        ' Capture first visible sheet's API name for file naming
        If firstAPIName = "" Then
            firstAPIName = wsSource.Range("A2").value
        End If
        
        ' Generate sheet name from A2 and G4
        sheetName = "API_" & wsSource.Range("A2").value & "_" & wsSource.Range("G4").value
        
        ' Skip if sheet already exists in target workbook
        If SheetExistsIn(wbNew, sheetName) Then GoTo NextSheet
        
        ' Create and populate the new sheet
        Set wsNew = wbNew.Sheets.Add(After:=wbNew.Sheets(wbNew.Sheets.count))
        wsNew.name = sheetName
        
        ' Ensure no freeze panes
        wsNew.Activate
        If wsNew.Parent.Windows(1).FreezePanes Then
            wsNew.Parent.Windows(1).FreezePanes = False
        End If
        
        ' Find the last column with data in Row 8
        lastCol = wsSource.Cells(8, wsSource.Columns.count).End(xlToLeft).Column
        
        ' Row 1: field names from source Row 8
        For i = 1 To lastCol
            wsNew.Cells(1, i).value = wsSource.Cells(8, i).value
        Next i
        
        ' Row 2: descriptions from source Row 7 (with ref formatting)
        For i = 1 To lastCol
            wsNew.Cells(2, i).value = FormatDescriptionCell(wsSource.Cells(7, i).value)
        Next i
        
        ' Row 3: "no" in A3, "yes" elsewhere
        wsNew.Cells(3, 1).value = "no"
        For i = 2 To lastCol
            wsNew.Cells(3, i).value = "yes"
        Next i
        
        ' Override A1 with MESSAGE
        wsNew.Cells(1, 1).value = "MESSAGE"
        
        ' Copy data from Row 9 onwards
        wsSource.Range(wsSource.Cells(9, 2), wsSource.Cells(wsSource.UsedRange.Rows.count, lastCol)).Copy _
            Destination:=wsNew.Cells(4, 2)
        
        ' Add entry to Control sheet
        wsControl.Cells(r, 1).value = sheetName
        wsControl.Cells(r, 2).value = wsSource.Range("D2").value
        wsControl.Cells(r, 3).value = "x"
        r = r + 1
        
        wsNew.Columns.AutoFit
        
NextSheet:
    Next wsSource
    
    wsControl.Columns.AutoFit
    
    ' Remove the default empty first sheet if still present
    On Error Resume Next
    If wbNew.Sheets(1).name = "Sheet1" Then wbNew.Sheets(1).Delete
    On Error GoTo ErrorHandler
    
    ' Prompt user to save
    wbName = "API_" & firstAPIName
    #If Mac Then
        filePath = Application.GetSaveAsFilename(InitialFileName:=wbName)
    #Else
        filePath = Application.GetSaveAsFilename(InitialFileName:=wbName, FileFilter:="Excel Files (*.xlsx), *.xlsx")
    #End If
    
    If CStr(filePath) = "False" Then
        MsgBox "Save canceled. Workbook was not saved.", vbExclamation
        Exit Sub
    End If
    
    Application.DisplayAlerts = False
    wbNew.SaveAs fileName:=filePath, FileFormat:=xlOpenXMLWorkbook
    Application.DisplayAlerts = True
    
    ' === Upload & Process ===
    If MsgBox("Do you want to send this file to M3?", vbYesNo + vbQuestion, "Upload to M3") = vbYes Then
        wbNew.Close SaveChanges:=True
        
        If UploadFileToM3(CStr(filePath)) Then
            If MsgBox("Upload successful. Do you want to process this file in M3 (EVS100MI.ImportFile)?", _
                       vbYesNo + vbQuestion, "Process in M3") = vbYes Then
                ProcessFileInM3 Dir(CStr(filePath))
            End If
        End If
    End If
    
    Exit Sub

ErrorHandler:
    Debug.Print "ExportToEVS100: ERROR - " & Err.description
    MsgBox "Error in ExportToEVS100: " & Err.description, vbCritical
End Sub


' =============================================================================
' Upload file to M3 via File Management REST API (PUT)
' =============================================================================
Private Function UploadFileToM3(filePath As String) As Boolean
    Dim config As httpConfig
    Dim response As httpResponse
    Dim uploadUrl As String
    Dim authHeader As String
    
    uploadUrl = m_s_MainUrl & "/M3/foundation-rest/file-management/v1/file/FileImport/" & Dir(filePath)
    authHeader = m_s_TokenType & " " & m_s_AccessToken
    
    config = BuildFileUploadConfig(uploadUrl, authHeader, filePath, 120)
    response = ExecuteRequest(config)
    
    If response.success Or response.statusCode = 200 Then
        Debug.Print "UploadFileToM3: Success (HTTP " & response.statusCode & ")"
        UploadFileToM3 = True
    Else
        Debug.Print "UploadFileToM3: Failed (HTTP " & response.statusCode & ") - " & response.errorMessage
        MsgBox "Upload failed (HTTP " & response.statusCode & "): " & response.errorMessage, vbExclamation
        UploadFileToM3 = False
    End If
End Function

' =============================================================================
' Process uploaded file in M3 via EVS100MI.ImportFile (POST)
' =============================================================================
Private Sub ProcessFileInM3(fileName As String)
    Dim response As httpResponse
    Dim body As String
    Dim apiUrl As String

    ' Ensure token is valid before calling the API
    EnsureAuthenticated

    ' EVS100MI.ImportFile only takes FNAM as input � no output fields to select
    body = "{""program"":""EVS100MI""," & _
           """transactions"":[{""transaction"":""ImportFile""," & _
           """record"":{""FNAM"":""" & fileName & """}}]}"

    apiUrl = m_s_MiPath
    apiUrl = apiUrl & "?extendedresult=true"
    
    If m_s_M3user <> "" Then
        apiUrl = apiUrl & "&m3user=" & m_s_M3user
    End If
    If m_s_Company <> "" Then apiUrl = apiUrl & "&cono=" & m_s_Company
    If m_s_Division <> "" Then apiUrl = apiUrl & "&divi=" & m_s_Division

    response = ExecuteApiPost(m_s_MainUrl, apiUrl, body)

    If response.success Then
        Debug.Print "ProcessFileInM3: Success"
        MsgBox "File processed successfully.", vbInformation, "Process in M3"
    Else
        Debug.Print "ProcessFileInM3: Failed (HTTP " & response.statusCode & ") - " & response.errorMessage
        MsgBox "Processing failed (HTTP " & response.statusCode & "): " & response.errorMessage, vbExclamation
    End If
End Sub

' =============================================================================
' Helpers
' =============================================================================
Private Function SheetExistsIn(wb As Workbook, name As String) As Boolean
    Dim Sh As Worksheet
    For Each Sh In wb.Sheets
        If LCase(Sh.name) = LCase(name) Then
            SheetExistsIn = True
            Exit Function
        End If
    Next Sh
    SheetExistsIn = False
End Function

Private Function FormatDescriptionCell(rawVal As String) As String
    Dim lines() As String
    Dim label As String, ref As String
    
    rawVal = Replace(rawVal, vbCrLf, vbLf)
    rawVal = Replace(rawVal, vbCr, vbLf)
    lines = Split(rawVal, vbLf)
    
    If UBound(lines) >= 1 Then
        label = Trim(lines(0))
        ref = Trim(lines(1))
        If Len(ref) >= 2 Then
            FormatDescriptionCell = label & " (" & Left(ref, 1) & ":" & Mid(ref, 2) & ")"
        Else
            FormatDescriptionCell = label
        End If
    Else
        FormatDescriptionCell = Trim(rawVal)
    End If
End Function

