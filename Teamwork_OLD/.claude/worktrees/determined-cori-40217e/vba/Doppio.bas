Attribute VB_Name = "Doppio"
''
' Doppio Group MI Excel Enhancement
' (c) Eric Pronovost - eric@doppiogroup.com
'
' @class Doppio
' @author eric@doppiogroup.com
''
Option Explicit

Public m_b_Multitenant As Boolean
Public fileResult As String
Public m_b_Webservice As Boolean
Public m_l_PrevErrorNumber As Long
Public m_l_Row As Long
Public m_obj_ColumnConditions As New ArrayList
Public m_obj_ColumnDescriptions As New ArrayList
Public m_obj_ColumnDirections As New ArrayList
Public m_obj_ColumnNames As New ArrayList
Public m_obj_ColumnTypes As New ArrayList
Public m_obj_JsonResponse As Object
Public m_obj_Layout As New ArrayList
Public m_obj_Records As Object
Public m_obj_Results As Object
Public m_obj_ws As Worksheet
Public m_s_AccessToken As String
Public m_s_RefreshToken As String
Public m_s_Company As String
Public m_s_CurlResult As String
Public m_s_Division As String
Public m_s_LoadedUrl As String
Public m_s_M3user As String
Public m_s_MainUrl As String
Public m_s_Meta As String
Public m_s_MiPath As String
Public m_s_MiUrl As String
Public m_s_PrevData As String
Public m_s_SelectedEnvironment As String
Public m_s_SoapBody As String
Public m_s_SoapHeader As String
Public m_s_stToken As String
Public m_s_stUrl As String
Public m_s_TokenType As String
Public m_s_WsPath As String
Public m_RecordCache As Collection
Public curlFormat As String
Public curlMethod As String
Public curlBody As String
Public url As String
  
Public ci As String
Public cs As String
Public pu As String
Public ot As String
Public saak As String
Public sask As String
Public iu As String
Public ti As String
Public oa As String
Public ru As String
Public maxrecs As Long
Public maxbulk As Integer
Public refreshSeconds As Integer
Public formatting, righttrim As Boolean
Public splitChar As String
Public originalSheet As Worksheet
Public encodedTenant As String
Public activeEnvironment As String
Public curlCommand As String
Public selectedEnvironment As String
Public maxtime As Integer
Public conoDivi As Long
Public manager As EnvironmentManager
Public MatrixBuilder As MatrixManager
Public pivoted As Boolean
Public currentPid As String

Sub apicall(main_url, mi_path, mi_url, body, ByVal apiType As String)
    Dim curlAuth As String
    Dim curlBody As String
    Dim curlData As String
    Dim curlHeader As String
    Dim curlPrefix As String
    Dim curlUrl As String
    Dim fileNumber As String
    Dim http As Object
    Dim json As Object
    Dim method As String
    Dim mi_result As String
    Dim position As Long
    Dim responseText As String
    Dim retry As Boolean
    Dim retryCount As Integer
    Dim script As String
    Dim tempFilePath As String
    
    On Error GoTo ErrorHandler

    Curl_Build main_url, mi_path, mi_url, body, apiType, script

    #If Mac Then
        retry = True
        retryCount = 0
    
        Do While retry And retryCount < 5
            retry = False
            retryCount = retryCount + 1
      
            ' Save script To a file
            Dim filePath As String
            Dim fileNum As Integer
            filePath = Environ("HOME") & "/curl_input.sh"
            fileNum = FreeFile
            Open filePath For Output As #fileNum
            Print #fileNum, script
            Close #fileNum
    
            ' Execute the AppleScript
            mi_result = ""
            Set m_obj_JsonResponse = Nothing
            Set m_obj_Results = Nothing
            Set m_obj_Records = Nothing
            On Error Resume Next
            
            ExecuteScriptWithRetry (script)
            
            If Err.Number <> 0 Then
                #If DEBUG_MODE Then
                    Debug.Print "apicall_Bridge  Err.Number: " & Err.Number & " [" & url & "]"
                #End If
                If Len(body) > 200 Then
                    #If DEBUG_MODE Then
                        Debug.Print "body: " & Left(body, 200) & " ..."
                    #End If
                End If
                mi_result = "NOK script error"
                ExitProgram
            Else
                'Debug.Print Now & " Load results start"
                Dim result As String
                tempFilePath = Environ("HOME") & "/curl_output.txt"
                fileNumber = FreeFile
                Open tempFilePath For Input As fileNumber
                result = Input$(LOF(fileNumber), fileNumber)
                Close fileNumber
                m_s_CurlResult = result
                'Debug.Print Now & " Load results end"
                If 1 = 2 Then
                    If Len(result) > 200 Then
                        #If DEBUG_MODE Then
                            Debug.Print "result: " & Left(result, 200) & " ..."
                        #End If
                    Else
                        #If DEBUG_MODE Then
                            Debug.Print "result: " & result
                        #End If
                    End If
                End If
                
                If Right(mi_url, Len(".meta")) = ".meta" Or mi_url = "" Then
                    m_s_Meta = result
                End If
    
                ' Check If tranaction failed
                Set json = JsonConverter.ParseJson(result)
                Dim terminationReason As String
                Dim connectionError As String
                terminationReason = json.item("terminationReason")
                connectionError = json.item("error")
                If terminationReason <> "" Or connectionError <> "" Then
                    result = "{""results"":[{""records"":[{""ZZUSID"":""" & terminationReason & connectionError & """}]}]}"
                End If
    
                ' Parse the JSON response
                DoppioUI.UI_ShowPleaseWait "Please Wait... Parse Results"
                
                If result = "" Then
                    Exit Sub
                End If
                
                'Debug.Print result
                Set m_obj_JsonResponse = JsonConverter.ParseJson(result)
                Set m_obj_Results = m_obj_JsonResponse.item("results")(1)
                Set m_obj_Records = m_obj_Results.item("records")
                Set m_obj_Results = m_obj_JsonResponse.item("results")
            End If
      
            ' Check For Unauthorized error And retry If necessary
            If Left(m_s_CurlResult, 9) = "{""error"":" Then
                Dim env As Environment
                Set env = manager.GetEnvironment(m_s_SelectedEnvironment)
                m_s_AccessToken = ""
                manager.AddEnvironment env.Name, env.tenant, env.tenant, m_s_AccessToken, env.url, env.User, env.company, env.division
                retry = True
                If Not apicall_Unauthorized(saak, main_url, mi_path, mi_url, body, apiType, script) Then
                    MsgBox "Error in API processing."
                End If
            End If
        Loop
    
    #Else
        retry = True
        retryCount = 0
        Dim result As String
            
        Do While retry And retryCount < 5
            retry = False
            retryCount = retryCount + 1
    
            ' Create a New HTTP request object
            Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
            http.Open curlMethod, url, False
            http.setRequestHeader "Accept", "application/" & curlFormat
            http.setRequestHeader "Content-Type", "application/" & curlFormat
            http.setRequestHeader "Authorization", m_s_TokenType & " " & m_s_AccessToken
            If body <> "" Then
                http.Send body
            Else
                http.Send
            End If
      
            ' Get the response text
            result = http.responseText
            If result = "" Then result = http.status
    
            ' Check For Unauthorized error And retry If necessary
            If Left(result, 9) = "{""error"":" Then
                position = RecordCache_Find(saak)
                If position > 0 And position <= m_RecordCache.count Then
                    m_RecordCache.Remove position
                End If
                position = RecordCache_Find(saak)
                retry = True

                'apicall_Unauthorized saak, main_url, mi_path, mi_url, body, apitype, script
                If Not apicall_Unauthorized(saak, main_url, mi_path, mi_url, body, apiType, script) Then
                    MsgBox "Error in API processing."
                End If
            Else
                If Right(mi_url, Len(".meta")) = ".meta" Then
                    m_s_Meta = responseText
                End If
        
                ' Check If tranaction failed
                Set json = JsonConverter.ParseJson(responseText)
                Dim terminationReason As String
                terminationReason = json.item("terminationReason")
                If terminationReason <> "" Then
                    result = "{""results"":[{""errorMessage"":""" & terminationReason & """}]}"
                End If
        
                ' Parse the JSON response
                DoppioUI.UI_ShowPleaseWait "Please Wait... Parse Results"
                Set m_obj_JsonResponse = JsonConverter.ParseJson(result)
                Set m_obj_Results = m_obj_JsonResponse.item("results")(1)
                Set m_obj_Records = m_obj_Results.item("records")
                Set m_obj_Results = m_obj_JsonResponse.item("results")
                m_s_CurlResult = result
            End If
        Loop
    #End If
    
    'Debug.Print Now & " ParseBulk start"
    Select Case apiType
    Case "API"
        If body <> "" Then
            Parse_Bulk_Results m_obj_Results
        End If
    Case "IPS"
        If curlMethod = "POST" Then
            Parse_Soap_Results result
        End If
    Case "XtendM3"
        Parse_M3X_Results
    Case Else
    End Select
    'Debug.Print Now & " ParseBulk end"

ErrorHandler:
    If IsSheetVisible("Log") Then
        Dim data As String
        body = Replace(body, "\""", """")
        data = mi_url
        If data = "" Then
            Dim pos As Long
            pos = InStr(1, body, "selectedColumns", vbTextCompare)
            If pos > 0 Then
                data = Left(body, pos - 3) & " ..."
            Else
                data = Left(body, 250)
            End If
        End If
        LogError data
    End If
    Resume Next

    Application.ScreenUpdating = True
    Application.Calculate
    DoEvents
    Application.ScreenUpdating = False
    DoppioUI.UI_KillPleaseWait

End Sub

Function apicall_Unauthorized(saak As String, ByVal main_url As String, ByVal mi_path As String, ByVal mi_url As String, ByVal body As String, ByVal apiType As String, script As String) As Boolean
    Dim temp_main_url As String
    Dim temp_mi_path As String
    Dim temp_mi_url As String
    Dim temp_body As String
    Dim temp_apitype As String
    Dim position As Long
    Dim retry As Boolean
    
    ' Store the current values into temporary variables
    temp_main_url = main_url
    temp_mi_path = mi_path
    temp_mi_url = mi_url
    temp_body = body
    temp_apitype = apiType

    ' Refresh tenant token
    Tenant_Token

    ' Reprocess the API type with the stored temporary values
    Curl_Build temp_main_url, temp_mi_path, temp_mi_url, temp_body, temp_apitype, script

    ' Return success
    apicall_Unauthorized = True
End Function

Sub Curl_Build(ByVal main_url As String, ByVal mi_path As String, ByVal mi_url As String, ByVal body As String, apiType As String, ByRef script As String)
    Dim curlPrefix As String
    Dim curlPostfix As String
    Dim curlHeader As String
    Dim curlAuth As String
    Dim curlData As String
    Dim curlContent As String
    Dim curlDataType As String
    
    ' Initialize URL and format
    url = main_url & mi_path & "/" & mi_url
    curlFormat = "json"
    curlContent = "json; charset=UTF-8"
    curlDataType = "--data-raw"
    'Debug.Print URL
    
    ' Handle API type logic
    If apiType = "" Then apiType = m_obj_ws.Range("Type").value
    If apiType = "" Then apiType = "API"
    Select Case apiType
    Case "API"
        If body = "" Then
            curlMethod = "GET"
        Else
            curlMethod = "POST"
            url = main_url & mi_path & "?maxrecs=" & maxrecs & "&extendedresult=true"
            If m_s_M3user <> "" Then url = url & "&m3user=" & m_s_M3user
            If righttrim Then
                url = url & "&righttrim=true"
            Else
                url = url & "&righttrim=false"
            End If
            If m_s_Company <> "" Then url = url & "&cono=" & m_s_Company
            If m_s_Division <> "" Then url = url & "&divi=" & m_s_Division
        End If
    Case "IDM"
            curlMethod = m_obj_ws.Range("G5").value
    Case "IPS"
        If body = "" Then
            curlMethod = "GET"
        Else
            curlMethod = "POST"
            curlFormat = "xml"
            curlContent = "xml"
        End If
    Case "FileMng"
        curlMethod = "PUT"
        curlContent = "octet-stream"
        curlDataType = "--data-binary"
        curlPostfix = " -w '%{http_code}'"
    Case "XtendM3"
        curlMethod = "PUT"
        curlPostfix = " -w '%{http_code}'"
    Case "FNC"
        curlMethod = "POST"
    End Select

    ' Construct cURL command
    curlPrefix = "curl --request " & curlMethod & " --max-time " & maxtime & " "
    curlHeader = "--header 'accept: application/" & curlFormat & "; charset=UTF-8' --header 'Content-Type: application/" & curlContent & "' "
    curlAuth = "--header 'Authorization: " & m_s_TokenType & " " & m_s_AccessToken & "' "
    curlData = ""
    If body <> "" Then
        curlBody = body
        curlBody = Replace(curlBody, "\", "\\\\")
        curlBody = Replace(curlBody, """", "\""")
        curlBody = Replace(curlBody, "'", "'\\''")
        curlBody = Replace(curlBody, "!!", "\\\""")
        curlData = curlDataType & " '" & curlBody & "'"
    End If

    ' Construct script for macOS
    script = "Do shell script """ & curlPrefix & "--location '" & url & "' " & curlHeader & curlAuth & curlData & curlPostfix & " > ~/curl_output.txt"""
    
    curlData = ""
    If body <> "" Then
        curlBody = body
        curlData = curlDataType & " '" & curlBody & "'"
    End If
    curlCommand = curlPrefix & "--location '" & url & "' " & curlHeader & curlAuth & curlData
    curlCommand = Replace(curlCommand, "?&", "?")
    
End Sub

Sub AutoFit_ColumnsAndRows(reload As Boolean, mandatory As Boolean)
    Dim ws As Worksheet
    Dim colNum As Long
    Dim lastColumn As Long
    Dim i As Long
    Dim index As Long
    Dim value As String
    Dim min_G_Width As Double
    Dim min_I_Width As Double
    Dim formatRange As Range
    Dim valuesArray() As Variant
    
    'Debug.Print Now & " AutoFit_ColumnsAndRows START"
    
    ' Speed up by disabling screen updating, events, And calculations
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    colNum = 2
    m_l_Row = 9
    Set ws = ThisWorkbook.ActiveSheet

    lastColumn = m_obj_ColumnNames.count() + 1
    If lastColumn < ws.Cells(8, ws.columns.count).End(xlToLeft).column Then
        lastColumn = ws.Cells(8, ws.columns.count).End(xlToLeft).column
    End If
    
    If reload Then
        'Debug.Print Now & " reload START"
        ReDim valuesArray(1 To 2, 1 To lastColumn + 1)
        
        If mandatory Then
            ' Only mandatory fields
            For i = 1 To lastColumn
                If (m_obj_ColumnConditions.item(i) = "1" Or m_obj_ColumnConditions.item(i)) Then
                    valuesArray(1, colNum) = m_obj_ColumnDescriptions.item(i)
                    valuesArray(2, colNum) = m_obj_ColumnNames.item(i)
                    colNum = colNum + 1
                End If
            Next i
        Else
            ' Pass 1: mandatory fields first
            For i = 1 To lastColumn
                If (m_obj_ColumnConditions.item(i) = "1" Or m_obj_ColumnConditions.item(i)) Then
                    valuesArray(1, colNum) = m_obj_ColumnDescriptions.item(i)
                    valuesArray(2, colNum) = m_obj_ColumnNames.item(i)
                    colNum = colNum + 1
                End If
            Next i
            
            ' Pass 2: then the rest
            For i = 1 To lastColumn
                If Not (m_obj_ColumnConditions.item(i) = "1" Or m_obj_ColumnConditions.item(i)) Then
                    valuesArray(1, colNum) = m_obj_ColumnDescriptions.item(i)
                    valuesArray(2, colNum) = m_obj_ColumnNames.item(i)
                    colNum = colNum + 1
                End If
            Next i
        End If
        ws.Rows("7:8").ClearContents
        valuesArray = TransformArray(valuesArray, conoDivi)
        lastColumn = UBound(valuesArray, 2)
        
        'Debug.Print Now & " reload END"
    Else
        'Debug.Print Now & " load START"
        On Error Resume Next
        ReDim valuesArray(1 To 2, 1 To lastColumn)
        lastColumn = ws.Cells(8, ws.columns.count).End(xlToLeft).column
        If 1 = 1 Then
            For i = 2 To lastColumn
                value = ws.Cells(8, i).value
                index = m_obj_ColumnNames.IndexOf(value)
                If index > 0 Then
                    valuesArray(1, i) = m_obj_ColumnDescriptions.item(index)
                    valuesArray(2, i) = m_obj_ColumnNames.item(index)
                    If m_obj_ColumnTypes.item(index) <> "A" Then
                        If formatRange Is Nothing Then
                            Set formatRange = ws.columns(colNum)
                        Else
                            Set formatRange = Union(formatRange, ws.columns(colNum))
                        End If
                    End If
                    Select Case m_obj_ColumnDirections.item(index)
                    Case "I"
                        If m_obj_ColumnConditions.item(index) = "1" Or m_obj_ColumnConditions.item(index) Then
                            ws.Cells(m_l_Row - 1, i).Interior.Color = RGB(130, 0, 0)
                        Else
                            ws.Cells(m_l_Row - 1, i).Interior.Color = RGB(64, 64, 64)
                        End If
                    Case Else
                        ws.Cells(m_l_Row - 1, i).Interior.Color = RGB(128, 128, 128)
                    End Select
                    colNum = colNum + 1
                Else
                    valuesArray(1, i) = ""
                    valuesArray(2, i) = value
                End If
            Next i
        End If
 
        ' Set the number format for the collected range in one go
        If formatting Then
            'Debug.Print Now & " format Range START"
            ws.Cells.NumberFormat = "@"
            If Not formatRange Is Nothing Then
                formatRange.NumberFormat = "General"
            End If
            'Debug.Print Now & " format Range END"
        End If
        
        ' Adjust row And column sizes
        If 1 = 1 Then
            'Debug.Print Now & " adjust size START"
            With ws
                .Rows("1:6").AutoFit
                .Rows("1:6").columns.AutoFit
                Rows(1).RowHeight = 60
                Rows(7).RowHeight = 36
                columns(1).ColumnWidth = 38
            
                min_G_Width = .Cells(7, 7).ColumnWidth
                min_I_Width = .Cells(7, 9).ColumnWidth
            
                .Rows(7).WrapText = False
                .Rows(7).columns.AutoFit
                .Rows(7).WrapText = True
                .Rows(7).columns.AutoFit
                    
                '.UsedRange.Rows.AutoFit
            End With
            'Debug.Print Now & " adjust size END"
                
        End If
        'Debug.Print Now & " load END"
    End If

    ' Write the data back to the worksheet
    'Debug.Print Now & " load array START"
    ws.Range(ws.Cells(m_l_Row - 2, 1), ws.Cells(m_l_Row - 1, lastColumn)).value = valuesArray
    'Debug.Print Now & " load array END"


    ' Loop through all columns in Row 7 And autofit, but ensure the width is at least minColumnWidth
    'Debug.Print Now & " min width START"
    For i = 1 To lastColumn
        If ws.Cells(7, i).ColumnWidth < 12 Then
            ws.columns(i).ColumnWidth = 12
        End If
        If i = 7 And ws.Cells(7, i).ColumnWidth < min_G_Width Then
            ws.columns(i).ColumnWidth = min_G_Width
        ElseIf i = 9 And ws.Cells(7, i).ColumnWidth < min_I_Width Then
            ws.columns(i).ColumnWidth = min_I_Width
        End If
    Next i

    Rows(1).RowHeight = 60
    Rows(7).RowHeight = 36
    columns(1).ColumnWidth = 38

    Application.GoTo Reference:="R9C2", Scroll:=True
    'Debug.Print Now & " min width END"

    ' Restore application settings
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    DoEvents
    'Debug.Print Now & " AutoFit_ColumnsAndRows END"
End Sub
Private Function TransformArray(valuesArray As Variant, CompanyDivision As Long) As Variant
    Dim tempArray() As Variant
    Dim i As Long, j As Long
    Dim lastCol As Long
    Dim conoPos As Long, diviPos As Long
    Dim actualCols As Long
    
    ' --- Remove trailing empty column ---
    lastCol = UBound(valuesArray, 2)
    If Trim(valuesArray(1, lastCol) & "") = "" And Trim(valuesArray(2, lastCol) & "") = "" Then
        ReDim Preserve valuesArray(1 To 2, 1 To lastCol - 1)
    End If
    lastCol = UBound(valuesArray, 2)
    
    ' --- Find CONO/DIVI positions ---
    conoPos = -1
    diviPos = -1
    For i = 1 To lastCol
        If valuesArray(2, i) = "CONO" Then conoPos = i
        If valuesArray(2, i) = "DIVI" Then diviPos = i
    Next i
    
    Select Case CompanyDivision
        Case 0
            ReDim tempArray(1 To 2, 1 To lastCol)
            tempArray(1, 1) = ""
            tempArray(2, 1) = ""
            For i = 2 To lastCol
                tempArray(1, i) = valuesArray(1, i)
                tempArray(2, i) = valuesArray(2, i)
            Next i
            
        Case 1
            ReDim tempArray(1 To 2, 1 To lastCol)
            tempArray(1, 1) = ""
            tempArray(2, 1) = ""
            j = 2
            If conoPos > 0 Then
                tempArray(1, j) = valuesArray(1, conoPos)
                tempArray(2, j) = valuesArray(2, conoPos)
                j = j + 1
            End If
            If diviPos > 0 Then
                tempArray(1, j) = valuesArray(1, diviPos)
                tempArray(2, j) = valuesArray(2, diviPos)
                j = j + 1
            End If
            For i = 2 To lastCol
                If i <> conoPos And i <> diviPos Then
                    tempArray(1, j) = valuesArray(1, i)
                    tempArray(2, j) = valuesArray(2, i)
                    j = j + 1
                End If
            Next i
            If j <= lastCol Then ReDim Preserve tempArray(1 To 2, 1 To j - 1)
            
        Case 2
            actualCols = lastCol - IIf(conoPos > 0, 1, 0) - IIf(diviPos > 0, 1, 0)
            ReDim tempArray(1 To 2, 1 To actualCols)
            tempArray(1, 1) = ""
            tempArray(2, 1) = ""
            j = 2
            For i = 2 To lastCol
                If i <> conoPos And i <> diviPos Then
                    tempArray(1, j) = valuesArray(1, i)
                    tempArray(2, j) = valuesArray(2, i)
                    j = j + 1
                End If
            Next i
            
        Case 3
            ReDim tempArray(1 To 2, 1 To lastCol)
            tempArray(1, 1) = ""
            tempArray(2, 1) = ""
            j = 2
            For i = 2 To lastCol
                If i <> conoPos And i <> diviPos Then
                    tempArray(1, j) = valuesArray(1, i)
                    tempArray(2, j) = valuesArray(2, i)
                    j = j + 1
                End If
            Next i
            If conoPos > 0 Then
                tempArray(1, j) = valuesArray(1, conoPos)
                tempArray(2, j) = valuesArray(2, conoPos)
                j = j + 1
            End If
            If diviPos > 0 Then
                tempArray(1, j) = valuesArray(1, diviPos)
                tempArray(2, j) = valuesArray(2, diviPos)
                j = j + 1
            End If
            ReDim Preserve tempArray(1 To 2, 1 To j - 1)
    End Select
    
    TransformArray = tempArray
End Function


' =============================================================================
' AUTOFIT CLICK (Modernized)
' =============================================================================

Public Sub AutoFit_Click_New()
    Dim ws As Worksheet
    Dim currentEnv As String
    
    On Error GoTo Cleanup
    
    Set ws = ThisWorkbook.ActiveSheet
    
    ' Check environment
    On Error Resume Next
    If Trim(ws.Range("Environment").value) = "" Then Exit Sub
    On Error GoTo Cleanup
    
    currentEnv = ws.Range("I2").value
    
    ' Update version
    DoppioUI.UI_UpdateVersion
    
    ' =========================================================================
    ' AUTHENTICATION - Same pattern as Process_Click
    ' =========================================================================
    If Doppio.m_s_AccessToken = "" Or _
       ws.Range("I3").value = "" Or _
       ws.Range("I4").value = "" Or _
       Doppio.activeEnvironment <> currentEnv Then
        
        #If DEBUG_MODE Then
            Debug.Print "AutoFit_Click_New: Need new token"
        #End If
        Doppio.Tenant_Token
    End If
    
    ' Get transactions
    GetTransactions_Click_New
    
    ' Load layout (skip for EXPORTMI with table definition)
    If Not (ws.Range("API").value = "EXPORTMI" And ws.Range("A3").value = "table:  ") Then
        Transaction_LoadLayout_New ws
    End If
    
    ' Set formulas and formatting
    SetFormulasAndFormatting_New ws
    
    DoppioUI.UI_ShowPleaseWait "Please wait... Autofitting Columns"
    
    ' Format column A as text
    ws.Range("A9:A" & ws.Cells(ws.Rows.count, "A").End(xlUp).row).NumberFormat = "@"
    
    ' Set up freeze panes
    Application.ScreenUpdating = False
    ws.Cells(1, 1).Select
    ActiveWindow.FreezePanes = False
    ws.Range("C9").Select
    ActiveWindow.FreezePanes = True
    Application.ScreenUpdating = True
    
    ' Autofit and filter
    Doppio.AutoFit_ColumnsAndRows False, False
    FilterRow8BasedOnPopulatedColumns_New ws
    
Cleanup:
    Application.ScreenUpdating = True
    DoppioUI.UI_KillPleaseWait
End Sub

' =============================================================================
' TRANSACTION LOAD LAYOUT (Modernized)
' =============================================================================

Private Sub Transaction_LoadLayout_New(ws As Worksheet)
    Dim apiType As String
    Dim url As String
    
    On Error GoTo ErrorHandler
    
    apiType = ws.Cells(2, 2).value
    If apiType = "" Then apiType = "API"
    
    Select Case apiType
        Case "API"
            If ws.Range("API").value <> "" And ws.Range("Transaction").value <> "" Then
                url = "MRS001MI/LstFields;maxrecs=0?MINM=" & ws.Range("API").value & _
                      "&TRNM=" & ws.Range("Transaction").value
                GetLayoutAPI_New url, ws
            End If
            
        Case "IPS"
            Doppio.m_b_Webservice = False
            url = ws.Cells(2, 1).value & "/"
            GetLayoutWS_New "/M3/ips/service", url, ws
            
        Case "XtendM3"
            url = ws.Range("Transaction").value
            GetLayoutM3X_New "/M3/extensibility/ionapi-doc", url, ws
    End Select
    
    Exit Sub
    
ErrorHandler:
    #If DEBUG_MODE Then
        Debug.Print "Transaction_LoadLayout_New: ERROR - " & Err.description
    #End If
End Sub

' =============================================================================
' SET FORMULAS AND FORMATTING (Modernized)
' =============================================================================

Public Sub SetFormulasAndFormatting_New(ws As Worksheet)
    Dim targetCell As Range
    
    On Error Resume Next
    
    ' Set formula and formatting for cell I6 (record count)
    Set targetCell = ws.Range("I6")
    targetCell.NumberFormat = "General"
    targetCell.Formula = "=MAX(COUNTA(B9:B1048576), COUNTA(C9:C1048576), COUNTA(D9:D1048576), COUNTA(E9:E1048576), COUNTA(F9:F1048576), COUNTA(G9:G1048576), COUNTA(H9:H1048576), COUNTA(I9:I1048576), COUNTA(J9:J1048576))"
    
    ' Only set B4-B6 formulas if not EXPORTMI sheet (no "table:" label in A3)
    If ws.Range("A3").value = "" Then
        ' B4 = NOK count
        Set targetCell = ws.Range("B4")
        targetCell.NumberFormat = "General"
        targetCell.Formula = "=COUNTIF(A:A, ""NOK *"")"
        
        ' B5 = OK count
        Set targetCell = ws.Range("B5")
        targetCell.NumberFormat = "General"
        targetCell.Formula = "=COUNTIF(A:A, ""OK"")"
        
        ' B6 = Remaining count
        Set targetCell = ws.Range("B6")
        targetCell.NumberFormat = "General"
        targetCell.Formula = "=SUM(I6-(B4+B5))"
    End If
    
    On Error GoTo 0
End Sub

' =============================================================================
' LEGACY WRAPPER FOR AUTOFIT
' =============================================================================

Public Sub AutoFit_Click()
    ' Call modernized version
    AutoFit_Click_New
End Sub
' =============================================================================







Function Base64EncodeVBA(text As String) As String
    Dim bytes() As Byte
    Dim objXML As Object
    Dim objNode As Object
    
    ' Convert the text to a byte array
    bytes = StrConv(text, vbFromUnicode)
    
    ' Create an XML document to handle Base64 encoding
    Set objXML = CreateObject("MSXML2.DOMDocument")
    Set objNode = objXML.createElement("b64")
    
    ' Encode the byte array as Base64
    objNode.dataType = "bin.base64"
    objNode.nodeTypedValue = bytes
    Base64EncodeVBA = objNode.text
    
    ' Remove line breaks
    Base64EncodeVBA = Replace(Base64EncodeVBA, vbLf, "")
    
    ' Remove trailing '=' characters (padding)
    Do While Right(Base64EncodeVBA, 1) = "="
        Base64EncodeVBA = Left(Base64EncodeVBA, Len(Base64EncodeVBA) - 1)
    Loop
    
    ' Clean up
    Set objNode = Nothing
    Set objXML = Nothing
End Function
Function Base64DecodeVBA(encodedText As String) As String
    Dim objXML As Object
    Dim objNode As Object
    Dim bytes() As Byte
    
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
    Base64DecodeVBA = StrConv(bytes, vbUnicode)
    
    ' Clean up
    Set objNode = Nothing
    Set objXML = Nothing
End Function

Private Function BulkMIBody_create(method, aInputFields, aInputValues)
    Dim json As String
    Dim jsonTransactions As String
    Dim nonBlankPairs() As String
    Dim selectedColumns() As String
    Dim pairCount As Integer
    Dim j As Integer

    jsonTransactions = ""

    ' Check For non-blank fields And values
    pairCount = 0
    For j = LBound(aInputFields) To UBound(aInputFields)
        If aInputFields(j) <> "" And aInputValues(j) <> "" Then
            ReDim Preserve nonBlankPairs(pairCount)
            aInputValues(j) = Replace(aInputValues(j), """", "!!")
            
            'if PAR1 has non-numeric values make them 0
            If aInputFields(j) = "PAR1" Then
                Dim originalString As String
                Dim modifiedString As String
                originalString = aInputValues(j)
                modifiedString = ReplaceAlphaWithZero(originalString)
                aInputValues(j) = modifiedString
            End If
            
            nonBlankPairs(pairCount) = """" & aInputFields(j) & """: """ & aInputValues(j) & """"
            pairCount = pairCount + 1
        End If
    Next j

    pairCount = 0
    For j = LBound(aInputFields) To UBound(aInputFields)
        If aInputFields(j) <> "" Then
            ReDim Preserve selectedColumns(pairCount)
            selectedColumns(pairCount) = """" & aInputFields(j) & """"
            pairCount = pairCount + 1
        End If
    Next j

    ' Create the JSON record using non-blank field-value pairs
    Dim jsonRecord As String
    jsonRecord = "{"
    jsonRecord = jsonRecord & Join(nonBlankPairs, ", ")
    jsonRecord = jsonRecord & "}"

    Dim jsonSelected As String
    jsonSelected = "["
    jsonSelected = jsonSelected & Join(selectedColumns, ", ")
    jsonSelected = jsonSelected & "]"
    'Debug.Print jsonSelected

    ' Add the transaction To the JSON transactions array
    jsonTransactions = jsonTransactions & _
                       "{" & """transaction"":""" & method & """,""record"":" & jsonRecord & ",""selectedColumns"":" & jsonSelected & "}"
    json = jsonTransactions

    'Debug.Print json

    BulkMIBody_create = json

End Function

Private Sub GetTransactionsAPI(api As String)
    Dim main_url As String
    Dim mi_path As String
    Dim mi_url As String
    Dim i As Long
    Dim pair As Variant
    Dim found As Boolean
    
    ' Initialize the cache
    RecordCache_Initialize
    found = False

    ' Construct API parameters
    main_url = iu & "/" & ti
    mi_path = "/M3/m3api-rest/v2/execute"
    mi_url = "MRS001MI/LstTransactions;maxrecs=0?MINM=" & api & "&returncols=MINM,TRNM,SIMU"

    ' Check the cache for the URL
    Call RecordCache_Retreive(mi_url, found)

    ' Make the API call if records are not cached
    If Not found Then
        apicall_Bridge main_url, mi_path, mi_url, "", "API"
        Call RecordCache_Store(mi_url)
    End If

    ' Process the retrieved or cached records
    Output_LstTransactions
End Sub

Sub ChangeCellColorBasedOnEnvironment()
    Dim SettingsSheet As Worksheet
    Dim environmentRange As Range
    Dim sheet As Worksheet
    Dim targetCell As Range
    Dim sourceCell As Range
    Dim foundCell As Range

    ' Set references to the sheets
    Set SettingsSheet = ThisWorkbook.Sheets("Environments")
    Set sheet = ActiveSheet

    ' Get the selected environment from the data validation cell
    m_s_SelectedEnvironment = sheet.Range("Environment").value

    ' Find the selected environment in the "A" column of the "Environments" sheet
    Set environmentRange = SettingsSheet.Range("A:A")
    Set foundCell = environmentRange.Find(What:=m_s_SelectedEnvironment, LookIn:=xlValues, LookAt:=xlWhole)

    If Not foundCell Is Nothing Then
        Set sourceCell = SettingsSheet.Cells(foundCell.row, 1)
        Set targetCell = sheet.Range("I2:I5")
        targetCell.Font.Color = sourceCell.Font.Color
        targetCell.Interior.Color = sourceCell.Interior.Color
        If ActiveSheet.Name <> "AvailableMIs" Then
            sheet.Tab.Color = sourceCell.Interior.Color
        End If
    End If
End Sub

Sub CheckAndUpdateValue()
    Dim targetValue As Variant
    Dim lookupRange As Range
    Dim firstValue As Variant

    ' Define your target cell And lookup range
    Set lookupRange = Worksheets("Transactions").Range("B:B")
    targetValue = ActiveSheet.Cells(4, 7).value
    firstValue = lookupRange.Cells(3).value
    ' Check If the target value is empty Or Not in the lookup range
    If targetValue = "" Or isError(Application.match(targetValue, lookupRange, 0)) Then
        ' If target value is empty Or Not in the lookup range, Set it To the first value of the list
        ActiveSheet.Cells(4, 7).value = firstValue
    End If
End Sub



Sub CleanSheet_Click()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim lastCol As Long
    Dim lastCell As Range
    Dim lastColLetter As String

    On Error Resume Next

    ' Unhide all sheets
    For Each ws In ThisWorkbook.Sheets
        ws.Visible = xlSheetVisible
    Next ws

    For Each ws In ThisWorkbook.Sheets
        ws.Activate

        ' Find the last used row
        Set lastCell = ws.Cells.Find(What:="*", After:=ws.Cells(1, 1), LookIn:=xlFormulas, LookAt:=xlPart, _
                                     SearchOrder:=xlByRows, SearchDirection:=xlPrevious, MatchCase:=False)
        If Not lastCell Is Nothing Then
            lastRow = lastCell.row
        Else
            lastRow = 1
        End If
        
        ' Exceptions
        If ws.Range("F4").value = "Transaction:  " Then
            lastRow = 8
        End If
        If ws.Name = "Environments" Then
            lastRow = 52
        End If
        
        ' Find the last used column
        Set lastCell = ws.Cells.Find(What:="*", After:=ws.Cells(1, 1), LookIn:=xlFormulas, LookAt:=xlPart, _
                                     SearchOrder:=xlByColumns, SearchDirection:=xlPrevious, MatchCase:=False)
        If Not lastCell Is Nothing Then
            lastCol = lastCell.column
        Else
            lastCol = 1
        End If

        lastColLetter = Split(ws.Cells(1, lastCol).Address, "$")(1)

        If lastRow < ws.Rows.count Then
            ws.Rows(lastRow + 1 & ":" & ws.Rows.count).Delete
        End If

        ' Delete all columns to the right of the last used column
        If lastCol < ws.columns.count Then
            ws.columns(lastCol + 1 & ":" & ws.columns.count).Delete
        End If
    Next ws
    On Error GoTo 0
End Sub

Sub ClearFields()
    ActiveSheet.Range("User").value = ""
    ActiveSheet.Range("Company").value = ""
    ActiveSheet.Range("Division").value = ""
    ActiveSheet.Range("Transaction").value = ""
    ci = ""
    cs = ""
End Sub

Sub ClearLogSheet()
    Dim logsheet As Worksheet
    Dim lastRow As Long

    ' Set the log sheet
    On Error Resume Next
    Set logsheet = ThisWorkbook.Sheets("Log")
    On Error GoTo 0
    If logsheet Is Nothing Then
        MsgBox "Log sheet Not found.", vbExclamation
        Exit Sub
    End If

    ' Find the last used row in the log sheet
    lastRow = logsheet.Cells(logsheet.Rows.count, 1).End(xlUp).row

    ' Clear rows from row 2 To the last used row
    If lastRow >= 2 Then
        logsheet.Rows("2:" & lastRow).ClearContents
    End If

    'MsgBox "Log sheet cleared.", vbInformation
End Sub

Sub ClearOutputArea(ws As Worksheet, targetColor As Long)
    Dim dataStartColumn As Long
    Dim dataStartRow As Long
    Dim dataEndColumn As Long
    Dim dataEndRow As Long
    Dim recordCount As Long

    ' Define the starting row for the grey area search
    dataStartRow = 8
    dataStartColumn = 0
    recordCount = ws.Range("I6").value

    ' Find the first column with the target grey color in row 8
    dataStartColumn = FindFirstColumnWithColor(ws, dataStartRow, targetColor)
    'For i = 1 To ws.Columns.Count
    '    If ws.Cells(dataStartRow, i).Interior.Color = targetColor Then
    '        dataStartColumn = i
    '        Exit For
    '    End If
    'Next i

    ' Only proceed if a grey column was found
    If dataStartColumn > 0 Then
        ' Find the last row with data in the start column
        'dataEndRow = ws.Cells(ws.Rows.Count, dataStartColumn).End(xlUp).row
        dataEndRow = dataStartRow + recordCount
        If dataEndRow < 9 Then dataEndRow = 9    ' Ensure lastRow is at least 9 if no data below row 9

        ' Find the last column with data, starting from row 9
        dataEndColumn = ws.Cells(9, ws.columns.count).End(xlToLeft).column
        If dataEndColumn < dataStartColumn Then dataEndColumn = dataStartColumn
        
        ' Clear the contents in the identified range
        ws.Range(ws.Cells(dataStartRow, dataStartColumn), ws.Cells(dataEndRow, dataEndColumn)).ClearContents
        'Debug.Print "Cleared range: " & ws.Range(ws.Cells(dataStartRow, dataStartColumn), ws.Cells(dataEndRow, dataEndColumn)).Address
    Else
        #If DEBUG_MODE Then
            Debug.Print "No grey column found to clear."
        #End If
    End If
End Sub

Sub ClearStatus()
    Dim lastRow As Long
    Dim clearRange As Range
    Dim columnToCheck As Range
    Dim firstBlankRow As Long
    Dim ws As Worksheet
    
    ' Ensure ws is Set To the active sheet
    Set m_obj_ws = ActiveSheet
    Set ws = ThisWorkbook.ActiveSheet

    Set columnToCheck = m_obj_ws.columns("A")
    lastRow = m_obj_ws.Cells(m_obj_ws.Rows.count, "A").End(xlUp).row

    On Error Resume Next
    firstBlankRow = columnToCheck.Find(What:="", After:=m_obj_ws.Cells(8, 1), _
                                       LookIn:=xlValues, LookAt:=xlWhole, _
                                       SearchOrder:=xlByRows, SearchDirection:=xlNext).row
    On Error GoTo 0

    If firstBlankRow > 6 Then
        Set clearRange = m_obj_ws.Range("A9:A" & firstBlankRow)
        If Application.WorksheetFunction.CountA(clearRange) > 0 Then
            clearRange.ClearContents
        End If
    End If

    ' Load transaction layout
    Transaction_LoadLayout_New ws
End Sub

Sub ColorEntireRowFromColumnToEnd()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Sheet1")       ' Change "Sheet1" To the name of your sheet

    Dim startColumn As Integer
    startColumn = 3                              ' Change this To the column where you want To start coloring

    Dim lastColumn As Integer
    lastColumn = ws.Cells(1, ws.columns.count).End(xlToLeft).column

    Dim colorRange As Range
    Set colorRange = ws.Rows(1).columns(startColumn & ":" & lastColumn).EntireRow

    ' Change the color To your desired color (e.g., RGB(255, 0, 0) For red)
    colorRange.Interior.Color = RGB(255, 0, 0)
End Sub

Sub CreateCopyWithNewName()
    Dim newFileName As String
    newFileName = Replace(ThisWorkbook.Name, ".xlsm", ".xlsx")
    Dim newPath As String
    newPath = ThisWorkbook.path & "\" & newFileName
    Dim newWorkbook As Workbook
    Set newWorkbook = Workbooks.Add
    ThisWorkbook.Sheets.Copy Before:=newWorkbook.Sheets(1)
    newWorkbook.SaveAs newPath, FileFormat:=xlOpenXMLWorkbook
    newWorkbook.Close False
End Sub

Function CreateJsonBodyWithQuery() As String
    Dim program As String
    Dim transaction As String
    Dim record As String
    Dim selectedColumns As String
    Dim query As String
    Dim body As String
    Dim ws As Worksheet
    
    Set ws = ThisWorkbook.ActiveSheet
    
    ' Define the components of the JSON
    program = "EXPORTMI"
    transaction = "Select"
    
    ' make sure a split character is set
    splitChar = ""
    If splitChar = "" Then
        splitChar = Sheets("Settings").Range("D12").value
    End If
    If splitChar = "" Then
        splitChar = Sheets("Settings").Range("E12").value
    End If
    If splitChar = "" Then
        splitChar = "^"
    End If
    
    ' Get the query from cell B6
    query = ws.Range("B6").value
    
    ' Escape double quotes in the query string for JSON format
    query = Replace(query, """", "\""")
    query = Replace(query, "select ", "")
    
    record = """SEPC"":""" & splitChar & """,""HDRS"":""0"",""QERY"":""" & query & """"
    selectedColumns = """REPL"""
    
    ' Build the JSON string
    body = "{""program"":""" & program & """," _
                                           & """transactions"":[{" _
                                           & """transaction"":""" & transaction & """," _
                                                                                    & """record"":{" & record & "}," _
                                                                                                                 & """selectedColumns"":[" & selectedColumns & "]" _
                                                                                                                 & "}]}"
    
    ' Return the JSON string
    CreateJsonBodyWithQuery = body
End Function

Sub Create_xlsx()

    Dim ws As Worksheet
    Dim btn As Button

    DeleteHiddenSheets

    ' Loop through all worksheets
    For Each ws In ThisWorkbook.Sheets
        
        #If DEBUG_MODE Then
            Debug.Print ws.Name
        #End If
        
        ' Remove buttons
        For Each btn In ws.Buttons
            btn.Delete
        Next btn
        
        ws.Activate
        
        ' Remove rows
        ws.Rows("2:6").Delete
        
        ' Reset freeze panes
        ws.Cells(1, 1).Select
        ActiveWindow.FreezePanes = False
        ws.Range("C4").Select
        ActiveWindow.FreezePanes = True
    
    Next ws

    Dim newFileName As String
    newFileName = Replace(ThisWorkbook.Name, ".xlsm", ".xlsx")

    Dim pathSeparator As String
    #If Mac Then
        pathSeparator = "/"
    #Else
        pathSeparator = "\"
    #End If

    Dim newPath As String
    newPath = ThisWorkbook.path & pathSeparator & newFileName

    Dim filePath As Variant
    #If Mac Then
        filePath = Application.GetSaveAsFilename()
    #Else
        filePath = Application.GetSaveAsFilename(InitialFileName:=newFileName, FileFilter:="Excel Files (*.xlsx), *.xlsx")
    #End If

    If filePath <> "False" Then
        ThisWorkbook.SaveAs fileName:=filePath, FileFormat:=xlOpenXMLWorkbook
    End If

End Sub

Sub Default_Buttons()
    Dim btn1, btn2, btn3, btn4, btn5 As Button
    Dim ws As Worksheet
    Set ws = ActiveSheet
    
    Rows(1).RowHeight = 60
    Rows(7).RowHeight = 36
    
    ' Add buttons
    #If Mac Then
        Set btn1 = ws.Buttons.Add(8, 80, 69, 29)
        Set btn2 = ws.Buttons.Add(8, 113, 69, 29)
        Set btn3 = ws.Buttons.Add(80, 80, 69, 19)
        Set btn4 = ws.Buttons.Add(80, 102, 69, 18)
        Set btn5 = ws.Buttons.Add(80, 123, 69, 19)
    #Else
        Set btn1 = ws.Buttons.Add(8, 78, 69, 27)
        Set btn2 = ws.Buttons.Add(8, 107, 69, 27)
        Set btn3 = ws.Buttons.Add(80, 78, 69, 17)
        Set btn4 = ws.Buttons.Add(80, 98, 69, 17)
        Set btn5 = ws.Buttons.Add(80, 117, 69, 17)
    #End If
    
    btn1.Caption = "Transactions"
    btn1.OnAction = "GetTransactions_Click"
    btn2.Caption = "Run"
    btn2.OnAction = "Process_Click"
    btn3.Caption = "Layout"
    btn3.OnAction = "GetLayoutAll_Click"
    btn4.Caption = "Mandatory"
    btn4.OnAction = "GetLayoutMan_Click"
    btn5.Caption = "Autofit"
    btn5.OnAction = "AutoFit_Click"
    
    ' Add labels
    ws.Range("B6").NumberFormat = "General"
    ws.Range("B4:B6").NumberFormat = "General"

    ws.Range("A1").value = "." & vbCrLf & "." & vbCrLf & "." & vbCrLf & "."
    ws.Range("A3").value = "_____________________________________"
    ws.Range("A4").value = "NOK:"
    ws.Range("A5").value = "OK:"
    ws.Range("A6").value = "To Process:"
    ws.Range("A1").WrapText = True
    ws.Range("A3:A6").Font.Bold = False
    ws.Range("A3:A6").HorizontalAlignment = xlRight
    ws.Range("A3:A6").VerticalAlignment = xlCenter
    ws.Range("A3").Font.Color = RGB(255, 255, 255)
    ws.Range("A4").Font.Color = RGB(255, 0, 0)
    ws.Range("A5").Font.Color = RGB(0, 176, 80)
    ws.Range("A6").Font.Color = RGB(0, 0, 0)
          
    ws.Range("B3").value = "_________"
    ws.Range("B4").Formula = "=COUNTIF(A:A, ""NOK *"")"
    ws.Range("B5").Formula = "=COUNTIF(A:A, ""OK"")"
    ws.Range("B6").Formula = "=SUM(I6-(B4+B5))"
    ws.Range("B3").Font.Color = RGB(255, 255, 255)
    ws.Range("B3:B6").HorizontalAlignment = xlLeft
    ws.Range("B3:B6").VerticalAlignment = xlCenter
    ws.Range("B4:B6").Font.Color = RGB(0, 0, 0)
    
    ws.Range("i6").NumberFormat = "General"
    ws.Range("I6").Formula = "=MAX(COUNTA(B9:B1048576), COUNTA(C9:C1048576), COUNTA(D9:D1048576), COUNTA(E9:E1048576), COUNTA(F9:F1048576), COUNTA(G9:G1048576), COUNTA(H9:H1048576), COUNTA(I9:I1048576), COUNTA(J9:J1048576))"
End Sub

Sub DeleteHiddenSheets()
    Dim ws As Worksheet
    Dim i As Integer

    ' Loop through all worksheets in reverse order
    For i = ThisWorkbook.Sheets.count To 1 Step -1
        Set ws = ThisWorkbook.Sheets(i)

        ' Check if the sheet is hidden
        If ws.Visible = xlSheetHidden Or ws.Visible = xlSheetVeryHidden Then
            On Error Resume Next
            Application.DisplayAlerts = False
            ws.Delete
            If Err.Number <> 0 Then
                Err.Clear
            End If
            Application.DisplayAlerts = True
            On Error GoTo 0
        End If
    Next i
End Sub

Sub DisplayElapsedTime(startTime As Single, ws As Worksheet)
    Dim endTime As Single
    Dim elapsedSeconds As Single
    Dim elapsedHours As Integer, elapsedMinutes As Integer, remainingSeconds As Integer
    Dim formattedTime As String
    
    endTime = Timer
    elapsedSeconds = endTime - startTime
    
    If elapsedSeconds > 3600 Then
        elapsedHours = Int(elapsedSeconds / 3600)
        elapsedMinutes = Int((elapsedSeconds - elapsedHours * 3600) / 60)
        remainingSeconds = Int(elapsedSeconds Mod 60)
        formattedTime = Format(elapsedHours, "0") & "h " & Format(elapsedMinutes, "00") & "m " & Format(remainingSeconds, "00") & "s"
        
    ElseIf elapsedSeconds > 60 Then
        elapsedMinutes = Int(elapsedSeconds / 60)
        remainingSeconds = Int(elapsedSeconds Mod 60)
        formattedTime = Format(elapsedMinutes, "0") & "m " & Format(remainingSeconds, "00") & "s"
        
    Else
        formattedTime = Format(elapsedSeconds, "0.00") & "s"
    End If
    
    ' Set the formatted time in cell G6 of the specified worksheet
    ws.Range("G6").value = formattedTime
End Sub

Sub Environments_GetUsers()
    Dim ws As Worksheet
    Dim jsonString As String
    Dim json As Object
    Dim encodedTenant As String
    Dim lastRow As Long
    Dim i As Long
    Dim script As String
    Dim tempFilePath As String
    Dim fileNumber As String
    Dim result As String
    Dim record As Variant
    Dim mi_url As String
    
    Set ws = ActiveSheet
    lastRow = ws.Cells(ws.Rows.count, "B").End(xlUp).row ' Get the last row in column B
    
    For i = 1 To lastRow                         ' Loop through each row in column B
        jsonString = ws.Cells(i, 2).value        ' Get value from column B
        If jsonString = "" Then Exit For         ' Exit if the cell is empty
        
        Set json = JsonConverter.ParseJson(jsonString)
        encodedTenant = jsonString

        ' Extract And assign values To fields
        ti = json.item("ti")
        ci = json.item("ci")
        cs = json.item("cs")
        iu = json.item("iu")
        pu = json.item("pu")
        ot = json.item("ot")
        saak = json.item("saak")
        sask = json.item("sask")

        ' == MAC
        tempFilePath = Environ("HOME") & "/curl_output.txt"
        Kill tempFilePath

        curlCommand = "curl --location --max-time 20 '" & pu & ot & "' " & _
                      "--header 'Content-Type: application/x-www-form-urlencoded' " & _
                      "--data-urlencode 'client_id=" & ci & "' " & _
                      "--data-urlencode 'client_secret=" & cs & "' " & _
                      "--data-urlencode 'grant_type=password' " & _
                      "--data-urlencode 'username=" & saak & "' " & _
                      "--data-urlencode 'password=" & sask & "'"

        'Debug.Print curlCommand

        script = "Do shell script """ & curlCommand & " > ~/curl_output.txt"""

        ExecuteScriptWithRetry (script)

        If Err.Number <> 0 Then
            #If DEBUG_MODE Then
                Debug.Print "Tenant Token  Err.Number: " & Err.Number & " [" & script & "]"
            #End If
            MsgBox "An error occurred While executing the curl command."
        Else
            tempFilePath = Environ("HOME") & "/curl_output.txt"
            fileNumber = FreeFile
            Open tempFilePath For Input As fileNumber
            result = Input$(LOF(fileNumber), fileNumber)
            Close fileNumber

            Set json = JsonConverter.ParseJson(result)
            m_s_AccessToken = json.item("access_token")
            m_s_TokenType = json.item("token_type")
        End If

        m_s_MainUrl = iu & "/" & ti
        m_s_MiPath = "/M3/m3api-rest/v2/execute"
        mi_url = "MRS001MI/GetUserInfo/?"
        
        m_b_Webservice = False
        apicall_Bridge m_s_MainUrl, m_s_MiPath, mi_url, "", "API"
    
        ' Assuming Records is a valid collection from apicall
        For Each record In m_obj_Records
            ws.Cells(i, 3).value = ti
            If record.item("ZZUSID") = "" Then
                ws.Cells(i, 4).value = record.item("USFN")
            Else
                ws.Cells(i, 4).value = record.item("ZZUSID")
            End If
            'ws.Cells(i, 5).value = Base64EncodeVBA(ws.Cells(i, 2).value)
            ws.Cells(i, 5).value = Base64EncodeVBA(encodedTenant)
            ws.Cells(i, 5).WrapText = False
        Next record
    Next i
    KillPleaseWait
End Sub

' =============================================================================
' MODERNIZED Environments_Load
' =============================================================================
' Uses new DoppioHttp module for faster HTTP calls while maintaining
' compatibility with existing code.
'
' Replace your existing Environments_Load with this version.
' =============================================================================

Sub Environments_Load()
    Dim ws As Worksheet
    Dim body As String
    Dim config As httpConfig
    Dim httpResponse As httpResponse
    Dim json As Object
    Dim userName As String
    Dim machineName As String
    Dim fullUserName As String
    Dim encodedUserName As String
    Dim encodedMachineName As String
    Dim mainUrl As String
    Dim miPath As String
    Dim miUrl As String
    Dim tokenUrl As String
    Dim tokenBody As String
    Dim jsonString As String
    Dim TenantID As String
    Dim clientId As String
    Dim clientSecret As String
    Dim ssoBase As String
    Dim tokenEndpoint As String
    Dim authEndpoint As String
    Dim saak As String
    Dim sask As String
    Dim envCount As Long
    Dim encodedUserInfo As String
    
    On Error GoTo ErrorHandler
    
    ' Define worksheet and clear columns A through E
    Set ws = ThisWorkbook.Sheets("Environments")
    ws.columns("A:H").ClearContents
    
    ' Get and decode JSON configuration from Environments sheet
    jsonString = ws.Range("M2").value
    jsonString = Base64DecodeVBA(jsonString)
    
    ' Parse the JSON configuration
    Set json = JsonConverter.ParseJson(jsonString)
    
    ' Extract values from the JSON object
    TenantID = json.item("ti")
    clientId = json.item("ci")
    clientSecret = json.item("cs")
    ssoBase = json.item("pu")
    tokenEndpoint = json.item("ot")
    authEndpoint = json.item("oa")
    saak = json.item("saak")
    sask = json.item("sask")
    
    ws.Range("L2").value = TenantID
    
    ' Set up API call parameters
    mainUrl = json.item("iu") & "/" & TenantID
    miPath = "/M3/m3api-rest/v2/execute"
    miUrl = ""
    m_s_Company = ""
    m_s_Division = ""
    m_s_M3user = ""
    
    #If DEBUG_MODE Then
        Debug.Print "Environments_Load_New: Starting..."
        Debug.Print "  TenantID: " & TenantID
        Debug.Print "  MainUrl: " & mainUrl
    #End If
    
    ' Get user and machine info (works on both Mac and Windows)
    GetUserAndMachineInfo userName, fullUserName, machineName
    encodedUserName = Replace(UrlEncode(userName), "%20", " ")
    encodedMachineName = Replace(UrlEncode(machineName), "%20", " ")
    
    #If DEBUG_MODE Then
        Debug.Print "  UserName: " & userName
        Debug.Print "  MachineName: " & machineName
    #End If
    
    ' === STEP 1: Get Access Token ===
    tokenUrl = ssoBase & tokenEndpoint
    tokenBody = "client_id=" & Core_UrlEncode(clientId) & _
                "&client_secret=" & Core_UrlEncode(clientSecret) & _
                "&grant_type=password" & _
                "&username=" & Core_UrlEncode(saak) & _
                "&password=" & Core_UrlEncode(sask)
    
    #If DEBUG_MODE Then
        Debug.Print "  Getting token from: " & tokenUrl
    #End If
    
    config.url = tokenUrl
    config.method = HttpMethod_POST
    config.contentType = "application/x-www-form-urlencoded"
    config.AcceptType = "application/json"
    config.authHeader = ""
    config.body = tokenBody
    config.TimeoutSeconds = 30
    
    httpResponse = DoppioHttp.ExecuteRequest(config)
    
    If Not httpResponse.success Then
        #If DEBUG_MODE Then
            Debug.Print "  Token request failed: " & httpResponse.errorMessage
        #End If
        MsgBox "Unable to get token: " & httpResponse.errorMessage, vbCritical
        Exit Sub
    End If
    
    Set json = JsonConverter.ParseJson(httpResponse.body)
    m_s_AccessToken = json.item("access_token")
    m_s_TokenType = json.item("token_type")
    
    If m_s_AccessToken = "" Then
        #If DEBUG_MODE Then
            Debug.Print "  No access token in response"
        #End If
        MsgBox "Unable to get access token", vbCritical
        Exit Sub
    End If
    
    #If DEBUG_MODE Then
        Debug.Print "  Got token (" & Len(m_s_AccessToken) & " chars)"
    #End If
    
    ' === STEP 2: Get tenant list (EXAUTH = 20) ===
    body = "{""program"":""EXPORTMI"",""transactions"":[{""transaction"":""Select"",""record"":{""SEPC"":""^"",""HDRS"":""0"",""QERY"":""EXPCID,EXTNNM,EXHASH from EXTXSM where EXAUTH = 20""},""selectedColumns"":[""REPL""]}]}"
    
    #If DEBUG_MODE Then
        Debug.Print "  Getting tenant list..."
    #End If
    httpResponse = ExecuteApiPost(mainUrl, miPath, body)
    
    If httpResponse.success Then
        ' Parse response and set m_obj_Results for Environment_Tenants
        Set json = JsonConverter.ParseJson(httpResponse.body)
        If Not json Is Nothing Then
            Set m_obj_Results = json.item("results")
            #If DEBUG_MODE Then
                Debug.Print "  Tenant list: Got " & m_obj_Results.count & " results"
            #End If
        End If
        
        ' Call Environment_Tenants to process the results
        Environment_Tenants mainUrl
    Else
        #If DEBUG_MODE Then
            Debug.Print "  Tenant list request failed: " & httpResponse.errorMessage
        #End If
    End If
    
    ' === STEP 3: Get user's authorized environments (EXAUTH = 1) ===
    body = "{""program"":""EXPORTMI"",""transactions"":[{""transaction"":""Select"",""record"":{""SEPC"":""^"",""HDRS"":""0"",""QERY"":""EXTNNM,EXM3ID from EXTXSM where EXPCID = " & encodedUserName & " and EXAUTH = 1""},""selectedColumns"":[""REPL""]}]}"
    
    #If DEBUG_MODE Then
        Debug.Print "  Getting user environments..."
    #End If
    httpResponse = ExecuteApiPost(mainUrl, miPath, body)
    
    If httpResponse.success Then
        ' Parse response and set m_obj_Results for Environment_List
        Set json = JsonConverter.ParseJson(httpResponse.body)
        If Not json Is Nothing Then
            Set m_obj_Results = json.item("results")
            #If DEBUG_MODE Then
                Debug.Print "  User environments: Got " & m_obj_Results.count & " results"
            #End If
        End If
        
        ' Also set fileResult for backward compatibility
        fileResult = httpResponse.body
    Else
        #If DEBUG_MODE Then
            Debug.Print "  User environments request failed: " & httpResponse.errorMessage
        #End If
    End If
    
    ' === STEP 4: Check if user has any environments ===
    envCount = Environment_List()
    
    If envCount = 0 Then
        #If DEBUG_MODE Then
            Debug.Print "  No environments found, requesting access..."
        #End If
        
        Set ws = ThisWorkbook.Sheets("Environments")
        ws.Cells(1, 1).value = "Access requested"
        encodedUserInfo = Base64EncodeVBA(GetUserSessionInfo)
        
        body = "{""program"":""EXT124MI"",""transactions"":[{""transaction"":""AddUsrInfo"",""record"":{""PCID"":""" & encodedUserName & """,""TNNM"":""" & encodedMachineName & """,""M3ID"":""unknown"",""AUTH"":""99"",""HASH"":""" & encodedUserInfo & """}}]}"
        
        httpResponse = ExecuteApiPost(mainUrl, miPath, body)
        
        If Not httpResponse.success Then
            #If DEBUG_MODE Then
                Debug.Print "  Access request failed: " & httpResponse.errorMessage
            #End If
        End If
    Else
        #If DEBUG_MODE Then
            Debug.Print "  Found " & envCount & " environments"
        #End If
    End If
    
    ' Cleanup
    ws.columns("A:D").WrapText = False
    KillPleaseWait
    
    #If DEBUG_MODE Then
        Debug.Print "Environments_Load_New: Complete"
    #End If
    Exit Sub
    
ErrorHandler:
    #If DEBUG_MODE Then
        Debug.Print "Environments_Load_New: ERROR - " & Err.description
    #End If
    MsgBox "Error in Environments_Load_New: " & Err.description, vbCritical
    KillPleaseWait
End Sub

' =============================================================================
' Helper: Get user and machine info (cross-platform)
' =============================================================================
Private Sub GetUserAndMachineInfo(ByRef userName As String, ByRef fullUserName As String, ByRef machineName As String)
    On Error Resume Next
    
    #If Mac Then
        userName = MacScript("do shell script ""whoami""")
        fullUserName = MacScript("do shell script ""id -F""")
        machineName = MacScript("do shell script ""scutil --get ComputerName""")
    #Else
        userName = Environ("USERNAME")
        fullUserName = Environ("USERDOMAIN")
        machineName = Environ("COMPUTERNAME")
    #End If
    
    On Error GoTo 0
End Sub

' =============================================================================
' Helper: Execute API POST request
' =============================================================================
Private Function ExecuteApiPost(mainUrl As String, miPath As String, body As String) As httpResponse
    Dim config As httpConfig
    Dim apiUrl As String
    
    apiUrl = mainUrl & miPath
    
    config.url = apiUrl
    config.method = HttpMethod_POST
    config.contentType = "application/json; charset=UTF-8"
    config.AcceptType = "application/json; charset=UTF-8"
    config.authHeader = m_s_TokenType & " " & m_s_AccessToken
    config.body = body
    config.TimeoutSeconds = 30
    
    ExecuteApiPost = DoppioHttp.ExecuteRequest(config)
End Function

' =============================================================================
' Helper: URL Encode (if not already available)
' =============================================================================
Private Function Core_UrlEncode(text As String) As String
    Dim i As Long
    Dim c As String
    Dim result As String
    
    For i = 1 To Len(text)
        c = Mid(text, i, 1)
        Select Case c
            Case "A" To "Z", "a" To "z", "0" To "9", "-", "_", ".", "~"
                result = result & c
            Case " "
                result = result & "+"
            Case Else
                result = result & "%" & Right("0" & Hex(Asc(c)), 2)
        End Select
    Next i
    
    Core_UrlEncode = result
End Function



Sub Environment_Tenants(main_url As String)
    Dim i As Long
    Dim envKey As String
    Dim envValue1 As String
    Dim envValue2 As String
    Dim b64env As String
    Dim replSplit As Variant
    Dim result As Object
    Dim records As Object
    Dim record As Object
    
    Set manager = New EnvironmentManager
    If Not m_obj_Results Is Nothing Then
        i = 1
        For Each result In m_obj_Results
            Set records = result.item("records")
            If Not records Is Nothing Then
                For Each record In records
                    replSplit = Split(record.item("REPL"), "^")
                    If UBound(replSplit) = 2 And replSplit(1) <> "" Then
                        b64env = replSplit(2)
                        envKey = replSplit(0)
                        envValue1 = replSplit(1)
                        envValue2 = Base64DecodeVBA(b64env)
                        manager.AddEnvironment envKey, envValue1, envValue2, "", main_url, m_s_M3user, m_s_Company, m_s_Division
                    End If
                Next record
            End If
        Next result
    End If
End Sub
Function Environment_List() As Long
    Dim ws As Worksheet
    Dim i As Long
    Dim env As Environment
    Dim envKey As String
    Dim replSplit As Variant
    Dim result As Object
    Dim records As Object
    Dim record As Object

    Set ws = ThisWorkbook.Sheets("Environments")
    i = 1

    If Not m_obj_Results Is Nothing Then
        For Each result In m_obj_Results
            Set records = result.item("records")
            If Not records Is Nothing Then
                For Each record In records
                    replSplit = Split(record.item("REPL"), "^")
                    If UBound(replSplit) = 1 And replSplit(1) <> "" Then
                        envKey = replSplit(0)
                        Set env = manager.GetEnvironment(envKey)
                        
                        If Not env Is Nothing Then
                            ws.Cells(i, 1).value = env.tenant
                            ws.Cells(i, 2).value = env.Details
                            ws.Cells(i, 4).value = replSplit(1)
                        Else
                            ws.Cells(i, 1).value = "Access requested"
                            ws.Cells(i, 2).value = " "
                            ws.Cells(i, 4).value = " "
                        End If
                        
                        ws.Cells(i, 3).value = " "
                        ws.Cells(i, 5).value = " "
                        i = i + 1
                    End If
                Next record
            End If
        Next result
    End If
    
    ' return records loaded
    Environment_List = i - 1
End Function
Sub Curl_LoadResults()
    Dim tempFilePath As String
    Dim fileNumber As Integer
    Dim fileLength As Long

    On Error GoTo ErrorHandler

    tempFilePath = Environ("HOME") & "/curl_output.txt"
    If Dir(tempFilePath) = "" Then
        Exit Sub
    End If
    fileNumber = FreeFile
    Open tempFilePath For Input As #fileNumber
    fileLength = LOF(fileNumber)
    If fileLength = 0 Then
        Close #fileNumber
        Exit Sub
    End If
    fileResult = Input$(fileLength, fileNumber)
    Close #fileNumber

    Set m_obj_JsonResponse = JsonConverter.ParseJson(fileResult)

    If Not m_obj_JsonResponse.exists("results") Then
        'MsgBox "Error: 'results' key not found in JSON response.", vbExclamation
        Exit Sub
    End If

    Set m_obj_Results = m_obj_JsonResponse.item("results")
    Exit Sub

ErrorHandler:
    MsgBox "Unexpected error: " & Err.description, vbCritical
    On Error GoTo 0
End Sub


Sub Environment_SetDefaultValue()
    Dim ws As Worksheet
    Dim namedRange As Name
    Dim defaultValue As Variant
    Dim environmentSheet As Worksheet
    Dim tstRow As Long

    ' Set the active sheet
    Set ws = ActiveSheet

    ' Set the "Environments" sheet
    Set environmentSheet = ThisWorkbook.Sheets("Environments")

    ' Find the first row containing a name ending with "TST" or "TEST" in column A, starting from row 1
    tstRow = 0
    On Error Resume Next
    tstRow = Application.WorksheetFunction.match("TST*", environmentSheet.columns("A"), 0)
    If tstRow = 0 Then
        tstRow = Application.WorksheetFunction.match("TEST*", environmentSheet.columns("A"), 0)
    End If
    On Error GoTo 0

    ' If a name ending with "TST" or "TEST" is found, set defaultValue to the corresponding row; otherwise, use A1
    If tstRow > 0 Then
        defaultValue = environmentSheet.Cells(tstRow, 1).value
    Else
        defaultValue = environmentSheet.Range("A1").value
    End If

    ' Assume environmentSheet and ws are set to the relevant sheets
    On Error Resume Next
    Set namedRange = ws.names("Environment")
    On Error GoTo 0
    
    If Not namedRange Is Nothing Then
        On Error Resume Next
        ' Attempt to access the RefersToRange property
        If Not namedRange.RefersToRange Is Nothing Then
            ' Valid range; assign the default value
            namedRange.RefersToRange.value = defaultValue
        Else
            ' Handle the broken reference in namedRange
            #If DEBUG_MODE Then
                Debug.Print "Error: namedRange refers to an invalid range."
                Debug.Print "namedRange RefersTo: " & namedRange.RefersTo
            #End If
            ' Remove invalid named range
            ws.names("Environment").Delete
            Set namedRange = Nothing
        End If
        On Error GoTo 0
    End If
    
    ' Create the named range with the default value if it doesn't exist
    If namedRange Is Nothing Then
        On Error Resume Next
        ws.names.Add Name:="Environment", RefersTo:="=" & environmentSheet.Name & "!$I$2"
        ws.Range("I2").value = defaultValue
        On Error GoTo 0
    End If
End Sub

Sub ExecuteScriptWithRetry(ByRef script As String)
    Dim scriptRetry As Integer
    Dim scriptSuccess As Boolean
    Dim scriptResult As String
    Dim currentMaxTime As Integer
    Dim initialMaxTime As Integer
    Dim maxTimePos As Integer
    Dim spacePos As Integer
    Dim maxTimeStr As String
    Dim retryIncrement As Integer

    scriptRetry = 0
    scriptSuccess = False
    maxTimePos = InStr(script, "--max-time ")

    If maxTimePos > 0 Then
        spacePos = InStr(maxTimePos + 11, script, " ")
        If spacePos > 0 Then
            maxTimeStr = Mid(script, maxTimePos + 11, spacePos - (maxTimePos + 11))
        Else
            maxTimeStr = Mid(script, maxTimePos + 11)
        End If
        If IsNumeric(maxTimeStr) Then
            initialMaxTime = CInt(maxTimeStr)
        Else
            initialMaxTime = 20
        End If
    Else
        initialMaxTime = 20
    End If
    
    currentMaxTime = initialMaxTime
    If initialMaxTime < 10 Then
        currentMaxTime = 15
    Else
        currentMaxTime = initialMaxTime
    End If

    Do While scriptRetry < 6 And Not scriptSuccess
        On Error Resume Next
        scriptResult = MacScript(script)
        
        If Err.Number = 0 Then
            scriptSuccess = True
        Else
            #If DEBUG_MODE Then
                Debug.Print "Err.Number: " & Err.Number
                Debug.Print "script: " & script & vbCr & scriptResult
            #End If
            scriptRetry = scriptRetry + 1
            currentMaxTime = currentMaxTime * 2 ' Double the max-time value
            
            If Not PromptUser("max-time " & initialMaxTime & " hit, retry with " & currentMaxTime & "?") Then
                KillPleaseWait
                Err.Clear
                Exit Sub
            End If
            
            script = Replace(script, "--max-time " & initialMaxTime, "--max-time " & currentMaxTime)
            initialMaxTime = currentMaxTime
            DoppioUI.UI_ShowPleaseWait "Retry Attempt " & scriptRetry & "...  (--max-time " & currentMaxTime & ")"
            Err.Clear
        End If
    Loop
End Sub

Sub ExitProgram()
    Dim response As VbMsgBoxResult

    response = MsgBox("Do you want To stop the program?", vbYesNo + vbQuestion, "Stop Program")

    If response = vbYes Then
        ''Application.Calculation = xlCalculationAutomatic
        ''Application.ScreenUpdating = True
        KillPleaseWait
        ''DoEvents
        End
    End If

End Sub

Function ExtractServiceName(inputString As String) As String
    Dim startPos As Integer
    Dim endPos As Integer

    startPos = InStr(inputString, "services/") + Len("services/")
    endPos = InStr(startPos, inputString, "?wsdl")
    If startPos > 0 And endPos > startPos Then
        ExtractServiceName = Mid(inputString, startPos, endPos - startPos)
    Else
        ExtractServiceName = "Service name Not found"
    End If

End Function



Function FindColumnIndex(key As Variant) As Integer
    Dim i As Integer
    For i = 2 To m_obj_ws.Cells(8, columns.count).End(xlToLeft).column
        If m_obj_ws.Cells(8, i).value = key Then
            FindColumnIndex = i
            Exit Function
        End If
    Next i
    FindColumnIndex = 0
End Function

Function FindFirstColumnWithColor(ws As Worksheet, dataStartRow As Long, targetColor As Long) As Long
    Dim i As Long

    ' Loop through each column in the specified row
    For i = 1 To ws.columns.count
        ' Check if the cell color matches the target color
        If ws.Cells(dataStartRow, i).Interior.Color = targetColor Then
            ' Set the function result to the column number and exit
            FindFirstColumnWithColor = i
            Exit Function
        End If
    Next i

    ' Return 0 if no matching color is found
    FindFirstColumnWithColor = 0
End Function

' Helper function to format XML with proper indentation and minimal carriage returns
Function FormatXML(node As Object, indentLevel As Integer) As String
    Dim i As Integer
    Dim child As Object
    Dim result As String
    Dim indent As String
    Dim textContent As String
    
    ' Create the indent string based on the level
    For i = 1 To indentLevel
        indent = indent & vbTab                  ' Use tab for indentation
    Next i
    
    ' Start formatting the current node
    result = indent & "<" & node.nodeName
    
    ' Add attributes if there are any
    If node.Attributes.Length > 0 Then
        For i = 0 To node.Attributes.Length - 1
            result = result & " " & node.Attributes(i).nodeName & "="""
            result = result & node.Attributes(i).text & """"
        Next i
    End If
    
    ' If the node has child nodes, process them
    If node.HasChildNodes Then
        result = result & ">" & vbCrLf
        For Each child In node.ChildNodes
            If child.NodeType = 1 Then           ' Element node
                result = result & FormatXML(child, indentLevel + 1)
            ElseIf child.NodeType = 3 Then       ' Text node (ignore if it's just whitespace)
                textContent = Trim(child.text)
                If Len(textContent) > 0 Then
                    result = result & indent & vbTab & textContent & vbCrLf
                End If
            End If
        Next child
        result = result & indent & "</" & node.nodeName & ">" & vbCrLf
    Else
        ' If it's an empty node, close it on the same line
        result = result & "/>" & vbCrLf
    End If
    
    FormatXML = result
End Function



Public Sub SortByRequiredAndName( _
    ByRef names As Object, _
    ByRef descriptions As Object, _
    ByRef types As Object, _
    ByRef conditions As Object, _
    ByRef directions As Object _
)
    Dim i As Long, j As Long, n As Long
    n = conditions.count
    If n <= 1 Then Exit Sub
    
    Dim condI As Boolean, condJ As Boolean
    Dim nameI As String, nameJ As String
    
    Dim tmpName As Variant, tmpDesc As Variant, tmpType As Variant
    Dim tmpCond As Variant, tmpDir As Variant
    
    ' Bubble style sort with two keys:
    '  1. Required fields first
    '  2. Alphabetical by name inside each group
    For i = 1 To n - 1
        For j = i + 1 To n
            condI = IsTruthy(conditions.item(i))
            condJ = IsTruthy(conditions.item(j))
            nameI = CStr(names.item(i))
            nameJ = CStr(names.item(j))
            
            If (condJ And Not condI) _
               Or ((condI = condJ) And (UCase(nameJ) < UCase(nameI))) Then
               
                ' --- swap across all parallel arrays ---
                tmpName = names.item(i)
                names.SetItem i, names.item(j)
                names.SetItem j, tmpName
                
                tmpDesc = descriptions.item(i)
                descriptions.SetItem i, descriptions.item(j)
                descriptions.SetItem j, tmpDesc
                
                tmpType = types.item(i)
                types.SetItem i, types.item(j)
                types.SetItem j, tmpType
                
                tmpCond = conditions.item(i)
                conditions.SetItem i, conditions.item(j)
                conditions.SetItem j, tmpCond
                
                tmpDir = directions.item(i)
                directions.SetItem i, directions.item(j)
                directions.SetItem j, tmpDir
            End If
        Next j
    Next i
End Sub

' Helper: treat True, 1, "1" as truthy
Private Function IsTruthy(v As Variant) As Boolean
    If VarType(v) = vbBoolean Then
        IsTruthy = v
    Else
        IsTruthy = (val(v) <> 0)
    End If
End Function



Sub GetLayoutM3X(m_s_MiPath As String, mi_url As String)
    Dim methodName As String
    Dim jsonObject As Object
    Dim pathKey As Variant
    Dim methodKey As Variant
    Dim parameters As Object
    Dim parameter As Object
    Dim flds, typ, description As String

    If m_s_stUrl <> "" Then
        ' Not done
    Else
        apicall_Bridge m_s_MainUrl, m_s_MiPath, mi_url, "", ""

        m_obj_ColumnNames.Initialize
        m_obj_ColumnDescriptions.Initialize
        m_obj_ColumnTypes.Initialize
        m_obj_ColumnConditions.Initialize
        m_obj_ColumnDirections.Initialize

        ' Set the JSON data And method name
        methodName = ActiveSheet.Range("Transaction").value

        Set jsonObject = JsonConverter.ParseJson(m_s_CurlResult)

        For Each pathKey In jsonObject.item("paths").keys
            For Each methodKey In jsonObject.item("paths").item(pathKey).keys
                If jsonObject.item("paths").item(pathKey).item(methodKey).exists("parameters") Then
                    Set parameters = jsonObject.item("paths").item(pathKey).item(methodKey).item("parameters")
                    For Each parameter In parameters
                        Dim Alias As String
                        Dim Name As String
                        Dim dataType As String
                        Dim Required As Boolean

                        Alias = parameter.item("name")
                        Name = parameter.item("name")
                        dataType = parameter.item("type")
                        Required = parameter.item("required")

                        m_obj_ColumnNames.Add Alias
                        flds = Name & vbCrLf
                        typ = dataType
                        description = flds & typ
                        m_obj_ColumnDescriptions.Add description
                        m_obj_ColumnTypes.Add dataType
                        m_obj_ColumnConditions.Add Required
                        m_obj_ColumnDirections.Add "I"
                    
                    Next parameter
                End If
            Next methodKey
        Next pathKey
    End If
    'AutoFit_ColumnsAndRows False, False
End Sub



Sub GetLayoutWS(m_s_MiPath As String, mi_url As String)
    Dim methodName As String
    Dim methodData As Object
    Dim inputItem As Object
    Dim colNum As Integer
    Dim url As String
    Dim found As Boolean
    
    url = mi_url
    
'    If m_s_stUrl <> "" Then
'        Schema_ExtractInputandOuput
'    Else

    ' Check the cache for the URL
    Call RecordCache_Retreive(url, found)

     ' Make the API call if records are not cached
    If Not found Then
        apicall_Bridge m_s_MainUrl, m_s_MiPath, mi_url, "", ""
        Call RecordCache_Store(url)
    Else
        Set m_obj_JsonResponse = m_obj_Records
    End If
    
    m_obj_ColumnNames.Initialize
    m_obj_ColumnDescriptions.Initialize
    m_obj_ColumnTypes.Initialize
    m_obj_ColumnConditions.Initialize
    m_obj_ColumnDirections.Initialize

    ' Set the JSON data And method name
    methodName = ActiveSheet.Range("Transaction").value

    colNum = 2
    ' Find the method data
    If Not m_obj_JsonResponse Is Nothing Then
        For Each methodData In m_obj_JsonResponse.item("methods")
            If methodData.item("name") = methodName Then

                ' add input fields
                For Each inputItem In methodData.item("program").item("visibleInput")
                    Dim Alias As String
                    Dim Name As String
                    Dim dataType As String
                    Dim fieldLength As Long
                    Dim Required As Boolean
                    Dim flds, typ, leng, description As String
                    
                    Alias = inputItem.item("alias")
                    Name = inputItem.item("name")
                    dataType = inputItem.item("datatype")
                    fieldLength = inputItem.item("fieldLength")
                    Required = inputItem.item("required")

                    Select Case dataType
                    Case "STRING"
                        dataType = "A"
                    Case "DECIMAL"
                        dataType = "N"
                    End Select

                    m_obj_ColumnNames.Add Alias
                    flds = Name & vbCrLf
                    typ = dataType
                    leng = fieldLength
                    description = flds & typ & leng
                    m_obj_ColumnDescriptions.Add description
                    m_obj_ColumnTypes.Add dataType
                    m_obj_ColumnConditions.Add Required
                    m_obj_ColumnDirections.Add "I"
                    
                Next inputItem

                ' add output fields
                Dim outputItem As Object
                
                For Each outputItem In methodData.item("program").item("outputs")
                    Alias = outputItem.item("alias")
                    Name = outputItem.item("name")
                    dataType = outputItem.item("datatype")
                    fieldLength = outputItem.item("fieldLength")
                    Required = False

                    Select Case dataType
                    Case "STRING"
                        dataType = "A"
                    Case "DECIMAL"
                        dataType = "N"
                    End Select

                Next outputItem

            End If
        Next methodData
    End If
        
'    End If
End Sub



' =============================================================================
' MODERNIZED GetLayout_Click
' =============================================================================
' Uses new DoppioHttp module for API calls
' Follows same authentication pattern as Process_Click
' =============================================================================

Public Sub GetLayout_Click_New(mandatory As Boolean)
    Dim ws As Worksheet
    Dim url As String
    Dim apiType As String
    Dim currentEnv As String
    
    On Error GoTo ErrorHandler
    
    ' Update version
    DoppioUI.UI_UpdateVersion
    
    ' Check environment
    If Range("Environment").value = "" Then Exit Sub
    
    Set ws = ActiveSheet
    currentEnv = ws.Range("I2").value
    
    ' Clear status
    ws.Range("G6").value = ""
    
    DoppioUI.UI_ShowPleaseWait "Loading Transactions Layout"
    Application.GoTo Reference:="R9C2", Scroll:=True
    
    ' Set progress indicator
    ws.Range("J3").value = 0
    
    ' =========================================================================
    ' AUTHENTICATION - Same pattern as Process_Click
    ' =========================================================================
    If Doppio.m_s_AccessToken = "" Or _
       ws.Range("I3").value = "" Or _
       ws.Range("I4").value = "" Or _
       Doppio.activeEnvironment <> currentEnv Then
        
        #If DEBUG_MODE Then
            Debug.Print "GetLayout_Click_New: Need new token"
        #End If
        Doppio.Tenant_Token
        Doppio.m_l_Row = 9
    End If
    
    ws.Range("J3").value = 2
    
    ' Get API type
    apiType = ws.Cells(2, 2).value
    If apiType = "" Then apiType = "API"
    
    ' Process based on API type
    Select Case apiType
        Case "API"
            ProcessLayoutAPI ws, mandatory
            
        Case "IPS"
            ProcessLayoutIPS ws, mandatory
            
        Case "XtendM3"
            ProcessLayoutM3X ws, mandatory
            
        Case "IDM"
            ProcessLayoutIDM ws, mandatory
            
        Case Else
            ' Unknown type
    End Select
    
    ' Post-processing
    Doppio.AutoFit_ColumnsAndRows False, mandatory
    FilterRow8BasedOnPopulatedColumns_New ws
    DoppioUI.UI_KillPleaseWait
    
    Exit Sub
    
ErrorHandler:
    DoppioUI.UI_KillPleaseWait
    #If DEBUG_MODE Then
        Debug.Print "GetLayout_Click_New: ERROR - " & Err.description
    #End If
End Sub

' =============================================================================
' API LAYOUT (MRS001MI/LstFields)
' =============================================================================

Private Sub ProcessLayoutAPI(ws As Worksheet, mandatory As Boolean)
    Dim url As String
    Dim apiName As String
    Dim transName As String
    Dim columnCount As Long
    
    On Error GoTo ErrorHandler
    
    ' Set MI path (used by Doppio module)
    Doppio.m_s_MiPath = "/M3/m3api-rest/v2/execute"
    
    ' Rename sheet
    Doppio.RenameSheet ""
    
    ' Get API and Transaction from worksheet
    apiName = ws.Range("API").value
    transName = ws.Range("Transaction").value
    
    ' Build URL for MRS001MI/LstFields
    url = "MRS001MI/LstFields;maxrecs=0?MINM=" & apiName & "&TRNM=" & transName
    
    ' Get current column count (safely)
    On Error Resume Next
    columnCount = Doppio.m_obj_ColumnNames.count()
    If Err.Number <> 0 Then columnCount = 0
    On Error GoTo ErrorHandler
    columnCount = 0
    
    #If DEBUG_MODE Then
        Debug.Print "ProcessLayoutAPI: URL = " & url
        Debug.Print "ProcessLayoutAPI: m_s_LoadedUrl = " & Doppio.m_s_LoadedUrl
        Debug.Print "ProcessLayoutAPI: columnCount = " & columnCount
    #End If
    
    ' Special handling for EXPORTMI/Select - reset loaded URL if columns not loaded
    If apiName = "EXPORTMI" And Left(transName, 6) = "Select" Then
        If columnCount < 0 Then
            Doppio.m_s_LoadedUrl = ""
        End If
    End If
    
    ' Only fetch if API and Transaction are set
    If apiName <> "" And transName <> "" Then
        ' Fetch if:
        ' 1. URL changed from what we last loaded, OR
        ' 2. No columns are loaded (even if URL matches - columns may have been cleared)
        If Doppio.m_s_LoadedUrl <> url Or columnCount = 0 Then
            #If DEBUG_MODE Then
                Debug.Print "ProcessLayoutAPI: Calling GetLayoutAPI_New (URL changed or no columns)"
            #End If
            GetLayoutAPI_New url, ws
        Else
            #If DEBUG_MODE Then
                Debug.Print "ProcessLayoutAPI: Skipping load - URL matches and columns exist"
            #End If
        End If
        Doppio.AutoFit_ColumnsAndRows True, mandatory
    End If
    
    Exit Sub
    
ErrorHandler:
    #If DEBUG_MODE Then
        Debug.Print "ProcessLayoutAPI: ERROR - " & Err.description
    #End If
End Sub

Private Sub GetLayoutAPI_New(url As String, ws As Worksheet)
    Dim maxbulk_hold As Integer
    
    On Error GoTo ErrorHandler
    
    ' Store URL
    Doppio.m_s_LoadedUrl = url
    
    ' Initialize collections
    Doppio.m_obj_ColumnNames.Initialize
    Doppio.m_obj_ColumnDescriptions.Initialize
    Doppio.m_obj_ColumnTypes.Initialize
    Doppio.m_obj_ColumnConditions.Initialize
    Doppio.m_obj_ColumnDirections.Initialize
    
    ' Hold maxbulk
    maxbulk_hold = Doppio.maxbulk
    Doppio.maxbulk = 1
    
    ' Fetch input columns first
    FetchAndProcessColumns_New url & "&TRTP=I&returncols=FLNM,FLDS,TYPE,LENG,MAND", "I"
    
    ' Fetch output columns
    FetchAndProcessColumns_New url & "&TRTP=O&returncols=FLNM,FLDS,TYPE,LENG,MAND", "O"
    
    ' Restore maxbulk
    Doppio.maxbulk = maxbulk_hold
    
    Exit Sub
    
ErrorHandler:
    #If DEBUG_MODE Then
        Debug.Print "GetLayoutAPI_New: ERROR - " & Err.description
    #End If
End Sub

Private Sub FetchAndProcessColumns_New(url As String, direction As String)
    Dim response As apiResponse
    Dim record As Object
    Dim flds As String, typ As String, leng As String, description As String
    Dim value As String
    Dim mandValue As Variant
    
    On Error GoTo ErrorHandler
    
    ' 1. Try to get from new Cache
    ' Cache_TryGetFromCache populates 'response' if found and valid
    If Not DoppioCache.Cache_TryGetFromCache(url, response) Then
        
        #If DEBUG_MODE Then
            Debug.Print "FetchAndProcessColumns_New: Cache miss, calling API"
        #End If
        
        ' 2. Execute API Call
        response = ExecuteLayoutCall(url)
        
        ' 3. Store in new Cache if successful
        If response.success And Not response.records Is Nothing Then
            DoppioCache.Cache_StoreInCache url, response
            #If DEBUG_MODE Then
                Debug.Print "FetchAndProcessColumns_New: Stored in cache"
            #End If
        End If
    Else
        #If DEBUG_MODE Then
            Debug.Print "FetchAndProcessColumns_New: Cache hit"
        #End If
    End If
    
    ' 4. Process records (from response object, not global variable)
    If response.success And Not response.records Is Nothing Then
        #If DEBUG_MODE Then
            Debug.Print "FetchAndProcessColumns_New: Processing " & response.records.count & " records"
        #End If
        
        For Each record In response.records
            value = record.item("FLNM")
            
            ' For input, add all; for output, only add if not already present
            ' We still write to Doppio globals because downstream AutoFit functions likely depend on them [cite: 312]
            If direction = "I" Or (direction = "O" And Not Doppio.m_obj_ColumnNames.Contains(value)) Then
                Doppio.m_obj_ColumnNames.Add record.item("FLNM")
                
                flds = record.item("FLDS") & vbCrLf
                typ = record.item("TYPE")
                leng = record.item("LENG")
                description = flds & typ & leng
                
                Doppio.m_obj_ColumnDescriptions.Add description
                Doppio.m_obj_ColumnTypes.Add record.item("TYPE")
                
                mandValue = record.item("MAND")
                If IsNull(mandValue) Or mandValue = "" Then
                    mandValue = 0
                End If
                Doppio.m_obj_ColumnConditions.Add mandValue
                Doppio.m_obj_ColumnDirections.Add direction
            End If
        Next record
        
        #If DEBUG_MODE Then
            Debug.Print "FetchAndProcessColumns_New: Total columns = " & Doppio.m_obj_ColumnNames.count()
        #End If
    Else
        #If DEBUG_MODE Then
            Debug.Print "FetchAndProcessColumns_New: No records to process or API failed"
        #End If
    End If
    
    Exit Sub
    
ErrorHandler:
    #If DEBUG_MODE Then
        Debug.Print "FetchAndProcessColumns_New: ERROR - " & Err.description
    #End If
End Sub

' =============================================================================
' IPS LAYOUT (Web Services)
' =============================================================================

Private Sub ProcessLayoutIPS(ws As Worksheet, mandatory As Boolean)
    Dim url As String
    
    On Error GoTo ErrorHandler
    
    ' Rename sheet
    If Left(ws.Name, 5) = "Sheet" Then Doppio.RenameSheet ""
    
    Doppio.m_b_Webservice = False
    url = ws.Cells(2, 1).value & "/"
    
    If Doppio.m_s_LoadedUrl <> url Then
        GetLayoutWS_New "/M3/ips/service", url, ws
        Doppio.m_s_LoadedUrl = url
    End If
    
    Doppio.AutoFit_ColumnsAndRows True, mandatory
    
    Exit Sub
    
ErrorHandler:
    #If DEBUG_MODE Then
        Debug.Print "ProcessLayoutIPS: ERROR - " & Err.description
    #End If
End Sub

Private Sub GetLayoutWS_New(path As String, url As String, ws As Worksheet)
    Dim response As apiResponse
    Dim methodName As String
    Dim methodData As Object
    Dim inputItem As Object
    Dim outputItem As Object
    Dim flds As String, typ As String, leng As String, description As String
    Dim aliasName As String, fieldName As String, dataType As String
    Dim fieldLength As Long, Required As Boolean
    
    On Error GoTo ErrorHandler
    
    ' Check cache
    response = TryGetFromCache_Layout(url)
    
    If Not response.success Then
        response = ExecuteSwaggerLayoutCall(path, url)
        If response.success Then
            DoppioCache.Cache_StoreDataInCache url, response.data
        End If
    End If
    
    ' Initialize collections
    Doppio.m_obj_ColumnNames.Initialize
    Doppio.m_obj_ColumnDescriptions.Initialize
    Doppio.m_obj_ColumnTypes.Initialize
    Doppio.m_obj_ColumnConditions.Initialize
    Doppio.m_obj_ColumnDirections.Initialize
    
    ' Get method name
    methodName = ws.Range("Transaction").value
    
    ' Process response
    If response.success And Not response.results Is Nothing Then
        For Each methodData In response.results.item("methods")
            If methodData.item("name") = methodName Then
                
                ' Add input fields
                For Each inputItem In methodData.item("program").item("visibleInput")
                    aliasName = inputItem.item("alias")
                    fieldName = inputItem.item("name")
                    dataType = inputItem.item("datatype")
                    fieldLength = inputItem.item("fieldLength")
                    Required = inputItem.item("required")
                    
                    ' Convert data types
                    Select Case dataType
                        Case "STRING": dataType = "A"
                        Case "DECIMAL": dataType = "N"
                    End Select
                    
                    Doppio.m_obj_ColumnNames.Add aliasName
                    description = fieldName & vbCrLf & dataType & fieldLength
                    Doppio.m_obj_ColumnDescriptions.Add description
                    Doppio.m_obj_ColumnTypes.Add dataType
                    Doppio.m_obj_ColumnConditions.Add Required
                    Doppio.m_obj_ColumnDirections.Add "I"
                Next inputItem
                
                ' Add output fields
                For Each outputItem In methodData.item("program").item("outputs")
                    aliasName = outputItem.item("alias")
                    fieldName = outputItem.item("name")
                    dataType = outputItem.item("datatype")
                    fieldLength = outputItem.item("fieldLength")
                    
                    Select Case dataType
                        Case "STRING": dataType = "A"
                        Case "DECIMAL": dataType = "N"
                    End Select
                    
                    Doppio.m_obj_ColumnNames.Add aliasName
                    description = fieldName & vbCrLf & dataType & fieldLength
                    Doppio.m_obj_ColumnDescriptions.Add description
                    Doppio.m_obj_ColumnTypes.Add dataType
                    Doppio.m_obj_ColumnConditions.Add False
                    Doppio.m_obj_ColumnDirections.Add "O"
                Next outputItem
                
                Exit For
            End If
        Next methodData
    End If
    
    Exit Sub
    
ErrorHandler:
    #If DEBUG_MODE Then
        Debug.Print "GetLayoutWS_New: ERROR - " & Err.description
    #End If
End Sub

' =============================================================================
' XTENDM3 LAYOUT
' =============================================================================

Private Sub ProcessLayoutM3X(ws As Worksheet, mandatory As Boolean)
    Dim url As String
    
    On Error GoTo ErrorHandler
    
    url = ws.Range("Transaction").value
    
    If Doppio.m_s_LoadedUrl <> url Then
        GetLayoutM3X_New "/M3/extensibility/ionapi-doc", url, ws
        Doppio.m_s_LoadedUrl = url
    End If
    
    Doppio.AutoFit_ColumnsAndRows True, mandatory
    
    Exit Sub
    
ErrorHandler:
    #If DEBUG_MODE Then
        Debug.Print "ProcessLayoutM3X: ERROR - " & Err.description
    #End If
End Sub

Private Sub GetLayoutM3X_New(path As String, url As String, ws As Worksheet)
    Dim response As apiResponse
    Dim methodName As String
    Dim jsonObject As Object
    Dim pathKey As Variant
    Dim methodKey As Variant
    Dim parameters As Object
    Dim parameter As Object
    Dim aliasName As String, dataType As String
    Dim Required As Boolean
    Dim flds As String, description As String
    
    On Error GoTo ErrorHandler
    
    ' Make API call
    response = ExecuteSwaggerLayoutCall(path, url)
    
    ' Initialize collections
    Doppio.m_obj_ColumnNames.Initialize
    Doppio.m_obj_ColumnDescriptions.Initialize
    Doppio.m_obj_ColumnTypes.Initialize
    Doppio.m_obj_ColumnConditions.Initialize
    Doppio.m_obj_ColumnDirections.Initialize
    
    methodName = ws.Range("Transaction").value
    
    ' Process response
    If response.success And Len(response.data) > 0 Then
        Set jsonObject = JsonConverter.ParseJson(response.data)
        
        For Each pathKey In jsonObject.item("paths").keys
            For Each methodKey In jsonObject.item("paths").item(pathKey).keys
                If jsonObject.item("paths").item(pathKey).item(methodKey).exists("parameters") Then
                    Set parameters = jsonObject.item("paths").item(pathKey).item(methodKey).item("parameters")
                    
                    For Each parameter In parameters
                        aliasName = parameter.item("name")
                        dataType = parameter.item("type")
                        Required = parameter.item("required")
                        
                        Doppio.m_obj_ColumnNames.Add aliasName
                        description = aliasName & vbCrLf & dataType
                        Doppio.m_obj_ColumnDescriptions.Add description
                        Doppio.m_obj_ColumnTypes.Add dataType
                        Doppio.m_obj_ColumnConditions.Add Required
                        Doppio.m_obj_ColumnDirections.Add "I"
                    Next parameter
                End If
            Next methodKey
        Next pathKey
    End If
    
    Exit Sub
    
ErrorHandler:
    #If DEBUG_MODE Then
        Debug.Print "GetLayoutM3X_New: ERROR - " & Err.description
    #End If
End Sub

' =============================================================================
' IDM LAYOUT
' =============================================================================

Private Sub ProcessLayoutIDM(ws As Worksheet, mandatory As Boolean)
    Dim searchSheet As Worksheet
    Dim result As String
    Dim lastRow As Long, rowIndex As Long
    Dim strTransaction As String
    Dim openParenPos As Integer, closeParenPos As Integer
    Dim curlMethod As String
    
    On Error GoTo ErrorHandler
    
    ' Rename sheet if needed
    If Left(ws.Name, 5) = "Sheet" Then
        strTransaction = ws.Range("Transaction").value
        strTransaction = Mid(strTransaction, 2)
        openParenPos = InStr(strTransaction, "(")
        closeParenPos = InStr(strTransaction, ")")
        If openParenPos > 0 Then
            curlMethod = Mid(strTransaction, openParenPos + 1, closeParenPos - openParenPos - 1)
            strTransaction = Trim(Left(strTransaction, openParenPos - 1))
        End If
        Doppio.RenameSheet curlMethod & " " & strTransaction
    End If
    
    ' Find the layout parameters from Transactions sheet
    Set searchSheet = ThisWorkbook.Sheets("Transactions")
    lastRow = searchSheet.Cells(searchSheet.Rows.count, "A").End(xlUp).row
    
    For rowIndex = 3 To lastRow
        If searchSheet.Cells(rowIndex, 1).value = ws.Range("API").value And _
           searchSheet.Cells(rowIndex, 2).value = ws.Range("Transaction").value Then
            result = searchSheet.Cells(rowIndex, 5).value
            Exit For
        End If
    Next rowIndex
    
    ' Get IDM layout
    Doppio.IDM_GetLayout result
    Doppio.AutoFit_ColumnsAndRows True, mandatory
    
    Exit Sub
    
ErrorHandler:
    #If DEBUG_MODE Then
        Debug.Print "ProcessLayoutIDM: ERROR - " & Err.description
    #End If
End Sub

' =============================================================================
' API CALL HELPERS
' =============================================================================

Private Function ExecuteLayoutCall(url As String) As apiResponse
    Dim config As httpConfig
    Dim httpResponse As httpResponse
    Dim response As apiResponse
    Dim apiUrl As String
    Dim json As Object
    
    On Error GoTo ErrorHandler
    
    ' Build URL
    apiUrl = Doppio.m_s_MainUrl & "/M3/m3api-rest/v2/execute/" & url
    
    ' Configure request
    config.url = apiUrl
    config.method = HttpMethod_GET
    config.contentType = "application/json"
    config.AcceptType = "application/json"
    config.authHeader = Doppio.m_s_TokenType & " " & Doppio.m_s_AccessToken
    config.TimeoutSeconds = 30
    config.body = ""
    
    #If DEBUG_MODE Then
        Debug.Print "ExecuteLayoutCall: " & apiUrl
    #End If
    
    ' Execute
    httpResponse = DoppioHttp.ExecuteRequest(config)
    
    ' Build response
    response.success = httpResponse.success
    response.data = httpResponse.body
    response.errorMessage = httpResponse.errorMessage
    
    ' Parse JSON
    If httpResponse.success And Len(httpResponse.body) > 0 Then
        Set json = JsonConverter.ParseJson(httpResponse.body)
        If Not json Is Nothing Then
            If json.exists("results") Then
                Set response.results = json.item("results")
                If response.results.count > 0 Then
                    Set response.records = response.results(1).item("records")
                    If Not response.records Is Nothing Then
                        response.recordCount = response.records.count
                    End If
                End If
            End If
        End If
    End If
    
    ExecuteLayoutCall = response
    Exit Function
    
ErrorHandler:
    response.success = False
    response.errorMessage = Err.description
    ExecuteLayoutCall = response
End Function

Private Function ExecuteSwaggerLayoutCall(path As String, url As String) As apiResponse
    Dim config As httpConfig
    Dim httpResponse As httpResponse
    Dim response As apiResponse
    Dim apiUrl As String
    Dim json As Object
    
    On Error GoTo ErrorHandler
    
    ' Build URL
    apiUrl = Doppio.m_s_MainUrl & path & url
    
    ' Configure request
    config.url = apiUrl
    config.method = HttpMethod_GET
    config.contentType = "application/json"
    config.AcceptType = "application/json"
    config.authHeader = Doppio.m_s_TokenType & " " & Doppio.m_s_AccessToken
    config.TimeoutSeconds = 30
    config.body = ""
    
    #If DEBUG_MODE Then
        Debug.Print "ExecuteSwaggerLayoutCall: " & apiUrl
    #End If
    
    ' Execute
    httpResponse = DoppioHttp.ExecuteRequest(config)
    
    ' Build response
    response.success = httpResponse.success
    response.data = httpResponse.body
    response.errorMessage = httpResponse.errorMessage
    
    ' Parse JSON
    If httpResponse.success And Len(httpResponse.body) > 0 Then
        Set json = JsonConverter.ParseJson(httpResponse.body)
        Set response.results = json
    End If
    
    ExecuteSwaggerLayoutCall = response
    Exit Function
    
ErrorHandler:
    response.success = False
    response.errorMessage = Err.description
    ExecuteSwaggerLayoutCall = response
End Function

Private Function TryGetFromCache_Layout(cacheKey As String) As apiResponse
    Dim response As apiResponse
    Dim found As Boolean
    
    found = DoppioCache.Cache_TryGetFromCache(cacheKey, response)
    
    If Not found Then
        response.success = False
    End If
    
    TryGetFromCache_Layout = response
End Function

' =============================================================================
' UI HELPERS
' =============================================================================

Public Sub FilterRow8BasedOnPopulatedColumns_New(ws As Worksheet)
    Dim lastColumn As Long
    Dim dataRange As Range
    Dim colorRange As Range
    
    On Error Resume Next
    
    lastColumn = ws.Cells(8, ws.columns.count).End(xlToLeft).column
    
    ' If transaction is blank, clear headers
    If ws.Range("Transaction").value = "" Then
        lastColumn = 1
        ws.Rows("7:8").ClearContents
        Set colorRange = ws.Range(ws.Cells(8, lastColumn + 1), ws.Cells(8, 200))
        colorRange.Interior.Color = RGB(128, 128, 128)
        Set colorRange = ws.Range(ws.Cells(7, lastColumn + 1), ws.Cells(7, 200))
        colorRange.Interior.Color = RGB(192, 0, 0)
    End If
    
    ws.AutoFilterMode = False
    
    If lastColumn > 1 Then
        Set dataRange = ws.Range(ws.Cells(8, 1), ws.Cells(8, lastColumn))
        dataRange.AutoFilter field:=1
        Set colorRange = ws.Range(ws.Cells(8, lastColumn + 1), ws.Cells(8, 200))
        colorRange.Interior.Color = RGB(128, 128, 128)
        Set colorRange = ws.Range(ws.Cells(7, lastColumn + 1), ws.Cells(7, 200))
        colorRange.Interior.Color = RGB(192, 0, 0)
    End If
    
    On Error GoTo 0
End Sub

' =============================================================================
' LEGACY WRAPPERS
' =============================================================================

Public Sub GetLayout_Click(mandatory As Boolean)
    ' Call modernized version
    GetLayout_Click_New mandatory
End Sub

Public Sub GetLayoutAll_Click()
    GetLayout_Click_New False
End Sub

Public Sub GetLayoutMan_Click()
    GetLayout_Click_New True
End Sub
' =============================================================================

Sub GetSingleMultiple()
    GetTransactions_Click
    m_obj_ws.Range("G5").NumberFormat = "General"
    m_obj_ws.Range("G5").Formula = "=IFNA(VLOOKUP(Transaction,Transactions!B:C,2,False),"""")"
    m_obj_ws.Calculate
    DoEvents
End Sub

Sub GetTransactionsAllWS()
    Dim swaggerArray, swagger, swaggerEndpoint, entity, entityParts
    Dim rowNum As Integer
    Dim resultSheet As Worksheet
    
    rowNum = 1

    apicall_Bridge m_s_MainUrl, "/M3/ips/service/ionapi-doc", "?&pageSize=1000", "", ""

    ' Parse JSON
    Set swaggerArray = m_obj_JsonResponse("swaggerCollection")("swagger")

    ' Set the worksheet To store the results
    Set resultSheet = ThisWorkbook.Sheets("Available WS")

    ' Clear the previous data in the sheet
    resultSheet.UsedRange.Clear

    ' Extract swagger endpoints And split into two fields
    For Each swagger In swaggerArray
        swaggerEndpoint = swagger("swaggerEndpoint")
        entity = swagger("entity")

        ' Split entity into two parts using the '#' delimiter
        entityParts = Split(entity, "#")

        ' Append the result To the output string
        resultSheet.Cells(rowNum, 1).value = entityParts(0)
        resultSheet.Cells(rowNum, 2).value = entityParts(1)
        rowNum = rowNum + 1
    Next swagger
End Sub

Sub GetTransactionsM3X(m_s_MiPath As String, mi_url As String)
    Dim swaggerArray, swagger, swaggerEndpoint, entity
    Dim rowNum As Integer
    Dim resultSheet As Worksheet
    
    rowNum = 3
    m_b_Webservice = False

    Set resultSheet = ThisWorkbook.Sheets("Transactions")
    resultSheet.UsedRange.Clear

    If m_s_stUrl <> "" Then
        'Schema_ExtractTransactions
    Else
        apicall_Bridge m_s_MainUrl, m_s_MiPath, mi_url, "", "API"

        Set swaggerArray = m_obj_JsonResponse.item("swaggerCollection").item("swagger")

        ' Extract swagger endpoints And split into two fields
        For Each swagger In swaggerArray
            swaggerEndpoint = swagger.item("swaggerEndpoint")
            entity = swagger.item("entity")

            ' Append the result To the output string
            resultSheet.Cells(rowNum, 1).value = "Extensibility"
            resultSheet.Cells(rowNum, 2).value = entity
            resultSheet.Cells(rowNum, 3).value = "S"
            rowNum = rowNum + 1
        Next swagger
    End If

End Sub

Sub GetTransactionsWS2(m_s_MiPath As String, mi_url As String)
    Dim swaggerArray As Variant, swagger As Variant
    Dim swaggerEndpoint As String, entity As String, entityParts As Variant
    Dim rowNum As Integer
    Dim resultSheet As Worksheet
    Dim i As Long
    Dim pair As Variant
    Dim found As Boolean
    Dim recordPair As String
    Dim serializedRecords As String

    rowNum = 3
    m_b_Webservice = False

    ' Initialize cache
    RecordCache_Initialize
    found = False

    ' Get the result sheet and clear it
    Set resultSheet = ThisWorkbook.Sheets("Transactions")
    resultSheet.UsedRange.Clear

    ' Check the schema URL
    If m_s_stUrl <> "" Then
        Schema_ExtractTransactions
    Else
        ' Check the cache for the URL
        For i = 1 To m_RecordCache.count
            pair = Split(m_RecordCache.item(i), "|")
            If pair(0) = mi_url Then
                Set swaggerArray = JsonConverter.ParseJson(pair(1)) ' Rehydrate cached records
                found = True
                Exit For
            End If
        Next i

        ' Make the API call if records are not cached
        If Not found Then
            apicall_Bridge m_s_MainUrl, m_s_MiPath, mi_url, "", ""

            ' Cache the records if retrieved
            If Not m_obj_JsonResponse Is Nothing Then
                swaggerArray = m_obj_JsonResponse("swaggerCollection")("swagger")
                serializedRecords = JsonConverter.ConvertToJson(swaggerArray)
                recordPair = mi_url & "|" & serializedRecords
                m_RecordCache.Add recordPair
                RecordCache_Dump
            End If
        End If

        ' Extract swagger endpoints and split into two fields
        For Each swagger In swaggerArray
            swaggerEndpoint = swagger("swaggerEndpoint")
            entity = swagger("entity")

            ' Split entity into two parts using the '#' delimiter
            entityParts = Split(entity, "#")

            ' Append the result to the output string
            resultSheet.Cells(rowNum, 1).value = entityParts(0)
            resultSheet.Cells(rowNum, 2).value = entityParts(1)
            resultSheet.Cells(rowNum, 3).value = "S"
            rowNum = rowNum + 1
        Next swagger
    End If
End Sub

Sub GetTransactionsWS(m_s_MiPath As String, mi_url As String)
    Dim swaggerArray, swagger, swaggerEndpoint, entity, entityParts
    Dim rowNum As Integer
    Dim resultSheet As Worksheet
    Dim found As Boolean
    Dim url As String
    
    rowNum = 3
    m_b_Webservice = False
    url = mi_url
    
    Set resultSheet = ThisWorkbook.Sheets("Transactions")
    resultSheet.UsedRange.Clear

'    If m_s_stUrl <> "" Then
'        Schema_ExtractTransactions
'    Else

    ' Check the cache for the URL
    Call RecordCache_Retreive(url, found)

     ' Make the API call if records are not cached
    If Not found Then
        apicall_Bridge m_s_MainUrl, m_s_MiPath, mi_url, "", ""
        Call RecordCache_Store(url)
    Else
        Set m_obj_JsonResponse = m_obj_Records
    End If

    If Not m_obj_JsonResponse Is Nothing Then
        Set swaggerArray = m_obj_JsonResponse.item("swaggerCollection").item("swagger")
    
        ' Extract swagger endpoints And split into two fields
        For Each swagger In swaggerArray
            swaggerEndpoint = swagger.item("swaggerEndpoint")
            entity = swagger.item("entity")
    
            ' Split entity into two parts using the '#' delimiter
            entityParts = Split(entity, "#")
    
            ' Append the result To the output string
            resultSheet.Cells(rowNum, 1).value = entityParts(0)
            resultSheet.Cells(rowNum, 2).value = entityParts(1)
            resultSheet.Cells(rowNum, 3).value = "S"
            rowNum = rowNum + 1
        Next swagger
    End If

End Sub



' =============================================================================
' MODERNIZED GetTransactions_Click
' =============================================================================
' Uses new DoppioHttp module for API calls
' Follows same authentication pattern as Process_Click
' =============================================================================

Public Sub GetTransactions_Click_New()
    Dim ws As Worksheet
    Dim api As String
    Dim apiType As String
    Dim currentEnv As String
    
    On Error GoTo ErrorHandler
    
    ' Update version display
    DoppioUI.UI_UpdateVersion
    
    ' Check environment
    If Range("Environment").value = "" Then Exit Sub
    
    Set ws = ActiveSheet
    currentEnv = ws.Range("I2").value
    
    ' Clear G6 if exists
    On Error Resume Next
    ws.Range("G6").value = ""
    On Error GoTo ErrorHandler
    
    DoppioUI.UI_ShowPleaseWait "Loading Transactions For API"
    Application.GoTo Reference:="R9C2", Scroll:=True
    
    ' Set version display
    With ws.Range("J2")
        .value = DOPPIO_VERSION
        .Font.Color = RGB(255, 255, 255)
        .Font.Size = 10
        .Font.Italic = True
        .HorizontalAlignment = xlLeft
    End With
    
    ' =========================================================================
    ' AUTHENTICATION - Same pattern as Process_Click
    ' =========================================================================
    If Doppio.m_s_AccessToken = "" Or _
       ws.Range("I3").value = "" Or _
       ws.Range("I4").value = "" Or _
       Doppio.activeEnvironment <> currentEnv Then
        
        #If DEBUG_MODE Then
            Debug.Print "GetTransactions_Click_New: Need new token"
        #End If
        Doppio.Tenant_Token
        Doppio.m_l_Row = 9
    End If
    
    ' Get API info
    api = ws.Cells(2, 1).value
    apiType = ws.Cells(2, 2).value
    If apiType = "" Then apiType = "API"
    
    ' Process based on API type
    Select Case apiType
        Case "API"
            GetTransactionsAPI_New api, ws
        Case "IPS"
            GetTransactionsIPS_New api, ws
        Case "XtendM3"
            GetTransactionsM3X_New ws
        Case "IDM"
            Doppio.IDM_Load_Methods
        Case Else
            ' Unknown type
    End Select
    
    ' Post-processing
    Doppio.SortTransactions
    Doppio.CheckAndUpdateValue
    DoppioUI.UI_KillPleaseWait
    UpdateG5Formula_New ws
    
    Exit Sub
    
ErrorHandler:
    DoppioUI.UI_KillPleaseWait
    #If DEBUG_MODE Then
        Debug.Print "GetTransactions_Click_New: ERROR - " & Err.description
    #End If
End Sub

' =============================================================================
' API TRANSACTIONS (MRS001MI/LstTransactions)
' =============================================================================

Private Sub GetTransactionsAPI_New(api As String, ws As Worksheet)
    Dim response As apiResponse
    Dim cacheKey As String
    Dim found As Boolean
    
    On Error GoTo ErrorHandler
    
    ' Initialize cache
    DoppioCache.RecordCache_Initialize
    
    ' Build cache key
    cacheKey = "MRS001MI/LstTransactions?MINM=" & api
    
    ' Check cache first
    response = TryGetFromCache(cacheKey)
    
    If Not response.success Then
        ' Make API call
        response = ExecuteTransactionCall("MRS001MI", "LstTransactions", "MINM=" & api & "&returncols=MINM,TRNM,SIMU")
        
        ' Store in cache if successful
        If response.success Then
            DoppioCache.Cache_StoreDataInCache cacheKey, response.data
        End If
    End If
    
    ' Output results to Transactions sheet
    If response.success Then
        OutputLstTransactions_New response
    End If
    
    Exit Sub
    
ErrorHandler:
    #If DEBUG_MODE Then
        Debug.Print "GetTransactionsAPI_New: ERROR - " & Err.description
    #End If
End Sub

' =============================================================================
' IPS TRANSACTIONS (Swagger)
' =============================================================================

Private Sub GetTransactionsIPS_New(api As String, ws As Worksheet)
    Dim response As apiResponse
    Dim cacheKey As String
    Dim url As String
    
    On Error GoTo ErrorHandler
    
    url = "?&pageSize=1000&search=" & api
    cacheKey = "IPS:" & url
    
    ' Check cache
    response = TryGetFromCache(cacheKey)
    
    If Not response.success Then
        ' Make API call
        response = ExecuteSwaggerCall("/M3/ips/service/ionapi-doc", url)
        
        If response.success Then
            DoppioCache.Cache_StoreDataInCache cacheKey, response.data
        End If
    End If
    
    ' Output results
    If response.success Then
        OutputSwaggerTransactions_New response
    End If
    
    Exit Sub
    
ErrorHandler:
    #If DEBUG_MODE Then
        Debug.Print "GetTransactionsIPS_New: ERROR - " & Err.description
    #End If
End Sub

' =============================================================================
' XTENDM3 TRANSACTIONS
' =============================================================================

Private Sub GetTransactionsM3X_New(ws As Worksheet)
    Dim response As apiResponse
    
    On Error GoTo ErrorHandler
    
    ' Make API call
    response = ExecuteSwaggerCall("M3/extensibility/ionapi-doc", "")
    
    ' Output results
    If response.success Then
        OutputM3XTransactions_New response
    End If
    
    Exit Sub
    
ErrorHandler:
    #If DEBUG_MODE Then
        Debug.Print "GetTransactionsM3X_New: ERROR - " & Err.description
    #End If
End Sub

' =============================================================================
' API CALL HELPERS
' =============================================================================

Private Function ExecuteTransactionCall(program As String, transaction As String, parameters As String) As apiResponse
    Dim config As httpConfig
    Dim httpResponse As httpResponse
    Dim response As apiResponse
    Dim apiUrl As String
    Dim json As Object
    
    On Error GoTo ErrorHandler
    
    ' Build URL
    apiUrl = Doppio.m_s_MainUrl & "/M3/m3api-rest/v2/execute/" & program & "/" & transaction & ";maxrecs=0?" & parameters
    
    ' Configure request
    config.url = apiUrl
    config.method = HttpMethod_GET
    config.contentType = "application/json"
    config.AcceptType = "application/json"
    config.authHeader = Doppio.m_s_TokenType & " " & Doppio.m_s_AccessToken
    config.TimeoutSeconds = 30
    config.body = ""
    
    #If DEBUG_MODE Then
        Debug.Print "ExecuteTransactionCall: " & apiUrl
    #End If
    
    ' Execute
    httpResponse = DoppioHttp.ExecuteRequest(config)
    
    ' Build response
    response.success = httpResponse.success
    response.data = httpResponse.body
    response.errorMessage = httpResponse.errorMessage
    
    ' Parse JSON
    If httpResponse.success And Len(httpResponse.body) > 0 Then
        Set json = JsonConverter.ParseJson(httpResponse.body)
        If Not json Is Nothing Then
            If json.exists("results") Then
                Set response.results = json.item("results")
                If response.results.count > 0 Then
                    Set response.records = response.results(1).item("records")
                    If Not response.records Is Nothing Then
                        response.recordCount = response.records.count
                    End If
                End If
            End If
        End If
    End If
    
    ExecuteTransactionCall = response
    Exit Function
    
ErrorHandler:
    response.success = False
    response.errorMessage = Err.description
    ExecuteTransactionCall = response
End Function

Private Function ExecuteSwaggerCall(path As String, queryString As String) As apiResponse
    Dim config As httpConfig
    Dim httpResponse As httpResponse
    Dim response As apiResponse
    Dim apiUrl As String
    Dim json As Object
    
    On Error GoTo ErrorHandler
    
    ' Build URL
    apiUrl = Doppio.m_s_MainUrl & "/" & path & queryString
    
    ' Configure request
    config.url = apiUrl
    config.method = HttpMethod_GET
    config.contentType = "application/json"
    config.AcceptType = "application/json"
    config.authHeader = Doppio.m_s_TokenType & " " & Doppio.m_s_AccessToken
    config.TimeoutSeconds = 30
    config.body = ""
    
    #If DEBUG_MODE Then
        Debug.Print "ExecuteSwaggerCall: " & apiUrl
    #End If
    
    ' Execute
    httpResponse = DoppioHttp.ExecuteRequest(config)
    
    ' Build response
    response.success = httpResponse.success
    response.data = httpResponse.body
    response.errorMessage = httpResponse.errorMessage
    
    ' Parse JSON for swagger
    If httpResponse.success And Len(httpResponse.body) > 0 Then
        Set json = JsonConverter.ParseJson(httpResponse.body)
        If Not json Is Nothing Then
            Set response.results = json
        End If
    End If
    
    ExecuteSwaggerCall = response
    Exit Function
    
ErrorHandler:
    response.success = False
    response.errorMessage = Err.description
    ExecuteSwaggerCall = response
End Function

Private Function TryGetFromCache(cacheKey As String) As apiResponse
    Dim response As apiResponse
    Dim found As Boolean
    
    found = DoppioCache.Cache_TryGetFromCache(cacheKey, response)
    
    If Not found Then
        response.success = False
    End If
    
    TryGetFromCache = response
End Function

' =============================================================================
' OUTPUT HELPERS
' =============================================================================

Private Sub OutputLstTransactions_New(response As apiResponse)
    Dim ws As Worksheet
    Dim record As Object
    Dim rowNum As Long
    
    On Error GoTo ErrorHandler
    
    Set ws = ThisWorkbook.Sheets("Transactions")
    ws.UsedRange.Clear
    
    If response.records Is Nothing Then Exit Sub
    
    rowNum = 3
    For Each record In response.records
        ws.Cells(rowNum, 1).value = record.item("MINM")
        ws.Cells(rowNum, 2).value = record.item("TRNM")
        ws.Cells(rowNum, 3).value = record.item("SIMU")
        rowNum = rowNum + 1
    Next record
    
    Exit Sub
    
ErrorHandler:
    #If DEBUG_MODE Then
        Debug.Print "OutputLstTransactions_New: ERROR - " & Err.description
    #End If
End Sub

Private Sub OutputSwaggerTransactions_New(response As apiResponse)
    Dim ws As Worksheet
    Dim swaggerArray As Object
    Dim swagger As Object
    Dim entity As String
    Dim entityParts() As String
    Dim rowNum As Long
    
    On Error GoTo ErrorHandler
    
    Set ws = ThisWorkbook.Sheets("Transactions")
    ws.UsedRange.Clear
    
    If response.results Is Nothing Then Exit Sub
    
    ' Get swagger collection
    Set swaggerArray = response.results.item("swaggerCollection").item("swagger")
    
    rowNum = 3
    For Each swagger In swaggerArray
        entity = swagger.item("entity")
        entityParts = Split(entity, "#")
        
        ws.Cells(rowNum, 1).value = entityParts(0)
        If UBound(entityParts) >= 1 Then
            ws.Cells(rowNum, 2).value = entityParts(1)
        End If
        ws.Cells(rowNum, 3).value = "S"
        rowNum = rowNum + 1
    Next swagger
    
    Exit Sub
    
ErrorHandler:
    #If DEBUG_MODE Then
        Debug.Print "OutputSwaggerTransactions_New: ERROR - " & Err.description
    #End If
End Sub

Private Sub OutputM3XTransactions_New(response As apiResponse)
    Dim ws As Worksheet
    Dim swaggerArray As Object
    Dim swagger As Object
    Dim entity As String
    Dim rowNum As Long
    
    On Error GoTo ErrorHandler
    
    Set ws = ThisWorkbook.Sheets("Transactions")
    ws.UsedRange.Clear
    
    If response.results Is Nothing Then Exit Sub
    
    ' Get swagger collection
    Set swaggerArray = response.results.item("swaggerCollection").item("swagger")
    
    rowNum = 3
    For Each swagger In swaggerArray
        entity = swagger.item("entity")
        
        ws.Cells(rowNum, 1).value = "Extensibility"
        ws.Cells(rowNum, 2).value = entity
        ws.Cells(rowNum, 3).value = "S"
        rowNum = rowNum + 1
    Next swagger
    
    Exit Sub
    
ErrorHandler:
    #If DEBUG_MODE Then
        Debug.Print "OutputM3XTransactions_New: ERROR - " & Err.description
    #End If
End Sub

Private Sub UpdateG5Formula_New(ws As Worksheet)
    On Error Resume Next
    ws.Range("G5").NumberFormat = "General"
    ws.Range("G5").Formula = "=IFNA(VLOOKUP(Transaction,Transactions!B:C,2,False),"""")"
    ws.Calculate
    DoEvents
    On Error GoTo 0
End Sub

' =============================================================================
' LEGACY WRAPPER
' =============================================================================

Public Sub GetTransactions_Click()
    ' Call modernized version
    GetTransactions_Click_New
End Sub
' =============================================================================


Sub HelpSheet()
    Sheets("Help").Visible = True
    Sheets("Help").Activate
    With ActiveWindow
        If .FreezePanes Then .FreezePanes = False
        .SplitColumn = 0
        .SplitRow = 5
        .FreezePanes = True
    End With
    Application.GoTo Reference:="R1C1", Scroll:=True
End Sub

Sub IDM_BuildJSONFromWorksheet(startCol As Integer, startRow As Integer, row As Integer, ByRef jsonResult As String)
    Dim ws As Worksheet
    Dim jsonArray As Collection
    Dim jsonItem As Dictionary
    Dim paramName As String
    Dim paramValue As String
    Dim columnIndex As Integer
    Dim jsonObject As Dictionary
    Dim itemObject As Dictionary
    Dim attrsObject As Dictionary
    
    JsonConverter.JsonOptions.UseDoubleForLargeNumbers = True

    ' Set your worksheet
    Set ws = ActiveSheet
    
    ' Initialize JSON collections and dictionaries
    Set jsonArray = New Collection
    Set jsonObject = New Dictionary
    Set itemObject = New Dictionary
    Set attrsObject = New Dictionary
    
    ' Start reading data from the specified starting column and row
    columnIndex = startCol
    currentPid = ""
    
    ' Loop through columns until an empty cell in the header row is found
    Do While ws.Cells(startRow - 1, columnIndex).value <> ""
        paramName = ws.Cells(startRow - 1, columnIndex).value
        paramValue = ws.Cells(row, columnIndex).value
        
        If paramName = "pid" Then
            currentPid = paramValue
        Else
            ' Create a JSON item and add it to the JSON array
            Set jsonItem = New Dictionary
            jsonItem.Add "name", paramName
            If paramName = "MDS_ID" Then
                jsonItem.Add "type", "21"
            Else
                jsonItem.Add "type", "1"
            End If
            jsonItem.Add "qual", paramName
            If paramValue <> "" Then
                jsonItem.Add "value", paramValue
            End If
            
            jsonArray.Add jsonItem
        End If
        
        columnIndex = columnIndex + 1
    Loop

    ' Add the JSON array to the attrs object with the key "attr"
    attrsObject.Add "attr", jsonArray
    
    ' Add attrs and pid to the item object
    itemObject.Add "attrs", attrsObject
    'itemObject.Add "pid", "string"  ' Replace "string" with the actual PID value if available
    
    ' Add the item object to the root JSON object
    jsonObject.Add "item", itemObject

    ' Convert the root JSON object to a JSON string
    jsonResult = JsonConverter.ConvertToJson(jsonObject)
    
    ' Output the JSON result
    'Debug.Print jsonResult
    
End Sub

Sub IDM_GetLayout(jsonParameters As String)
    Dim jsonArray As Object
    Dim param As Variant
    Dim paramDetails As Object
    Dim aliasName As String
    Dim Name As String
    Dim dataType As String
    Dim Required As Boolean
    
    ' Initialize collections or arrays if needed
    m_obj_ColumnNames.Initialize
    m_obj_ColumnDescriptions.Initialize
    m_obj_ColumnTypes.Initialize
    m_obj_ColumnConditions.Initialize
    m_obj_ColumnDirections.Initialize

    If jsonParameters = "" Then
        Exit Sub
    End If

    ' Parse the JSON string into a usable object
    Set jsonArray = JsonConverter.ParseJson(jsonParameters)

    ' Loop through each parameter in the "parameters" array
    For Each param In jsonArray
        Set paramDetails = param
        aliasName = paramDetails.item("name")
        Name = Replace(paramDetails.item("name"), "$", "")
        dataType = paramDetails.item("type")
        Required = paramDetails.item("required")

        Select Case dataType
        Case "string"
            dataType = "A"
        Case "integer"
            dataType = "N"
        Case "boolean"
            dataType = "N"
        Case "object"
            dataType = "O"
        Case Else
            dataType = "A"
        End Select

        Name = Name & vbCrLf & dataType
                    
        ' Add to collections or arrays
        m_obj_ColumnNames.Add aliasName
        m_obj_ColumnDescriptions.Add Name
        m_obj_ColumnTypes.Add dataType
        m_obj_ColumnConditions.Add Required
        m_obj_ColumnDirections.Add "I"
    Next param

End Sub

Sub IDM_Load_Methods()
    'Dim m_obj_JsonResponse As Object
    Dim paths As Object
    Dim pathKey As Variant
    Dim methodKey As Variant
    Dim methodDetails As Object
    Dim rowNum As Long
    Dim tag As Variant
    Dim resultSheet As Worksheet
    Dim api, transaction, method, summary As String
    Dim ws As Worksheet
    Set ws = ActiveSheet
    Dim found As Boolean

    ' Set the initial row number and worksheet
    Set resultSheet = ThisWorkbook.Sheets("Transactions")
    resultSheet.UsedRange.Clear
    rowNum = 3
    
    ' Load last result from the file
    'tempFilePath = Environ("HOME") & "/curl_output.txt"
    'fileNumber = FreeFile
    'Open tempFilePath For Input As fileNumber
    'm_s_CurlResult = Input$(LOF(fileNumber), fileNumber)
    'Close fileNumber
    
    Call RecordCache_Retreive(ws.Range("API").value, found)
    
    If found Then
        IDM_Load_Cache
    Else
        apicall_Bridge m_s_MainUrl, "/IDM/api", "ionapi-doc", "", "API"
        Set m_obj_JsonResponse = JsonConverter.ParseJson(m_s_CurlResult)
        Set paths = m_obj_JsonResponse.item("paths")
        
        ' Loop through the keys in the "paths" object (dictionary)
        Dim jsonArray As Object
        Set jsonArray = JsonConverter.ParseJson("[]")
        For Each pathKey In paths.keys
            For Each methodKey In paths.item(pathKey).keys
                If IsObject(paths.item(pathKey).item(methodKey)) Then
                    Set methodDetails = paths.item(pathKey).item(methodKey)
                    transaction = pathKey
                    method = UCase(methodKey)
                    For Each tag In methodDetails.item("tags")
                        api = tag
                    Next tag
                    summary = methodDetails.item("summary")
                    
                    ' Output values to the worksheet
                    If api = ws.Range("API").value Then
                        resultSheet.Cells(rowNum, 1).value = api
                        resultSheet.Cells(rowNum, 2).value = transaction & " (" & method & ")"
                        resultSheet.Cells(rowNum, 3).value = method
                        resultSheet.Cells(rowNum, 4).value = summary
                       
                        ' Check if parameters exist
                        If methodDetails.exists("parameters") Then
                            If IsObject(methodDetails.item("parameters")) Then
                                Dim parametersJson As String
                                Dim jsonValue As Object
                                Set jsonValue = methodDetails.item("parameters")
                                parametersJson = JsonConverter.ConvertToJson(jsonValue)
                                resultSheet.Cells(rowNum, 5).value = parametersJson
                            End If
                        End If
                        
                        ' Check if requestBody exists
                        If methodDetails.exists("requestBody") Then
                            If IsObject(methodDetails.item("requestBody")) Then
                                Dim requestBodyJson As String
                                Dim rbValue As Object
                                Set rbValue = methodDetails.item("requestBody")
                                requestBodyJson = JsonConverter.ConvertToJson(rbValue)
                                resultSheet.Cells(rowNum, 6).value = requestBodyJson
                            End If
                        End If
                        
                        rowNum = rowNum + 1
                        
                        Dim jsonObject As Object
                        Dim finalJson As String
                        Set jsonObject = JsonConverter.ParseJson("{}")
                        jsonObject.item("api") = api
                        jsonObject.item("transaction") = transaction & " (" & method & ")"
                        jsonObject.item("method") = method
                        jsonObject.item("summary") = summary
                        jsonObject.item("parameters") = parametersJson
                        jsonArray.Add jsonObject
                        
                    End If
                End If
            Next methodKey
        Next pathKey
        
        Set m_obj_Records = jsonArray
        Call RecordCache_Store(ws.Range("API").value)
    End If
    
End Sub

Sub IDM_Load_Results(startCol As Integer, row As Integer)
    Dim jsonResponse As Object
    Dim jsonValue As Object
    Dim paths As Object
    Dim pathKey As Variant
    Dim col As Integer
    Dim isMultiple As Boolean
     
    ' Load last result from the file
    'tempFilePath = Environ("HOME") & "/curl_output.txt"
    'fileNumber = FreeFile
    'Open tempFilePath For Input As fileNumber
    'm_s_CurlResult = input$(LOF(fileNumber), fileNumber)
    'Close fileNumber
    
    ' Parse the JSON string into a usable object
    Set jsonResponse = JsonConverter.ParseJson(m_s_CurlResult)
    
    ' Check if it's a single record or multiple records
    isMultiple = False
    If jsonResponse.exists("items") Then
        ' Check if "items.item" is an array or a collection (use TypeName)
        If TypeName(jsonResponse.item("items").item("item")) = "Collection" Or TypeName(jsonResponse.item("items").item("item")) = "Array" Then
            Set paths = jsonResponse.item("items").item("item")
            isMultiple = True
        End If
    ElseIf jsonResponse.exists("item") Then
        Set paths = jsonResponse.item("item")
        m_obj_ws.Cells(row, 1).value = "OK"
        m_obj_ws.Cells(row, 1).Font.Color = "40000"
    ElseIf jsonResponse.exists("terminationReason") Then
        m_obj_ws.Cells(row, 1).value = "NOK " & jsonResponse.item("terminationReason")
        m_obj_ws.Cells(row, 1).Font.Color = "255"
        Exit Sub
    ElseIf jsonResponse.exists("error") Then
        m_obj_ws.Cells(row, 1).value = "NOK " & jsonResponse.item("error").item("message")
        m_obj_ws.Cells(row, 1).Font.Color = "255"
        Exit Sub
    End If
    
    ' Write headers (only once)
    If row = 9 Then
        col = startCol
        
        If isMultiple Then
            ' Multiple records case, write headers from the first item
            Dim firstItem As Object
            Set firstItem = paths(1)             ' Assuming the first item exists in the collection/array
            For Each pathKey In firstItem.keys
                If pathKey = "pid" Then
                    m_obj_ws.Cells(8, col).value = pathKey ' Set header
                    col = col + 1
                End If
            Next pathKey
        Else
            ' Single record case, write headers from the single item
            'For Each pathKey In paths.Keys
                'm_obj_ws.Cells(8, col).value = pathKey ' Set header
                'col = col + 1
            'Next pathKey
        End If
    End If
    
    ' Reset column and row
    'row = row + 1
    'startRow = row
    'col = startCol
    
    ' Write values
    If isMultiple Then
        ' Multiple records case
        Dim item As Object
        For Each item In paths
            col = startCol
            For Each pathKey In item.keys
                If IsObject(item.item(pathKey)) Then
                    Set jsonValue = item.item(pathKey)
                    'm_obj_ws.Cells(row, col).value = JsonConverter.ConvertToJson(jsonValue)
                    If pathKey = "attrs" Then
                        Call IDM_Process_JSON_Parameters(JsonConverter.ConvertToJson(jsonValue), col, 8, row)
                    End If
                Else
                    If pathKey = "pid" Then
                        m_obj_ws.Cells(row, col).value = item.item(pathKey)
                        col = col + 1
                    End If
                End If
            Next pathKey
            row = row + 1
        Next item
    Else
        ' Single record case
        For Each pathKey In paths.keys
            If IsObject(paths.item(pathKey)) Then
                Set jsonValue = paths.item(pathKey)
                'm_obj_ws.Cells(row, col).value = JsonConverter.ConvertToJson(jsonValue)
                If pathKey = "attrs" Then
                    Call IDM_Process_JSON_Parameters(JsonConverter.ConvertToJson(jsonValue), startCol, 8, row)
                End If
            Else
                'm_obj_ws.Cells(row, col).value = paths(pathKey)
                'col = col + 1
            End If
        Next pathKey
    End If
End Sub

Sub IDM_Load_Tags()
    Dim jsonResponse As Object
    Dim paths As Object
    Dim pathItem As Variant
    Dim httpMethod As Variant
    Dim tagsArray As Object
    Dim tag As Variant
    
    Dim uniqueTags As Collection
    Dim resultSheet As Worksheet
    Dim rowNum As Long
    
    Dim tempFilePath As String
    Dim fileNumber As Integer
    Dim key As String
    
    ' Set worksheet and start row
    Set resultSheet = ThisWorkbook.Sheets("AvailableMIs")
    rowNum = 9
    
    ' --------------------------------------------------
    ' Debug: load curl output (same pattern as yours)
    ' --------------------------------------------------
    If 1 = 1 Then
        tempFilePath = Environ("HOME") & "/curl_output.txt"
        fileNumber = FreeFile
        Open tempFilePath For Input As fileNumber
        m_s_CurlResult = Input$(LOF(fileNumber), fileNumber)
        Close fileNumber
    End If
    
    ' --------------------------------------------------
    ' Parse JSON
    ' --------------------------------------------------
    Set jsonResponse = JsonConverter.ParseJson(m_s_CurlResult)
    Set paths = jsonResponse.item("paths")
    
    ' Collection for unique tags (Mac & Windows)
    Set uniqueTags = New Collection
    
    ' --------------------------------------------------
    ' Walk paths ? methods ? tags
    ' --------------------------------------------------
    For Each pathItem In paths.keys
        For Each httpMethod In paths.item(pathItem).keys
            If paths.item(pathItem).item(httpMethod).exists("tags") Then
                Set tagsArray = paths.item(pathItem).item(httpMethod).item("tags")
                For Each tag In tagsArray
                    ' Case-insensitive uniqueness key
                    key = LCase(CStr(tag))
                    On Error Resume Next
                    uniqueTags.Add CStr(tag), key
                    On Error GoTo 0
                Next tag
            End If
        Next httpMethod
    Next pathItem
    
    ' --------------------------------------------------
    ' Output unique tags
    ' --------------------------------------------------
    For Each tag In uniqueTags
        resultSheet.Cells(rowNum, 3).value = tag
        resultSheet.Cells(rowNum, 7).value = tag
        rowNum = rowNum + 1
    Next tag

End Sub



Sub IDM_Process()
    Dim openParenPos As Integer
    Dim closeParenPos As Integer
    Dim row As Integer
    Dim lastColumn As Long
    Dim column As Long
    Dim arrayIndex As Long
    Dim startTime As Single
    Dim startRow As Integer
    Dim strTransaction As String
    Dim recordCount As String
    Dim maxbulk_hold As Integer
    
    Set m_obj_ws = ActiveSheet
       
    If m_s_AccessToken = "" Or m_obj_ws.Range("I3").value = "" Or m_obj_ws.Range("I4").value = "" Or activeEnvironment <> m_obj_ws.Range("I2") Then
        Tenant_Token
    Else
        m_obj_ws.Range("J3").value = 2
    End If
    
    If m_obj_ws.Range("G5").value = "" Then GetTransactions_Click
    
    ClearStatus
    maxbulk_hold = maxbulk
    maxbulk = 1
    
    ' get transaction and method
    strTransaction = m_obj_ws.Range("Transaction").value
    strTransaction = Mid(strTransaction, 2)
    openParenPos = InStr(strTransaction, "(")
    closeParenPos = InStr(strTransaction, ")")
    If openParenPos > 0 Then
        curlMethod = Mid(strTransaction, openParenPos + 1, closeParenPos - openParenPos - 1)
        strTransaction = Trim(Left(strTransaction, openParenPos - 1))
    End If

    recordCount = m_obj_ws.Range("I6").value
    m_s_MiPath = "/IDM/api"
    startRow = 9
    
    ' extract url key
    Dim urlKey As String
    Dim startPos As Integer
    Dim endPos As Integer
    startPos = InStr(strTransaction, "{") + 1
    endPos = InStr(strTransaction, "}")
    If startPos > 0 And endPos > startPos Then
        urlKey = Mid(strTransaction, startPos, endPos - startPos)
    Else
        urlKey = ""
    End If
    
    ' clear if method is not PUT
    If curlMethod <> "PUT" Then
        ClearOutputArea m_obj_ws, RGB(128, 128, 128)
    End If

    ' Get the Header columns To use from the Data sheet
    ReDim strInputFields(1 To 400)
    lastColumn = m_obj_ws.Cells(8, m_obj_ws.columns.count).End(xlToLeft).column
    row = startRow
 
    ' populate the key fields
    arrayIndex = 0
    For column = 2 To lastColumn
        If m_obj_ws.Cells(row - 1, column) <> "" Then
            arrayIndex = arrayIndex + 1
            If arrayIndex > 400 Then Exit For
            strInputFields(arrayIndex) = m_obj_ws.Cells(row - 1, column)
        End If
    Next column
    
    ' === build body parm START ===
    If curlMethod = "PUT" Then
        row = startRow
        Dim attributeStart As Integer
        Dim bodyColumn As Long
        Dim pidColumn As Long
        
        ' Find first column with gray color (attributes)
        attributeStart = FindFirstColumnWithColor(m_obj_ws, startRow - 1, RGB(128, 128, 128))
        
        ' Find the column that contains "body" in row 8
        bodyColumn = 0
        Dim col As Long
        lastColumn = m_obj_ws.Cells(8, m_obj_ws.columns.count).End(xlToLeft).column
        For col = 1 To lastColumn
            If LCase(Trim(m_obj_ws.Cells(8, col).value)) = "body" Then
                bodyColumn = col
                Exit For
            End If
        Next col
        For col = 1 To lastColumn
            If LCase(Trim(m_obj_ws.Cells(8, col).value)) = "pid" Then
                pidColumn = col
                Exit For
            End If
        Next col
        
        If bodyColumn = 0 Then
            Exit Sub
        End If
               
        ' Loop through rows and build JSON
        If m_obj_ws.Cells(startRow - 1, attributeStart).value <> "" Then
            While row < (startRow + recordCount)
                Dim jsonResult As String
                Call IDM_BuildJSONFromWorksheet(attributeStart, 9, row, jsonResult)
                
                ' Write JSON to the 'body' column
                m_obj_ws.Cells(row, bodyColumn).value = jsonResult
                If m_obj_ws.Cells(row, pidColumn).value = "" Then m_obj_ws.Cells(row, pidColumn).value = currentPid
                row = row + 1
            Wend
        End If
        
        ' Clear attributes area
        ClearOutputArea m_obj_ws, RGB(128, 128, 128)
        row = startRow
        AutoFit_Click
        'KillPleaseWait
    End If
    ' === build body parm END ===
    
    DoppioUI.UI_ShowPleaseWait "Please Wait... Calling API"
    startTime = Timer
    
    ' Get the values from the current row on the Data sheet.
    While row < (startRow + recordCount)
        Dim value As Variant
        Dim field As String
        Dim fieldValue As String
        Dim firstParam As Boolean
        Dim body As String
        Dim mi_url As String
        
        mi_url = strTransaction & "?"
        firstParam = True
        lastColumn = m_obj_ws.Cells(8, m_obj_ws.columns.count).End(xlToLeft).column
        
        For column = 2 To lastColumn
            field = strInputFields(column - 1)
            value = m_obj_ws.Cells(row, column)
            If isError(value) Then
                value = ""
            End If
            fieldValue = value
            
            ' add key to url
            If field = urlKey Then
                mi_url = Replace(mi_url, "{" & urlKey & "}", fieldValue)
            
            ' Check if field is "body"
            ElseIf field = "body" Then
                body = fieldValue
            
            ' Append field as a URL parameter if value is not empty
            ElseIf fieldValue <> "" Then
                If Not firstParam Then
                    mi_url = mi_url & "&"
                End If
                mi_url = mi_url & IDM_URLEncode(field) & "=" & IDM_URLEncode(fieldValue)
                firstParam = False
            Else
                Exit For
            End If
        Next column
        
        'mi_url = temp_url
        
        ' === call REST API
        apicall_Bridge m_s_MainUrl, "/IDM/api", mi_url, body, "IDM"
        IDM_Load_Results lastColumn + 1, row
            
        row = row + 1
    Wend
    
    AutoFit_Click
    KillPleaseWait
    maxbulk = maxbulk_hold
    
    ' update timer
    DisplayElapsedTime startTime, m_obj_ws

End Sub

Sub IDM_Process_JSON_Parameters(jsonParameters As String, startCol As Integer, startRow As Integer, row As Integer)
    Dim jsonObject As Object
    Dim jsonArray As Object
    Dim attr As Object
    Dim paramName As String
    Dim paramValue As String
    Dim columnIndex As Integer
    
    ' Set your worksheet here
    Set m_obj_ws = ActiveSheet
    
    ' Parse the JSON string into a usable object
    Set jsonObject = JsonConverter.ParseJson(jsonParameters)
    
    ' Access the "attr" array
    Set jsonArray = jsonObject.item("attr")
    
    ' Start inserting data from the specified start column
    columnIndex = startCol
    
    ' Loop through each item in the "attr" array
    For Each attr In jsonArray
        paramName = attr.item("name")
        
        ' Check if "value" exists, otherwise leave it blank
        If attr.exists("value") Then
            paramValue = attr.item("value")
        Else
            paramValue = ""
        End If
        
        ' Place the name in row 2 and value in row 3
        If startRow = row - 1 Then
            m_obj_ws.Cells(startRow, columnIndex).value = paramName
            m_obj_ws.Cells(row, columnIndex).value = paramValue
            columnIndex = columnIndex + 1
        Else
            Dim paramColumnIndex As Long
            Dim searchRow As Long: searchRow = 8 ' Row where paramName is located
            Dim foundCell As Range
        
            ' Find the column index where paramName is located in row 8
            With m_obj_ws.Rows(searchRow)
                Set foundCell = .Find(What:=paramName, LookIn:=xlValues, LookAt:=xlWhole)
                If Not foundCell Is Nothing Then
                    columnIndex = foundCell.column
                    m_obj_ws.Cells(row, columnIndex).value = paramValue
                Else
                    MsgBox "Parameter name not found in row " & searchRow, vbExclamation
                End If
            End With
        End If
    Next attr

End Sub

Function IDM_URLEncode(str As String) As String
    Dim i As Integer
    Dim ch As String
    Dim encodedStr As String
    
    encodedStr = ""
    
    ' Loop through each character in the input string
    For i = 1 To Len(str)
        ch = Mid(str, i, 1)
        
        Select Case ch
        Case " "
            encodedStr = encodedStr & "%20"
        Case "$"
            encodedStr = encodedStr & "%24"
        Case "/"
            encodedStr = encodedStr & "%2F"
        Case "["
            encodedStr = encodedStr & "%5B"
        Case "]"
            encodedStr = encodedStr & "%5D"
        Case "@"
            encodedStr = encodedStr & "%40"
        Case "="
            encodedStr = encodedStr & "%3D"
        Case ">"
            encodedStr = encodedStr & "%3E"
        Case "<"
            encodedStr = encodedStr & "%3C"
        Case """"
            encodedStr = encodedStr & "%22"
            'Case "("
            'encodedStr = encodedStr & "%28"
            'Case ")"
            'encodedStr = encodedStr & "%29"
        Case "&"
            encodedStr = encodedStr & "%26"
        Case ":"
            encodedStr = encodedStr & "%3A"
        Case Else
            ' Append the character as is if it doesn't need encoding
            encodedStr = encodedStr & ch
        End Select
    Next i
    
    ' Return the fully encoded string
    IDM_URLEncode = encodedStr
End Function

Function IsAlpha(char As String) As Boolean
    IsAlpha = (char Like "[A-Z']")
End Function

Function IsSheetVisible(sheetName As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(sheetName)
    On Error GoTo 0
    If ws Is Nothing Then
        IsSheetVisible = False
    Else
        IsSheetVisible = (ws.Visible = xlSheetVisible)
    End If
End Function

Public Sub KillPleaseWait()
    On Error GoTo ErrorHandler
    Application.ScreenUpdating = True
    DoEvents
    m_obj_ws.Shapes("Wait").Delete
    Exit Sub

ErrorHandler:
    ' Handle any errors that occurred
    ' For example, If the shape "Wait" does Not exist, this can be logged Or ignored As needed
    If Err.Number <> 0 Then
        ' Log the error Or show a message box
        ' MsgBox "An error occurred: " & Err.Description
    End If
    On Error GoTo 0
End Sub

' =============================================================================
' MODERNIZED LoadSourceAPIFormats
' =============================================================================
' Uses DoppioHttp for API calls
' =============================================================================

Public Sub LoadSourceAPIFormats_New()
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet, resultSheet As Worksheet
    Dim rowNum As Long, lastRow As Long
    Dim last_minm As String
    Dim record As Object, minm As String, mids As String
    Dim sortRange As Range
    Dim totalRecords As Long
    Dim i As Long
    Dim dataArray() As Variant
    Dim outputRow As Long
    Dim response As apiResponse
    Dim swaggerList As Object
    Dim rng As Range
    
    Set ws = ActiveSheet
    
    ' Clear old values
    Set resultSheet = ThisWorkbook.Sheets("AvailableMIs")
    rowNum = 9
    lastRow = resultSheet.Cells(resultSheet.Rows.count, "A").End(xlUp).row
    If lastRow > 8 Then
        resultSheet.Range("A9:J" & lastRow).ClearContents
    End If

    DoppioUI.UI_ShowPleaseWait "Loading Source API Formats"
        
    ' Force connect to source environment
    Doppio.Tenant_Token
    
    ' ==========================================
    ' Load APIs (MRS001MI/LstPrograms)
    ' ==========================================
    response = ExecuteMICall_Load("MRS001MI/LstPrograms;maxrecs=0?&returncols=MINM,MIDS")
    
    If response.success And Not response.records Is Nothing Then
        rowNum = 9
        totalRecords = response.records.count
        If totalRecords > 0 Then
            ReDim dataArray(1 To totalRecords, 1 To 2)
            i = 1
            For Each record In response.records
                dataArray(i, 1) = record.item("MINM")
                dataArray(i, 2) = record.item("MIDS")
                i = i + 1
            Next record
            resultSheet.Range(resultSheet.Cells(rowNum, 1), resultSheet.Cells(rowNum + totalRecords - 1, 1)).value = Application.index(dataArray, 0, 1)
            resultSheet.Range(resultSheet.Cells(rowNum, 5), resultSheet.Cells(rowNum + totalRecords - 1, 5)).value = Application.index(dataArray, 0, 2)
        End If
    End If
    
    ' ==========================================
    ' Load Foundation MT (/M3/foundation-rest/service/execute/ionapi-doc)
    ' ==========================================
    response = ExecuteSwaggerCall_Load("/M3/foundation-rest/service/execute/ionapi-doc/?&pageSize=1000")
    
    If response.success And Not response.results Is Nothing Then
        On Error Resume Next
        Set swaggerList = response.results.item("swaggerCollection").item("swagger")
        On Error GoTo ErrorHandler
        
        If Not swaggerList Is Nothing Then
            rowNum = 9
            totalRecords = swaggerList.count
            If totalRecords > 0 Then
                ReDim dataArray(1 To totalRecords, 1 To 2)
                i = 1
                For Each record In swaggerList
                    dataArray(i, 1) = record.item("entity")
                    dataArray(i, 2) = record.item("desc")
                    i = i + 1
                Next record
                resultSheet.Range(resultSheet.Cells(rowNum, 9), resultSheet.Cells(rowNum + totalRecords - 1, 9)).value = Application.index(dataArray, 0, 1)
                resultSheet.Range(resultSheet.Cells(rowNum, 10), resultSheet.Cells(rowNum + totalRecords - 1, 10)).value = Application.index(dataArray, 0, 2)
            End If
        End If
    End If
    
    ' ==========================================
    ' Load IPS (/M3/ips/service/ionapi-doc)
    ' ==========================================
    response = ExecuteSwaggerCall_Load("/M3/ips/service/ionapi-doc/?&pageSize=1000")
    
    If response.success And Not response.results Is Nothing Then
        On Error Resume Next
        Set swaggerList = response.results.item("swaggerCollection").item("swagger")
        On Error GoTo ErrorHandler
        
        If Not swaggerList Is Nothing Then
            rowNum = 9
            outputRow = 0
            last_minm = ""
            totalRecords = swaggerList.count
            
            If totalRecords > 0 Then
                ReDim dataArray(1 To totalRecords, 1 To 2)
                
                For Each record In swaggerList
                    minm = record.item("entity")
                    mids = record.item("desc")
                    
                    If InStr(minm, "#") > 0 Then minm = Left(minm, InStr(minm, "#") - 1)
                    If InStr(mids, "-") > 0 Then mids = Mid(mids, InStr(mids, "-") + 1)
                    
                    If last_minm <> minm Then
                        last_minm = minm
                        outputRow = outputRow + 1
                        dataArray(outputRow, 1) = minm
                        dataArray(outputRow, 2) = mids
                    End If
                Next record
                
                If outputRow > 0 Then
                    resultSheet.Range(resultSheet.Cells(rowNum, 2), resultSheet.Cells(rowNum + outputRow - 1, 2)).value = Application.index(dataArray, 0, 1)
                    resultSheet.Range(resultSheet.Cells(rowNum, 6), resultSheet.Cells(rowNum + outputRow - 1, 6)).value = Application.index(dataArray, 0, 2)
                End If
            End If
        End If
    End If
    
    ' ==========================================
    ' Load IDM (/IDM/api/ionapi-doc)
    ' ==========================================
    response = ExecuteSwaggerCall_Load("/IDM/api/ionapi-doc")
    
    If response.success And Not response.results Is Nothing Then
        ' Set the response for IDM_Load_Tags to use
        Set Doppio.m_obj_JsonResponse = response.results
        Doppio.IDM_Load_Tags
    End If
    
    ' ==========================================
    ' Load XtendM3 (static entry)
    ' ==========================================
    resultSheet.Cells(9, 4).value = "Extensibility"
    resultSheet.Cells(9, 8).value = "Infor XtendM3"
    
    ' Sort results
    Set sortRange = resultSheet.Range("A9:B" & resultSheet.Cells(resultSheet.Rows.count, 1).End(xlUp).row)
    sortRange.Sort Key1:=sortRange.columns(1), Order1:=xlAscending, Header:=xlYes
    sortRange.columns.AutoFit
    
    ' Check if "API" field exists on the panel
    On Error Resume Next
    Set rng = ws.Range("API")
    On Error GoTo ErrorHandler
    
    If Not rng Is Nothing Then
        ws.Cells(2, 2).value = "API"
        If ws.Range("API").value = "" Then
            ws.Range("API").value = "CRS111MI"
        End If
        GetTransactions_Click_New
        Doppio.Settings_CopyDefaults
    End If
    
    DoppioUI.UI_KillPleaseWait
    Exit Sub
    
ErrorHandler:
    DoppioUI.UI_KillPleaseWait
    #If DEBUG_MODE Then
        Debug.Print "LoadSourceAPIFormats_New: ERROR - " & Err.description
    #End If
End Sub

' =============================================================================
' HELPER: Execute MI API call (returns records)
' =============================================================================

Private Function ExecuteMICall_Load(endpoint As String) As apiResponse
    Dim config As httpConfig
    Dim httpResponse As httpResponse
    Dim response As apiResponse
    Dim apiUrl As String
    Dim json As Object
    
    On Error GoTo ErrorHandler
    
    ' Build URL for MI call
    apiUrl = Doppio.m_s_MainUrl & "/M3/m3api-rest/v2/execute/" & endpoint
    
    ' Configure request
    config.url = apiUrl
    config.method = HttpMethod_GET
    config.contentType = "application/json"
    config.AcceptType = "application/json"
    config.authHeader = Doppio.m_s_TokenType & " " & Doppio.m_s_AccessToken
    config.TimeoutSeconds = 30
    config.body = ""
    
    #If DEBUG_MODE Then
        Debug.Print "ExecuteMICall_Load: " & apiUrl
    #End If
    
    ' Execute
    httpResponse = DoppioHttp.ExecuteRequest(config)
    
    ' Build response
    response.success = httpResponse.success
    response.data = httpResponse.body
    response.errorMessage = httpResponse.errorMessage
    
    ' Parse JSON - MI response has "results" array
    If httpResponse.success And Len(httpResponse.body) > 0 Then
        Set json = JsonConverter.ParseJson(httpResponse.body)
        
        If Not json Is Nothing Then
            If json.exists("results") Then
                Set response.results = json.item("results")
                If response.results.count > 0 Then
                    Set response.records = response.results(1).item("records")
                    If Not response.records Is Nothing Then
                        response.recordCount = response.records.count
                    End If
                End If
            End If
        End If
    End If
    
    ExecuteMICall_Load = response
    Exit Function
    
ErrorHandler:
    #If DEBUG_MODE Then
        Debug.Print "ExecuteMICall_Load: ERROR - " & Err.description
    #End If
    response.success = False
    response.errorMessage = Err.description
    ExecuteMICall_Load = response
End Function

' =============================================================================
' HELPER: Execute Swagger API call (returns swagger collection)
' =============================================================================

Private Function ExecuteSwaggerCall_Load(endpoint As String) As apiResponse
    Dim config As httpConfig
    Dim httpResponse As httpResponse
    Dim response As apiResponse
    Dim apiUrl As String
    Dim json As Object
    
    On Error GoTo ErrorHandler
    
    ' Build URL for Swagger call
    apiUrl = Doppio.m_s_MainUrl & endpoint
    
    ' Configure request
    config.url = apiUrl
    config.method = HttpMethod_GET
    config.contentType = "application/json"
    config.AcceptType = "application/json"
    config.authHeader = Doppio.m_s_TokenType & " " & Doppio.m_s_AccessToken
    config.TimeoutSeconds = 30
    config.body = ""
    
    #If DEBUG_MODE Then
        Debug.Print "ExecuteSwaggerCall_Load: " & apiUrl
    #End If
    
    ' Execute
    httpResponse = DoppioHttp.ExecuteRequest(config)
    
    ' Build response
    response.success = httpResponse.success
    response.data = httpResponse.body
    response.errorMessage = httpResponse.errorMessage
    
    ' Parse JSON - Swagger response is the raw JSON object
    If httpResponse.success And Len(httpResponse.body) > 0 Then
        Set json = JsonConverter.ParseJson(httpResponse.body)
        Set response.results = json
    End If
    
    ExecuteSwaggerCall_Load = response
    Exit Function
    
ErrorHandler:
    #If DEBUG_MODE Then
        Debug.Print "ExecuteSwaggerCall_Load: ERROR - " & Err.description
    #End If
    response.success = False
    response.errorMessage = Err.description
    ExecuteSwaggerCall_Load = response
End Function

' =============================================================================
' LEGACY WRAPPER
' =============================================================================

Public Sub LoadSourceAPIFormats()
    LoadSourceAPIFormats_New
End Sub


Sub LogError(ByVal data As String)
    Dim json As Object
    Dim wasTerminated As Boolean
    Dim nrOfSuccessfullTransactions As Long
    Dim nrOfFailedTransactions As Long
    Dim nextRow As Long
    Dim currentTime As String
    Dim logsheet As Worksheet
    
    ' Set the log sheet
    On Error Resume Next
    Set logsheet = ThisWorkbook.Sheets("Log")
    On Error GoTo 0
    If logsheet Is Nothing Then
        ' If the log sheet doesn't exist, create it
        'Set logSheet = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        'logSheet.name = "Log"
        ' Add headers
        logsheet.Cells(1, 1).value = "Timestamp"
        logsheet.Cells(1, 2).value = "Error Number"
        logsheet.Cells(1, 3).value = "Error Description"
        logsheet.Cells(1, 4).value = "Data"
        logsheet.Cells(1, 5).value = "wasTerminated"
        logsheet.Cells(1, 6).value = "nrOfSuccessfullTransactions"
        logsheet.Cells(1, 7).value = "nrOfFailedTransactions"
        logsheet.Cells(1, 8).value = "curlCommand"
        logsheet.Cells(1, 9).value = "curlResult"
    End If

    ' Find the Next empty row
    nextRow = logsheet.Cells(logsheet.Rows.count, 1).End(xlUp).row + 1

    ' Get the current time
    currentTime = Format(Now, "yyyy-mm-dd hh:mm:ss")

    ' Try To parse the JSON string
    On Error GoTo JsonError
    Set json = JsonConverter.ParseJson(m_s_CurlResult)
    On Error GoTo 0

    ' Extract the required fields
    wasTerminated = json.item("wasTerminated")
    nrOfSuccessfullTransactions = json.item("nrOfSuccessfullTransactions")
    nrOfFailedTransactions = json.item("nrOfFailedTransactions")

    ' Log the error
    If (Err.Number <> m_l_PrevErrorNumber Or data <> m_s_PrevData) Then
        logsheet.Cells(nextRow, 1).value = currentTime
        logsheet.Cells(nextRow, 2).value = Err.Number
        logsheet.Cells(nextRow, 3).value = Err.description
        logsheet.Cells(nextRow, 4).value = data
        logsheet.Cells(nextRow, 5).value = wasTerminated
        logsheet.Cells(nextRow, 6).value = nrOfSuccessfullTransactions
        logsheet.Cells(nextRow, 7).value = nrOfFailedTransactions
        logsheet.Cells(nextRow, 8).value = curlCommand
        If nrOfSuccessfullTransactions = 0 Then
            logsheet.Cells(nextRow, 9).value = m_s_CurlResult
        End If
        m_l_PrevErrorNumber = Err.Number
        m_s_PrevData = data
    End If

    Exit Sub

JsonError:
    ' Handle JSON parsing error
    #If DEBUG_MODE Then
        Debug.Print "Error parsing JSON: " & Err.description
    #End If
    logsheet.Cells(nextRow, 1).value = currentTime
    logsheet.Cells(nextRow, 2).value = Err.Number
    logsheet.Cells(nextRow, 3).value = "JSON Parsing Error: " & Err.description
    logsheet.Cells(nextRow, 4).value = data
    logsheet.Cells(nextRow, 8).value = curlCommand
    logsheet.Cells(nextRow, 9).value = m_s_CurlResult
    m_l_PrevErrorNumber = Err.Number
    m_s_PrevData = data
    On Error GoTo 0
End Sub

' =============================================================================
' MODERNIZED Log_Activity
' =============================================================================
' Uses DoppioHttp for API calls
' Cross-platform compatible (Mac/Windows)
' =============================================================================

Public Sub Log_Activity_New()
    On Error GoTo ErrorHandler
    
    Dim jsonString As String
    Dim json As Object
    Dim userName As String, machineName As String, fullUserName As String
    Dim encodedUserName As String, encodedMachineName As String
    Dim encodedSelectedEnvironment As String
    Dim mainUrl As String, miPath As String, miUrl As String
    Dim tokenEndpoint As String, ssoBase As String
    Dim clientId As String, clientSecret As String
    Dim encTenant As String, urlSelectedEnvironment As String
    Dim envSheet As Worksheet
    
    ' Save current token (we use a different one for logging)
    Dim saveToken As String, saveType As String
    saveType = Doppio.m_s_TokenType
    saveToken = Doppio.m_s_AccessToken
    
    ' Decode and parse config
    jsonString = "eyJ0aSI6IkRPUFBJT19ERU0iLCJjbiI6Ik0zIERhdGEgTG9hZGVyIiwiZHQiOiIxMiIsImNpIjoiRE9QUElPX0RFTX5abS1tNkRMZTlCWHBuWDM0SVVfQ3RDaGdYX2R2b3lCUXFTQ2hOR2JSUm1NIiwiY3MiOiJacHN0MXhpbFNHTEJrdGFyemNWZFdSbkEyb2tNbHNCTUdjOW1wUTdoZE1RZW12ZDdYR1RucFFicWVXTVpNMHBUb3BXcE9FQmIwUFlaMnVMbUxRUjY3ZyIsIml1IjoiaHR0cHM6Ly9taW5nbGUtaW9uYXBpLmluZm9yY2xvdWRzdWl0ZS5jb20iLCJwdSI6Imh0dHBzOi8vbWluZ2xlLXNzby5pbmZvcmNsb3Vkc3VpdGUuY29tOjQ0My9ET1BQSU9fREVNL2FzLyIsIm9hIjoiYXV0aG9yaXphdGlvbi5vYXV0aDIiLCJvdCI6InRva2VuLm9hdXRoMiIsIm9yIjoicmV2b2tlX3Rva2VuLm9hdXRoMiIsImV2IjoiVTE0NzgzNTgxMDEiLCJ2IjoiMS4wIiwic2FhayI6IkRPUFBJT19ERU0jek9lYkI0eGwxZGlFeTIxaWtNUXNJSzFfaW1sZjdycFBuRGEyYTZxOEtkdFE1Vy1hZDhHY2o2NEd3OVVILUplSU9mbGdWenJuZlRKLVdJSDFZYlU4MlEiLCJzYXNrIjoiMzhKcElrcTVIWUhEbGhwcUVuT040cDIyQ2o2NEVkNVdwbE5zaVNjMXRhRUd4NFV3aGZRblRGMnZYUDhha1R5Ri12S3llSWFocjlwX1RzM3NubmJIU2cifQo"
    jsonString = Doppio.Base64DecodeVBA(jsonString)
    Set json = JsonConverter.ParseJson(jsonString)
    
    ' Extract from JSON config
    ssoBase = json.item("pu")
    tokenEndpoint = json.item("ot")
    clientId = json.item("ci")
    clientSecret = json.item("cs")
    
    mainUrl = json.item("iu") & "/" & json.item("ti")
    miPath = "/M3/m3api-rest/v2/execute"
    
    urlSelectedEnvironment = UrlEncode_Log(Doppio.m_s_SelectedEnvironment)
    encTenant = Doppio.Base64EncodeVBA(Doppio.encodedTenant)
    
    ' Get user/machine info (cross-platform)
    GetUserInfo userName, fullUserName, machineName
    
    encodedUserName = UrlEncode_Log(userName)
    encodedMachineName = UrlEncode_Log(machineName)
    encodedSelectedEnvironment = UrlEncode_Log(Doppio.m_s_SelectedEnvironment)
    
    ' Get token for logging service
    If Not GetLogActivityToken(ssoBase & tokenEndpoint, clientId, clientSecret, json.item("saak"), json.item("sask")) Then
        GoTo Cleanup
    End If
    
    ' Update user information
    miUrl = mainUrl & miPath & "/EXT123MI/UpdUsrInfo?" & _
            "M3NM=" & encodedSelectedEnvironment & _
            "&M3ID=" & Doppio.m_s_M3user & _
            "&PCNM=" & encodedMachineName & _
            "&PCID=" & encodedUserName & _
            "&VERS=" & DOPPIO_VERSION & _
            "&HASH=" & encTenant & _
            "&TNAL=" & urlSelectedEnvironment
    
    ExecuteLogCall miUrl
    
    ' Add user information
    miUrl = mainUrl & miPath & "/EXT123MI/AddUsrInfo?" & _
            "M3NM=" & encodedSelectedEnvironment & _
            "&M3ID=" & Doppio.m_s_M3user & _
            "&PCNM=" & encodedMachineName & _
            "&PCID=" & encodedUserName & _
            "&AUTH=1" & _
            "&VERS=" & DOPPIO_VERSION & _
            "&TNAL=" & urlSelectedEnvironment & _
            "&HASH=" & encTenant
    
    ExecuteLogCall miUrl
    
Cleanup:
    ' Restore original token
    Doppio.m_s_TokenType = saveType
    Doppio.m_s_AccessToken = saveToken
    Exit Sub
    
ErrorHandler:
    #If DEBUG_MODE Then
        Debug.Print "Log_Activity_New: ERROR - " & Err.description
    #End If
    ' Restore original token even on error
    Doppio.m_s_TokenType = saveType
    Doppio.m_s_AccessToken = saveToken
End Sub

' =============================================================================
' HELPER: Get user/machine info (cross-platform)
' =============================================================================

Private Sub GetUserInfo(ByRef userName As String, ByRef fullUserName As String, ByRef machineName As String)
    On Error Resume Next
    
    #If Mac Then
        userName = MacScript("do shell script ""whoami""")
        fullUserName = MacScript("do shell script ""id -F""")
        machineName = MacScript("do shell script ""scutil --get ComputerName""")
    #Else
        userName = Environ("USERNAME")
        fullUserName = Environ("USERDOMAIN")
        machineName = Environ("COMPUTERNAME")
    #End If
    
    On Error GoTo 0
End Sub

' =============================================================================
' HELPER: Get token for log activity service
' =============================================================================

Private Function GetLogActivityToken(tokenUrl As String, clientId As String, clientSecret As String, saak As String, sask As String) As Boolean
    Dim config As httpConfig
    Dim response As httpResponse
    Dim json As Object
    Dim body As String
    
    On Error GoTo ErrorHandler
    
    ' Build form body
    body = "client_id=" & clientId & _
           "&client_secret=" & clientSecret & _
           "&grant_type=password" & _
           "&username=" & saak & _
           "&password=" & sask
    
    ' Configure request
    config.url = tokenUrl
    config.method = HttpMethod_POST
    config.contentType = "application/x-www-form-urlencoded"
    config.AcceptType = "application/json"
    config.authHeader = ""
    config.TimeoutSeconds = 5
    config.body = body
    
    #If DEBUG_MODE Then
        Debug.Print "GetLogActivityToken: URL = " & tokenUrl
    #End If
    
    ' Execute request
    response = DoppioHttp.ExecuteRequest(config)
    
    If response.success And Len(response.body) > 0 Then
        Set json = JsonConverter.ParseJson(response.body)
        If Not json Is Nothing Then
            Doppio.m_s_AccessToken = json.item("access_token")
            Doppio.m_s_TokenType = json.item("token_type")
            GetLogActivityToken = True
            #If DEBUG_MODE Then
                Debug.Print "GetLogActivityToken: Success"
            #End If
            Exit Function
        End If
    End If
    
    #If DEBUG_MODE Then
        Debug.Print "GetLogActivityToken: Failed - " & response.errorMessage
    #End If
    GetLogActivityToken = False
    Exit Function
    
ErrorHandler:
    #If DEBUG_MODE Then
        Debug.Print "GetLogActivityToken: ERROR - " & Err.description
    #End If
    GetLogActivityToken = False
End Function

' =============================================================================
' HELPER: Execute log API call
' =============================================================================

Private Sub ExecuteLogCall(url As String)
    Dim config As httpConfig
    Dim response As httpResponse
    
    On Error GoTo ErrorHandler
    
    ' Configure request
    config.url = url
    config.method = HttpMethod_GET
    config.contentType = "application/json"
    config.AcceptType = "application/json"
    config.authHeader = Doppio.m_s_TokenType & " " & Doppio.m_s_AccessToken
    config.TimeoutSeconds = 5
    config.body = ""
    
    #If DEBUG_MODE Then
        Debug.Print "ExecuteLogCall: " & Left(url, 100) & "..."
    #End If
    
    ' Execute request (fire and forget - don't care about response)
    response = DoppioHttp.ExecuteRequest(config)
    
    #If DEBUG_MODE Then
        Debug.Print "ExecuteLogCall: Status = " & response.statusCode
    #End If
    
    Exit Sub
    
ErrorHandler:
    #If DEBUG_MODE Then
        Debug.Print "ExecuteLogCall: ERROR - " & Err.description
    #End If
End Sub

' =============================================================================
' HELPER: URL Encode (local version to avoid dependency)
' =============================================================================

Private Function UrlEncode_Log(text As String) As String
    Dim i As Long
    Dim c As String
    Dim result As String
    
    result = ""
    For i = 1 To Len(text)
        c = Mid(text, i, 1)
        Select Case c
            Case "A" To "Z", "a" To "z", "0" To "9", "-", "_", ".", "~"
                result = result & c
            Case " "
                result = result & "%20"
            Case Else
                result = result & "%" & Right("0" & Hex(Asc(c)), 2)
        End Select
    Next i
    
    UrlEncode_Log = result
End Function

' =============================================================================
' LEGACY WRAPPER
' =============================================================================

Public Sub Log_Activity()
    Log_Activity_New
End Sub



Function MacBase64Encode(inputString As String) As String
    ' Escape special characters in the input string
    inputString = Replace(inputString, "\", "\\")

    ' Construct the AppleScript command
    Dim appleScriptCmd As String
    appleScriptCmd = "Do shell script ""printf %s " & inputString & " | base64"""
    'Debug.Print appleScriptCmd

    ' Execute the AppleScript command
    MacBase64Encode = MacScript(appleScriptCmd)
End Function

Sub MoveToTopOfFrozenSectionOnActiveSheet()
    Dim ws As Worksheet
    Set ws = ActiveSheet
    ws.Activate
    ActiveWindow.FreezePanes = False
    ws.Cells(9, 3).Activate
    ActiveWindow.FreezePanes = True
End Sub

Sub Output_GetUserInfo()
    Dim ws As Worksheet
    Dim record As Object
    
    Set ws = ActiveSheet
    
    Application.ScreenUpdating = False
    
    For Each record In m_obj_Records
        If record.item("ZZUSID") = "" Then
            ws.Cells(3, 9).value = m_s_M3user
        Else
            ws.Cells(3, 9).value = m_s_M3user
        End If
        If ws.Cells(4, 9) = "" Then
            ws.Cells(4, 9).value = record.item("ZDCONO")
            ws.Cells(5, 9).value = record.item("ZDDIVI")
        End If
    Next record
    
    Application.ScreenUpdating = True

    m_s_M3user = ws.Range("User").value
    m_s_Company = ws.Range("Company").value
    m_s_Division = ws.Range("Division").value

End Sub

Sub Output_LstTransactions()
    Dim ws As Worksheet
    Dim record As Object
    
    Set ws = ThisWorkbook.Sheets("Transactions")

    ' Clear existing data
    ws.UsedRange.Clear

    ' Check If the Records collection is Not empty
    If Not (m_obj_Records Is Nothing) Then
        Dim rowNum As Long
        rowNum = 3

        ' Loop through the Records collection
        For Each record In m_obj_Records
            ws.Cells(rowNum, 1).value = record.item("MINM")
            ws.Cells(rowNum, 2).value = record.item("TRNM")
            ws.Cells(rowNum, 3).value = record.item("SIMU")
            rowNum = rowNum + 1
        Next record
    End If

End Sub

Sub Parse_Bulk_Results(results As Object)
    Dim message, nrOfTransactions, errorMessage, nrOfSuccessfullTransactions As String
    Dim totalRows As Long
    Dim outputRange As Range
    Dim result As Object
    Dim outputRow As Long
    Dim rowNum As Long
    Dim lastCol As Integer
    Dim LastRunTime As Single
    
    If Not m_obj_JsonResponse Is Nothing Then
        nrOfTransactions = m_obj_JsonResponse.item("nrOfSuccessfullTransactions") + m_obj_JsonResponse.item("nrOfFailedTransactions")
    Else
        Exit Sub
    End If
    
    ' Process records in the result
    rowNum = Application.Max(m_l_Row - maxbulk, 9)

    ' Determine total rows for array storage
    If Not m_obj_Results Is Nothing Then
        'totalRows = 1
        For Each result In m_obj_Results
            Set m_obj_Records = result.item("records")
            If Not m_obj_Records Is Nothing Then
                totalRows = m_obj_Records.count
            End If
        Next result
    End If
    totalRows = Application.Max(nrOfTransactions, totalRows)
    'Debug.Print "totalRows: " & totalRows

    ' Define the size of the output array
    lastCol = m_obj_ws.Cells(8, m_obj_ws.columns.count).End(xlToLeft).column
    ReDim outputData(1 To totalRows, 1 To lastCol)
    outputRow = 1
    'Debug.Print "lastCol: " & lastCol
    
    If Not m_obj_Results Is Nothing Then
        For Each result In m_obj_Results
            nrOfSuccessfullTransactions = m_obj_JsonResponse.item("nrOfSuccessfullTransactions")
            
            message = Trim(Left(result.item("errorMessage"), 240))
            errorMessage = "OK"
            If message <> "" Then
                errorMessage = "NOK " & message
            End If
            
            ' Process parameters
            Call ProcessParameters(result, outputData, outputRow)
                
            ' Add the errorMessage to outputData (Column A)
            outputData(outputRow, 1) = errorMessage
                
            Set m_obj_Records = result.item("records")

            If Not m_obj_Records Is Nothing Then
                Dim recordCount As Long
                recordCount = m_obj_Records.count

                If recordCount > 100000 Then
                    Application.ScreenUpdating = True
                    If Not PromptUser("There are " & recordCount & " records To process, continue?") Then
                        KillPleaseWait
                        Exit Sub
                    End If
                    Application.Calculate
                    DoEvents
                    Application.ScreenUpdating = False
                End If

                Dim record As Object
                LastRunTime = Now

                ' ==================================================
                ' Use a Collection instead of Scripting.Dictionary for Mac compatibility
                Dim columnCache As Collection
                Set columnCache = New Collection
            
                Dim keyCache As Collection
                Set keyCache = New Collection
            
                ' Check if there are multiple records
                Dim nrOfRecords As Long
                nrOfRecords = m_obj_Records.count
                Dim i As Integer
                Dim value As Variant

                ' load data from resuls/records
                If IsRecordsEmpty(m_obj_Records) Then
                    For i = 2 To lastCol
                        value = Cells((rowNum + outputRow) - 1, i).value
                        outputData(outputRow, i) = value
                    Next i
                Else
                    For Each record In m_obj_Records
                        Dim key As Variant
                        Dim columnIndex As Integer
                        Dim isFound As Boolean
                        
                        For Each key In record.keys
                            If outputData(outputRow, 1) = "" Then
                                outputData(outputRow, 1) = "OK"
                            End If
                    
                            ' Check if column index for this key has already been cached
                            isFound = False
                            For i = 1 To columnCache.count
                                If keyCache(i) = key Then
                                    columnIndex = columnCache(i)
                                    isFound = True
                                    Exit For
                                End If
                            Next i
                            'columnIndex = i
                            
                            If Not isFound Then
                                ' Find column index and cache it for future use
                                columnIndex = FindColumnIndex(key)
                                columnCache.Add columnIndex
                                keyCache.Add key
                            End If
                        
                            ' Only populate if valid column index found
                            If columnIndex > 0 Then
                                value = record.item(key)
                                value = Replace(value, "??", "")
                                ' Write value to outputData array if not empty
                                'If value <> "" Then
                                'Debug.Print "outputRow: " & outputRow & " columnIndex:" & columnIndex
                                outputData(outputRow, columnIndex) = value
                                'End If
                            End If
                        Next key
                
                        ' Only increment outputRow if the transaction is successful
                        If nrOfSuccessfullTransactions = 1 And nrOfRecords > 1 Then
                            outputRow = outputRow + 1
                        End If
                    Next record
                End If
                ' ==================================================

            End If
            'rowNum = rowNum + 1
            outputRow = outputRow + 1
        Next result
    End If
    
    ' Define output range and write array to sheet at once
    'Debug.Print "rowNum: " & rowNum & "  totalRows: " & totalRows & "  lastCol: " & lastCol
    Set outputRange = m_obj_ws.Cells(rowNum, 1).Resize(totalRows, lastCol)
    outputRange.value = outputData
    
    ' Loop through Column A in the outputRange to apply font colors
    Dim cell As Range
    For Each cell In outputRange.columns(1).Cells ' Loop through Column A
        If cell.value = "OK" Then
            cell.Font.Color = 40000              ' Apply font color for "OK"
        Else
            cell.Font.Color = COLOR_ERROR                ' Apply font color for other values
        End If
    Next cell
    
    'KillPleaseWait
    Exit Sub

    'ErrorHandler:
    '    MsgBox "An error occurred: " & Err.description
    '    KillPleaseWait
    '    On Error GoTo 0
End Sub
Private Function IsRecordsEmpty(records As Collection) As Boolean
    IsRecordsEmpty = False
    If m_obj_Records.count = 0 Then
        IsRecordsEmpty = True
    ElseIf m_obj_Records.count = 1 And m_obj_Records(1).count = 0 Then
        IsRecordsEmpty = True
    End If
End Function
Function Parse_Fault_String(xmlText) As String
    Dim xmlDoc As Object
    Dim ns As String

    Set xmlDoc = CreateObject("MSXML2.DOMDocument.6.0")

    ' Load the XML text
    xmlDoc.LoadXML xmlText

    If m_b_Multitenant Then
        ns = "xmlns:SOAP-ENV='http://schemas.xmlsoap.org/soap/envelope/'"
    Else
        ns = "xmlns:soap='http://schemas.xmlsoap.org/soap/envelope/'"
    End If
    xmlDoc.SetProperty "SelectionNamespaces", ns
    xmlDoc.SetProperty "SelectionLanguage", "XPath"

    ' Find the <faultstring> element
    Dim faultStringNode As Object
    Set faultStringNode = xmlDoc.SelectSingleNode("//faultstring")

    ' Extract the value of <faultstring>
    Dim faultString As String
    If Not faultStringNode Is Nothing Then
        faultString = faultStringNode.text
    End If

    Parse_Fault_String = faultString
End Function

Sub Parse_JSON_Response(ByVal json As Object)
    Dim results As Object
    Dim item As Object
    
    Set results = json.item("results")

    Dim ws As Worksheet
    Set ws = ActiveSheet

    Dim i As Long
    Dim errorMessage As String

    ' Start populating from Row 9
    i = 9

    For Each item In results
        If item("errorMessage") <> "" Then
            errorMessage = item("errorMessage")
        Else
            errorMessage = "OK"
        End If

        ' Output the errorMessage Or "OK" in Column A starting from Row 9
        ws.Cells(i, 1).value = errorMessage

        ' Increment the row counter
        i = i + 1
    Next item
End Sub

Sub Parse_M3X_Results()
    Dim errorMessage, fontColor As String

    errorMessage = "OK"
    fontColor = "40000"
    If m_s_CurlResult = "403" Then
        errorMessage = "NOK 403 Forbidden"
        fontColor = "255"
    End If
    m_obj_ws.Cells(m_l_Row, 1).value = errorMessage
    m_obj_ws.Cells(m_l_Row, 1).Font.Color = fontColor

End Sub

Sub Parse_Soap_Results(result As String)
    Dim errorMessage As String
    Dim fontColor As String
    
    errorMessage = Parse_Fault_String(result)
    If errorMessage <> "" Or (result = "" And m_s_MiUrl <> "MRS001MI/GetUserInfo/?") Then
        If errorMessage = "" Then
            errorMessage = "Unknown Error"
        End If
        errorMessage = "NOK " & errorMessage
        fontColor = "255"
    Else
        errorMessage = "OK"
        fontColor = "40000"
    End If
    m_obj_ws.Cells(m_l_Row, 1).value = errorMessage
    m_obj_ws.Cells(m_l_Row, 1).Font.Color = fontColor

    Dim xmlDoc As Object
    Set xmlDoc = CreateObject("MSXML2.DOMDocument.6.0")

    ' Load the XML content
    xmlDoc.LoadXML result

    ' Check If the XML was loaded successfully
    If xmlDoc.parseError <> 0 Then
        MsgBox "Error loading XML: " & xmlDoc.parseError.reason
        Exit Sub
    End If

    ' Get node name
    Dim ns As String
    Dim originalString As String
    Dim nodeName As String
    originalString = m_obj_ws.Range("API").value
    If Len(originalString) >= 3 Then
        If Mid(originalString, 3, 1) = "S" Then
            nodeName = Left(originalString, 6)
        Else
            nodeName = Left(originalString, 5)
        End If
    Else
        nodeName = originalString
    End If

    Dim resultNode As Object
    If m_b_Multitenant Then
        ns = "xmlns:chg='http://schemas.infor.com/ips/" & m_obj_ws.Range("API").value & "/" & m_obj_ws.Range("Transaction").value & "'"
        xmlDoc.SetProperty "SelectionNamespaces", ns
        xmlDoc.SetProperty "SelectionLanguage", "XPath"
        Set resultNode = xmlDoc.SelectSingleNode("//chg:" & nodeName)
    Else
        ns = "xmlns:ns='http://your.company.net/" & m_obj_ws.Range("API").value & "/" & m_obj_ws.Range("Transaction").value & "'"
        xmlDoc.SetProperty "SelectionNamespaces", ns
        xmlDoc.SetProperty "SelectionLanguage", "XPath"
        Set resultNode = xmlDoc.SelectSingleNode("//ns:" & nodeName)
    End If

    Dim colNum As Integer
    Dim childNode As Object
    Dim key, value As String
    If Not resultNode Is Nothing Then
        colNum = 2
        For Each childNode In resultNode.ChildNodes
            key = childNode.nodeName
            value = childNode.text

            Dim parts() As String
            parts = Split(key, ":")
            If UBound(parts) > 0 Then
                key = parts(UBound(parts))
            End If

            ' Find the column index based on the key in row 8
            Dim columnIndex, i As Integer
            columnIndex = 0
            For i = 2 To m_obj_ws.Cells(8, columns.count).End(xlToLeft).column
                If m_obj_ws.Cells(8, i).value = key Then
                    columnIndex = i
                    Exit For
                End If
            Next i
            'Debug.Print key & vbTab & columnIndex

            If columnIndex > 0 Then
                m_obj_ws.Cells(m_l_Row, 1).value = "OK"
                m_obj_ws.Cells(m_l_Row, 1).Font.Color = "40000"
                If value <> "" Then
                    m_obj_ws.Cells(m_l_Row, columnIndex).value = value
                End If
            End If

        Next childNode
    End If

End Sub

Sub Postman_Build()
    ' Declare variables to hold field values
    Dim url As String
    Dim startPos As Long
    Dim endPos As Long
    Dim dataPosition As Long
    Dim body As String
    Dim json As Object, infoObj As Object, authObj As Object, oauth2Array As Object
    Dim script As String
      
    Set m_obj_ws = ActiveSheet
    
    Doppio.Tenant_Information
  
    'Curl_Build m_s_MainUrl, m_s_MiPath, m_s_MiUrl, body, "API", script

    ' extract url
    startPos = InStr(curlCommand, "'")
    endPos = InStr(startPos + 1, curlCommand, "'")
    If startPos > 0 And endPos > startPos Then
        url = Mid(curlCommand, startPos + 1, endPos - startPos - 1)
    Else
        url = ""
    End If
    
    ' extract body variable
    dataPosition = InStr(curlCommand, "--data-raw")
    If dataPosition > 0 Then
        body = Trim(Mid(curlCommand, dataPosition + Len("--data-raw")))
        body = Replace(body, "'", "")
    Else
        body = ""
    End If

    ' Initialize main JSON object
    Set json = JsonConverter.ParseJson("{}")
    
    ' Create "info" object
    Set infoObj = JsonConverter.ParseJson("{}")
    infoObj.item("_postman_id") = "4cc39a68-a241-4334-ba6a-0abff2c7e49b"
    infoObj.item("name") = ti
    infoObj.item("schema") = "https://schema.getpostman.com/json/collection/v2.1.0/collection.json"
    infoObj.item("_exporter_id") = "4894281"
    infoObj.item("_collection_link") = "https://crimson-meadow-524739.postman.co/workspace/CTOS-Workspace~0511ca45-2c21-4156-ad9d-89da5e304070/collection/4894281-4cc39a68-a241-4334-ba6a-0abff2c7e49b?action=share&source=collection_link&creator=4894281"
    
    ' Add "info" object to main JSON
    '  json.item("info") = infoObj
    json.Add "info", infoObj
    
    ' Create "item" array and object
    Dim itemArray, itemObj As Object
    Set itemArray = JsonConverter.ParseJson("[]")
    Set itemObj = JsonConverter.ParseJson("{}")
    
    ' Populate "item" object
    itemObj.item("name") = m_obj_ws.Range("API").value & " " & m_obj_ws.Range("Transaction").value
    
    ' Create "request" object
    Dim requestObj As Object
    Set requestObj = JsonConverter.ParseJson("{}")
    requestObj.item("method") = "POST"
    
    ' Create "header" array
    Dim headerArray As Object, headerObj As Object
    Set headerArray = JsonConverter.ParseJson("[]")
    Set headerObj = JsonConverter.ParseJson("{}")
    headerObj.item("key") = "Content-Type"
    headerObj.item("value") = "application/json"
    headerArray.Add headerObj
    
    ' Add "header" to "request" object
    '  requestObj("header") = headerArray
    requestObj.Add "header", headerArray
    
    ' Create "body" object
    Dim bodyObj As Object, optionsObj As Object, rawObj As Object
    Set bodyObj = JsonConverter.ParseJson("{}")
    bodyObj.item("mode") = "raw"
    bodyObj.item("raw") = body
    
    ' Create "options" object for "body"
    Set optionsObj = JsonConverter.ParseJson("{}")
    Set rawObj = JsonConverter.ParseJson("{}")
    rawObj.item("language") = "json"
    '  optionsObj("raw") = rawObj
    optionsObj.Add "raw", rawObj
    '  bodyObj("options") = optionsObj
    bodyObj.Add "options", optionsObj
  
    ' Add "body" to "request" object
    '  requestObj("body") = bodyObj
    requestObj.Add "body", bodyObj
    
    ' Add parsed URL to the request object
    Dim parsedUrlObj As Object
    Set parsedUrlObj = Postman_ParseURL(url)
    '  requestObj("url") = parsedUrlObj("url")
    requestObj.Add "url", parsedUrlObj.item("url")
    
    ' Add "description"
    requestObj.item("description") = "this is the description"
    
    ' Add "request" to "item" object
    '  itemObj("request") = requestObj
    itemObj.Add "request", requestObj
    '  itemObj("response") = JsonConverter.ParseJson("[]")
    itemObj.Add "response", JsonConverter.ParseJson("[]")
    
    ' Add "item" object to "item" array
    itemArray.Add itemObj
    
    ' Add "item" array to main JSON
    '  json.item("item") = itemArray
    json.Add "item", itemArray
    
    ' Create "auth" object
    Set authObj = JsonConverter.ParseJson("{}")
    authObj.item("type") = "oauth2"
    
    ' Create "oauth2" array
    Set oauth2Array = JsonConverter.ParseJson("[]")
    
    ' Add OAuth2 fields
    oauth2Array.Add Postman_CreateOAuthField("password", sask)
    oauth2Array.Add Postman_CreateOAuthField("username", saak)
    oauth2Array.Add Postman_CreateOAuthField("accessTokenUrl", pu & ot)
    oauth2Array.Add Postman_CreateOAuthField("grant_type", "password_credentials")
    oauth2Array.Add Postman_CreateOAuthField("clientSecret", cs)
    oauth2Array.Add Postman_CreateOAuthField("clientId", ci)
    oauth2Array.Add Postman_CreateOAuthField("tokenName", ti)
    oauth2Array.Add Postman_CreateOAuthField("addTokenTo", "header")
    
    ' Add OAuth2 array to auth object
    '  authObj("oauth2") = oauth2Array
    authObj.Add "oauth2", oauth2Array
    
    ' Add "auth" object to main JSON
    '  json.item("auth") = authObj
    json.Add "auth", authObj
    
    ' Create "event" array and add to main JSON (empty array in this case)
    '  json.item("event") = JsonConverter.ParseJson("[]")
    json.Add "event", JsonConverter.ParseJson("[]")
    
    ' Output the JSON structure
    Dim jsonString As String
    jsonString = JsonConverter.ConvertToJson(json)
    
    ' Output the final JSON in Immediate Window (Ctrl+G to see it)
    SampleRESTPopup jsonString, "Collection Import"
    
End Sub

' Helper function to create oauth2 fields
Function Postman_CreateOAuthField(key As String, value As String) As Object
    Dim field As Object
    Set field = JsonConverter.ParseJson("{}")
    field.item("key") = key
    field.item("value") = value
    field.item("type") = "string"
    Set Postman_CreateOAuthField = field
End Function

Function Postman_ParseURL(fullUrl As String)
    
    ' Variables to hold different parts of the URL
    Dim protocol As String, host As String, path As String, queryString As String
    Dim hostParts() As String, pathParts() As String, queryParams() As String
    Dim queryKeys() As String, queryValues() As String
    Dim i As Long
    Dim remainingUrl, pathStart, pathAndQuery, queryStart As String
    
    ' Initialize JSON object
    Dim jsonObject As Object
    Set jsonObject = JsonConverter.ParseJson("{}")
    
    ' Create item structure in the JSON object
    Dim item As Object
    Set item = JsonConverter.ParseJson("{}")
    
    ' Split URL into protocol, host, path, and query parameters
    protocol = Left(fullUrl, InStr(fullUrl, "://") - 1)
    remainingUrl = Mid(fullUrl, InStr(fullUrl, "://") + 3)
    pathStart = InStr(remainingUrl, "/")
    host = Left(remainingUrl, pathStart - 1)
    hostParts = Split(host, ".")
    pathAndQuery = Mid(remainingUrl, pathStart + 1)
    queryStart = InStr(pathAndQuery, "?")
    If queryStart > 0 Then
        path = Left(pathAndQuery, queryStart - 1)
    Else
        path = pathAndQuery
    End If
    pathParts = Split(path, "/")
    If queryStart > 0 Then
        queryString = Mid(pathAndQuery, queryStart + 1)
        queryParams = Split(queryString, "&")
        ReDim queryKeys(UBound(queryParams))
        ReDim queryValues(UBound(queryParams))
        For i = LBound(queryParams) To UBound(queryParams)
            If InStr(queryParams(i), "=") > 0 Then
                queryKeys(i) = Split(queryParams(i), "=")(0)
                queryValues(i) = Split(queryParams(i), "=")(1)
            End If
        Next i
    End If
    
    ' Add URL components to the item structure
    With item
        ' Add raw
        .Add "raw", fullUrl
        
        ' Add protocol
        .Add "protocol", protocol
        
        ' Add host
        Dim hostJson As Object
        Set hostJson = JsonConverter.ParseJson("[]")
        For i = LBound(hostParts) To UBound(hostParts)
            hostJson.Add hostParts(i)
        Next i
        .Add "host", hostJson
        
        ' Add path
        Dim pathJson As Object
        Set pathJson = JsonConverter.ParseJson("[]")
        For i = LBound(pathParts) To UBound(pathParts)
            pathJson.Add pathParts(i)
        Next i
        .Add "path", pathJson
        
        ' Add query parameters
        Dim queryJson As Object
        Set queryJson = JsonConverter.ParseJson("[]")
        For i = LBound(queryKeys) To UBound(queryKeys)
            If Len(queryKeys(i)) > 0 Then
                Dim queryParam As Object
                Set queryParam = JsonConverter.ParseJson("{}")
                queryParam.Add "key", queryKeys(i)
                queryParam.Add "value", queryValues(i)
                queryJson.Add queryParam
            End If
        Next i
        .Add "query", queryJson
    End With
    
    ' Add the item structure to the main JSON object
    jsonObject.Add "url", item
    
    ' Print the resulting JSON
    Set Postman_ParseURL = jsonObject
End Function

Sub PrettyPrintData(inputText As String)
    ' Check if the input is JSON or XML
    If Left(inputText, 1) = "{" Then
        ' If it starts with "{", assume it's JSON and call the JSON pretty-print routine
        PrettyPrintJSON inputText
    ElseIf Left(inputText, 1) = "<" Then
        ' If it starts with "<", assume it's XML and call the XML pretty-print routine
        'PrettyPrintXML inputText
    Else
        MsgBox "Input is neither valid JSON nor XML."
    End If
End Sub

Sub PrettyPrintJSON(jsonText As String)
    Dim json As Object
    Dim prettyJSON As String
    
    On Error GoTo ErrorHandler                   ' Handle any JSON parsing errors
    
    ' Parse and pretty print JSON
    Set json = JsonConverter.ParseJson(jsonText)
    prettyJSON = JsonConverter.ConvertToJson(json, Whitespace:=2) ' Indentation of 2 spaces
    #If DEBUG_MODE Then
        Debug.Print prettyJSON                       ' Output to Immediate window
    #End If
    Exit Sub
    
ErrorHandler:
    MsgBox "Error parsing JSON: " & Err.description
    
End Sub

Sub PrettyPrintXML()
    Dim xmlDoc As Object
    Dim xmlText As String
    Dim prettyXML As String
    
    ' Your XML string
    xmlText = "<?xml version=""1.0"" encoding=""UTF-8""?>" & _
              "<SOAP-ENV:Envelope xmlns:SOAP-ENV=""http://schemas.xmlsoap.org/soap/envelope/"" " & _
              "xmlns:chg=""http://schemas.infor.com/ips/MWS980WS/RepairSelection"" " & _
              "xmlns:cred=""http://lawson.com/ws/credentials"">" & _
              "<SOAP-ENV:Header><cred:lws><cred:company>501</cred:company>" & _
              "<cred:division>USA</cred:division></cred:lws></SOAP-ENV:Header>" & _
              "<SOAP-ENV:Body><chg:RepairSelection><chg:MWS980>" & _
              "<chg:SelectionOrientation>0</chg:SelectionOrientation>" & _
              "<chg:AnalysisRound>WHS AL1</chg:AnalysisRound>" & _
              "</chg:MWS980></chg:RepairSelection></SOAP-ENV:Body>" & _
              "</SOAP-ENV:Envelope>"
    
    ' Load XML document
    Set xmlDoc = CreateObject("MSXML2.DOMDocument.6.0")
    xmlDoc.async = False
    xmlDoc.validateOnParse = False
    xmlDoc.preserveWhiteSpace = True
    xmlDoc.LoadXML xmlText
    
    ' Check if XML loaded successfully
    If xmlDoc.parseError.ErrorCode <> 0 Then
        MsgBox "Error in XML: " & xmlDoc.parseError.reason
        Exit Sub
    End If
    
    ' Format the XML by adding indentation and line breaks
    prettyXML = FormatXML(xmlDoc.DocumentElement, 0)
    
    ' Output to Immediate window
    #If DEBUG_MODE Then
        Debug.Print prettyXML
    #End If
End Sub

Sub ProcessParameters(result As Object, outputData As Variant, outputRow As Long)
    If result.exists("parameters") Then
        If Not result.item("parameters") Is Nothing Then
            Dim parameter As Variant
            Dim columnIndex As Integer
            columnIndex = 2
    
            For Each parameter In result.item("parameters").keys
                Dim paramValue As Variant
                paramValue = result.item("parameters").item(parameter)
                columnIndex = FindColumnIndex(parameter)
                If columnIndex > 0 Then outputData(outputRow, columnIndex) = paramValue
            Next parameter
        End If
    End If
End Sub



Private Function Process_BuildM3X(aInputFields, aInputValues)
    Dim url, inputField As String
    Dim i As Integer
    
    For i = 1 To 100
        If aInputValues(i) <> "" Then
            inputField = aInputFields(i)
            If inputField = "uuid" Then
                url = aInputValues(i) + "?"
            Else
                url = url + inputField & "=" & aInputValues(i)
            End If
        End If
    Next i
    Process_BuildM3X = url

End Function
Sub Clear_Token()
    m_s_AccessToken = ""
End Sub

Sub PromptSaveAsDialogMac()
    ' Display Save As dialog

End Sub

Function PromptUser(ByVal promptMessage As String) As Boolean
    Dim response As VbMsgBoxResult

    On Error Resume Next
    response = MsgBox(promptMessage, vbYesNo + vbQuestion, "Continue Processing")
    On Error GoTo 0

    If response = vbYes Then
        PromptUser = True
    End If
End Function

Function ReadFile(filePath As String) As String
    Dim fileNumber As Integer
    fileNumber = FreeFile

    Open filePath For Input As fileNumber
    ReadFile = Input$(LOF(fileNumber), fileNumber)
    Close fileNumber
End Function

Sub RenameSheet(proposedName As String)
    Dim position As Integer
    
    ' if name not passed then propose one
    If proposedName = "" Then
        proposedName = ActiveSheet.Range("API").value & " " & ActiveSheet.Range("Transaction").value
        If Len(proposedName) > 31 Then
            proposedName = ActiveSheet.Range("API").value
        End If
    End If
    proposedName = Replace(proposedName, "/", "")
    proposedName = Replace(proposedName, ":", "-")
    
    ' Find the position of the first "("
    position = InStr(1, proposedName, "(")
    If position > 0 Then
        proposedName = Trim(Left(proposedName, position - 1))
    End If
    
    If Not ActiveSheet.Name = proposedName Then
        If ActiveSheet.Range("API").value <> "" And ActiveSheet.Range("Transaction").value <> "" Then
            Dim newSheet As String
            newSheet = proposedName
            Dim sheetNumber As Integer
            sheetNumber = 0
            If Settings_SheetMatch(newSheet) Then
                sheetNumber = sheetNumber + 1
            End If
            Do While Settings_SheetMatch(newSheet & " (" & sheetNumber & ")")
                If ActiveSheet.Name <> newSheet & " (" & sheetNumber & ")" Then
                    sheetNumber = sheetNumber + 1
                Else
                    Exit Do
                End If
            Loop
            newSheet = newSheet & " (" & sheetNumber & ")"
            ActiveSheet.Name = Replace(newSheet, " (0)", "")
        End If
    End If

End Sub

Function ReplaceAlphaWithZero(inputString As String) As String
    Dim i As Integer
    Dim resultString As String
    resultString = inputString

    For i = 511 To Len(inputString)
        If i < 1263 Or i > 1563 Then
            If IsAlpha(Mid(inputString, i, 1)) Then
                Mid(resultString, i, 1) = "0"
            End If
        End If
    Next i

    ReplaceAlphaWithZero = resultString
End Function
Sub SampleREST()
    Dim strServiceName As String, strMethod As String
    Dim concatenatedResult As String
    Dim lastCol As Long, i As Long
    Dim fieldName As String
    Dim fieldValue As String
    
    Set m_obj_ws = ActiveSheet
    
    m_l_Row = 9
    m_s_Company = m_obj_ws.Range("Company").value
    m_s_Division = m_obj_ws.Range("Division").value
    strServiceName = m_obj_ws.Range("API").value
    strMethod = m_obj_ws.Range("Transaction").value
    
    ' Check if m_obj_ColumnNames needs refresh
    If m_obj_ColumnNames.count = -1 Or _
       InStr(m_s_LoadedUrl, strServiceName) = 0 Or _
       InStr(m_s_LoadedUrl, strMethod) = 0 Then
        Call GetLayout_Click(False)
    End If
    
    concatenatedResult = ""
    
    ' Find last column in row 8
    lastCol = m_obj_ws.Cells(8, m_obj_ws.columns.count).End(xlToLeft).column
    
    ' Loop through fields in row 8, starting from column 2
    For i = 2 To lastCol
        fieldName = Trim(m_obj_ws.Cells(8, i).value)
        fieldValue = Trim(m_obj_ws.Cells(9, i).value)
        
        If fieldName <> "" And fieldValue <> "" Then
            ' Optional: check that this field is an input field in m_obj_ColumnDirections
            'If m_obj_ColumnDirections.item(fieldName) = "I" Then
                concatenatedResult = concatenatedResult & fieldName & "=" & fieldValue & "&"
            'End If
        End If
    Next i
    
    ' Remove trailing "&"
    If Right(concatenatedResult, 1) = "&" Then
        concatenatedResult = Left(concatenatedResult, Len(concatenatedResult) - 1)
    End If
    
    ' Build final URL
    concatenatedResult = strServiceName & "/" & strMethod & "?" & concatenatedResult
    
    SampleRESTPopup concatenatedResult, "REST Sample"
End Sub

' DoppioGroup Module Code
Sub SampleRESTPopup(passedString As String, title As String)

    Dim myForm As SampleREST
    Set myForm = New SampleREST
    myForm.FormTitle = title
    myForm.TextBox1.text = passedString
    myForm.StartUpPosition = 0
    myForm.Left = (Application.Width - myForm.Width) / 2
    myForm.Top = (Application.Height - myForm.Height) / 2
    myForm.Show

End Sub

Sub Schema_ExtractTransactions()
    Dim xmlDoc As Object
    Dim transactions As New ArrayList
    Dim rowNum As Integer
    Dim ws As Worksheet
    
    transactions.Initialize
    rowNum = 3
    Set ws = ThisWorkbook.Sheets("Transactions")

    ' Check If running on Windows Or Mac
    #If Mac Then
        ' For Mac, use MSXML2.DOMDocument
        Set xmlDoc = CreateObject("MSXML2.DOMDocument")
    #Else
        ' For Windows, use DOMDocument
        Set xmlDoc = CreateObject("MSXML2.DOMDocument.6.0")
    #End If

    apicall_Bridge m_s_stUrl, "/mws/services", ActiveSheet.Range("API").value & ".meta", "", ""

    ' Load the XML string
    xmlDoc.LoadXML m_s_Meta
    ' Check For errors in the XML
    If xmlDoc.parseError.ErrorCode <> 0 Then
        MsgBox "Error in XML: " & xmlDoc.parseError.reason
        Exit Sub
    End If

    ' Register the namespace
    xmlDoc.SetProperty "SelectionNamespaces", "xmlns:v2='http://schemas.intentia.net/mws/meta/webservice/v2.2/' xmlns:v21='http://schemas.intentia.net/mws/meta/mpd/v21/' xmlns:v211='http://schemas.intentia.net/mws/meta/mpdi/v21/'"

    ' Navigate To the operations node
    Dim operationsNode As Object
    Set operationsNode = xmlDoc.SelectSingleNode("/v2:webservice/v2:operations")
    If Not operationsNode Is Nothing Then
        Dim operationNode As Object
        For Each operationNode In operationsNode.ChildNodes
            Dim operationName As String
            operationName = operationNode.Attributes.getNamedItem("name").text
            transactions.Add (operationName)
        Next operationNode
    End If

    transactions.Sort
    Dim i As Long
    For i = 1 To transactions.count
        'Debug.Print transactions.Item(i)
        ws.Cells(rowNum, 1).value = ActiveSheet.Range("API").value
        ws.Cells(rowNum, 2).value = transactions.item(i)
        '        Schema_ExtractInputandOuput xmlDoc, transactions.Item(i)
        rowNum = rowNum + 1
    Next i

End Sub

Sub Schema_ExtractTypeAndLength(xmlDoc As Object, operation As String, panelName As String, fieldName As String, Alias As String, Required As String)

    ' Specify the namespace prefixes
    xmlDoc.SetProperty "SelectionNamespaces", "xmlns:v2='http://schemas.intentia.net/mws/meta/webservice/v2.2/' xmlns:v21='http://schemas.intentia.net/mws/meta/mpd/v21/' xmlns:v211='http://schemas.intentia.net/mws/meta/mpdm/v21/'"

    ' Construct the XPath expression based on the provided panelName And fieldName
    Dim xpath As String
    xpath = "/v2:webservice/v2:operations/v2:operation[@name='" & operation & "']/v21:mpd/v21:mpdmetadata/v211:panels/v211:panel[@name='" & panelName & "']/v211:fields/v211:field[@name='" & fieldName & "']"

    ' Select the specified field node
    Dim fieldNode As Object
    Set fieldNode = xmlDoc.SelectSingleNode(xpath)

    ' Check If the field node is found
    If Not fieldNode Is Nothing Then
        ' Extract type And fieldLength values
        Dim fieldType As String
        Dim fieldLength As String
        Dim langTxt As String

        fieldType = fieldNode.SelectSingleNode("v211:type").text
        fieldLength = fieldNode.SelectSingleNode("v211:fieldLength").text
        langTxt = fieldNode.SelectSingleNode("v211:langTxt").text

        m_obj_Layout.Add Alias & vbTab & langTxt & vbTab & fieldType & vbTab & fieldLength & vbTab & Required
    End If

End Sub

Sub SetFormulasAndFormatting()
    Dim ws As Worksheet
    Dim targetCell As Range
    
    ' Set the worksheet variable to your target sheet
    Set ws = ActiveSheet
    
    GetTransactions_Click
    
    ' Set formula and formatting for cell I6
    Set targetCell = ws.Range("I6")
    targetCell.NumberFormat = "General"
    targetCell.Formula = "=MAX(COUNTA(B9:B1048576), COUNTA(C9:C1048576), COUNTA(D9:D1048576), COUNTA(E9:E1048576), COUNTA(F9:F1048576), COUNTA(G9:G1048576), COUNTA(H9:H1048576), COUNTA(I9:I1048576), COUNTA(J9:J1048576))"
    
    If ws.Range("A3").value = "" Then
        ' Set formula and formatting for cell B4
        Set targetCell = ws.Range("B4")
        targetCell.NumberFormat = "General"
        targetCell.Formula = "=COUNTIF(A:A, ""NOK *"")"
        
        ' Set formula and formatting for cell B5
        Set targetCell = ws.Range("B5")
        targetCell.NumberFormat = "General"
        targetCell.Formula = "=COUNTIF(A:A, ""OK"")"
        
        ' Set formula and formatting for cell B6
        Set targetCell = ws.Range("B6")
        targetCell.NumberFormat = "General"
        targetCell.Formula = "=SUM(I6-(B4+B5))"
    End If
    
End Sub

' =============================================================================
' MODERNIZED Settings
' =============================================================================
' Handles various settings commands
' Uses Doppio.* for global variables
' =============================================================================

Public Sub Settings_New(value As String)
    Dim startIndex As Integer
    Dim endIndex As Integer
    Dim numericValue As Long
    Dim ws As Worksheet
    Dim cacheSheet As Worksheet
    Dim tempFilePath As String
    Dim activeTransaction As String
    
    Set ws = ThisWorkbook.ActiveSheet
    
    ' ==========================================
    ' Sheet Creation Commands
    ' ==========================================
    If value = "New sheet" Or value = "ns" Then
        Doppio.Settings_NewSheet
        Exit Sub
    End If
    
    ' ==========================================
    ' Export/Report Commands
    ' ==========================================
    If value = "xlsx" Or value = "xls" Or value = "report" Then
        Doppio.Create_xlsx
        Exit Sub
    End If
    
    ' ==========================================
    ' Code Generation Commands
    ' ==========================================
    If value = "samplerest" Or value = "rest" Then
        Doppio.SampleREST
        Exit Sub
    End If
    
    If value = "curl" Then
        Doppio.SampleRESTPopup Doppio.curlCommand, "CURL Sample"
        Exit Sub
    End If
    
    If value = "postman" Then
        Doppio.Postman_Build
        Exit Sub
    End If
    
    If value = "python" Then
        Doppio.Python_GenerateFunction
        Exit Sub
    End If
    
    ' ==========================================
    ' Data Manipulation Commands
    ' ==========================================
    If value = "pivot" Then
        Doppio.RunBuildMatrix
        Exit Sub
    End If
    
    If value = "unpivot" Then
        Doppio.RunUnpivotMatrix
        Exit Sub
    End If
    
    If value = "decode" Then
        Doppio.ToggleBase64Button "Decode"
        Exit Sub
    End If
    
    If value = "encode" Then
        Doppio.ToggleBase64Button "Encode"
        Exit Sub
    End If
    
    If value = "evs100" Then
        Doppio.ExportToEVS100
        Exit Sub
    End If
    
    ' ==========================================
    ' Prep Command - Full Reset
    ' ==========================================
    If value = "prep" Then
        Doppio.Settings_DeleteSheets
        
        With ActiveSheet
            .Range("API").value = "CRS111MI"
            .Range("Type").value = "API"
            .Range("Transaction").value = ""
            .Range("Environment").value = ""
            .Range("User").value = ""
            .Range("Company").value = ""
            .Range("Division").value = ""
            .Range("G6").value = ""
        End With
        
        ' Hide system sheets
        On Error Resume Next
        Sheets("Master").Visible = False
        Sheets("Log").Visible = False
        Sheets("Cache").Visible = False
        Sheets("Environments").Visible = False
        Sheets("AvailableMIs").Visible = False
        Sheets("Transactions").Visible = False
        Sheets("Versions").Visible = False
        Sheets("Help").Visible = True
        On Error GoTo 0
        
        ' Clear Environments sheet
        On Error Resume Next
        Set ws = Worksheets("Environments")
        On Error GoTo 0
        If Not ws Is Nothing Then
            ws.Cells.ClearContents
            ws.columns("A").ClearContents
        End If
        
        Set ws = ActiveSheet
        FilterRow8BasedOnPopulatedColumns_New ws
        ws.Rows("7:8").ClearContents
        
        Application.Calculation = xlCalculationAutomatic
        DoEvents
        DoppioUI.UI_KillPleaseWait
        
        ActiveSheet.Name = "Sheet1"
        SetFormulasAndFormatting_New ws
        Doppio.Settings_CopyDefaults
        
        ' Clean up temp files
        On Error Resume Next
        tempFilePath = Environ("HOME") & "/curl_output.txt"
        Kill tempFilePath
        tempFilePath = Environ("HOME") & "/curl_input.sh"
        Kill tempFilePath
        On Error GoTo 0
        
        AutoFit_Click_New
        Exit Sub
    End If
    
    ' ==========================================
    ' Reset Command
    ' ==========================================
    If value = "reset" Then
        Doppio.CleanSheet_Click
        
        activeTransaction = ActiveSheet.Range("Transaction").value
        ActiveSheet.Range("Transaction").value = ""
        FilterRow8BasedOnPopulatedColumns_New ws
        ActiveSheet.Range("Transaction").value = activeTransaction
        
        Doppio.Settings_DeleteSheets
        ActiveSheet.Range("G6").value = ""
        
        ' Hide system sheets
        On Error Resume Next
        Sheets("Master").Visible = False
        Sheets("Log").Visible = False
        Sheets("Cache").Visible = False
        Sheets("Settings").Visible = False
        Sheets("Environments").Visible = False
        Sheets("AvailableMIs").Visible = False
        Sheets("Transactions").Visible = False
        Sheets("Versions").Visible = False
        On Error GoTo 0
        
        Doppio.Settings_CopyDefaults
        
        ' Clean up temp files
        On Error Resume Next
        tempFilePath = Environ("HOME") & "/curl_output.txt"
        Kill tempFilePath
        tempFilePath = Environ("HOME") & "/curl_input.sh"
        Kill tempFilePath
        On Error GoTo 0
        
        ActiveSheet.Name = "Sheet1"
        SetFormulasAndFormatting_New ws
        Exit Sub
    End If
    
    ' ==========================================
    ' Sheet Visibility Toggles
    ' ==========================================
    If value = "environments" Then
        Sheets("Environments").Visible = Not Sheets("Environments").Visible
        Exit Sub
    End If
    
    If value = "master" Then
        Sheets("Master").Visible = Not Sheets("Master").Visible
        Exit Sub
    End If
    
    If value = "log" Then
        Sheets("Log").Visible = Not Sheets("Log").Visible
        Exit Sub
    End If
    
    If value = "apis" Then
        Sheets("AvailableMIs").Visible = Not Sheets("AvailableMIs").Visible
        Exit Sub
    End If
    
    If value = "transactions" Then
        Sheets("Transactions").Visible = Not Sheets("Transactions").Visible
        Exit Sub
    End If
    
    If value = "unhide" Then
        Sheets("Environments").Visible = True
        Sheets("AvailableMIs").Visible = True
        Exit Sub
    End If
    
    If value = "hide" Then
        On Error Resume Next
        Sheets("Master").Visible = False
        Sheets("Log").Visible = False
        Sheets("Cache").Visible = False
        Sheets("Settings").Visible = False
        Sheets("Environments").Visible = False
        Sheets("AvailableMIs").Visible = False
        Sheets("Transactions").Visible = False
        Sheets("Help").Visible = False
        Sheets("Versions").Visible = False
        On Error GoTo 0
        Exit Sub
    End If
    
    ' ==========================================
    ' Cache Commands
    ' ==========================================
    If value = "cache" Then
        Doppio.RecordCache_Display
        Exit Sub
    End If
    
    If value = "clear cache" Or value = "cc" Then
        DoppioCache.RecordCache_Reset
        On Error Resume Next
        Set cacheSheet = ThisWorkbook.Sheets("Cache")
        On Error GoTo 0
        If Not cacheSheet Is Nothing Then
            cacheSheet.Rows("2:" & cacheSheet.Rows.count).ClearContents
        End If
        Exit Sub
    End If
    
    If value = "load cache" Then
        RecordCache_Load
        Exit Sub
    End If
    
    ' ==========================================
    ' Numeric Settings (maxrecs, maxbulk, refresh)
    ' ==========================================
    If Left(value, Len("maxrecs")) = "maxrecs" Then
        startIndex = InStr(value, "=")
        If startIndex > 0 Then
            endIndex = Len(value)
            numericValue = CLng(Mid(value, startIndex + 1, endIndex - startIndex))
            Doppio.maxrecs = numericValue
            DoppioConfig.Config_MaxRecords = numericValue
            #If DEBUG_MODE Then
                Debug.Print "Settings: maxrecs = " & numericValue
            #End If
        End If
        Exit Sub
    End If
    
    If Left(value, Len("maxbulk")) = "maxbulk" Then
        startIndex = InStr(value, "=")
        If startIndex > 0 Then
            endIndex = Len(value)
            numericValue = CLng(Mid(value, startIndex + 1, endIndex - startIndex))
            Doppio.maxbulk = CInt(numericValue)
            #If DEBUG_MODE Then
                Debug.Print "Settings: maxbulk = " & Doppio.maxbulk
            #End If
        End If
        Exit Sub
    End If
    
    If Left(value, Len("refresh")) = "refresh" Then
        startIndex = InStr(value, "=")
        If startIndex > 0 Then
            endIndex = Len(value)
            numericValue = CLng(Mid(value, startIndex + 1, endIndex - startIndex))
            Doppio.refreshSeconds = CInt(numericValue)
            #If DEBUG_MODE Then
                Debug.Print "Settings: refreshSeconds = " & Doppio.refreshSeconds
            #End If
        End If
        Exit Sub
    End If
    
    ' ==========================================
    ' Other Commands
    ' ==========================================
    If value = "defaults" Then
        Doppio.Settings_CopyDefaults
        Doppio.SettingsSheet
        Exit Sub
    End If
    
    If value = "help" Then
        Doppio.HelpSheet
        Exit Sub
    End If
    
    If value = "settings" Then
        Doppio.SettingsSheet
        Exit Sub
    End If
    
    If value = "clear" Or value = "clr" Then
        ActiveSheet.Rows("9:" & ActiveSheet.Rows.count).ClearContents
        ActiveSheet.columns("A:" & ActiveSheet.columns.count).ClearContents
        Exit Sub
    End If
    
    If value = "clearstatus" Or value = "clrsts" Then
        ActiveSheet.Range("A9:A" & ActiveSheet.Rows.count).ClearContents
        Exit Sub
    End If
    
    If Left(value, Len("load s")) = "load s" Then
        LoadSourceAPIFormats
        Exit Sub
    End If
End Sub

' =============================================================================
' LEGACY WRAPPER
' =============================================================================

Public Sub settings(value As String)
    Settings_New value
End Sub


Sub SettingsSheet()
    Dim ws As Worksheet
    
    On Error GoTo ErrorHandler
    
    Set ws = ThisWorkbook.Sheets("Settings")
    
    ' Update settings values from DoppioConfig
    Dim settings As ApiSettings
    settings = DoppioConfig.Config_ApiSettings
    
    Application.ScreenUpdating = False
    
    With ws
        .Visible = True
        .Range("maxrecs").value = settings.MaxRecords
        .Range("maxbulk").value = settings.maxbulk
        .Range("refreshSeconds").value = settings.refreshSeconds
        .Range("formatting").value = settings.formatting
        .Range("righttrim").value = settings.righttrim
        .Range("splitChar").value = settings.splitChar
        .Range("maxtime").value = settings.MaxTimeout
        .Range("conoDivi").value = settings.conoDivi
        .Activate
    End With
    
    ' Freeze the top 4 rows
    With ActiveWindow
        If .FreezePanes Then .FreezePanes = False
        .SplitColumn = 0
        .SplitRow = 4
        .FreezePanes = True
    End With
    
    ' Position cursor
    ws.Range("D7").Select
    
    Application.ScreenUpdating = True
    Exit Sub
    
ErrorHandler:
    Application.ScreenUpdating = True
    MsgBox "Error showing settings: " & Err.description, vbExclamation
End Sub

Sub Settings_CopyDefaults()
    Dim ws As Worksheet, settings As Worksheet
    Dim wasHidden As Boolean
    Dim screenUpdatingState As Boolean
    Dim weChangedScreenUpdating As Boolean

    ' Store the current ScreenUpdating state
'    screenUpdatingState = Application.ScreenUpdating
'    weChangedScreenUpdating = False

    ' Only turn off screen updating if it's currently on
'    If screenUpdatingState Then
'        Application.ScreenUpdating = False
'        weChangedScreenUpdating = True
'    End If

    Set ws = ActiveSheet
    Set settings = ThisWorkbook.Sheets("Settings")

    ' Check if the sheet is hidden
    wasHidden = (settings.Visible <> xlSheetVisible)

    ' If it was hidden, make it visible temporarily
    If wasHidden Then settings.Visible = xlSheetVisible

    settings.Range("E7:E14").Copy
    settings.Range("D7:D14").PasteSpecial Paste:=xlPasteValues

    Application.CutCopyMode = False

    ' Set the fields with values from the Settings sheet
    maxrecs = settings.Range("maxrecs").value
    maxbulk = settings.Range("maxbulk").value
    refreshSeconds = settings.Range("refreshSeconds").value
    righttrim = settings.Range("righttrim").value
    formatting = settings.Range("formatting").value
    splitChar = settings.Range("splitChar").value
    maxtime = settings.Range("maxtime").value
    conoDivi = settings.Range("conoDivi").value
    
    If wasHidden Then settings.Visible = xlSheetVeryHidden
    ws.Activate

    ' Restore screen updating only if we changed it
    'If weChangedScreenUpdating Then Application.ScreenUpdating = True
End Sub

Sub Settings_DeleteSheets()
    Dim ws As Worksheet
    Dim protectedSheets As Variant
    Dim sheetName As String

    ' List of sheets that should Not be deleted
    protectedSheets = Array("Environments", "Master", "AvailableMIs", "Transactions", "Versions", "Help", "Settings", "Log", "Cache")

    ' Add the active sheet To the protected sheets list
    If Not ActiveSheet Is Nothing Then
        ReDim Preserve protectedSheets(UBound(protectedSheets) + 1)
        protectedSheets(UBound(protectedSheets)) = ActiveSheet.Name
    End If

    Application.DisplayAlerts = False
    For Each ws In ThisWorkbook.Sheets
        sheetName = ws.Name
        If isError(Application.match(sheetName, protectedSheets, 0)) Then
            Worksheets(sheetName).Delete
        End If
    Next ws
    Application.DisplayAlerts = True
End Sub

Sub Settings_NewSheet()
    Dim masterSheetName, activeEnv, activeUser, activeCompany, activeDivision, activeAPI, activeType, activeTransaction As String
    Dim ws As Worksheet
    
    Set ws = ThisWorkbook.ActiveSheet
    
    activeEnv = ActiveSheet.Range("Environment").value
    activeUser = ActiveSheet.Range("User").value
    activeCompany = ActiveSheet.Range("Company").value
    activeDivision = ActiveSheet.Range("Division").value
    activeAPI = ActiveSheet.Range("API").value
    activeType = ActiveSheet.Range("Type").value
    activeTransaction = ActiveSheet.Range("Transaction").value

    masterSheetName = "Master"
    Dim newSheetBaseName As String
    newSheetBaseName = "Sheet"
    Dim sheetNumber As Integer
    sheetNumber = 1

    Do While Settings_SheetExists(newSheetBaseName & sheetNumber)
        sheetNumber = sheetNumber + 1
    Loop

    Dim newSheetName As String
    newSheetName = newSheetBaseName & sheetNumber

    Sheets(masterSheetName).Copy After:=Sheets(Sheets.count)
    Sheets(Sheets.count).Name = newSheetName
    Sheets(newSheetName).Visible = xlSheetVisible
    Sheets(newSheetName).Activate
    Default_Buttons
    
    ActiveSheet.Range("Environment").value = activeEnv
    ActiveSheet.Range("User").value = activeUser
    ActiveSheet.Range("Company").value = activeCompany
    ActiveSheet.Range("Division").value = activeDivision
    ActiveSheet.Range("API").value = activeAPI
    ActiveSheet.Range("Type").value = activeType
    ActiveSheet.Range("Transaction").value = ""
    FilterRow8BasedOnPopulatedColumns_New ws
    ActiveSheet.Range("Transaction").value = activeTransaction
    ChangeCellColorBasedOnEnvironment
    ActiveSheet.Range("A2").Select
End Sub

Function Settings_SheetExists(partialSheetName As String) As Boolean
    Dim ws As Worksheet
    For Each ws In Worksheets
        If InStr(1, ws.Name, partialSheetName) > 0 Then
            Settings_SheetExists = True
            Exit Function
        End If
    Next ws
    Settings_SheetExists = False
End Function

Function Settings_SheetMatch(exactSheetName As String) As Boolean
    Dim ws As Worksheet
    For Each ws In Worksheets
        If UCase(ws.Name) = UCase(exactSheetName) Then
            Settings_SheetMatch = True
            Exit Function
        End If
    Next ws
    Settings_SheetMatch = False
End Function

Public Sub ShowPleaseWait(message As String)
    On Error GoTo ErrorHandler

    Dim boxsize As Integer
    boxsize = 50

    ' Calculate the required box size
    Dim i As Integer
    For i = 1 To Len(message)
        If Mid(message, i, 1) = vbCr Then
            boxsize = boxsize + 22
        End If
    Next i

    ' Ensure ws is Set To the active sheet
    Set m_obj_ws = ActiveSheet

    ' Define positions For the shape And button
    Dim shp As Shape
    Dim btnLeftPosition As Double
    Dim btnTopPosition As Double
    Dim leftPosition As Double
    Dim topPosition As Double

    With m_obj_ws
        ' Delete existing shape If it exists
        On Error Resume Next
        .Shapes("Wait").Delete
        On Error GoTo 0

        ' Add a New shape And position it around row 7
        leftPosition = .Cells(7, 1).Left
        topPosition = .Cells(7, 1).Top

        Set shp = .Shapes.AddShape(msoShapeRectangle, leftPosition + 350, topPosition + 100, 300, boxsize)
        shp.Name = "Wait"

        ' Customize the shape
        With shp
            .TextFrame.Characters.text = vbCr & vbTab & message
            .Fill.ForeColor.RGB = RGB(0, 0, 0)
            .TextFrame.Characters.Font.Color = RGB(255, 255, 255)
            .TextFrame.Characters.Font.Name = "Avenir"
            .TextFrame.Characters.Font.FontStyle = "Bold"
        End With

        btnLeftPosition = leftPosition + 575
        btnTopPosition = topPosition - 60 + boxsize + 10
    End With

    Application.ScreenUpdating = True
    DoEvents
    Application.ScreenUpdating = False

    Exit Sub

ErrorHandler:
    ' Handle any errors that occurred
    MsgBox "An error occurred: " & Err.description
    On Error GoTo 0
End Sub

Private Function soapXMLBody_create(method, program, aInputColumns, aInputFields, aInputValues)
    Dim xmlString As String
    Dim i As Integer
    Dim level1, level2, level3, level4 As String
    Dim currentLevel As String
    Dim inputField As String

    level1 = program
    xmlString = ""
    xmlString = xmlString & "<SOAP-ENV:Body>"
    xmlString = xmlString & "<chg:" & method & ">" 'namespace + webservice + method

    For i = 1 To 100
        If aInputValues(i) <> "" Then
            inputField = aInputFields(i)
            'Debug.Print currentLevel & vbTab & aInputValues(i)
            If currentLevel = "" Then
                currentLevel = level1
                xmlString = xmlString & "<chg:" & level1 & ">"
            End If
            If InStr(aInputFields(i), ":") Then
                'related program
                If Left(aInputFields(i), InStr(aInputFields(i), ":") - 1) <> currentLevel Then
                    'New related program
                    currentLevel = Left(aInputFields(i), InStr(aInputFields(i), ":") - 1)
                    If level2 = "" Then
                        level2 = currentLevel
                    ElseIf level3 = "" Then
                        level3 = currentLevel
                    ElseIf level4 = "" Then
                        level4 = currentLevel
                    End If
                    'add New level
                    xmlString = xmlString & "<chg:" & currentLevel & ">"
                End If
                inputField = Mid(aInputFields(i), InStr(aInputFields(i), ":") + 1, 15)
            End If
            If aInputColumns(i) <> 0 Then
                xmlString = xmlString & "<chg:" & inputField & ">" & aInputValues(i) & "</chg:" & inputField & ">"
            End If
        End If
    Next i

    ' end levels
    If level4 <> "" Then xmlString = xmlString & "</chg:" & level4 & ">"
    If level3 <> "" Then xmlString = xmlString & "</chg:" & level3 & ">"
    If level2 <> "" Then xmlString = xmlString & "</chg:" & level2 & ">"
    xmlString = xmlString & "</chg:" & level1 & ">"
    xmlString = xmlString & "</chg:" & method & ">"
    xmlString = xmlString & "</SOAP-ENV:Body>"
    xmlString = xmlString & "</SOAP-ENV:Envelope>"

    'Debug.Print xmlString
    soapXMLBody_create = xmlString
End Function

Private Function soapXMLHeader_create(webservice, method, cono, divi)
    Dim xmlString As String

    xmlString = ""
    xmlString = xmlString & "<?xml version=""1.0"" encoding=""UTF-8""?>"
    If m_b_Multitenant Then
        xmlString = xmlString & "<SOAP-ENV:Envelope xmlns:SOAP-ENV=""http://schemas.xmlsoap.org/soap/envelope/"" xmlns:chg=""http://schemas.infor.com/ips/" & webservice & "/" & method & """ xmlns:cred=""http://lawson.com/ws/credentials"">"
    Else
        xmlString = xmlString & "<SOAP-ENV:Envelope xmlns:SOAP-ENV=""http://schemas.xmlsoap.org/soap/envelope/"" xmlns:chg=""http://your.company.net/" & webservice & "/" & method & """ xmlns:cred=""http://lawson.com/ws/credentials"">"
    End If
    xmlString = xmlString & "<SOAP-ENV:Header>"
    xmlString = xmlString & "<cred:lws>"
    xmlString = xmlString & "<cred:company>" + cono + "</cred:company>"
    xmlString = xmlString & "<cred:division>" + divi + "</cred:division>"
    xmlString = xmlString & "</cred:lws>"
    xmlString = xmlString & "</SOAP-ENV:Header>"

    'Debug.Print xmlString
    soapXMLHeader_create = xmlString
End Function

Sub SortTransactions()
    Dim ws As Worksheet
    Dim sortRange As Range

    ' Specify the worksheet
    Set ws = ThisWorkbook.Sheets("Transactions")

    ' Define the range To be sorted
    With ws
        Set sortRange = .Range(.Cells(3, "A"), .Cells(.Cells(.Rows.count, "A").End(xlUp).row, "F"))
    End With

    ' Sort the range by the values in column B
    With sortRange
        .Sort Key1:=.columns(2), Order1:=xlAscending, Header:=xlNo
    End With
End Sub

Function SplitString(ByVal inputString As String, ByVal delimiter As String, ByVal index As Integer) As String
    Dim arr() As String
    arr = Split(inputString, delimiter)
    If index <= UBound(arr) + 1 Then
        SplitString = arr(index - 1)
    Else
        SplitString = ""
    End If
End Function

Sub ST_ExtractWebServices()
    Dim lines() As String
    Dim refLines() As String
    Dim i As Integer, j As Integer
    Dim ws As Worksheet
    Dim rowNum As Integer
    Dim serviceName As String
    
    apicall_Bridge m_s_stUrl, "/mws-ws", "", "", ""
    lines = Split(m_s_Meta, vbLf)
    ReDim refLines(0 To UBound(lines))
    For i = LBound(lines) To UBound(lines)
        If Left(lines(i), 7) = "<a href" Then
            refLines(j) = lines(i)
            j = j + 1
        End If
    Next i
    ReDim Preserve refLines(0 To j - 1)

    Set ws = ThisWorkbook.Sheets("AvailableMIs")
    rowNum = 3
    For i = LBound(refLines) To UBound(refLines)
        If InStr(refLines(i), "/services/") > 0 Then
            serviceName = ExtractServiceName(refLines(i))
            ws.Cells(rowNum, 2).value = serviceName
            rowNum = rowNum + 1
        End If
    Next i
End Sub





Sub Tenant_Token()
    Dim ws As Worksheet
    Dim env As Environment
    Dim tokenUrl As String
    Dim body As String
    Dim config As httpConfig
    Dim httpResponse As httpResponse
    Dim json As Object
    Dim apiResponse As apiResponse
    
    ' New variables for sheet handling
    Dim wsEnv As Worksheet
    Dim rngFound As Range
    Dim bIsCached As Boolean ' Flag to track if we loaded from sheet
    
    On Error GoTo ErrorHandler
    
    Set ws = ActiveSheet
    ws.Range("J3").value = 0
    m_s_stToken = ""
    bIsCached = False
    
    ' Update selected environment if changed
    If m_s_SelectedEnvironment <> ws.Range("I2").value Then
        m_s_SelectedEnvironment = ws.Range("I2").value
    End If
    
    ' Load tenant configuration
    Doppio.Tenant_Information
    
    ' Safety Check: Ensure Manager is Loaded
    If manager Is Nothing Then
        MsgBox "The Environment Manager is not loaded. Please close and re-open the workbook to initialize.", vbCritical
        ws.Range("J3").value = 0
        Exit Sub
    End If
    
    Set env = manager.GetEnvironment(m_s_SelectedEnvironment)
    
    If env Is Nothing Then
        MsgBox "Environment '" & m_s_SelectedEnvironment & "' not found in manager.", vbCritical
        ws.Range("J3").value = 0
        Exit Sub
    End If
    
    ' ---------------------------------------------------------
    ' STEP 1: Check Environments Sheet for existing cached session
    ' ---------------------------------------------------------
    
    ' Reset globals
    m_s_AccessToken = ""
'    m_s_M3user = ""
    m_s_Company = ""
    m_s_Division = ""
    
    On Error Resume Next
    Set wsEnv = ThisWorkbook.Sheets("Environments")
    On Error GoTo ErrorHandler
    
    If Not wsEnv Is Nothing Then
        Set rngFound = wsEnv.columns("A").Find(What:=m_s_SelectedEnvironment, _
                                               LookIn:=xlValues, _
                                               LookAt:=xlWhole)
        
        If Not rngFound Is Nothing Then
            ' Retrieve values
            m_s_AccessToken = Trim(wsEnv.Cells(rngFound.row, "E").value)
'            m_s_M3user = Trim(wsEnv.Cells(rngFound.row, "D").value)
            m_s_Company = Trim(wsEnv.Cells(rngFound.row, "G").value)
            m_s_Division = Trim(wsEnv.Cells(rngFound.row, "H").value)
            
            ' INTEGRITY CHECK:
            ' If ANY field is blank, cache is invalid.
            If m_s_AccessToken = "" Or _
               m_s_M3user = "" Or _
               m_s_Company = "" Or _
               m_s_Division = "" Then
               
                #If DEBUG_MODE Then
                    Debug.Print "Tenant_Token: Cache incomplete. Forcing fresh login."
                #End If
                m_s_AccessToken = ""
                bIsCached = False
            Else
                #If DEBUG_MODE Then
                    Debug.Print "Tenant_Token: Cache Hit! Skipping validation."
                #End If
                ws.Range("User").value = m_s_M3user
                ws.Range("Company").value = m_s_Company
                ws.Range("Division").value = m_s_Division
                bIsCached = True
            End If
        End If
    End If

    ' ---------------------------------------------------------
    ' STEP 2: Logic Split - Cached vs Fresh
    ' ---------------------------------------------------------
    
    ' Set common URLs (Needed for both paths)
    m_s_MainUrl = iu & "/" & ti
    m_s_MiPath = "/M3/m3api-rest/v2/execute"
    m_s_WsPath = "/M3/ips/service"
    m_s_MiUrl = "MRS001MI/GetUserInfo/?"
    m_b_Webservice = False
    
    ' PATH A: We have a valid cache
    If bIsCached Then
        ' We assume the token is valid. If it's expired, the next API call
        ' in the main process will fail (401) and should trigger a retry anyway.
        m_s_TokenType = "Bearer"
        ws.Range("J3").value = 2 ' Set status to Connected
        
    ' PATH B: No cache, perform full login and validation
    Else
        ' 1. Get Token
        tokenUrl = pu & ot
        body = "client_id=" & Core_UrlEncode(ci) & _
               "&client_secret=" & Core_UrlEncode(cs) & _
               "&grant_type=password" & _
               "&username=" & Core_UrlEncode(saak) & _
               "&password=" & Core_UrlEncode(sask)
        
        #If DEBUG_MODE Then
            Debug.Print "Tenant_Token: Fetching new token..."
        #End If
        
        config.url = tokenUrl
        config.method = HttpMethod_POST
        config.contentType = "application/x-www-form-urlencoded"
        config.AcceptType = "application/json"
        config.authHeader = ""
        config.body = body
        config.TimeoutSeconds = 30
        
        httpResponse = DoppioHttp.ExecuteRequest(config)
        
        If Not httpResponse.success Then
            MsgBox "Unable to get token: " & httpResponse.errorMessage, vbCritical
            ws.Range("J3").value = 0
            Exit Sub
        End If
        
        Set json = JsonConverter.ParseJson(httpResponse.body)
        
        ' Error checking
        On Error Resume Next
        Dim errorMsg As String
        errorMsg = ""
        If Not json Is Nothing Then errorMsg = json.item("error")
        On Error GoTo ErrorHandler
        
        If errorMsg <> "" Then
            MsgBox "Unable to connect: " & errorMsg, vbCritical
            ws.Range("J3").value = 0
            Exit Sub
        End If
        
        If Not json Is Nothing Then
            m_s_AccessToken = json.item("access_token")
            m_s_TokenType = json.item("token_type")
            'If json.exists("refresh_token") Then m_s_RefreshToken = json("refresh_token")
        End If
        
        If m_s_TokenType = "" Then m_s_TokenType = "Bearer"
        
        ' 2. Validate & Get User Info (Since we didn't have it in cache)
        #If DEBUG_MODE Then
            Debug.Print "Tenant_Token: Validating new session..."
        #End If
        ws.Range("J3").value = 1
        
        apiResponse = ExecuteNewApiCall("MRS001MI", "GetUserInfo", "")
        
        If apiResponse.success And Not apiResponse.records Is Nothing Then
            Set m_obj_Records = apiResponse.records
            Output_GetUserInfo ' Sets m_s_M3user, Company, Division globals
            DoEvents
            
            ' 3. Save to Cache (Using the helper sub)
            UpdateEnvironmentCache m_s_SelectedEnvironment, _
                                   m_s_AccessToken, _
                                   m_s_M3user, _
                                   m_s_Company, _
                                   m_s_Division
                                   
            ws.Range("J3").value = 2
        Else
            ' Login failed
            If Not rngFound Is Nothing Then
                 wsEnv.Range("E" & rngFound.row & ":H" & rngFound.row).ClearContents
            End If
            MsgBox "Connection failed during validation.", vbExclamation
            ws.Range("J3").value = 0
            Exit Sub
        End If
        
    End If
    
    ' ---------------------------------------------------------
    ' Final Cleanup (Runs for both paths)
    ' ---------------------------------------------------------
    
    ' Update status check
    If m_s_AccessToken = "" Or ws.Cells(4, 9).value = "" Then
        ws.Range("J3").value = 0
    Else
        ws.Range("J3").value = 2
        activeEnvironment = ws.Range("I2").value
    End If
    
    ' Save to memory manager
    manager.AddEnvironment env.Name, env.tenant, env.Details, m_s_AccessToken, m_s_MainUrl, m_s_M3user, m_s_Company, m_s_Division
    
    If Not bIsCached Then Log_Activity
    
    Exit Sub
    
ErrorHandler:
    #If DEBUG_MODE Then
        Debug.Print "Tenant_Token: ERROR - " & Err.description
    #End If
    MsgBox "Error in Tenant_Token: " & Err.description, vbCritical
    ws.Range("J3").value = 0
End Sub



''
' UpdateEnvironmentCache
' Updates specific fields in the Environments sheet for a given environment.
' Pass empty strings "" for arguments you do not wish to update.
''
Public Sub UpdateEnvironmentCache(envName As String, Optional token As String = "", Optional m3user As String = "", Optional company As String = "", Optional division As String = "")
    Dim wsEnv As Worksheet
    Dim rngFound As Range
    
    On Error Resume Next
    Set wsEnv = ThisWorkbook.Sheets("Environments")
    On Error GoTo 0
    
    If wsEnv Is Nothing Then Exit Sub
    
    ' Find the environment row
    Set rngFound = wsEnv.columns("A").Find(What:=envName, LookIn:=xlValues, LookAt:=xlWhole)
    
    If Not rngFound Is Nothing Then
        On Error Resume Next
        ' Update Token (Column E) if provided
        If token <> "" Then wsEnv.Cells(rngFound.row, "E").value = token
        
        ' Update User (Column F) if provided
        If m3user <> "" Then wsEnv.Cells(rngFound.row, "F").value = m3user
        
        ' Update Company (Column G) if provided
        If company <> "" Then wsEnv.Cells(rngFound.row, "G").value = company
        
        ' Update Division (Column H) if provided
        If division <> "" Then wsEnv.Cells(rngFound.row, "H").value = division
        On Error GoTo 0
        
        #If DEBUG_MODE Then
            Debug.Print "UpdateEnvironmentCache: Updated cache for " & envName
        #End If
    Else
        #If DEBUG_MODE Then
            Debug.Print "UpdateEnvironmentCache: Warning - Environment '" & envName & "' not found."
        #End If
    End If
End Sub











Sub TEST_CalculateBase64()
    Dim userName As String
    Dim Password As String
    Dim encodedString As String

    ' Define username And password
    userName = "inforbc\ctosqadata"
    Password = "cNLkmiOjs6cfCbKe"

    ' Concatenate username And password With a colon
    encodedString = userName & ":" & Password

    ' Encode the concatenated string To Base64
    encodedString = Base64EncodeVBA(encodedString)

    ' Display the result
    MsgBox "Base64 encoded string: " & encodedString
End Sub

Private Sub UpdateG5Formula()
    If Application.WorksheetFunction.isError(m_obj_ws.Range("G5")) Or m_obj_ws.Range("G5").value = "" Then
        With m_obj_ws.Range("G5")
            .NumberFormat = "General"
            .Formula = "=IFNA(VLOOKUP(Transaction,Transactions!B:C,2,False),"""")"
        End With
        m_obj_ws.Calculate
        DoEvents
    End If
End Sub

Sub UpdateVersion()
    Dim currentVersion As String
    Dim targetCell As Range
    
    Set targetCell = ActiveSheet.Range("J2")
    currentVersion = targetCell.value
    If currentVersion <> DOPPIO_VERSION Then
        targetCell.value = DOPPIO_VERSION
    End If
End Sub

Function UrlEncode(str As String) As String
    Dim i As Integer
    Dim c As String
    Dim encoded As String
    encoded = ""
    
    For i = 1 To Len(str)
        c = Mid(str, i, 1)
        Select Case c
        Case "0" To "9", "A" To "Z", "a" To "z", "-", "_", ".", "~"
            encoded = encoded & c
        Case " "
            encoded = encoded & "%20"
        Case Else
            ' Ignore (remove) any other special characters
            ' No action taken for characters outside the allowed set
        End Select
    Next i
    
    UrlEncode = encoded
End Function

' Custom Function To write data To a file
Sub WriteFile(filePath As String, data As String)
    Dim fileNumber As Integer
    fileNumber = FreeFile

    ' Open the file For writing
    Open filePath For Output As fileNumber

    ' Write the data To the file
    Print #fileNumber, data

    ' Close the file
    Close fileNumber
End Sub

Sub RecordCache_Dump()
    Dim cacheSheet As Worksheet
    Dim ws As Worksheet
    Dim i As Long, lastRow As Long
    Dim pair As Variant
    Dim dataArr() As Variant
    Dim buttonNames, buttonActions, buttonPositions As Variant
    Dim exists As Boolean
    Dim shp As Shape
    Dim btn As Object
    Dim recordCount As Long
    
    ' Optimize Performance
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    ' Initialize cache if needed
    If m_RecordCache Is Nothing Then RecordCache_Initialize

    ' Check if "Cache" sheet exists, else create it
    On Error Resume Next
    Set cacheSheet = ThisWorkbook.Sheets("Cache")
    If cacheSheet Is Nothing Then
        Set cacheSheet = ThisWorkbook.Sheets.Add
        cacheSheet.Name = "Cache"
        cacheSheet.Tab.Color = vbBlack
    End If
    On Error GoTo 0
    
    ' If "Cache" is active, clear and set headers
    If ActiveSheet.Name = "Cache" Then
        cacheSheet.Cells.Clear
        cacheSheet.Rows(1).RowHeight = 97

        ' Write headers
        cacheSheet.Cells(1, 1).value = "Key (URL)"
        cacheSheet.Cells(1, 2).value = "Value (Serialized JSON)"
        cacheSheet.Cells(1, 1).Font.Bold = True
        cacheSheet.Cells(1, 2).Font.Bold = True
        
        ' Set column widths
        cacheSheet.columns(1).ColumnWidth = 70
        cacheSheet.columns(2).ColumnWidth = 70
        
        ' Define button properties
        buttonNames = Array("Dump Cache", "Load", "Reset", "Close")
        buttonActions = Array("RecordCache_Dump", "RecordCache_Load", "RecordCache_Reset", "RecordCache_Close")
        buttonPositions = Array(28, 112, 197, 281)

        ' Loop through buttons and add if not exists
        For i = LBound(buttonNames) To UBound(buttonNames)
            exists = False
            
            ' Check if button exists
            For Each shp In cacheSheet.Shapes
                If shp.Type = msoFormControl Then
                    If LCase(shp.Name) = LCase(buttonNames(i)) Then
                        exists = True
                        Exit For
                    End If
                End If
            Next shp
            
            ' Add button only if it does not exist
            If Not exists Then
                Set btn = cacheSheet.Buttons.Add(buttonPositions(i), 26, 71, 30)
                btn.Caption = buttonNames(i)
                btn.OnAction = buttonActions(i)
            End If
        Next i
    End If

    ' Exit if recordCache is empty
    recordCount = m_RecordCache.count
    If recordCount = 0 Then GoTo Cleanup
    
    ' Prepare data array for faster writing
    ReDim dataArr(1 To recordCount, 1 To 2)
    For i = 1 To recordCount
        pair = Split(m_RecordCache.item(i), "|")
        dataArr(i, 1) = pair(0) ' Cache Key
        dataArr(i, 2) = pair(1) ' Cached Value (JSON)
    Next i

    ' Write data to sheet in bulk
    cacheSheet.Range("A2").Resize(recordCount, 2).value = dataArr

    ' Freeze Top Row if Active
    If ActiveSheet Is cacheSheet Then
        cacheSheet.Rows(2).Select
        ActiveWindow.FreezePanes = True
    End If

    ' Enable AutoFilter
    If Not cacheSheet.AutoFilterMode Then cacheSheet.Rows(1).AutoFilter

    ' Sort Data
    With cacheSheet.Sort
        .SortFields.Clear
        .SortFields.Add key:=cacheSheet.columns(1), Order:=xlAscending
        .SetRange cacheSheet.Range("A1:B" & recordCount + 1)
        .Header = xlYes
        .Apply
    End With

Cleanup:
    ' Restore Application Settings
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
End Sub

Sub ZZZ_RecordCache_Load()
    Dim cacheSheet As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim key As String
    Dim value As String
    Dim pair As Variant
    
    ' Ensure the recordCache is initialized
    If m_RecordCache Is Nothing Then
        RecordCache_Initialize
    End If

    ' Check if the "Cache" sheet exists
    On Error Resume Next
    Set cacheSheet = ThisWorkbook.Sheets("Cache")
    If cacheSheet Is Nothing Then
        Exit Sub
    End If
    On Error GoTo 0

    ' Check if the sheet contains any data
    lastRow = cacheSheet.Cells(cacheSheet.Rows.count, 1).End(xlUp).row
    If lastRow < 2 Then
        Exit Sub
    End If

    ' Load data from the sheet into recordCache, avoiding duplicates
    For i = 2 To lastRow
        key = cacheSheet.Cells(i, 1).value
        value = cacheSheet.Cells(i, 2).value
        
        Dim found As Boolean
        Dim j As Long
        found = False
        For j = 1 To m_RecordCache.count
            pair = Split(m_RecordCache.item(j), "|")
            If pair(0) = key Then
                found = True
                Exit For
            End If
        Next j
        
        If key <> "" And value <> "" And Not found Then
            m_RecordCache.Add key & "|" & value
        End If
    Next i
End Sub



Function RecordCache_Find(saak As String) As Long
    Dim i As Long
    Dim pair As Variant
    
    If m_RecordCache Is Nothing Then
        RecordCache_Initialize
    End If
    
    For i = 1 To m_RecordCache.count
        pair = Split(m_RecordCache.item(i), "|")
        If pair(0) = saak Then
            RecordCache_Find = i
            Exit Function
        End If
    Next i
    RecordCache_Find = 0
End Function

Sub RecordCache_Initialize()
    If m_RecordCache Is Nothing Then
        Set m_RecordCache = New Collection
        RecordCache_Load
    End If
End Sub

Sub RecordCache_Display()
    Dim ws As Worksheet
    Set ws = Sheets("Cache")
    
    With ws
        .Visible = Not .Visible
        If .Visible = xlSheetVisible Then
            .Activate
            Sheets("Cache").Activate
            With ActiveWindow
                If .FreezePanes Then .FreezePanes = False
                .SplitColumn = 1
                .SplitRow = 1
                .FreezePanes = True
            End With
            Application.GoTo Reference:="R7C4", Scroll:=False
        End If
    End With
End Sub

Sub RecordCache_Retreive(url As String, found As Boolean)
    Dim i As Long
    Dim pair() As String
    Dim test_url As String

    found = False
    
    If Not m_RecordCache Is Nothing Then
        If m_RecordCache.count > 0 Then
            For i = m_RecordCache.count To 1 Step -1 ' Loop backwards to safely remove items
                pair = Split(m_RecordCache.item(i), "|")
                test_url = Replace(url, ";maxrecs=0", "")
        
                If pair(0) = test_url Then
                    ' Check if the value is an error message
                    If Left$(Trim(pair(1)), 9) = "{""error"":" Then
                        m_RecordCache.Remove i
                        Exit For
                    End If
        
                    On Error Resume Next
                    Set m_obj_Records = JsonConverter.ParseJson(pair(1))
                    If Err.Number <> 0 Then
                        m_RecordCache.Remove i
                    Else
                        found = True
                    End If
                    On Error GoTo 0
                    Exit For
                End If
            Next i
        End If
    End If

End Sub

Sub RecordCache_Store(url As String)
    Dim serializedRecords As String
    Dim recordPair As String
    RecordCache_Initialize
    
    If Not m_obj_Records Is Nothing Then
        serializedRecords = JsonConverter.ConvertToJson(m_obj_Records)
        url = Replace(url, ";maxrecs=0", "")
        recordPair = url & "|" & serializedRecords
        m_RecordCache.Add recordPair
        RecordCache_Dump
        Exit Sub
    End If
    If Not m_obj_JsonResponse Is Nothing Then
        serializedRecords = JsonConverter.ConvertToJson(m_obj_JsonResponse)
        url = Replace(url, ";maxrecs=0", "")
        recordPair = url & "|" & serializedRecords
        m_RecordCache.Add recordPair
        RecordCache_Dump
    End If
End Sub

Sub IDM_Load_Cache()
    Dim jsonData As Object
    Dim methodDetails As Object
    Dim jsonArray As Object
    Dim jsonValue As Object
    Dim parametersJson As String
    Dim cleanedString As String
    Dim rowNum As Long
    Dim i As Long
    Dim resultSheet As Worksheet

    Set resultSheet = ThisWorkbook.Sheets("Transactions")
    Set jsonData = m_obj_Records
    
    rowNum = 3
    For Each methodDetails In jsonData
        With methodDetails
            resultSheet.Cells(rowNum, 1).value = .item("api")
            resultSheet.Cells(rowNum, 2).value = .item("transaction")
            resultSheet.Cells(rowNum, 3).value = .item("method")
            resultSheet.Cells(rowNum, 4).value = .item("summary")
            If .exists("parameters") Then
                parametersJson = JsonConverter.ConvertToJson(.item("parameters"))
                cleanedString = Replace(parametersJson, "\\\", "~")
                cleanedString = Replace(cleanedString, "\", "")
                cleanedString = Replace(cleanedString, "~", "\")
                cleanedString = Mid(cleanedString, 2, Len(cleanedString) - 2)
                resultSheet.Cells(rowNum, 5).value = cleanedString
            End If
        End With
        rowNum = rowNum + 1
    Next methodDetails
End Sub

Function GetFolderPath() As String
    ' for MacOS if you are having directory authority issues
    Dim folderPath As String
    On Error Resume Next
    folderPath = MacScript("choose folder as string")
    If Err.Number <> 0 Then
        folderPath = ""
        Err.Clear
    End If
    On Error GoTo 0
    GetFolderPath = folderPath
End Function


Sub Python_GenerateFunction()
    Dim sql As String
    Dim tableName As String
    Dim columnList As String
    Dim whereclause As String
    Dim pyCode As String
    Dim columns As Variant
    Dim i As Integer

    sql = Range("B6").value
    If InStr(sql, "*") > 0 Then
        MsgBox "Wildcard '*' found in field list. Please specify columns explicitly.", vbExclamation, "Invalid Field List"
        Exit Sub
    End If

    Dim selectPos As Long: selectPos = InStr(1, sql, "select", vbTextCompare) + 6
    Dim fromPos As Long: fromPos = InStr(1, sql, "from", vbTextCompare)
    Dim wherePos As Long: wherePos = InStr(1, sql, "where", vbTextCompare)
    If wherePos = 0 Then wherePos = Len(sql) + 1

    columnList = Trim(Mid(sql, selectPos, fromPos - selectPos))
    tableName = Trim(Mid(sql, fromPos + 4, wherePos - fromPos - 4))
    If wherePos <= Len(sql) Then
        whereclause = Trim(Mid(sql, wherePos + 5))
    Else
        whereclause = ""
    End If

    columns = Split(columnList, ",")

    pyCode = "def extract_" & tableName & "(cur):" & vbCr
    pyCode = pyCode & "    print(""Processing " & tableName & """)" & vbCr & vbCr

    pyCode = pyCode & "    # Drop the table if it exists" & vbCr
    pyCode = pyCode & "    cur.execute(""DROP TABLE IF EXISTS M3_" & tableName & """)" & vbCr & vbCr

    pyCode = pyCode & "    # Create the table" & vbCr
    pyCode = pyCode & "    create_table_sql = """"""" & vbCr
    pyCode = pyCode & "    CREATE TABLE IF NOT EXISTS M3_" & tableName & " (" & vbCr
    For i = 0 To UBound(columns)
        pyCode = pyCode & "        " & Trim(columns(i)) & " TEXT"
        If i < UBound(columns) Then
            pyCode = pyCode & "," & vbCr
        Else
            pyCode = pyCode & vbCr
        End If
    Next i
    pyCode = pyCode & "    )""""""" & vbCr
    pyCode = pyCode & "    cur.execute(create_table_sql)" & vbCr & vbCr

    pyCode = pyCode & "    # Define the payload for the API request" & vbCr
    pyCode = pyCode & "    payload = {" & vbCr
    pyCode = pyCode & "        ""program"": ""EXPORTMI""," & vbCr
    pyCode = pyCode & "        ""transactions"": [{" & vbCr
    pyCode = pyCode & "            ""transaction"": ""Select""," & vbCr
    pyCode = pyCode & "            ""record"": {" & vbCr
    pyCode = pyCode & "                ""SEPC"": ""^""," & vbCr
    pyCode = pyCode & "                ""HDRS"": ""0""," & vbCr
    If whereclause <> "" Then
        pyCode = pyCode & "                ""QERY"": """ & columnList & " from " & tableName & " where " & whereclause & """" & vbCr
    Else
        pyCode = pyCode & "                ""QERY"": """ & columnList & " from " & tableName & """" & vbCr
    End If
    pyCode = pyCode & "            }," & vbCr
    pyCode = pyCode & "            ""selectedColumns"": [""REPL""]" & vbCr
    pyCode = pyCode & "        }]" & vbCr
    pyCode = pyCode & "    }" & vbCr & vbCr

    pyCode = pyCode & "    # Insert data into the table" & vbCr
    pyCode = pyCode & "    insert_sql = ""INSERT INTO M3_" & tableName & " (" & columnList & ") VALUES ("
    For i = 1 To UBound(columns)
        pyCode = pyCode & "?, "
    Next i
    pyCode = pyCode & "?)""" & vbCr

    pyCode = pyCode & "    with requests.Session() as session:" & vbCr
    pyCode = pyCode & "        data = post_to_m3(payload, session)" & vbCr
    pyCode = pyCode & "    records = data['results'][0]['records']" & vbCr
    pyCode = pyCode & "    for record in tqdm(records, desc=""Inserting " & tableName & """):" & vbCr
    pyCode = pyCode & "        cur.execute(insert_sql, record[""REPL""].split(""^"")[:" & UBound(columns) + 1 & "])" & vbCr
    pyCode = pyCode & "    cur.connection.commit()" & vbCr

    SampleRESTPopup pyCode, "Python Sample"
End Sub


Sub ValidateSelectedEnvironment()
    Dim selectedEnv As String
    Dim envCell As Range
    Dim isValid As Boolean
    Dim wsEnv As Worksheet
    Dim firstEnv As String

    On Error GoTo ExitSub
    selectedEnv = ActiveSheet.Range("Environment").value
    Set wsEnv = ThisWorkbook.Sheets("Environments")
    isValid = False

    ' Loop through Environments!A:A
    For Each envCell In wsEnv.Range("A1", wsEnv.Cells(wsEnv.Rows.count, "A").End(xlUp))
        If Trim(envCell.value) <> "" Then
            If firstEnv = "" Then firstEnv = envCell.value
            If envCell.value = selectedEnv Then
                isValid = True
                Exit For
            End If
        End If
    Next envCell

    ' If not valid, set to first value
    If Not isValid Then
        ActiveSheet.Range("Environment").value = firstEnv
    End If

ExitSub:
End Sub

Sub RunBuildMatrix()
    Dim ws As Worksheet
    Set ws = ActiveSheet

    Set MatrixBuilder = New MatrixManager
    MatrixBuilder.BuildMatrix ws
'    pivoted = True
End Sub

Sub RunUnpivotMatrix()
    Dim ws As Worksheet
    
'    If Not pivoted Then Exit Sub
    
    Set ws = ActiveSheet

    If MatrixBuilder Is Nothing Then
'        MsgBox "You must build the matrix first.", vbExclamation
'        Exit Sub
        Set MatrixBuilder = New MatrixManager
    End If

    MatrixBuilder.RefreshMatrix ws, ws.Range("B8")
    MatrixBuilder.Unpivot ws
'    pivoted = False

    AutoFit_Click
    AutoFit_Click
End Sub
Function GetUserSessionInfo() As String
    Dim userName As String, userDomain As String, computerName As String
    Dim userProfile As String, osName As String, osVersion As String
    Dim localIP As String, publicIP As String
    Dim jsonText As String

    On Error Resume Next

    #If Mac Then
        userName = MacScript("do shell script ""whoami""")
        userDomain = "N/A"
        computerName = MacScript("do shell script ""scutil --get ComputerName""")
        userProfile = Environ("HOME")
        osName = "macOS"
        osVersion = MacScript("do shell script ""sw_vers -productVersion""")
        localIP = MacScript("do shell script ""ipconfig getifaddr en0 || echo Unavailable""")
    #Else
        Dim netObj As Object
        Set netObj = CreateObject("WScript.Network")
        userName = netObj.userName
        userDomain = netObj.userDomain
        computerName = netObj.computerName
        userProfile = Environ("USERPROFILE")
        localIP = GetLocalIPAddress()
        osName = GetOSName()
        osVersion = GetOSVersion()
    #End If

    publicIP = GetPublicIPAddress()

    jsonText = "{""userName"":""" & userName & """, " & _
               """userDomain"":""" & userDomain & """, " & _
               """computerName"":""" & computerName & """, " & _
               """localIP"":""" & localIP & """, " & _
               """publicIP"":""" & publicIP & """, " & _
               """userProfile"":""" & userProfile & """, " & _
               """osName"":""" & osName & """, " & _
               """osVersion"":""" & osVersion & """, " & _
               """sheetVersion"":""" & DOPPIO_VERSION & """}"

    GetUserSessionInfo = jsonText
End Function

Function GetLocalIPAddress() As String
    On Error GoTo ErrHandler
    Dim objWMIService As Object, colItems As Object, objItem As Object
    Set objWMIService = GetObject("winmgmts:\\.\root\cimv2")
    Set colItems = objWMIService.ExecQuery("Select * from Win32_NetworkAdapterConfiguration Where IPEnabled = True")

    For Each objItem In colItems
        If Not IsNull(objItem.ipAddress) Then
            GetLocalIPAddress = objItem.ipAddress(0)
            Exit Function
        End If
    Next
ErrHandler:
    GetLocalIPAddress = "Unavailable"
End Function

Function GetOSName() As String
    Dim objWMIService As Object, colOS As Object, objOS As Object
    On Error Resume Next
    Set objWMIService = GetObject("winmgmts:\\.\root\cimv2")
    Set colOS = objWMIService.ExecQuery("Select * from Win32_OperatingSystem")

    For Each objOS In colOS
        GetOSName = objOS.Caption
        Exit For
    Next
End Function

Function GetOSVersion() As String
    Dim objWMIService As Object, colOS As Object, objOS As Object
    On Error Resume Next
    Set objWMIService = GetObject("winmgmts:\\.\root\cimv2")
    Set colOS = objWMIService.ExecQuery("Select * from Win32_OperatingSystem")

    For Each objOS In colOS
        GetOSVersion = objOS.VERSION
        Exit For
    Next
End Function

Function GetPublicIPAddress() As String
    On Error GoTo Fail

    #If Mac Then
        GetPublicIPAddress = MacScript("do shell script ""curl -s https://api.ipify.org""")
    #Else
        Dim http As Object
        Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
        http.Open "GET", "https://api.ipify.org", False
        http.Send
        GetPublicIPAddress = http.responseText
    #End If
    Exit Function

Fail:
    GetPublicIPAddress = "Unavailable"
End Function
Sub DecodeBase64InCell()
    Dim cell As Range
    Dim encodedText As String
    Dim decodedText As String
    
    ' Set target cell (defaults to currently active cell)
    Set cell = ActiveCell
    
    ' Read base64 string from the current cell
    encodedText = Trim(cell.value)
    
    ' Skip empty cells
    If encodedText = "" Then
        MsgBox "The selected cell is empty.", vbExclamation
        Exit Sub
    End If
    
    ' Decode and output to the same cell
    On Error GoTo DecodeError
    decodedText = Base64DecodeVBA(encodedText)
    
    cell.value = decodedText
    cell.WrapText = False
    ToggleBase64Button ("Decode")
    
    Exit Sub
    
DecodeError:
    ToggleBase64Button ("Decode")
    MsgBox "Failed to decode Base64 string: " & Err.description, vbCritical
End Sub
Sub EncodeBase64InCell()
    Dim cell As Range
    Dim plainText As String
    Dim encodedText As String

    Set cell = ActiveCell
    plainText = Trim(cell.value)

    If plainText = "" Then
        MsgBox "The selected cell is empty.", vbExclamation
        Exit Sub
    End If

    On Error GoTo EncodeError
    encodedText = Base64EncodeVBA(plainText)
    
    cell.value = encodedText
    cell.WrapText = False
    ToggleBase64Button ("Encode")
    
    Exit Sub

EncodeError:
    ToggleBase64Button ("Encode")
    MsgBox "Failed to encode string: " & Err.description, vbCritical
End Sub
Sub ToggleBase64Button(btnType As String)
    Dim ws As Worksheet
    Dim btn As Button
    Dim btnName As String
    Dim found As Boolean

    btnName = btnType
    Set ws = ActiveSheet
    found = False

    ' Check for existing button
    For Each btn In ws.Buttons
        If btn.Name = btnName Then
            btn.Delete
            found = True
            Exit Sub
        End If
    Next btn

    ' If not found, create it
    If Not found Then
        #If Mac Then
            Set btn = ws.Buttons.Add(8, 113, 69, 29)
        #Else
            Set btn = ws.Buttons.Add(8, 107, 69, 27)
        #End If
        
        With btn
            .Caption = btnType
            If btnType = "Decode" Then
                .OnAction = "DecodeBase64InCell"
            ElseIf btnType = "Encode" Then
                .OnAction = "EncodeBase64InCell"
            End If
            .Name = btnName
        End With
    End If
End Sub
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
    
    'Doppio.Tenant_Information
    Tenant_Token
    ClearStatus
    KillPleaseWait
    
    Set wbSource = ThisWorkbook
    
    ' Create new workbook
    Set wbNew = Workbooks.Add(xlWBATWorksheet)
    
    ' Add Control sheet at the start
    Set wsControl = wbNew.Sheets(1)
    wsControl.Name = "Control"
    wsControl.Range("A1").value = "Worksheet"
    wsControl.Range("B1").value = "Description"
    wsControl.Range("C1").value = "Data"
    r = 2 ' Row counter for Control sheet
    Dim firstAPIName As String
    firstAPIName = ""
    
    ' Loop through all visible sheets in source workbook
    For Each wsSource In wbSource.Sheets
        If wsSource.Visible = xlSheetVisible Then
            
            ' Capture first visible sheet's API name for file naming
            If firstAPIName = "" Then
                firstAPIName = wsSource.Range("A2").value
            End If
            
            ' Generate names from values in A2 and G4
            sheetName = "API_" & wsSource.Range("A2").value & "_" & wsSource.Range("G4").value
            
            ' Check if sheet already exists in wbNew
            Dim exists As Boolean
            exists = False
            Dim sh As Worksheet
            For Each sh In wbNew.Sheets
                If LCase(sh.Name) = LCase(sheetName) Then
                    exists = True
                    Exit For
                End If
            Next sh
            
            If Not exists Then
                ' Create a new sheet in target workbook
                Set wsNew = wbNew.Sheets.Add(After:=wbNew.Sheets(wbNew.Sheets.count))
                wsNew.Name = sheetName
                
                ' Ensure no freeze panes
                wsNew.Activate
                If wsNew.Parent.Windows(1).FreezePanes Then
                    wsNew.Parent.Windows(1).FreezePanes = False
                End If
                
                ' Find the last column with data in Row 8
                lastCol = wsSource.Cells(8, wsSource.columns.count).End(xlToLeft).column
                
                ' Row 1 from Row 8
                For i = 1 To lastCol
                    wsNew.Cells(1, i).value = wsSource.Cells(8, i).value
                Next i
                
                ' Row 2 from Row 7 with formatting
                For i = 1 To lastCol
                    rawVal = wsSource.Cells(7, i).value
                    rawVal = Replace(rawVal, vbCrLf, vbLf)
                    rawVal = Replace(rawVal, vbCr, vbLf)
                    lines = Split(rawVal, vbLf)
                    
                    If UBound(lines) >= 1 Then
                        label = Trim(lines(0))
                        ref = Trim(lines(1))
                        If Len(ref) >= 2 Then
                            colLetter = Left(ref, 1)
                            rowNum = Mid(ref, 2)
                            wsNew.Cells(2, i).value = label & " (" & colLetter & ":" & rowNum & ")"
                        Else
                            wsNew.Cells(2, i).value = label
                        End If
                    Else
                        wsNew.Cells(2, i).value = Trim(rawVal)
                    End If
                Next i
                
                ' Row 3 values: "no" in A3, "yes" elsewhere
                wsNew.Cells(3, 1).value = "no"
                For i = 2 To lastCol
                    wsNew.Cells(3, i).value = "yes"
                Next i
                
                ' Insert "MESSAGE" in A1
                wsNew.Cells(1, 1).value = "MESSAGE"
                
                ' Copy data from Row 9 onwards
                wsSource.Range(wsSource.Cells(9, 2), wsSource.Cells(wsSource.UsedRange.Rows.count, lastCol)).Copy _
                    Destination:=wsNew.Cells(4, 2)
                
                ' Add entry to Control sheet
                wsControl.Cells(r, 1).value = sheetName
                wsControl.Cells(r, 2).value = wsSource.Range("D2").value
                wsControl.Cells(r, 3).value = "x"
                r = r + 1
                
                wsNew.columns.AutoFit
            End If
        End If
    Next wsSource
    
    wsControl.columns.AutoFit
    
    ' Remove the default empty first sheet if still present
    On Error Resume Next
    If wbNew.Sheets(1).Name = "Sheet1" Then wbNew.Sheets(1).Delete
    On Error GoTo 0
    
    ' Save workbook
    wbName = "API_" & firstAPIName
    '& "_" & Format(Now, "yyyymmdd_hhmmss") & ".xlsx"
    #If Mac Then
        filePath = Application.GetSaveAsFilename(InitialFileName:=wbName)
    #Else
        filePath = Application.GetSaveAsFilename(InitialFileName:=wbName, FileFilter:="Excel Files (*.xlsx), *.xlsx")
    #End If
    
    Dim statusCode As String
    Dim executeBody As String
    Dim script As String
    Dim uploadUrl As String
    Dim body As String
    Dim tempFilePath As String
    tempFilePath = Environ("HOME") & "/curl_output.txt"
    If filePath <> "False" Then
        Application.DisplayAlerts = False
        wbNew.SaveAs fileName:=filePath, FileFormat:=xlOpenXMLWorkbook
        Application.DisplayAlerts = True
        
        ' === Prompt to upload ===
        If MsgBox("Do you want to send this file to M3?", vbYesNo + vbQuestion, "Upload to M3") = vbYes Then
            wbNew.Close SaveChanges:=True
            uploadUrl = "file/FileImport/" & Dir(filePath)
            body = "@" & filePath
            Curl_Build m_s_MainUrl, "/M3/foundation-rest/file-management/v1", uploadUrl, body, "FileMng", script
            #If Mac Then
                ExecuteScriptWithRetry (script)
                statusCode = ReadFileToString(tempFilePath)
            #Else
                statusCode = ParseAndExecuteCurl_Regex(script)
            #End If
            
            ' === Prompt to process ===
            If statusCode = "201" Then
                If MsgBox("Upload successful. Do you want to process this file in M3 (EVS100MI.ImportFile)?", vbYesNo + vbQuestion, "Process in M3") = vbYes Then
                    executeBody = "{""program"":""EVS100MI"",""transactions"":[{""transaction"":""ImportFile"",""record"":{""FNAM"":""" & Dir(filePath) & """},""selectedColumns"":[""FNAM""]}]}"
                    Curl_Build m_s_MainUrl, m_s_MiPath, "", executeBody, "API", script
                    #If Mac Then
                        ExecuteScriptWithRetry (script)
                        statusCode = ReadFileToString(tempFilePath)
                    #Else
                        ParseAndExecuteCurl_Regex (script)
                    #End If
                End If
            End If
        End If
    Else
        MsgBox "Save canceled. Workbook was not saved.", vbExclamation
    End If
End Sub
Function ReadFileToString(filePath As String) As String
    Dim fileNumber As Integer
    Dim fileContent As String
    
    On Error GoTo ErrHandler
    
    fileNumber = FreeFile
    Open filePath For Input As fileNumber
        fileContent = Input$(LOF(fileNumber), fileNumber)
    Close fileNumber
    
    ReadFileToString = fileContent
    Exit Function
    
ErrHandler:
    ' Return empty string on error
    ReadFileToString = ""
    On Error GoTo 0
End Function
Function ParseAndExecuteCurl_Regex(script As String) As Long
    Dim curlCommand As String
    Dim method As String, url As String, bodyFile As String, bodyRaw As String
    Dim headers As Object, headerMatches As Object, match As Object
    Dim http As Object
    Dim result As String
    Dim fileStream As Object, fileData() As Byte
    Dim regEx As Object
    
    ' Your corrected curl command
    'Debug.Print script
    curlCommand = script
    
    ' Create dictionary for headers
    Set headers = CreateObject("Scripting.Dictionary")
    
    ' Extract method
    Set regEx = CreateObject("VBScript.RegExp")
    regEx.Pattern = "--request\s+(\w+)"
    regEx.IgnoreCase = True
    regEx.Global = False
    If regEx.Test(curlCommand) Then
        method = regEx.Execute(curlCommand)(0).SubMatches(0)
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
        Dim pos As Long
        pos = InStr(match.SubMatches(0), ":")
        If pos > 0 Then
            headers(Trim(Left(match.SubMatches(0), pos - 1))) = Trim(Mid(match.SubMatches(0), pos + 1))
        End If
    Next
    
    ' Extract data-binary file
    regEx.Pattern = "--data-binary\s+'@([^']+)'"
    If regEx.Test(curlCommand) Then
        bodyFile = regEx.Execute(curlCommand)(0).SubMatches(0)
        bodyFile = Replace(bodyFile, "\\\\", "\")
    End If
    
    ' Extract data-raw body
    regEx.Pattern = "--data-raw\s+'([^']+)'"
    If regEx.Test(curlCommand) Then
        bodyRaw = regEx.Execute(curlCommand)(0).SubMatches(0)
        bodyRaw = Replace(bodyRaw, "\""", """") ' convert \" to "
    End If
    
    ' === Prepare HTTP request ===
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    http.Open method, url, False
    
    ' Apply headers
    Dim hdr As Variant
    For Each hdr In headers.keys
        http.setRequestHeader hdr, headers(hdr)
    Next
    
    ' === Send body if present ===
    If Len(bodyFile) > 0 Then
        Set fileStream = CreateObject("ADODB.Stream")
        fileStream.Type = 1 'binary
        fileStream.Open
        fileStream.LoadFromFile bodyFile
        fileData = fileStream.Read
        fileStream.Close
        http.Send fileData
    ElseIf Len(bodyRaw) > 0 Then
        http.Send bodyRaw
    Else
        http.Send
    End If
    
    ' Output response
    'Debug.Print "Status: " & http.Status
    'Debug.Print "Response: " & http.responseText
    
    ParseAndExecuteCurl_Regex = http.status
End Function

''
' Load tenant information from Environments sheet
''
Sub Tenant_Information()
    Dim SettingsSheet As Worksheet
    Dim environmentRange As Range
    Dim targetCell As Range
    Dim jsonString As String
    Dim json As Object
    Dim User As String, Password As String

    Set SettingsSheet = ThisWorkbook.Sheets("Environments")
    Set environmentRange = SettingsSheet.Range("A:A")

    m_s_SelectedEnvironment = ActiveSheet.Range("Environment").value
    If m_s_SelectedEnvironment = "" Then
        ClearFields
    Else
        ChangeCellColorBasedOnEnvironment

        Set targetCell = environmentRange.Find(What:=m_s_SelectedEnvironment, LookIn:=xlValues, LookAt:=xlWhole)

        If Not targetCell Is Nothing Then
            jsonString = targetCell.Offset(0, 1).value
            m_s_M3user = targetCell.Offset(0, 3).value
            Set json = JsonConverter.ParseJson(jsonString)
            encodedTenant = jsonString

            ti = json.item("ti")
            ci = json.item("ci")
            cs = json.item("cs")
            iu = json.item("iu")
            pu = json.item("pu")
            oa = json.item("oa")
            ot = json.item("ot")
            ru = json.item("ru")
            saak = json.item("saak")
            sask = json.item("sask")
            User = json.item("user")
            Password = json.item("password")
            m_s_stUrl = json.item("url")
        End If
    End If

    m_b_Multitenant = True
    If m_s_stUrl <> "" Then
        m_b_Multitenant = False
        m_s_stToken = Base64EncodeVBA(User & ":" & Password)
    End If
    m_s_MainUrl = iu & "/" & ti

    If manager Is Nothing Then
        Set manager = New EnvironmentManager
    End If
    
    Dim env As Environment
    Set env = manager.GetEnvironment(m_s_SelectedEnvironment)
    If Not manager.HasEnvironment(m_s_SelectedEnvironment) Then
        manager.AddEnvironment m_s_SelectedEnvironment, m_s_SelectedEnvironment, encodedTenant, "", m_s_MainUrl, m_s_M3user, "", ""
    ElseIf env.User = "" Then
        manager.AddEnvironment m_s_SelectedEnvironment, m_s_SelectedEnvironment, encodedTenant, "", m_s_MainUrl, m_s_M3user, "", ""
    End If
End Sub


Sub TestEnvironmentManager()
    Dim testManager As EnvironmentManager
    Dim env As Environment

    Set testManager = New EnvironmentManager
    testManager.AddEnvironment "TestEnv1", "https://api.example.com1", "OAuthToken1", "", "", "", "", ""
    Set env = testManager.GetEnvironment("TestEnv1")

    If Not env Is Nothing Then
        MsgBox "URL: " & env.Name & vbCrLf & "Token: " & env.Details
    End If
End Sub
Public Sub ClearEnvironmentTokens(Optional envName As String = "")
    ' Clear module-level tokens
    m_s_AccessToken = ""
    m_s_RefreshToken = ""
    m_s_TokenType = ""
    activeEnvironment = ""
    m_s_SelectedEnvironment = ""
    
    ' Clear manager cache
    If Not manager Is Nothing Then
        If envName <> "" Then
            manager.ClearEnvironment envName
        End If
    End If
End Sub



