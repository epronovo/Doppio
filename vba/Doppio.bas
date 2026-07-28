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
Public m_b_TokenAttemptedThisCycle As Boolean   ' Prevents repeated token attempts in one event cycle
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
Public m_b_AutoFitToggle As Boolean  ' False = use row 7, True = use row 9
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
Public maxRecs As Long
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
Public sheetNaming As Integer
Public manager As EnvironmentManager
Public MatrixBuilder As MatrixManager
Public pivoted As Boolean
Public currentPid As String
Public b_AddingSheet As Boolean  ' guard flag set True around programmatic sheet adds to suppress Workbook_NewSheet

Public m_OriginalSheet As Worksheet

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
            url = main_url & mi_path & "?maxrecs=" & maxRecs & "&extendedresult=true"
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
        ' Body is already valid JSON. Embedded in AppleScript Do shell script "..."
        ' so \ and " must be escaped for AppleScript; ' needs shell escaping.
        curlBody = body
        curlBody = Replace(curlBody, "\", "\\")    ' AppleScript: \ → \\
        curlBody = Replace(curlBody, """", "\""") ' AppleScript: " → \"
        curlBody = Replace(curlBody, "'", "'\''") ' shell: ' → '\''
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

Public Sub AutoFit_Click()
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
    UI_UpdateVersion
    
    ' =========================================================================
    ' AUTHENTICATION - Same pattern as Process_Click
    ' =========================================================================
    If m_s_AccessToken = "" Or _
       ws.Range("I3").value = "" Or _
       ws.Range("I4").value = "" Or _
       activeEnvironment <> currentEnv Then
        
        Debug.Print "AutoFit_Click_New: Need new token"
        Tenant_Token
    End If
    
    ' Get transactions
    GetTransactions_Click
    
    ' Load layout (skip for EXPORTMI with table definition)
    If Not (ws.Range("API").value = "EXPORTMI" And ws.Range("A3").value = "table:  ") Then
        Transaction_LoadLayout_New ws
    End If
    
    ' Set formulas and formatting
    SetFormulasAndFormatting_New ws
    
    'UI_UI_ShowPleaseWait "Please wait... Autofitting Columns"
    
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
    AutoFit_ColumnsAndRows False, False
    FilterRow8BasedOnPopulatedColumns ws
    
Cleanup:
    Application.ScreenUpdating = True
    UI_KillPleaseWait
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
    Dim settings As ApiSettings
    Dim prevEnableEvents As Boolean

    On Error GoTo ErrorHandler

    ' Get settings from Doppio_Config (single source of truth)
    settings = Config_ApiSettings

    prevEnableEvents = Application.EnableEvents

    ' Speed up by disabling screen updating, events, and calculations
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    
    colNum = 2
    m_l_Row = 9
    Set ws = ThisWorkbook.ActiveSheet
    lastColumn = m_obj_ColumnNames.count() + 1
    If lastColumn < ws.Cells(8, ws.Columns.count).End(xlToLeft).Column Then
        lastColumn = ws.Cells(8, ws.Columns.count).End(xlToLeft).Column
    End If
    
    If reload Then
        ReDim valuesArray(1 To 2, 1 To lastColumn + 1)
        
        If mandatory Then
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
        valuesArray = TransformArray(valuesArray, settings.conoDivi)
        lastColumn = UBound(valuesArray, 2)
        
    Else
        On Error Resume Next
        ReDim valuesArray(1 To 2, 1 To lastColumn)
        lastColumn = ws.Cells(8, ws.Columns.count).End(xlToLeft).Column
        
        For i = 2 To lastColumn
            value = ws.Cells(8, i).value
            index = m_obj_ColumnNames.IndexOf(value)
            If index > 0 Then
                valuesArray(1, i) = m_obj_ColumnDescriptions.item(index)
                valuesArray(2, i) = m_obj_ColumnNames.item(index)
                
                If m_obj_ColumnTypes.item(index) <> "A" Then
                    If formatRange Is Nothing Then
                        Set formatRange = ws.Columns(i)
                    Else
                        Set formatRange = Union(formatRange, ws.Columns(i))
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
        On Error GoTo ErrorHandler
        
        ' Apply number formatting
        If settings.formatting Then
            ws.Cells.NumberFormat = "@"
            If Not formatRange Is Nothing Then
                formatRange.NumberFormat = "General"
            End If
        End If
        
        ' Adjust row and column sizes
        With ws
            .Rows("1:6").AutoFit
            .Rows("1:6").Columns.AutoFit
            .Rows(1).RowHeight = 60
            .Rows(7).RowHeight = 36
            .Columns(1).ColumnWidth = 38
            
            min_G_Width = .Cells(7, 7).ColumnWidth
            min_I_Width = .Cells(7, 9).ColumnWidth
            
            m_b_AutoFitToggle = Not m_b_AutoFitToggle  ' flip for next call

            If m_b_AutoFitToggle Then
                ' Second call: autofit all columns based on content from row 7 downward
                .Range(.Cells(7, 1), .Cells(.UsedRange.row + .UsedRange.Rows.count - 1, .UsedRange.Column + .UsedRange.Columns.count - 1)).Columns.AutoFit
            Else
                ' First call: autofit based on row 7 descriptions (original behaviour)
                .Rows(7).WrapText = False
                .Rows("7:8").Columns.AutoFit
                .Rows(7).WrapText = True
                .Rows("7:8").Columns.AutoFit
            End If
        End With
    End If
    
    ' Write the data back to the worksheet
    ws.Range(ws.Cells(m_l_Row - 2, 1), ws.Cells(m_l_Row - 1, lastColumn)).value = valuesArray
    
    ' Enforce minimum column widths
    For i = 1 To lastColumn
        If ws.Cells(7, i).ColumnWidth < 12 Then
            ws.Columns(i).ColumnWidth = 12
        End If
        If i = 7 And ws.Cells(7, i).ColumnWidth < min_G_Width Then
            ws.Columns(i).ColumnWidth = min_G_Width
        ElseIf i = 9 And ws.Cells(7, i).ColumnWidth < min_I_Width Then
            ws.Columns(i).ColumnWidth = min_I_Width
        End If
    Next i
    
    ' Final fixed sizes
    ws.Rows(1).RowHeight = 60
    ws.Rows(7).RowHeight = 36
    ws.Columns(1).ColumnWidth = 38
    Application.GoTo Reference:="R9C2", Scroll:=True

Cleanup:
    Application.ScreenUpdating = True
    Application.EnableEvents = prevEnableEvents
    Application.Calculation = xlCalculationAutomatic
    DoEvents
    Exit Sub

ErrorHandler:
    Debug.Print "AutoFit_ColumnsAndRows: ERROR - " & Err.description
    Resume Cleanup
End Sub

Public Sub FilterRow8BasedOnPopulatedColumns(ws As Worksheet)
    Dim lastColumn As Long
    Dim dataRange As Range
    Dim colorRange As Range
    
    On Error Resume Next
    
    lastColumn = ws.Cells(8, ws.Columns.count).End(xlToLeft).Column
    
    ' If transaction is blank, clear headers
    If ws.Range("Transaction").value = "" Then
        lastColumn = 1
        ws.Rows("7:8").ClearContents
        Set colorRange = ws.Range(ws.Cells(8, lastColumn + 1), ws.Cells(8, 200))
        colorRange.Interior.Color = RGB(128, 128, 128)
    End If
    
    ws.AutoFilterMode = False
    
    If lastColumn > 1 Then
        Set dataRange = ws.Range(ws.Cells(8, 1), ws.Cells(8, lastColumn))
        dataRange.AutoFilter field:=1
        Set colorRange = ws.Range(ws.Cells(8, lastColumn + 1), ws.Cells(8, 200))
        colorRange.Interior.Color = RGB(128, 128, 128)
    End If
    
    On Error GoTo 0
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
                GetLayoutAPI url, ws
            End If
            
        Case "IPS"
            m_b_Webservice = False
            On Error Resume Next
            Application.Run "Doppio_IPS.GetLayoutWS", ws.Range("API").value, ws
            On Error GoTo ErrorHandler
            
        Case "XtendM3"
            url = ws.Range("Transaction").value
            GetLayoutM3X "/M3/extensibility/ionapi-doc", url, ws
    End Select
    
    Exit Sub
    
ErrorHandler:
    Debug.Print "Transaction_LoadLayout_New: ERROR - " & Err.description
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
    Dim prevEnableEvents As Boolean

    prevEnableEvents = Application.EnableEvents

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

    On Error Resume Next

    ' Unhide all sheets
    For Each ws In ThisWorkbook.Sheets
        ws.Visible = xlSheetVisible
    Next ws

    For Each ws In ThisWorkbook.Sheets
            If ws.name = "Logos" Then GoTo NextSheet
            
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
            If ws.name = "Environments" Then
                lastRow = 52
            End If
            ' Find the last used column
            Set lastCell = ws.Cells.Find(What:="*", After:=ws.Cells(1, 1), LookIn:=xlFormulas, LookAt:=xlPart, _
                                         SearchOrder:=xlByColumns, SearchDirection:=xlPrevious, MatchCase:=False)
            If Not lastCell Is Nothing Then
                lastCol = lastCell.Column
            Else
                lastCol = 1
            End If
            lastColLetter = Split(ws.Cells(1, lastCol).Address, "$")(1)
            If lastRow < ws.Rows.count Then
                ws.Rows(lastRow + 1 & ":" & ws.Rows.count).Delete
            End If
            ' Delete all columns to the right of the last used column
            If lastCol < ws.Columns.count Then
                ws.Columns(lastCol + 1 & ":" & ws.Columns.count).Delete
            End If
NextSheet:
        Next ws
    
    On Error GoTo 0

    Application.EnableEvents = prevEnableEvents
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
End Sub

Sub ClearFields()
    ActiveSheet.Range("User").value = ""
    ActiveSheet.Range("Company").value = ""
    ActiveSheet.Range("Division").value = ""
    ActiveSheet.Range("Transaction").value = ""
    ci = ""
    cs = ""
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
        dataEndColumn = ws.Cells(9, ws.Columns.count).End(xlToLeft).Column
        If dataEndColumn < dataStartColumn Then dataEndColumn = dataStartColumn
        
        ' Clear the contents in the identified range
        ws.Range(ws.Cells(dataStartRow, dataStartColumn), ws.Cells(dataEndRow, dataEndColumn)).ClearContents
        'Debug.Print "Cleared range: " & ws.Range(ws.Cells(dataStartRow, dataStartColumn), ws.Cells(dataEndRow, dataEndColumn)).Address
    Else
        Debug.Print "ClearOutputArea: No grey column found to clear."
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

    Set columnToCheck = m_obj_ws.Columns("A")
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
    lastColumn = ws.Cells(1, ws.Columns.count).End(xlToLeft).Column

    Dim colorRange As Range
    Set colorRange = ws.Rows(1).Columns(startColumn & ":" & lastColumn).EntireRow

    ' Change the color To your desired color (e.g., RGB(255, 0, 0) For red)
    colorRange.Interior.Color = RGB(255, 0, 0)
End Sub

Sub CreateCopyWithNewName()
    Dim newFileName As String
    newFileName = Replace(ThisWorkbook.name, ".xlsm", ".xlsx")
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

    ' Save before any changes so the original xlsm is preserved
    ThisWorkbook.Save

    DeleteHiddenSheets

    ' Switch away from the active sheet before the loop so that when the loop
    ' reaches it, ws.Activate performs a real switch and window state refreshes correctly
    Dim wsOriginal As Worksheet
    Dim wsSwap As Worksheet
    Set wsOriginal = ActiveSheet
    For Each wsSwap In ThisWorkbook.Sheets
        If wsSwap.Visible = xlSheetVisible And wsSwap.name <> wsOriginal.name Then
            wsSwap.Activate
            Exit For
        End If
    Next wsSwap

    ' Loop through all visible worksheets
    For Each ws In ThisWorkbook.Sheets

        If ws.Visible <> xlSheetVisible Then GoTo NextWs

        Debug.Print "Create_xlsx: " & ws.name

        ' Remove buttons
        For Each btn In ws.Buttons
            btn.Delete
        Next btn

        ' Remove EnvSuffix image if present
        On Error Resume Next
        ws.Shapes("EnvSuffix").Delete
        On Error GoTo 0

        ws.Activate

        ' Remove rows
        ws.Rows("2:6").Delete

        ' Copy format from B2 to A2
        ws.Range("B2").Copy
        ws.Range("A2").PasteSpecial Paste:=xlPasteFormats
        Application.CutCopyMode = False

        ' Reset freeze panes — scroll to origin first to ensure C4 lands correctly
        ActiveWindow.FreezePanes = False
        ActiveWindow.ScrollRow = 1
        ActiveWindow.ScrollColumn = 1
        ws.Range("C4").Select
        ActiveWindow.FreezePanes = True

NextWs:
    Next ws

    ' Return to the original sheet so it is the active tab in the saved file
    wsOriginal.Activate

    Dim newFileName As String
    newFileName = Replace(ThisWorkbook.name, ".xlsm", ".xlsx")

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

''
' Environments_GetUsers - Modernized
' Two passes over the rows in column B of the active sheet:
'   LOOP 1 authenticates against each tenant's SSO endpoint (ExecuteRequest)
'          and writes the tenant ID (C) and bearer access token (E).
'   LOOP 2 reuses the stored token to call MRS001MI/GetUserInfo
'          (ExecuteApiPost) and writes the M3 user ID (D).
'
' Column layout written per row:
'   C = Tenant ID (ti)         -- loop 1
'   E = Bearer access token    -- loop 1
'   D = M3 User ID (ZZUSID) or Full Name (USFN) if ZZUSID is blank -- loop 2
''
Sub Environments_GetUsers()
    Dim ws              As Worksheet
    Dim jsonString      As String
    Dim json            As Object
    Dim lastRow         As Long
    Dim i               As Long
    Dim tenantId        As String
    Dim clientId        As String
    Dim clientSecret    As String
    Dim instanceUrl     As String
    Dim ssoBase         As String
    Dim tokenEndpoint   As String
    Dim saak            As String
    Dim sask            As String
    Dim tokenUrl        As String
    Dim tokenBody       As String
    Dim miBody          As String
    Dim config          As httpConfig
    Dim httpResp        As httpResponse
    Dim results         As Object
    Dim resultItem      As Object
    Dim records         As Object
    Dim record          As Object

    On Error GoTo ErrorHandler

    Set ws = ActiveSheet
    lastRow = ws.Cells(ws.Rows.count, "B").End(xlUp).row

    ' --- Clear previous results in columns C, D and E ---
    If lastRow >= 1 Then ws.Range("C1:E" & lastRow).ClearContents

    ' Suppress the 401 re-auth prompt for the whole routine; this batch handles
    ' auth itself (a bad token just blanks that row's user).
    g_b_SuppressAuthPrompt = True

    ' =========================================================================
    ' LOOP 1: Get the bearer token for every row and write tenant ID (C) + token (E)
    ' =========================================================================
    For i = 1 To lastRow
        jsonString = ws.Cells(i, 2).value
        If jsonString = "" Then Exit For

        ' --- Parse tenant credentials from column B ---
        Set json = ParseJson(jsonString)

        tenantId = json.item("ti")
        clientId = json.item("ci")
        clientSecret = json.item("cs")
        ssoBase = json.item("pu")
        tokenEndpoint = json.item("ot")
        saak = json.item("saak")
        sask = json.item("sask")

        ' --- Get Access Token ---
        tokenUrl = ssoBase & tokenEndpoint
        tokenBody = "client_id=" & Core_UrlEncode(clientId) & _
                    "&client_secret=" & Core_UrlEncode(clientSecret) & _
                    "&grant_type=password" & _
                    "&username=" & Core_UrlEncode(saak) & _
                    "&password=" & Core_UrlEncode(sask)

        config.url = tokenUrl
        config.method = HttpMethod_POST
        config.contentType = "application/x-www-form-urlencoded"
        config.AcceptType = "application/json"
        config.authHeader = ""
        config.body = tokenBody
        config.timeoutSeconds = 30

        httpResp = ExecuteRequest(config)

        If Not httpResp.success Then
            Debug.Print "Environments_GetUsers: Token failed for row " & i & " - " & httpResp.errorMessage
            GoTo NextTokenRow
        End If

        Set json = ParseJson(httpResp.body)
        m_s_AccessToken = json.item("access_token")
        m_s_TokenType = json.item("token_type")

        If m_s_AccessToken = "" Then
            Debug.Print "Environments_GetUsers: Empty token for row " & i
            GoTo NextTokenRow
        End If

        ' --- Write tenant ID (C) and bearer token (E) ---
        ws.Cells(i, 3).value = tenantId
        ws.Cells(i, 5).value = m_s_AccessToken
        ws.Cells(i, 5).WrapText = False

NextTokenRow:
    Next i

    ' =========================================================================
    ' LOOP 2: Use the stored token for each row to look up the M3 user ID (D)
    ' =========================================================================
    For i = 1 To lastRow
        jsonString = ws.Cells(i, 2).value
        If jsonString = "" Then Exit For

        ' --- Default the user (D) to blank; overwritten only if one is returned ---
        ws.Cells(i, 4).value = ""

        ' --- Skip rows that have no token from loop 1 ---
        m_s_AccessToken = ws.Cells(i, 5).value
        If m_s_AccessToken = "" Then GoTo NextUserRow

        Set json = ParseJson(jsonString)
        tenantId = json.item("ti")
        instanceUrl = json.item("iu")

        ' --- Call MRS001MI/GetUserInfo ---
        m_s_MainUrl = instanceUrl & "/" & tenantId
        m_s_MiPath = "/M3/m3api-rest/v2/execute"

        miBody = "{""program"":""MRS001MI"",""transactions"":[{""transaction"":""GetUserInfo"",""record"":{}}]}"

        httpResp = ExecuteApiPost(m_s_MainUrl, m_s_MiPath, miBody)

        If Not httpResp.success Then
            Debug.Print "Environments_GetUsers: MRS001MI call failed for row " & i & " - " & httpResp.errorMessage
            GoTo NextUserRow
        End If

        ' --- Parse results and write user ID (D) ---
        ' Response may be HTML (e.g. an error page) instead of JSON. If parsing
        ' fails, blank out the user and move on rather than aborting the run.
        Set json = Nothing
        On Error Resume Next
        Set json = ParseJson(httpResp.body)
        On Error GoTo ErrorHandler

        If json Is Nothing Then
            Debug.Print "Environments_GetUsers: Non-JSON response for row " & i
            ws.Cells(i, 4).value = ""
            GoTo NextUserRow
        End If

        Set results = json.item("results")
        If results Is Nothing Then
            ws.Cells(i, 4).value = ""
            GoTo NextUserRow
        End If

        For Each resultItem In results
            On Error Resume Next
            Set records = resultItem.item("records")
            On Error GoTo ErrorHandler

            If Not records Is Nothing Then
                For Each record In records
                    If record.item("ZZUSID") = "" Then
                        ws.Cells(i, 4).value = record.item("USFN")
                    Else
                        ws.Cells(i, 4).value = record.item("ZZUSID")
                    End If
                Next record
            End If
        Next resultItem

NextUserRow:
    Next i

    g_b_SuppressAuthPrompt = False
    KillPleaseWait
    Debug.Print "Environments_GetUsers: Complete"
    Exit Sub

ErrorHandler:
    g_b_SuppressAuthPrompt = False
    Debug.Print "Environments_GetUsers: ERROR at row " & i & " - " & Err.description
    MsgBox "Error in Environments_GetUsers: " & Err.description, vbCritical
    KillPleaseWait
End Sub

' =============================================================================
' MODERNIZED Environments_Load
' =============================================================================
' Uses new Doppio_Http module for faster HTTP calls while maintaining
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
    Dim tenantId As String
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
    ws.Columns("A:H").ClearContents
    
    ' Get and decode JSON configuration from Environments sheet
    jsonString = ws.Range("M2").value
    jsonString = Base64Decode(jsonString)
    
    ' Parse the JSON configuration
    Set json = ParseJson(jsonString)
    
    ' Extract values from the JSON object
    tenantId = json.item("ti")
    clientId = json.item("ci")
    clientSecret = json.item("cs")
    ssoBase = json.item("pu")
    tokenEndpoint = json.item("ot")
    authEndpoint = json.item("oa")
    saak = json.item("saak")
    sask = json.item("sask")
    
    ws.Range("L2").value = tenantId
    
    ' Set up API call parameters
    mainUrl = json.item("iu") & "/" & tenantId
    miPath = "/M3/m3api-rest/v2/execute"
    miUrl = ""
    m_s_Company = ""
    m_s_Division = ""
    m_s_M3user = ""
    
    Debug.Print "Environments_Load: Starting..."
    Debug.Print "  TenantID: " & tenantId
    Debug.Print "  MainUrl: " & mainUrl
    
    ' Get user and machine info (works on both Mac and Windows)
    GetUserAndMachineInfo userName, fullUserName, machineName
    encodedUserName = Replace(UrlEncode(userName), "%20", " ")
    encodedMachineName = Replace(UrlEncode(machineName), "%20", " ")
    
    Debug.Print "  UserName: " & userName
    Debug.Print "  MachineName: " & machineName
    
    ' === STEP 1: Get Access Token ===
    tokenUrl = ssoBase & tokenEndpoint
    tokenBody = "client_id=" & Core_UrlEncode(clientId) & _
                "&client_secret=" & Core_UrlEncode(clientSecret) & _
                "&grant_type=password" & _
                "&username=" & Core_UrlEncode(saak) & _
                "&password=" & Core_UrlEncode(sask)
    
    Debug.Print "  Getting token from: " & tokenUrl
    
    config.url = tokenUrl
    config.method = HttpMethod_POST
    config.contentType = "application/x-www-form-urlencoded"
    config.AcceptType = "application/json"
    config.authHeader = ""
    config.body = tokenBody
    config.timeoutSeconds = 30
    
    httpResponse = ExecuteRequest(config)
    
    If Not httpResponse.success Then
        Debug.Print "  Token request failed: " & httpResponse.errorMessage
        MsgBox "Unable to get token: " & httpResponse.errorMessage, vbCritical
        Exit Sub
    End If
    
    Set json = ParseJson(httpResponse.body)
    m_s_AccessToken = json.item("access_token")
    m_s_TokenType = json.item("token_type")
    
    If m_s_AccessToken = "" Then
        Debug.Print "Environments_Load: No access token in response"
        MsgBox "Unable to get access token", vbCritical
        Exit Sub
    End If

    Debug.Print "Environments_Load: Got token (" & Len(m_s_AccessToken) & " chars)"
    
    ' === STEP 2: Get tenant list (EXAUTH = 20) ===
    body = "{""program"":""EXPORTMI"",""transactions"":[{""transaction"":""Select"",""record"":{""SEPC"":""^"",""HDRS"":""0"",""QERY"":""EXPCID,EXTNNM,EXHASH from EXTXSM where EXAUTH = 20""},""selectedColumns"":[""REPL""]}]}"

    Debug.Print "Environments_Load: Getting tenant list..."
    httpResponse = ExecuteApiPost(mainUrl, miPath, body)
    
    If httpResponse.success Then
        ' Parse response and set m_obj_Results for Environment_Tenants
        Set json = ParseJson(httpResponse.body)
        If Not json Is Nothing Then
            Set m_obj_Results = json.item("results")
            Debug.Print "  Tenant list: Got " & m_obj_Results.count & " results"
        End If
        
        ' Call Environment_Tenants to process the results
        Environment_Tenants mainUrl
    Else
        Debug.Print "  Tenant list request failed: " & httpResponse.errorMessage
    End If
    
    ' === STEP 3: Get user's authorized environments (EXAUTH = 1) ===
    body = "{""program"":""EXPORTMI"",""transactions"":[{""transaction"":""Select"",""record"":{""SEPC"":""^"",""HDRS"":""0"",""QERY"":""EXTNNM,EXM3ID from EXTXSM where EXPCID = '" & encodedUserName & "' and EXAUTH = 1""},""selectedColumns"":[""REPL""]}]}"

    Debug.Print "Environments_Load: Getting user environments..."
    httpResponse = ExecuteApiPost(mainUrl, miPath, body)
    
    If httpResponse.success Then
        ' Parse response and set m_obj_Results for Environment_List
        Set json = ParseJson(httpResponse.body)
        If Not json Is Nothing Then
            Set m_obj_Results = json.item("results")
            Debug.Print "  User environments: Got " & m_obj_Results.count & " results"
        End If
        
        ' Also set fileResult for backward compatibility
        fileResult = httpResponse.body
    Else
        Debug.Print "  User environments request failed: " & httpResponse.errorMessage
    End If
    
    ' === STEP 4: Check if user has any environments ===
    envCount = Environment_List()
    
    If envCount = 0 Then
        Debug.Print "Environments_Load: No environments found, requesting access..."

        Set ws = ThisWorkbook.Sheets("Environments")
        ws.Cells(1, 1).value = "Access requested"
        encodedUserInfo = Base64Encode(GetUserSessionInfo)
        
        body = "{""program"":""EXT124MI"",""transactions"":[{""transaction"":""AddUsrInfo"",""record"":{""PCID"":""" & encodedUserName & """,""TNNM"":""" & encodedMachineName & """,""M3ID"":""unknown"",""AUTH"":""99"",""HASH"":""" & encodedUserInfo & """}}]}"
        
        httpResponse = ExecuteApiPost(mainUrl, miPath, body)
        
        If Not httpResponse.success Then
            Debug.Print "  Access request failed: " & httpResponse.errorMessage
        End If
    Else
        Debug.Print "Environments_Load: Found " & envCount & " environments"
    End If
    
    ' Cleanup
    ws.Columns("A:D").WrapText = False
    KillPleaseWait
    
    Debug.Print "Environments_Load: Complete"
    Exit Sub
    
ErrorHandler:
    Debug.Print "Environments_Load: ERROR - " & Err.description
    MsgBox "Error in Environments_Load: " & Err.description, vbCritical
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
        machineName = Split(machineName, " (")(0)
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
Public Function ExecuteApiPost(mainUrl As String, miPath As String, body As String) As httpResponse
    Dim config As httpConfig
    Dim apiUrl As String
    
    apiUrl = mainUrl & miPath
    
    config.url = apiUrl
    config.method = HttpMethod_POST
    config.contentType = "application/json; charset=UTF-8"
    config.AcceptType = "application/json; charset=UTF-8"
    config.authHeader = m_s_TokenType & " " & m_s_AccessToken
    config.body = body
    config.timeoutSeconds = 30
    
    ExecuteApiPost = ExecuteRequest(config)
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
                        envValue2 = Base64Decode(b64env)
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

    Set m_obj_JsonResponse = ParseJson(fileResult)

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
    Dim namedRange As name
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
    tstRow = Application.WorksheetFunction.match("TST*", environmentSheet.Columns("A"), 0)
    If tstRow = 0 Then
        tstRow = Application.WorksheetFunction.match("TEST*", environmentSheet.Columns("A"), 0)
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
            Debug.Print "Environment_SetDefaultValue: Error: namedRange refers to an invalid range."
            Debug.Print "Environment_SetDefaultValue: namedRange RefersTo: " & namedRange.RefersTo
            ' Remove invalid named range
            ws.names("Environment").Delete
            Set namedRange = Nothing
        End If
        On Error GoTo 0
    End If
    
    ' Create the named range with the default value if it doesn't exist
    If namedRange Is Nothing Then
        On Error Resume Next
        ws.names.Add name:="Environment", RefersTo:="=" & environmentSheet.name & "!$I$2"
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
            Debug.Print "ExecuteScriptWithRetry: Err.Number: " & Err.Number
            Debug.Print "ExecuteScriptWithRetry: script: " & script & vbCr & scriptResult
            scriptRetry = scriptRetry + 1
            currentMaxTime = currentMaxTime * 2 ' Double the max-time value
            
            If Not PromptUser("max-time " & initialMaxTime & " hit, retry with " & currentMaxTime & "?") Then
                KillPleaseWait
                Err.Clear
                Exit Sub
            End If
            
            script = Replace(script, "--max-time " & initialMaxTime, "--max-time " & currentMaxTime)
            initialMaxTime = currentMaxTime
            UI_ShowPleaseWait "Retry Attempt " & scriptRetry & "...  (--max-time " & currentMaxTime & ")"
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
    For i = 2 To m_obj_ws.Cells(8, Columns.count).End(xlToLeft).Column
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
    For i = 1 To ws.Columns.count
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

' =============================================================================
' MODERNIZED GetLayout_Click
' =============================================================================
' Uses new Doppio_Http module for API calls
' Follows same authentication pattern as Process_Click
' =============================================================================

Public Sub GetLayout_Click(mandatory As Boolean)
    Dim ws As Worksheet
    Dim url As String
    Dim apiType As String
    Dim currentEnv As String
    
    On Error GoTo ErrorHandler
    
    ' Update version
    UI_UpdateVersion
    
    ' Check environment
    If Range("Environment").value = "" Then Exit Sub
    
    Set ws = ActiveSheet
    currentEnv = ws.Range("I2").value
    
    ' Clear status
    ws.Range("G6").value = ""
    
'    UI_UI_ShowPleaseWait "Loading Transactions Layout"
    Application.GoTo Reference:="R9C2", Scroll:=True
    
    ' Set progress indicator
    ws.Range("J3").value = 0
    
    ' =========================================================================
    ' AUTHENTICATION - Same pattern as Process_Click
    ' =========================================================================
    If m_s_AccessToken = "" Or _
       ws.Range("I3").value = "" Or _
       ws.Range("I4").value = "" Or _
       activeEnvironment <> currentEnv Then
        
        Debug.Print "GetLayout_Click_New: Need new token"
        Tenant_Token
        m_l_Row = 9
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
            On Error Resume Next
            Application.Run "Doppio_IPS.ProcessLayoutIPS", ws, mandatory
            On Error GoTo ErrorHandler
            
        Case "XtendM3"
            ProcessLayoutM3X ws, mandatory
            
        Case "IDM"
'            ProcessLayoutIDM ws, mandatory
            
        Case Else
            ' Unknown type
    End Select
    
    ' Post-processing
    AutoFit_ColumnsAndRows False, mandatory
    FilterRow8BasedOnPopulatedColumns ws
    UI_KillPleaseWait
    
    Exit Sub
    
ErrorHandler:
    UI_KillPleaseWait
    Debug.Print "GetLayout_Click_New: ERROR - " & Err.description
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
    m_s_MiPath = "/M3/m3api-rest/v2/execute"
    
    ' Rename sheet
    Dim exportmiTable As String
    exportmiTable = Trim(ws.Range("B3").value)
    If ws.Range("API").value = "EXPORTMI" And exportmiTable Like "*[A-Za-z0-9]*" Then
        UI_RenameSheet "EXPORTMI for " & exportmiTable
    Else
        UI_RenameSheet ""
    End If
    
    ' Get API and Transaction from worksheet
    apiName = ws.Range("API").value
    transName = ws.Range("Transaction").value
    
    ' Build URL for MRS001MI/LstFields
    url = "MRS001MI/LstFields;maxrecs=0?MINM=" & apiName & "&TRNM=" & transName
    
    ' Get current column count (safely)
    On Error Resume Next
    columnCount = m_obj_ColumnNames.count()
    If Err.Number <> 0 Then columnCount = 0
    On Error GoTo ErrorHandler
    columnCount = 0
    
    Debug.Print "ProcessLayoutAPI: URL = " & url
    Debug.Print "ProcessLayoutAPI: m_s_LoadedUrl = " & m_s_LoadedUrl
    Debug.Print "ProcessLayoutAPI: columnCount = " & columnCount
    
    ' Special handling for EXPORTMI/Select - reset loaded URL if columns not loaded
    If apiName = "EXPORTMI" And Left(transName, 6) = "Select" Then
        If columnCount < 0 Then
            m_s_LoadedUrl = ""
        End If
    End If
    
    ' Only fetch if API and Transaction are set
    If apiName <> "" And transName <> "" Then
        ' Fetch if:
        ' 1. URL changed from what we last loaded, OR
        ' 2. No columns are loaded (even if URL matches - columns may have been cleared)
        If m_s_LoadedUrl <> url Or columnCount = 0 Then
            Debug.Print "ProcessLayoutAPI: Calling GetLayoutAPI (URL changed or no columns)"
            GetLayoutAPI url, ws
        Else
            Debug.Print "ProcessLayoutAPI: Skipping load - URL matches and columns exist"
        End If
        AutoFit_ColumnsAndRows True, mandatory
    End If
    
    Exit Sub
    
ErrorHandler:
    Debug.Print "ProcessLayoutAPI: ERROR - " & Err.description
End Sub

Private Sub GetLayoutAPI(url As String, ws As Worksheet)
    Dim maxbulk_hold As Integer
    
    On Error GoTo ErrorHandler
    
    ' Store URL
    m_s_LoadedUrl = url
    
    ' Initialize collections
    m_obj_ColumnNames.Initialize
    m_obj_ColumnDescriptions.Initialize
    m_obj_ColumnTypes.Initialize
    m_obj_ColumnConditions.Initialize
    m_obj_ColumnDirections.Initialize
    
    ' Hold maxbulk
    maxbulk_hold = maxbulk
    maxbulk = 1
    
    ' Fetch input columns first
    FetchAndProcessColumns url & "&TRTP=I&returncols=FLNM,FLDS,TYPE,LENG,MAND", "I"
    
    ' Fetch output columns
    FetchAndProcessColumns url & "&TRTP=O&returncols=FLNM,FLDS,TYPE,LENG,MAND", "O"
    
    ' Restore maxbulk
    maxbulk = maxbulk_hold
    
    Exit Sub
    
ErrorHandler:
    Debug.Print "GetLayoutAPI: ERROR - " & Err.description
End Sub

Private Sub FetchAndProcessColumns(url As String, direction As String)
    Dim response As apiResponse
    Dim record As Object
    Dim flds As String, typ As String, leng As String, description As String
    Dim value As String
    Dim mandValue As Variant
    
    On Error GoTo ErrorHandler
    
    ' 1. Try to get from new Cache
    ' Cache_TryGetFromCache populates 'response' if found and valid
    If Not Cache_TryGetFromCache(url, response) Then
        
        Debug.Print "FetchAndProcessColumns: Cache miss, calling API"
        
        ' 2. Execute API Call
        response = ExecuteLayoutCall(url)
        
        ' 3. Store in new Cache if successful
        If response.success And Not response.records Is Nothing Then
            Cache_StoreInCache url, response
            Debug.Print "FetchAndProcessColumns: Stored in cache"
        End If
    Else
        Debug.Print "FetchAndProcessColumns: Cache hit"
    End If
    
    ' 4. Process records (from response object, not global variable)
    If response.success And Not response.records Is Nothing Then
        Debug.Print "FetchAndProcessColumns: Processing " & response.records.count & " records"
        
        For Each record In response.records
            value = record.item("FLNM")
            
            ' For input, add all; for output, only add if not already present
            ' We still write to Doppio globals because downstream AutoFit functions likely depend on them [cite: 312]
            If direction = "I" Or (direction = "O" And Not m_obj_ColumnNames.Contains(value)) Then
                m_obj_ColumnNames.Add record.item("FLNM")
                
                flds = record.item("FLDS") & vbCrLf
                typ = record.item("TYPE")
                leng = record.item("LENG")
                description = flds & typ & leng
                
                m_obj_ColumnDescriptions.Add description
                m_obj_ColumnTypes.Add record.item("TYPE")
                
                mandValue = record.item("MAND")
                If IsNull(mandValue) Or mandValue = "" Then
                    mandValue = 0
                End If
                m_obj_ColumnConditions.Add mandValue
                m_obj_ColumnDirections.Add direction
            End If
        Next record
        
        Debug.Print "FetchAndProcessColumns: Total columns = " & m_obj_ColumnNames.count()
    Else
        Debug.Print "FetchAndProcessColumns: No records to process or API failed"
    End If
    
    Exit Sub
    
ErrorHandler:
    Debug.Print "FetchAndProcessColumns: ERROR - " & Err.description
End Sub

' =============================================================================
' XTENDM3 LAYOUT
' =============================================================================

Private Sub ProcessLayoutM3X(ws As Worksheet, mandatory As Boolean)
    Dim url As String
    
    On Error GoTo ErrorHandler
    
    url = ws.Range("Transaction").value
    
    If m_s_LoadedUrl <> url Then
        GetLayoutM3X "/M3/extensibility/ionapi-doc", url, ws
        m_s_LoadedUrl = url
    End If
    
    AutoFit_ColumnsAndRows True, mandatory
    
    Exit Sub
    
ErrorHandler:
    Debug.Print "ProcessLayoutM3X: ERROR - " & Err.description
End Sub

Private Sub GetLayoutM3X(path As String, url As String, ws As Worksheet)
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
    m_obj_ColumnNames.Initialize
    m_obj_ColumnDescriptions.Initialize
    m_obj_ColumnTypes.Initialize
    m_obj_ColumnConditions.Initialize
    m_obj_ColumnDirections.Initialize
    
    methodName = ws.Range("Transaction").value
    
    ' Process response
    If response.success And Len(response.data) > 0 Then
        Set jsonObject = ParseJson(response.data)
        
        For Each pathKey In jsonObject.item("paths").keys
            For Each methodKey In jsonObject.item("paths").item(pathKey).keys
                If jsonObject.item("paths").item(pathKey).item(methodKey).exists("parameters") Then
                    Set parameters = jsonObject.item("paths").item(pathKey).item(methodKey).item("parameters")
                    
                    For Each parameter In parameters
                        aliasName = parameter.item("name")
                        dataType = parameter.item("type")
                        Required = parameter.item("required")
                        
                        m_obj_ColumnNames.Add aliasName
                        description = aliasName & vbCrLf & dataType
                        m_obj_ColumnDescriptions.Add description
                        m_obj_ColumnTypes.Add dataType
                        m_obj_ColumnConditions.Add Required
                        m_obj_ColumnDirections.Add "I"
                    Next parameter
                End If
            Next methodKey
        Next pathKey
    End If
    
    Exit Sub
    
ErrorHandler:
    Debug.Print "GetLayoutM3X: ERROR - " & Err.description
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
    apiUrl = m_s_MainUrl & "/M3/m3api-rest/v2/execute/" & url
    
    ' Configure request
    config.url = apiUrl
    config.method = HttpMethod_GET
    config.contentType = "application/json"
    config.AcceptType = "application/json"
    config.authHeader = m_s_TokenType & " " & m_s_AccessToken
    config.timeoutSeconds = 30
    config.body = ""
    
    Debug.Print "ExecuteLayoutCall: " & apiUrl
    
    ' Execute
    httpResponse = ExecuteRequest(config)
    
    ' Build response
    response.success = httpResponse.success
    response.data = httpResponse.body
    response.errorMessage = httpResponse.errorMessage
    
    ' Parse JSON
    If httpResponse.success And Len(httpResponse.body) > 0 Then
        Set json = ParseJson(httpResponse.body)
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
    apiUrl = m_s_MainUrl & path & url
    
    ' Configure request
    config.url = apiUrl
    config.method = HttpMethod_GET
    config.contentType = "application/json"
    config.AcceptType = "application/json"
    config.authHeader = m_s_TokenType & " " & m_s_AccessToken
    config.timeoutSeconds = 30
    config.body = ""
    
    Debug.Print "ExecuteSwaggerLayoutCall: " & apiUrl
    
    ' Execute
    httpResponse = ExecuteRequest(config)
    
    ' Build response
    response.success = httpResponse.success
    response.data = httpResponse.body
    response.errorMessage = httpResponse.errorMessage
    
    ' Parse JSON
    If httpResponse.success And Len(httpResponse.body) > 0 Then
        Set json = ParseJson(httpResponse.body)
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
    
    found = Cache_TryGetFromCache(cacheKey, response)
    
    If Not found Then
        response.success = False
    End If
    
    TryGetFromCache_Layout = response
End Function

' =============================================================================
' UI HELPERS
' =============================================================================



' =============================================================================
' LEGACY WRAPPERS
' =============================================================================

Public Sub GetLayoutAll_Click()
    GetLayout_Click False
End Sub

Public Sub GetLayoutMan_Click()
    GetLayout_Click True
End Sub
' =============================================================================

Sub GetSingleMultiple()
    GetTransactions_Click
    m_obj_ws.Range("G5").NumberFormat = "General"
    m_obj_ws.Range("G5").Formula = "=IFNA(VLOOKUP(Transaction,Transactions!B:C,2,False),"""")"
    m_obj_ws.Calculate
    DoEvents
End Sub

' =============================================================================
' MODERNIZED GetTransactions_Click
' =============================================================================
' Uses new Doppio_Http module for API calls
' Follows same authentication pattern as Process_Click
' =============================================================================

Public Sub GetTransactions_Click()
    Dim ws As Worksheet
    Dim api As String
    Dim apiType As String
    Dim currentEnv As String
    
    On Error GoTo ErrorHandler
    
    ' Update version display
    UI_UpdateVersion
    
    ' Check environment
    If Range("Environment").value = "" Then Exit Sub
    
    Set ws = ActiveSheet
    currentEnv = ws.Range("I2").value
    
    ' Clear G6 if exists
    On Error Resume Next
    ws.Range("G6").value = ""
    On Error GoTo ErrorHandler
    
 '   UI_UI_ShowPleaseWait "Loading Transactions For API"
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
    If m_s_AccessToken = "" Or _
       ws.Range("I3").value = "" Or _
       ws.Range("I4").value = "" Or _
       activeEnvironment <> currentEnv Then
        
        Debug.Print "GetTransactions_Click: Need new token"
        Tenant_Token
        m_l_Row = 9
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
            On Error Resume Next
            Application.Run "Doppio_IPS.GetTransactionsIPS", api, ws
            On Error GoTo ErrorHandler
        Case "XtendM3"
            GetTransactionsM3X_New ws
        Case "IDM"
'            IDM_Load_Methods
        Case Else
            ' Unknown type
    End Select
    
    ' Post-processing
    SortTransactions
    CheckAndUpdateValue
    UI_KillPleaseWait
    UpdateG5Formula_New ws
    
    Exit Sub
    
ErrorHandler:
    UI_KillPleaseWait
    Debug.Print "GetTransactions_Click: ERROR - " & Err.description
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
    RecordCache_Initialize
    
    ' Build cache key
    cacheKey = "MRS001MI/LstTransactions?MINM=" & api
    
    ' Check cache first
    response = TryGetFromCache(cacheKey)
    
    ' Re-fetch if: no cache hit, records is Nothing, or cache returned 0 records
    If Not response.success Or response.records Is Nothing Or response.recordCount = 0 Then
        Debug.Print "GetTransactionsAPI_New: Cache records=" & response.recordCount & " - calling API"
        response = ExecuteTransactionCall("MRS001MI", "LstTransactions", "MINM=" & api & "&returncols=MINM,TRNM,SIMU")

        ' Store in cache if successful
        If response.success Then
            Cache_StoreDataInCache cacheKey, response.data
        End If
    Else
        Debug.Print "GetTransactionsAPI_New: Cache hit - " & response.recordCount & " records"
    End If

    ' Output results to Transactions sheet
    If response.success Then
        Output_LstTransactions response
    End If
    
    Exit Sub
    
ErrorHandler:
    Debug.Print "GetTransactionsAPI_New: ERROR - " & Err.description
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
    Debug.Print "GetTransactionsM3X_New: ERROR - " & Err.description
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
    apiUrl = m_s_MainUrl & "/M3/m3api-rest/v2/execute/" & program & "/" & transaction & ";maxrecs=0?" & parameters
    
    ' Configure request
    config.url = apiUrl
    config.method = HttpMethod_GET
    config.contentType = "application/json"
    config.AcceptType = "application/json"
    config.authHeader = m_s_TokenType & " " & m_s_AccessToken
    config.timeoutSeconds = 30
    config.body = ""
    
    Debug.Print "ExecuteTransactionCall: " & apiUrl
    
    ' Execute
    httpResponse = ExecuteRequest(config)
    
    ' Build response
    response.success = httpResponse.success
    response.data = httpResponse.body
    response.errorMessage = httpResponse.errorMessage
    
    ' Parse JSON
    If httpResponse.success And Len(httpResponse.body) > 0 Then
        Set json = ParseJson(httpResponse.body)
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
    apiUrl = m_s_MainUrl & "/" & path & queryString
    
    ' Configure request
    config.url = apiUrl
    config.method = HttpMethod_GET
    config.contentType = "application/json"
    config.AcceptType = "application/json"
    config.authHeader = m_s_TokenType & " " & m_s_AccessToken
    config.timeoutSeconds = 30
    config.body = ""
    
    Debug.Print "ExecuteSwaggerCall: " & apiUrl
    
    ' Execute
    httpResponse = ExecuteRequest(config)
    
    ' Build response
    response.success = httpResponse.success
    response.data = httpResponse.body
    response.errorMessage = httpResponse.errorMessage
    
    ' Parse JSON for swagger
    If httpResponse.success And Len(httpResponse.body) > 0 Then
        Set json = ParseJson(httpResponse.body)
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
    
    found = Cache_TryGetFromCache(cacheKey, response)
    
    If Not found Then
        response.success = False
    End If
    
    TryGetFromCache = response
End Function

' =============================================================================
' OUTPUT HELPERS
' =============================================================================

Private Sub Output_LstTransactions(response As apiResponse)
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
    Debug.Print "Output_LstTransactions: ERROR - " & Err.description
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
    Debug.Print "OutputSwaggerTransactions_New: ERROR - " & Err.description
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
    Debug.Print "OutputM3XTransactions_New: ERROR - " & Err.description
End Sub

Private Sub UpdateG5Formula_New(ws As Worksheet)
    On Error Resume Next
    
    If ws.Range("G5").value = "" Then
        ws.Range("G5").NumberFormat = "General"
        ws.Range("G5").Formula = "=IFNA(VLOOKUP(Transaction,Transactions!B:C,2,False),"""")"
        ws.Calculate
        DoEvents
    End If
    
    On Error GoTo 0
End Sub

Sub HelpSheet()
    Sheets("Help").Visible = True
    Sheets("Help").Activate
    With ActiveWindow
        If .FreezePanes Then .FreezePanes = False
        .SplitColumn = 40
        .SplitRow = 40
        .FreezePanes = True
    End With
    Application.GoTo Reference:="R1C1", Scroll:=True
End Sub

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
' Uses Doppio_Http for API calls
' =============================================================================

Public Sub LoadSourceAPIFormats()
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

    UI_ShowPleaseWait "Loading Source API Formats"
        
    ' Force connect to source environment
    Tenant_Token
    
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
        Set m_obj_JsonResponse = response.results
        'IDM_Load_Tags
    End If
    
    ' ==========================================
    ' Load XtendM3 (static entry)
    ' ==========================================
    resultSheet.Cells(9, 4).value = "Extensibility"
    resultSheet.Cells(9, 8).value = "Infor XtendM3"
    
    ' Sort results
    Set sortRange = resultSheet.Range("A9:B" & resultSheet.Cells(resultSheet.Rows.count, 1).End(xlUp).row)
    sortRange.Sort Key1:=sortRange.Columns(1), Order1:=xlAscending, header:=xlYes
    sortRange.Columns.AutoFit
    
    ' Check if "API" field exists on the panel
    On Error Resume Next
    Set rng = ws.Range("API")
    On Error GoTo ErrorHandler
    
    If Not rng Is Nothing Then
        ws.Cells(2, 2).value = "API"
        If ws.Range("API").value = "" Then
            ws.Range("API").value = "CRS111MI"
        End If
        GetTransactions_Click
        Settings_CopyDefaults
    End If
    
    UI_KillPleaseWait
    Exit Sub
    
ErrorHandler:
    UI_KillPleaseWait
    Debug.Print "LoadSourceAPIFormats_New: ERROR - " & Err.description
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
    apiUrl = m_s_MainUrl & "/M3/m3api-rest/v2/execute/" & endpoint
    
    ' Configure request
    config.url = apiUrl
    config.method = HttpMethod_GET
    config.contentType = "application/json"
    config.AcceptType = "application/json"
    config.authHeader = m_s_TokenType & " " & m_s_AccessToken
    config.timeoutSeconds = 30
    config.body = ""
    
    Debug.Print "ExecuteMICall_Load: " & apiUrl
    
    ' Execute
    httpResponse = ExecuteRequest(config)
    
    ' Build response
    response.success = httpResponse.success
    response.data = httpResponse.body
    response.errorMessage = httpResponse.errorMessage
    
    ' Parse JSON - MI response has "results" array
    If httpResponse.success And Len(httpResponse.body) > 0 Then
        Set json = ParseJson(httpResponse.body)
        
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
    Debug.Print "ExecuteMICall_Load: ERROR - " & Err.description
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
    apiUrl = m_s_MainUrl & endpoint
    
    ' Configure request
    config.url = apiUrl
    config.method = HttpMethod_GET
    config.contentType = "application/json"
    config.AcceptType = "application/json"
    config.authHeader = m_s_TokenType & " " & m_s_AccessToken
    config.timeoutSeconds = 30
    config.body = ""
    
    Debug.Print "ExecuteSwaggerCall_Load: " & apiUrl
    
    ' Execute
    httpResponse = ExecuteRequest(config)
    
    ' Build response
    response.success = httpResponse.success
    response.data = httpResponse.body
    response.errorMessage = httpResponse.errorMessage
    
    ' Parse JSON - Swagger response is the raw JSON object
    If httpResponse.success And Len(httpResponse.body) > 0 Then
        Set json = ParseJson(httpResponse.body)
        Set response.results = json
    End If
    
    ExecuteSwaggerCall_Load = response
    Exit Function
    
ErrorHandler:
    Debug.Print "ExecuteSwaggerCall_Load: ERROR - " & Err.description
    response.success = False
    response.errorMessage = Err.description
    ExecuteSwaggerCall_Load = response
End Function

' =============================================================================
' MODERNIZED Log_Activity
' =============================================================================
' Uses Doppio_Http for API calls
' Cross-platform compatible (Mac/Windows)
' =============================================================================

Public Sub Log_Activity()
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
    saveType = m_s_TokenType
    saveToken = m_s_AccessToken
    
    ' Decode and parse config
    jsonString = "eyJ0aSI6IkRPUFBJT19ERU0iLCJjbiI6Ik0zIERhdGEgTG9hZGVyIiwiZHQiOiIxMiIsImNpIjoiRE9QUElPX0RFTX5abS1tNkRMZTlCWHBuWDM0SVVfQ3RDaGdYX2R2b3lCUXFTQ2hOR2JSUm1NIiwiY3MiOiJacHN0MXhpbFNHTEJrdGFyemNWZFdSbkEyb2tNbHNCTUdjOW1wUTdoZE1RZW12ZDdYR1RucFFicWVXTVpNMHBUb3BXcE9FQmIwUFlaMnVMbUxRUjY3ZyIsIml1IjoiaHR0cHM6Ly9taW5nbGUtaW9uYXBpLmluZm9yY2xvdWRzdWl0ZS5jb20iLCJwdSI6Imh0dHBzOi8vbWluZ2xlLXNzby5pbmZvcmNsb3Vkc3VpdGUuY29tOjQ0My9ET1BQSU9fREVNL2FzLyIsIm9hIjoiYXV0aG9yaXphdGlvbi5vYXV0aDIiLCJvdCI6InRva2VuLm9hdXRoMiIsIm9yIjoicmV2b2tlX3Rva2VuLm9hdXRoMiIsImV2IjoiVTE0NzgzNTgxMDEiLCJ2IjoiMS4wIiwic2FhayI6IkRPUFBJT19ERU0jek9lYkI0eGwxZGlFeTIxaWtNUXNJSzFfaW1sZjdycFBuRGEyYTZxOEtkdFE1Vy1hZDhHY2o2NEd3OVVILUplSU9mbGdWenJuZlRKLVdJSDFZYlU4MlEiLCJzYXNrIjoiMzhKcElrcTVIWUhEbGhwcUVuT040cDIyQ2o2NEVkNVdwbE5zaVNjMXRhRUd4NFV3aGZRblRGMnZYUDhha1R5Ri12S3llSWFocjlwX1RzM3NubmJIU2cifQo"
    jsonString = Base64Decode(jsonString)
    Set json = ParseJson(jsonString)
    
    ' Extract from JSON config
    ssoBase = json.item("pu")
    tokenEndpoint = json.item("ot")
    clientId = json.item("ci")
    clientSecret = json.item("cs")
    
    mainUrl = json.item("iu") & "/" & json.item("ti")
    miPath = "/M3/m3api-rest/v2/execute"
    
    urlSelectedEnvironment = UrlEncode_Log(m_s_SelectedEnvironment)
    encTenant = Base64Encode(encodedTenant)
    
    ' Get user/machine info (cross-platform)
    GetUserInfo userName, fullUserName, machineName
    
    encodedUserName = UrlEncode_Log(userName)
    encodedMachineName = UrlEncode_Log(machineName)
    encodedSelectedEnvironment = UrlEncode_Log(m_s_SelectedEnvironment)
    
    ' Get token for logging service
    If Not GetLogActivityToken(ssoBase & tokenEndpoint, clientId, clientSecret, json.item("saak"), json.item("sask")) Then
        GoTo Cleanup
    End If
    
    ' Update user information
    miUrl = mainUrl & miPath & "/EXT123MI/UpdUsrInfo?" & _
            "M3NM=" & encodedSelectedEnvironment & _
            "&M3ID=" & m_s_M3user & _
            "&PCNM=" & encodedMachineName & _
            "&PCID=" & encodedUserName & _
            "&VERS=" & DOPPIO_VERSION & _
            "&HASH=" & encTenant & _
            "&TNAL=" & urlSelectedEnvironment
    
    ExecuteLogCall miUrl
    
    ' Add user information
    miUrl = mainUrl & miPath & "/EXT123MI/AddUsrInfo?" & _
            "M3NM=" & encodedSelectedEnvironment & _
            "&M3ID=" & m_s_M3user & _
            "&PCNM=" & encodedMachineName & _
            "&PCID=" & encodedUserName & _
            "&AUTH=1" & _
            "&VERS=" & DOPPIO_VERSION & _
            "&TNAL=" & urlSelectedEnvironment & _
            "&HASH=" & encTenant
    
    ExecuteLogCall miUrl
    
Cleanup:
    ' Restore original token
    m_s_TokenType = saveType
    m_s_AccessToken = saveToken
    Exit Sub
    
ErrorHandler:
    Debug.Print "Log_Activity_New: ERROR - " & Err.description
    ' Restore original token even on error
    m_s_TokenType = saveType
    m_s_AccessToken = saveToken
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
        machineName = Split(machineName, " (")(0)
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
    config.timeoutSeconds = 5
    config.body = body
    
    Debug.Print "GetLogActivityToken: URL = " & tokenUrl
    
    ' Execute request
    response = ExecuteRequest(config)
    
    If response.success And Len(response.body) > 0 Then
        Set json = ParseJson(response.body)
        If Not json Is Nothing Then
            m_s_AccessToken = json.item("access_token")
            m_s_TokenType = json.item("token_type")
            GetLogActivityToken = True
            Debug.Print "GetLogActivityToken: Success"
            Exit Function
        End If
    End If
    
    Debug.Print "GetLogActivityToken: Failed - " & response.errorMessage
    GetLogActivityToken = False
    Exit Function
    
ErrorHandler:
    Debug.Print "GetLogActivityToken: ERROR - " & Err.description
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
    config.authHeader = m_s_TokenType & " " & m_s_AccessToken
    config.timeoutSeconds = 5
    config.body = ""
    
    Debug.Print "ExecuteLogCall: " & Left(url, 100) & "..."
    
    ' Execute request (fire and forget - don't care about response)
    response = ExecuteRequest(config)
    
    Debug.Print "ExecuteLogCall: Status = " & response.statusCode
    
    Exit Sub
    
ErrorHandler:
    Debug.Print "ExecuteLogCall: ERROR - " & Err.description
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
    lastCol = m_obj_ws.Cells(8, m_obj_ws.Columns.count).End(xlToLeft).Column
    
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
' Handles various settings commands
' Uses * for global variables
' =============================================================================

Public Sub Keywords(value As String)
    Dim startIndex As Integer
    Dim endIndex As Integer
    Dim numericValue As Long
    Dim ws As Worksheet
    Dim wsOld As Worksheet
    Dim wsNew As Worksheet
    Dim cacheSheet As Worksheet
    Dim tempFilePath As String
    Dim activeTransaction As String
    Dim savedEnv As String
    
    Set ws = ThisWorkbook.ActiveSheet
    Application.Calculation = xlCalculationAutomatic
    DoEvents
    UI_KillPleaseWait
    
    If Keywords_CustomModules(value) Then Exit Sub
     
    If value = "New sheet" Or value = "ns" Then
        Settings_NewSheet
        Exit Sub
    End If
    
    If value = "xlsx" Or value = "xls" Or value = "report" Then
        Create_xlsx
        Exit Sub
    End If
    
    If value = "samplerest" Or value = "rest" Then
        SampleREST
        Exit Sub
    End If
    
    If value = "curl" Then
        Dim curlScript As String
        Dim curlMiUrl As String
        Dim curlWs As Worksheet
        Dim curlApiName As String
        Dim curlTransaction As String
        Dim curlApiType As String
        Dim curlBuiltBody As String
        Dim curlLastCol As Integer
        Dim curlFieldName As String
        Dim curlFieldValue As String
        Dim curlFirstField As Boolean
        Dim curlRecord As String
        Dim ci2 As Integer

        Set curlWs = ActiveSheet
        curlApiName = curlWs.Range("API").value
        curlTransaction = curlWs.Range("Transaction").value
        curlApiType = curlWs.Range("Type").value
        curlMiUrl = curlApiName & "/" & curlTransaction

        ' Build body from row 8 (fields) and row 9 (values)
        curlLastCol = curlWs.Cells(8, curlWs.Columns.count).End(xlToLeft).Column
        curlRecord = ""
        curlFirstField = True
        For ci2 = 2 To curlLastCol
            curlFieldName = CStr(curlWs.Cells(8, ci2).value)
            If curlFieldName <> "" Then
                curlFieldValue = CStr(curlWs.Cells(9, ci2).value)
                curlFieldValue = Replace(curlFieldValue, "\", "\\")
                curlFieldValue = Replace(curlFieldValue, """", "\""")
                If Not curlFirstField Then curlRecord = curlRecord & ","
                curlRecord = curlRecord & """" & curlFieldName & """:""" & curlFieldValue & """"
                curlFirstField = False
            End If
        Next ci2

        curlBuiltBody = "{""program"":""" & curlApiName & """,""transactions"":[" & _
                        "{""transaction"":""" & curlTransaction & """,""record"":{" & curlRecord & "}}]}"

        Curl_Build m_s_MainUrl, m_s_MiPath, curlMiUrl, curlBuiltBody, curlApiType, curlScript
        SampleRESTPopup curlCommand, "CURL Sample"
        Exit Sub
    End If
    
    If value = "postman" Then
        Postman_Build
        Exit Sub
    End If
    
    If value = "pivot" Then
        RunBuildMatrix
        Exit Sub
    End If
    
    If value = "unpivot" Then
        RunUnpivotMatrix
        Exit Sub
    End If
    
    If value = "decode" Then
        ToggleBase64Button "Decode"
        Exit Sub
    End If
    
    If value = "encode" Then
        ToggleBase64Button "Encode"
        Exit Sub
    End If
    
    If value = "evs100" Then
        ExportToEVS100
        Exit Sub
    End If
    
    If value = "prep" Or value = "reset" Then
        Application.ScreenUpdating = False
        Application.Calculation = xlCalculationManual

        UI_RemoveButtons
        UI_DefaultButtons
        UI_DeleteSheets
        CleanSheet_Click
        ResetCountFormat ws
        
        If value = "prep" Then
            With ActiveSheet
                .Range("API").value = "CRS111MI"
                .Range("Type").value = "API"
                .Range("Transaction").value = "List"
                .Range("Environment").value = ""
                .Range("User").value = ""
                .Range("Company").value = ""
                .Range("Division").value = ""
                .Range("G6").value = ""
            End With
        End If
        
        ' Hide system sheets (keep Help visible)
        UI_HideSystemSheets showHelp:=True
        
        ' Clear Environments sheet
        If value = "prep" Then
            On Error Resume Next
            Set ws = Worksheets("Environments")
            On Error GoTo 0
            If Not ws Is Nothing Then
                Dim savedK2M2 As Variant
                savedK2M2 = ws.Range("K2:M2").value
                ws.Cells.ClearContents
                ws.Range("K2:M2").value = savedK2M2
            End If
        End If
        
        Set ws = ActiveSheet
        If value = "prep" Then
            FilterRow8BasedOnPopulatedColumns ws
            ws.Rows("7:8").ClearContents
            ws.Range(ws.Cells(3, 11), ws.Cells(ws.Rows.count, ws.Columns.count)).ClearContents
        End If
        
        Application.Calculation = xlCalculationAutomatic
        Application.ScreenUpdating = True
        DoEvents
        UI_KillPleaseWait
        
        ' Clean up temp files
        On Error Resume Next
        tempFilePath = Environ("HOME") & "/curl_output.txt"
        Kill tempFilePath
        tempFilePath = Environ("HOME") & "/curl_input.sh"
        Kill tempFilePath
        On Error GoTo 0

        If value = "prep" Then
            ActiveSheet.name = "Sheet1"
            SetFormulasAndFormatting_New ws
            Settings_CopyDefaults
            AutoFit_Click
            ws.Activate
        Else
            ' reset: copy a fresh sheet from Master so Worksheet_Change events
            ' are reliably registered (reusing the existing sheet after heavy
            ' cleanup leaves its event plumbing in a broken state)
            savedEnv = ws.Range("Environment").value
            Set wsOld = ws

            ' Free up the "Sheet1" tab name so the new copy can use it
            If wsOld.name = "Sheet1" Then wsOld.name = "_TempReset_"

            b_AddingSheet = True
            Sheets("Master").Copy After:=Sheets(Sheets.count)
            b_AddingSheet = False

            Set wsNew = Sheets(Sheets.count)
            wsNew.name = "Sheet1"
            wsNew.Visible = xlSheetVisible
            wsNew.Activate
            UI_DefaultButtons

            wsNew.Range("Environment").value = savedEnv
            SetFormulasAndFormatting_New wsNew
            Settings_CopyDefaults

            Application.DisplayAlerts = False
            wsOld.Delete
            Application.DisplayAlerts = True

            GetLayoutAll_Click
            wsNew.Activate
        End If

        Exit Sub
    End If
       
    If value = "environments" Or value = "env" Then
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
    
    If value = "ver" Then
        If Sheets("Versions").Visible Then
            Sheets("Versions").Visible = False
        Else
            Sheets("Versions").Visible = True
            Sheets("Versions").Activate
        End If
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
        UI_HideSystemSheets
        Exit Sub
    End If
    
    If Left(value, Len("maxrecs")) = "maxrecs" Then
        startIndex = InStr(value, "=")
        If startIndex > 0 Then
            endIndex = Len(value)
            numericValue = CLng(Mid(value, startIndex + 1, endIndex - startIndex))
            maxRecs = numericValue
            Config_MaxRecords = numericValue
            Debug.Print "Keywords: Settings: maxrecs = " & numericValue
        End If
        Exit Sub
    End If

    If Left(value, Len("maxbulk")) = "maxbulk" Then
        startIndex = InStr(value, "=")
        If startIndex > 0 Then
            endIndex = Len(value)
            numericValue = CLng(Mid(value, startIndex + 1, endIndex - startIndex))
            maxbulk = CInt(numericValue)
            Debug.Print "Keywords: Settings: maxbulk = " & maxbulk
        End If
        Exit Sub
    End If

    If Left(value, Len("refresh")) = "refresh" Then
        startIndex = InStr(value, "=")
        If startIndex > 0 Then
            endIndex = Len(value)
            numericValue = CLng(Mid(value, startIndex + 1, endIndex - startIndex))
            refreshSeconds = CInt(numericValue)
            Debug.Print "Keywords: Settings: refreshSeconds = " & refreshSeconds
        End If
        Exit Sub
    End If
    
    If value = "defaults" Then
        Settings_CopyDefaults
        Exit Sub
    End If
    
    If value = "help" Then
        HelpSheet
        Exit Sub
    End If
    
    If value = "settings" Then
        SettingsSheet
        Exit Sub
    End If
    
    If value = "clear" Or value = "clr" Then
        ActiveSheet.Rows("9:" & ActiveSheet.Rows.count).ClearContents
        ResetCountFormat ActiveSheet
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
    
    If Left(value, Len("load e")) = "load e" Then
        Environments_Load
        Exit Sub
    End If
End Sub

Sub SettingsSheet()
    Dim ws As Worksheet
    
    On Error GoTo ErrorHandler
    
    ' Remember the current sheet before switching
    Set m_OriginalSheet = ActiveSheet
    
    Set ws = ThisWorkbook.Sheets("Settings")
    
    Dim settings As ApiSettings
    settings = Config_ApiSettings
    
    Application.ScreenUpdating = False
    
    With ws
        .Visible = True
        .Range("maxrecs").value = maxRecs
        .Range("maxbulk").value = maxbulk
        .Range("refreshSeconds").value = refreshSeconds
        .Range("formatting").value = formatting
        .Range("righttrim").value = righttrim
        .Range("splitChar").value = splitChar
        .Range("maxtime").value = maxtime
        .Range("conoDivi").value = conoDivi
        .Range("naming").value = sheetNaming
        .Activate
    End With
    
    With ActiveWindow
        If .FreezePanes Then .FreezePanes = False
        .SplitColumn = 40
        .SplitRow = 40
        .FreezePanes = True
    End With
    
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

    Set ws = ActiveSheet
    Set settings = ThisWorkbook.Sheets("Settings")

    wasHidden = (settings.Visible <> xlSheetVisible)

    settings.Range("D7:D15").value = settings.Range("E7:E15").value

    ' Copy developer default explicitly in case its row falls outside D7:D15
    settings.Range("developer").value = settings.Range("developer").Offset(0, 1).value

    maxRecs = settings.Range("maxrecs").value
    maxbulk = settings.Range("maxbulk").value
    refreshSeconds = settings.Range("refreshSeconds").value
    righttrim = settings.Range("righttrim").value
    formatting = settings.Range("formatting").value
    splitChar = settings.Range("splitChar").value
    maxtime = settings.Range("maxtime").value
    conoDivi = settings.Range("conoDivi").value
    sheetNaming = settings.Range("naming").value

    Config_LoadSettingsFromSheet

    If wasHidden Then settings.Visible = xlSheetVeryHidden

    ws.Activate
    UI_RefreshDeveloperButtons
End Sub


Sub Settings_NewSheet()
    ' Note: In VBA, you must specify "As String" for each variable,
    ' otherwise they default to Variant type.
    Dim masterSheetName As String, activeUser As String, activeCompany As String
    Dim activeDivision As String, activeAPI As String, activeType As String, activeTransaction As String
    Dim activeEnv As String
    Dim ws As Worksheet
    Dim wsNew As Worksheet ' <-- 1. Create a variable for the new sheet
    
    Set ws = ThisWorkbook.ActiveSheet ' This is your original/source sheet
    
    activeEnv = ws.Range("Environment").value
    activeUser = ws.Range("User").value
    activeCompany = ws.Range("Company").value
    activeDivision = ws.Range("Division").value
    activeAPI = ws.Range("API").value
    activeType = ws.Range("Type").value
    activeTransaction = ws.Range("Transaction").value

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

    ' Create the new sheet (flag prevents Workbook_NewSheet from firing our handler)
    b_AddingSheet = True
    Sheets(masterSheetName).Copy After:=Sheets(Sheets.count)
    b_AddingSheet = False
    
    ' 2. Bind your new sheet variable right here
    Set wsNew = Sheets(Sheets.count)
    
    wsNew.name = newSheetName
    wsNew.Visible = xlSheetVisible
    wsNew.Activate
    UI_DefaultButtons
    
    ' 3. Use wsNew instead of ActiveSheet to be perfectly explicit
    wsNew.Range("Environment").value = activeEnv
    wsNew.Range("User").value = activeUser
    wsNew.Range("Company").value = activeCompany
    wsNew.Range("Division").value = activeDivision
    wsNew.Range("API").value = activeAPI
    wsNew.Range("Type").value = activeType
    wsNew.Range("Transaction").value = ""
    
    ' 4. Pass wsNew here so it filters the NEW sheet, not the original one
    FilterRow8BasedOnPopulatedColumns wsNew
    
    wsNew.Range("Transaction").value = activeTransaction
    UI_UpdateEnvironmentColors
    wsNew.Range("A2").Select
    
    ' 5. Pass wsNew to your logo script!
    Call InsertLogoCrossPlatform(activeEnv, wsNew)
End Sub

Function Settings_SheetExists(partialSheetName As String) As Boolean
    Dim ws As Worksheet
    For Each ws In Worksheets
        If InStr(1, ws.name, partialSheetName) > 0 Then
            Settings_SheetExists = True
            Exit Function
        End If
    Next ws
    Settings_SheetExists = False
End Function

Function Settings_SheetMatch(exactSheetName As String) As Boolean
    Dim ws As Worksheet
    For Each ws In Worksheets
        If UCase(ws.name) = UCase(exactSheetName) Then
            Settings_SheetMatch = True
            Exit Function
        End If
    Next ws
    Settings_SheetMatch = False
End Function

Sub Settings_SaveValues()
    Dim settings As Worksheet
    Set settings = ThisWorkbook.Sheets("Settings")

    ' Write current global values back to the named ranges (D column)
    settings.Range("maxrecs").value = maxRecs
    settings.Range("maxbulk").value = maxbulk
    settings.Range("refreshSeconds").value = refreshSeconds
    settings.Range("righttrim").value = righttrim
    settings.Range("formatting").value = formatting
    settings.Range("splitChar").value = splitChar
    settings.Range("maxtime").value = maxtime
    settings.Range("conoDivi").value = conoDivi
    settings.Range("naming").value = sheetNaming
End Sub

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
        .Sort Key1:=.Columns(2), Order1:=xlAscending, header:=xlNo
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

Sub Tenant_Token(Optional wsTarget As Worksheet = Nothing)
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

    ' ---- Skip if a token attempt already ran this event cycle ----
    If m_b_TokenAttemptedThisCycle Then
        Debug.Print "Tenant_Token: Skipped - token already attempted this cycle."
        Exit Sub
    End If
    m_b_TokenAttemptedThisCycle = True

    On Error GoTo ErrorHandler

    ' Use the caller-supplied sheet when provided (e.g. from Process_Click),
    ' otherwise fall back to the active sheet as before.
    If wsTarget Is Nothing Then
        Set ws = ActiveSheet
    Else
        Set ws = wsTarget
    End If
    ws.Range("J3").value = 0
    m_s_stToken = ""
    bIsCached = False
    
    ' Update selected environment if changed
    If m_s_SelectedEnvironment <> ws.Range("I2").value Then
        m_s_SelectedEnvironment = ws.Range("I2").value
    End If
    
    ' Load tenant configuration
    Tenant_Information
    
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
        Set rngFound = wsEnv.Columns("A").Find(What:=m_s_SelectedEnvironment, _
                                               LookIn:=xlValues, _
                                               LookAt:=xlWhole)
        
        If Not rngFound Is Nothing Then
            ' Retrieve values
            m_s_AccessToken = Trim(wsEnv.Cells(rngFound.row, "E").value)
            m_s_M3user = Trim(wsEnv.Cells(rngFound.row, "D").value)
            m_s_Company = Trim(wsEnv.Cells(rngFound.row, "G").value)
            m_s_Division = Trim(wsEnv.Cells(rngFound.row, "H").value)
            
            ' INTEGRITY CHECK:
            ' If ANY field is blank, cache is invalid.
            If m_s_AccessToken = "" Or _
               m_s_M3user = "" Or _
               m_s_Company = "" Then
               
                Debug.Print "Tenant_Token: Cache incomplete. Forcing fresh login."
                m_s_AccessToken = ""
                bIsCached = False
            Else
                Debug.Print "Tenant_Token: Cache Hit! Skipping validation."
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
        
        Debug.Print "Tenant_Token: Fetching new token..."
        
        config.url = tokenUrl
        config.method = HttpMethod_POST
        config.contentType = "application/x-www-form-urlencoded"
        config.AcceptType = "application/json"
        config.authHeader = ""
        config.body = body
        config.timeoutSeconds = 30
        
        httpResponse = ExecuteRequest(config)
        
        If Not httpResponse.success Then
            ws.Range("J3").value = 0
            ws.Range("User").value = "unauthorized"
            Debug.Print "Tenant_Token: httpResponse="; httpResponse.errorMessage
            Exit Sub
        End If
        
        Set json = ParseJson(httpResponse.body)
        
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
        
        ' 2. Validate & Get User Info
        Debug.Print "Tenant_Token: Validating new session..."
        ws.Range("J3").value = 1
        
        body = "{""program"":""MRS001MI"",""transactions"":[{""transaction"":""GetUserInfo""}]}"
        httpResponse = ExecuteApiPost(m_s_MainUrl, m_s_MiPath & "?m3user=" & m_s_M3user, body)
        
        If httpResponse.success Then
            Set json = ParseJson(httpResponse.body)
            If Not json Is Nothing Then
                Set m_obj_Records = json.item("results")(1).item("records")
            End If
        End If
        
        If httpResponse.success And Not m_obj_Records Is Nothing Then
            Output_GetUserInfo
            DoEvents
            
            ' 3. Save to Cache
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
    manager.AddEnvironment env.name, env.tenant, env.Details, m_s_AccessToken, m_s_MainUrl, m_s_M3user, m_s_Company, m_s_Division
    
    If Not bIsCached Then Log_Activity
    
    Exit Sub
    
ErrorHandler:
    Debug.Print "Tenant_Token: ERROR - " & Err.description
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
    Set rngFound = wsEnv.Columns("A").Find(What:=envName, LookIn:=xlValues, LookAt:=xlWhole)
    
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
        
        Debug.Print "UpdateEnvironmentCache: Updated cache for " & envName
    Else
        Debug.Print "UpdateEnvironmentCache: Warning - Environment '" & envName & "' not found."
    End If
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
        GetOSVersion = objOS.version
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
    decodedText = Base64Decode(encodedText)
    
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
    encodedText = Base64Encode(plainText)
    
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
        If btn.name = btnName Then
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
            .name = btnName
        End With
    End If
End Sub

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

    ' Only read from ActiveSheet when the caller hasn't already set
    ' m_s_SelectedEnvironment (e.g. Tenant_Token sets it from the correct
    ' ws reference before calling here, so we must not overwrite it).
    If m_s_SelectedEnvironment = "" Then
        m_s_SelectedEnvironment = ActiveSheet.Range("Environment").value
    End If
    If m_s_SelectedEnvironment = "" Then
        ClearFields
    Else
        UI_UpdateEnvironmentColors

        Set targetCell = environmentRange.Find(What:=m_s_SelectedEnvironment, LookIn:=xlValues, LookAt:=xlWhole)

        If Not targetCell Is Nothing Then
            jsonString = targetCell.Offset(0, 1).value
            m_s_M3user = targetCell.Offset(0, 3).value
            Set json = ParseJson(jsonString)
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
        m_s_stToken = Base64Encode(User & ":" & Password)
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
''
' Export all standard VBA modules to /Users/ericpronovost/Doppio/vba/
' Run from the Immediate Window: ExportModules
' Requires: Trust access to the VBA project object model
'   (Excel Options ? Trust Center ? Macro Settings)
''
Public Sub ExportModules()
    Const EXPORT_PATH As String = "/Users/ericpronovost/Doppio/vba/"
    Dim c As Object
    Dim exported As Integer

    exported = 0
    For Each c In ThisWorkbook.VBProject.VBComponents
        If c.Type <> 100 And c.name <> "SampleREST" Then
            c.Export EXPORT_PATH & c.name & ".bas"
            exported = exported + 1
        End If
    Next c

    MsgBox exported & " modules exported to " & EXPORT_PATH, vbInformation, "Export Complete"
End Sub

''
' Import all *.bas files from /Users/ericpronovost/Doppio/vba/ into the workbook.
' Existing standard/class modules with matching names are removed first so the
' file on disk becomes the authoritative version.  Document modules (sheets,
' ThisWorkbook) are skipped because they cannot be removed/re-imported.
'
' Run from the Immediate Window: ImportModules
' Requires: Trust access to the VBA project object model
'   (Excel Options ? Trust Center ? Macro Settings)
''
''
' Import class modules only from the export folder.
' Run this first, then Import_FixCollisions, then Import_StandardModules.
'
' Immediate Window: Doppio.Import_ClassModules
''
Public Sub Import_ClassModules()
    Const IMPORT_PATH As String = "/Users/ericpronovost/Doppio/vba/"
    Dim fileName As String
    Dim moduleName As String
    Dim comp As Object
    Dim imported As Integer
    Dim skipped As Integer

    imported = 0
    skipped = 0

    fileName = Dir(IMPORT_PATH & "*.bas")
    Do While fileName <> ""
        moduleName = Left(fileName, Len(fileName) - 4)

        If moduleName = "SampleREST" Then GoTo Import_SkipClassFile
        If Not Import_IsClassFile(IMPORT_PATH & fileName) Then GoTo Import_SkipClassFile

        Set comp = Nothing
        On Error Resume Next
        Set comp = ThisWorkbook.VBProject.VBComponents(moduleName)
        On Error GoTo 0

        If Not comp Is Nothing Then
            If comp.Type = 100 Then
                skipped = skipped + 1
                GoTo Import_SkipClassFile
            End If
            On Error Resume Next
            ThisWorkbook.VBProject.VBComponents.Remove comp
            DoEvents
            Err.Clear
            ThisWorkbook.VBProject.VBComponents.Import IMPORT_PATH & fileName
            On Error GoTo 0
        Else
            On Error Resume Next
            ThisWorkbook.VBProject.VBComponents.Import IMPORT_PATH & fileName
            On Error GoTo 0
        End If

        imported = imported + 1
Import_SkipClassFile:
        fileName = Dir()
    Loop

    Dim msg As String
    msg = imported & " class module(s) imported from " & IMPORT_PATH
    If skipped > 0 Then msg = msg & Chr(13) & skipped & " document module(s) skipped."
    msg = msg & Chr(13) & Chr(13) & "Run Doppio.Import_FixCollisions, then Doppio.Import_StandardModules."
    MsgBox msg, vbInformation, "Class Modules Imported"
End Sub

''
' Import standard (non-class) modules only from the export folder.
' Class modules are excluded � run Import_ClassModules separately if needed.
' Run Import_FixCollisions afterwards if any module got a ""1"" suffix.
'
' Immediate Window: Doppio.Import_StandardModules
''
Public Sub Import_StandardModules()
    Const IMPORT_PATH As String = "/Users/ericpronovost/Doppio/vba/"
    Dim fileName As String
    Dim moduleName As String
    Dim comp As Object
    Dim imported As Integer
    Dim skipped As Integer

    imported = 0
    skipped = 0

    fileName = Dir(IMPORT_PATH & "*.bas")
    Do While fileName <> ""
        moduleName = Left(fileName, Len(fileName) - 4)

        If moduleName = "SampleREST" Then GoTo Import_SkipStdFile
        If Import_IsClassFile(IMPORT_PATH & fileName) Then GoTo Import_SkipStdFile

        Set comp = Nothing
        On Error Resume Next
        Set comp = ThisWorkbook.VBProject.VBComponents(moduleName)
        On Error GoTo 0

        If Not comp Is Nothing Then
            If comp.Type = 100 Then
                skipped = skipped + 1
                GoTo Import_SkipStdFile
            End If
            On Error Resume Next
            ThisWorkbook.VBProject.VBComponents.Remove comp
            DoEvents
            Err.Clear
            ThisWorkbook.VBProject.VBComponents.Import IMPORT_PATH & fileName
            On Error GoTo 0
        Else
            On Error Resume Next
            ThisWorkbook.VBProject.VBComponents.Import IMPORT_PATH & fileName
            On Error GoTo 0
        End If

        imported = imported + 1
Import_SkipStdFile:
        fileName = Dir()
    Loop
    
    Dim msg As String
    msg = imported & " class module(s) imported from " & IMPORT_PATH
    If skipped > 0 Then msg = msg & Chr(13) & skipped & " document module(s) skipped."
    'msg = msg & Chr(13) & Chr(13) & "Run Doppio.Import_FixCollisions, then Doppio.Import_StandardModules."
    MsgBox msg, vbInformation, "Class Modules Imported"
End Sub

''
' Fix "ModuleName1" collisions left behind after an import.
' Processes class modules first (Type 2) then standard modules (Type 1)
' so typed references like "As New ArrayList" always resolve correctly.
' Uses rename-swap so the correct name is never absent from the project.
'
' Immediate Window: Doppio.Import_FixCollisions
''
Public Sub Import_FixCollisions()
    Dim comp As Object
    Dim compOld As Object
    Dim moduleName As String
    Dim fixed As Integer
    Dim targetType As Integer
    Dim pass As Integer

    fixed = 0

    For pass = 1 To 2
        If pass = 1 Then targetType = 2 Else targetType = 1

        For Each comp In ThisWorkbook.VBProject.VBComponents
            If comp.Type = targetType And Len(comp.name) > 1 Then
                If Right(comp.name, 1) = "1" Then
                    moduleName = Left(comp.name, Len(comp.name) - 1)

                    Set compOld = Nothing
                    On Error Resume Next
                    Set compOld = ThisWorkbook.VBProject.VBComponents(moduleName)
                    On Error GoTo 0

                    If Not compOld Is Nothing Then
                        On Error Resume Next
                        compOld.name = moduleName & "_OLD"
                        DoEvents
                        comp.name = moduleName
                        DoEvents
                        If Err.Number = 0 Then
                            ThisWorkbook.VBProject.VBComponents.Remove compOld
                            DoEvents
                            fixed = fixed + 1
                        End If
                        Err.Clear
                        On Error GoTo 0
                    Else
                        On Error Resume Next
                        comp.name = moduleName
                        If Err.Number = 0 Then fixed = fixed + 1
                        Err.Clear
                        On Error GoTo 0
                    End If
                End If
            End If
        Next comp
    Next pass

    MsgBox fixed & " module(s) renamed.", vbInformation, "Fix Collisions Complete"
End Sub

''
' Returns True if a .bas file starts with "VERSION 1.0 CLASS".
''
Private Function Import_IsClassFile(filePath As String) As Boolean
    Dim f As Integer
    Dim firstLine As String

    On Error GoTo Import_NotClass
    f = FreeFile
    Open filePath For Input As #f
    Line Input #f, firstLine
    Close #f
    Import_IsClassFile = (Left(firstLine, 17) = "VERSION 1.0 CLASS")
    Exit Function

Import_NotClass:
    Import_IsClassFile = False
    If f > 0 Then Close #f
End Function

''
' Handles keyword commands for optional/custom modules.
' Returns True if the keyword was handled, so Keywords can Exit Sub.
' Add new optional module keywords here to keep them isolated from core keywords.
''
Public Function Keywords_CustomModules(ByVal value As String) As Boolean
    Keywords_CustomModules = False

    ' journal
    If LCase(value) = "journal" Or LCase(value) = "jrn" Then
        Dim hasJournal As Boolean
        On Error Resume Next
        hasJournal = Application.Run("Xtra_PrelimJournal.ModuleExists")
        On Error GoTo 0
        If Not hasJournal Then
            MsgBox "The Xtra_PrelimJournal module is not installed.", vbExclamation, "Missing Module"
            Keywords_CustomModules = True
            Exit Function
        End If
        On Error Resume Next
        Application.Run "Xtra_PrelimJournal.PreliminaryPrep"
        If Err.Number <> 0 Then
            MsgBox "Error running Journal process: " & Err.description, vbCritical, "Error"
            Err.Clear
        End If
        On Error GoTo 0
        Keywords_CustomModules = True
        Exit Function
    End If

    ' agreements
    If LCase(value) = "agreement" Or LCase(value) = "agr" Then
        Dim hasAgreements As Boolean
        On Error Resume Next
        hasAgreements = Application.Run("Xtra_SupplierAgreements.ModuleExists")
        On Error GoTo 0
        If Not hasAgreements Then
            MsgBox "The Xtra_SupplierAgreements module is not installed.", vbExclamation, "Missing Module"
            Keywords_CustomModules = True
            Exit Function
        End If
        On Error Resume Next
        Application.Run "Xtra_SupplierAgreements.RunAgreementProcess"
        If Err.Number <> 0 Then
            MsgBox "Error running Agreement process: " & Err.description, vbCritical, "Error"
            Err.Clear
        End If
        On Error GoTo 0
        Keywords_CustomModules = True
        Exit Function
    End If

    ' pricelists
    If LCase(value) = "pricelist" Or LCase(value) = "prc" Then
        Dim hasPriceList As Boolean
        On Error Resume Next
        hasPriceList = Application.Run("Xtra_PriceList.ModuleExists")
        On Error GoTo 0
        If Not hasPriceList Then
            MsgBox "The DoppioPriceList module is not installed.", vbExclamation, "Missing Module"
            Keywords_CustomModules = True
            Exit Function
        End If
        On Error Resume Next
        Application.Run "Xtra_PriceList.SetupPriceListSheets"
        If Err.Number <> 0 Then
            MsgBox "Error running Xtra_PriceList setup: " & Err.description, vbCritical, "Error"
            Err.Clear
        End If
        On Error GoTo 0
        Keywords_CustomModules = True
        Exit Function
    End If

    ' Item load
    If LCase(value) = "itemload" Or LCase(value) = "npi" Then
        Dim hasNPI As Boolean

        UI_ShowPleaseWait "Building NPI Sheet"

        On Error Resume Next
        hasNPI = Application.Run("NPI.ModuleExists")
        On Error GoTo 0
        If Not hasNPI Then
            MsgBox "The NPI module is not installed.", vbExclamation, "Missing Module"
            Keywords_CustomModules = True
            Exit Function
        End If
        On Error Resume Next
        Application.Run "NPI.SetupNPISheet"
        UI_KillPleaseWait
        If Err.Number <> 0 Then
            MsgBox "Error running NPI setup: " & Err.description, vbCritical, "Error"
            Err.Clear
        End If
        On Error GoTo 0
        Keywords_CustomModules = True
        Exit Function
    End If

    ' transpose / trn
    ' Transposes the data table (B7 through last used cell).
    ' Data rows are counted from B9; transpose is blocked if >= 50 rows.
    ' Transposed result is placed starting at B9.
    If LCase(value) = "transpose" Or LCase(value) = "trn" Then
        Dim trn_ws As Worksheet
        Set trn_ws = ActiveSheet

        trn_ws.Range("A9:A" & trn_ws.Rows.count).ClearContents
        
        ' -- Count data rows (B9 downward) ----------------------------------
        Dim trn_lastDataRow As Long
        trn_lastDataRow = trn_ws.Cells(trn_ws.Rows.count, "B").End(xlUp).row
        If trn_lastDataRow < 9 Then
            MsgBox "No data found in column B starting at row 9.", vbInformation, "Transpose"
            Keywords_CustomModules = True
            Exit Function
        End If

        Dim trn_dataRowCount As Long
        trn_dataRowCount = trn_lastDataRow - 8  ' row 9 = 1 data row, etc.

        ' Cap at 100 rows -- silently drop anything beyond that
        If trn_dataRowCount > 200 Then trn_lastDataRow = 8 + 200

        ' -- Find last used column (check rows 7-9) -------------------------
        Dim trn_lastCol As Long
        Dim trn_tmpCol As Long
        trn_lastCol = trn_ws.Cells(7, trn_ws.Columns.count).End(xlToLeft).Column
        trn_tmpCol = trn_ws.Cells(8, trn_ws.Columns.count).End(xlToLeft).Column
        If trn_tmpCol > trn_lastCol Then trn_lastCol = trn_tmpCol
        trn_tmpCol = trn_ws.Cells(9, trn_ws.Columns.count).End(xlToLeft).Column
        If trn_tmpCol > trn_lastCol Then trn_lastCol = trn_tmpCol

        If trn_lastCol < 2 Then
            MsgBox "No data found in the table (expected to start at column B).", vbInformation, "Transpose"
            Keywords_CustomModules = True
            Exit Function
        End If

        ' -- Read source range: B7 to lastCol x lastDataRow -----------------
        Dim trn_srcRows As Long   ' number of rows in source  (row 7 .. lastDataRow)
        Dim trn_srcCols As Long   ' number of cols in source  (col B .. lastCol)
        trn_srcRows = trn_lastDataRow - 7 + 1
        trn_srcCols = trn_lastCol - 2 + 1

        Dim trn_srcData As Variant
        trn_srcData = trn_ws.Range(trn_ws.Cells(7, 2), trn_ws.Cells(trn_lastDataRow, trn_lastCol)).value

        ' -- Build transposed array: srcCols rows x srcRows cols ------------
        Dim trn_out() As Variant
        ReDim trn_out(1 To trn_srcCols, 1 To trn_srcRows)
        Dim trn_r As Long, trn_c As Long
        For trn_r = 1 To trn_srcRows
            For trn_c = 1 To trn_srcCols
                trn_out(trn_c, trn_r) = trn_srcData(trn_r, trn_c)
            Next trn_c
        Next trn_r

        Application.ScreenUpdating = False

        ' Remove auto-filter if present
        If trn_ws.AutoFilterMode Then trn_ws.AutoFilterMode = False

        ' -- Clear area large enough to cover both source and destination ---
        ' Source  : rows 7..lastDataRow,        cols B..lastCol
        ' Dest    : rows 9..(9+srcCols-1),      cols B..(B+srcRows-1)
        Dim trn_clearLastRow As Long
        Dim trn_clearLastCol As Long
        trn_clearLastRow = Application.WorksheetFunction.Max(trn_lastDataRow, 9 + trn_srcCols - 1)
        trn_clearLastCol = Application.WorksheetFunction.Max(trn_lastCol, 2 + trn_srcRows - 1)
        trn_ws.Range(trn_ws.Cells(7, 2), trn_ws.Cells(trn_clearLastRow, trn_clearLastCol)).ClearContents

        ' -- Write header row at row 8 --------------------------------------
        ' B8 = "Field Description", C8 = "Field", D8..end = "Data"
        Dim trn_hdrLastCol As Long
        trn_hdrLastCol = 2 + trn_srcRows - 1   ' rightmost column of the transposed block

        ' Clear any previous content / colour in row 8 over this range
        trn_ws.Range(trn_ws.Cells(8, 2), trn_ws.Cells(8, trn_hdrLastCol)).ClearContents

        ' Write labels
        trn_ws.Cells(8, 2).value = "Field Description Type and Length__"
        trn_ws.Cells(8, 3).value = "Field"
        Dim trn_h As Long
        For trn_h = 4 To trn_hdrLastCol
            trn_ws.Cells(8, trn_h).value = "Data"
        Next trn_h

        ' Style: same light-gray background + white bold text as output field headers
        With trn_ws.Range(trn_ws.Cells(8, 2), trn_ws.Cells(8, trn_hdrLastCol))
            .Interior.Color = RGB(128, 128, 128)
            .Font.Color = RGB(255, 255, 255)
            .Font.Bold = True
'            .HorizontalAlignment = xlCenter
        End With

        ' -- Write transposed data starting at B9 ---------------------------
        Dim trn_dataRange As Range
        Set trn_dataRange = trn_ws.Range(trn_ws.Cells(9, 2), _
                                         trn_ws.Cells(9 + trn_srcCols - 1, 2 + trn_srcRows - 1))
        trn_dataRange.value = trn_out
        trn_dataRange.HorizontalAlignment = xlLeft
        trn_dataRange.VerticalAlignment = xlTop

        Application.ScreenUpdating = True

        ' Refit columns/rows
        AutoFit_Click
        AutoFit_ColumnsAndRows False, False

        Keywords_CustomModules = True
        Exit Function
    End If

End Function


