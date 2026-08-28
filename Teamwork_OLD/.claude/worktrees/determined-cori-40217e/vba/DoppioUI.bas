Attribute VB_Name = "DoppioUI"
''
' Doppio UI Module
' Excel user interface operations, dialogs, and formatting
'
' @module DoppioUI
' @author Doppio Group - eric@doppiogroup.com
' @version 2.0
''
Option Explicit

' =============================================================================
' PLEASE WAIT DIALOG
' =============================================================================

''
' Show a "Please Wait" message box on the active sheet
' @param message - Message to display
''
Public Sub UI_ShowPleaseWait(message As String)
    Dim ws As Worksheet
    Dim shp As Shape
    Dim boxHeight As Integer
    Dim leftPosition As Double
    Dim topPosition As Double
    Dim i As Integer
    
    On Error GoTo ErrorHandler
    
    Set ws = ActiveSheet
    boxHeight = 50
    
    ' Calculate height based on line breaks
    For i = 1 To Len(message)
        If Mid(message, i, 1) = vbCr Then
            boxHeight = boxHeight + 22
        End If
    Next i
    
    ' Delete existing shape if present
    On Error Resume Next
    ws.Shapes("Wait").Delete
    On Error GoTo ErrorHandler
    
    ' Position around row 7
    leftPosition = ws.Cells(7, 1).Left
    topPosition = ws.Cells(7, 1).Top
    
    ' Add shape
    Set shp = ws.Shapes.AddShape(msoShapeRectangle, leftPosition + 350, topPosition + 100, 300, boxHeight)
    shp.Name = "Wait"
    
    ' Style the shape
    With shp
        .TextFrame.Characters.text = vbCr & vbTab & message
        .Fill.ForeColor.RGB = RGB(0, 0, 0)
        .TextFrame.Characters.Font.Color = RGB(255, 255, 255)
        .TextFrame.Characters.Font.Name = "Avenir"
        .TextFrame.Characters.Font.FontStyle = "Bold"
    End With
    
    ' Force UI update
    Application.ScreenUpdating = True
    DoEvents
    Application.ScreenUpdating = False
    
    Exit Sub
    
ErrorHandler:
    ' Silently fail - UI helper shouldn't crash the app
    On Error GoTo 0
End Sub

''
' Remove the "Please Wait" message box
''
Public Sub UI_KillPleaseWait()
    On Error Resume Next
    ActiveSheet.Shapes("Wait").Delete
    Application.ScreenUpdating = True
    DoEvents
    On Error GoTo 0
End Sub

' =============================================================================
' ENVIRONMENT UI
' =============================================================================

''
' Update cell colors based on selected environment
''
Public Sub UI_UpdateEnvironmentColors()
    Dim SettingsSheet As Worksheet
    Dim ws As Worksheet
    Dim environmentRange As Range
    Dim foundCell As Range
    Dim sourceCell As Range
    Dim targetCell As Range
    Dim envName As String
    
    On Error GoTo ErrorHandler
    
    Set SettingsSheet = ThisWorkbook.Sheets("Environments")
    Set ws = ActiveSheet
    
    envName = DoppioConfig.Config_SelectedEnvironment
    If envName = "" Then Exit Sub
    
    ' Find the environment in settings
    Set environmentRange = SettingsSheet.Range("A:A")
    Set foundCell = environmentRange.Find(What:=envName, LookIn:=xlValues, LookAt:=xlWhole)
    
    If Not foundCell Is Nothing Then
        Set sourceCell = SettingsSheet.Cells(foundCell.row, 1)
        Set targetCell = ws.Range("I2:I5")
        
        ' Copy colors
        targetCell.Font.Color = sourceCell.Font.Color
        targetCell.Interior.Color = sourceCell.Interior.Color
        
        ' Update sheet tab color (except for AvailableMIs)
        If ws.Name <> "AvailableMIs" Then
            ws.Tab.Color = sourceCell.Interior.Color
        End If
    End If
    
    Exit Sub
    
ErrorHandler:
    On Error GoTo 0
End Sub

''
' Clear environment-related fields on the active sheet
''
Public Sub UI_ClearEnvironmentFields()
    On Error Resume Next
    With ActiveSheet
        .Range("User").value = ""
        .Range("Company").value = ""
        .Range("Division").value = ""
    End With
    On Error GoTo 0
End Sub

' =============================================================================
' STATUS COLUMN
' =============================================================================

''
' Clear the status column (A) from row 9 down
''
Public Sub UI_ClearStatus()
    On Error Resume Next
    ActiveSheet.Range("A9:A" & ActiveSheet.Rows.count).ClearContents
    On Error GoTo 0
End Sub

''
' Set status for a specific row
' @param row - Row number
' @param status - Status text ("OK", "NOK", etc.)
' @param Optional isError - True if this is an error status
''
Public Sub UI_SetRowStatus(row As Long, status As String, Optional isError As Boolean = False)
    On Error Resume Next
    With ActiveSheet.Cells(row, 1)
        .value = status
        If isError Then
            .Font.Color = COLOR_ERROR
        Else
            .Font.Color = COLOR_SUCCESS
        End If
    End With
    On Error GoTo 0
End Sub

' =============================================================================
' DATA AREA
' =============================================================================

''
' Clear the output area (data rows)
' @param ws - Worksheet to clear
' @param Optional headerColor - Color to check for output columns
''
Public Sub UI_ClearOutputArea(ws As Worksheet, Optional headerColor As Long = 0)
    Dim lastCol As Long
    Dim col As Long
    Dim startCol As Long
    
    On Error Resume Next
    
    If headerColor = 0 Then headerColor = COLOR_OUTPUT
    
    ' Find first column with the specified header color
    lastCol = ws.Cells(8, ws.columns.count).End(xlToLeft).column
    startCol = 0
    
    For col = 2 To lastCol
        If ws.Cells(8, col).Interior.Color = headerColor Then
            startCol = col
            Exit For
        End If
    Next col
    
    ' Clear from that column to the end
    If startCol > 0 Then
        ws.Range(ws.Cells(9, startCol), ws.Cells(ws.Rows.count, lastCol)).ClearContents
    End If
    
    On Error GoTo 0
End Sub

''
' Find the first column with a specific header color
' @param ws - Worksheet to search
' @param headerRow - Row containing headers
' @param targetColor - Color to find
' @return Long - Column number or 0 if not found
''
Public Function UI_FindFirstColumnWithColor(ws As Worksheet, headerRow As Long, targetColor As Long) As Long
    Dim lastCol As Long
    Dim col As Long
    
    On Error Resume Next
    lastCol = ws.Cells(headerRow, ws.columns.count).End(xlToLeft).column
    
    For col = 2 To lastCol
        If ws.Cells(headerRow, col).Interior.Color = targetColor Then
            UI_FindFirstColumnWithColor = col
            Exit Function
        End If
    Next col
    
    UI_FindFirstColumnWithColor = 0
    On Error GoTo 0
End Function

' =============================================================================
' COLUMN FORMATTING
' =============================================================================

''
' Auto-fit columns and apply formatting
' @param ws - Worksheet to format
' @param Optional reload - Whether to reload column definitions
' @param Optional mandatoryOnly - Show only mandatory fields
''
Public Sub UI_AutoFitColumns(ws As Worksheet, Optional reload As Boolean = False, Optional mandatoryOnly As Boolean = False)
    Dim lastColumn As Long
    Dim i As Long
    
    On Error GoTo ErrorHandler
    
    ' Optimize performance
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    
    ' Get last column
    lastColumn = ws.Cells(8, ws.columns.count).End(xlToLeft).column
    
    ' Auto-fit rows and columns
    With ws
        .Rows("1:6").AutoFit
        .Rows(1).RowHeight = 60
        .Rows(7).RowHeight = 36
        .columns(1).ColumnWidth = 38
        
        .Rows(7).WrapText = False
        .Rows(7).columns.AutoFit
        .Rows(7).WrapText = True
        .Rows(7).columns.AutoFit
    End With
    
    ' Ensure minimum column widths
    For i = 1 To lastColumn
        If ws.Cells(7, i).ColumnWidth < 12 Then
            ws.columns(i).ColumnWidth = 12
        End If
    Next i
    
    ' Restore settings
    Application.GoTo Reference:="R9C2", Scroll:=True
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    DoEvents
    
    Exit Sub
    
ErrorHandler:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    On Error GoTo 0
End Sub

''
' Apply column header colors based on field direction
' @param ws - Worksheet
' @param col - Column number
' @param direction - Field direction ("I" = Input, "O" = Output, "B" = Both)
' @param isMandatory - Whether field is mandatory
''
Public Sub UI_ApplyColumnHeaderColor(ws As Worksheet, col As Long, direction As String, isMandatory As Boolean)
    Dim headerRow As Long
    headerRow = 8  ' Standard header row
    
    Select Case direction
        Case "I"
            If isMandatory Then
                ws.Cells(headerRow, col).Interior.Color = COLOR_MANDATORY
            Else
                ws.Cells(headerRow, col).Interior.Color = COLOR_OPTIONAL
            End If
        Case Else
            ws.Cells(headerRow, col).Interior.Color = COLOR_OUTPUT
    End Select
End Sub

' =============================================================================
' ROW FILTERING
' =============================================================================

''
' Filter row 8 to show only populated columns
''
Public Sub UI_FilterRow8BasedOnPopulatedColumns()
    Dim ws As Worksheet
    Dim lastCol As Long
    Dim dataLastRow As Long
    Dim col As Long
    Dim hasData As Boolean
    Dim dataRange As Range
    
    On Error Resume Next
    
    Set ws = ActiveSheet
    lastCol = ws.Cells(8, ws.columns.count).End(xlToLeft).column
    dataLastRow = ws.Cells(ws.Rows.count, 2).End(xlUp).row
    
    If dataLastRow < 9 Then dataLastRow = 9
    
    For col = 2 To lastCol
        Set dataRange = ws.Range(ws.Cells(9, col), ws.Cells(dataLastRow, col))
        hasData = Application.WorksheetFunction.CountA(dataRange) > 0
        
        If hasData Then
            ws.columns(col).Hidden = False
        Else
            ' Don't hide columns that have headers (they might be needed)
            If ws.Cells(8, col).value = "" Then
                ws.columns(col).Hidden = True
            End If
        End If
    Next col
    
    On Error GoTo 0
End Sub

' =============================================================================
' TIMER DISPLAY
' =============================================================================

''
' Display elapsed time since start
' @param startTime - Timer value at start
' @param ws - Worksheet to display on
''
Public Sub UI_DisplayElapsedTime(startTime As Single, ws As Worksheet)
    Dim elapsed As Single
    Dim minutes As Integer
    Dim seconds As Integer
    Dim timeStr As String
    
    elapsed = Timer - startTime
    minutes = Int(elapsed / 60)
    seconds = Int(elapsed Mod 60)
    
    If minutes > 0 Then
        timeStr = minutes & "m " & seconds & "s"
    Else
        timeStr = seconds & "s"
    End If
    
    On Error Resume Next
    ws.Range("G6").value = timeStr
    On Error GoTo 0
End Sub

' =============================================================================
' SHEET MANAGEMENT
' =============================================================================

''
' Rename the active sheet (safely)
' @param newName - New sheet name
''
Public Sub UI_RenameSheet(newName As String)
    Dim safeName As String
    Dim i As Integer
    
    On Error Resume Next
    
    ' Remove invalid characters
    safeName = newName
    safeName = Replace(safeName, "/", "-")
    safeName = Replace(safeName, "\", "-")
    safeName = Replace(safeName, "?", "")
    safeName = Replace(safeName, "*", "")
    safeName = Replace(safeName, "[", "(")
    safeName = Replace(safeName, "]", ")")
    safeName = Replace(safeName, ":", "-")
    
    ' Limit length
    If Len(safeName) > 31 Then
        safeName = Left(safeName, 31)
    End If
    
    ' Check if name already exists
    i = 1
    Do While Core_SheetExists(safeName)
        If Len(safeName) > 28 Then
            safeName = Left(safeName, 28)
        End If
        safeName = safeName & " (" & i & ")"
        i = i + 1
    Loop
    
    ActiveSheet.Name = safeName
    
    On Error GoTo 0
End Sub

''
' Copy the Master sheet to create a new working sheet
''
Public Sub UI_CreateNewSheet()
    Dim masterSheetName As String
    Dim newSheetName As String
    Dim sheetNumber As Integer
    Dim activeEnv As String
    Dim activeUser As String
    Dim activeCompany As String
    Dim activeDivision As String
    Dim activeAPI As String
    Dim activeType As String
    Dim ws As Worksheet
    
    On Error GoTo ErrorHandler
    
    Set ws = ThisWorkbook.ActiveSheet
    
    ' Save current values
    activeEnv = ActiveSheet.Range("Environment").value
    activeUser = ActiveSheet.Range("User").value
    activeCompany = ActiveSheet.Range("Company").value
    activeDivision = ActiveSheet.Range("Division").value
    activeAPI = ActiveSheet.Range("API").value
    activeType = ActiveSheet.Range("Type").value
    
    masterSheetName = "Master"
    sheetNumber = 1
    
    ' Find unique sheet name
    Do While Core_SheetExists("Sheet" & sheetNumber)
        sheetNumber = sheetNumber + 1
    Loop
    newSheetName = "Sheet" & sheetNumber
    
    ' Copy master sheet
    ThisWorkbook.Sheets(masterSheetName).Copy After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.count)
    ThisWorkbook.Sheets(ThisWorkbook.Sheets.count).Name = newSheetName
    ThisWorkbook.Sheets(newSheetName).Visible = xlSheetVisible
    ThisWorkbook.Sheets(newSheetName).Activate
    
    ' Restore values
    ActiveSheet.Range("Environment").value = activeEnv
    ActiveSheet.Range("User").value = activeUser
    ActiveSheet.Range("Company").value = activeCompany
    ActiveSheet.Range("Division").value = activeDivision
    ActiveSheet.Range("API").value = activeAPI
    ActiveSheet.Range("Type").value = activeType
    ActiveSheet.Range("Transaction").value = ""
    
    ' Add default buttons
    UI_AddDefaultButtons
    FilterRow8BasedOnPopulatedColumns_New ws
    UI_UpdateEnvironmentColors
    
    ActiveSheet.Range("A2").Select
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error creating new sheet: " & Err.description, vbExclamation
End Sub

' =============================================================================
' BUTTONS
' =============================================================================

''
' Add default buttons to the active sheet
''
Public Sub UI_AddDefaultButtons()
    Dim btn As Button
    Dim ws As Worksheet
    
    On Error GoTo ErrorHandler
    
    Set ws = ActiveSheet
    
    ' Remove existing buttons
    On Error Resume Next
    ws.Buttons.Delete
    On Error GoTo ErrorHandler
    
    ' Add buttons (positions differ between Mac and Windows)
    #If Mac Then
        Set btn = ws.Buttons.Add(8, 80, 69, 29)
    #Else
        Set btn = ws.Buttons.Add(8, 78, 69, 27)
    #End If
    btn.Caption = "Run"
    btn.OnAction = "Run_Click"
    
    #If Mac Then
        Set btn = ws.Buttons.Add(8, 113, 69, 29)
    #Else
        Set btn = ws.Buttons.Add(8, 107, 69, 27)
    #End If
    btn.Caption = "Layout"
    btn.OnAction = "GetLayout_Click"
    
    #If Mac Then
        Set btn = ws.Buttons.Add(80, 80, 69, 19)
    #Else
        Set btn = ws.Buttons.Add(80, 78, 69, 17)
    #End If
    btn.Caption = "Transactions"
    btn.OnAction = "GetTransactions_Click"
    
    #If Mac Then
        Set btn = ws.Buttons.Add(80, 102, 69, 18)
    #Else
        Set btn = ws.Buttons.Add(80, 98, 69, 17)
    #End If
    btn.Caption = "Autofit"
    btn.OnAction = "AutoFit_Click"
    
    #If Mac Then
        Set btn = ws.Buttons.Add(80, 123, 69, 19)
    #Else
        Set btn = ws.Buttons.Add(80, 117, 69, 17)
    #End If
    btn.Caption = "New Sheet"
    btn.OnAction = "NewSheet_Click"
    
    Exit Sub
    
ErrorHandler:
    ' Buttons are optional, don't crash
    On Error GoTo 0
End Sub

' =============================================================================
' USER PROMPTS
' =============================================================================

''
' Prompt user with Yes/No question
' @param message - Question to ask
' @return Boolean - True if user clicked Yes
''
Public Function UI_PromptUser(message As String) As Boolean
    Dim result As VbMsgBoxResult
    result = MsgBox(message, vbYesNo + vbQuestion, "Doppio")
    UI_PromptUser = (result = vbYes)
End Function

''
' Prompt user to stop the program
''
Public Sub UI_PromptExitProgram()
    If PromptUser("Do you want to stop the program?") Then
        KillPleaseWait
        End
    End If
End Sub

' =============================================================================
' VERSION DISPLAY
' =============================================================================

''
' Update version display on the active sheet
''
Public Sub UI_UpdateVersion()
    On Error Resume Next
    ActiveSheet.Range("J2").value = DOPPIO_VERSION
    On Error GoTo 0
End Sub

' =============================================================================
' LOGGING
' =============================================================================

''
' Log an error to the Log sheet
' @param data - Data/context for the error
''
Public Sub UI_LogError(data As String)
    Dim ws As Worksheet
    Dim nextRow As Long
    
    On Error Resume Next
    
    If Not IsSheetVisible("Log") Then Exit Sub
    
    Set ws = ThisWorkbook.Sheets("Log")
    nextRow = ws.Cells(ws.Rows.count, 1).End(xlUp).row + 1
    
    ws.Cells(nextRow, 1).value = Now
    ws.Cells(nextRow, 2).value = DoppioConfig.Config_SelectedEnvironment
    ws.Cells(nextRow, 3).value = Left(data, 500)
    
    On Error GoTo 0
End Sub

' =============================================================================
' FREEZE PANES
' =============================================================================

''
' Set up standard freeze panes for a data sheet
' @param ws - Worksheet to configure
''
Public Sub UI_SetupFreezePanes(ws As Worksheet)
    On Error Resume Next
    
    ws.Activate
    ActiveWindow.FreezePanes = False
    ws.Range("C9").Select
    ActiveWindow.FreezePanes = True
    
    On Error GoTo 0
End Sub

