Attribute VB_Name = "Doppio_UI"
''
' Doppio UI Module
' Excel user interface operations, dialogs, and formatting
'
' @module Doppio_UI
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
    shp.name = "Wait"
    
    ' Style the shape
    With shp
        .TextFrame.Characters.text = vbCr & vbTab & message
        .Fill.ForeColor.RGB = RGB(0, 0, 0)
        .TextFrame.Characters.Font.Color = RGB(255, 255, 255)
        .TextFrame.Characters.Font.name = "Avenir"
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
    Dim wsLogos As Worksheet, ws As Worksheet
    Dim sourceCell As Range, targetCell As Range
    Dim envValue As String, envPrefix As String, envSuffix As String
    Dim r As Long, c As Variant, found As Boolean
    Dim shp As Shape
    
    On Error GoTo ErrorHandler
    
    Set wsLogos = ThisWorkbook.Sheets("Logos")
    Set ws = ActiveSheet
    
    ' Get the selected environment value [cite: 6]
    envValue = ws.Range("Environment").value
    If envValue = "" Then Exit Sub
    
    ' Identify the suffix and prefix for formatting
    envSuffix = UCase(Right(Trim(envValue), 3))
    envPrefix = envValue
    If InStr(envPrefix, " ") > 0 Then envPrefix = Split(envPrefix, " ")(0)
    If InStr(envPrefix, "_") > 0 Then envPrefix = Split(envPrefix, "_")(0)
    envPrefix = LCase(envPrefix)
    
    ' Find logo/format in Logos sheet cols B,D,F,H [cite: 6]
    For Each c In Array(2, 4, 6, 8)
        For r = 1 To wsLogos.Cells(wsLogos.Rows.count, c).End(xlUp).row
            If LCase(Trim(wsLogos.Cells(r, c).value)) = envPrefix Then
                Set sourceCell = wsLogos.Cells(r, c)
                found = True: Exit For
            End If
        Next r
        If found Then Exit For
    Next c
    
    If found Then
        Set targetCell = ws.Range("I2:I5")
        targetCell.Font.Color = sourceCell.Font.Color
        targetCell.Interior.Color = sourceCell.Interior.Color

        ' Update B7:GV7 with environment colors
        With ws.Range("B7:GV7")
            .Font.Color = sourceCell.Font.Color
            .Interior.Color = sourceCell.Interior.Color
        End With

        ' Update A8 with environment colors
        With ws.Range("A8")
            .Font.Color = sourceCell.Font.Color
            .Interior.Color = sourceCell.Interior.Color
        End With
        With ws.Range("B8:GV8")
            .Font.Color = sourceCell.Font.Color
        End With

        ' Update sheet tab color (except for AvailableMIs) [cite: 8, 154]
        If ws.name <> "AvailableMIs" Then
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
    lastCol = ws.Cells(8, ws.Columns.count).End(xlToLeft).Column
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
    lastCol = ws.Cells(headerRow, ws.Columns.count).End(xlToLeft).Column
    
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
    lastColumn = ws.Cells(8, ws.Columns.count).End(xlToLeft).Column
    
    ' Auto-fit rows and columns
    With ws
        .Rows("1:6").AutoFit
        .Rows(1).RowHeight = 60
        .Rows(7).RowHeight = 36
        .Columns(1).ColumnWidth = 38
        
        .Rows(7).WrapText = False
        .Rows(7).Columns.AutoFit
        .Rows(7).WrapText = True
        .Rows(7).Columns.AutoFit
    End With
    
    ' Ensure minimum column widths
    For i = 1 To lastColumn
        If ws.Cells(7, i).ColumnWidth < 12 Then
            ws.Columns(i).ColumnWidth = 12
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
    lastCol = ws.Cells(8, ws.Columns.count).End(xlToLeft).Column
    dataLastRow = ws.Cells(ws.Rows.count, 2).End(xlUp).row
    
    If dataLastRow < 9 Then dataLastRow = 9
    
    For col = 2 To lastCol
        Set dataRange = ws.Range(ws.Cells(9, col), ws.Cells(dataLastRow, col))
        hasData = Application.WorksheetFunction.CountA(dataRange) > 0
        
        If hasData Then
            ws.Columns(col).Hidden = False
        Else
            ' Don't hide columns that have headers (they might be needed)
            If ws.Cells(8, col).value = "" Then
                ws.Columns(col).Hidden = True
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
' Rename the active sheet safely.
' When newName is "" the name is auto-built from the sheet's API + Transaction
' named ranges (same logic previously in the standalone RenameSheet routine).
' @param newName - Desired name, or "" to auto-name from API/Transaction ranges
''
Public Sub UI_RenameSheet(newName As String)
    Dim baseName As String
    Dim safeName As String
    Dim suffix As String
    Dim position As Integer
    Dim i As Integer

    ' --- Auto-name when no name supplied ---
    If newName = "" Then
        On Error Resume Next
        Dim apiVal As String, txnVal As String
        apiVal = Trim(ActiveSheet.Range("API").value)
        txnVal = Trim(ActiveSheet.Range("Transaction").value)
        On Error GoTo 0
        If apiVal = "" Or txnVal = "" Then Exit Sub   ' nothing useful to name from

        ' Apply naming method from Settings (sheetNaming global):
        '   0 = API + Transaction   (e.g. "CRS620MI AddSupplier")
        '   1 = Transaction only    (e.g. "AddSupplier")
        '   2 = API only            (e.g. "CRS620MI")
        '   3 = First 6 of API      (e.g. "CRS620")
        Select Case sheetNaming
            Case 1:  newName = txnVal
            Case 2:  newName = apiVal
            Case 3:  newName = Left(apiVal, 6)
            Case Else   ' 0 or unrecognised — API + Transaction
                newName = apiVal & " " & txnVal
                If Len(newName) > 31 Then newName = apiVal  ' fall back if too long
        End Select
    End If

    ' --- Strip everything from the first "(" onward ---
    position = InStr(1, newName, "(")
    If position > 0 Then newName = Trim(Left(newName, position - 1))

    ' --- Sanitize characters that are invalid in sheet names ---
    safeName = newName
    safeName = Replace(safeName, "/", "")
    safeName = Replace(safeName, "\", "-")
    safeName = Replace(safeName, "?", "")
    safeName = Replace(safeName, "*", "")
    safeName = Replace(safeName, "[", "(")
    safeName = Replace(safeName, "]", ")")
    safeName = Replace(safeName, ":", "-")
    safeName = Trim(safeName)

    ' --- Enforce 31-character limit ---
    If Len(safeName) > 31 Then safeName = Left(safeName, 31)

    ' --- Skip rename if name hasn't changed ---
    If ActiveSheet.name = safeName Then Exit Sub

    ' --- De-duplicate: append (1), (2) … if the name belongs to another sheet ---
    baseName = safeName
    i = 1
    Do While Core_SheetExists(safeName) And ActiveSheet.name <> safeName
        suffix = " (" & i & ")"
        If Len(baseName) + Len(suffix) > 31 Then
            safeName = Left(baseName, 31 - Len(suffix)) & suffix
        Else
            safeName = baseName & suffix
        End If
        i = i + 1
    Loop

    On Error Resume Next
    ActiveSheet.name = safeName
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
    ThisWorkbook.Sheets(ThisWorkbook.Sheets.count).name = newSheetName
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
    FilterRow8BasedOnPopulatedColumns ws
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
' Remove all form buttons from a worksheet
' @param Optional ws - Worksheet to clear buttons from; defaults to ActiveSheet
''
Public Sub UI_RemoveButtons(Optional ws As Worksheet = Nothing)
    Dim btn As Button

    If ws Is Nothing Then Set ws = ActiveSheet

    On Error Resume Next
    For Each btn In ws.Buttons
        btn.Delete
    Next btn
    On Error GoTo 0
End Sub


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
    Dim NextRow As Long
    
    On Error Resume Next
    
    If Not IsSheetVisible("Log") Then Exit Sub
    
    Set ws = ThisWorkbook.Sheets("Log")
    NextRow = ws.Cells(ws.Rows.count, 1).End(xlUp).row + 1
    
    ws.Cells(NextRow, 1).value = Now
    ws.Cells(NextRow, 2).value = Config_SelectedEnvironment
    ws.Cells(NextRow, 3).value = Left(data, 500)
    
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

Public Sub ResetCountFormat(ws As Worksheet)
    With ws.Range("I6")
        .Font.Bold = False
        .Font.Color = RGB(0, 0, 0)
    End With
End Sub

Public Sub CheckMaxRecordsWarning(ws As Worksheet)
    Dim maxRecs As Long
    
    ws.Calculate
    maxRecs = Config_ApiSettings.MaxRecords
    
    If (ws.Range("I6").value >= maxRecs And maxRecs <> 0 And ws.Range("G5").value = "M") Or _
       (maxRecs = 0 And ws.Range("I6").value = 10000) Then
        With ws.Range("I6")
            .Font.Bold = True
            .Font.Color = RGB(255, 0, 0)
        End With
    Else
        ResetCountFormat ws
    End If
End Sub

' =============================================================================
' WORKSHEET CHANGE HANDLER
' =============================================================================

''
' Centralised Worksheet_Change logic for all Doppio sheets.
' Call this from each sheet's Worksheet_Change event with a single line:
'
'     Private Sub Worksheet_Change(ByVal Target As Range)
'         UI_HandleWorksheetChange Target
'     End Sub
'
' @param Target - The Range passed in by the Worksheet_Change event
''
Public Sub UI_HandleWorksheetChange(ByVal Target As Range)

    Const API_ADDRESS As String = "$A$2"
    Const TRANSACTION_ADDRESS As String = "$G$4"
    Const ENVIRONMENT_ADDRESS As String = "$I$2"
    Const KEYWORD_ADDRESS As String = "$A$7"
    Const TOP_CELL_ADDRESS As String = "B9"
    Const OUTPUT_KEYWORD_ADDRESS As String = "A7"

    Dim keyword As String

    On Error GoTo ErrorHandler

    ' When you change the MI program
    If Target.Address(True, True) = API_ADDRESS Then
        GetTransactions_Click
        Target.Worksheet.Range(TRANSACTION_ADDRESS).Select
    End If

    ' When you change the transaction
    If Target.Address(True, True) = TRANSACTION_ADDRESS Then
        Target.Worksheet.Range(TOP_CELL_ADDRESS).Select
        If ActiveSheet.Range("API") = "EXPORTMI" And Left(ActiveSheet.Range("Transaction"), 6) = "Select" Then
            Dim userResponse As VbMsgBoxResult
            userResponse = MsgBox("Would you like to parse reply field (REPL) results?", vbYesNo + vbQuestion, "Parse Results")
            If userResponse = vbYes Then
                If InStr(1, ActiveSheet.Range("B9").value, " from ", vbTextCompare) > 0 Then
                    ActiveSheet.Range("B6").value = "select " & ActiveSheet.Range("B9").value
                End If
                ActiveSheet.Rows("7:" & ActiveSheet.Rows.count).ClearContents
                UI_RemoveButtons
                EXPORTMI_AddButtons
                EXPORTMI_ParseSQLQuery
            Else
                ActiveSheet.Rows("9:" & ActiveSheet.Rows.count).ClearContents
                If Len(ActiveSheet.Range("B6").value) > 7 Then
                    ActiveSheet.Range("C9").value = Sheets("Settings").Range("D12").value
                    ActiveSheet.Range("B9").value = Mid(ActiveSheet.Range("B6").value, 8)
                End If
                UI_RemoveButtons
                UI_DefaultButtons
                ClearStatus
                GetLayoutAll_Click
                KillPleaseWait
            End If
        Else
            If ActiveSheet.Range("API") <> "EXPORTMI" And ActiveSheet.Range("A3").value = "table:  " Then
                UI_RemoveButtons
                UI_DefaultButtons
            End If
            ClearStatus
            KillPleaseWait
        End If
    End If

    ' When you change the environment
    If Target.Address(True, True) = ENVIRONMENT_ADDRESS Then
        Target.Worksheet.Range(TOP_CELL_ADDRESS).Select
        Call InsertLogoCrossPlatform(Target.value, ActiveSheet)
        
        ' Reset the token-attempt flag at the start of each user-initiated event
        m_b_TokenAttemptedThisCycle = False

        UI_ShowPleaseWait "Please Wait... Authenticating"

        ' Suppress cascading Worksheet_Change events while we blank these cells
        Application.EnableEvents = False
        ActiveSheet.Range("User").value = ""
        ActiveSheet.Range("Company").value = ""
        ActiveSheet.Range("Division").value = ""
        Application.EnableEvents = True

        If Target.value <> "" Then
            ClearEnvironmentTokens Target.value
            Tenant_Token
            ClearStatus
            AutoFit_Click
        End If
    End If

    ' When Company or Division is blanked out - if both are now empty, clear
    ' the tenant info on the Environments sheet and fetch a fresh token
    Dim compRangeAddr As String, divRangeAddr As String
    compRangeAddr = ""
    divRangeAddr = ""
    On Error Resume Next
    compRangeAddr = Target.Worksheet.Range("Company").Address(True, True)
    divRangeAddr = Target.Worksheet.Range("Division").Address(True, True)
    On Error GoTo ErrorHandler

    If (compRangeAddr <> "" And divRangeAddr <> "") And _
       (Target.Address(True, True) = compRangeAddr Or Target.Address(True, True) = divRangeAddr) Then

        Dim companyNow As String, divisionNow As String, envNameCD As String
        companyNow = ""
        divisionNow = ""
        envNameCD = ""
        On Error Resume Next
        companyNow = Target.Worksheet.Range("Company").value
        divisionNow = Target.Worksheet.Range("Division").value
        envNameCD = Target.Worksheet.Range("Environment").value
        On Error GoTo ErrorHandler

        If envNameCD <> "" And companyNow = "" And divisionNow = "" Then
            ' Clear company/division columns on the Environments sheet for this tenant
            Dim wsEnvClear As Worksheet
            Dim envRowClear As Range
            On Error Resume Next
            Set wsEnvClear = ThisWorkbook.Sheets("Environments")
            On Error GoTo ErrorHandler

            If Not wsEnvClear Is Nothing Then
                On Error Resume Next
                Set envRowClear = wsEnvClear.Columns("A").Find(What:=envNameCD, LookIn:=xlValues, LookAt:=xlWhole)
                On Error GoTo ErrorHandler
                If Not envRowClear Is Nothing Then
                    wsEnvClear.Cells(envRowClear.row, "G").value = ""
                    wsEnvClear.Cells(envRowClear.row, "H").value = ""
                End If
            End If

            ClearEnvironmentTokens envNameCD
            Tenant_Token
        End If
    End If

    ' When you type a keyword into A7
    If Target.Address = KEYWORD_ADDRESS Then

        Application.EnableEvents = False
        Application.ScreenUpdating = False

        ' Capture the worksheet reference now ? Keywords() may navigate away,
        ' which causes Target.Worksheet to fail if resolved after the call.
        Dim wsKeyword As Worksheet
        Set wsKeyword = Target.Worksheet

        keyword = LCase(Target.value)
        Target.ClearContents

        Keywords (keyword)

        Dim finalTarget As Range

        Select Case LCase(keyword)
            Case "clear"
                wsKeyword.Activate
                MoveToTopOfFrozenSectionOnActiveSheet
                Set finalTarget = wsKeyword.Range(TOP_CELL_ADDRESS)

            Case "help", "settings", "ver", "npi", "itemload"
                ' handled inside Keywords()

            Case "xlsx", "xls", "report"
                ' Return to the original sheet
                wsKeyword.Activate

            Case "ns", "new sheet", "jrn", "journal"
                ' Keywords() activated the new sheet; land on the API cell
                Set finalTarget = ActiveSheet.Range("A2")

            Case "prep", "reset"
                ' Keywords() activated the correct sheet; set finalTarget on
                ' ActiveSheet (not wsKeyword, which may have been deleted by reset)
                ' so the GoTo below fires after EnableEvents=True and DoEvents,
                ' giving A7 proper keyboard focus on the first keystroke
                Set finalTarget = ActiveSheet.Range(OUTPUT_KEYWORD_ADDRESS)

            Case Else
                wsKeyword.Activate
                MoveToTopOfFrozenSectionOnActiveSheet
                Set finalTarget = wsKeyword.Range(OUTPUT_KEYWORD_ADDRESS)
        End Select

        ' If Company or Division is blank after the keyword ran, the token process
        ' didn't populate defaults yet — force a refresh now to get them
        ' (skip for prep/reset/prc/pricelist since wsKeyword may point to a deleted sheet)
        If LCase(keyword) <> "prep" And LCase(keyword) <> "reset" _
           And LCase(keyword) <> "prc" And LCase(keyword) <> "pricelist" _
           And LCase(keyword) <> "xlsx" And LCase(keyword) <> "xls" _
           And LCase(keyword) <> "report" Then
            If wsKeyword.Range("Environment").value <> "" Then
                If wsKeyword.Range("Company").value = "" Or wsKeyword.Range("Division").value = "" Then
                    Tenant_Token
                End If
            End If
        End If

        Application.EnableEvents = True
        Application.ScreenUpdating = True
        DoEvents

        ' Goto AFTER DoEvents so screen repaint and event flush can't steal focus back
        If Not finalTarget Is Nothing Then
            Application.GoTo Reference:=finalTarget
        End If
    End If

    Exit Sub

ErrorHandler:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    MsgBox "An error occurred: " & Err.description, vbExclamation, "Error"
End Sub

' =============================================================================
' WORKBOOK EVENT HANDLERS
' =============================================================================

''
' Centralised Workbook_Open logic.
' Call this from ThisWorkbook with a single line:
'
'     Private Sub Workbook_Open()
'         UI_HandleWorkbookOpen
'     End Sub
''
Public Sub UI_HandleWorkbookOpen()

    Const EXCLUDE_SHEETS As String = "Settings,Log,Environments,Transactions,AvailableMIs,Help,Versions,Logos"

    Dim sheet As Worksheet
    Dim excludeArr As Variant
    Dim settings As Worksheet

    RecordCache_RemoveBearerTokens
    Environments_Load

    Application.ScreenUpdating = False

    excludeArr = Split(EXCLUDE_SHEETS, ",")

    For Each sheet In ThisWorkbook.Worksheets
        If isError(Application.match(sheet.name, excludeArr, 0)) Then
            On Error Resume Next
            sheet.Range("J3").value = 0
            On Error GoTo 0
        End If
    Next sheet

    Set settings = ThisWorkbook.Sheets("Settings")
    With settings
        maxRecs = .Range("maxrecs").value
        maxbulk = .Range("maxbulk").value
        refreshSeconds = .Range("refreshSeconds").value
        righttrim = .Range("righttrim").value
        formatting = .Range("formatting").value
        splitChar = .Range("splitChar").value
        maxtime = .Range("maxtime").value
        conoDivi = .Range("conoDivi").value
        sheetNaming = .Range("naming").value
    End With

    Application.ScreenUpdating = True

End Sub

''
' Centralised Workbook_SheetActivate logic.
' Call this from ThisWorkbook with a single line:
'
'     Private Sub Workbook_SheetActivate(ByVal Sh As Object)
'         UI_HandleSheetActivate Sh
'     End Sub
'
' @param Sh - The sheet being activated, passed from Workbook_SheetActivate
''
Public Sub UI_HandleSheetActivate(ByVal Sh As Object)

    Const SYSTEM_SHEETS As String = "Master,Log,Cache,Settings,Environments,AvailableMIs,Transactions,Help,Versions,Logos"

    Dim systemArr As Variant
    Dim matchResult As Variant
    Dim apiName As String
    Dim g5Value As String

    systemArr = Split(SYSTEM_SHEETS, ",")

    On Error Resume Next
    matchResult = Application.match(Sh.name, systemArr, 0)
    On Error GoTo 0

    If Not isError(matchResult) Then Exit Sub  ' system/master sheet ? nothing to do

    On Error Resume Next
    apiName = Sh.Range("A2").value
    g5Value = Sh.Range("G5").value
    On Error GoTo 0

    If apiName = "" Then Exit Sub

    ' Don't interrupt an active copy/cut operation
    If Application.CutCopyMode <> False Then Exit Sub

    ' Load transactions only if G5 is blank (avoids redundant reloads)
    If Trim(g5Value) = "" Then
        GetTransactions_Click
    End If

End Sub

''
' Centralised Workbook_NewSheet logic.
' Call this from ThisWorkbook with a single line:
'
'     Private Sub Workbook_NewSheet(ByVal Sh As Object)
'         UI_HandleNewSheet Sh
'     End Sub
'
' Delegates to Settings_NewSheet, which copies from Master and sets up the
' sheet properly. The b_AddingSheet flag in bas prevents the Copy
' inside Settings_NewSheet from triggering this handler a second time.
'
' @param Sh - The new sheet, passed from Workbook_NewSheet
''
Public Sub UI_HandleNewSheet(ByVal Sh As Object)

    ' Ignore programmatic sheet additions (e.g. from Settings_NewSheet itself)
    If b_AddingSheet Then Exit Sub

    ' A user clicked "+". Delete the blank sheet Excel just created and
    ' run Settings_NewSheet instead so it copies from Master correctly.
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    Sh.Delete
    Application.DisplayAlerts = True
    Application.EnableEvents = True

    Settings_NewSheet

End Sub

' =============================================================================
' SYSTEM SHEET VISIBILITY
' =============================================================================

''
' Hide all system/utility sheets.
' @param showHelp - Pass True to leave the Help sheet visible (default False)
''
Public Sub UI_HideSystemSheets(Optional showHelp As Boolean = False)
    Dim sheetNames As Variant
    Dim i As Integer

    sheetNames = Array("Master", "Log", "Cache", "Settings", "Environments", _
                       "AvailableMIs", "Transactions", "Versions", "Logos", "Help")

    On Error Resume Next
    For i = 0 To UBound(sheetNames)
        Sheets(sheetNames(i)).Visible = False
    Next i
    On Error GoTo 0

    If showHelp Then
        On Error Resume Next
        Sheets("Help").Visible = True
        On Error GoTo 0
    End If
End Sub


' =============================================================================
' SHEET MANAGEMENT
' =============================================================================

Public Sub UI_DeleteSheets()
    Dim ws As Worksheet
    Dim protectedSheets As Variant
    Dim sheetName As String

    ' List of sheets that should Not be deleted
    protectedSheets = Array("Environments", "Master", "AvailableMIs", "Transactions", "Versions", "Help", "Settings", "Log", "Cache", "Logos")

    ' Add the active sheet To the protected sheets list
    If Not ActiveSheet Is Nothing Then
        ReDim Preserve protectedSheets(UBound(protectedSheets) + 1)
        protectedSheets(UBound(protectedSheets)) = ActiveSheet.name
    End If

    Application.DisplayAlerts = False
    For Each ws In ThisWorkbook.Sheets
        sheetName = ws.name
        If isError(Application.match(sheetName, protectedSheets, 0)) Then
            Worksheets(sheetName).Delete
        End If
    Next ws
    Application.DisplayAlerts = True
End Sub


Public Sub UI_DefaultButtons()
    Dim btn1, btn2, btn3, btn4, btn5, btn6 As Button
    Dim ws As Worksheet
    Set ws = ActiveSheet

    Rows(1).RowHeight = 60
    Rows(7).RowHeight = 36

    ' Add buttons
    #If Mac Then
        Set btn1 = ws.Buttons.Add(8, 82, 69, 29)
        Set btn2 = ws.Buttons.Add(8, 115, 69, 29)
        Set btn3 = ws.Buttons.Add(80, 82, 69, 19)
        Set btn4 = ws.Buttons.Add(80, 104, 69, 18)
        Set btn5 = ws.Buttons.Add(80, 125, 69, 19)
        Set btn6 = ws.Buttons.Add(151, 82, 12, 62)
    #Else
        Set btn1 = ws.Buttons.Add(8, 78, 69, 27)
        Set btn2 = ws.Buttons.Add(8, 107, 69, 27)
        Set btn3 = ws.Buttons.Add(80, 78, 69, 17)
        Set btn4 = ws.Buttons.Add(80, 98, 69, 17)
        Set btn5 = ws.Buttons.Add(80, 117, 69, 17)
        Set btn6 = ws.Buttons.Add(151, 78, 12, 61)
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
    btn6.Caption = "Rep"
    btn6.OnAction = "Xtra_ReplaceValues.Xtra_ReplaceValues"
    btn6.Visible = Config_Developer

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
    ws.Range("I6").Formula = "=MAX(COUNTA(B9:B1048576),COUNTA(C9:C1048576),COUNTA(D9:D1048576),COUNTA(E9:E1048576),COUNTA(F9:F1048576),COUNTA(G9:G1048576),COUNTA(H9:H1048576),COUNTA(I9:I1048576),COUNTA(J9:J1048576),COUNTA(K9:K1048576),COUNTA(L9:L1048576),COUNTA(M9:M1048576),COUNTA(N9:N1048576),COUNTA(O9:O1048576),COUNTA(P9:P1048576),COUNTA(Q9:Q1048576),COUNTA(R9:R1048576),COUNTA(S9:S1048576),COUNTA(T9:T1048576),COUNTA(U9:U1048576),COUNTA(V9:V1048576),COUNTA(W9:W1048576),COUNTA(X9:X1048576),COUNTA(Y9:Y1048576),COUNTA(Z9:Z1048576),COUNTA(AA9:AA1048576),COUNTA(AB9:AB1048576),COUNTA(AC9:AC1048576),COUNTA(AD9:AD1048576),COUNTA(AE9:AE1048576))"
End Sub

''
' Refresh developer-only button visibility on the active sheet.
' Call this after settings are loaded (e.g. from Settings_CopyDefaults).
''
Public Sub UI_RefreshDeveloperButtons()
    Dim ws As Worksheet
    Dim btn As Button
    Set ws = ActiveSheet
    For Each btn In ws.Buttons
        If btn.OnAction = "Xtra_ReplaceValues.Xtra_ReplaceValues" Then
            btn.Visible = Config_Developer
        End If
    Next btn
End Sub


