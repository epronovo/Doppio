Attribute VB_Name = "Xtra_ReplaceValues"

' Tracks whether values have been replaced this session.
' Resets to False when the workbook is closed/reopened (safe default: run forward first).
Private m_bReplaced As Boolean

Sub Xtra_ReplaceValues()
    '--------------------------------------------------------------------------
    ' Xtra_ReplaceValues
    ' Toggles between two states:
    '
    '   FORWARD (m_bReplaced = False):
    '     Looks up each cell value against CTTX40/CTTX15 in the EXPORTMI sheet
    '     and replaces it with the corresponding CTSTKY value.
    '
    '   REVERSE (m_bReplaced = True):
    '     Looks up each cell value against CTSTKY in the EXPORTMI sheet
    '     and replaces it with the corresponding CTTX40 value (CTTX15 fallback).
    '
    ' EXPORTMI sheet layout (row 8 = headers, row 9+ = data):
    '   CTTX40   (full text)
    '   CTTX15   (truncated text)
    '   CTSTKY   (key / replacement value)
    '--------------------------------------------------------------------------

    If m_bReplaced Then
        Call ReplaceValues_Reverse
    Else
        Call ReplaceValues_Forward
    End If

End Sub

' =============================================================================
' FORWARD: CTTX40 / CTTX15  →  CTSTKY
' =============================================================================
Private Sub ReplaceValues_Forward()

    Dim wsActive        As Worksheet
    Dim wsExport        As Worksheet
    Dim exportSheetName As String

    Dim lastCol         As Long
    Dim lastRow         As Long
    Dim exportLastRow   As Long

    Dim colHeader       As String
    Dim cellValue       As String
    Dim replaceValue    As String

    Dim headerRow       As Long
    Dim dataStartRow    As Long

    Dim c               As Long
    Dim r               As Long
    Dim e               As Long

    Dim colCTTX40       As Long
    Dim colCTTX15       As Long
    Dim colCTSTKY       As Long

    Dim found           As Boolean
    Dim changeCount     As Long

    headerRow = 8
    dataStartRow = 9

    Set wsActive = ActiveSheet

    lastCol = wsActive.Cells(headerRow, wsActive.Columns.count).End(xlToLeft).Column
    lastRow = wsActive.UsedRange.row + wsActive.UsedRange.Rows.count - 1
    If lastRow < dataStartRow Then
        MsgBox "No data rows found below row " & headerRow & ".", vbInformation
        Exit Sub
    End If

    changeCount = 0

    For c = 1 To lastCol

        colHeader = Trim(CStr(wsActive.Cells(headerRow, c).value))
        If colHeader = "" Then GoTo NextColumn

        exportSheetName = "EXPORTMI for " & colHeader

        Set wsExport = Nothing
        On Error Resume Next
        Set wsExport = ThisWorkbook.Sheets(exportSheetName)
        On Error GoTo 0
        If wsExport Is Nothing Then GoTo NextColumn

        colCTTX40 = 0: colCTTX15 = 0: colCTSTKY = 0
        Dim expHeaderCol As Long
        expHeaderCol = wsExport.Cells(headerRow, wsExport.Columns.count).End(xlToLeft).Column

        Dim h As Long
        For h = 1 To expHeaderCol
            Select Case UCase(Trim(CStr(wsExport.Cells(headerRow, h).value)))
                Case "CTTX40":  colCTTX40 = h
                Case "CTTX15":  colCTTX15 = h
                Case "CTSTKY":  colCTSTKY = h
            End Select
        Next h

        If colCTSTKY = 0 Then GoTo NextColumn
        If colCTTX40 = 0 And colCTTX15 = 0 Then GoTo NextColumn

        exportLastRow = wsExport.Cells(wsExport.Rows.count, colCTSTKY).End(xlUp).row
        If exportLastRow < dataStartRow Then GoTo NextColumn

        Dim arrTX40() As String
        Dim arrTX15() As String
        Dim arrSTKY() As String
        Dim expRows   As Long
        expRows = exportLastRow - dataStartRow + 1

        ReDim arrTX40(1 To expRows)
        ReDim arrTX15(1 To expRows)
        ReDim arrSTKY(1 To expRows)

        For e = 1 To expRows
            If colCTTX40 > 0 Then arrTX40(e) = Trim(CStr(wsExport.Cells(dataStartRow + e - 1, colCTTX40).value))
            If colCTTX15 > 0 Then arrTX15(e) = Trim(CStr(wsExport.Cells(dataStartRow + e - 1, colCTTX15).value))
            arrSTKY(e) = Trim(CStr(wsExport.Cells(dataStartRow + e - 1, colCTSTKY).value))
        Next e

        For r = dataStartRow To lastRow

            cellValue = Trim(CStr(wsActive.Cells(r, c).value))
            If cellValue = "" Then GoTo NextRow

            replaceValue = "": found = False

            For e = 1 To expRows
                If colCTTX40 > 0 Then
                    If LCase(arrTX40(e)) = LCase(cellValue) Then
                        replaceValue = arrSTKY(e): found = True: Exit For
                    End If
                End If
                If Not found And colCTTX15 > 0 Then
                    If LCase(arrTX15(e)) = LCase(cellValue) Then
                        replaceValue = arrSTKY(e): found = True: Exit For
                    End If
                End If
            Next e

            If found And replaceValue <> "" Then
                wsActive.Cells(r, c).value = replaceValue
                changeCount = changeCount + 1
            End If

NextRow:
        Next r

NextColumn:
    Next c

    m_bReplaced = True

'    MsgBox "Xtra_ReplaceValues complete." & vbNewLine & _
'           changeCount & " value(s) replaced." & vbNewLine & vbNewLine & _
'           "Press the button again to restore original values.", vbInformation, "Done"

End Sub

' =============================================================================
' REVERSE: CTSTKY  →  CTTX40 (CTTX15 fallback)
' =============================================================================
Private Sub ReplaceValues_Reverse()

    Dim wsActive        As Worksheet
    Dim wsExport        As Worksheet
    Dim exportSheetName As String

    Dim lastCol         As Long
    Dim lastRow         As Long
    Dim exportLastRow   As Long

    Dim colHeader       As String
    Dim cellValue       As String
    Dim replaceValue    As String

    Dim headerRow       As Long
    Dim dataStartRow    As Long

    Dim c               As Long
    Dim r               As Long
    Dim e               As Long

    Dim colCTTX40       As Long
    Dim colCTTX15       As Long
    Dim colCTSTKY       As Long

    Dim found           As Boolean
    Dim changeCount     As Long

    headerRow = 8
    dataStartRow = 9

    Set wsActive = ActiveSheet

    lastCol = wsActive.Cells(headerRow, wsActive.Columns.count).End(xlToLeft).Column
    lastRow = wsActive.UsedRange.row + wsActive.UsedRange.Rows.count - 1
    If lastRow < dataStartRow Then
        MsgBox "No data rows found below row " & headerRow & ".", vbInformation
        Exit Sub
    End If

    changeCount = 0

    For c = 1 To lastCol

        colHeader = Trim(CStr(wsActive.Cells(headerRow, c).value))
        If colHeader = "" Then GoTo NextColumn

        exportSheetName = "EXPORTMI for " & colHeader

        Set wsExport = Nothing
        On Error Resume Next
        Set wsExport = ThisWorkbook.Sheets(exportSheetName)
        On Error GoTo 0
        If wsExport Is Nothing Then GoTo NextColumn

        colCTTX40 = 0: colCTTX15 = 0: colCTSTKY = 0
        Dim expHeaderCol As Long
        expHeaderCol = wsExport.Cells(headerRow, wsExport.Columns.count).End(xlToLeft).Column

        Dim h As Long
        For h = 1 To expHeaderCol
            Select Case UCase(Trim(CStr(wsExport.Cells(headerRow, h).value)))
                Case "CTTX40":  colCTTX40 = h
                Case "CTTX15":  colCTTX15 = h
                Case "CTSTKY":  colCTSTKY = h
            End Select
        Next h

        If colCTSTKY = 0 Then GoTo NextColumn
        If colCTTX40 = 0 And colCTTX15 = 0 Then GoTo NextColumn

        exportLastRow = wsExport.Cells(wsExport.Rows.count, colCTSTKY).End(xlUp).row
        If exportLastRow < dataStartRow Then GoTo NextColumn

        Dim arrTX40() As String
        Dim arrTX15() As String
        Dim arrSTKY() As String
        Dim expRows   As Long
        expRows = exportLastRow - dataStartRow + 1

        ReDim arrTX40(1 To expRows)
        ReDim arrTX15(1 To expRows)
        ReDim arrSTKY(1 To expRows)

        For e = 1 To expRows
            If colCTTX40 > 0 Then arrTX40(e) = Trim(CStr(wsExport.Cells(dataStartRow + e - 1, colCTTX40).value))
            If colCTTX15 > 0 Then arrTX15(e) = Trim(CStr(wsExport.Cells(dataStartRow + e - 1, colCTTX15).value))
            arrSTKY(e) = Trim(CStr(wsExport.Cells(dataStartRow + e - 1, colCTSTKY).value))
        Next e

        For r = dataStartRow To lastRow

            cellValue = Trim(CStr(wsActive.Cells(r, c).value))
            If cellValue = "" Then GoTo NextRow

            replaceValue = "": found = False

            ' Reverse: match against CTSTKY, restore CTTX40 (CTTX15 fallback)
            For e = 1 To expRows
                If LCase(arrSTKY(e)) = LCase(cellValue) Then
                    If colCTTX40 > 0 And arrTX40(e) <> "" Then
                        replaceValue = arrTX40(e)
                    ElseIf colCTTX15 > 0 And arrTX15(e) <> "" Then
                        replaceValue = arrTX15(e)
                    End If
                    found = True
                    Exit For
                End If
            Next e

            If found And replaceValue <> "" Then
                wsActive.Cells(r, c).value = replaceValue
                changeCount = changeCount + 1
            End If

NextRow:
        Next r

NextColumn:
    Next c

    m_bReplaced = False

'    MsgBox "Xtra_ReplaceValues: values restored." & vbNewLine & _
'           changeCount & " value(s) reverted.", vbInformation, "Restored"

End Sub
