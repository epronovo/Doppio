Attribute VB_Name = "M3_LoadManager"
' ============================================================
'  M3 Configuration Load Manager  -  macOS version
'
'  Approach:
'    - Codes are read from the sheet and embedded directly
'      into the Python script as a list literal.
'    - Python script is written via VBA Open (TMPDIR writes OK).
'    - Python is executed via MacScript do shell script.
'    - Results are read back via MacScript cat.
'    - No COM objects, no Scripting.Dictionary, no ADODB.
'
'  Prerequisites:
'    - Python 3  (verify: which python3 in Terminal)
'    - doppio.db at the path in DB_PATH below
'
'  How to use:
'    Option+F8 -> RunAPICheck -> Run
' ============================================================

Option Explicit

' ---- Configuration ------------------------------------------
Private Const DB_PATH               As String = "/Users/ericpronovost/sqlite/doppio.db"
Private Const PYTHON_PATH           As String = "/Library/Frameworks/Python.framework/Versions/3.12/bin/python3"
Private Const DATA_START            As Long = 5
Private Const COL_CODE              As Long = 3
Private Const COL_API               As Long = 8
Private Const COL_LTYPE             As Long = 9
Private Const COL_LRESP             As Long = 10
Private Const RESPONSIBLE_API       As String = "Eric"
Private Const RESPONSIBLE_MANUAL    As String = "Lead Consultant"
Private Const HIGH_VOLUME_THRESHOLD As Long = 50
Private Const TMP_RESULTS           As String = "m3_results_out.csv"
Private Const TMP_SCRIPT            As String = "m3_check.py"

' ============================================================
'  MAIN
' ============================================================
Public Sub RunAPICheck()
    Dim ws      As Worksheet
    Dim lastRow As Long
    Dim updated As Long
    Dim errors  As Long

    On Error GoTo ErrHandler

    Set ws = ThisWorkbook.Sheets(ThisWorkbook.Sheets(1).name)
    lastRow = FindLastDataRow(ws)

    If lastRow < DATA_START Then
        MsgBox "No data rows found in '" & ThisWorkbook.Sheets(1).name & "'.", vbInformation
        Exit Sub
    End If

    Dim tmpDir As String
    tmpDir = Environ("TMPDIR")
    If Right(tmpDir, 1) <> "/" Then tmpDir = tmpDir & "/"

    Dim pathResults As String: pathResults = tmpDir & TMP_RESULTS
    Dim pathScript  As String: pathScript = tmpDir & TMP_SCRIPT

    ' Step 1 - read codes from sheet, build Python list literal
    Dim codesList As String
    codesList = BuildCodesList(ws, DATA_START, lastRow)

    ' Step 2 - write Python script using VBA Open (TMPDIR writes are allowed)
    WritePythonScript pathScript, codesList, pathResults

    ' Step 3 - execute via shell (fire and forget - quoting issues make return unreliable)
    RunScript pathScript

    ' Step 4 - read results via shell cat into parallel arrays
    Dim resCodes()  As String
    Dim resValues() As String
    Dim resCount    As Long
    ReadResultsCSV pathResults, resCodes, resValues, resCount

    If resCount = 0 Then
        MsgBox "No results returned. Check:" & Chr(13) & Chr(13) & _
               "  1. PYTHON_PATH = " & PYTHON_PATH & Chr(13) & _
               "     Verify: open Terminal, type  which python3" & Chr(13) & Chr(13) & _
               "  2. DB_PATH = " & DB_PATH & Chr(13) & _
               "     Verify the file exists at that path" & Chr(13) & Chr(13) & _
               "  3. Excel has Automation permission" & Chr(13) & _
               "     System Settings -> Privacy -> Automation -> Excel", _
               vbCritical, "M3 Load Manager"
        GoTo Cleanup
    End If

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    Dim r As Long
    For r = DATA_START To lastRow
        On Error GoTo ErrRow
        Dim progCode As String
        progCode = Trim(CStr(ws.Cells(r, COL_CODE).value))
        If progCode = "" Then GoTo NextRow

        Dim checkCode As String
        checkCode = Trim(Split(progCode, "/")(0))

        Dim apiExists As Boolean
        apiExists = LookupCode(checkCode, resCodes, resValues, resCount)

        With ws.Cells(r, COL_API)
            .value = IIf(apiExists, "Yes", "No")
            ApplyStatusFormat .Cells, .value
        End With

        Dim loadType As String
        If apiExists Then
            loadType = "API"
        ElseIf GetSheetRowCount(progCode) >= HIGH_VOLUME_THRESHOLD Then
            loadType = "Webservice"
        Else
            loadType = "Manual"
        End If

        With ws.Cells(r, COL_LTYPE)
            .value = loadType
            ApplyLoadTypeFormat .Cells, loadType
        End With

        ws.Cells(r, COL_LRESP).value = _
            IIf(loadType = "Manual", RESPONSIBLE_MANUAL, RESPONSIBLE_API)

        updated = updated + 1
        GoTo NextRow
ErrRow:
        errors = errors + 1
        Resume NextRow
NextRow:
    Next r

    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic

    MsgBox "API check complete." & Chr(13) & Chr(13) & _
           "Rows updated : " & updated & Chr(13) & _
           "Errors       : " & errors, vbInformation, "M3 Load Manager"

Cleanup:
    CleanTempFiles tmpDir
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Exit Sub

ErrHandler:
    MsgBox "Error " & Err.Number & ": " & Err.description, vbCritical, "M3 Load Manager"
    Resume Cleanup
End Sub

' ============================================================
'  BUILD CODES LIST
'  Returns a Python list literal: ["CRS040","CRS025","PDS010"]
' ============================================================
Private Function BuildCodesList(ws As Worksheet, firstRow As Long, _
                                 lastRow As Long) As String
    Dim items As String
    Dim r     As Long
    For r = firstRow To lastRow
        Dim code As String
        code = Trim(CStr(ws.Cells(r, COL_CODE).value))
        If code <> "" Then
            code = Trim(Split(code, "/")(0))
            If items <> "" Then items = items & ","
            items = items & Chr(34) & code & Chr(34)
        End If
    Next r
    BuildCodesList = "[" & items & "]"
End Function

' ============================================================
'  WRITE PYTHON SCRIPT
'  VBA Open works for writes to TMPDIR - only reads are blocked
' ============================================================
Private Sub WritePythonScript(scriptPath As String, codesList As String, _
                               resultsPath As String)
    Dim fNum As Integer
    fNum = FreeFile
    Open scriptPath For Output As #fNum
    Print #fNum, "import sqlite3, sys, os"
    Print #fNum, "db_path  = r'" & DB_PATH & "'"
    Print #fNum, "out_file = r'" & resultsPath & "'"
    Print #fNum, "codes    = " & codesList
    Print #fNum, "if not os.path.exists(db_path):"
    Print #fNum, "    sys.exit('DB not found: ' + db_path)"
    Print #fNum, "conn = sqlite3.connect(db_path)"
    Print #fNum, "cur  = conn.cursor()"
    Print #fNum, "with open(out_file, 'w') as out:"
    Print #fNum, "    for code in codes:"
    Print #fNum, "        try:"
    Print #fNum, "            cur.execute('SELECT MINM FROM cmipgm WHERE MINM = ?', (code + 'MI',))"
    Print #fNum, "            out.write(code + ',' + ('1' if cur.fetchone() else '0') + chr(10))"
    Print #fNum, "        except Exception as e:"
    Print #fNum, "            out.write(code + ',error' + chr(10))"
    Print #fNum, "conn.close()"
    Close #fNum
End Sub

' ============================================================
'  RUN SCRIPT  /  READ RESULTS  /  LOOKUP
' ============================================================
Private Sub RunScript(scriptPath As String)
    ' Single-quoted paths inside the shell command avoid all nested-quote issues
    Dim appleCmd As String
    appleCmd = "do shell script " & Chr(34) & _
               PYTHON_PATH & " '" & scriptPath & "'" & _
               Chr(34)
    On Error Resume Next
    MacScript (appleCmd)
    On Error GoTo 0
End Sub

Private Sub ReadResultsCSV(filePath As String, _
                            ByRef outCodes() As String, _
                            ByRef outValues() As String, _
                            ByRef outCount As Long)
    outCount = 0

    ' Try VBA Open first; fall back to shell cat if sandboxed
    Dim csvContent As String
    csvContent = TryReadFile(filePath)
    If csvContent = "" Then csvContent = TryShellCat(filePath)
    If csvContent = "" Then Exit Sub

    ' Normalise line endings (MacScript returns Chr(13), VBA Open returns Chr(10))
    csvContent = Replace(csvContent, Chr(13) & Chr(10), Chr(13))
    csvContent = Replace(csvContent, Chr(10), Chr(13))

    Dim lines() As String
    lines = Split(csvContent, Chr(13))
    ReDim outCodes(UBound(lines))
    ReDim outValues(UBound(lines))

    Dim i As Long
    For i = 0 To UBound(lines)
        Dim parts() As String
        parts = Split(Trim(lines(i)), ",")
        If UBound(parts) >= 1 Then
            Dim k As String: k = Trim(parts(0))
            If k <> "" Then
                outCodes(outCount) = k
                outValues(outCount) = Trim(parts(1))
                outCount = outCount + 1
            End If
        End If
    Next i
End Sub

Private Function TryReadFile(filePath As String) As String
    Dim result As String
    Dim fNum   As Integer
    Dim line   As String
    On Error Resume Next
    fNum = FreeFile
    Open filePath For Input As #fNum
    If Err.Number <> 0 Then TryReadFile = "": Exit Function
    On Error GoTo 0
    Do While Not EOF(fNum)
        Line Input #fNum, line
        result = result & line & Chr(13)
    Loop
    Close #fNum
    TryReadFile = result
End Function

Private Function TryShellCat(filePath As String) As String
    Dim result As String
    On Error Resume Next
    result = MacScript("do shell script ""cat '" & filePath & "'""")
    On Error GoTo 0
    TryShellCat = result
End Function

Private Function LookupCode(code As String, _
                             codes() As String, _
                             values() As String, _
                             count As Long) As Boolean
    Dim i As Long
    For i = 0 To count - 1
        If codes(i) = code Then
            LookupCode = (values(i) = "1")
            Exit Function
        End If
    Next i
    LookupCode = False
End Function

Private Sub CleanTempFiles(tmpDir As String)
    On Error Resume Next
    MacScript ("do shell script ""rm -f " & _
        Chr(34) & tmpDir & TMP_RESULTS & Chr(34) & " " & _
        Chr(34) & tmpDir & TMP_SCRIPT & Chr(34) & """")
    On Error GoTo 0
End Sub

' ============================================================
'  SHEET HELPERS
' ============================================================
Private Function FindLastDataRow(ws As Worksheet) As Long
    Dim r As Long
    r = ws.Cells(ws.Rows.count, COL_CODE).End(xlUp).row
    If r < DATA_START Then r = DATA_START - 1
    FindLastDataRow = r
End Function

Private Function GetSheetRowCount(progCode As String) As Long
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Sheets
        If InStr(1, ws.name, Split(progCode, "/")(0), vbTextCompare) > 0 Then
            GetSheetRowCount = ws.Cells(ws.Rows.count, 1).End(xlUp).row - 4
            If GetSheetRowCount < 0 Then GetSheetRowCount = 0
            Exit Function
        End If
    Next ws
    GetSheetRowCount = 0
End Function

' ============================================================
'  FORMATTING
' ============================================================
Private Sub ApplyStatusFormat(c As Range, status As String)
    Select Case status
        Case "Yes":  c.Interior.Color = RGB(198, 239, 206): c.Font.Color = RGB(0, 97, 0)
        Case "No":   c.Interior.Color = RGB(255, 199, 206): c.Font.Color = RGB(156, 0, 6)
        Case Else:   c.Interior.Color = RGB(242, 242, 242): c.Font.Color = RGB(89, 89, 89)
    End Select
    c.Font.Bold = (status = "Yes" Or status = "No")
End Sub

Private Sub ApplyLoadTypeFormat(c As Range, loadType As String)
    Select Case loadType
        Case "API":        c.Interior.Color = RGB(198, 239, 206): c.Font.Color = RGB(0, 97, 0)
        Case "Webservice": c.Interior.Color = RGB(255, 235, 156): c.Font.Color = RGB(156, 101, 0)
        Case "Manual":     c.Interior.Color = RGB(217, 217, 217): c.Font.Color = RGB(64, 64, 64)
        Case Else:         c.Interior.Color = RGB(255, 255, 255): c.Font.Color = RGB(0, 0, 0)
    End Select
End Sub

' ============================================================
'  PUBLIC UTILITIES
' ============================================================
Public Sub ResetLoadColumns()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(ThisWorkbook.Sheets(1).name)
    Dim lastRow As Long
    lastRow = FindLastDataRow(ws)
    If lastRow < DATA_START Then Exit Sub
    If MsgBox("Clear all API / Load Type / Responsibility values?", _
              vbYesNo + vbQuestion, "Reset") = vbNo Then Exit Sub
    Dim r As Long
    For r = DATA_START To lastRow
        Dim c As Range
        For Each c In ws.Range(ws.Cells(r, COL_API), ws.Cells(r, COL_LRESP))
            c.value = "": c.Interior.ColorIndex = xlNone
            c.Font.Color = RGB(0, 0, 0): c.Font.Bold = False
        Next c
        With ws.Cells(r, COL_API)
            .value = "Pending check"
            .Font.Italic = True
            .Font.Color = RGB(127, 127, 127)
        End With
    Next r
    MsgBox "Reset complete.", vbInformation
End Sub

Public Sub OverrideRow()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(ThisWorkbook.Sheets(1).name)
    Dim r As Long: r = ActiveCell.row
    If r < DATA_START Then
        MsgBox "Click a data row first.", vbExclamation: Exit Sub
    End If
    Dim lt As String
    lt = InputBox("Load Type for row " & r & ":" & Chr(13) & _
                  "(API / Webservice / Manual)", "Override", _
                  ws.Cells(r, COL_LTYPE).value)
    If lt = "" Then Exit Sub
    lt = StrConv(lt, vbProperCase)
    If lt = "Api" Then lt = "API"
    If lt <> "API" And lt <> "Webservice" And lt <> "Manual" Then
        MsgBox "Use: API, Webservice, or Manual", vbExclamation: Exit Sub
    End If
    ws.Cells(r, COL_LTYPE).value = lt
    ApplyLoadTypeFormat ws.Cells(r, COL_LTYPE), lt
    ws.Cells(r, COL_LRESP).value = IIf(lt = "Manual", RESPONSIBLE_MANUAL, RESPONSIBLE_API)
End Sub

Public Sub ShowTempDir()
    MsgBox "Sandbox temp dir: " & Environ("TMPDIR"), vbInformation, "TMPDIR"
End Sub


