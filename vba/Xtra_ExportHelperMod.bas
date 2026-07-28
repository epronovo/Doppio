Attribute VB_Name = "Xtra_ExportHelperMod"
''
' Doppio Group - Xtra_ExportHelper
' (c) Eric Pronovost - eric@doppiogroup.com
'
' Helps users build an EXPORTMI call for INDEX tables (8-character table
' names such as MITTRA00). Fired from the Run button (EXPORTMI_BuildSQLQuery)
' BEFORE the SQL query is built, so the fields/where cells it fills flow
' straight into the query.
'
' Trigger (kept simple):
'   - The "table:" label is present in A3, AND
'   - The table name in B3 is exactly 8 characters long.
'
' What it does:
'   1. Calls MNS120MI / Get with FILE = <table name> to fetch the table's
'      key columns (KEY1..KEY9).
'   2. B4 (fields:) is filled with the non-blank keys, comma separated.
'   3. B5 (where:)  is set to the first key that does NOT end in "CONO",
'      with " <> abc" appended.
'
Option Explicit

Public Sub Xtra_ExportHelper(ws As Worksheet)
    Dim tableName As String
    Dim response As apiResponse
    Dim rec As Object
    Dim i As Long
    Dim k As String
    Dim keys As String
    Dim whereKey As String

    On Error GoTo ErrorHandler

    ' --- Guard: only fire for an 8-char table name under a "table:" label ---
    If LCase(Trim(CStr(ws.Range("A3").value))) <> "table:" Then Exit Sub

    tableName = UCase(Trim(CStr(ws.Range("B3").value)))
    If Len(tableName) <> 8 Then Exit Sub

    ' Respect anything the user already typed: only fill blank cells.
    ' If both fields (B4) and where (B5) are already set, do nothing.
    Dim fieldsVal As String
    Dim needFields As Boolean, needWhere As Boolean
    fieldsVal = Trim(CStr(ws.Range("B4").value))
    needFields = (fieldsVal = "" Or fieldsVal = "*")
    needWhere = (Trim(CStr(ws.Range("B5").value)) = "")
    If Not needFields And Not needWhere Then Exit Sub

    ' --- Authenticate and fetch in a single click ---
    ' On the very first Run the session/config isn't primed until the first token
    ' attempt completes, so one try can come up empty. Prime-and-retry here (up to
    ' two attempts) so the fields fill on the first Run without a second click.
    ' Clearing m_b_TokenAttemptedThisCycle lets Tenant_Token actually run each pass
    ' (it is otherwise skipped once per event cycle).
    Dim attempt As Long
    Dim gotData As Boolean
    gotData = False

    For attempt = 1 To 2
        m_b_TokenAttemptedThisCycle = False
        If EXPORTMI_EnsureAuth(ws) Then
            ' MNS120MI/Get returns the table's key fields (KEY1..KEY9)
            response = ExecuteMiGet("MNS120MI", "Get", "FILE=" & tableName)
            If response.success Then
                If Not response.records Is Nothing Then
                    If response.records.count > 0 Then
                        gotData = True
                        Exit For
                    End If
                End If
            End If
        End If
    Next attempt

    If Not gotData Then Exit Sub

    Set rec = response.records(1)

    ' --- Build the fields list and pick the where key ---
    keys = ""
    whereKey = ""
    For i = 1 To 9
        k = ""
        On Error Resume Next
        k = Trim(CStr(rec.item("KEY" & i)))
        On Error GoTo ErrorHandler

        If k <> "" Then
            If keys <> "" Then keys = keys & ", "
            keys = keys & k

            ' First key that does NOT end in CONO drives the where clause
            If whereKey = "" And Right(UCase(k), 4) <> "CONO" Then
                whereKey = k
            End If
        End If
    Next i

    If needFields And keys <> "" Then ws.Range("B4").value = keys
    If needWhere And whereKey <> "" Then ws.Range("B5").value = whereKey & " <> abc"

    ' Cap this auto-built index-table run at 1000 records so a first pass stays
    ' quick, and let the user know how to pull the full list.
    On Error Resume Next
    ThisWorkbook.Sheets("Settings").Range("maxrecs").value = 1000
    Config_MaxRecords = 1000
    On Error GoTo ErrorHandler

    ' Make sure B4/B5 are painted on screen before the modal popup blocks redraw
    DoEvents

    MsgBox "This index table is limited to 1000 records for now." & vbNewLine & vbNewLine & _
           "Set maxrecs to 0 on the Settings sheet when you're ready to run the full list.", _
           vbInformation, "Xtra ExportHelper"

    Exit Sub

ErrorHandler:
    Debug.Print "Xtra_ExportHelper: ERROR - " & Err.description
End Sub
