VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MatrixManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Class name: MatrixManager
Option Explicit

Private MatrixData As MatrixData

' === UTILITY METHODS ===
Private Function InCollection(col As Collection, value As String) As Boolean
    Dim i As Long
    For i = 1 To col.count
        If col(i) = value Then
            InCollection = True
            Exit Function
        End If
    Next i
    InCollection = False
End Function

Private Function FindIndex(arr() As String, value As String) As Long
    Dim i As Long
    For i = LBound(arr) To UBound(arr)
        If arr(i) = value Then
            FindIndex = i
            Exit Function
        End If
    Next i
    FindIndex = -1
End Function

' === MATRIX BUILD ===
Public Sub BuildMatrix(ws As Worksheet)
    Dim lastRow As Long, lastCol As Long
    Dim dataRange As Range, data As Variant
    Dim rowKeys As New Collection, colKeys As New Collection
    Dim rowMap() As String, colMap() As String
    Dim outputData() As Variant
    Dim i As Long, rowKey As String, colKey As String
    Dim rowIndex As Long, colIndex As Long
    Dim totalRows As Long
    Dim outputStartCell As Range, outputRange As Range
'    Dim Headers(1 To 3) As String
    Dim startCell As Range
    Dim outputOffsetColumns As Long

    Set startCell = ws.Range("B8")
    outputOffsetColumns = 0
    lastRow = ws.Cells(ws.Rows.count, startCell.Column).End(xlUp).row
    lastCol = ws.Cells(startCell.row, ws.Columns.count).End(xlToLeft).Column
    Set dataRange = ws.Range(startCell.Offset(1, 0), ws.Cells(lastRow, lastCol))
    
    ' Safety check: ensure only 3 columns are selected
    If dataRange.Columns.count <> 3 Then Exit Sub

    data = dataRange.value

    ' Store validated input and headers for use in Unpivot
    Set MatrixData = New MatrixData
'    Headers(1) = CStr(ws.Cells(startCell.row, startCell.column).value)
'    Headers(2) = CStr(ws.Cells(startCell.row, startCell.column + 1).value)
'    Headers(3) = CStr(ws.Cells(startCell.row, startCell.column + 2).value)
'    MatrixData.SetHeaders Headers
    ws.Range("C6").value = CStr(ws.Cells(startCell.row, startCell.Column).value)
    ws.Range("D6").value = CStr(ws.Cells(startCell.row, startCell.Column + 1).value)
    ws.Range("E6").value = CStr(ws.Cells(startCell.row, startCell.Column + 2).value)

    ' Collect unique row and column keys
    For i = 1 To UBound(data, 1)
        rowKey = CStr(data(i, 1))
        colKey = CStr(data(i, 2))
        If Not InCollection(rowKeys, rowKey) Then rowKeys.Add rowKey
        If Not InCollection(colKeys, colKey) Then colKeys.Add colKey
    Next i

    totalRows = rowKeys.count + 1
    lastCol = colKeys.count + 1

    ReDim rowMap(1 To rowKeys.count)
    ReDim colMap(1 To colKeys.count)
    For i = 1 To rowKeys.count: rowMap(i) = rowKeys(i): Next i
    For i = 1 To colKeys.count: colMap(i) = colKeys(i): Next i

    ReDim outputData(1 To totalRows, 1 To lastCol)
    outputData(1, 1) = CStr(ws.Cells(startCell.row, startCell.Column).value)
    For i = 1 To colKeys.count
        outputData(1, i + 1) = colKeys(i)
    Next i

    For i = 1 To rowKeys.count
        outputData(i + 1, 1) = rowKeys(i)
    Next i

    For i = 1 To UBound(data, 1)
        rowKey = CStr(data(i, 1))
        colKey = CStr(data(i, 2))
        rowIndex = FindIndex(rowMap, rowKey) + 1
        colIndex = FindIndex(colMap, colKey) + 1
        outputData(rowIndex, colIndex) = data(i, 3)
    Next i

    ActiveSheet.Rows("7:" & ActiveSheet.Rows.count).ClearContents

    Set outputStartCell = startCell.Offset(0, outputOffsetColumns)
    Set outputRange = outputStartCell.Resize(totalRows, lastCol)
    outputRange.value = outputData

    AutoFit_Click
    With outputRange
        .VerticalAlignment = xlCenter
        .Columns.AutoFit
    End With

    ' headers
    With outputRange.Offset(0, 1).Resize(outputRange.Rows.count, outputRange.Columns.count - 1)
        .HorizontalAlignment = xlCenter
    End With
    ws.Rows(8).HorizontalAlignment = xlLeft
    Set outputRange = outputStartCell.Resize(1, lastCol)
    With outputRange
        .Interior.Color = RGB(64, 64, 64)
    End With
    
    ' Save headers to store
    MatrixData.inputData = outputData
    Set MatrixData.InputRange = outputRange
End Sub

' In MatrixBuilder class or a general module
Public Sub RefreshMatrix(ws As Worksheet, startCell As Range)
    Dim lastRow As Long, lastCol As Long
    Dim currentRange As Range
    Dim currentData As Variant

    If MatrixData Is Nothing Then
        Set MatrixData = New MatrixData
    End If
    
    ' Detect the current matrix size starting from startCell
    lastRow = ws.Cells(ws.Rows.count, startCell.Column).End(xlUp).row
    lastCol = ws.Cells(startCell.row, ws.Columns.count).End(xlToLeft).Column

    ' Read the current matrix from the sheet
    Set currentRange = ws.Range(startCell, ws.Cells(lastRow, lastCol))
    currentData = currentRange.value

    ' Update the store with the new values
    MatrixData.inputData = currentData
    Set MatrixData.InputRange = currentRange
End Sub

' === MATRIX UNPIVOT ===
Public Sub Unpivot(ws As Worksheet)
    Dim lastRow As Long, lastCol As Long
    Dim inputData As Variant, outputData() As Variant
    Dim i As Long, j As Long, outRow As Long
    Dim numRows As Long, numCols As Long
    Dim outputStart As Range
'    Dim Headers() As String
    Dim outputRowOffset As Long
    
    If MatrixData Is Nothing Then
'        MsgBox "MatrixData not initialized. Please run BuildMatrix first.", vbExclamation
'        Exit Sub
        Set MatrixData = New MatrixData
    End If

    outputRowOffset = 2
    inputData = MatrixData.inputData
'    Headers = MatrixData.Headers

    numRows = UBound(inputData, 1)
    numCols = UBound(inputData, 2)
    ReDim outputData(1 To (numRows - 1) * (numCols - 1), 1 To 3)

    outRow = 1
    For i = 2 To numRows
        For j = 2 To numCols
            outputData(outRow, 1) = inputData(i, 1)
            outputData(outRow, 2) = inputData(1, j)
            outputData(outRow, 3) = inputData(i, j)
            outRow = outRow + 1
        Next j
    Next i

    ReDim Preserve outputData(1 To outRow - 1, 1 To 3)

    ' Reset formatting from B8 down
    lastRow = ws.Cells(ws.Rows.count, 2).End(xlUp).row
    lastCol = ws.Cells(8, ws.Columns.count).End(xlToLeft).Column
    If lastRow >= 8 And lastCol >= 2 Then
        With ws.Range(ws.Cells(8, 2), ws.Cells(lastRow, lastCol))
            .HorizontalAlignment = xlLeft
            .VerticalAlignment = xlTop
        End With
    End If

    ' Clear previous output starting from row 7
    ws.Rows("7:" & ws.Rows.count).ClearContents

    ' Determine output start cell
    Set outputStart = ws.Range("B7").Offset(outputRowOffset, 0)

    ' Write unpivoted data
    With outputStart.Resize(UBound(outputData, 1), 3)
        .value = outputData
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlTop
        .Columns.AutoFit
    End With

    ' Write headers one row above the output
'    With outputStart.Offset(-1, 0).Resize(1, 3)
'        .value = Array(Headers(1), Headers(2), Headers(3))
'        .HorizontalAlignment = xlLeft
'        .VerticalAlignment = xlTop
'    End With
    ws.Range("B8").value = ws.Range("C6").value
    ws.Range("C8").value = ws.Range("D6").value
    ws.Range("D8").value = ws.Range("E6").value
    ws.Range("C6").value = ""
    ws.Range("D6").value = ""
    ws.Range("E6").value = ""
    
    
End Sub


