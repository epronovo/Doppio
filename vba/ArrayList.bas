VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ArrayList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private values() As Variant

Public Sub Initialize()
    ReDim values(0)
End Sub

Public Sub Add(item As Variant)
    Dim index As Long
    index = UBound(values) + 1
    ReDim Preserve values(0 To index)
    values(index) = item
End Sub

Public Sub Sort()
    Dim i As Long, j As Long
    Dim temp As Variant
    For i = LBound(values) To UBound(values) - 1
        For j = i + 1 To UBound(values)
            If values(i) > values(j) Then
                temp = values(i)
                values(i) = values(j)
                values(j) = temp
            End If
        Next j
    Next i
End Sub

Public Function item(index As Long) As Variant
    If index >= 1 And index <= UBound(values) Then
        item = values(index)
    End If
End Function

Public Function count() As Long
    On Error GoTo ErrorHandler
    count = UBound(values)
    Exit Function
ErrorHandler:
    count = -1
End Function

Public Function Contains(item As Variant) As Boolean
    Dim i As Long
    For i = LBound(values) To UBound(values)
        If values(i) = item Then
            Contains = True
            Exit Function
        End If
    Next i
    Contains = False
End Function

Public Function IndexOf(item As Variant) As Long
    Dim i As Long
    For i = LBound(values) To UBound(values)
        If values(i) = item Then
            IndexOf = i
            Exit Function
        End If
    Next i
    ' If the item is not found, return -1
    IndexOf = -1
End Function

