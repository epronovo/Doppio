VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MatrixData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Class name: MatrixDataStore
Option Explicit

Private pInputRange As Range
Private pInputData As Variant
Private pHeaders() As String

Public Property Let headers(value() As String)
    pHeaders = value
End Property

Public Property Get headers() As String()
    headers = pHeaders
End Property

Public Sub ZZZ_SetHeaders(value() As String)
    pHeaders = value
End Sub

Public Property Set InputRange(rng As Range)
    Set pInputRange = rng
End Property

Public Property Get InputRange() As Range
    Set InputRange = pInputRange
End Property

Public Property Let inputData(val As Variant)
    pInputData = val
End Property

Public Property Get inputData() As Variant
    inputData = pInputData
End Property
