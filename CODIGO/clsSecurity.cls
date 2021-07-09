VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSecurity"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
#If AntiExternos = 1 Then

' Módulo "anti-externos", ponele. By WyroX
' La idea es complicársela un poco a un sniffer.

Private XorArray(137) As Byte

Public Sub Initialize()
    Call Rnd(-1)
    Call Randomize(129)

    Dim i As Integer
    For i = 0 To UBound(XorArray)
        XorArray(i) = Rnd * 256
    Next
End Sub

Public Sub XorData(ByRef Data() As Byte, ByVal UpperBound As Long, ByRef Index As Integer)
    Dim i As Long
    For i = 0 To UpperBound
        Data(i) = Data(i) Xor XorArray(Index)
        Index = Index + 7
        If Index > UBound(XorArray) Then Index = Index - UBound(XorArray) - 1
    Next
End Sub

#End If

