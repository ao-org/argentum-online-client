VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsNetWriter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'    Argentum 20 - Game Client Program
'    Copyright (C) 2025 - Noland Studios
'
'    This program is free software: you can redistribute it and/or modify
'    it under the terms of the GNU Affero General Public License as published by
'    the Free Software Foundation, either version 3 of the License, or
'    (at your option) any later version.
'
'    This program is distributed in the hope that it will be useful,
'    but WITHOUT ANY WARRANTY; without even the implied warranty of
'    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'    GNU Affero General Public License for more details.
'    You should have received a copy of the GNU Affero General Public License
'    along with this program.  If not, see <https://www.gnu.org/licenses/>.
#If DIRECT_PLAY = 1 Then
    Option Base 0
    Option Explicit
    Private oMsg()  As Byte
    Private loffset As Long


Public Sub Clear()
    loffset = NewBuffer(oMsg)
End Sub
Public Sub WriteInt16(ByVal v As Integer)
    AddDataToBuffer oMsg, v, SIZE_INTEGER, loffset
End Sub

Public Sub WriteInt8(ByVal v As Byte)
    AddDataToBuffer oMsg, v, SIZE_BYTE, loffset
End Sub
Public Sub send()
On Error GoTo send_error:
    Err.Clear
    ReDim Preserve oMsg(0 To loffset - 1)
#If Developer = 1 Then
    Call SaveStringInFile("Sending:" & ByteArrayToDecimalString(oMsg), "send_debug.txt")
#End If
    dpc.send oMsg, 0, DPNSEND_GUARANTEED Or DPNSEND_NOLOOPBACK
    Me.Clear
    Exit Sub
send_error:
    Me.Clear
    If Err.Number <> 0 Then
        Call HandleDPlayError(Err.Number, Err.Description, "modnetwork.send", Erl)
    End If
End Sub
Public Sub WriteInt32(ByVal v As Long)
    AddDataToBuffer oMsg, v, SIZE_LONG, loffset
End Sub
Public Sub WriteReal32(ByVal v As Single)
    AddDataToBuffer oMsg, v, SIZE_SINGLE, loffset
End Sub
Public Sub WriteBool(ByVal v As Boolean)
    AddDataToBuffer oMsg, v, SIZE_BOOLEAN, loffset
End Sub
Public Sub WriteString8(ByVal v As String)
    Debug.Assert Len(v) > 0
        If Len(v) = 0 Then
        ' this is a workaround because the game is trying to send empty strings and DP will crash if we call AddStringToBuffer with an empty str
        v = "0"
    End If
    AddStringToBuffer oMsg, v, loffset
End Sub

Public Sub WriteSafeArrayInt8(value() As Byte)
    Debug.Assert (UBound(value) - LBound(value)) + 1 > 0
    WriteInt32 LBound(value)
    WriteInt32 UBound(value)
    Dim i As Integer
    For i = LBound(value) To UBound(value)
        WriteInt8 value(i)
    Next i
End Sub
Private Sub Class_Initialize()
 loffset = NewBuffer(oMsg)
End Sub

#End If

