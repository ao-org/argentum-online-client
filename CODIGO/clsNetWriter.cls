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

Private oMsg() As Byte
Private lOffset As Long

Public Sub Clear()
    lOffset = NewBuffer(oMsg)
End Sub
Public Sub WriteInt16(ByVal v As Integer)
    AddDataToBuffer oMsg, v, SIZE_INTEGER, lOffset
End Sub

Public Sub WriteInt8(ByVal v As Byte)
    AddDataToBuffer oMsg, v, SIZE_BYTE, lOffset
End Sub
Public Sub send()
    dpc.send oMsg, 0, DPNSEND_NOLOOPBACK & DPNSEND_GUARANTEED
    Me.Clear
End Sub
Public Sub WriteString8(ByVal v As String)
    AddStringToBuffer oMsg, v, lOffset
End Sub

Private Sub Class_Initialize()
 lOffset = NewBuffer(oMsg)
End Sub

#End If

