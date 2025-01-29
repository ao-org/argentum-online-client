VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsNetReader"
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

Private rcv_data As DxVBLibA.DPNMSG_RECEIVE
Private loffset As Long
Public Function GetAvailable() As Integer
    GetAvailable = rcv_data.lDataSize - loffset
End Function
Public Sub set_data(data As DxVBLibA.DPNMSG_RECEIVE)
    rcv_data = data
    loffset = 0
End Sub

Public Sub Clear()
    loffset = NewBuffer(oMsg)
End Sub
Public Function ReadInt16() As Integer
    Call GetDataFromBuffer(rcv_data.ReceivedData, ReadInt16, SIZE_INTEGER, loffset)
End Function
Public Function ReadInt32() As Long
    Call GetDataFromBuffer(rcv_data.ReceivedData, ReadInt32, SIZE_LONG, loffset)
End Function
Public Function ReadBool() As Boolean
    Call GetDataFromBuffer(rcv_data.ReceivedData, ReadBool, SIZE_BOOLEAN, loffset)
End Function
Public Function ReadReal32() As Single
    Call GetDataFromBuffer(rcv_data.ReceivedData, ReadReal32, SIZE_SINGLE, loffset)
End Function
Public Function ReadInt8() As Byte
    Call GetDataFromBuffer(rcv_data.ReceivedData, ReadInt8, SIZE_BYTE, loffset)
End Function
Public Function ReadString8() As String
    ReadString8 = GetStringFromBuffer(rcv_data.ReceivedData, loffset)
End Function

#End If
