Attribute VB_Name = "MD5"
'    Argentum 20 - Game Client Program
'    Copyright (C) 2022 - Noland Studios
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
'
'
Option Explicit

Private Declare Sub MDFile Lib "aamd532.dll" (ByVal f As String, ByVal r As String)
Private Declare Sub MDStringFix Lib "aamd532.dll" (ByVal f As String, ByVal t As Long, ByVal r As String)

Public Function MD5String(ByVal P As String) As String
    On Error Goto MD5String_Err
    Dim r As String * 32, t As Long
    r = Space(32)
    t = Len(P)
    MDStringFix P, t, r
    MD5String = r
    Exit Function
MD5String_Err:
    Call TraceError(Err.Number, Err.Description, "MD5.MD5String", Erl)
End Function

Public Function MD5File(ByVal f As String) As String
    On Error Goto MD5File_Err
    Dim r As String * 32
    r = Space(32)
    MDFile f, r
    MD5File = r
    Exit Function
MD5File_Err:
    Call TraceError(Err.Number, Err.Description, "MD5.MD5File", Erl)
End Function

Public Function hexMd52Asc(ByVal MD5 As String) As String
    On Error Goto hexMd52Asc_Err
    Dim i As Long
    Dim L As String
    
    If Len(MD5) And &H1 Then MD5 = "0" & MD5
    
    For i = 1 To Len(MD5) \ 2
        L = mid$(MD5, (2 * i) - 1, 2)
        hexMd52Asc = hexMd52Asc & Chr$(hexHex2Dec(L))
    Next i
    Exit Function
hexMd52Asc_Err:
    Call TraceError(Err.Number, Err.Description, "MD5.hexMd52Asc", Erl)
End Function

Public Function hexHex2Dec(ByVal hex As String) As Long
    On Error Goto hexHex2Dec_Err
    hexHex2Dec = Val("&H" & hex)
    Exit Function
hexHex2Dec_Err:
    Call TraceError(Err.Number, Err.Description, "MD5.hexHex2Dec", Erl)
End Function

Public Function txtOffset(ByVal Text As String, ByVal off As Integer) As String
    On Error Goto txtOffset_Err
    Dim i As Long
    Dim L As String
    
    For i = 1 To Len(Text)
        L = mid$(Text, i, 1)
        txtOffset = txtOffset & Chr$((Asc(L) + off) And &HFF)
    Next i
    Exit Function
txtOffset_Err:
    Call TraceError(Err.Number, Err.Description, "MD5.txtOffset", Erl)
End Function


