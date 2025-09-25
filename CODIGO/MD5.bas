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
Private Declare Sub MDFile Lib "aamd532.dll" (ByVal f As String, ByVal R As String)
Private Declare Sub MDStringFix Lib "aamd532.dll" (ByVal f As String, ByVal t As Long, ByVal R As String)

Public Function MD5String(ByVal p As String) As String
    Dim R As String * 32, t As Long
    R = Space(32)
    t = Len(p)
    MDStringFix p, t, R
    MD5String = R
End Function

Public Function MD5File(ByVal f As String) As String
    Dim R As String * 32
    R = Space(32)
    MDFile f, R
    MD5File = R
End Function

Public Function hexMd52Asc(ByVal MD5 As String) As String
    Dim i As Long
    Dim l As String
    If Len(MD5) And &H1 Then MD5 = "0" & MD5
    For i = 1 To Len(MD5) \ 2
        l = mid$(MD5, (2 * i) - 1, 2)
        hexMd52Asc = hexMd52Asc & Chr$(hexHex2Dec(l))
    Next i
End Function

Public Function hexHex2Dec(ByVal hex As String) As Long
    hexHex2Dec = val("&H" & hex)
End Function

Public Function txtOffset(ByVal text As String, ByVal off As Integer) As String
    Dim i As Long
    Dim l As String
    For i = 1 To Len(text)
        l = mid$(text, i, 1)
        txtOffset = txtOffset & Chr$((Asc(l) + off) And &HFF)
    Next i
End Function
