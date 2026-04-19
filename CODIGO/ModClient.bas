Attribute VB_Name = "ModClient"
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
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Const GWL_EXSTYLE As Long = (-20)
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Const WS_EX_TRANSPARENT As Long = &H20&
Private Const WM_NCLBUTTONDOWN = &HA1
Private Const HTCAPTION = 2
Private Const WS_EX_LAYERED = &H80000
Private Const LWA_ALPHA = &H2&

Public Sub Make_Transparent_Richtext(ByVal hWnd As Long)
    'If Win2kXP Then
    On Error GoTo Make_Transparent_Richtext_Err
    Call SetWindowLong(hWnd, GWL_EXSTYLE, WS_EX_TRANSPARENT)
    Exit Sub
Make_Transparent_Richtext_Err:
    Call RegistrarError(Err.Number, Err.Description, "ModClient.Make_Transparent_Richtext", Erl)
    Resume Next
End Sub

Public Sub NameMapa(ByVal map As Long)
    On Error GoTo NameMapa_Err
    frmMain.NameMapa.Caption = MapDat.map_name
    If QueRender = 0 Then
        Letter_Set 0, MapDat.map_name
    End If
    Exit Sub
NameMapa_Err:
    Call RegistrarError(Err.Number, Err.Description, "ModClient.NameMapa", Erl)
    Resume Next
End Sub

