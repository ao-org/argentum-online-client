Attribute VB_Name = "dx_ui"
' Argentum 20 Game Client
'
'    Copyright (C) 2025 Noland Studios LTD
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
'
'    You should have received a copy of the GNU Affero General Public License
'    along with this program.  If not, see <https://www.gnu.org/licenses/>.
'
'    This program was based on Argentum Online 0.11.6
'    Copyright (C) 2002 Márquez Pablo Ignacio
'
'    Argentum Online is based on Baronsoft's VB6 Online RPG
'    You can contact the original creator of ORE at aaron@baronsoft.com
'    for more information about ORE please visit http://www.baronsoft.com/
'
'
'
Option Explicit


Private Declare Function CreateFontA Lib "gdi32" _
    (ByVal nHeight As Long, ByVal nWidth As Long, ByVal nEscapement As Long, ByVal nOrientation As Long, _
     ByVal fnWeight As Long, ByVal fdwItalic As Long, ByVal fdwUnderline As Long, ByVal fdwStrikeOut As Long, _
     ByVal fdwCharSet As Long, ByVal fdwOutputPrecision As Long, ByVal fdwClipPrecision As Long, _
     ByVal fdwQuality As Long, ByVal fdwPitchAndFamily As Long, ByVal lpszFace As String) As Long

' Font constants
Private Const FW_NORMAL As Long = 400
Private Const DEFAULT_CHARSET As Long = 1
Private Const OUT_DEFAULT_PRECIS As Long = 0
Private Const DEFAULT_QUALITY As Long = 0
Private Const DEFAULT_PITCH As Long = 0
Private Const FF_DONTCARE As Long = 0

Public g_connectScreen As clsUIConnectScreen
Public g_Font As D3DXFont
' Virtual-Key codes
Public Const VK_LBUTTON As Long = &H1
' Mouse position API
Private Type POINTAPI
    x As Long
    y As Long
End Type
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function ScreenToClient Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long
Public g_MouseX As Long
Public g_MouseY As Long
Public g_MouseButtons As Long
Public UIRenderer As clsUIRenderer


' Update current mouse state relative to the client window
Public Sub UpdateMouse(ByVal hWnd As Long)
    On Error Goto UpdateMouse_Err
    Dim pt As POINTAPI
    If GetCursorPos(pt) Then
        ScreenToClient hWnd, pt
        g_MouseX = pt.x
        g_MouseY = pt.y
    End If
    Exit Sub
UpdateMouse_Err:
    Call TraceError(Err.Number, Err.Description, "dx_ui.UpdateMouse", Erl)
End Sub

Public Sub init_connect_screen(ByRef dev As Direct3DDevice8)
    On Error Goto init_connect_screen_Err
    If g_Font Is Nothing Then
        Dim hFont As Long
        hFont = CreateFontA(20, 0, 0, 0, FW_NORMAL, 0, 0, 0, DEFAULT_CHARSET, OUT_DEFAULT_PRECIS, OUT_DEFAULT_PRECIS, DEFAULT_QUALITY, DEFAULT_PITCH Or FF_DONTCARE, "Arial")
        Set g_Font = DirectD3D8.CreateFont(dev, hFont)
    End If
    Set g_connectScreen = New clsUIConnectScreen: g_connectScreen.Init dev, g_Font
    Exit Sub
init_connect_screen_Err:
    Call TraceError(Err.Number, Err.Description, "dx_ui.init_connect_screen", Erl)
End Sub




Public Sub init_dx_ui(ByRef dev As Direct3DDevice8)
    On Error Goto init_dx_ui_Err
#If DXUI Then
    Set UIRenderer = New clsUIRenderer
    Call UIRenderer.Init(DirectDevice, 1000)
    Call init_connect_screen(DirectDevice)
#End If
    Exit Sub
init_dx_ui_Err:
    Call TraceError(Err.Number, Err.Description, "dx_ui.init_dx_ui", Erl)
End Sub
