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
' Common UI color constants (opaque)
Public Const UI_COLOR_WHITE As Long = &HFFFFFF     ' &HFFFFFFFF
Public Const UI_COLOR_BLACK As Long = &H0        ' 0xFF000000
Public Const UI_COLOR_RED   As Long = &HFF0000   ' 0xFFFF0000
Public Const UI_COLOR_GREEN As Long = &HFF00FF   ' 0xFF00FF00
Public Const UI_COLOR_BLUE  As Long = &HFF       ' 0xFF0000FF
Public g_connectScreen      As clsUIConnectScreen
' Virtual-Key codes
Public Const VK_LBUTTON     As Long = &H1

' Mouse position API
Private Type POINTAPI
    x As Long
    y As Long
End Type

Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function ScreenToClient Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long
Public g_MouseX       As Long
Public g_MouseY       As Long
Public g_MouseButtons As Long
Public UIRenderer     As clsUIRenderer

' Update current mouse state relative to the client window
Public Sub UpdateMouse(ByVal hWnd As Long)
    Dim pt As POINTAPI
    If GetCursorPos(pt) Then
        ScreenToClient hWnd, pt
        g_MouseX = pt.x
        g_MouseY = pt.y
    End If
End Sub

Public Sub init_connect_screen(ByRef dev As Direct3DDevice8)
    Set g_connectScreen = New clsUIConnectScreen: g_connectScreen.Init dev
End Sub

Public Sub init_dx_ui(ByRef dev As Direct3DDevice8)
    #If DXUI Then
        Set UIRenderer = New clsUIRenderer
        Call UIRenderer.Init(DirectDevice, 1000)
        Call init_connect_screen(DirectDevice)
    #End If
End Sub
