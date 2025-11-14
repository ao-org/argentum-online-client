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
Public Const UI_COLOR_BLUE   As Long = &HFF0000   ' 0xFFFF0000
Public Const UI_COLOR_GREEN As Long = &HFF00     ' 0xFF00FF00
Public Const UI_COLOR_RED  As Long = &HFF       ' 0xFF0000FF
Public Const UI_COLOR_GOLD  As Long = &HFFFFFF              ' 0xFF0000FF
Public Const UI_COLOR_GRAY  As Long = &HFFFF00              ' 0xFF0000FF
Public Const UI_MAX_QUADS As Long = 2000
Public g_connectScreen      As clsUIConnectScreen
Public g_statisticsScreen   As clsUIStatistics

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
Public UITextures As clsTexManager
    

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

Public Sub init_statistic_screen(ByRef dev As Direct3DDevice8)
    Set g_statisticsScreen = New clsUIStatistics
    Call g_statisticsScreen.Init(dev)
End Sub


Private Sub preload_ui_textures()
#If DXUI Then
    ' Preload all UI textures you will use (ids are examples)
    Dim ids()
    Dim w As Long: Dim h As Long
    ids = Array(14, 21, 13000, 13001, 13002, 13030, 13031, 13032, 13033, 13034, 13035, 13036, 13037, 13038, 13039, 13043)     ' <- put your real atlas/skin ids here

    Dim i As Long, tex As Direct3DTexture8
    For i = LBound(ids) To UBound(ids)
        Set tex = UITextures.GetTexture(ids(i), w, h)  ' fetch or load from manager cache
        Debug.Assert Not tex Is Nothing
        Debug.Print "UI texture bytes: "; CStr(UITextures.GetAllocatedBytes())

    Next i
#End If
End Sub

Public Sub init_dx_ui(ByRef dev As Direct3DDevice8)
    #If DXUI Then
        Set UITextures = New clsTexManager
        Call UITextures.Init(DirectD3D8, dev)
        Set UIRenderer = New clsUIRenderer
        preload_ui_textures
            
        Call UIRenderer.Init(DirectDevice, UI_MAX_QUADS)
        Call init_connect_screen(DirectDevice)
        Call init_statistic_screen(DirectDevice)
        'Agregar aca init statistics ui
    #End If
End Sub

Public Function KeyCodeToChar(ByVal KeyCode As Integer, ByVal Shift As Integer) As String
    ' Handles typical alphanumeric and punctuation keys
    Dim ch As String

    Select Case KeyCode
        Case vbKeyA To vbKeyZ
            If (Shift And vbShiftMask) <> 0 Then
                ch = Chr$(KeyCode)
            Else
                ch = LCase$(Chr$(KeyCode))
            End If

        Case vbKey0 To vbKey9
            If (Shift And vbShiftMask) <> 0 Then
                ' Handle shifted numbers (e.g. !@#$%)
                Select Case KeyCode
                    Case vbKey1: ch = "!"
                    Case vbKey2: ch = "@"
                    Case vbKey3: ch = "#"
                    Case vbKey4: ch = "$"
                    Case vbKey5: ch = "%"
                    Case vbKey6: ch = "^"
                    Case vbKey7: ch = "&"
                    Case vbKey8: ch = "*"
                    Case vbKey9: ch = "("
                    Case vbKey0: ch = ")"
                End Select
            Else
                ch = Chr$(KeyCode)
            End If

        Case vbKeySpace
            ch = " "

        Case vbKeyBack
            ch = "{BACK}"

        Case vbKeyReturn
            ch = "{ENTER}"

        Case vbKeyTab
            ch = "{TAB}"

        Case vbKeyDelete
            ch = "{DEL}"

        Case vbKeyLeft
            ch = "{LEFT}"
        Case vbKeyRight
            ch = "{RIGHT}"

        Case vbKeyUp
            ch = "{UP}"

        Case vbKeyDown
            ch = "{DOWN}"

        Case "190"
            If (Shift And vbShiftMask) <> 0 Then ch = ">" Else ch = "."

'        Case vbKeyComma
'            If (Shift And vbShiftMask) <> 0 Then ch = "<" Else ch = ","
'
'        Case vbKeyMinus
'            If (Shift And vbShiftMask) <> 0 Then ch = "_" Else ch = "-"
'
'        Case vbKeyEqual
'            If (Shift And vbShiftMask) <> 0 Then ch = "+" Else ch = "="

        Case Else
            ' ignore function keys, control, etc.
            ch = vbNullString
    End Select

    KeyCodeToChar = ch
End Function
