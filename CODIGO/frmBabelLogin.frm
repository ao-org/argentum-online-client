VERSION 5.00
Begin VB.Form frmBabelUI 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "Argentum20"
   ClientHeight    =   3015
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4560
   Icon            =   "frmBabelLogin.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox UIRenderArea 
      BackColor       =   &H80000007&
      BorderStyle     =   0  'None
      Height          =   3015
      Left            =   0
      ScaleHeight     =   3015
      ScaleWidth      =   4455
      TabIndex        =   0
      Top             =   0
      Width           =   4455
   End
End
Attribute VB_Name = "frmBabelUI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Argentum 20 Game Client
'
'    Copyright (C) 2023 Noland Studios LTD
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

Private Type t_InputState
    LastMouseX As Single
    LastMouseY As Single
    ShiftState As Integer
    LastButton As Integer
    'this is not the same, as LastButton, a mouse mouve with no button press
    'will set LastButton to 0 but we need to keep track o the last pressed.
    LastButtonPress As Integer
End Type
Private InputState As t_InputState
Dim FrmMove As Boolean
Dim DragX, DragY As Integer

Private Sub Form_Load()
On Error GoTo Form_Load_Err
    
    ' Seteamos el caption hay que poner 20 aniversario
    Me.Caption = "Argentum20"
    
    ' Removemos la barra de titulo pero conservando el caption para la barra de tareas
    Call Form_RemoveTitleBar(Me)
    Debug.Assert D3DWindow.BackBufferWidth <> 0
    Debug.Assert D3DWindow.BackBufferHeight <> 0
    Me.ScaleMode = vbPixel
    Me.Width = D3DWindow.BackBufferWidth * screen.TwipsPerPixelX
    Me.Height = (D3DWindow.BackBufferHeight) * screen.TwipsPerPixelY
    
    UIRenderArea.ScaleMode = vbPixel
    UIRenderArea.Width = D3DWindow.BackBufferWidth * screen.TwipsPerPixelX
    UIRenderArea.Height = D3DWindow.BackBufferHeight * screen.TwipsPerPixelY
    UIRenderArea.Top = 0 '20 * screen.TwipsPerPixelY 'keep for debug
    Call FormParser.Parse_Form(Me)
    Call Subclass(Me.hwnd, AddressOf BabelEditWndProc)
    Exit Sub

Form_Load_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmBabelUI.Form_Load", Erl)
    Resume Next
End Sub

Private Sub UIRenderArea_Click()
    If g_game_state.state = e_state_gameplay_screen Then
        If IsInsideGameplayArea(InputState.LastMouseX, InputState.LastMouseY) Then
            Call OnClick(InputState.LastButton, InputState.ShiftState)
        End If
    End If
End Sub

Private Sub UIRenderArea_DblClick()
On Error GoTo UIRenderArea_DblClick_Err
    'nasty hack to solve vb6 events issues, vb6 mouse events goes in this way:
    'MouseDown, MouseUp, Click, DblClick, and MouseUp
    'on double click events we miss the second mouse down, because js dont get the full mouse down + up it doesn handle the double click itself
    Dim btnConvert As MouseButton
    btnConvert = ConvertMouseButton(button)
    Call BabelSendMouseEvent(InputState.LastMouseX, InputState.LastMouseY, kType_MouseDown, kButton_Left)
    If g_game_state.state = e_state_gameplay_screen Then
        If IsInsideGameplayArea(InputState.LastMouseX, InputState.LastMouseY) Then
            If Not frmComerciar.visible And Not frmBancoObj.visible Then
                If InputState.LastButtonPress = vbLeftButton Then
                    Call WriteDoubleClick(tX, tY)
                End If
            End If
        End If
    End If
    Exit Sub
UIRenderArea_DblClick_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmBabelUI.Form_KeyDown", Erl)
End Sub

Private Sub UIRenderArea_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo Form_KeyDown_Err
    Dim CapsState As Boolean
    CapsState = GetKeyState(vbKeyCapital)
    Call BabelSendKeyEvent(KeyCode, Shift, kType_RawKeyDown, CapsState, False)
    Exit Sub
Form_KeyDown_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmBabelUI.Form_KeyDown", Erl)
End Sub

Private Sub UIRenderArea_KeyPress(KeyAscii As Integer)
On Error GoTo RenderArea_KeyPress_Err
    Dim CapsState As Boolean
    CapsState = GetKeyState(vbKeyCapital)
    Call BabelSendKeyEvent(KeyAscii, Shift, kType_Char, CapsState, False)
    Exit Sub
RenderArea_KeyPress_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmBabelUI.RenderArea_KeyPress", Erl)
End Sub

Private Sub UIRenderArea_KeyUp(KeyCode As Integer, Shift As Integer)
On Error GoTo Form_KeyUp_Err
    Dim CapsState As Boolean
    CapsState = GetKeyState(vbKeyCapital)
    Call BabelSendKeyEvent(KeyCode, Shift, kType_KeyUp, CapsState, False)
    If KeyCode = vbKeyV And Shift = vbCtrlMask Then
        ' Retrieve clipboard content
        Dim clipboardText As String
        clipboardText = GetClipboardText()
        ' Prevent further processing of the Ctrl+V keystroke
        Call PasteText(clipboardText)
        KeyCode = 0
    End If
#If DEBUGGING = 1 Or Developer = 1 Then
    If Not DebugInitialized Then
        If (Shift = (vbCtrlMask Or vbShiftMask)) And KeyCode = 68 Then 'shift + d
            frmDebugUI.Show
        End If
    End If
#End If
    If g_game_state.state = e_state_gameplay_screen Then
        Call HandleKeyUp(KeyCode, Shift)
    End If
    Exit Sub
Form_KeyUp_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmBabelUI.Form_KeyUp", Erl)
End Sub

Private Sub UIRenderArea_MouseDown(button As Integer, Shift As Integer, x As Single, y As Single)
    Dim btnConvert As MouseButton
    btnConvert = ConvertMouseButton(button)
    If InputState.LastMouseY < 30 Then
        FrmMove = True
        DragX = x
        DragY = y
    End If
    Call UpdateMouseState(x / screen.TwipsPerPixelX, y / screen.TwipsPerPixelY, button, Shift)
    Call BabelSendMouseEvent(x / screen.TwipsPerPixelX, y / screen.TwipsPerPixelY, kType_MouseDown, button)
End Sub

Private Sub UIRenderArea_MouseMove(button As Integer, Shift As Integer, x As Single, y As Single)
    If FrmMove Then
        Dim nX As Long
        Dim nY As Long
        nX = Me.Left + x - DragX
        nY = Me.Top + y - DragY
        Me.Left = nX
        Me.Top = nY
    End If
    Call UpdateMouseState(x / screen.TwipsPerPixelX, y / screen.TwipsPerPixelY, button, Shift)
    Call BabelSendMouseEvent(InputState.LastMouseX, InputState.LastMouseY, kType_MouseMove, kButton_None)
    
    
End Sub

Private Sub UIRenderArea_MouseUp(button As Integer, Shift As Integer, x As Single, y As Single)
    If FrmMove Then
        Dim nX As Long
        Dim nY As Long
        nX = Me.Left + x - DragX
        nY = Me.Top + y - DragY
        Me.Left = nX
        Me.Top = nY
        FrmMove = False
    End If
    Dim btnConvert As MouseButton
    btnConvert = ConvertMouseButton(button)
    Call UpdateMouseState(x / screen.TwipsPerPixelX, y / screen.TwipsPerPixelY, button, Shift)
    Call BabelSendMouseEvent(InputState.LastMouseX, InputState.LastMouseY, kType_MouseUp, button)
    
    If g_game_state.state = e_state_gameplay_screen Then
        If IsInsideGameplayArea(InputState.LastMouseX, InputState.LastMouseY) Then
            Call HandleGameplayAreaMouseUp(button, InputState.LastMouseX - GameplayDrawAreaRect.Left, InputState.LastMouseY - GameplayDrawAreaRect.Top, Me.Top, Me.Left, Me.Height, GameplayDrawAreaRect)
        End If
    End If
End Sub


Private Sub UpdateMouseState(ByVal PosX As Integer, ByVal PosY As Integer, ByVal LastButton As Integer, ByVal Shift As Integer)
    InputState.LastMouseX = PosX
    InputState.LastMouseY = PosY
    InputState.LastButton = LastButton
    If LastButton > 0 Then InputState.LastButtonPress = LastButton
    InputState.ShiftState = Shift
    MouseX = InputState.LastMouseX - gameplay_render_offset.x
    MouseY = InputState.LastMouseY - gameplay_render_offset.y
End Sub

Private Function IsInsideGameplayArea(ByVal MouseX As Integer, ByVal MouseY As Integer) As Boolean
    If MouseX < gameplay_render_offset.x Then Exit Function 'Render_Main_Rect
    If MouseX > gameplay_render_offset.x + Render_Main_Rect.Right Then Exit Function
    If MouseY < gameplay_render_offset.y Then Exit Function
    If MouseY > gameplay_render_offset.y + Render_Main_Rect.Bottom Then Exit Function
    IsInsideGameplayArea = True
End Function

Private Function GetClipboardText() As String
    Dim clipboardData As Long
    Dim clipboardText As String
    Dim pointer As Long
    Dim length As Long
    ' Check if text format is available in clipboard
    If IsClipboardFormatAvailable(CF_TEXT) Then
        ' Open the clipboard
        If OpenClipboard(Me.hwnd) Then
            ' Get the clipboard data
            clipboardData = GetClipboardData(CF_TEXT)
            ' Lock the memory to get a pointer to the data
            pointer = GlobalLock(clipboardData)
            ' Calculate the length of the text
            length = lstrlenA(pointer)
            ' Allocate space for the text
            clipboardText = Space$(length)
            ' Copy the text from the clipboard to the string variable
            CopyMemory ByVal clipboardText, ByVal pointer, length
            ' Unlock the memory
            GlobalUnlock pointer
            ' Close the clipboard
            CloseClipboard
        End If
    ElseIf IsClipboardFormatAvailable(CF_UNICODETEXT) Then
        ' Open the clipboard
        If OpenClipboard(Me.hwnd) Then
            ' Get the clipboard data
            clipboardData = GetClipboardData(CF_UNICODETEXT)
            ' Lock the memory to get a pointer to the data
            pointer = GlobalLock(clipboardData)
            ' Calculate the length of the Unicode text
            length = lstrlenA(pointer) * 2 ' Length in bytes
            ' Allocate space for the Unicode text
            clipboardText = Space$(length \ 2)
            ' Copy the Unicode text from the clipboard to the string variable
            CopyMemory ByVal StrPtr(clipboardText), ByVal pointer, length
            ' Unlock the memory
            GlobalUnlock pointer
            ' Close the clipboard
            CloseClipboard
        End If
    End If
    ' Return the clipboard text
    GetClipboardText = clipboardText
End Function
