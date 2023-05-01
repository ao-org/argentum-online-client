VERSION 5.00
Begin VB.Form frmBabelLogin 
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
      Height          =   3015
      Left            =   0
      ScaleHeight     =   2955
      ScaleWidth      =   4395
      TabIndex        =   0
      Top             =   0
      Width           =   4455
   End
End
Attribute VB_Name = "frmBabelLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private LastMouseX As Single
Private LastMouseY As Single

Private Sub Form_Load()
On Error GoTo Form_Load_Err
    
    ' Seteamos el caption hay que poner 20 aniversario
    Me.Caption = "Login"
    
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
    Call InitializeUI(D3DWindow.BackBufferWidth, D3DWindow.BackBufferHeight, BytesPerPixel)
    UIRenderArea.Top = 0 '20 * screen.TwipsPerPixelY 'keep for debug
    Exit Sub

Form_Load_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmBabelLogin.Form_Load", Erl)
    Resume Next
End Sub

Private Sub Form_MouseDown(button As Integer, Shift As Integer, x As Single, y As Single)
    FrmMove = True
    DragX = x
    Dragy = y
End Sub
Private Sub Form_MouseMove(button As Integer, Shift As Integer, x As Single, y As Single)
Dim nX, nY
    If FrmMove Then
        nX = Me.Left + x - DragX
        nY = Me.Top + y - Dragy
        Me.Left = nX
        Me.Top = nY
    End If
End Sub
Private Sub Form_MouseUp(button As Integer, Shift As Integer, x As Single, y As Single)
    Dim nX, nY
    nX = Me.Left + x - DragX
    nY = Me.Top + y - Dragy
    Me.Left = nX
    Me.Top = nY
    FrmMove = False
    Debug.Print "form mouse up"
End Sub

Private Sub UIRenderArea_DblClick()
On Error GoTo UIRenderArea_DblClick_Err
    'nasty hack to solve vb6 events issues, vb6 moyse events goes in this way:
    'MouseDown, MouseUp, Click, DblClick, and MouseUp
    'on double click events we miss the second mouse down, because js dont get the full mouse down + up it doesn handle the double click itself
    Dim btnConvert As MouseButton
    btnConvert = ConvertMouseButton(button)
    Call BabelSendMouseEvent(LastMouseX, LastMouseY, kType_MouseDown, kButton_Left)
    Exit Sub
UIRenderArea_DblClick_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmBabelLogin.Form_KeyDown", Erl)
End Sub

Private Sub UIRenderArea_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo Form_KeyDown_Err
    Dim CapsState As Boolean
    CapsState = GetKeyState(vbKeyCapital)
    Call BabelSendKeyEvent(KeyCode, Shift, kType_RawKeyDown, CapsState, False)
    Exit Sub
Form_KeyDown_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmBabelLogin.Form_KeyDown", Erl)
End Sub

Private Sub UIRenderArea_KeyPress(KeyAscii As Integer)
On Error GoTo RenderArea_KeyPress_Err
    Dim CapsState As Boolean
    CapsState = GetKeyState(vbKeyCapital)
    Call BabelSendKeyEvent(KeyAscii, Shift, kType_Char, CapsState, False)
    Exit Sub
RenderArea_KeyPress_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmBabelLogin.RenderArea_KeyPress", Erl)
End Sub

Private Sub UIRenderArea_KeyUp(KeyCode As Integer, Shift As Integer)
On Error GoTo Form_KeyUp_Err
    Dim CapsState As Boolean
    CapsState = GetKeyState(vbKeyCapital)
    Call BabelSendKeyEvent(KeyCode, Shift, kType_KeyUp, CapsState, False)
    If Not DebugInitialized Then
        If Shift And KeyCode = 68 Then 'shift + d
            frmDebugUI.Show
        End If
    End If
    Exit Sub
Form_KeyUp_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmBabelLogin.Form_KeyUp", Erl)
End Sub

Private Sub UIRenderArea_MouseDown(button As Integer, Shift As Integer, x As Single, y As Single)
    Dim btnConvert As MouseButton
    btnConvert = ConvertMouseButton(button)
    Call BabelSendMouseEvent(x / screen.TwipsPerPixelX, y / screen.TwipsPerPixelY, kType_MouseDown, button)
End Sub

Private Sub UIRenderArea_MouseMove(button As Integer, Shift As Integer, x As Single, y As Single)
    Call BabelSendMouseEvent(x / screen.TwipsPerPixelX, y / screen.TwipsPerPixelY, kType_MouseMove, kButton_None)
End Sub

Private Sub UIRenderArea_MouseUp(button As Integer, Shift As Integer, x As Single, y As Single)
    Dim btnConvert As MouseButton
    btnConvert = ConvertMouseButton(button)
    LastMouseX = x / screen.TwipsPerPixelX
    LastMouseY = y / screen.TwipsPerPixelY
    Call BabelSendMouseEvent(LastMouseX, LastMouseY, kType_MouseUp, button)
End Sub
