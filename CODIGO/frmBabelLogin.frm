VERSION 5.00
Begin VB.Form frmBabelLogin 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3015
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4560
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
    Me.Height = (D3DWindow.BackBufferHeight + 20) * screen.TwipsPerPixelY
    
    UIRenderArea.ScaleMode = vbPixel
    UIRenderArea.Width = D3DWindow.BackBufferWidth * screen.TwipsPerPixelX
    UIRenderArea.Height = D3DWindow.BackBufferHeight * screen.TwipsPerPixelY
    Call InitializeUI(D3DWindow.BackBufferWidth, D3DWindow.BackBufferHeight, 4)
    UIRenderArea.Top = 0 '20 * screen.TwipsPerPixelY
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
    Call BabelSendMouseEvent(x / screen.TwipsPerPixelX, y / screen.TwipsPerPixelY, kType_MouseUp, button)
End Sub
