VERSION 5.00
Begin VB.Form frmDebugUI 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Inspector"
   ClientHeight    =   11520
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   15360
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   768
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1024
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.PictureBox RenderArea 
      BackColor       =   &H00000000&
      Height          =   11520
      Left            =   0
      ScaleHeight     =   768
      ScaleMode       =   0  'User
      ScaleWidth      =   1024
      TabIndex        =   0
      Top             =   0
      Width           =   15360
   End
End
Attribute VB_Name = "frmDebugUI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim FrmMove As Boolean
Dim DragX, Dragy As Integer
Dim Pressing As Boolean

Private Sub Form_Load()
    
    On Error GoTo Form_Load_Err
    
    ' Seteamos el caption hay que poner 20 aniversario
    Me.Caption = "Inspector"
    
    ' Removemos la barra de titulo pero conservando el caption para la barra de tareas
    Call Form_RemoveTitleBar(Me)
    Debug.Assert D3DWindow.BackBufferWidth <> 0
    Debug.Assert D3DWindow.BackBufferHeight <> 0
    Me.ScaleMode = vbPixel
    Me.Width = D3DWindow.BackBufferWidth * screen.TwipsPerPixelX
    Me.Height = (D3DWindow.BackBufferHeight + 20) * screen.TwipsPerPixelY
    RenderArea.Top = 20
    Call InitializeInspectorUI(D3DWindow.BackBufferWidth, D3DWindow.BackBufferHeight)
    Exit Sub

Form_Load_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmDebugUi.Form_Load", Erl)
    Resume Next
    
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    FrmMove = True
    DragX = x * screen.TwipsPerPixelX
    Dragy = y * screen.TwipsPerPixelY
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim nX, nY As Long
    If FrmMove Then
        nX = Me.Left + x * screen.TwipsPerPixelX - DragX
        nY = Me.Top + y * screen.TwipsPerPixelY - Dragy
        Me.Left = nX
        Me.Top = nY
    End If
End Sub
Private Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim nX, nY As Long
    nX = Me.Left + x * screen.TwipsPerPixelX - DragX
    nY = Me.Top + y * screen.TwipsPerPixelY - Dragy
    Me.Left = nX
    Me.Top = nY
    FrmMove = False
End Sub

Private Sub RenderArea_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo Form_KeyDown_Err
    Dim CapsState As Boolean
    CapsState = GetKeyState(vbKeyCapital)
    Call BabelSendKeyEvent(KeyCode, Shift, kType_RawKeyDown, CapsState, True)
    Exit Sub

Form_KeyDown_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmDebugUI.Form_KeyDown", Erl)
End Sub

Private Sub RenderArea_KeyPress(KeyAscii As Integer)
On Error GoTo RenderArea_KeyPress_Err
    Dim CapsState As Boolean
    CapsState = GetKeyState(vbKeyCapital)
    Call BabelSendKeyEvent(KeyAscii, Shift, kType_Char, CapsState, True)
    Exit Sub
RenderArea_KeyPress_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmDebugUI.RenderArea_KeyPress", Erl)
End Sub

Private Sub RenderArea_KeyUp(KeyCode As Integer, Shift As Integer)
On Error GoTo Form_KeyUp_Err
    Dim CapsState As Boolean
    CapsState = GetKeyState(vbKeyCapital)
    Call BabelSendKeyEvent(KeyCode, Shift, kType_KeyUp, CapsState, True)
    Exit Sub

Form_KeyUp_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmDebugUI.Form_KeyUp", Erl)
End Sub

Private Sub RenderArea_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbLeftButton Then
        Call SendDebugMouseEvent(x, y, kType_MouseDown, kButton_Left)
    ElseIf Button = vbRightButton Then
        Call SendDebugMouseEvent(x, y, kType_MouseDown, kButton_Right)
    ElseIf Button = vbMiddleButton Then
        Call SendDebugMouseEvent(x, y, kType_MouseDown, kButton_Middle)
    End If
End Sub

Private Sub RenderArea_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call SendDebugMouseEvent(x, y, kType_MouseMove, kButton_None)
End Sub

Private Sub RenderArea_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbLeftButton Then
        Call SendDebugMouseEvent(x, y, kType_MouseUp, kButton_Left)
    ElseIf Button = vbRightButton Then
        Call SendDebugMouseEvent(x, y, kType_MouseUp, kButton_Right)
    ElseIf Button = vbMiddleButton Then
        Call SendDebugMouseEvent(x, y, kType_MouseUp, kButton_Middle)
    End If
End Sub
