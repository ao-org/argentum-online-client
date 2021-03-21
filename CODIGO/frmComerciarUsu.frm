VERSION 5.00
Begin VB.Form frmComerciarUsu 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   5595
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6510
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   373
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   434
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   465
      Left            =   3960
      ScaleHeight     =   32
      ScaleMode       =   0  'User
      ScaleWidth      =   32
      TabIndex        =   5
      Top             =   1740
      Width           =   480
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   1470
      Left            =   3840
      TabIndex        =   4
      Top             =   2700
      Width           =   2250
   End
   Begin VB.TextBox txtCant 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Left            =   3870
      TabIndex        =   3
      Text            =   "1"
      Top             =   4500
      Width           =   555
   End
   Begin VB.ListBox List2 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   2550
      Left            =   690
      TabIndex        =   1
      Top             =   1590
      Width           =   2430
   End
   Begin VB.Image Image2 
      Height          =   270
      Left            =   4860
      Top             =   1650
      Width           =   270
   End
   Begin VB.Image Image1 
      Height          =   270
      Left            =   4845
      Top             =   1980
      Width           =   285
   End
   Begin VB.Image cmdOfrecer 
      Height          =   405
      Left            =   4740
      Tag             =   "0"
      Top             =   4425
      Width           =   1380
   End
   Begin VB.Image cmdRechazar 
      Height          =   480
      Left            =   495
      Tag             =   "0"
      Top             =   4710
      Width           =   1440
   End
   Begin VB.Image cmdAceptar 
      Height          =   495
      Left            =   1935
      Tag             =   "0"
      Top             =   4710
      Width           =   1335
   End
   Begin VB.Image Command2 
      Height          =   525
      Left            =   3960
      Tag             =   "0"
      Top             =   4980
      Width           =   2130
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Left            =   780
      TabIndex        =   2
      Top             =   4200
      Width           =   2295
   End
   Begin VB.Label lblEstadoResp 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Esperando respuesta..."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   255
      Left            =   720
      TabIndex        =   0
      Top             =   4440
      Visible         =   0   'False
      Width           =   2490
   End
End
Attribute VB_Name = "frmComerciarUsu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**************************************************************
' frmComerciarUsu.frm
'
'**************************************************************

'**************************************************************************
'This program is free software; you can redistribute it and/or modify
'it under the terms of the Affero General Public License;
'either version 1 of the License, or any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'Affero General Public License for more details.
'
'You should have received a copy of the Affero General Public License
'along with this program; if not, you can find it at http://www.affero.org/oagpl.html
'**************************************************************************

Option Explicit

Dim Item            As Boolean

Const WM_SYSCOMMAND As Long = &H112&

Const MOUSE_MOVE    As Long = &HF012&

Private Declare Function ReleaseCapture Lib "user32" () As Long

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Long) As Long

Public LastIndex1   As Integer

Public LasActionBuy As Boolean

Private Sub moverForm()
    
    On Error GoTo moverForm_Err
    

    Dim res As Long

    ReleaseCapture
    res = SendMessage(Me.hwnd, WM_SYSCOMMAND, MOUSE_MOVE, 0)

    
    Exit Sub

moverForm_Err:
    Call RegistrarError(Err.number, Err.Description, "frmComerciarUsu.moverForm", Erl)
    Resume Next
    
End Sub

Private Sub cmdAceptar_Click()
    
    On Error GoTo cmdAceptar_Click_Err
    
    Call WriteUserCommerceOk

    
    Exit Sub

cmdAceptar_Click_Err:
    Call RegistrarError(Err.number, Err.Description, "frmComerciarUsu.cmdAceptar_Click", Erl)
    Resume Next
    
End Sub

Private Sub cmdAceptar_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

    ' cmdAceptar.Picture = LoadInterface(Language + "comercioseguro_aceptarpress.bmp")
    'cmdAceptar.Tag = "1"
End Sub

Private Sub cmdAceptar_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    On Error GoTo cmdAceptar_MouseMove_Err
    

    If cmdAceptar.Tag = "0" Then
        cmdAceptar.Picture = LoadInterface(Language + "comercioseguro_aceptarhover.bmp")
        cmdAceptar.Tag = "1"

    End If
    
    cmdRechazar.Picture = Nothing
    cmdRechazar.Tag = "0"

    cmdOfrecer.Picture = Nothing
    cmdOfrecer.Tag = "0"

    Command2.Picture = Nothing
    Command2.Tag = "0"

    
    Exit Sub

cmdAceptar_MouseMove_Err:
    Call RegistrarError(Err.number, Err.Description, "frmComerciarUsu.cmdAceptar_MouseMove", Erl)
    Resume Next
    
End Sub

Private Sub cmdOfrecer_Click()
    
    On Error GoTo cmdOfrecer_Click_Err
    

    If Item = True Then
        If List1.ListIndex < 0 Then Exit Sub
        If List1.ItemData(List1.ListIndex) <= 0 Then Exit Sub
    
        '    If Val(txtCant.Text) > List1.ItemData(List1.ListIndex) Or _
        '        Val(txtCant.Text) <= 0 Then Exit Sub
    ElseIf Item = False Then

        '    If Val(txtCant.Text) > UserGLD Then
        '        Exit Sub
        '    End If
    End If

    If Item = True Then
        Call WriteUserCommerceOffer(List1.ListIndex + 1, Val(txtCant.Text))
    ElseIf Item = False Then
        Call WriteUserCommerceOffer(FLAGORO, Val(txtCant.Text))
    Else
        Exit Sub

    End If

    lblEstadoResp.Visible = True

    
    Exit Sub

cmdOfrecer_Click_Err:
    Call RegistrarError(Err.number, Err.Description, "frmComerciarUsu.cmdOfrecer_Click", Erl)
    Resume Next
    
End Sub

Private Sub cmdOfrecer_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

    'cmdOfrecer.Picture = LoadInterface(Language + "comercioseguro_ofrecerpress.bmp")
    'cmdOfrecer.Tag = "1"
End Sub

Private Sub cmdOfrecer_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    On Error GoTo cmdOfrecer_MouseMove_Err
    

    If cmdOfrecer.Tag = "0" Then
        cmdOfrecer.Picture = LoadInterface(Language + "comercioseguro_ofrecerhover.bmp")
        cmdOfrecer.Tag = "1"

    End If

    
    Exit Sub

cmdOfrecer_MouseMove_Err:
    Call RegistrarError(Err.number, Err.Description, "frmComerciarUsu.cmdOfrecer_MouseMove", Erl)
    Resume Next
    
End Sub

Private Sub cmdRechazar_Click()
    
    On Error GoTo cmdRechazar_Click_Err
    
    Call WriteUserCommerceReject

    
    Exit Sub

cmdRechazar_Click_Err:
    Call RegistrarError(Err.number, Err.Description, "frmComerciarUsu.cmdRechazar_Click", Erl)
    Resume Next
    
End Sub

Private Sub cmdRechazar_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

    '  cmdRechazar.Picture = LoadInterface(Language + "comercioseguro_rechazarpress.bmp")
    ' cmdRechazar.Tag = "1"
End Sub

Private Sub cmdRechazar_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    On Error GoTo cmdRechazar_MouseMove_Err
    

    If cmdRechazar.Tag = "0" Then
        cmdRechazar.Picture = LoadInterface(Language + "comercioseguro_rechazarhover.bmp")
        cmdRechazar.Tag = "1"

    End If

    cmdAceptar.Picture = Nothing
    cmdAceptar.Tag = "0"

    cmdOfrecer.Picture = Nothing
    cmdOfrecer.Tag = "0"

    Command2.Picture = Nothing
    Command2.Tag = "0"
    
    
    Exit Sub

cmdRechazar_MouseMove_Err:
    Call RegistrarError(Err.number, Err.Description, "frmComerciarUsu.cmdRechazar_MouseMove", Erl)
    Resume Next
    
End Sub

Private Sub Command2_Click()
    
    On Error GoTo Command2_Click_Err
    
    Call WriteUserCommerceEnd

    
    Exit Sub

Command2_Click_Err:
    Call RegistrarError(Err.number, Err.Description, "frmComerciarUsu.Command2_Click", Erl)
    Resume Next
    
End Sub

Private Sub Command2_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

    '  Command2.Picture = LoadInterface(Language + "comercioseguro_cancelarpress.bmp")
    '  Command2.Tag = "1"
End Sub

Private Sub Command2_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    On Error GoTo Command2_MouseMove_Err
    

    If Command2.Tag = "0" Then
        Command2.Picture = LoadInterface(Language + "comercioseguro_cancelarhover.bmp")
        Command2.Tag = "1"

    End If

    
    Exit Sub

Command2_MouseMove_Err:
    Call RegistrarError(Err.number, Err.Description, "frmComerciarUsu.Command2_MouseMove", Erl)
    Resume Next
    
End Sub

Private Sub Form_Deactivate()
    'Me.SetFocus
    'Picture1.SetFocus

End Sub

Private Sub Form_Load()
    
    On Error GoTo Form_Load_Err
    
    Call FormParser.Parse_Form(Me)
    'Carga las imagenes...?
    lblEstadoResp.Visible = False
    Item = True

    
    Exit Sub

Form_Load_Err:
    Call RegistrarError(Err.number, Err.Description, "frmComerciarUsu.Form_Load", Erl)
    Resume Next
    
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
    
    On Error GoTo Form_KeyPress_Err
    

    If (KeyAscii = 27) Then
        Unload Me

    End If

    
    Exit Sub

Form_KeyPress_Err:
    Call RegistrarError(Err.number, Err.Description, "frmComerciarUsu.Form_KeyPress", Erl)
    Resume Next
    
End Sub

Private Sub Form_LostFocus()
    
    On Error GoTo Form_LostFocus_Err
    
    Me.SetFocus
    picture1.SetFocus

    
    Exit Sub

Form_LostFocus_Err:
    Call RegistrarError(Err.number, Err.Description, "frmComerciarUsu.Form_LostFocus", Erl)
    Resume Next
    
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    On Error GoTo Form_MouseMove_Err
    
    cmdAceptar.Picture = Nothing
    cmdAceptar.Tag = "0"

    cmdRechazar.Picture = Nothing
    cmdRechazar.Tag = "0"

    cmdOfrecer.Picture = Nothing
    cmdOfrecer.Tag = "0"

    Command2.Picture = Nothing
    Command2.Tag = "0"
    moverForm

    
    Exit Sub

Form_MouseMove_Err:
    Call RegistrarError(Err.number, Err.Description, "frmComerciarUsu.Form_MouseMove", Erl)
    Resume Next
    
End Sub

Private Sub Image1_Click()
    
    On Error GoTo Image1_Click_Err
    
    Image1.Picture = LoadInterface(Language + "comercioseguro_opbjeto.bmp")
    Image2.Picture = Nothing
    List1.Enabled = True
    Item = True

    
    Exit Sub

Image1_Click_Err:
    Call RegistrarError(Err.number, Err.Description, "frmComerciarUsu.Image1_Click", Erl)
    Resume Next
    
End Sub

Private Sub Image2_Click()
    
    On Error GoTo Image2_Click_Err
    
    Image2.Picture = LoadInterface(Language + "comercioseguro_oro.bmp")
    Image1.Picture = Nothing
    List1.Enabled = False
    Item = False

    
    Exit Sub

Image2_Click_Err:
    Call RegistrarError(Err.number, Err.Description, "frmComerciarUsu.Image2_Click", Erl)
    Resume Next
    
End Sub

Private Sub List1_Click()
    
    On Error GoTo List1_Click_Err
    
    DibujaGrh frmMain.Inventario.GrhIndex(List1.ListIndex + 1)

    
    Exit Sub

List1_Click_Err:
    Call RegistrarError(Err.number, Err.Description, "frmComerciarUsu.List1_Click", Erl)
    Resume Next
    
End Sub

Public Sub DibujaGrh(grh As Long)
    
    On Error GoTo DibujaGrh_Err
    
    Call Grh_Render_To_Hdc(picture1, (grh), 0, 0)

    
    Exit Sub

DibujaGrh_Err:
    Call RegistrarError(Err.number, Err.Description, "frmComerciarUsu.DibujaGrh", Erl)
    Resume Next
    
End Sub

Private Sub List2_Click()
    
    On Error GoTo List2_Click_Err
    

    If List2.ListIndex >= 0 Then
        DibujaGrh OtroInventario(List2.ListIndex + 1).GrhIndex
        Label3.Caption = "Cantidad: " & List2.ItemData(List2.ListIndex)
        cmdAceptar.Enabled = True
        cmdRechazar.Enabled = True
    Else
        cmdAceptar.Enabled = False
        cmdRechazar.Enabled = False

    End If

    
    Exit Sub

List2_Click_Err:
    Call RegistrarError(Err.number, Err.Description, "frmComerciarUsu.List2_Click", Erl)
    Resume Next
    
End Sub

Private Sub txtCant_Change()
    
    On Error GoTo txtCant_Change_Err
    

    If Val(txtCant.Text) < 1 Then txtCant.Text = "1"
    
    If Val(txtCant.Text) > 2147483647 Then txtCant.Text = "2147483647"

    
    Exit Sub

txtCant_Change_Err:
    Call RegistrarError(Err.number, Err.Description, "frmComerciarUsu.txtCant_Change", Erl)
    Resume Next
    
End Sub

Private Sub txtCant_KeyDown(KeyCode As Integer, Shift As Integer)
    
    On Error GoTo txtCant_KeyDown_Err
    

    If Not ((KeyCode >= 48 And KeyCode <= 57) Or KeyCode = vbKeyBack Or KeyCode = vbKeyDelete Or (KeyCode >= 37 And KeyCode <= 40)) Then
        'txtCant = KeyCode
        KeyCode = 0

    End If

    
    Exit Sub

txtCant_KeyDown_Err:
    Call RegistrarError(Err.number, Err.Description, "frmComerciarUsu.txtCant_KeyDown", Erl)
    Resume Next
    
End Sub

Private Sub txtCant_KeyPress(KeyAscii As Integer)
    
    On Error GoTo txtCant_KeyPress_Err
    

    If Not ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = vbKeyBack Or KeyAscii = vbKeyDelete Or (KeyAscii >= 37 And KeyAscii <= 40)) Then
        'txtCant = KeyCode
        KeyAscii = 0

    End If

    
    Exit Sub

txtCant_KeyPress_Err:
    Call RegistrarError(Err.number, Err.Description, "frmComerciarUsu.txtCant_KeyPress", Erl)
    Resume Next
    
End Sub

'[/Alejo]

