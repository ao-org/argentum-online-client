VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmComerciarUsu 
   BackColor       =   &H00808080&
   BorderStyle     =   0  'None
   ClientHeight    =   8745
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8625
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   583
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   575
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      BackColor       =   &H00070406&
      BorderStyle     =   0  'None
      ForeColor       =   &H00757575&
      Height          =   270
      Left            =   4890
      TabIndex        =   9
      Text            =   "Escribe un mensaje..."
      Top             =   6990
      Width           =   3000
   End
   Begin VB.TextBox txtOro 
      Alignment       =   2  'Center
      BackColor       =   &H000D1312&
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
      Left            =   840
      TabIndex        =   7
      Text            =   "1"
      Top             =   7005
      Width           =   1755
   End
   Begin VB.PictureBox picInvOtherSell 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   510
      Left            =   4770
      ScaleHeight     =   34
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   210
      TabIndex        =   6
      Top             =   3675
      Width           =   3150
   End
   Begin VB.PictureBox picInvUserSell 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   510
      Left            =   4770
      ScaleHeight     =   34
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   210
      TabIndex        =   5
      Top             =   2160
      Width           =   3150
   End
   Begin VB.PictureBox picInv 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   3675
      Left            =   735
      ScaleHeight     =   245
      ScaleMode       =   0  'User
      ScaleWidth      =   210
      TabIndex        =   4
      Top             =   2160
      Width           =   3150
   End
   Begin VB.TextBox txtCant 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H000D1312&
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
      Left            =   1290
      TabIndex        =   1
      Text            =   "1"
      Top             =   6030
      Width           =   675
   End
   Begin RichTextLib.RichTextBox RecTxt 
      Height          =   1665
      Left            =   4830
      TabIndex        =   10
      TabStop         =   0   'False
      ToolTipText     =   "Mensajes del servidor"
      Top             =   5235
      Width           =   3075
      _ExtentX        =   5424
      _ExtentY        =   2937
      _Version        =   393217
      BackColor       =   459782
      BorderStyle     =   0
      Enabled         =   -1  'True
      HideSelection   =   0   'False
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      DisableNoScroll =   -1  'True
      Appearance      =   0
      TextRTF         =   $"frmComerciarUsu.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label lblItemName 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Vacío"
      ForeColor       =   &H00E0E0E0&
      Height          =   195
      Left            =   3420
      TabIndex        =   13
      Top             =   6600
      Width           =   435
   End
   Begin VB.Label lblUserItemName 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Vacío"
      ForeColor       =   &H00E0E0E0&
      Height          =   195
      Left            =   7530
      TabIndex        =   12
      Top             =   2835
      Width           =   435
   End
   Begin VB.Label lblOtherItemName 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Vacío"
      ForeColor       =   &H00E0E0E0&
      Height          =   195
      Left            =   7530
      TabIndex        =   11
      Top             =   4320
      Width           =   435
   End
   Begin VB.Image Image1 
      Height          =   495
      Left            =   8160
      Top             =   0
      Width           =   495
   End
   Begin VB.Image cmdOfrecerOro 
      Height          =   420
      Left            =   2775
      Top             =   6900
      Width           =   1125
   End
   Begin VB.Image cmdMenos 
      Height          =   285
      Left            =   780
      Top             =   5985
      Width           =   285
   End
   Begin VB.Image cmdMas 
      Height          =   285
      Left            =   2160
      Top             =   5985
      Width           =   285
   End
   Begin VB.Label lblEstadoResp 
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
      Left            =   5760
      TabIndex        =   0
      Top             =   7440
      Visible         =   0   'False
      Width           =   2490
   End
   Begin VB.Label lblMyGold 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H0000EAFF&
      Height          =   255
      Left            =   1080
      TabIndex        =   8
      Top             =   6600
      Width           =   2055
   End
   Begin VB.Label lblOroMiOferta 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H0000EAFF&
      Height          =   255
      Left            =   5160
      TabIndex        =   3
      Top             =   2835
      Width           =   1335
   End
   Begin VB.Label lblOro 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H0000EAFF&
      Height          =   255
      Left            =   5160
      TabIndex        =   2
      Top             =   4320
      Width           =   1335
   End
   Begin VB.Image cmdOfrecer 
      Height          =   420
      Left            =   2775
      Tag             =   "0"
      Top             =   5940
      Width           =   1125
   End
   Begin VB.Image cmdRechazar 
      Height          =   420
      Left            =   2070
      Tag             =   "0"
      Top             =   7995
      Width           =   1980
   End
   Begin VB.Image cmdAceptar 
      Height          =   420
      Left            =   4590
      Tag             =   "0"
      Top             =   7995
      Width           =   1980
   End
End
Attribute VB_Name = "frmComerciarUsu"
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


Option Explicit

Dim Item            As Boolean

Const WM_SYSCOMMAND As Long = &H112&

Const MOUSE_MOVE    As Long = &HF012&

Private Declare Function ReleaseCapture Lib "user32" () As Long

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Long) As Long

Public LastIndex1   As Integer

Public WithEvents InvUser As clsGrapchicalInventory
Attribute InvUser.VB_VarHelpID = -1
Public WithEvents InvUserSell As clsGrapchicalInventory
Attribute InvUserSell.VB_VarHelpID = -1
Public WithEvents InvOtherSell As clsGrapchicalInventory
Attribute InvOtherSell.VB_VarHelpID = -1

Public LasActionBuy As Boolean

Private Sub cmdAceptar_Click()
    On Error Goto cmdAceptar_Click_Err
    
    On Error GoTo cmdAceptar_Click_Err
    
    Call WriteUserCommerceOk

    
    Exit Sub

cmdAceptar_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmComerciarUsu.cmdAceptar_Click", Erl)
    Resume Next
    
    Exit Sub
cmdAceptar_Click_Err:
    Call TraceError(Err.Number, Err.Description, "frmComerciarUsu.cmdAceptar_Click", Erl)
End Sub

Private Sub cmdAceptar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Goto cmdAceptar_MouseDown_Err

     cmdAceptar.Picture = LoadInterface("boton-aceptar-off.bmp")
    cmdAceptar.Tag = "1"
    Exit Sub
cmdAceptar_MouseDown_Err:
    Call TraceError(Err.Number, Err.Description, "frmComerciarUsu.cmdAceptar_MouseDown", Erl)
End Sub

Private Sub cmdAceptar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Goto cmdAceptar_MouseMove_Err
    
    On Error GoTo cmdAceptar_MouseMove_Err
    

    If cmdAceptar.Tag = "0" Then
        cmdAceptar.Picture = LoadInterface("boton-aceptar-over.bmp")
        cmdAceptar.Tag = "1"

    End If
    
    cmdRechazar.Picture = Nothing
    cmdRechazar.Tag = "0"

    cmdOfrecer.Picture = Nothing
    cmdOfrecer.Tag = "0"


    
    Exit Sub

cmdAceptar_MouseMove_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmComerciarUsu.cmdAceptar_MouseMove", Erl)
    Resume Next
    
    Exit Sub
cmdAceptar_MouseMove_Err:
    Call TraceError(Err.Number, Err.Description, "frmComerciarUsu.cmdAceptar_MouseMove", Erl)
End Sub

Private Sub cmdAceptar_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Goto cmdAceptar_MouseUp_Err
     cmdAceptar.Picture = LoadInterface("boton-aceptar-over.bmp")
    cmdAceptar.Tag = "1"
    Exit Sub
cmdAceptar_MouseUp_Err:
    Call TraceError(Err.Number, Err.Description, "frmComerciarUsu.cmdAceptar_MouseUp", Erl)
End Sub


Private Sub cmdOfrecer_Click()
    On Error Goto cmdOfrecer_Click_Err
    
    On Error GoTo cmdOfrecer_Click_Err
    

    If InvUser.SelectedItem > 0 Then
        Call WriteUserCommerceOffer(InvUser.SelectedItem, Val(txtCant.Text))
    End If
    

    
    Exit Sub

cmdOfrecer_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmComerciarUsu.cmdOfrecer_Click", Erl)
    Resume Next
    
    Exit Sub
cmdOfrecer_Click_Err:
    Call TraceError(Err.Number, Err.Description, "frmComerciarUsu.cmdOfrecer_Click", Erl)
End Sub

Private Sub cmdOfrecer_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Goto cmdOfrecer_MouseDown_Err

    cmdOfrecer.Picture = LoadInterface("boton-ofrecer-off.bmp")
    cmdOfrecer.Tag = "1"
    Exit Sub
cmdOfrecer_MouseDown_Err:
    Call TraceError(Err.Number, Err.Description, "frmComerciarUsu.cmdOfrecer_MouseDown", Erl)
End Sub

Private Sub cmdOfrecer_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Goto cmdOfrecer_MouseMove_Err
    
    On Error GoTo cmdOfrecer_MouseMove_Err
    

    If cmdOfrecer.Tag = "0" Then
        cmdOfrecer.Picture = LoadInterface("boton-ofrecer-over.bmp")
        cmdOfrecer.Tag = "1"

    End If

    
    Exit Sub

cmdOfrecer_MouseMove_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmComerciarUsu.cmdOfrecer_MouseMove", Erl)
    Resume Next
    
    Exit Sub
cmdOfrecer_MouseMove_Err:
    Call TraceError(Err.Number, Err.Description, "frmComerciarUsu.cmdOfrecer_MouseMove", Erl)
End Sub

Private Sub cmdOfrecerOro_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Goto cmdOfrecerOro_MouseDown_Err
    cmdOfrecerOro.Picture = LoadInterface("boton-ofrecer-off.bmp")
    cmdOfrecerOro.Tag = "1"
    Exit Sub
cmdOfrecerOro_MouseDown_Err:
    Call TraceError(Err.Number, Err.Description, "frmComerciarUsu.cmdOfrecerOro_MouseDown", Erl)
End Sub
Private Sub cmdOfrecerOro_Click()
    On Error Goto cmdOfrecerOro_Click_Err
 On Error GoTo cmdOfrecerOro_Click_Err

    If Val(txtOro.Text) > 0 Then
        Call WriteUserCommerceOffer(FLAGORO, Val(txtOro.Text))
    End If
        
    Exit Sub

cmdOfrecerOro_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmComerciarUsu.cmdOfrecerOro_Click", Erl)
    Resume Next
    Exit Sub
cmdOfrecerOro_Click_Err:
    Call TraceError(Err.Number, Err.Description, "frmComerciarUsu.cmdOfrecerOro_Click", Erl)
End Sub

Private Sub cmdOfrecerOro_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Goto cmdOfrecerOro_MouseMove_Err

    On Error GoTo cmdOfrecerOro_MouseMove_Err
    

    If cmdOfrecerOro.Tag = "0" Then
        cmdOfrecerOro.Picture = LoadInterface("boton-ofrecer-over.bmp")
        cmdOfrecerOro.Tag = "1"

    End If

    
    Exit Sub

cmdOfrecerOro_MouseMove_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmComerciarUsu.cmdOfrecerOro_MouseMove", Erl)
    Resume Next
    Exit Sub
cmdOfrecerOro_MouseMove_Err:
    Call TraceError(Err.Number, Err.Description, "frmComerciarUsu.cmdOfrecerOro_MouseMove", Erl)
End Sub

Private Sub cmdOfrecerOro_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Goto cmdOfrecerOro_MouseUp_Err
    cmdOfrecerOro.Picture = LoadInterface("boton-ofrecer-over.bmp")
    cmdOfrecerOro.Tag = "1"
    Exit Sub
cmdOfrecerOro_MouseUp_Err:
    Call TraceError(Err.Number, Err.Description, "frmComerciarUsu.cmdOfrecerOro_MouseUp", Erl)
End Sub

Private Sub cmdRechazar_Click()
    On Error Goto cmdRechazar_Click_Err
    
    On Error GoTo cmdRechazar_Click_Err
    
    Call WriteUserCommerceReject

    
    Exit Sub

cmdRechazar_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmComerciarUsu.cmdRechazar_Click", Erl)
    Resume Next
    
    Exit Sub
cmdRechazar_Click_Err:
    Call TraceError(Err.Number, Err.Description, "frmComerciarUsu.cmdRechazar_Click", Erl)
End Sub

Private Sub cmdRechazar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Goto cmdRechazar_MouseDown_Err

     cmdRechazar.Picture = LoadInterface("boton-rechazar-off.bmp")
     cmdRechazar.Tag = "1"
    Exit Sub
cmdRechazar_MouseDown_Err:
    Call TraceError(Err.Number, Err.Description, "frmComerciarUsu.cmdRechazar_MouseDown", Erl)
End Sub

Private Sub cmdRechazar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Goto cmdRechazar_MouseMove_Err
    
    On Error GoTo cmdRechazar_MouseMove_Err
    

    If cmdRechazar.Tag = "0" Then
        cmdRechazar.Picture = LoadInterface("boton-rechazar-over.bmp")
        cmdRechazar.Tag = "1"

    End If

    cmdAceptar.Picture = Nothing
    cmdAceptar.Tag = "0"

    cmdOfrecer.Picture = Nothing
    cmdOfrecer.Tag = "0"

    
    
    Exit Sub

cmdRechazar_MouseMove_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmComerciarUsu.cmdRechazar_MouseMove", Erl)
    Resume Next
    
    Exit Sub
cmdRechazar_MouseMove_Err:
    Call TraceError(Err.Number, Err.Description, "frmComerciarUsu.cmdRechazar_MouseMove", Erl)
End Sub



Private Sub Command2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Goto Command2_MouseDown_Err

    '  Command2.Picture = LoadInterface("comercioseguro_cancelarpress.bmp")
    '  Command2.Tag = "1"
    Exit Sub
Command2_MouseDown_Err:
    Call TraceError(Err.Number, Err.Description, "frmComerciarUsu.Command2_MouseDown", Erl)
End Sub


Private Sub Form_Deactivate()
    On Error Goto Form_Deactivate_Err
    'Me.SetFocus
    'Picture1.SetFocus

    Exit Sub
Form_Deactivate_Err:
    Call TraceError(Err.Number, Err.Description, "frmComerciarUsu.Form_Deactivate", Erl)
End Sub

Private Sub Form_Load()
    On Error Goto Form_Load_Err
        
    On Error GoTo Form_Load_Err
    
    Call Aplicar_Transparencia(Me.hwnd, 240)
    
    Call FormParser.Parse_Form(Me)
    'Carga las imagenes...?
    lblEstadoResp.Visible = False
    Item = True
    Me.Picture = LoadInterface("ventanacomercio.bmp")
    AddtoRichTextBox frmComerciarUsu.RecTxt, "Antes de aceptar la transacción asegúrate de tener suficiente espacio en tu inventario, de lo contrario los items sobrantes caerán al piso.", 255, 19, 19, 1, 0
    Exit Sub
Form_Load_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmComerciarUsu.Form_Load", Erl)
    Resume Next
    
    Exit Sub
Form_Load_Err:
    Call TraceError(Err.Number, Err.Description, "frmComerciarUsu.Form_Load", Erl)
End Sub
Private Sub cmdMenos_Click()
    On Error Goto cmdMenos_Click_Err
    If Val(txtCant.Text) > 0 Then
        txtCant.Text = Val(txtCant.Text - 1)
    End If
    Exit Sub
cmdMenos_Click_Err:
    Call TraceError(Err.Number, Err.Description, "frmComerciarUsu.cmdMenos_Click", Erl)
End Sub
Private Sub cmdMas_Click()
    On Error Goto cmdMas_Click_Err
    If Val(txtCant.Text) < 10000 Then
        txtCant.Text = Val(txtCant.Text + 1)
    End If
    Exit Sub
cmdMas_Click_Err:
    Call TraceError(Err.Number, Err.Description, "frmComerciarUsu.cmdMas_Click", Erl)
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
    On Error Goto Form_KeyPress_Err
    
    On Error GoTo Form_KeyPress_Err
    

    If (KeyAscii = 27) Then
        Unload Me

    End If

    
    Exit Sub

Form_KeyPress_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmComerciarUsu.Form_KeyPress", Erl)
    Resume Next
    
    Exit Sub
Form_KeyPress_Err:
    Call TraceError(Err.Number, Err.Description, "frmComerciarUsu.Form_KeyPress", Erl)
End Sub

Private Sub Form_LostFocus()
    On Error Goto Form_LostFocus_Err
    
    On Error GoTo Form_LostFocus_Err
    
    Me.SetFocus

    
    Exit Sub

Form_LostFocus_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmComerciarUsu.Form_LostFocus", Erl)
    Resume Next
    
    Exit Sub
Form_LostFocus_Err:
    Call TraceError(Err.Number, Err.Description, "frmComerciarUsu.Form_LostFocus", Erl)
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Goto Form_MouseMove_Err
    
    On Error GoTo Form_MouseMove_Err
    
    cmdAceptar.Picture = Nothing
    cmdAceptar.Tag = "0"

    cmdRechazar.Picture = Nothing
    cmdRechazar.Tag = "0"

    cmdOfrecer.Picture = Nothing
    cmdOfrecer.Tag = "0"

    MoverForm Me.hwnd

    
    Exit Sub

Form_MouseMove_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmComerciarUsu.Form_MouseMove", Erl)
    Resume Next
    
    Exit Sub
Form_MouseMove_Err:
    Call TraceError(Err.Number, Err.Description, "frmComerciarUsu.Form_MouseMove", Erl)
End Sub


Private Sub Image1_Click()
    On Error Goto Image1_Click_Err
    On Error GoTo Image1_Click_Err
    
    Call WriteUserCommerceReject

    
    Exit Sub

Image1_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmComerciarUsu.Image1_Click", Erl)
    Resume Next
    
    Exit Sub
Image1_Click_Err:
    Call TraceError(Err.Number, Err.Description, "frmComerciarUsu.Image1_Click", Erl)
End Sub

Private Sub picInv_Click()
    On Error Goto picInv_Click_Err
    
    If InvUser.SelectedItem <> 0 Then
        
        Me.lblItemName.Caption = ObjData(InvUser.OBJIndex(InvUser.SelectedItem)).Name
    
    Else
        
        Me.lblItemName.Caption = "Vacío"
        
    End If
    
    Exit Sub
picInv_Click_Err:
    Call TraceError(Err.Number, Err.Description, "frmComerciarUsu.picInv_Click", Erl)
End Sub

Private Sub picInv_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Goto picInv_MouseMove_Err
    Call picInv_Click
    Exit Sub
picInv_MouseMove_Err:
    Call TraceError(Err.Number, Err.Description, "frmComerciarUsu.picInv_MouseMove", Erl)
End Sub

Private Sub picInvUserSell_Click()
    On Error Goto picInvUserSell_Click_Err
    
    If InvUserSell.SelectedItem <> 0 Then
        
        Me.lblUserItemName.Caption = ObjData(InvUserSell.OBJIndex(InvUserSell.SelectedItem)).Name
    
    Else
        
        Me.lblUserItemName.Caption = "Vacío"
        
    End If
    
    Exit Sub
picInvUserSell_Click_Err:
    Call TraceError(Err.Number, Err.Description, "frmComerciarUsu.picInvUserSell_Click", Erl)
End Sub

Private Sub picInvOtherSell_Click()
    On Error Goto picInvOtherSell_Click_Err
    
    If InvOtherSell.SelectedItem <> 0 Then
        
        Me.lblOtherItemName.Caption = ObjData(InvOtherSell.OBJIndex(InvOtherSell.SelectedItem)).Name
    
    Else
        
        Me.lblOtherItemName.Caption = "Vacío"
    
    End If
    
    Exit Sub
picInvOtherSell_Click_Err:
    Call TraceError(Err.Number, Err.Description, "frmComerciarUsu.picInvOtherSell_Click", Erl)
End Sub

Private Sub picInv_Paint()
    On Error Goto picInv_Paint_Err
    Call frmComerciarUsu.InvUser.ReDraw
    Exit Sub
picInv_Paint_Err:
    Call TraceError(Err.Number, Err.Description, "frmComerciarUsu.picInv_Paint", Erl)
End Sub

Private Sub picInvOtherSell_Paint()
    On Error Goto picInvOtherSell_Paint_Err
    Call frmComerciarUsu.InvOtherSell.ReDraw
    Exit Sub
picInvOtherSell_Paint_Err:
    Call TraceError(Err.Number, Err.Description, "frmComerciarUsu.picInvOtherSell_Paint", Erl)
End Sub

Private Sub picInvUserSell_Paint()
    On Error Goto picInvUserSell_Paint_Err
    Call frmComerciarUsu.InvUserSell.ReDraw
    Exit Sub
picInvUserSell_Paint_Err:
    Call TraceError(Err.Number, Err.Description, "frmComerciarUsu.picInvUserSell_Paint", Erl)
End Sub

Private Sub Text1_GotFocus()
    On Error Goto Text1_GotFocus_Err
    If Text1.Text = "Escribe un mensaje..." Then
        Text1.Text = ""
        Text1.ForeColor = vbWhite
    End If
    Exit Sub
Text1_GotFocus_Err:
    Call TraceError(Err.Number, Err.Description, "frmComerciarUsu.Text1_GotFocus", Erl)
End Sub


Private Sub Text1_KeyPress(KeyAscii As Integer)
    On Error Goto Text1_KeyPress_Err
    If KeyAscii = 13 Then
        If Text1.Text <> "" Then Call WriteCommerceSendChatMessage(Text1.Text)
        Text1.Text = ""
        KeyAscii = 0
    End If
    Exit Sub
Text1_KeyPress_Err:
    Call TraceError(Err.Number, Err.Description, "frmComerciarUsu.Text1_KeyPress", Erl)
End Sub


Private Sub Text1_LostFocus()
    On Error Goto Text1_LostFocus_Err
    If Text1.Text = "" Then
        Text1.Text = "Escribe un mensaje..."
        Text1.ForeColor = &H757575
    End If
    Exit Sub
Text1_LostFocus_Err:
    Call TraceError(Err.Number, Err.Description, "frmComerciarUsu.Text1_LostFocus", Erl)
End Sub

Private Sub txtCant_Change()
    On Error Goto txtCant_Change_Err
    
    On Error GoTo txtCant_Change_Err
    

    If Val(txtCant.Text) < 1 Then txtCant.Text = "1"
    
    If Val(txtCant.Text) > 2147483647 Then txtCant.Text = "2147483647"

    
    Exit Sub

txtCant_Change_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmComerciarUsu.txtCant_Change", Erl)
    Resume Next
    
    Exit Sub
txtCant_Change_Err:
    Call TraceError(Err.Number, Err.Description, "frmComerciarUsu.txtCant_Change", Erl)
End Sub

Private Sub txtCant_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Goto txtCant_KeyDown_Err
    
    On Error GoTo txtCant_KeyDown_Err
    

    If Not ((KeyCode >= 48 And KeyCode <= 57) Or KeyCode = vbKeyBack Or KeyCode = vbKeyDelete Or (KeyCode >= 37 And KeyCode <= 40)) Then
        'txtCant = KeyCode
        KeyCode = 0

    End If

    
    Exit Sub

txtCant_KeyDown_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmComerciarUsu.txtCant_KeyDown", Erl)
    Resume Next
    
    Exit Sub
txtCant_KeyDown_Err:
    Call TraceError(Err.Number, Err.Description, "frmComerciarUsu.txtCant_KeyDown", Erl)
End Sub

Private Sub txtCant_KeyPress(KeyAscii As Integer)
    On Error Goto txtCant_KeyPress_Err
    
    On Error GoTo txtCant_KeyPress_Err
    

    If Not ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = vbKeyBack Or KeyAscii = vbKeyDelete Or (KeyAscii >= 37 And KeyAscii <= 40)) Then
        'txtCant = KeyCode
        KeyAscii = 0

    End If

    
    Exit Sub

txtCant_KeyPress_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmComerciarUsu.txtCant_KeyPress", Erl)
    Resume Next
    
    Exit Sub
txtCant_KeyPress_Err:
    Call TraceError(Err.Number, Err.Description, "frmComerciarUsu.txtCant_KeyPress", Erl)
End Sub

'[/Alejo]

