VERSION 5.00
Begin VB.Form frmGoliath 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Operación bancaria"
   ClientHeight    =   7212
   ClientLeft      =   0
   ClientTop       =   -72
   ClientWidth     =   8172
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.4
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   601
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   681
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frameRetirar 
      BackColor       =   &H80000007&
      BorderStyle     =   0  'None
      Height          =   3330
      Left            =   600
      TabIndex        =   0
      Top             =   2100
      Visible         =   0   'False
      Width           =   2670
      Begin VB.TextBox txtDatos 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Left            =   825
         MaxLength       =   30
         TabIndex        =   3
         Text            =   "0"
         Top             =   1200
         Width           =   975
      End
      Begin VB.TextBox txtname 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   450
         MaxLength       =   30
         TabIndex        =   2
         Top             =   2040
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.Image cmdAceptar 
         Height          =   420
         Left            =   375
         Tag             =   "0"
         Top             =   2520
         Width           =   1980
      End
      Begin VB.Label lblDatos 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   3015
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.Image cmdMenos 
         Height          =   315
         Left            =   360
         Tag             =   "0"
         Top             =   1155
         Width           =   315
      End
      Begin VB.Image cmdMas 
         Height          =   315
         Left            =   1980
         Tag             =   "0"
         Top             =   1140
         Width           =   315
      End
      Begin VB.Image Image3 
         Height          =   3330
         Left            =   0
         Tag             =   "0"
         Top             =   0
         Width           =   2670
      End
   End
   Begin VB.Timer tmrNumber 
      Enabled         =   0   'False
      Interval        =   30
      Left            =   0
      Top             =   0
   End
   Begin VB.Image cmdCerrar 
      Height          =   420
      Left            =   7680
      Top             =   0
      Width           =   510
   End
   Begin VB.Image cmdTransferir 
      Height          =   420
      Left            =   5250
      Top             =   6220
      Width           =   1830
   End
   Begin VB.Image cmdBoveda 
      Height          =   420
      Left            =   4570
      Top             =   5010
      Width           =   1830
   End
   Begin VB.Image cmdDepositar 
      Height          =   420
      Left            =   3170
      Top             =   6230
      Width           =   1830
   End
   Begin VB.Image cmdRetirar 
      Height          =   420
      Left            =   1080
      Top             =   6230
      Width           =   1830
   End
   Begin VB.Label gold 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   255
      Left            =   4440
      TabIndex        =   1
      Top             =   1300
      Width           =   1095
   End
End
Attribute VB_Name = "frmGoliath"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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
'*****************************************************************
'frmGoliath - ImperiumAO - v1.3.0
'*****************************************************************
'Respective portions copyrighted by contributors listed below.
'
'This library is free software; you can redistribute it and/or
'modify it under the terms of the GNU Lesser General Public
'License as published by the Free Software Foundation version 2.1 of
'the License
'
'This library is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU
'Lesser General Public License for more details.
'
'You should have received a copy of the GNU Lesser General Public
'License along with this library; if not, write to the Free Software
'Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA
'*****************************************************************

'*****************************************************************
'Augusto José Rando (barrin@imperiumao.com.ar)
'   - First Relase
'*****************************************************************

Option Explicit

Private TipoOperacion       As Byte

Private OroDep             As Long
Private m_Number             As Integer

Private m_Increment          As Integer

Private m_Interval           As Integer

Private cBotonBoveda As clsGraphicalButton
Private cBotonRetirar As clsGraphicalButton
Private cBotonDepositar As clsGraphicalButton
Private cBotonTransferir As clsGraphicalButton
Private cBotonAceptar As clsGraphicalButton
Private cBotonCerrar As clsGraphicalButton
Private cBotonMas As clsGraphicalButton
Private cBotonMenos As clsGraphicalButton

Public Sub ParseBancoInfo(ByVal oro As Long, ByVal Items As Byte)
    
    On Error GoTo ParseBancoInfo_Err

    OroDep = oro
    gold.Caption = OroDep
    
    txtDatos.BackColor = RGB(17, 18, 12)
    gold.ForeColor = RGB(235, 164, 14)
    lblDatos.ForeColor = RGB(235, 164, 14)
    
    txtname.BackColor = RGB(17, 18, 12)
    Me.Show vbModeless, frmMain
    
    Exit Sub

ParseBancoInfo_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmGoliath.ParseBancoInfo", Erl)
    Resume Next
    
End Sub


Public Sub UpdateBankGld(ByVal oro As Long)
    
    On Error GoTo ParseBancoInfo_Err

    OroDep = oro
    gold.Caption = OroDep
    
    Exit Sub

ParseBancoInfo_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmGoliath.ParseBancoInfo", Erl)
    Resume Next
    
End Sub

Private Sub Form_Load()
    
    On Error GoTo Form_Load_Err
    
    Call Aplicar_Transparencia(Me.hwnd, 240)
    
    Call FormParser.Parse_Form(Me)

    Me.Picture = LoadInterface("ventanabanco.bmp")
    
    Call LoadButtons
    Exit Sub

Form_Load_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmGoliath.Form_Load", Erl)
    Resume Next
    
End Sub

Private Sub LoadButtons()
       
    Set cBotonBoveda = New clsGraphicalButton
    Set cBotonRetirar = New clsGraphicalButton
    Set cBotonDepositar = New clsGraphicalButton
    Set cBotonTransferir = New clsGraphicalButton
    Set cBotonAceptar = New clsGraphicalButton
    Set cBotonCerrar = New clsGraphicalButton
    Set cBotonMas = New clsGraphicalButton
    Set cBotonMenos = New clsGraphicalButton

    Call cBotonBoveda.Initialize(cmdBoveda, "boton-ver-boveda-default.bmp", _
                                                "boton-ver-boveda-over.bmp", _
                                                "boton-ver-boveda-off.bmp", Me)
    
    Call cBotonRetirar.Initialize(cmdRetirar, "boton-retirar-default.bmp", _
                                                "boton-retirar-over.bmp", _
                                                "boton-retirar-off.bmp", Me)
                                                
    Call cBotonDepositar.Initialize(cmdDepositar, "boton-depositar-default.bmp", _
                                                "boton-depositar-over.bmp", _
                                                "boton-depositar-off.bmp", Me)
    
    Call cBotonTransferir.Initialize(cmdTransferir, "boton-transferir-default.bmp", _
                                                "boton-transferir-over.bmp", _
                                                "boton-transferir-off.bmp", Me)
                                                
    Call cBotonAceptar.Initialize(cmdAceptar, "boton-aceptar-default.bmp", _
                                                "boton-aceptar-over.bmp", _
                                                "boton-aceptar-off.bmp", Me)
                                                
    Call cBotonCerrar.Initialize(cmdCerrar, "boton-cerrar-default.bmp", _
                                                "boton-cerrar-over.bmp", _
                                                "boton-cerrar-off.bmp", Me)
                                                
    Call cBotonMas.Initialize(cmdMas, "boton-sm-mas-default.bmp", _
                                                "boton-sm-mas-over.bmp", _
                                                "boton-sm-mas-off.bmp", Me)
                                                
    Call cBotonMenos.Initialize(cmdMenos, "boton-sm-menos-default.bmp", _
                                                "boton-sm-menos-over.bmp", _
                                                "boton-sm-menos-off.bmp", Me)
End Sub


Private Sub Form_KeyPress(KeyAscii As Integer)
    
    On Error GoTo Form_KeyPress_Err
    
    If (KeyAscii = 27) Then
        Unload Me
    End If
    
    Exit Sub

Form_KeyPress_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmGoliath.Form_KeyPress", Erl)
    Resume Next
    
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    On Error GoTo Form_MouseDown_Err
    
    frameRetirar.Visible = False
    
    Exit Sub

Form_MouseDown_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmGoliath.Form_MouseDown", Erl)
    Resume Next
    
End Sub

Private Sub cmdDepositar_Click()
    lblDatos.Caption = ""
    Image3.Picture = LoadInterface("ventanabanco-depositar.bmp")
    frameRetirar.Visible = True
    txtname.Visible = False
    TipoOperacion = 1
End Sub

Private Sub cmdRetirar_Click()
    lblDatos.Caption = ""
    Image3.Picture = LoadInterface("ventanabanco-retirar.bmp")
    frameRetirar.Visible = True
    txtname.Visible = False
    TipoOperacion = 2
End Sub

Private Sub cmdTransferir_Click()
    Image3.Picture = LoadInterface("ventanabanco-transferir.bmp")
    frameRetirar.Visible = True
    txtname.Visible = True
    TipoOperacion = 3
End Sub
Private Sub cmdBoveda_Click()
    Call WriteBankStart
    Unload Me
End Sub
Private Sub cmdCerrar_Click()
    Unload Me
End Sub
Private Sub cmdAceptar_Click()
    
    Select Case TipoOperacion

        Case 0
            Unload Me
    
        Case 1 'Depositar
            'Negativos y ceros
            If Val(txtDatos.Text) < 1 Then lblDatos.Caption = "Cantidad inválida."
            
            If UserGLD <= 0 Then
                lblDatos.Caption = "No tienes oro para depositar."
                Exit Sub
            End If
    
            Call WriteBankDepositGold(min(Val(txtDatos.Text), UserGLD))

        Case 2 'Retirar
            'Negativos y ceros
            If Val(txtDatos.Text) < 1 Then lblDatos.Caption = "Cantidad inválida."
            
            If OroDep <= 0 Then
                lblDatos.Caption = "No tienes oro en la cuenta."
                Exit Sub
            End If
    
            Call WriteBankExtractGold(min(Val(txtDatos.Text), OroDep))

        Case 3 'Transferir
            'Negativos y ceros
            If Val(txtDatos.Text) < 1 Then
                lblDatos.Caption = "Cantidad inválida, reintente."

                Exit Sub
            End If
            
            If OroDep <= 0 Then
                lblDatos.Caption = "No tienes oro en la cuenta."
                Exit Sub
            End If

            If txtname.Text <> "" Then
                Call WriteTransFerGold(min(Val(txtDatos.Text), OroDep), txtname.Text)
                Unload Me
            Else
                lblDatos.Caption = "¡Nombre de destino inválido!"
                txtDatos.Text = ""
            End If
    End Select

End Sub

Private Sub cmdMas_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    m_Increment = 1
    tmrNumber.Interval = 30
    tmrNumber.Enabled = True
End Sub

Private Sub cmdMenos_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    txtDatos.Text = str((Val(txtDatos.Text) - 1))
    m_Increment = -1
    tmrNumber.Interval = 30
    tmrNumber.Enabled = True
End Sub

Private Sub cmdMas_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    tmrNumber.Enabled = False
End Sub

Private Sub cmdMenos_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    tmrNumber.Enabled = False
End Sub


Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MoverForm Me.hwnd
End Sub

Private Sub tmrNumber_Timer()
    
    On Error GoTo tmrNumber_Timer_Err

    Const MIN_NUMBER = 1

    Const MAX_NUMBER = 10000

    txtDatos = txtDatos + m_Increment

    If txtDatos < MIN_NUMBER Then
        txtDatos = MIN_NUMBER
    ElseIf txtDatos > MAX_NUMBER Then
        txtDatos = MAX_NUMBER

    End If

    txtDatos.Text = format$(txtDatos)
    
    If m_Interval > 1 Then
        m_Interval = m_Interval - 1
        tmrNumber.Interval = m_Interval

    End If

    
    Exit Sub

tmrNumber_Timer_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmGoliath.tmrNumber_Timer", Erl)
    Resume Next
    
End Sub

Private Sub txtname_KeyPress(KeyAscii As Integer)
    If KeyAscii = 8 Then
        Exit Sub
    End If
    
    If InStr(" abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ", Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txtDatos_KeyPress(KeyAscii As Integer)
    If KeyAscii = 8 Then
        Exit Sub
    End If
    
    If InStr("0123456789", Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub
