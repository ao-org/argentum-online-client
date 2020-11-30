VERSION 5.00
Begin VB.Form frmGoliath 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Operación bancaria"
   ClientHeight    =   7215
   ClientLeft      =   0
   ClientTop       =   -75
   ClientWidth     =   8175
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
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
   ScaleHeight     =   7215
   ScaleWidth      =   8175
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrNumber 
      Enabled         =   0   'False
      Interval        =   30
      Left            =   0
      Top             =   0
   End
   Begin VB.TextBox txtname 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   0
      EndProperty
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
      Left            =   1080
      MaxLength       =   30
      TabIndex        =   3
      Top             =   4300
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.TextBox txtDatos 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   0
      EndProperty
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   1450
      MaxLength       =   30
      TabIndex        =   0
      Text            =   "0"
      Top             =   3460
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Image cmdMasMenos 
      Height          =   315
      Index           =   1
      Left            =   2610
      Tag             =   "0"
      Top             =   3400
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Image cmdMasMenos 
      Height          =   315
      Index           =   0
      Left            =   990
      Tag             =   "0"
      Top             =   3420
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Label lblDatos 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "¡Cantidad invalida!"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   760
      TabIndex        =   1
      Top             =   5280
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Image Image2 
      Height          =   420
      Left            =   1000
      Tag             =   "0"
      Top             =   4780
      Visible         =   0   'False
      Width           =   1980
   End
   Begin VB.Image operacion 
      Height          =   420
      Index           =   4
      Left            =   7680
      Top             =   0
      Width           =   510
   End
   Begin VB.Image operacion 
      Height          =   420
      Index           =   3
      Left            =   5250
      Top             =   6220
      Width           =   1830
   End
   Begin VB.Image operacion 
      Height          =   420
      Index           =   2
      Left            =   4570
      Top             =   5010
      Width           =   1830
   End
   Begin VB.Image operacion 
      Height          =   420
      Index           =   0
      Left            =   3170
      Top             =   6230
      Width           =   1830
   End
   Begin VB.Image operacion 
      Height          =   420
      Index           =   1
      Left            =   1080
      Top             =   6230
      Width           =   1830
   End
   Begin VB.Label gold 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   255
      Left            =   4440
      TabIndex        =   2
      Top             =   1300
      Width           =   1095
   End
   Begin VB.Image Image3 
      Height          =   3330
      Left            =   630
      Tag             =   "0"
      Top             =   2260
      Width           =   2670
   End
End
Attribute VB_Name = "frmGoliath"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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

Private QueOperacion       As Byte

Private OroDep             As Long
Private m_Number             As Integer

Private m_Increment          As Integer

Private m_Interval           As Integer

Public Sub ParseBancoInfo(ByVal oro As Long, ByVal Items As Byte)
    
    On Error GoTo ParseBancoInfo_Err
    

    OroDep = oro
    gold.Caption = OroDep

    Me.Picture = LoadInterface("ventanabanco.bmp")
    HayFormularioAbierto = True
    
    txtDatos.BackColor = RGB(17, 18, 12)
    gold.ForeColor = RGB(235, 164, 14)
    lblDatos.ForeColor = RGB(235, 164, 14)
    
    txtname.BackColor = RGB(17, 18, 12)
    Me.Show vbModeless, frmMain
    

    Exit Sub

    
    Exit Sub

ParseBancoInfo_Err:
    Call RegistrarError(Err.number, Err.Description, "frmGoliath.ParseBancoInfo", Erl)
    Resume Next
    
End Sub

Private Sub cmdMasMenos_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    
    On Error GoTo cmdMasMenos_MouseDown_Err
    

    Call Sound.Sound_Play(SND_CLICK)

    Select Case Index

        Case 0
            cmdMasMenos(Index).Picture = LoadInterface("boton-sm-menos-off.bmp")
            cmdMasMenos(Index).Tag = "1"
            txtDatos.Text = str((Val(txtDatos.Text) - 1))
            m_Increment = -1

        Case 1
            cmdMasMenos(Index).Picture = LoadInterface("boton-sm-mas-off.bmp")
            cmdMasMenos(Index).Tag = "1"
            m_Increment = 1

    End Select

    tmrNumber.Interval = 30
    tmrNumber.Enabled = True

    
    Exit Sub

cmdMasMenos_MouseDown_Err:
    Call RegistrarError(Err.number, Err.Description, "frmGoliath.cmdMasMenos_MouseDown", Erl)
    Resume Next
    
End Sub

Private Sub cmdMasMenos_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    
    On Error GoTo cmdMasMenos_MouseMove_Err
    

    Select Case Index

        Case 0

            If cmdMasMenos(Index).Tag = "0" Then
                cmdMasMenos(Index).Picture = LoadInterface("boton-sm-menos-over.bmp")
                cmdMasMenos(Index).Tag = "1"

            End If

        Case 1

            If cmdMasMenos(Index).Tag = "0" Then
                cmdMasMenos(Index).Picture = LoadInterface("boton-sm-mas-over.bmp")
                cmdMasMenos(Index).Tag = "1"

            End If

    End Select

    
    Exit Sub

cmdMasMenos_MouseMove_Err:
    Call RegistrarError(Err.number, Err.Description, "frmGoliath.cmdMasMenos_MouseMove", Erl)
    Resume Next
    
End Sub

Private Sub cmdMasMenos_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    
    On Error GoTo cmdMasMenos_MouseUp_Err
    
    Call Form_MouseMove(Button, Shift, x, y)
    tmrNumber.Enabled = False

    
    Exit Sub

cmdMasMenos_MouseUp_Err:
    Call RegistrarError(Err.number, Err.Description, "frmGoliath.cmdMasMenos_MouseUp", Erl)
    Resume Next
    
End Sub
Private Sub Form_Load()
    
    On Error GoTo Form_Load_Err
    
    Call FormParser.Parse_Form(Me)
    
    Exit Sub

Form_Load_Err:
    Call RegistrarError(Err.number, Err.Description, "frmGoliath.Form_Load", Erl)
    Resume Next
    
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    
    On Error GoTo Form_KeyPress_Err
    
    If (KeyAscii = 27) Then
        Unload Me
    End If
    
    Exit Sub

Form_KeyPress_Err:
    Call RegistrarError(Err.number, Err.Description, "frmGoliath.Form_KeyPress", Erl)
    Resume Next
    
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    On Error GoTo Form_MouseDown_Err
    
    operacion(1).Tag = "0"
    operacion(2).Tag = "0"
    operacion(3).Tag = "0"

    Image3.Picture = Nothing
    txtDatos.Visible = False
    txtname.Visible = False
    lblDatos.Visible = False
    Image2.Visible = False
    cmdMasMenos(0).Visible = False
    cmdMasMenos(1).Visible = False
    
    
    Exit Sub

Form_MouseDown_Err:
    Call RegistrarError(Err.number, Err.Description, "frmGoliath.Form_MouseDown", Erl)
    Resume Next
    
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    On Error GoTo Form_MouseMove_Err
    
    operacion(0).Tag = "0"
    operacion(1).Tag = "0"
    operacion(2).Tag = "0"
    operacion(3).Tag = "0"
    'operacion(4).Tag = "0"

    operacion(0).Picture = Nothing
    operacion(1).Picture = Nothing
    operacion(2).Picture = Nothing
    operacion(3).Picture = Nothing

    Image2.Picture = Nothing
    Image2.Tag = "0"
        
     If cmdMasMenos(0).Tag = "1" Then
        cmdMasMenos(0).Picture = Nothing
        cmdMasMenos(0).Tag = "0"
    End If

    If cmdMasMenos(1).Tag = "1" Then
        cmdMasMenos(1).Picture = Nothing
        cmdMasMenos(1).Tag = "0"
    End If

    
    Exit Sub

Form_MouseMove_Err:
    Call RegistrarError(Err.number, Err.Description, "frmGoliath.Form_MouseMove", Erl)
    Resume Next
    
End Sub

Private Sub Image2_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    On Error GoTo Image2_MouseDown_Err
    
    Image2 = LoadInterface("boton-aceptar-ES-off.bmp")
    
    Exit Sub

Image2_MouseDown_Err:
    Call RegistrarError(Err.number, Err.Description, "frmGoliath.Image2_MouseDown", Erl)
    Resume Next
    
End Sub

Private Sub Image2_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    On Error GoTo Image2_MouseMove_Err
    
    If Image2.Tag = "0" Then
        Image2.Picture = LoadInterface("boton-aceptar-ES-over.bmp")
        Image2.Tag = "1"

    End If
    
    Exit Sub

Image2_MouseMove_Err:
    Call RegistrarError(Err.number, Err.Description, "frmGoliath.Image2_MouseMove", Erl)
    Resume Next
    
End Sub

Private Sub Image2_Click()
    
    On Error GoTo Image2_Click_Err
    

    Select Case QueOperacion

        Case 0
            Unload Me
    
        Case 5 'Depositar
    
            'Negativos y ceros
            If Val(txtDatos.Text) < 1 Then lblDatos.Caption = "Cantidad inválida."
            
            If UserGLD <= 0 Then
                lblDatos.Caption = "No tienes oro para depositar."
                Exit Sub
            End If
    
            Call WriteBankDepositGold(min(Val(txtDatos.Text), UserGLD))
            Unload Me

        Case 1 'Retirar
    
            'Negativos y ceros
            If Val(txtDatos.Text) < 1 Then lblDatos.Caption = "Cantidad inválida."
            
            If OroDep <= 0 Then
                lblDatos.Caption = "No tienes oro en la cuenta."
                Exit Sub
            End If
    
            Call WriteBankExtractGold(min(Val(txtDatos.Text), OroDep))
            Unload Me

        Case 4 'Transferir - Destino - Cantidad
            'Negativos y ceros
            If Val(txtDatos.Text) < 1 Then
                lblDatos.Caption = "Cantidad inválida, reintente."
                'txtDatos.Text = ""
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

    
    Exit Sub

Image2_Click_Err:
    Call RegistrarError(Err.number, Err.Description, "frmGoliath.Image2_Click", Erl)
    Resume Next
    
End Sub

Private Sub Image3_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    On Error GoTo Image3_MouseMove_Err
    
     If cmdMasMenos(0).Tag = "1" Then
        cmdMasMenos(0).Picture = Nothing
        cmdMasMenos(0).Tag = "0"
    End If

    If cmdMasMenos(1).Tag = "1" Then
        cmdMasMenos(1).Picture = Nothing
        cmdMasMenos(1).Tag = "0"
    End If
    
    
    Image2.Picture = Nothing
    Image2.Tag = "0"
    
    Exit Sub

Image3_MouseMove_Err:
    Call RegistrarError(Err.number, Err.Description, "frmGoliath.Image3_MouseMove", Erl)
    Resume Next
    
End Sub

Private Sub operacion_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    
    On Error GoTo operacion_MouseDown_Err
    

    Select Case Index

        Case 0 ' depositar
            lblDatos.Caption = ""
            Image3.Picture = LoadInterface("ventanabanco-depositar.bmp")
            operacion(Index) = LoadInterface("boton-depositar-ES-off.bmp")
            operacion(1).Tag = "0"
            operacion(2).Tag = "0"
            operacion(3).Tag = "0"
            txtDatos.Visible = True
            txtname.Visible = False
            lblDatos.Visible = True
            Image2.Visible = True
            cmdMasMenos(0).Visible = True
            cmdMasMenos(1).Visible = True
            
        Case 1 ' retirar
            lblDatos.Caption = ""
            Image3.Picture = LoadInterface("ventanabanco-retirar.bmp")
            operacion(Index) = LoadInterface("boton-retirar-ES-off.bmp")
            operacion(0).Tag = "0"
            operacion(2).Tag = "0"
            operacion(3).Tag = "0"
            txtDatos.Visible = True
            txtname.Visible = False
            lblDatos.Visible = True
            Image2.Visible = True
            cmdMasMenos(0).Visible = True
            cmdMasMenos(1).Visible = True

        Case 2 ' boveda
            lblDatos.Caption = ""
            'Image3.Picture = LoadInterface("ventanabanco.bmp")
            operacion(Index) = LoadInterface("boton-ver-boveda-es-off.bmp")
            operacion(1).Tag = "0"
            operacion(0).Tag = "0"
            operacion(3).Tag = "0"
            txtDatos.Visible = False
            txtname.Visible = False
        
        Case 3 ' transfer

            Image3.Picture = LoadInterface("ventanabanco-transferir.bmp")
            operacion(Index) = LoadInterface("boton-transferir-es-off.bmp")
            operacion(1).Tag = "0"
            operacion(2).Tag = "0"
            operacion(0).Tag = "0"
            txtDatos.Visible = True
            txtname.Visible = True
            lblDatos.Visible = True
            Image2.Visible = True
            cmdMasMenos(0).Visible = True
            cmdMasMenos(1).Visible = True
        Case 4 'Cerrar
            Unload Me
    End Select

    
    Exit Sub

operacion_MouseDown_Err:
    Call RegistrarError(Err.number, Err.Description, "frmGoliath.operacion_MouseDown", Erl)
    Resume Next
    
End Sub

Private Sub operacion_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    
    On Error GoTo operacion_MouseMove_Err
    
    Select Case Index

        Case 0 ' depositar

            If operacion(Index).Tag = "0" Then
                operacion(Index).Picture = LoadInterface("boton-depositar-ES-over.bmp")
                operacion(Index).Tag = "1"

            End If

            operacion(1).Tag = "0"
            operacion(2).Tag = "0"
            operacion(3).Tag = "0"
            'operacion(4).Tag = "0"

        Case 1 ' retirar

            If operacion(Index).Tag = "0" Then
                operacion(Index).Picture = LoadInterface("boton-retirar-ES-over.bmp")
                operacion(Index).Tag = "1"

            End If

            operacion(0).Tag = "0"
            operacion(2).Tag = "0"
            operacion(3).Tag = "0"

        Case 2 ' boveda

            If operacion(Index).Tag = "0" Then
                operacion(Index).Picture = LoadInterface("boton-ver-boveda-es-over.bmp")
                operacion(Index).Tag = "1"

            End If

            operacion(0).Tag = "0"
            operacion(1).Tag = "0"
            operacion(3).Tag = "0"

        Case 3 ' transfer

            If operacion(Index).Tag = "0" Then
                operacion(Index).Picture = LoadInterface("boton-transferir-es-over.bmp")
                operacion(Index).Tag = "1"

            End If

            operacion(1).Tag = "0"
            operacion(2).Tag = "0"
            operacion(0).Tag = "0"
    
    End Select

    
    Exit Sub

operacion_MouseMove_Err:
    Call RegistrarError(Err.number, Err.Description, "frmGoliath.operacion_MouseMove", Erl)
    Resume Next
    
End Sub

Private Sub operacion_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    
    On Error GoTo operacion_MouseUp_Err
    

    Select Case Index

        Case 0 ' depositar
            QueOperacion = 5
            'Me.Picture = LoadInterface("goliathdepositar.bmp")
            operacion(1).Tag = "0"
            operacion(2).Tag = "0"
            operacion(3).Tag = "0"
           ' operacion(4).Tag = "0"

        Case 1 ' retirar
            QueOperacion = 1
           ' Me.Picture = LoadInterface("goliathretiro.bmp")
            operacion(0).Tag = "0"
            operacion(2).Tag = "0"
            operacion(3).Tag = "0"
           ' operacion(4).Tag = "0"

        Case 2 ' boveda
            Call WriteBankStart
            Unload Me

        Case 3 ' transfer
            QueOperacion = 4
            lblDatos.Visible = True
    
    End Select

    
    Exit Sub

operacion_MouseUp_Err:
    Call RegistrarError(Err.number, Err.Description, "frmGoliath.operacion_MouseUp", Erl)
    Resume Next
    
End Sub

Private Sub tmrNumber_Timer()
    
    On Error GoTo tmrNumber_Timer_Err
    

    Const MIN_NUMBER = 1

    Const MAX_NUMBER = 10000

    m_Number = m_Number + m_Increment

    If m_Number < MIN_NUMBER Then
        m_Number = MIN_NUMBER
    ElseIf m_Number > MAX_NUMBER Then
        m_Number = MAX_NUMBER

    End If

    txtDatos.Text = format$(m_Number)
    
    If m_Interval > 1 Then
        m_Interval = m_Interval - 1
        tmrNumber.Interval = m_Interval

    End If

    
    Exit Sub

tmrNumber_Timer_Err:
    Call RegistrarError(Err.number, Err.Description, "frmGoliath.tmrNumber_Timer", Erl)
    Resume Next
    
End Sub
