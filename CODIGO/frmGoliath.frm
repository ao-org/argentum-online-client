VERSION 5.00
Begin VB.Form frmGoliath 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Operación bancaria"
   ClientHeight    =   5610
   ClientLeft      =   0
   ClientTop       =   -75
   ClientWidth     =   6510
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
   ScaleHeight     =   5610
   ScaleWidth      =   6510
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
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
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   3850
      MaxLength       =   30
      TabIndex        =   0
      Text            =   "0"
      Top             =   1740
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3760
      TabIndex        =   8
      Top             =   1140
      Width           =   2295
   End
   Begin VB.Label gold 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   255
      Left            =   1320
      TabIndex        =   7
      Top             =   2100
      Width           =   1095
   End
   Begin VB.Label operacion 
      BackStyle       =   0  'Transparent
      Height          =   855
      Index           =   4
      Left            =   2400
      TabIndex        =   6
      Tag             =   "0"
      Top             =   4080
      Width           =   495
   End
   Begin VB.Label operacion 
      BackStyle       =   0  'Transparent
      Height          =   855
      Index           =   3
      Left            =   1960
      TabIndex        =   5
      Tag             =   "0"
      Top             =   4080
      Width           =   495
   End
   Begin VB.Label operacion 
      BackStyle       =   0  'Transparent
      Height          =   855
      Index           =   2
      Left            =   1460
      TabIndex        =   4
      Tag             =   "0"
      Top             =   4080
      Width           =   495
   End
   Begin VB.Label operacion 
      BackStyle       =   0  'Transparent
      Height          =   855
      Index           =   1
      Left            =   1040
      TabIndex        =   3
      Tag             =   "0"
      Top             =   4080
      Width           =   495
   End
   Begin VB.Label operacion 
      BackStyle       =   0  'Transparent
      Height          =   855
      Index           =   0
      Left            =   480
      TabIndex        =   2
      Tag             =   "0"
      Top             =   4080
      Width           =   495
   End
   Begin VB.Image Image3 
      Height          =   1275
      Left            =   360
      Tag             =   "0"
      Top             =   4040
      Width           =   2835
   End
   Begin VB.Image Image2 
      Height          =   450
      Left            =   4020
      Tag             =   "0"
      Top             =   4780
      Width           =   1830
   End
   Begin VB.Label lblDatos 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
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
      Left            =   3360
      TabIndex        =   1
      Top             =   1180
      Visible         =   0   'False
      Width           =   3135
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

Private QueOperacion As Byte

Private CantTransferencia As Long
Private EtapaTransferencia As Byte
Private OroDep As Long

Public Sub ParseBancoInfo(ByVal oro As Long, ByVal Items As Byte)

OroDep = oro
gold.Caption = "$" & OroDep

Me.Picture = LoadInterface("goliath.bmp")
HayFormularioAbierto = True
Me.Show vbModeless, frmmain

Exit Sub
End Sub


Private Sub Form_Click()
'Me.Picture = LoadInterface("goliath.bmp")
'QueOperacion = 0
'gold.Visible = True
'txtDatos.Visible = False
End Sub

Private Sub Form_Load()
Call FormParser.Parse_Form(Me)
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
If (KeyAscii = 27) Then
    Unload Me
End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)




operacion(1).Tag = "0"
operacion(2).Tag = "0"
operacion(3).Tag = "0"
operacion(4).Tag = "0"


Image3.Picture = Nothing

Image2.Picture = Nothing


Image2.Tag = "0"

End Sub

Private Sub Image2_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Image2.Tag = "0" Then
        Image2.Picture = LoadInterface("goliathaceptar.bmp")
        Image2.Tag = "1"
    End If
End Sub


Private Sub Image2_Click()

Select Case QueOperacion
    Case 0
        Unload Me
    
    Case 5 'Depositar
    
        'Negativos y ceros
        If (Val(txtDatos.Text) < 1 And (UCase$(txtDatos.Text) <> "TODO")) Then lblDatos.Caption = "Cantidad inválida."
    
        If Val(txtDatos.Text) <= UserGLD Or UCase$(txtDatos.Text) = "TODO" Then
                    
    
            Call WriteBankDepositGold(IIf(Val(txtDatos.Text) > 0, Val(txtDatos.Text), UserGLD))
            Unload Me
        Else
            lblDatos.Caption = "No tienes esa cantidad, reintenta."
        End If
    Case 1 'Retirar
    
        'Negativos y ceros
        If (Val(txtDatos.Text) < 1 And (UCase$(txtDatos.Text) <> "TODO")) Then lblDatos.Caption = "Cantidad inválida."
    
        If Val(txtDatos.Text) <= OroDep Or UCase$(txtDatos.Text) = "TODO" Then
            Call WriteBankExtractGold(IIf(Val(txtDatos.Text) > 0, Val(txtDatos.Text), OroDep))
            Unload Me
        Else
            lblDatos.Caption = "No tienes esa cantidad, reintenta."
        End If
    Case 4 'Transferir - Destino - Cantidad
        If EtapaTransferencia = 0 Then
        
            'Negativos y ceros
            If Val(txtDatos.Text) < 1 Then
                Label1.Caption = "Cantidad inválida, reintenta."
                txtDatos.Text = ""
                Exit Sub
            End If
            
            If Val(txtDatos.Text) <= OroDep Then
                CantTransferencia = Val(txtDatos.Text)
                Label1.Caption = "¿A quién le deseas transferir?"
                EtapaTransferencia = 1
                txtDatos.Text = ""
            Else
                Label1.Caption = "No tienes esa cantidad depositada."
                txtDatos.Text = ""
            End If
        ElseIf EtapaTransferencia = 1 Then
            If txtDatos.Text <> "" Then
            Call WriteTransFerGold(CantTransferencia, txtDatos.Text)
                Unload Me
            Else
                Label1.Caption = "¡Nombre de destino inválido!"
                txtDatos.Text = ""
            End If
        End If
End Select

End Sub


Private Sub lstBanco_Click()



End Sub


Private Sub operacion_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
Select Case Index
    Case 0 ' depositar
        lblDatos.Caption = ""
        Me.Picture = LoadInterface("goliathdepositar.bmp")
        operacion(1).Tag = "0"
        operacion(2).Tag = "0"
        operacion(3).Tag = "0"
        operacion(4).Tag = "0"
        txtDatos.Visible = True
        gold.Visible = False
        Label1.Visible = False
    Case 1 ' retirar
        lblDatos.Caption = ""

        Me.Picture = LoadInterface("goliathretiro.bmp")
        operacion(0).Tag = "0"
        operacion(2).Tag = "0"
        operacion(3).Tag = "0"
        operacion(4).Tag = "0"
        txtDatos.Visible = True
        gold.Visible = False
        Label1.Visible = False
    Case 2 ' boveda
        lblDatos.Caption = ""
        Me.Picture = LoadInterface("goliathboveda.bmp")

        operacion(1).Tag = "0"
        operacion(0).Tag = "0"
        operacion(3).Tag = "0"
        operacion(4).Tag = "0"
        txtDatos.Visible = False
        Label1.Visible = False
        
    Case 3 ' transfer
        

        Me.Picture = LoadInterface("goliathtransferencia.bmp")

        operacion(1).Tag = "0"
        operacion(2).Tag = "0"
        operacion(0).Tag = "0"
        operacion(4).Tag = "0"
        txtDatos.Visible = True
        gold.Visible = False
        Label1.Caption = "Ingrese la cantidad a transferir"
        Label1.Visible = True
        
    Case 4 ' shop
        lblDatos.Caption = ""
        Me.Picture = LoadInterface("goliathshop.bmp")

        operacion(1).Tag = "0"
        operacion(2).Tag = "0"
        operacion(3).Tag = "0"
        operacion(0).Tag = "0"
        txtDatos.Visible = False
        Label1.Visible = False
    
    End Select
End Sub

Private Sub operacion_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
Select Case Index
    Case 0 ' depositar
        If operacion(Index).Tag = "0" Then
            Image3.Picture = LoadInterface("goliathdepositoover.bmp")
            operacion(Index).Tag = "1"
        End If
        operacion(1).Tag = "0"
        operacion(2).Tag = "0"
        operacion(3).Tag = "0"
        operacion(4).Tag = "0"
    Case 1 ' retirar
        If operacion(Index).Tag = "0" Then
            Image3.Picture = LoadInterface("goliathretirarover.bmp")
            operacion(Index).Tag = "1"
        End If
        operacion(0).Tag = "0"
        operacion(2).Tag = "0"
        operacion(3).Tag = "0"
        operacion(4).Tag = "0"
    Case 2 ' boveda
        If operacion(Index).Tag = "0" Then
            Image3.Picture = LoadInterface("goliathbovedaover.bmp")
            operacion(Index).Tag = "1"
        End If
        operacion(0).Tag = "0"
        operacion(1).Tag = "0"
        operacion(3).Tag = "0"
        operacion(4).Tag = "0"
    Case 3 ' transfer
        If operacion(Index).Tag = "0" Then
            Image3.Picture = LoadInterface("goliathtransferenciaover.bmp")
            operacion(Index).Tag = "1"
        End If
        operacion(1).Tag = "0"
        operacion(2).Tag = "0"
        operacion(0).Tag = "0"
        operacion(4).Tag = "0"
        
    Case 4 ' shop
        If operacion(Index).Tag = "0" Then
            Image3.Picture = LoadInterface("goliathshopover.bmp")
            operacion(Index).Tag = "1"
        End If
        
        operacion(1).Tag = "0"
        operacion(2).Tag = "0"
        operacion(3).Tag = "0"
        operacion(0).Tag = "0"
    
    
    End Select
End Sub

Private Sub operacion_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
Select Case Index
    Case 0 ' depositar
        QueOperacion = 5
        Me.Picture = LoadInterface("goliathdepositar.bmp")
        operacion(1).Tag = "0"
        operacion(2).Tag = "0"
        operacion(3).Tag = "0"
        operacion(4).Tag = "0"
    Case 1 ' retirar
        QueOperacion = 1
        Me.Picture = LoadInterface("goliathretiro.bmp")
        operacion(0).Tag = "0"
        operacion(2).Tag = "0"
        operacion(3).Tag = "0"
        operacion(4).Tag = "0"
    Case 2 ' boveda
        Call WriteBankStart
        Unload Me
    Case 3 ' transfer
    QueOperacion = 4
        EtapaTransferencia = 0
        lblDatos.Caption = "¿Qué cantidad deseas transferir?"
    Case 4 ' shop
        Call WriteTraerShop
        Unload Me
    
    
    End Select
End Sub

