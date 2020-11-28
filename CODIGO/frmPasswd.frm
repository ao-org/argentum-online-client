VERSION 5.00
Begin VB.Form frmPasswd 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4620
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   4110
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   4620
   ScaleWidth      =   4110
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Ingrese los siguientes datos"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3495
      Left            =   120
      TabIndex        =   3
      Top             =   0
      Width           =   3855
      Begin VB.TextBox txtPasswdCheck 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         IMEMode         =   3  'DISABLE
         Left            =   120
         PasswordChar    =   "*"
         TabIndex        =   8
         Top             =   2400
         Width           =   3510
      End
      Begin VB.TextBox txtPasswd 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         IMEMode         =   3  'DISABLE
         Left            =   120
         PasswordChar    =   "*"
         TabIndex        =   7
         Top             =   1800
         Width           =   3510
      End
      Begin VB.TextBox txtCorreo 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   600
         Width           =   3510
      End
      Begin VB.TextBox txtCorreoCheck 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   1200
         Width           =   3495
      End
      Begin VB.TextBox TxtVerif 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   3000
         Width           =   3510
      End
      Begin VB.Label Label3 
         Caption         =   "Dirección de correo electronico:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   120
         TabIndex        =   14
         Top             =   360
         Width           =   3555
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Verifiación del password:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   13
         Top             =   2160
         Width           =   2175
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Password:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   12
         Top             =   1560
         Width           =   915
      End
      Begin VB.Label Label6 
         Caption         =   "Verificación del correo electronico:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   120
         TabIndex        =   11
         Top             =   960
         Width           =   3555
      End
      Begin VB.Label lblverificacion 
         Caption         =   "Codigo"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   3000
         TabIndex        =   10
         Top             =   2760
         Width           =   615
      End
      Begin VB.Label Label7 
         Caption         =   "Escriba los siguientes caracteres:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   2760
         Width           =   3615
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Volver"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   120
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Top             =   3600
      Width           =   1065
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Aceptar"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   3000
      MousePointer    =   99  'Custom
      TabIndex        =   1
      Top             =   3600
      Width           =   960
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Autocompletar"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1440
      TabIndex        =   0
      Top             =   3600
      Width           =   1455
   End
   Begin VB.Line Line2 
      X1              =   120
      X2              =   120
      Y1              =   3960
      Y2              =   4560
   End
   Begin VB.Label Estado 
      Alignment       =   2  'Center
      Caption         =   "Rellene la siguiente información por favor."
      Height          =   480
      Left            =   120
      TabIndex        =   15
      Top             =   4080
      Width           =   3855
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   4080
      Y1              =   3960
      Y2              =   3960
   End
   Begin VB.Line Line3 
      X1              =   4080
      X2              =   120
      Y1              =   4560
      Y2              =   4560
   End
   Begin VB.Line Line4 
      X1              =   4080
      X2              =   4080
      Y1              =   3960
      Y2              =   4560
   End
End
Attribute VB_Name = "frmPasswd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Argentum Online 0.11.6
'
'Copyright (C) 2002 Márquez Pablo Ignacio
'Copyright (C) 2002 Otto Perez
'Copyright (C) 2002 Aaron Perkins
'Copyright (C) 2002 Matías Fernando Pequeño
'
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
'
'Argentum Online is based on Baronsoft's VB6 Online RPG
'You can contact the original creator of ORE at aaron@baronsoft.com
'for more information about ORE please visit http://www.baronsoft.com/
'
'
'You can contact me at:
'morgolock@speedy.com.ar
'www.geocities.com/gmorgolock
'Calle 3 número 983 piso 7 dto A
'La Plata - Pcia, Buenos Aires - Republica Argentina
'Código Postal 1900
'Pablo Ignacio Márquez


Option Explicit

Function CheckDatos() As Boolean

If txtPasswd.Text <> txtPasswdCheck.Text Then
    MsgBox "Los passwords que tipeo no coinciden, por favor vuelva a ingresarlos."
    Exit Function
End If

CheckDatos = True

End Function

Private Sub Command1_Click()

If CheckDatos() Then

    UserPassword = txtPasswd.Text
    'UserEmail = txtCorreo.Text
    

    
#If UsarWrench = 1 Then
    frmMain.Socket1.HostName = IPdelServidor
    frmMain.Socket1.RemotePort = PuertoDelServidor
#End If
    
    EstadoLogin = E_MODO.CrearNuevoPj
    
#If UsarWrench = 1 Then
    If Not frmMain.Socket1.Connected Then
#Else
    If frmMain.Winsock1.State <> sckConnected Then
#End If
        MsgBox "Error: Se ha perdido la conexion con el server."
        Unload Me
        
    Else
        Call Login
    End If
End If

End Sub

Private Sub Command2_Click()
    EstadoLogin = E_MODO.Dados
    Unload Me
End Sub

Private Sub Command3_Click()
txtCorreo = "pablito_3_15@hotmail.com"
txtCorreoCheck = "pablito_3_15@hotmail.com"
txtPasswd = "a"
txtPasswdCheck = "a"
TxtVerif = lblverificacion
Call Command1_Click
End Sub

Private Sub Form_Load()
Dim Caracteres As Long
Caracteres = RandomNumber(1000, 9000)
lblverificacion.Caption = Caracteres
End Sub

