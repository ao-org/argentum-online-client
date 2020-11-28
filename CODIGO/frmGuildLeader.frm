VERSION 5.00
Begin VB.Form frmGuildLeader 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Administración del Clan"
   ClientHeight    =   6330
   ClientLeft      =   45
   ClientTop       =   210
   ClientWidth     =   6075
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
   ScaleHeight     =   6330
   ScaleWidth      =   6075
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame4 
      Caption         =   "Estadisticas"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   120
      TabIndex        =   13
      Top             =   120
      Width           =   2895
      Begin VB.Label Miembros 
         Caption         =   "El clan cuenta con x miembros"
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   840
         Width           =   2535
      End
      Begin VB.Label porciento 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0%"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   490
         Width           =   2415
      End
      Begin VB.Label expcount 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "400 / 500"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   490
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.Shape Shape1 
         BorderWidth     =   2
         Height          =   255
         Left            =   120
         Top             =   480
         Width           =   2415
      End
      Begin VB.Shape EXPBAR 
         BackColor       =   &H000000C0&
         BorderColor     =   &H00FFFFFF&
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   210
         Left            =   135
         Top             =   495
         Width           =   960
      End
      Begin VB.Label beneficios 
         BackStyle       =   0  'Transparent
         Caption         =   "No atacarse / Chat de clan / Pedir ayuda (K) / Verse Invisible / Marca de clan / Verse vida / Max miembros: 25"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   120
         TabIndex        =   16
         Top             =   1320
         Width           =   2655
      End
      Begin VB.Label expacu 
         Caption         =   "Beneficios:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   1080
         Width           =   2055
      End
      Begin VB.Label nivel 
         Caption         =   "Nivel:"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Clanes"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   3120
      TabIndex        =   9
      Top             =   2160
      Width           =   2895
      Begin VB.ListBox guildslist 
         Height          =   1230
         ItemData        =   "frmGuildLeader.frx":0000
         Left            =   120
         List            =   "frmGuildLeader.frx":0002
         TabIndex        =   11
         Top             =   240
         Width           =   2655
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Detalles"
         Height          =   375
         Left            =   120
         MouseIcon       =   "frmGuildLeader.frx":0004
         MousePointer    =   99  'Custom
         TabIndex        =   10
         Top             =   1560
         Width           =   2655
      End
   End
   Begin VB.Frame txtnews 
      Caption         =   "Noticias para clan"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   120
      TabIndex        =   6
      Top             =   4320
      Width           =   5895
      Begin VB.CommandButton Command5 
         Caption         =   "Editar descripción del clan"
         Height          =   375
         Left            =   3120
         MouseIcon       =   "frmGuildLeader.frx":0156
         MousePointer    =   99  'Custom
         TabIndex        =   19
         Top             =   1440
         Width           =   2655
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Actualizar"
         Height          =   375
         Left            =   120
         MouseIcon       =   "frmGuildLeader.frx":02A8
         MousePointer    =   99  'Custom
         TabIndex        =   8
         Top             =   1440
         Width           =   2415
      End
      Begin VB.TextBox txtguildnews 
         Height          =   1095
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   7
         Top             =   240
         Width           =   5655
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Miembros"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   120
      TabIndex        =   3
      Top             =   2160
      Width           =   2895
      Begin VB.CommandButton Command2 
         Caption         =   "Detalles"
         Height          =   375
         Left            =   120
         MouseIcon       =   "frmGuildLeader.frx":03FA
         MousePointer    =   99  'Custom
         TabIndex        =   5
         Top             =   1560
         Width           =   2655
      End
      Begin VB.ListBox members 
         Height          =   1230
         ItemData        =   "frmGuildLeader.frx":054C
         Left            =   120
         List            =   "frmGuildLeader.frx":054E
         TabIndex        =   4
         Top             =   240
         Width           =   2655
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Solicitudes de ingreso"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   3120
      TabIndex        =   0
      Top             =   120
      Width           =   2895
      Begin VB.CommandButton cmdElecciones 
         Caption         =   "Abrir elecciones"
         Enabled         =   0   'False
         Height          =   375
         Left            =   120
         MouseIcon       =   "frmGuildLeader.frx":0550
         MousePointer    =   99  'Custom
         TabIndex        =   12
         Top             =   1935
         Width           =   2655
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Detalles"
         Height          =   375
         Left            =   120
         MouseIcon       =   "frmGuildLeader.frx":06A2
         MousePointer    =   99  'Custom
         TabIndex        =   2
         Top             =   1440
         Width           =   2655
      End
      Begin VB.ListBox solicitudes 
         Height          =   1035
         ItemData        =   "frmGuildLeader.frx":07F4
         Left            =   120
         List            =   "frmGuildLeader.frx":07F6
         TabIndex        =   1
         Top             =   240
         Width           =   2655
      End
   End
End
Attribute VB_Name = "frmGuildLeader"
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

Private Sub cmdElecciones_Click()
    Call WriteGuildOpenElections
    Unload Me

End Sub

Private Sub Command1_Click()

    If solicitudes.ListIndex = -1 Then Exit Sub
    
    frmCharInfo.frmType = CharInfoFrmType.frmMembershipRequests
    Call WriteGuildMemberInfo(solicitudes.List(solicitudes.ListIndex))

    'Unload Me
End Sub

Private Sub Command2_Click()

    If members.ListIndex = -1 Then Exit Sub
    
    frmCharInfo.frmType = CharInfoFrmType.frmMembers
    Call WriteGuildMemberInfo(members.List(members.ListIndex))

    'Unload Me
End Sub

Private Sub Command3_Click()

    Dim k As String

    k = Replace(txtguildnews, vbCrLf, "º")
    
    Call WriteGuildUpdateNews(k)

End Sub

Private Sub Command4_Click()
    frmGuildBrief.EsLeader = True
    Call WriteGuildRequestDetails(guildslist.List(guildslist.ListIndex))

    'Unload Me
End Sub

Private Sub Command5_Click()

    Dim fdesc As String

    fdesc = InputBox("Ingrese la descripción:", "Modificar descripción")

    fdesc = Replace(fdesc, vbCrLf, "º", , , vbBinaryCompare)
    
    If Not AsciiValidos(fdesc) Then
        MsgBox "La descripcion contiene caracteres invalidos"
        Exit Sub
    Else
        Call WriteClanCodexUpdate(fdesc)

    End If

End Sub

Private Sub Command6_Click()

    'Call frmGuildURL.Show(vbModeless, frmGuildLeader)
    'Unload Me
End Sub

Private Sub Command7_Click()
    Call WriteGuildPeacePropList

End Sub

Private Sub Command9_Click()
    Call WriteGuildAlliancePropList

End Sub

Private Sub expne_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    porciento.Visible = True
    expcount.Visible = False

End Sub

Private Sub Form_Load()
    Call FormParser.Parse_Form(Me)

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    porciento.Visible = True
    expcount.Visible = False

End Sub

Private Sub Form_Unload(Cancel As Integer)

    'frmMain.SetFocus
End Sub

Private Sub Frame4_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    porciento.Visible = True
    expcount.Visible = False

End Sub

Private Sub porciento_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    porciento.Visible = False
    expcount.Visible = True

End Sub

