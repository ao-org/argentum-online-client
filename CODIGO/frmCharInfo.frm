VERSION 5.00
Begin VB.Form frmCharInfo 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Información del personaje"
   ClientHeight    =   6000
   ClientLeft      =   45
   ClientTop       =   210
   ClientWidth     =   6360
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
   ScaleHeight     =   6000
   ScaleWidth      =   6360
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton desc 
      Caption         =   "Peticion"
      Height          =   495
      Left            =   2655
      MouseIcon       =   "frmCharInfo.frx":0000
      MousePointer    =   99  'Custom
      TabIndex        =   14
      Top             =   5385
      Width           =   1000
   End
   Begin VB.CommandButton Echar 
      Caption         =   "Echar"
      Height          =   495
      Left            =   1395
      MouseIcon       =   "frmCharInfo.frx":0152
      MousePointer    =   99  'Custom
      TabIndex        =   13
      Top             =   5385
      Width           =   1000
   End
   Begin VB.CommandButton Aceptar 
      Caption         =   "Aceptar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5085
      MouseIcon       =   "frmCharInfo.frx":02A4
      MousePointer    =   99  'Custom
      TabIndex        =   12
      Top             =   5385
      Width           =   1000
   End
   Begin VB.CommandButton Rechazar 
      Caption         =   "Rechazar"
      Height          =   495
      Left            =   3870
      MouseIcon       =   "frmCharInfo.frx":03F6
      MousePointer    =   99  'Custom
      TabIndex        =   11
      Top             =   5385
      Width           =   1000
   End
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "Cerrar"
      Height          =   495
      Left            =   120
      MouseIcon       =   "frmCharInfo.frx":0548
      MousePointer    =   99  'Custom
      TabIndex        =   10
      Top             =   5385
      Width           =   1000
   End
   Begin VB.Frame Frame1 
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
      Height          =   3120
      Left            =   135
      TabIndex        =   9
      Top             =   2115
      Width           =   6075
      Begin VB.TextBox txtMiembro 
         Height          =   1110
         Left            =   135
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   22
         Top             =   1800
         Width           =   5790
      End
      Begin VB.TextBox txtPeticiones 
         Height          =   1110
         Left            =   135
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   15
         Top             =   450
         Width           =   5790
      End
      Begin VB.Label lblMiembro 
         Caption         =   "Ultimos clanes en los que participó:"
         Height          =   255
         Left            =   135
         TabIndex        =   23
         Top             =   1620
         Width           =   2985
      End
      Begin VB.Label lblSolicitado 
         Caption         =   "Ultimas membresías solicitadas:"
         Height          =   255
         Left            =   135
         TabIndex        =   21
         Top             =   270
         Width           =   2985
      End
   End
   Begin VB.Frame charinfo 
      Caption         =   "General"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1980
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6075
      Begin VB.Label reputacion 
         Caption         =   "Reputacion:"
         Height          =   255
         Left            =   3060
         TabIndex        =   20
         Top             =   1440
         Width           =   2445
      End
      Begin VB.Label criminales 
         Caption         =   "Criminales asesinados:"
         Height          =   255
         Left            =   3060
         TabIndex        =   19
         Top             =   1200
         Width           =   2895
      End
      Begin VB.Label Ciudadanos 
         Caption         =   "Ciudadanos asesinados:"
         Height          =   255
         Left            =   3060
         TabIndex        =   18
         Top             =   960
         Width           =   2850
      End
      Begin VB.Label ejercito 
         Caption         =   "Faccion:"
         Height          =   255
         Left            =   3060
         TabIndex        =   17
         Top             =   720
         Width           =   2880
      End
      Begin VB.Label guildactual 
         Caption         =   "Clan Actual:"
         Height          =   255
         Left            =   3030
         TabIndex        =   16
         Top             =   480
         Width           =   2880
      End
      Begin VB.Label status 
         Caption         =   "Status:"
         Height          =   255
         Left            =   3060
         TabIndex        =   8
         Top             =   1680
         Width           =   2760
      End
      Begin VB.Label Banco 
         Caption         =   "Banco:"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   1680
         Width           =   2985
      End
      Begin VB.Label Oro 
         Caption         =   "Oro:"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   1440
         Width           =   2805
      End
      Begin VB.Label Genero 
         Caption         =   "Genero:"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   960
         Width           =   2895
      End
      Begin VB.Label Raza 
         Caption         =   "Raza:"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   480
         Width           =   2880
      End
      Begin VB.Label Clase 
         Caption         =   "Clase:"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   720
         Width           =   3270
      End
      Begin VB.Label Nivel 
         Caption         =   "Nivel:"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   1200
         Width           =   3105
      End
      Begin VB.Label Nombre 
         Caption         =   "Nombre:"
         Height          =   255
         Left            =   135
         TabIndex        =   1
         Top             =   240
         Width           =   5640
      End
   End
End
Attribute VB_Name = "frmCharInfo"
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

Public Enum CharInfoFrmType

    frmMembers
    frmMembershipRequests

End Enum

Public frmType As CharInfoFrmType

Private Sub Aceptar_Click()
    
    On Error GoTo Aceptar_Click_Err
    
    Call WriteGuildAcceptNewMember(Trim$(Right$(nombre, Len(nombre) - 8)))
    Unload frmGuildLeader
    Call WriteRequestGuildLeaderInfo
    Unload Me

    
    Exit Sub

Aceptar_Click_Err:
    Call RegistrarError(Err.number, Err.Description, "frmCharInfo.Aceptar_Click", Erl)
    Resume Next
    
End Sub

Private Sub Command1_Click()
    
    On Error GoTo Command1_Click_Err
    
    Unload Me

    
    Exit Sub

Command1_Click_Err:
    Call RegistrarError(Err.number, Err.Description, "frmCharInfo.Command1_Click", Erl)
    Resume Next
    
End Sub

Private Sub desc_Click()
    
    On Error GoTo desc_Click_Err
    
    Call WriteGuildRequestJoinerInfo(Right$(nombre, Len(nombre) - 8))

    
    Exit Sub

desc_Click_Err:
    Call RegistrarError(Err.number, Err.Description, "frmCharInfo.desc_Click", Erl)
    Resume Next
    
End Sub

Private Sub Echar_Click()
    
    On Error GoTo Echar_Click_Err
    
    Call WriteGuildKickMember(Right$(nombre, Len(nombre) - 8))
    Unload frmGuildLeader
    Call WriteRequestGuildLeaderInfo
    Unload Me

    
    Exit Sub

Echar_Click_Err:
    Call RegistrarError(Err.number, Err.Description, "frmCharInfo.Echar_Click", Erl)
    Resume Next
    
End Sub

Private Sub Form_Load()
    
    On Error GoTo Form_Load_Err
    
    Call FormParser.Parse_Form(Me)

    
    Exit Sub

Form_Load_Err:
    Call RegistrarError(Err.number, Err.Description, "frmCharInfo.Form_Load", Erl)
    Resume Next
    
End Sub

Private Sub Rechazar_Click()
    
    On Error GoTo Rechazar_Click_Err
    
    Load frmCommet
    frmCommet.t = RECHAZOPJ
    frmCommet.nombre = Right$(nombre, Len(nombre) - 8)
    frmCommet.Caption = "Ingrese motivo para rechazo"
    frmCommet.Show vbModeless, frmCharInfo

    
    Exit Sub

Rechazar_Click_Err:
    Call RegistrarError(Err.number, Err.Description, "frmCharInfo.Rechazar_Click", Erl)
    Resume Next
    
End Sub
