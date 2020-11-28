VERSION 5.00
Begin VB.Form frmBindKey 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Asignar acción"
   ClientHeight    =   2700
   ClientLeft      =   0
   ClientTop       =   -75
   ClientWidth     =   3375
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
   ScaleHeight     =   2700
   ScaleWidth      =   3375
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtComandoEnvio 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   240
      Left            =   840
      TabIndex        =   0
      Text            =   "Comando a enviar"
      Top             =   1120
      Width           =   1695
   End
   Begin VB.Label validez 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Macro valido unicamente para AutoUsar"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   45
      TabIndex        =   7
      Top             =   685
      Width           =   3255
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Height          =   195
      Left            =   480
      TabIndex        =   6
      Top             =   2000
      Width           =   1575
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Height          =   200
      Left            =   480
      TabIndex        =   5
      Top             =   1630
      Width           =   1815
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Height          =   200
      Left            =   480
      TabIndex        =   4
      Top             =   1350
      Width           =   1815
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   480
      TabIndex        =   3
      Top             =   795
      Width           =   1455
   End
   Begin VB.Image optAccionImg 
      Height          =   315
      Index           =   2
      Left            =   120
      Top             =   1610
      Width           =   315
   End
   Begin VB.Image optAccionImg 
      Height          =   315
      Index           =   3
      Left            =   120
      Top             =   1920
      Width           =   315
   End
   Begin VB.Image optAccionImg 
      Height          =   315
      Index           =   1
      Left            =   120
      Top             =   1290
      Width           =   315
   End
   Begin VB.Image optAccionImg 
      Height          =   315
      Index           =   0
      Left            =   120
      Top             =   930
      Width           =   315
   End
   Begin VB.Image cmdAccept 
      Height          =   480
      Left            =   1810
      Tag             =   "0"
      Top             =   2220
      Width           =   1575
   End
   Begin VB.Image cmdCancel 
      Height          =   480
      Left            =   0
      Tag             =   "0"
      Top             =   2220
      Width           =   1815
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "/"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   255
      Left            =   720
      TabIndex        =   2
      Top             =   1090
      Width           =   225
   End
   Begin VB.Label lblTecla 
      BackStyle       =   0  'Transparent
      Caption         =   "F7"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00005FB3&
      Height          =   375
      Left            =   2160
      TabIndex        =   1
      Top             =   180
      Width           =   975
   End
End
Attribute VB_Name = "frmBindKey"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
