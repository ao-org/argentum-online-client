VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Tienda AO20"
   ClientHeight    =   4575
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5370
   LinkTopic       =   "Form1"
   ScaleHeight     =   4575
   ScaleWidth      =   5370
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command2 
      Caption         =   "Cancelar"
      Height          =   495
      Left            =   240
      TabIndex        =   7
      Top             =   3960
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Comprar"
      Height          =   495
      Left            =   3960
      TabIndex        =   6
      Top             =   3960
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Height          =   3070
      Left            =   3240
      TabIndex        =   5
      Top             =   635
      Width           =   1935
   End
   Begin VB.ListBox frmShop 
      Height          =   2985
      Left            =   240
      TabIndex        =   4
      Top             =   720
      Width           =   2775
   End
   Begin VB.PictureBox Picture1 
      Height          =   495
      Left            =   3360
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   3
      Top             =   120
      Width           =   495
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   120
      Width           =   3015
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "12"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   4350
      TabIndex        =   2
      Top             =   390
      Width           =   285
   End
   Begin VB.Label Label1 
      Caption         =   "Creditos:"
      Height          =   495
      Left            =   4200
      TabIndex        =   1
      Top             =   120
      Width           =   855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub List1_Click()

End Sub
