VERSION 5.00
Begin VB.Form Manual 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Manual de Argentum"
   ClientHeight    =   6780
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6225
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6780
   ScaleWidth      =   6225
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Indice general"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6735
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   6015
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         Caption         =   "Clanes"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   1800
         TabIndex        =   13
         Top             =   5040
         Width           =   2175
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         Caption         =   "Quest"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   1800
         TabIndex        =   12
         Top             =   4680
         Width           =   2175
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         Caption         =   "Skills"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   1800
         TabIndex        =   11
         Top             =   4320
         Width           =   2175
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         Caption         =   "SECCION ACTUALMENTE EN CONSTRUCCIÓN"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   360
         TabIndex        =   10
         Top             =   5520
         Width           =   5535
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         Caption         =   "Herreria"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   1800
         TabIndex        =   9
         Top             =   3960
         Width           =   2175
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         Caption         =   "Pociones"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   2160
         TabIndex        =   8
         Top             =   3600
         Width           =   1455
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         Caption         =   "Alquimia"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   2160
         TabIndex        =   7
         Top             =   3240
         Width           =   1455
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "Atributos"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   2160
         TabIndex        =   6
         Top             =   2520
         Width           =   1455
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Caption         =   "Razas"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   2160
         TabIndex        =   5
         Top             =   2160
         Width           =   1455
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Categorias"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   0
         TabIndex        =   4
         Top             =   1800
         Width           =   5775
      End
      Begin VB.Label Label2 
         Caption         =   $"Manual.frx":0000
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   120
         TabIndex        =   3
         Top             =   600
         Width           =   5775
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Conceptos Basicos"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   5775
      End
      Begin VB.Label Label21 
         Alignment       =   2  'Center
         Caption         =   "Mineria y Herreria"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   1800
         TabIndex        =   1
         Top             =   2880
         Width           =   2175
      End
   End
End
Attribute VB_Name = "Manual"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Picture1_Click()

End Sub

Private Sub Picture8_Click()

End Sub

Private Sub Label10_Click()
    
    On Error GoTo Label10_Click_Err
    
    Unload Me
    manual_skill.Show , frmMain

    
    Exit Sub

Label10_Click_Err:
    Call RegistrarError(Err.number, Err.Description, "Manual.Label10_Click", Erl)
    Resume Next
    
End Sub

Private Sub Label11_Click()
    
    On Error GoTo Label11_Click_Err
    
    Unload Me
    Manual_Quest.Show , frmMain

    
    Exit Sub

Label11_Click_Err:
    Call RegistrarError(Err.number, Err.Description, "Manual.Label11_Click", Erl)
    Resume Next
    
End Sub

Private Sub Label12_Click()
    
    On Error GoTo Label12_Click_Err
    
    Unload Me
    Manual_Clanes.Show , frmMain

    
    Exit Sub

Label12_Click_Err:
    Call RegistrarError(Err.number, Err.Description, "Manual.Label12_Click", Erl)
    Resume Next
    
End Sub

Private Sub Label21_Click()
    
    On Error GoTo Label21_Click_Err
    
    Unload Me
    Manual_Mineria.Show , frmMain

    
    Exit Sub

Label21_Click_Err:
    Call RegistrarError(Err.number, Err.Description, "Manual.Label21_Click", Erl)
    Resume Next
    
End Sub

Private Sub Label4_Click()
    
    On Error GoTo Label4_Click_Err
    
    Unload Me
    Manual_Razas.Show , frmMain

    
    Exit Sub

Label4_Click_Err:
    Call RegistrarError(Err.number, Err.Description, "Manual.Label4_Click", Erl)
    Resume Next
    
End Sub

Private Sub Label5_Click()
    
    On Error GoTo Label5_Click_Err
    
    Unload Me
    Manual_Atributos.Show , frmMain

    
    Exit Sub

Label5_Click_Err:
    Call RegistrarError(Err.number, Err.Description, "Manual.Label5_Click", Erl)
    Resume Next
    
End Sub

Private Sub Label6_Click()
    
    On Error GoTo Label6_Click_Err
    
    Unload Me
    Manual_Alquimia.Show , frmMain

    
    Exit Sub

Label6_Click_Err:
    Call RegistrarError(Err.number, Err.Description, "Manual.Label6_Click", Erl)
    Resume Next
    
End Sub

Private Sub Label7_Click()
    
    On Error GoTo Label7_Click_Err
    
    Unload Me
    Manual_Pociones.Show , frmMain

    
    Exit Sub

Label7_Click_Err:
    Call RegistrarError(Err.number, Err.Description, "Manual.Label7_Click", Erl)
    Resume Next
    
End Sub

Private Sub Label8_Click()
    
    On Error GoTo Label8_Click_Err
    
    Unload Me
    Manual_Herreria.Show , frmMain

    
    Exit Sub

Label8_Click_Err:
    Call RegistrarError(Err.number, Err.Description, "Manual.Label8_Click", Erl)
    Resume Next
    
End Sub
