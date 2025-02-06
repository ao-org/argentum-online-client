VERSION 5.00
Begin VB.Form frmEstadisticas 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   0  'None
   Caption         =   "Estadisticas"
   ClientHeight    =   8655
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7095
   Icon            =   "FrmEstadisticas.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   577
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   473
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Image command1 
      Height          =   300
      Index           =   49
      Left            =   2040
      Tag             =   "0"
      Top             =   720
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.Image command1 
      Height          =   300
      Index           =   48
      Left            =   2520
      Tag             =   "0"
      Top             =   720
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.Image command1 
      Height          =   300
      Index           =   43
      Left            =   5460
      Tag             =   "0"
      Top             =   4545
      Width           =   330
   End
   Begin VB.Image command1 
      Height          =   300
      Index           =   42
      Left            =   6360
      Tag             =   "0"
      Top             =   4545
      Width           =   330
   End
   Begin VB.Image command1 
      Height          =   300
      Index           =   29
      Left            =   1440
      Tag             =   "0"
      Top             =   240
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.Image command1 
      Height          =   300
      Index           =   28
      Left            =   1920
      Tag             =   "0"
      Top             =   240
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.Label text1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "100"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000EA4EB&
      Height          =   210
      Index           =   0
      Left            =   720
      TabIndex        =   27
      Top             =   600
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Label text1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "100"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000EA4EB&
      Height          =   210
      Index           =   22
      Left            =   5955
      TabIndex        =   26
      Top             =   4560
      Width           =   285
   End
   Begin VB.Label text1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "100"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000EA4EB&
      Height          =   210
      Index           =   15
      Left            =   360
      TabIndex        =   25
      Top             =   600
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Label Puntos 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "120"
      ForeColor       =   &H8000000B&
      Height          =   195
      Left            =   4920
      TabIndex        =   24
      Top             =   1380
      Width           =   285
   End
   Begin VB.Label skills 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   150
      Index           =   23
      Left            =   3840
      TabIndex        =   23
      Top             =   7440
      Width           =   1620
   End
   Begin VB.Image command1 
      Height          =   300
      Index           =   46
      Left            =   6360
      Tag             =   "0"
      Top             =   6360
      Width           =   330
   End
   Begin VB.Image command1 
      Height          =   300
      Index           =   47
      Left            =   5460
      Tag             =   "0"
      Top             =   6360
      Width           =   330
   End
   Begin VB.Label text1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "100"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000EA4EB&
      Height          =   210
      Index           =   24
      Left            =   5955
      TabIndex        =   22
      Top             =   6390
      Width           =   285
   End
   Begin VB.Image imgCerrar 
      Height          =   420
      Left            =   6630
      Tag             =   "0"
      Top             =   0
      Width           =   465
   End
   Begin VB.Image command1 
      Height          =   300
      Index           =   45
      Left            =   5460
      Top             =   4155
      Width           =   330
   End
   Begin VB.Image command1 
      Height          =   300
      Index           =   44
      Left            =   6360
      Top             =   4155
      Width           =   330
   End
   Begin VB.Label text1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "100"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000EA4EB&
      Height          =   210
      Index           =   23
      Left            =   5955
      TabIndex        =   21
      Top             =   4200
      Width           =   285
   End
   Begin VB.Image command1 
      Height          =   300
      Index           =   41
      Left            =   5460
      Tag             =   "0"
      Top             =   3375
      Width           =   330
   End
   Begin VB.Image command1 
      Height          =   300
      Index           =   39
      Left            =   5460
      Tag             =   "0"
      Top             =   3765
      Width           =   330
   End
   Begin VB.Image command1 
      Height          =   300
      Index           =   38
      Left            =   6360
      Tag             =   "0"
      Top             =   3765
      Width           =   330
   End
   Begin VB.Image command1 
      Height          =   300
      Index           =   40
      Left            =   6360
      Tag             =   "0"
      Top             =   3375
      Width           =   330
   End
   Begin VB.Label text1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "100"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000EA4EB&
      Height          =   210
      Index           =   21
      Left            =   5955
      TabIndex        =   20
      Top             =   3420
      Width           =   285
   End
   Begin VB.Label text1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "100"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000EA4EB&
      Height          =   210
      Index           =   20
      Left            =   5955
      TabIndex        =   19
      Top             =   3810
      Width           =   285
   End
   Begin VB.Image command1 
      Height          =   300
      Index           =   37
      Left            =   5460
      Tag             =   "0"
      Top             =   3000
      Width           =   330
   End
   Begin VB.Image command1 
      Height          =   300
      Index           =   36
      Left            =   6360
      Tag             =   "0"
      Top             =   3000
      Width           =   330
   End
   Begin VB.Label text1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "100"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000EA4EB&
      Height          =   210
      Index           =   19
      Left            =   5955
      TabIndex        =   18
      Top             =   3045
      Width           =   285
   End
   Begin VB.Label skills 
      BackStyle       =   0  'Transparent
      Height          =   255
      Index           =   0
      Left            =   5040
      TabIndex        =   17
      Top             =   0
      Width           =   615
   End
   Begin VB.Image Image1 
      Height          =   420
      Left            =   2565
      Tag             =   "1"
      Top             =   7890
      Width           =   1980
   End
   Begin VB.Label text1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "100"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000EA4EB&
      Height          =   210
      Index           =   1
      Left            =   2805
      TabIndex        =   16
      Top             =   2265
      Width           =   270
   End
   Begin VB.Label text1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "100"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000EA4EB&
      Height          =   210
      Index           =   2
      Left            =   2805
      TabIndex        =   15
      Top             =   6405
      Width           =   285
   End
   Begin VB.Label text1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "100"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000EA4EB&
      Height          =   210
      Index           =   3
      Left            =   2805
      TabIndex        =   14
      Top             =   4155
      Width           =   285
   End
   Begin VB.Label text1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "100"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000EA4EB&
      Height          =   210
      Index           =   4
      Left            =   2805
      TabIndex        =   13
      Top             =   3765
      Width           =   285
   End
   Begin VB.Label text1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "100"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000EA4EB&
      Height          =   210
      Index           =   5
      Left            =   2805
      TabIndex        =   12
      Top             =   2640
      Width           =   285
   End
   Begin VB.Label text1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "100"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000EA4EB&
      Height          =   210
      Index           =   6
      Left            =   2805
      TabIndex        =   11
      Top             =   5670
      Width           =   285
   End
   Begin VB.Label text1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "100"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000EA4EB&
      Height          =   210
      Index           =   7
      Left            =   5955
      TabIndex        =   10
      Top             =   5625
      Width           =   285
   End
   Begin VB.Label text1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "100"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000EA4EB&
      Height          =   210
      Index           =   8
      Left            =   5955
      TabIndex        =   9
      Top             =   5250
      Width           =   285
   End
   Begin VB.Label text1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "100"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000EA4EB&
      Height          =   210
      Index           =   9
      Left            =   5955
      TabIndex        =   8
      Top             =   6015
      Width           =   285
   End
   Begin VB.Label text1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "100"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000EA4EB&
      Height          =   210
      Index           =   10
      Left            =   2805
      TabIndex        =   7
      Top             =   4545
      Width           =   285
   End
   Begin VB.Label text1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "100"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000EA4EB&
      Height          =   210
      Index           =   11
      Left            =   2805
      TabIndex        =   6
      Top             =   6780
      Width           =   285
   End
   Begin VB.Label text1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "100"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000EA4EB&
      Height          =   210
      Index           =   12
      Left            =   2805
      TabIndex        =   5
      Top             =   5280
      Width           =   270
   End
   Begin VB.Label text1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "100"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000EA4EB&
      Height          =   210
      Index           =   13
      Left            =   2805
      TabIndex        =   4
      Top             =   4905
      Width           =   285
   End
   Begin VB.Label text1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "100"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000EA4EB&
      Height          =   210
      Index           =   14
      Left            =   5955
      TabIndex        =   3
      Top             =   6780
      Width           =   285
   End
   Begin VB.Label text1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "100"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000EA4EB&
      Height          =   210
      Index           =   16
      Left            =   2805
      TabIndex        =   2
      Top             =   3030
      Width           =   285
   End
   Begin VB.Label text1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "100"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000EA4EB&
      Height          =   210
      Index           =   17
      Left            =   5955
      TabIndex        =   1
      Top             =   2265
      Width           =   285
   End
   Begin VB.Label text1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "100"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000EA4EB&
      Height          =   210
      Index           =   18
      Left            =   5955
      TabIndex        =   0
      Top             =   2655
      Width           =   285
   End
   Begin VB.Image command1 
      Height          =   300
      Index           =   0
      Left            =   3225
      Tag             =   "0"
      Top             =   2235
      Width           =   330
   End
   Begin VB.Image command1 
      Height          =   300
      Index           =   2
      Left            =   3225
      Tag             =   "0"
      Top             =   6360
      Width           =   330
   End
   Begin VB.Image command1 
      Height          =   300
      Index           =   4
      Left            =   3225
      Tag             =   "0"
      Top             =   4110
      Width           =   330
   End
   Begin VB.Image command1 
      Height          =   300
      Index           =   6
      Left            =   3225
      Tag             =   "0"
      Top             =   3735
      Width           =   330
   End
   Begin VB.Image command1 
      Height          =   300
      Index           =   8
      Left            =   3225
      Tag             =   "0"
      Top             =   2610
      Width           =   330
   End
   Begin VB.Image command1 
      Height          =   300
      Index           =   10
      Left            =   3225
      Tag             =   "0"
      Top             =   5625
      Width           =   330
   End
   Begin VB.Image command1 
      Height          =   300
      Index           =   12
      Left            =   6360
      Tag             =   "0"
      Top             =   5610
      Width           =   330
   End
   Begin VB.Image command1 
      Height          =   300
      Index           =   14
      Left            =   6360
      Tag             =   "0"
      Top             =   5220
      Width           =   330
   End
   Begin VB.Image command1 
      Height          =   300
      Index           =   16
      Left            =   6360
      Tag             =   "0"
      Top             =   6000
      Width           =   330
   End
   Begin VB.Image command1 
      Height          =   300
      Index           =   18
      Left            =   3225
      Tag             =   "0"
      Top             =   4500
      Width           =   330
   End
   Begin VB.Image command1 
      Height          =   300
      Index           =   20
      Left            =   3225
      Tag             =   "0"
      Top             =   6750
      Width           =   330
   End
   Begin VB.Image command1 
      Height          =   300
      Index           =   22
      Left            =   3225
      Tag             =   "0"
      Top             =   5250
      Width           =   330
   End
   Begin VB.Image command1 
      Height          =   300
      Index           =   24
      Left            =   3225
      Tag             =   "0"
      Top             =   4875
      Width           =   330
   End
   Begin VB.Image command1 
      Height          =   300
      Index           =   26
      Left            =   6360
      Tag             =   "0"
      Top             =   6750
      Width           =   330
   End
   Begin VB.Image command1 
      Height          =   300
      Index           =   30
      Left            =   3225
      Tag             =   "0"
      Top             =   3000
      Width           =   330
   End
   Begin VB.Image command1 
      Height          =   300
      Index           =   32
      Left            =   6360
      Tag             =   "0"
      Top             =   2235
      Width           =   330
   End
   Begin VB.Image command1 
      Height          =   300
      Index           =   34
      Left            =   6360
      Tag             =   "0"
      Top             =   2610
      Width           =   330
   End
   Begin VB.Image command1 
      Height          =   300
      Index           =   3
      Left            =   2325
      Tag             =   "0"
      Top             =   6360
      Width           =   330
   End
   Begin VB.Image command1 
      Height          =   300
      Index           =   5
      Left            =   2325
      Tag             =   "0"
      Top             =   4110
      Width           =   330
   End
   Begin VB.Image command1 
      Height          =   300
      Index           =   7
      Left            =   2325
      Tag             =   "0"
      Top             =   3735
      Width           =   330
   End
   Begin VB.Image command1 
      Height          =   300
      Index           =   9
      Left            =   2325
      Tag             =   "0"
      Top             =   2610
      Width           =   330
   End
   Begin VB.Image command1 
      Height          =   300
      Index           =   11
      Left            =   2325
      Tag             =   "0"
      Top             =   5625
      Width           =   330
   End
   Begin VB.Image command1 
      Height          =   300
      Index           =   13
      Left            =   5460
      Tag             =   "0"
      Top             =   5610
      Width           =   330
   End
   Begin VB.Image command1 
      Height          =   300
      Index           =   15
      Left            =   5460
      Tag             =   "0"
      Top             =   5220
      Width           =   330
   End
   Begin VB.Image command1 
      Height          =   300
      Index           =   17
      Left            =   5460
      Tag             =   "0"
      Top             =   6000
      Width           =   330
   End
   Begin VB.Image command1 
      Height          =   300
      Index           =   19
      Left            =   2325
      Tag             =   "0"
      Top             =   4500
      Width           =   330
   End
   Begin VB.Image command1 
      Height          =   300
      Index           =   21
      Left            =   2325
      Tag             =   "0"
      Top             =   6750
      Width           =   330
   End
   Begin VB.Image command1 
      Height          =   300
      Index           =   23
      Left            =   2325
      Tag             =   "0"
      Top             =   5250
      Width           =   330
   End
   Begin VB.Image command1 
      Height          =   300
      Index           =   25
      Left            =   2325
      Tag             =   "0"
      Top             =   4875
      Width           =   330
   End
   Begin VB.Image command1 
      Height          =   300
      Index           =   27
      Left            =   5460
      Tag             =   "0"
      Top             =   6750
      Width           =   330
   End
   Begin VB.Image command1 
      Height          =   300
      Index           =   31
      Left            =   2325
      Tag             =   "0"
      Top             =   3000
      Width           =   330
   End
   Begin VB.Image command1 
      Height          =   300
      Index           =   33
      Left            =   5460
      Tag             =   "0"
      Top             =   2235
      Width           =   330
   End
   Begin VB.Image command1 
      Height          =   300
      Index           =   35
      Left            =   5460
      Tag             =   "0"
      Top             =   2610
      Width           =   330
   End
   Begin VB.Image command1 
      Height          =   300
      Index           =   1
      Left            =   2325
      Tag             =   "0"
      Top             =   2235
      Width           =   330
   End
End
Attribute VB_Name = "frmEstadisticas"
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
Option Explicit

Public bmoving      As Boolean

Public dX           As Integer

Public dy           As Integer

' Constantes para SendMessage
Const WM_SYSCOMMAND As Long = &H112&

Const MOUSE_MOVE    As Long = &HF012&

Private Declare Function ReleaseCapture Lib "user32" () As Long

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Long) As Long

Private RealizoCambios                As Long

Private PonerloEnRojo(1 To NUMSKILLS) As Boolean

Private cBotonAceptar As clsGraphicalButton
Private cBotonCerrar As clsGraphicalButton


Public Sub Iniciar_Labels()
    
    On Error GoTo Iniciar_Labels_Err
    

    'Iniciamos los labels con los valores de los atributos y los skills
    Dim i As Integer

    For i = 1 To NUMSKILLS
        If UserSkills(i) > 100 Then UserSkills(i) = 100

        Text1(i).Caption = UserSkills(i)
    Next
    
    Exit Sub

Iniciar_Labels_Err:
    Call RegistrarError(Err.number, Err.Description, "frmEstadisticas.Iniciar_Labels", Erl)
    Resume Next
    
End Sub

Private Sub loadButtons()
       
    Set cBotonAceptar = New clsGraphicalButton
    Set cBotonCerrar = New clsGraphicalButton


    Call cBotonAceptar.Initialize(Image1, "boton-aceptar-default.bmp", _
                                                "boton-aceptar-over.bmp", _
                                                "boton-aceptar-off.bmp", Me)
                                                
                                                
    Call cBotonCerrar.Initialize(imgCerrar, "boton-cerrar-default.bmp", _
                                                "boton-cerrar-over.bmp", _
                                                "boton-cerrar-off.bmp", Me)
    
End Sub

Private Sub Command1_Click(Index As Integer)
    
    On Error GoTo Command1_Click_Err
    
    Dim indice

    Dim skilloriginal

    indice = Index \ 2 + 1

    If (Index And &H1) = 0 Then
        If Alocados > 0 Then
            indice = Index \ 2 + 1

            If indice > NUMSKILLS Then indice = NUMSKILLS
            If Val(Text1(indice).Caption) < MAXSKILLPOINTS Then
                Text1(indice).Caption = Val(Text1(indice).Caption) + 1
                flags(indice) = flags(indice) + 1
                Alocados = Alocados - 1
                RealizoCambios = RealizoCambios + 1

            End If
            
        End If

    Else

        If Alocados < SkillPoints Then
        
            indice = Index \ 2 + 1

            If Val(Text1(indice).Caption) > 0 And flags(indice) > 0 Then
                Text1(indice).Caption = Val(Text1(indice).Caption) - 1
                flags(indice) = flags(indice) - 1
                Alocados = Alocados + 1
                RealizoCambios = RealizoCambios - 1

            End If

        End If

    End If

    puntos.Caption = Alocados

    Dim ladder As Byte

    ladder = Val(Text1(indice).Caption)

    If UserSkills(indice) < ladder Then
        Text1(indice).ForeColor = vbRed
        PonerloEnRojo(indice) = True

    End If

    If UserSkills(indice) = ladder Then
        Text1(indice).ForeColor = &H40C0&
        RealizoCambios = RealizoCambios - 1
        PonerloEnRojo(indice) = False

    End If

    
    Exit Sub

Command1_Click_Err:
    Call RegistrarError(Err.number, Err.Description, "frmEstadisticas.Command1_Click", Erl)
    Resume Next
    
End Sub

Private Sub Command2_Click()
    
    On Error GoTo Command2_Click_Err
    
    Unload Me

    
    Exit Sub

Command2_Click_Err:
    Call RegistrarError(Err.number, Err.Description, "frmEstadisticas.Command2_Click", Erl)
    Resume Next
    
End Sub

Private Sub command1_DblClick(Index As Integer)
    Command1_Click (Index)
End Sub

Private Sub Command1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    
    On Error GoTo Command1_MouseDown_Err
    
    Set Command1(Index).Picture = LoadInterface(IIf(Index Mod 2 = 1, "boton-sm-flecha-izq-off.bmp", "boton-sm-flecha-der-off.bmp"))
    Command1(Index).Tag = "1"
    
    Exit Sub

Command1_MouseDown_Err:
    Call RegistrarError(Err.number, Err.Description, "frmEstadisticas.Command1_MouseDown", Erl)
    Resume Next
    
End Sub

Private Sub Command1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    
    On Error GoTo Command1_MouseMove_Err
    
    If Command1(Index).Tag = "0" Then
        Set Command1(Index).Picture = LoadInterface(IIf(Index Mod 2 = 1, "boton-sm-flecha-izq-over.bmp", "boton-sm-flecha-der-over.bmp"))
        Command1(Index).Tag = "1"
    End If

    Exit Sub

Command1_MouseMove_Err:
    Call RegistrarError(Err.number, Err.Description, "frmEstadisticas.Command1_MouseMove", Erl)
    Resume Next
    
End Sub

Private Sub command1_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    
    On Error GoTo command1_MouseUp_Err
    
    Set Command1(Index) = Nothing
    Command1(Index).Tag = "0"
    
    Exit Sub

command1_MouseUp_Err:
    Call RegistrarError(Err.number, Err.Description, "frmEstadisticas.command1_MouseUp", Erl)
    Resume Next
    
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    
    On Error GoTo Form_KeyDown_Err
    
    If KeyCode = 27 Then
        Unload Me

    End If

    
    Exit Sub

Form_KeyDown_Err:
    Call RegistrarError(Err.number, Err.Description, "frmEstadisticas.Form_KeyDown", Erl)
    Resume Next
    
End Sub

Private Sub Form_Load()
    
    On Error GoTo Form_Load_Err
    
    Call Aplicar_Transparencia(Me.hwnd, 240)
    
    Call FormParser.Parse_Form(Me)
    'Image1.Picture = LoadInterface("botonlargoaceptar.bmp")
    RealizoCambios = 0
    ReDim flags(1 To NUMSKILLS)
    
    Call loadButtons
    
    Exit Sub

Form_Load_Err:
    Call RegistrarError(Err.number, Err.Description, "frmEstadisticas.Form_Load", Erl)
    Resume Next
    
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    On Error GoTo Form_MouseMove_Err
    
    MoverForm Me.hwnd
    'If Image1.Tag = "1" Then
    ' Image1.Picture = LoadInterface("botonlargoaceptar.bmp")
    '    Image1.Tag = "0"
    'End If
    Image1.Picture = Nothing
    
    Dim A As Integer

    For A = 1 To NUMSKILLS

        If Not PonerloEnRojo(A) Then
            Text1(A).ForeColor = &HEA4EB

            'Skills(a).ForeColor = vbWhite
        End If

        If PonerloEnRojo(A) = True Then
            Text1(A).ForeColor = vbRed

        End If

    Next A
    
    Exit Sub

Form_MouseMove_Err:
    Call RegistrarError(Err.number, Err.Description, "frmEstadisticas.Form_MouseMove", Erl)
    Resume Next
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    
    On Error GoTo Form_QueryUnload_Err
    
    Unload Me

    
    Exit Sub

Form_QueryUnload_Err:
    Call RegistrarError(Err.number, Err.Description, "frmEstadisticas.Form_QueryUnload", Erl)
    Resume Next
    
End Sub

Private Sub Image1_Click()
    
    On Error GoTo Image1_Click_Err
    

    If RealizoCambios >= 1 Then
        If MsgBox(JsonLanguage.Item("MENSAJEBOX_GUARDAR_SKILLPOINTS"), vbYesNo) = vbYes Then

            Dim skillChanges(NUMSKILLS) As Byte

            Dim i                       As Long

            For i = 1 To NUMSKILLS
                skillChanges(i) = CByte(Text1(i).Caption) - UserSkills(i)
                'Actualizamos nuestros datos locales
                UserSkills(i) = Val(Text1(i).Caption)
        
            Next i
    
            Call WriteModifySkills(skillChanges())
    
            SkillPoints = Alocados
            Unload Me

        End If

    End If

    Unload Me
    
    For i = 1 To NUMSKILLS
        PonerloEnRojo(i) = False
    Next i

    
    Exit Sub

Image1_Click_Err:
    Call RegistrarError(Err.number, Err.Description, "frmEstadisticas.Image1_Click", Erl)
    Resume Next
    
End Sub


Private Sub imgCerrar_Click()
    
    On Error GoTo imgCerrar_Click_Err
    
    If RealizoCambios >= 1 Then
        If MsgBox(JsonLanguage.Item("MENSAJEBOX_CAMBIOS_SKILLPOINTS"), vbYesNo) = vbYes Then

            Dim skillChanges(NUMSKILLS) As Byte

            Dim i                       As Long

            For i = 1 To NUMSKILLS
                skillChanges(i) = CByte(Text1(i).Caption) - UserSkills(i)
                'Actualizamos nuestros datos locales
                UserSkills(i) = Val(Text1(i).Caption)
            Next i
    
            Call WriteModifySkills(skillChanges())
    
            SkillPoints = Alocados
            Unload Me

        End If

    End If
    
    For i = 1 To NUMSKILLS
        PonerloEnRojo(i) = False
    Next i

    Unload Me

    
    Exit Sub

imgCerrar_Click_Err:
    Call RegistrarError(Err.number, Err.Description, "frmEstadisticas.imgCerrar_Click", Erl)
    Resume Next
    
End Sub



Private Sub skills_Click(Index As Integer)
    
    On Error GoTo skills_Click_Err
    
    AddtoRichTextBox frmMain.RecTxt, JsonLanguage.Item("MENSAJE_INFORMACION_DE_SKILL") & SkillsDesc(Index), 2, 51, 223, 1, 1

    Exit Sub

skills_Click_Err:
    Call RegistrarError(Err.number, Err.Description, "frmEstadisticas.skills_Click", Erl)
    Resume Next
    
End Sub

Private Sub Skills_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    
    On Error GoTo Skills_MouseMove_Err
    

    Dim A As Integer

    For A = 1 To NUMSKILLS

        If Not PonerloEnRojo(A) Then
            Text1(A).ForeColor = &HEA4EB

        End If

        'Skills(a).ForeColor = vbWhite
    Next A

    Text1(Index).ForeColor = vbBlue

    'Skills(index).ForeColor = vbBlue
    
    Exit Sub

Skills_MouseMove_Err:
    Call RegistrarError(Err.number, Err.Description, "frmEstadisticas.Skills_MouseMove", Erl)
    Resume Next
    
End Sub

