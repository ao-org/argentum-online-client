VERSION 5.00
Begin VB.Form frmEstadisticas 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   0  'None
   Caption         =   "Estadisticas"
   ClientHeight    =   7110
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8115
   Icon            =   "FrmEstadisticas.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   474
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   541
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label lblpuntosbattle 
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
      ForeColor       =   &H000040C0&
      Height          =   195
      Left            =   2400
      TabIndex        =   64
      Top             =   4560
      Width           =   975
   End
   Begin VB.Label lblcredito 
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
      ForeColor       =   &H000040C0&
      Height          =   255
      Left            =   6600
      TabIndex        =   63
      Top             =   2025
      Width           =   855
   End
   Begin VB.Label lbldiasrestantes 
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
      ForeColor       =   &H000040C0&
      Height          =   255
      Left            =   5940
      TabIndex        =   62
      Top             =   1800
      Width           =   1335
   End
   Begin VB.Label lbldonador 
      BackStyle       =   0  'Transparent
      Caption         =   "Activo"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   255
      Left            =   5520
      TabIndex        =   61
      Top             =   1575
      Width           =   855
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   12
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   255
      Left            =   9360
      TabIndex        =   60
      Top             =   240
      Width           =   255
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Neutral"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   195
      Index           =   9
      Left            =   1320
      TabIndex        =   59
      Top             =   4350
      UseMnemonic     =   0   'False
      Width           =   2340
   End
   Begin VB.Label Fami 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "10"
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   6
         Charset         =   255
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   120
      Index           =   1
      Left            =   5160
      TabIndex        =   58
      Top             =   8580
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label Fami 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Ladder"
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   6
         Charset         =   255
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   0
      Left            =   3750
      TabIndex        =   57
      Top             =   8580
      Visible         =   0   'False
      Width           =   945
   End
   Begin VB.Label Fami 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0/0"
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   6
         Charset         =   255
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   165
      Index           =   4
      Left            =   3480
      TabIndex        =   56
      Top             =   10080
      Visible         =   0   'False
      Width           =   840
   End
   Begin VB.Label Fami 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Ninguna"
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   6
         Charset         =   255
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Index           =   5
      Left            =   3960
      TabIndex        =   55
      Top             =   10560
      Visible         =   0   'False
      Width           =   900
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
      Index           =   18
      Left            =   5040
      TabIndex        =   52
      Top             =   6000
      Width           =   1470
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
      Index           =   17
      Left            =   5160
      TabIndex        =   51
      Top             =   5820
      Width           =   1350
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
      Index           =   16
      Left            =   4680
      TabIndex        =   50
      Top             =   5640
      Width           =   1830
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
      Index           =   15
      Left            =   5280
      TabIndex        =   49
      Top             =   5460
      Width           =   1230
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
      Index           =   14
      Left            =   5280
      TabIndex        =   48
      Top             =   5280
      Width           =   1230
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
      Index           =   13
      Left            =   4920
      TabIndex        =   47
      Top             =   5100
      Width           =   1590
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
      Index           =   12
      Left            =   4800
      TabIndex        =   46
      Top             =   4920
      Width           =   1710
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
      Index           =   11
      Left            =   5400
      TabIndex        =   45
      Top             =   4740
      Width           =   1110
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
      Index           =   10
      Left            =   4440
      TabIndex        =   44
      Top             =   4560
      Width           =   2070
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
      Index           =   9
      Left            =   5520
      TabIndex        =   43
      Top             =   4380
      Width           =   990
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
      Index           =   8
      Left            =   5040
      TabIndex        =   42
      Top             =   4200
      Width           =   1470
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
      Index           =   7
      Left            =   5400
      TabIndex        =   41
      Top             =   4020
      Width           =   1110
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
      Index           =   6
      Left            =   5520
      TabIndex        =   40
      Top             =   3840
      Width           =   990
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
      Index           =   5
      Left            =   5640
      TabIndex        =   39
      Top             =   3660
      Width           =   870
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
      Index           =   4
      Left            =   4560
      TabIndex        =   38
      Top             =   3480
      Width           =   1950
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
      Index           =   3
      Left            =   4440
      TabIndex        =   37
      Top             =   3300
      Width           =   2070
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
      Index           =   2
      Left            =   5880
      TabIndex        =   36
      Top             =   3120
      Width           =   630
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
      Index           =   1
      Left            =   5880
      TabIndex        =   35
      Top             =   2925
      Width           =   630
   End
   Begin VB.Label skills 
      BackStyle       =   0  'Transparent
      Height          =   255
      Index           =   0
      Left            =   6360
      TabIndex        =   34
      Top             =   120
      Width           =   615
   End
   Begin VB.Image Image1 
      Height          =   465
      Left            =   3225
      Tag             =   "1"
      Top             =   6495
      Width           =   1785
   End
   Begin VB.Label Label6 
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
      ForeColor       =   &H000040C0&
      Height          =   195
      Index           =   7
      Left            =   2190
      TabIndex        =   33
      Top             =   5505
      Width           =   705
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "10 min"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   195
      Index           =   5
      Left            =   2310
      TabIndex        =   32
      Top             =   5940
      Width           =   705
   End
   Begin VB.Label Label6 
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
      ForeColor       =   &H000040C0&
      Height          =   195
      Index           =   0
      Left            =   2760
      TabIndex        =   31
      Top             =   4140
      UseMnemonic     =   0   'False
      Width           =   555
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
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
      Height          =   195
      Index           =   5
      Left            =   720
      TabIndex        =   30
      Top             =   720
      Visible         =   0   'False
      Width           =   675
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
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
      Height          =   195
      Index           =   2
      Left            =   720
      TabIndex        =   29
      Top             =   360
      Visible         =   0   'False
      Width           =   675
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Hombre"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   195
      Index           =   6
      Left            =   1380
      TabIndex        =   28
      Top             =   1920
      Width           =   1380
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Elfo"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   195
      Index           =   8
      Left            =   1095
      TabIndex        =   27
      Top             =   1710
      Width           =   900
   End
   Begin VB.Label puntos 
      AutoSize        =   -1  'True
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
      ForeColor       =   &H000040C0&
      Height          =   210
      Left            =   1980
      TabIndex        =   26
      Top             =   5730
      Width           =   90
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
      ForeColor       =   &H000040C0&
      Height          =   210
      Index           =   1
      Left            =   6720
      TabIndex        =   25
      Top             =   2910
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
      ForeColor       =   &H000040C0&
      Height          =   210
      Index           =   2
      Left            =   6720
      TabIndex        =   24
      Top             =   3090
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
      ForeColor       =   &H000040C0&
      Height          =   210
      Index           =   3
      Left            =   6720
      TabIndex        =   23
      Top             =   3270
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
      ForeColor       =   &H000040C0&
      Height          =   210
      Index           =   4
      Left            =   6720
      TabIndex        =   22
      Top             =   3465
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
      ForeColor       =   &H000040C0&
      Height          =   210
      Index           =   5
      Left            =   6720
      TabIndex        =   21
      Top             =   3645
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
      ForeColor       =   &H000040C0&
      Height          =   210
      Index           =   6
      Left            =   6720
      TabIndex        =   20
      Top             =   3810
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
      ForeColor       =   &H000040C0&
      Height          =   210
      Index           =   7
      Left            =   6720
      TabIndex        =   19
      Top             =   3990
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
      ForeColor       =   &H000040C0&
      Height          =   210
      Index           =   8
      Left            =   6720
      TabIndex        =   18
      Top             =   4170
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
      ForeColor       =   &H000040C0&
      Height          =   210
      Index           =   9
      Left            =   6720
      TabIndex        =   17
      Top             =   4350
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
      ForeColor       =   &H000040C0&
      Height          =   210
      Index           =   10
      Left            =   6720
      TabIndex        =   16
      Top             =   4530
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
      ForeColor       =   &H000040C0&
      Height          =   210
      Index           =   11
      Left            =   6720
      TabIndex        =   15
      Top             =   4710
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
      ForeColor       =   &H000040C0&
      Height          =   210
      Index           =   12
      Left            =   6720
      TabIndex        =   14
      Top             =   4890
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
      ForeColor       =   &H000040C0&
      Height          =   210
      Index           =   13
      Left            =   6720
      TabIndex        =   13
      Top             =   5070
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
      ForeColor       =   &H000040C0&
      Height          =   210
      Index           =   14
      Left            =   6720
      TabIndex        =   12
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
      ForeColor       =   &H000040C0&
      Height          =   210
      Index           =   15
      Left            =   6720
      TabIndex        =   11
      Top             =   5430
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
      ForeColor       =   &H000040C0&
      Height          =   210
      Index           =   16
      Left            =   6720
      TabIndex        =   10
      Top             =   5610
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
      ForeColor       =   &H000040C0&
      Height          =   210
      Index           =   17
      Left            =   6720
      TabIndex        =   9
      Top             =   5790
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
      ForeColor       =   &H000040C0&
      Height          =   210
      Index           =   18
      Left            =   6720
      TabIndex        =   8
      Top             =   5970
      Width           =   285
   End
   Begin VB.Label text1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   150
      Index           =   0
      Left            =   8400
      TabIndex        =   7
      Top             =   240
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.Image command1 
      Height          =   180
      Index           =   0
      Left            =   7035
      Top             =   2925
      Width           =   210
   End
   Begin VB.Image command1 
      Height          =   180
      Index           =   2
      Left            =   7035
      Top             =   3105
      Width           =   210
   End
   Begin VB.Image command1 
      Height          =   180
      Index           =   4
      Left            =   7035
      Top             =   3285
      Width           =   210
   End
   Begin VB.Image command1 
      Height          =   180
      Index           =   6
      Left            =   7035
      Top             =   3465
      Width           =   210
   End
   Begin VB.Image command1 
      Height          =   180
      Index           =   8
      Left            =   7035
      Top             =   3645
      Width           =   210
   End
   Begin VB.Image command1 
      Height          =   180
      Index           =   10
      Left            =   7035
      Top             =   3825
      Width           =   210
   End
   Begin VB.Image command1 
      Height          =   180
      Index           =   12
      Left            =   7035
      Top             =   4005
      Width           =   210
   End
   Begin VB.Image command1 
      Height          =   180
      Index           =   14
      Left            =   7035
      Top             =   4185
      Width           =   210
   End
   Begin VB.Image command1 
      Height          =   180
      Index           =   16
      Left            =   7035
      Top             =   4365
      Width           =   210
   End
   Begin VB.Image command1 
      Height          =   180
      Index           =   18
      Left            =   7035
      Top             =   4545
      Width           =   210
   End
   Begin VB.Image command1 
      Height          =   180
      Index           =   20
      Left            =   7035
      Top             =   4725
      Width           =   210
   End
   Begin VB.Image command1 
      Height          =   180
      Index           =   22
      Left            =   7035
      Top             =   4905
      Width           =   210
   End
   Begin VB.Image command1 
      Height          =   180
      Index           =   24
      Left            =   7035
      Top             =   5085
      Width           =   210
   End
   Begin VB.Image command1 
      Height          =   180
      Index           =   26
      Left            =   7035
      Top             =   5265
      Width           =   210
   End
   Begin VB.Image command1 
      Height          =   180
      Index           =   28
      Left            =   7035
      Top             =   5445
      Width           =   210
   End
   Begin VB.Image command1 
      Height          =   180
      Index           =   30
      Left            =   7035
      Top             =   5625
      Width           =   210
   End
   Begin VB.Image command1 
      Height          =   180
      Index           =   32
      Left            =   7035
      Top             =   5805
      Width           =   210
   End
   Begin VB.Image command1 
      Height          =   180
      Index           =   34
      Left            =   7035
      Top             =   5985
      Width           =   210
   End
   Begin VB.Image command1 
      Height          =   180
      Index           =   3
      Left            =   6540
      Top             =   3105
      Width           =   210
   End
   Begin VB.Image command1 
      Height          =   180
      Index           =   5
      Left            =   6540
      Top             =   3285
      Width           =   210
   End
   Begin VB.Image command1 
      Height          =   180
      Index           =   7
      Left            =   6540
      Top             =   3465
      Width           =   210
   End
   Begin VB.Image command1 
      Height          =   180
      Index           =   9
      Left            =   6540
      Top             =   3645
      Width           =   210
   End
   Begin VB.Image command1 
      Height          =   180
      Index           =   11
      Left            =   6540
      Top             =   3825
      Width           =   210
   End
   Begin VB.Image command1 
      Height          =   180
      Index           =   13
      Left            =   6540
      Top             =   4005
      Width           =   210
   End
   Begin VB.Image command1 
      Height          =   180
      Index           =   15
      Left            =   6540
      Top             =   4185
      Width           =   210
   End
   Begin VB.Image command1 
      Height          =   180
      Index           =   17
      Left            =   6540
      Top             =   4365
      Width           =   210
   End
   Begin VB.Image command1 
      Height          =   180
      Index           =   19
      Left            =   6540
      Top             =   4545
      Width           =   210
   End
   Begin VB.Image command1 
      Height          =   180
      Index           =   21
      Left            =   6540
      Top             =   4725
      Width           =   210
   End
   Begin VB.Image command1 
      Height          =   180
      Index           =   23
      Left            =   6540
      Top             =   4905
      Width           =   210
   End
   Begin VB.Image command1 
      Height          =   180
      Index           =   25
      Left            =   6540
      Top             =   5085
      Width           =   210
   End
   Begin VB.Image command1 
      Height          =   180
      Index           =   27
      Left            =   6540
      Top             =   5265
      Width           =   210
   End
   Begin VB.Image command1 
      Height          =   180
      Index           =   29
      Left            =   6540
      Top             =   5445
      Width           =   210
   End
   Begin VB.Image command1 
      Height          =   180
      Index           =   31
      Left            =   6540
      Top             =   5625
      Width           =   210
   End
   Begin VB.Image command1 
      Height          =   180
      Index           =   33
      Left            =   6540
      Top             =   5805
      Width           =   210
   End
   Begin VB.Image command1 
      Height          =   180
      Index           =   35
      Left            =   6540
      Top             =   5985
      Width           =   210
   End
   Begin VB.Image command1 
      Height          =   180
      Index           =   1
      Left            =   6540
      Top             =   2925
      Width           =   210
   End
   Begin VB.Image estado 
      Height          =   390
      Left            =   7440
      Top             =   9840
      Visible         =   0   'False
      Width           =   1425
   End
   Begin VB.Label Label6 
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
      ForeColor       =   &H000040C0&
      Height          =   195
      Index           =   1
      Left            =   2910
      TabIndex        =   6
      Top             =   3930
      UseMnemonic     =   0   'False
      Width           =   555
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Paladin"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   195
      Index           =   4
      Left            =   1155
      TabIndex        =   5
      Top             =   1515
      Width           =   2220
   End
   Begin VB.Label Atri 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "15"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   210
      Index           =   4
      Left            =   1935
      TabIndex        =   4
      Top             =   3240
      Width           =   180
   End
   Begin VB.Label Atri 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "15"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   210
      Index           =   3
      Left            =   1875
      TabIndex        =   3
      Top             =   3030
      Width           =   180
   End
   Begin VB.Label Atri 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "15"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   210
      Index           =   2
      Left            =   1515
      TabIndex        =   2
      Top             =   2805
      Width           =   180
   End
   Begin VB.Label Atri 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "15"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   210
      Index           =   1
      Left            =   1290
      TabIndex        =   1
      Top             =   2625
      Width           =   180
   End
   Begin VB.Label Label6 
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
      ForeColor       =   &H000040C0&
      Height          =   195
      Index           =   3
      Left            =   2475
      TabIndex        =   0
      Top             =   5295
      WhatsThisHelpID =   8000
      Width           =   705
   End
   Begin VB.Label Fami 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0/0"
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   6
         Charset         =   255
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   165
      Index           =   2
      Left            =   4080
      TabIndex        =   54
      Top             =   8850
      Visible         =   0   'False
      Width           =   840
   End
   Begin VB.Label Fami 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0/0"
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   6
         Charset         =   255
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   135
      Index           =   3
      Left            =   5160
      TabIndex        =   53
      Top             =   10155
      Visible         =   0   'False
      Width           =   480
   End
End
Attribute VB_Name = "frmEstadisticas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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

Private Sub moverForm()

    Dim res As Long

    ReleaseCapture
    res = SendMessage(Me.hwnd, WM_SYSCOMMAND, MOUSE_MOVE, 0)

End Sub

Public Sub Iniciar_Labels()

    'Iniciamos los labels con los valores de los atributos y los skills
    Dim i As Integer

    For i = 1 To NUMATRIBUTOS 'Colocado
        Atri(i).Caption = UserAtributos(i)
    Next

    For i = 1 To NUMSKILLS

        If UserSkills(i) > 100 Then
            UserSkills(i) = 100

        End If

        ' Skills(I).Caption = SkillsNames(I) & ":"
        Text1(i).Caption = UserSkills(i)
    Next

    Select Case UserEstadisticas.Alineacion

        Case 0
            Label6(9).Caption = "Neutral"
            Label6(9).ForeColor = RGB(127, 127, 127)

        Case 1
            Label6(9).Caption = "Ciudadano"
            Label6(9).ForeColor = RGB(0, 128, 255)

        Case 2
            Label6(9).Caption = "Caos"
            Label6(9).ForeColor = RGB(128, 0, 0)

        Case 3
            Label6(9).Caption = "Imperial"
            Label6(9).ForeColor = RGB(33, 133, 132)
        
        Case Else
            Label6(9).Caption = "Desconocido"

    End Select

    'estado = LoadInterface("estadisticascriminal.bmp")
    'Else
    '   Label6(9).Caption = "Ciudadano"
    'estado = LoadInterface("estadisticasciudadano.bmp")
    'End If

    With UserEstadisticas

        Label6(0).Caption = .CriminalesMatados 'Colocado
        Label6(1).Caption = .CiudadanosMatados 'Colocado
        Label6(3).Caption = .NpcsMatados
        Label6(4).Caption = .Clase 'Colocado
        Label6(5).Caption = .PenaCarcel & " min"
        Label6(6).Caption = .Genero
        Label6(7).Caption = .VecesQueMoriste
        Label6(8).Caption = .Raza

        If .Donador = 0 Then
            lbldonador.Caption = "Inactivo"
        Else
            lbldonador.Caption = "Activo"

        End If

        'lbldonador.Caption = .Donador
        lbldiasrestantes.Caption = .DiasRestantes
        lblcredito.Caption = .CreditoDonador
        lblpuntosbattle.Caption = .BattlePuntos
    
    End With

End Sub

Private Sub Command1_Click(Index As Integer)

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

End Sub

Private Sub Command2_Click()
    Unload Me

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 27 Then
        Unload Me

    End If

End Sub

Private Sub Form_Load()
    Call FormParser.Parse_Form(Me)
    'Image1.Picture = LoadInterface("botonlargoaceptar.bmp")
    RealizoCambios = 0
    ReDim flags(1 To NUMSKILLS)

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    moverForm
    'If Image1.Tag = "1" Then
    ' Image1.Picture = LoadInterface("botonlargoaceptar.bmp")
    '    Image1.Tag = "0"
    'End If
    Image1.Picture = Nothing

    Dim a As Integer

    For a = 1 To NUMSKILLS

        If Not PonerloEnRojo(a) Then
            Text1(a).ForeColor = &H40C0&

            'Skills(a).ForeColor = vbWhite
        End If

        If PonerloEnRojo(a) = True Then
            Text1(a).ForeColor = vbRed

        End If

    Next a

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Unload Me

End Sub

Private Sub Image1_Click()

    If RealizoCambios >= 1 Then
        If MsgBox("Realizo cambios en sus skillpoints ¿desea guardar antes de salir?", vbYesNo) = vbYes Then

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

End Sub

Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

    'Image1 = LoadInterface("aceptarpress.bmp")
End Sub

Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

    Image1.Picture = LoadInterface("estadisticaswidehover.bmp")

End Sub

Private Sub Label1_Click()

    If RealizoCambios >= 1 Then
        If MsgBox("Realizo cambios en sus skillpoints ¿desea guardar antes de salir?", vbYesNo) = vbYes Then

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

End Sub

Private Sub skills_Click(Index As Integer)
    AddtoRichTextBox frmmain.RecTxt, "Información del skill> " & SkillsDesc(Index), 2, 51, 223, 1, 1

End Sub

Private Sub Skills_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)

    Dim a As Integer

    For a = 1 To NUMSKILLS

        If Not PonerloEnRojo(a) Then
            Text1(a).ForeColor = &H40C0&

        End If

        'Skills(a).ForeColor = vbWhite
    Next a

    Text1(Index).ForeColor = vbBlue

    'Skills(index).ForeColor = vbBlue
End Sub

