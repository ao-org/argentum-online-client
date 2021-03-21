VERSION 5.00
Begin VB.Form frmEstadisticas 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   0  'None
   Caption         =   "Estadisticas"
   ClientHeight    =   8265
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10350
   Icon            =   "FrmEstadisticas.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   551
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   690
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
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
      Left            =   6930
      TabIndex        =   76
      Top             =   2700
      Width           =   1620
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
      Left            =   6930
      TabIndex        =   75
      Top             =   2325
      Width           =   1620
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
      Left            =   3675
      TabIndex        =   74
      Top             =   3090
      Width           =   1620
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
      ForeColor       =   &H000EA4EB&
      Height          =   210
      Index           =   5
      Left            =   1920
      TabIndex        =   73
      Top             =   5310
      Width           =   180
   End
   Begin VB.Image command1 
      Height          =   300
      Index           =   46
      Left            =   9615
      Tag             =   "0"
      Top             =   6375
      Width           =   330
   End
   Begin VB.Image command1 
      Height          =   300
      Index           =   47
      Left            =   8715
      Tag             =   "0"
      Top             =   6375
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
      Left            =   9195
      TabIndex        =   72
      Top             =   6390
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
      Index           =   24
      Left            =   6960
      TabIndex        =   71
      Top             =   6465
      Width           =   1620
   End
   Begin VB.Image imgCerrar 
      Height          =   420
      Left            =   9885
      Tag             =   "0"
      Top             =   0
      Width           =   465
   End
   Begin VB.Label lblcredito 
      Caption         =   "Label2"
      Height          =   255
      Left            =   1200
      TabIndex        =   70
      Top             =   10560
      Width           =   975
   End
   Begin VB.Label lbldiasrestantes 
      Caption         =   "Label2"
      Height          =   255
      Left            =   1200
      TabIndex        =   69
      Top             =   10320
      Width           =   975
   End
   Begin VB.Label lbldonador 
      Caption         =   "Label2"
      Height          =   255
      Left            =   1200
      TabIndex        =   68
      Top             =   10080
      Width           =   975
   End
   Begin VB.Image estado 
      Height          =   390
      Left            =   4080
      Top             =   9360
      Visible         =   0   'False
      Width           =   1425
   End
   Begin VB.Image command1 
      BorderStyle     =   1  'Fixed Single
      Height          =   180
      Index           =   45
      Left            =   1440
      Top             =   9000
      Width           =   210
   End
   Begin VB.Image command1 
      BorderStyle     =   1  'Fixed Single
      Height          =   180
      Index           =   44
      Left            =   2160
      Top             =   9000
      Width           =   210
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
      Index           =   23
      Left            =   1800
      TabIndex        =   67
      Top             =   9000
      Width           =   285
   End
   Begin VB.Label skills 
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
      Left            =   6600
      TabIndex        =   66
      Top             =   12720
      Width           =   1470
   End
   Begin VB.Image command1 
      BorderStyle     =   1  'Fixed Single
      Height          =   180
      Index           =   43
      Left            =   1440
      Top             =   8760
      Width           =   210
   End
   Begin VB.Image command1 
      BorderStyle     =   1  'Fixed Single
      Height          =   180
      Index           =   42
      Left            =   2160
      Top             =   8760
      Width           =   210
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
      Index           =   22
      Left            =   1800
      TabIndex        =   65
      Top             =   8760
      Width           =   285
   End
   Begin VB.Label skills 
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
      Index           =   22
      Left            =   1080
      TabIndex        =   64
      Top             =   9600
      Width           =   1620
   End
   Begin VB.Image command1 
      Height          =   300
      Index           =   41
      Left            =   8715
      Tag             =   "0"
      Top             =   3375
      Width           =   330
   End
   Begin VB.Image command1 
      Height          =   300
      Index           =   39
      Left            =   8715
      Tag             =   "0"
      Top             =   3765
      Width           =   330
   End
   Begin VB.Image command1 
      Height          =   300
      Index           =   38
      Left            =   9615
      Tag             =   "0"
      Top             =   3765
      Width           =   330
   End
   Begin VB.Image command1 
      Height          =   300
      Index           =   40
      Left            =   9615
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
      Left            =   9195
      TabIndex        =   63
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
      Left            =   9195
      TabIndex        =   62
      Top             =   3810
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
      ForeColor       =   &H000EA4EB&
      Height          =   150
      Index           =   20
      Left            =   6930
      TabIndex        =   61
      Top             =   3825
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
      Index           =   21
      Left            =   6930
      TabIndex        =   60
      Top             =   3465
      Width           =   1620
   End
   Begin VB.Image command1 
      Height          =   300
      Index           =   37
      Left            =   8715
      Tag             =   "0"
      Top             =   3000
      Width           =   330
   End
   Begin VB.Image command1 
      Height          =   300
      Index           =   36
      Left            =   9615
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
      Left            =   9195
      TabIndex        =   59
      Top             =   3060
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
      Index           =   19
      Left            =   6930
      TabIndex        =   58
      Top             =   3075
      Width           =   1620
   End
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
      Left            =   6480
      TabIndex        =   57
      Top             =   8880
      Width           =   975
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
      ForeColor       =   &H000EA4EB&
      Height          =   195
      Index           =   9
      Left            =   1920
      TabIndex        =   56
      Top             =   3135
      UseMnemonic     =   0   'False
      Width           =   1260
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
      TabIndex        =   55
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
      TabIndex        =   54
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
      TabIndex        =   53
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
      TabIndex        =   52
      Top             =   10440
      Visible         =   0   'False
      Width           =   900
   End
   Begin VB.Label skills 
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
      Left            =   1080
      TabIndex        =   49
      Top             =   9360
      Width           =   1620
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
      Left            =   6930
      TabIndex        =   48
      Top             =   6840
      Width           =   1620
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
      Left            =   3675
      TabIndex        =   47
      Top             =   4200
      Width           =   1620
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
      Left            =   3675
      TabIndex        =   46
      Top             =   5340
      Width           =   1620
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
      Left            =   6930
      TabIndex        =   45
      Top             =   6090
      Width           =   1620
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
      Left            =   3675
      TabIndex        =   44
      Top             =   4950
      Width           =   1620
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
      Left            =   6930
      TabIndex        =   43
      Top             =   5700
      Width           =   1620
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
      Left            =   6930
      TabIndex        =   42
      Top             =   4965
      Width           =   1620
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
      Left            =   6930
      TabIndex        =   41
      Top             =   4560
      Width           =   1620
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
      Left            =   3675
      TabIndex        =   40
      Top             =   5715
      Width           =   1620
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
      Left            =   3675
      TabIndex        =   39
      Top             =   2700
      Width           =   1620
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
      Left            =   3675
      TabIndex        =   38
      Top             =   4560
      Width           =   1620
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
      Left            =   3675
      TabIndex        =   37
      Top             =   3840
      Width           =   1620
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
      Left            =   3675
      TabIndex        =   36
      Top             =   6450
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
      Index           =   1
      Left            =   3675
      TabIndex        =   35
      Top             =   2340
      Width           =   1620
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
      Height          =   420
      Left            =   4200
      Tag             =   "1"
      Top             =   7500
      Width           =   1980
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
      Left            =   6120
      TabIndex        =   33
      Top             =   8520
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
      ForeColor       =   &H000EA4EB&
      Height          =   195
      Index           =   5
      Left            =   1920
      TabIndex        =   32
      Top             =   3405
      Width           =   1185
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
      ForeColor       =   &H000EA4EB&
      Height          =   195
      Index           =   0
      Left            =   1920
      TabIndex        =   31
      Top             =   6660
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
      ForeColor       =   &H000EA4EB&
      Height          =   195
      Index           =   6
      Left            =   1920
      TabIndex        =   28
      Top             =   2850
      Width           =   1260
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
      ForeColor       =   &H000EA4EB&
      Height          =   195
      Index           =   8
      Left            =   1920
      TabIndex        =   27
      Top             =   2580
      Width           =   1140
   End
   Begin VB.Label puntos 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000EA4EB&
      Height          =   240
      Left            =   8160
      TabIndex        =   26
      Top             =   1365
      Width           =   105
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
      Left            =   5940
      TabIndex        =   25
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
      Left            =   5940
      TabIndex        =   24
      Top             =   6420
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
      Left            =   5940
      TabIndex        =   23
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
      Index           =   4
      Left            =   5940
      TabIndex        =   22
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
      Index           =   5
      Left            =   5940
      TabIndex        =   21
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
      Left            =   5940
      TabIndex        =   20
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
      Left            =   9195
      TabIndex        =   19
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
      ForeColor       =   &H000EA4EB&
      Height          =   210
      Index           =   8
      Left            =   9195
      TabIndex        =   18
      Top             =   4920
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
      Left            =   9195
      TabIndex        =   17
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
      Index           =   10
      Left            =   5940
      TabIndex        =   16
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
      Index           =   11
      Left            =   9195
      TabIndex        =   15
      Top             =   6000
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
      Left            =   5940
      TabIndex        =   14
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
      Left            =   5940
      TabIndex        =   13
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
      Index           =   14
      Left            =   9195
      TabIndex        =   12
      Top             =   6795
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
      Left            =   1800
      TabIndex        =   11
      Top             =   8520
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
      Left            =   5940
      TabIndex        =   10
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
      Left            =   9195
      TabIndex        =   9
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
      Left            =   9195
      TabIndex        =   8
      Top             =   2655
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
      Height          =   300
      Index           =   0
      Left            =   6360
      Tag             =   "0"
      Top             =   2235
      Width           =   330
   End
   Begin VB.Image command1 
      Height          =   300
      Index           =   2
      Left            =   6360
      Tag             =   "0"
      Top             =   6375
      Width           =   330
   End
   Begin VB.Image command1 
      Height          =   300
      Index           =   4
      Left            =   6360
      Tag             =   "0"
      Top             =   3735
      Width           =   330
   End
   Begin VB.Image command1 
      Height          =   300
      Index           =   6
      Left            =   6360
      Tag             =   "0"
      Top             =   4500
      Width           =   330
   End
   Begin VB.Image command1 
      Height          =   300
      Index           =   8
      Left            =   6360
      Tag             =   "0"
      Top             =   2610
      Width           =   330
   End
   Begin VB.Image command1 
      Height          =   300
      Index           =   10
      Left            =   6360
      Tag             =   "0"
      Top             =   5625
      Width           =   330
   End
   Begin VB.Image command1 
      Height          =   300
      Index           =   12
      Left            =   9615
      Tag             =   "0"
      Top             =   4485
      Width           =   330
   End
   Begin VB.Image command1 
      Height          =   300
      Index           =   14
      Left            =   9615
      Tag             =   "0"
      Top             =   4860
      Width           =   330
   End
   Begin VB.Image command1 
      Height          =   300
      Index           =   16
      Left            =   9615
      Tag             =   "0"
      Top             =   5610
      Width           =   330
   End
   Begin VB.Image command1 
      Height          =   300
      Index           =   18
      Left            =   6360
      Tag             =   "0"
      Top             =   4875
      Width           =   330
   End
   Begin VB.Image command1 
      Height          =   300
      Index           =   20
      Left            =   9615
      Tag             =   "0"
      Top             =   6000
      Width           =   330
   End
   Begin VB.Image command1 
      Height          =   300
      Index           =   22
      Left            =   6360
      Tag             =   "0"
      Top             =   5250
      Width           =   330
   End
   Begin VB.Image command1 
      Height          =   300
      Index           =   24
      Left            =   6360
      Tag             =   "0"
      Top             =   4110
      Width           =   330
   End
   Begin VB.Image command1 
      Height          =   300
      Index           =   26
      Left            =   9615
      Tag             =   "0"
      Top             =   6765
      Width           =   330
   End
   Begin VB.Image command1 
      BorderStyle     =   1  'Fixed Single
      Height          =   180
      Index           =   28
      Left            =   2160
      Top             =   8520
      Width           =   210
   End
   Begin VB.Image command1 
      Height          =   300
      Index           =   30
      Left            =   6360
      Tag             =   "0"
      Top             =   3000
      Width           =   330
   End
   Begin VB.Image command1 
      Height          =   300
      Index           =   32
      Left            =   9615
      Tag             =   "0"
      Top             =   2235
      Width           =   330
   End
   Begin VB.Image command1 
      Height          =   300
      Index           =   34
      Left            =   9615
      Tag             =   "0"
      Top             =   2610
      Width           =   330
   End
   Begin VB.Image command1 
      Height          =   300
      Index           =   3
      Left            =   5460
      Tag             =   "0"
      Top             =   6375
      Width           =   330
   End
   Begin VB.Image command1 
      Height          =   300
      Index           =   5
      Left            =   5460
      Tag             =   "0"
      Top             =   3735
      Width           =   330
   End
   Begin VB.Image command1 
      Height          =   300
      Index           =   7
      Left            =   5460
      Tag             =   "0"
      Top             =   4500
      Width           =   330
   End
   Begin VB.Image command1 
      Height          =   300
      Index           =   9
      Left            =   5460
      Tag             =   "0"
      Top             =   2610
      Width           =   330
   End
   Begin VB.Image command1 
      Height          =   300
      Index           =   11
      Left            =   5460
      Tag             =   "0"
      Top             =   5625
      Width           =   330
   End
   Begin VB.Image command1 
      Height          =   300
      Index           =   13
      Left            =   8715
      Tag             =   "0"
      Top             =   4485
      Width           =   330
   End
   Begin VB.Image command1 
      Height          =   300
      Index           =   15
      Left            =   8715
      Tag             =   "0"
      Top             =   4860
      Width           =   330
   End
   Begin VB.Image command1 
      Height          =   300
      Index           =   17
      Left            =   8715
      Tag             =   "0"
      Top             =   5610
      Width           =   330
   End
   Begin VB.Image command1 
      Height          =   300
      Index           =   19
      Left            =   5460
      Tag             =   "0"
      Top             =   4875
      Width           =   330
   End
   Begin VB.Image command1 
      Height          =   300
      Index           =   21
      Left            =   8715
      Tag             =   "0"
      Top             =   6000
      Width           =   330
   End
   Begin VB.Image command1 
      Height          =   300
      Index           =   23
      Left            =   5460
      Tag             =   "0"
      Top             =   5250
      Width           =   330
   End
   Begin VB.Image command1 
      Height          =   300
      Index           =   25
      Left            =   5460
      Tag             =   "0"
      Top             =   4110
      Width           =   330
   End
   Begin VB.Image command1 
      Height          =   300
      Index           =   27
      Left            =   8715
      Tag             =   "0"
      Top             =   6765
      Width           =   330
   End
   Begin VB.Image command1 
      BorderStyle     =   1  'Fixed Single
      Height          =   180
      Index           =   29
      Left            =   1440
      Top             =   8520
      Width           =   210
   End
   Begin VB.Image command1 
      Height          =   300
      Index           =   31
      Left            =   5460
      Tag             =   "0"
      Top             =   3000
      Width           =   330
   End
   Begin VB.Image command1 
      Height          =   300
      Index           =   33
      Left            =   8715
      Tag             =   "0"
      Top             =   2235
      Width           =   330
   End
   Begin VB.Image command1 
      Height          =   300
      Index           =   35
      Left            =   8715
      Tag             =   "0"
      Top             =   2610
      Width           =   330
   End
   Begin VB.Image command1 
      Height          =   300
      Index           =   1
      Left            =   5460
      Tag             =   "0"
      Top             =   2235
      Width           =   330
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
      ForeColor       =   &H000EA4EB&
      Height          =   195
      Index           =   1
      Left            =   1920
      TabIndex        =   6
      Top             =   6390
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
      ForeColor       =   &H000EA4EB&
      Height          =   195
      Index           =   4
      Left            =   1920
      TabIndex        =   5
      Top             =   2325
      Width           =   1260
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
      ForeColor       =   &H000EA4EB&
      Height          =   210
      Index           =   4
      Left            =   1920
      TabIndex        =   4
      Top             =   5025
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
      ForeColor       =   &H000EA4EB&
      Height          =   210
      Index           =   3
      Left            =   1920
      TabIndex        =   3
      Top             =   4740
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
      ForeColor       =   &H000EA4EB&
      Height          =   210
      Index           =   2
      Left            =   1920
      TabIndex        =   2
      Top             =   4470
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
      ForeColor       =   &H000EA4EB&
      Height          =   210
      Index           =   1
      Left            =   1920
      TabIndex        =   1
      Top             =   4230
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
      ForeColor       =   &H000EA4EB&
      Height          =   195
      Index           =   3
      Left            =   1920
      TabIndex        =   0
      Top             =   6900
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
      TabIndex        =   51
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
      TabIndex        =   50
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
    
    On Error GoTo moverForm_Err
    

    Dim res As Long

    ReleaseCapture
    res = SendMessage(Me.hwnd, WM_SYSCOMMAND, MOUSE_MOVE, 0)

    
    Exit Sub

moverForm_Err:
    Call RegistrarError(Err.number, Err.Description, "frmEstadisticas.moverForm", Erl)
    Resume Next
    
End Sub

Public Sub Iniciar_Labels()
    
    On Error GoTo Iniciar_Labels_Err
    

    'Iniciamos los labels con los valores de los atributos y los skills
    Dim i As Integer

    For i = 1 To NUMATRIBUTOS 'Colocado
        Atri(i).Caption = UserAtributos(i)
    Next

    For i = 1 To NUMSKILLS
        If UserSkills(i) > 100 Then UserSkills(i) = 100

        Text1(i).Caption = UserSkills(i)
    Next

    Select Case UserEstadisticas.Alineacion

        Case 0
            Label6(9).Caption = "Criminal"
            Label6(9).ForeColor = RGB(255, 0, 0)

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

    'estado = LoadInterface(Language +"\estadisticascriminal.bmp")
    'Else
    '   Label6(9).Caption = "Ciudadano"
    'estado = LoadInterface(Language +"\estadisticasciudadano.bmp")
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

    
    Exit Sub

Iniciar_Labels_Err:
    Call RegistrarError(Err.number, Err.Description, "frmEstadisticas.Iniciar_Labels", Erl)
    Resume Next
    
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
    
    Call FormParser.Parse_Form(Me)
    'Image1.Picture = LoadInterface(Language +"\botonlargoaceptar.bmp")
    RealizoCambios = 0
    ReDim flags(1 To NUMSKILLS)

    
    Exit Sub

Form_Load_Err:
    Call RegistrarError(Err.number, Err.Description, "frmEstadisticas.Form_Load", Erl)
    Resume Next
    
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    On Error GoTo Form_MouseMove_Err
    
    moverForm
    'If Image1.Tag = "1" Then
    ' Image1.Picture = LoadInterface(Language +"\botonlargoaceptar.bmp")
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
    
    If Image1.Tag = "1" Then
        Set Image1.Picture = Nothing
        Image1.Tag = "0"
    End If
    
    If imgCerrar.Tag = "1" Then
        Set imgCerrar.Picture = Nothing
        imgCerrar.Tag = "0"
    End If

    For A = 0 To NUMSKILLS * 2 - 1
        If Command1(A).Tag = "1" Then
            Set Command1(A).Picture = Nothing
            Command1(A).Tag = "0"
        End If
    Next

    
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

    
    Exit Sub

Image1_Click_Err:
    Call RegistrarError(Err.number, Err.Description, "frmEstadisticas.Image1_Click", Erl)
    Resume Next
    
End Sub

Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    On Error GoTo Image1_MouseDown_Err
    
    Image1 = LoadInterface(Language + "\boton-aceptar-ES-off.bmp")
    Image1.Tag = "1"
    
    Exit Sub

Image1_MouseDown_Err:
    Call RegistrarError(Err.number, Err.Description, "frmEstadisticas.Image1_MouseDown", Erl)
    Resume Next
    
End Sub

Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    On Error GoTo Image1_MouseMove_Err
    

    If Image1.Tag = "0" Then
        Image1 = LoadInterface(Language +"\boton-aceptar-ES-over.bmp")
        Image1.Tag = "1"
    End If

    
    Exit Sub

Image1_MouseMove_Err:
    Call RegistrarError(Err.number, Err.Description, "frmEstadisticas.Image1_MouseMove", Erl)
    Resume Next
    
End Sub

Private Sub imgCerrar_Click()
    
    On Error GoTo imgCerrar_Click_Err
    
    If RealizoCambios >= 1 Then
        If MsgBox("Realizó cambios en sus skillpoints ¿desea guardar antes de salir?", vbYesNo) = vbYes Then

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

Private Sub imgCerrar_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    On Error GoTo imgCerrar_MouseDown_Err
    
    imgCerrar.Picture = LoadInterface(Language +"\boton-cerrar-off.bmp")
    imgCerrar.Tag = "1"
    
    Exit Sub

imgCerrar_MouseDown_Err:
    Call RegistrarError(Err.number, Err.Description, "frmEstadisticas.imgCerrar_MouseDown", Erl)
    Resume Next
    
End Sub

Private Sub imgCerrar_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    On Error GoTo imgCerrar_MouseMove_Err
    
    If imgCerrar.Tag = "0" Then
        imgCerrar.Picture = LoadInterface(Language +"\boton-cerrar-over.bmp")
        imgCerrar.Tag = "1"
    End If
    
    Exit Sub

imgCerrar_MouseMove_Err:
    Call RegistrarError(Err.number, Err.Description, "frmEstadisticas.imgCerrar_MouseMove", Erl)
    Resume Next
    
End Sub


Private Sub skills_Click(Index As Integer)
    
    On Error GoTo skills_Click_Err
    
    AddtoRichTextBox frmMain.RecTxt, "Información del skill> " & SkillsDesc(Index), 2, 51, 223, 1, 1

    
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
