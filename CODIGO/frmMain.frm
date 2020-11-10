VERSION 5.00
Object = "{33101C00-75C3-11CF-A8A0-444553540000}#1.0#0"; "CSWSK32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmmain 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   13305
   ClientLeft      =   345
   ClientTop       =   240
   ClientWidth     =   20295
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   -1  'True
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   Icon            =   "frmMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   Picture         =   "frmMain.frx":57E2
   ScaleHeight     =   815.007
   ScaleMode       =   0  'User
   ScaleWidth      =   1356.391
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin SocketWrenchCtrl.Socket Socket1 
      Left            =   9240
      Top             =   2640
      _Version        =   65536
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      AutoResolve     =   0   'False
      Backlog         =   1
      Binary          =   -1  'True
      Blocking        =   0   'False
      Broadcast       =   0   'False
      BufferSize      =   10240
      HostAddress     =   ""
      HostFile        =   ""
      HostName        =   ""
      InLine          =   0   'False
      Interval        =   0
      KeepAlive       =   0   'False
      Library         =   ""
      Linger          =   0
      LocalPort       =   0
      LocalService    =   ""
      Protocol        =   0
      RemotePort      =   0
      RemoteService   =   ""
      ReuseAddress    =   0   'False
      Route           =   -1  'True
      Timeout         =   1000
      Type            =   1
      Urgent          =   0   'False
   End
   Begin VB.CommandButton createObj 
      Caption         =   "Crear OBJ"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4800
      TabIndex        =   38
      Top             =   0
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton panelGM 
      Caption         =   "Panel GM"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3240
      TabIndex        =   37
      Top             =   0
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.PictureBox panel 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4875
      Left            =   11520
      Picture         =   "frmMain.frx":30469
      ScaleHeight     =   4875
      ScaleWidth      =   3705
      TabIndex        =   34
      Top             =   2400
      Width           =   3705
      Begin VB.ListBox hlst 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   3390
         Left            =   255
         TabIndex        =   35
         TabStop         =   0   'False
         Top             =   630
         Visible         =   0   'False
         Width           =   3195
      End
      Begin VB.PictureBox picInv 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   3840
         Left            =   280
         ScaleHeight     =   256
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   210
         TabIndex        =   36
         Top             =   740
         Width           =   3150
      End
      Begin VB.Image imgSpellInfo 
         Height          =   345
         Left            =   2490
         Tag             =   "1"
         Top             =   4230
         Width           =   375
      End
      Begin VB.Image imgInvLock 
         Height          =   210
         Index           =   2
         Left            =   30
         Top             =   4185
         Width           =   210
      End
      Begin VB.Image imgInvLock 
         Height          =   210
         Index           =   1
         Left            =   30
         Top             =   3660
         Width           =   210
      End
      Begin VB.Image imgInvLock 
         Height          =   210
         Index           =   0
         Left            =   30
         Top             =   3160
         Width           =   210
      End
      Begin VB.Image imgInventario 
         Height          =   420
         Left            =   10
         Tag             =   "0"
         Top             =   10
         Width           =   1830
      End
      Begin VB.Image imgHechizos 
         Height          =   420
         Left            =   1880
         Tag             =   "0"
         Top             =   20
         Width           =   1830
      End
      Begin VB.Image cmdMoverHechi 
         Height          =   285
         Index           =   0
         Left            =   3370
         MouseIcon       =   "frmMain.frx":6B533
         MousePointer    =   99  'Custom
         Tag             =   "0"
         Top             =   4550
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.Image cmdMoverHechi 
         Height          =   285
         Index           =   1
         Left            =   3370
         MouseIcon       =   "frmMain.frx":6B685
         MousePointer    =   99  'Custom
         Picture         =   "frmMain.frx":6B7D7
         Tag             =   "0"
         Top             =   4260
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.Image cmdlanzar 
         Height          =   585
         Left            =   0
         Tag             =   "1"
         Top             =   4260
         Visible         =   0   'False
         Width           =   2415
      End
   End
   Begin VB.PictureBox Panels 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4410
      Left            =   16680
      ScaleHeight     =   294
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   261
      TabIndex        =   32
      Top             =   7200
      Width           =   3915
      Begin VB.Image Image2 
         Height          =   495
         Index           =   1
         Left            =   600
         Tag             =   "1"
         Top             =   0
         Width           =   3405
      End
   End
   Begin VB.TextBox SendTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00202020&
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
      Height          =   360
      Left            =   600
      MaxLength       =   160
      MultiLine       =   -1  'True
      TabIndex        =   24
      TabStop         =   0   'False
      ToolTipText     =   "Chat"
      Top             =   1828
      Visible         =   0   'False
      Width           =   8184
   End
   Begin VB.PictureBox panelInf 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3000
      Left            =   15840
      ScaleHeight     =   200
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   304
      TabIndex        =   10
      Top             =   3120
      Width           =   4560
      Begin VB.Image manualboton 
         Height          =   390
         Left            =   2565
         Tag             =   "0"
         Top             =   1230
         Visible         =   0   'False
         Width           =   1185
      End
      Begin VB.Image rankingBoton 
         Height          =   375
         Left            =   2565
         Tag             =   "0"
         Top             =   675
         Visible         =   0   'False
         Width           =   1230
      End
   End
   Begin VB.PictureBox PicMeteo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1350
      Left            =   18000
      ScaleHeight     =   1350
      ScaleWidth      =   1290
      TabIndex        =   9
      Top             =   9840
      Width           =   1290
   End
   Begin VB.PictureBox MiniMap 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1500
      Left            =   9576
      ScaleHeight     =   100
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   100
      TabIndex        =   1
      ToolTipText     =   "Tu posicion en el mapa, click para mas info."
      Top             =   600
      Width           =   1500
      Begin VB.Shape personaje 
         BackColor       =   &H00FF0000&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00FF0000&
         FillColor       =   &H00FFFFFF&
         Height          =   60
         Index           =   5
         Left            =   0
         Shape           =   3  'Circle
         Top             =   0
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.Shape personaje 
         BackColor       =   &H00FF00FF&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00FF00FF&
         FillColor       =   &H00FFFFFF&
         Height          =   60
         Index           =   4
         Left            =   0
         Shape           =   3  'Circle
         Top             =   0
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.Shape personaje 
         BackColor       =   &H000080FF&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H000080FF&
         FillColor       =   &H00FFFFFF&
         Height          =   60
         Index           =   3
         Left            =   0
         Shape           =   3  'Circle
         Top             =   0
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.Shape personaje 
         BackColor       =   &H0000C000&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H0000C000&
         FillColor       =   &H00FFFFFF&
         Height          =   60
         Index           =   2
         Left            =   0
         Shape           =   3  'Circle
         Top             =   0
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.Shape personaje 
         BackColor       =   &H0000FFFF&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H0000FFFF&
         FillColor       =   &H00FFFFFF&
         Height          =   60
         Index           =   1
         Left            =   0
         Shape           =   3  'Circle
         Top             =   0
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.Shape personaje 
         BackColor       =   &H000000C0&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00FFFFFF&
         FillColor       =   &H0000FFFF&
         Height          =   90
         Index           =   0
         Left            =   450
         Shape           =   5  'Rounded Square
         Top             =   750
         Width           =   150
      End
   End
   Begin VB.Timer macrotrabajo 
      Enabled         =   0   'False
      Left            =   2760
      Top             =   2400
   End
   Begin VB.Timer Second 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   10440
      Top             =   2640
   End
   Begin VB.PictureBox renderer 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   9060
      Left            =   180
      ScaleHeight     =   604
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   732
      TabIndex        =   3
      Top             =   2286
      Width           =   10982
      Begin VB.Timer LlamaDeclan 
         Enabled         =   0   'False
         Interval        =   60000
         Left            =   10200
         Top             =   8160
      End
      Begin VB.Timer cerrarcuenta 
         Enabled         =   0   'False
         Interval        =   200
         Left            =   6960
         Top             =   120
      End
      Begin VB.Timer Contadores 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   1200
         Top             =   7800
      End
      Begin VB.Timer HoraFantasiaTimer 
         Enabled         =   0   'False
         Interval        =   600
         Left            =   240
         Top             =   7920
      End
      Begin VB.Timer TimerLluvia 
         Enabled         =   0   'False
         Interval        =   50
         Left            =   600
         Top             =   120
      End
      Begin VB.Timer TimerNiebla 
         Interval        =   100
         Left            =   1080
         Top             =   120
      End
      Begin VB.Timer Timerping 
         Enabled         =   0   'False
         Interval        =   7000
         Left            =   4440
         Top             =   120
      End
      Begin VB.Timer MacroLadder 
         Enabled         =   0   'False
         Interval        =   1300
         Left            =   1560
         Top             =   120
      End
      Begin VB.Timer Efecto 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   2040
         Top             =   120
      End
   End
   Begin RichTextLib.RichTextBox RecTxt 
      Height          =   1273
      Left            =   240
      TabIndex        =   7
      TabStop         =   0   'False
      ToolTipText     =   "Mensajes del servidor"
      Top             =   490
      Width           =   9187
      _ExtentX        =   16193
      _ExtentY        =   2249
      _Version        =   393217
      BackColor       =   0
      BorderStyle     =   0
      Enabled         =   -1  'True
      HideSelection   =   0   'False
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      DisableNoScroll =   -1  'True
      Appearance      =   0
      TextRTF         =   $"frmMain.frx":6BE55
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label manabar 
      Alignment       =   2  'Center
      BackColor       =   &H80000006&
      BackStyle       =   0  'Transparent
      Caption         =   "999/999"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   12840
      TabIndex        =   14
      Top             =   8718
      Width           =   1335
   End
   Begin VB.Image Temp2 
      Height          =   495
      Left            =   13990
      Tag             =   "0"
      ToolTipText     =   "Deshabilitado por el momento"
      Top             =   10122
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.Image Temp1 
      Height          =   495
      Left            =   13376
      Tag             =   "0"
      ToolTipText     =   "Deshabilitado por el momento"
      Top             =   10122
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.Image imgBugReport 
      Height          =   495
      Left            =   14220
      Top             =   10872
      Width           =   1020
   End
   Begin VB.Label ObjLbl 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "ItemData"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   495
      Left            =   11400
      TabIndex        =   33
      Top             =   7260
      Visible         =   0   'False
      Width           =   3900
   End
   Begin VB.Label ms 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "30 ms"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   8280
      TabIndex        =   31
      Top             =   240
      Width           =   1065
   End
   Begin VB.Label fps 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Fps: 200"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   8280
      TabIndex        =   30
      ToolTipText     =   "Numero de usuarios online"
      Top             =   65
      Width           =   1065
   End
   Begin VB.Image EstadisticasBoton 
      Height          =   420
      Left            =   14738
      Tag             =   "0"
      Top             =   1453
      Width           =   465
   End
   Begin VB.Image Image5 
      Height          =   495
      Left            =   12165
      Tag             =   "0"
      ToolTipText     =   "Grupo"
      Top             =   10122
      Width           =   540
   End
   Begin VB.Image clanimg 
      Height          =   495
      Left            =   12765
      Tag             =   "0"
      ToolTipText     =   "Clanes"
      Top             =   10122
      Width           =   540
   End
   Begin VB.Image QuestBoton 
      Height          =   495
      Left            =   14610
      Tag             =   "0"
      ToolTipText     =   "Quests"
      Top             =   10122
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.Label Coord 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "000 X:00 Y: 00"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   9480
      TabIndex        =   29
      Top             =   240
      Width           =   1695
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre del pj"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFC0C0&
      Height          =   495
      Left            =   11400
      TabIndex        =   28
      Top             =   600
      Width           =   3825
   End
   Begin VB.Label lblPorcLvl 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "33.33%"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   225
      Left            =   12987
      TabIndex        =   26
      Top             =   1224
      Width           =   675
   End
   Begin VB.Label lblArmor 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0/0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H008080FF&
      Height          =   255
      Left            =   14610
      TabIndex        =   23
      Top             =   9450
      Width           =   525
   End
   Begin VB.Label lblHelm 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0/0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   255
      Left            =   13620
      TabIndex        =   22
      Top             =   9452
      Width           =   615
   End
   Begin VB.Label lblShielder 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0/0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF80&
      Height          =   255
      Left            =   12750
      TabIndex        =   21
      Top             =   9452
      Width           =   495
   End
   Begin VB.Label lblWeapon 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0/0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   11760
      TabIndex        =   20
      Top             =   9452
      Width           =   615
   End
   Begin VB.Label AgilidadLbl 
      BackStyle       =   0  'Transparent
      Caption         =   "00"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   14400
      TabIndex        =   19
      Top             =   7885
      Width           =   450
   End
   Begin VB.Label Fuerzalbl 
      BackStyle       =   0  'Transparent
      Caption         =   "00"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   15000
      TabIndex        =   18
      Top             =   7885
      Width           =   330
   End
   Begin VB.Label GldLbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "100000"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   240
      Left            =   11880
      TabIndex        =   17
      Top             =   7885
      Width           =   720
   End
   Begin VB.Image TiendaBoton 
      Height          =   405
      Left            =   17040
      Tag             =   "0"
      Top             =   8895
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Image OpcionesBoton 
      Height          =   315
      Left            =   11431
      Tag             =   "0"
      Top             =   65
      Width           =   315
   End
   Begin VB.Label oxigenolbl 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "15"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   13785
      TabIndex        =   11
      ToolTipText     =   "Oxigeno acumulado"
      Top             =   7885
      Width           =   255
   End
   Begin VB.Image mapMundo 
      Height          =   495
      Left            =   11550
      Tag             =   "0"
      ToolTipText     =   "Mapa del mundo"
      Top             =   10122
      Width           =   540
   End
   Begin VB.Image CombateIcon 
      Height          =   180
      Left            =   8828
      Picture         =   "frmMain.frx":6BECC
      Top             =   1812
      Width           =   555
   End
   Begin VB.Image globalIcon 
      Height          =   180
      Left            =   8828
      Picture         =   "frmMain.frx":6C450
      Top             =   2008
      Width           =   555
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   10320
      TabIndex        =   8
      ToolTipText     =   "Activar / desactivar chat globales"
      Top             =   1800
      Width           =   855
   End
   Begin VB.Label lblResis 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "25%"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   16200
      TabIndex        =   6
      ToolTipText     =   "Tu daño magico"
      Top             =   10680
      Width           =   420
   End
   Begin VB.Label lbldm 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "25%"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   16320
      TabIndex        =   5
      ToolTipText     =   "Tu defensa magica"
      Top             =   11160
      Width           =   420
   End
   Begin VB.Image Image3 
      Height          =   210
      Left            =   284
      Tag             =   "0"
      ToolTipText     =   "Modo de chat"
      Top             =   1894
      Width           =   255
   End
   Begin VB.Image Image4 
      Height          =   315
      Index           =   1
      Left            =   14977
      Tag             =   "0"
      Top             =   65
      Width           =   315
   End
   Begin VB.Image Image4 
      Height          =   315
      Index           =   0
      Left            =   14603
      Tag             =   "0"
      Top             =   65
      Width           =   315
   End
   Begin VB.Label onlines 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Onlines: 1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   255
      Left            =   12480
      TabIndex        =   4
      ToolTipText     =   "Numero de usuarios online"
      Top             =   120
      Width           =   1665
   End
   Begin VB.Shape MainViewShp 
      BorderColor     =   &H000000FF&
      Height          =   2430
      Left            =   15600
      Top             =   7920
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   7680
      TabIndex        =   2
      Top             =   1680
      Width           =   450
   End
   Begin VB.Image PicCorreo 
      Height          =   435
      Left            =   13560
      Picture         =   "frmMain.frx":6C9D4
      Top             =   10920
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.Image PicResu 
      Height          =   510
      Left            =   11970
      Picture         =   "frmMain.frx":6D654
      ToolTipText     =   "Seguro de grupo"
      Top             =   10872
      Width           =   510
   End
   Begin VB.Image PicResuOn 
      Height          =   510
      Left            =   11970
      Picture         =   "frmMain.frx":6E466
      ToolTipText     =   "Seguro de grupo"
      Top             =   10872
      Visible         =   0   'False
      Width           =   510
   End
   Begin VB.Label NameMapa 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Mapa Desconocido"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   9456
      TabIndex        =   0
      Top             =   33
      Width           =   1740
   End
   Begin VB.Label lblLvl 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "255"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   11520
      TabIndex        =   27
      Top             =   960
      Width           =   3525
   End
   Begin VB.Label exp 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "99999/99999"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   11573
      TabIndex        =   25
      Top             =   1535
      Width           =   3045
   End
   Begin VB.Image ExpBar 
      Height          =   240
      Left            =   11573
      Picture         =   "frmMain.frx":6F278
      Top             =   1545
      Width           =   3045
   End
   Begin VB.Label HpBar 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "999/999"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   11880
      TabIndex        =   15
      Top             =   8301
      Width           =   3255
   End
   Begin VB.Image Hpshp 
      Height          =   240
      Left            =   11865
      Picture         =   "frmMain.frx":72111
      Top             =   8309
      Width           =   3240
   End
   Begin VB.Label AGUbar 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "100"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   165
      Left            =   13680
      TabIndex        =   12
      Top             =   9134
      Width           =   495
   End
   Begin VB.Label hambar 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "100"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   165
      Left            =   14618
      TabIndex        =   13
      Top             =   9142
      Width           =   495
   End
   Begin VB.Image AGUAsp 
      Height          =   135
      Left            =   13680
      Picture         =   "frmMain.frx":749D5
      Top             =   9150
      Width           =   480
   End
   Begin VB.Image PicSeg 
      Appearance      =   0  'Flat
      Height          =   510
      Left            =   11446
      Picture         =   "frmMain.frx":74D79
      ToolTipText     =   "Seguro de ataque"
      Top             =   10872
      Width           =   510
   End
   Begin VB.Image Image6 
      Height          =   510
      Left            =   11446
      Picture         =   "frmMain.frx":75B8B
      ToolTipText     =   "Seguro de ataque"
      Top             =   10872
      Width           =   510
   End
   Begin VB.Image PicSegClanOn 
      Appearance      =   0  'Flat
      Height          =   510
      Left            =   12494
      Picture         =   "frmMain.frx":7699D
      ToolTipText     =   "Seguro de clan"
      Top             =   10875
      Width           =   510
   End
   Begin VB.Image PicSegClanOff 
      Appearance      =   0  'Flat
      Height          =   510
      Left            =   12494
      Picture         =   "frmMain.frx":777AF
      ToolTipText     =   "Seguro de ataque"
      Top             =   10875
      Width           =   510
   End
   Begin VB.Image COMIDAsp 
      Height          =   120
      Left            =   14618
      Picture         =   "frmMain.frx":785C1
      Top             =   9158
      Width           =   480
   End
   Begin VB.Image MANShp 
      Height          =   240
      Left            =   11865
      Picture         =   "frmMain.frx":78905
      Top             =   8709
      Width           =   3240
   End
   Begin VB.Label stabar 
      Alignment       =   2  'Center
      BackColor       =   &H0000FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "999/999"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   165
      Left            =   11880
      TabIndex        =   16
      Top             =   9142
      Width           =   1335
   End
   Begin VB.Image STAShp 
      Height          =   135
      Left            =   11850
      Picture         =   "frmMain.frx":7B1C9
      Top             =   9153
      Width           =   1335
   End
   Begin VB.Menu mnuObj 
      Caption         =   "Objeto"
      Visible         =   0   'False
      Begin VB.Menu mnuTirar 
         Caption         =   "Tirar"
      End
      Begin VB.Menu mnuUsar 
         Caption         =   "Usar"
      End
      Begin VB.Menu mnuEquipar 
         Caption         =   "Equipar"
      End
   End
   Begin VB.Menu mnuNpc 
      Caption         =   "NPC"
      Visible         =   0   'False
      Begin VB.Menu mnuNpcDesc 
         Caption         =   "Descripcion"
      End
      Begin VB.Menu mnuNpcComerciar 
         Caption         =   "Comerciar"
         Visible         =   0   'False
      End
   End
End
Attribute VB_Name = "frmmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'You can contact me at:
'morgolock@speedy.com.ar
'
'Calle 3 número 983 piso 7 dto A
'La Plata - Pcia, Buenos Aires - Republica Argentina
'Código Postal 1900
'Pablo Ignacio Márquez
'Call ParseUserCommand("/CMSG " & stxtbuffercmsg)
Option Explicit

Public WithEvents Inventario As clsGrapchicalInventory
Attribute Inventario.VB_VarHelpID = -1

Private Const WS_EX_TRANSPARENT = &H20&

Private Const GWL_EXSTYLE = (-20)

Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

' Constantes para SendMessage
Const WM_SYSCOMMAND As Long = &H112&

Const MOUSE_MOVE    As Long = &HF012&

Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long

Private MenuNivel As Byte

Private Type POINTAPI

    x As Long
    y As Long

End Type

Dim Mouse As POINTAPI

Private Declare Function ReleaseCapture Lib "user32" () As Long

Public MouseBoton As Long

Public MouseShift As Long

Public IsPlaying  As Byte

Public bmoving    As Boolean

Public dX         As Integer

Public dy         As Integer

' Constantes para SendMessage

Const HWND_TOPMOST = -1

Const HWND_NOTOPMOST = -2

Const SWP_NOSIZE = &H1

Const SWP_NOMOVE = &H2

Const SWP_NOACTIVATE = &H10

Const SWP_SHOWWINDOW = &H40

Private Declare Sub SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As String) As Long

Private Const EM_GETLINE = &HC4

Private Const EM_LINELENGTH = &HC1

Private Sub clanimg_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

    If clanimg.Tag = "0" Then
        clanimg.Picture = LoadInterface("claniluminado.bmp")
        clanimg.Tag = "1"

    End If

End Sub

Private Sub clanimg_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

    If frmGuildLeader.Visible Then Unload frmGuildLeader
    Call WriteRequestGuildLeaderInfo

End Sub

Private Sub cmdlanzar_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    cmdlanzar.Picture = LoadInterface("boton-lanzar-ES-off.bmp")
    cmdlanzar.Tag = "1"

End Sub

Private Sub cmdlanzar_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call Form_MouseMove(Button, Shift, x, y)

End Sub

Private Sub cmdMoverHechi_Click(Index As Integer)

    If hlst.ListIndex = -1 Then Exit Sub

    Dim sTemp As String

    Select Case Index

        Case 1 'subir

            If hlst.ListIndex = 0 Then Exit Sub

        Case 0 'bajar

            If hlst.ListIndex = hlst.ListCount - 1 Then Exit Sub

    End Select

    Call WriteMoveSpell(Index, hlst.ListIndex + 1)
    
    Select Case Index

        Case 1 'subir
            sTemp = hlst.List(hlst.ListIndex - 1)
            hlst.List(hlst.ListIndex - 1) = hlst.List(hlst.ListIndex)
            hlst.List(hlst.ListIndex) = sTemp
            hlst.ListIndex = hlst.ListIndex - 1

        Case 0 'bajar
            sTemp = hlst.List(hlst.ListIndex + 1)
            hlst.List(hlst.ListIndex + 1) = hlst.List(hlst.ListIndex)
            hlst.List(hlst.ListIndex) = sTemp
            hlst.ListIndex = hlst.ListIndex + 1

    End Select

End Sub

Public Sub ControlSeguroParty(ByVal Mostrar As Boolean)

    If Mostrar Then
        If Not PicResu.Visible Then
            PicResu.Visible = True
            PicResuOn.Visible = False

        End If

    Else

        If PicResu.Visible Then
            PicResu.Visible = False
            PicResuOn.Visible = True

        End If

    End If

End Sub

Public Sub DibujarSeguro()
    PicSeg.Visible = True

End Sub

Public Sub DesDibujarSeguro()
    PicSeg.Visible = False

End Sub

Private Sub cmdMoverHechi_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)

    Select Case Index

        Case 0

            If cmdMoverHechi(Index).Tag = "0" Then
                cmdMoverHechi(Index).Picture = LoadInterface("boton-sm-flecha-aba-over.bmp")
                cmdMoverHechi(Index).Tag = "1"

            End If

        Case 1

            If cmdMoverHechi(Index).Tag = "0" Then
                cmdMoverHechi(Index).Picture = LoadInterface("boton-sm-flecha-arr-over.bmp")
                cmdMoverHechi(Index).Tag = "1"

            End If
    
    End Select
    
End Sub

Private Sub CombateIcon_Click()

    If ChatCombate = 0 Then
        ChatCombate = 1
        CombateIcon.Picture = LoadInterface("infoapretado.bmp")
    Else
        ChatCombate = 0
        CombateIcon.Picture = LoadInterface("info.bmp")

    End If

    Call WriteMacroPos

End Sub

Private Sub Contadores_Timer()

    If UserEstado = 1 Then Exit Sub
    If InviCounter > 0 Then
        InviCounter = InviCounter - 1

    End If

    If ScrollExpCounter > 0 Then
        ScrollExpCounter = ScrollExpCounter - 1

    End If

    If ScrollOroCounter > 0 Then
        ScrollOroCounter = ScrollOroCounter - 1

    End If

    If OxigenoCounter > 0 Then
        OxigenoCounter = OxigenoCounter - 1

    End If

    If DrogaCounter > 0 Then
        DrogaCounter = DrogaCounter - 1

        If DrogaCounter = 12 Then
            frmmain.Fuerzalbl.ForeColor = vbWhite
            frmmain.AgilidadLbl.ForeColor = vbWhite
        ElseIf DrogaCounter = 11 Then
            frmmain.Fuerzalbl.ForeColor = RGB(204, 0, 0)
            frmmain.AgilidadLbl.ForeColor = RGB(204, 0, 0)
        ElseIf DrogaCounter = 10 Then
            frmmain.Fuerzalbl.ForeColor = vbWhite
            frmmain.AgilidadLbl.ForeColor = vbWhite
        ElseIf DrogaCounter = 9 Then
            frmmain.Fuerzalbl.ForeColor = RGB(204, 0, 0)
            frmmain.AgilidadLbl.ForeColor = RGB(204, 0, 0)
        ElseIf DrogaCounter = 8 Then
            frmmain.Fuerzalbl.ForeColor = vbWhite
            frmmain.AgilidadLbl.ForeColor = vbWhite
        ElseIf DrogaCounter = 7 Then
            frmmain.Fuerzalbl.ForeColor = RGB(204, 0, 0)
            frmmain.AgilidadLbl.ForeColor = RGB(204, 0, 0)
        ElseIf DrogaCounter = 6 Then
            frmmain.Fuerzalbl.ForeColor = vbWhite
            frmmain.AgilidadLbl.ForeColor = vbWhite
        ElseIf DrogaCounter = 5 Then
            frmmain.Fuerzalbl.ForeColor = RGB(204, 0, 0)
            frmmain.AgilidadLbl.ForeColor = RGB(204, 0, 0)
        ElseIf DrogaCounter = 4 Then
            frmmain.Fuerzalbl.ForeColor = vbWhite
            frmmain.AgilidadLbl.ForeColor = vbWhite
        ElseIf DrogaCounter = 3 Then
            frmmain.Fuerzalbl.ForeColor = RGB(204, 0, 0)
            frmmain.AgilidadLbl.ForeColor = RGB(204, 0, 0)
        ElseIf DrogaCounter = 2 Then
            frmmain.Fuerzalbl.ForeColor = vbWhite
            frmmain.AgilidadLbl.ForeColor = vbWhite
        ElseIf DrogaCounter = 1 Then
            frmmain.Fuerzalbl.ForeColor = RGB(204, 0, 0)
            frmmain.AgilidadLbl.ForeColor = RGB(204, 0, 0)

        End If

    End If

    If InviCounter = 0 And ScrollExpCounter = 0 And ScrollOroCounter = 0 And DrogaCounter = 0 And OxigenoCounter = 0 Then
        Contadores.Enabled = False

    End If

End Sub

Private Sub createObj_Click()
    Dim i As Long
    For i = 1 To NumOBJs

        If ObjData(i).name <> "" Then

            Dim subelemento As ListItem

            Set subelemento = FrmObjetos.ListView1.ListItems.Add(, , ObjData(i).name)
            
            subelemento.SubItems(1) = i

        End If

    Next i
    FrmObjetos.Show , Me
End Sub

Private Sub Efecto_Timer()
    Call engine.Map_Base_Light_Set(Map_light_baseBackup)
    Efecto.Enabled = False
    EfectoEnproceso = False

End Sub

Private Sub EstadisticasBoton_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    EstadisticasBoton.Picture = LoadInterface("boton-estadisticas-off.bmp")
    EstadisticasBoton.Tag = "1"

End Sub

Private Sub EstadisticasBoton_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

    If EstadisticasBoton.Tag = "0" Then
        EstadisticasBoton.Picture = LoadInterface("boton-estadisticas-over.bmp")
        EstadisticasBoton.Tag = "1"

    End If

End Sub

Private Sub EstadisticasBoton_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    LlegaronAtrib = False
    LlegaronSkills = False
    Call WriteRequestAtributes
    Call WriteRequestSkills
    Call WriteRequestMiniStats
            
    Call FlushBuffer

    Do While Not LlegaronSkills Or Not LlegaronAtrib
        DoEvents 'esperamos a que lleguen y mantenemos la interfaz viva
    Loop
            
    Alocados = SkillPoints
    frmEstadisticas.puntos.Caption = SkillPoints
    frmEstadisticas.Iniciar_Labels
    frmEstadisticas.Picture = LoadInterface("VentanaEstadisticas.bmp")
    HayFormularioAbierto = True
    frmEstadisticas.Show , frmmain
    LlegaronAtrib = False
    LlegaronSkills = False

End Sub

Private Sub exp_Click()
    Call WriteScrollInfo

End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)

    If Not SendTxt.Visible Then
        If Not pausa And frmmain.Visible And Not frmComerciar.Visible And Not frmComerciarUsu.Visible And Not frmBancoObj.Visible And Not frmGoliath.Visible Then
    
            If Accionar(KeyCode) Then
                Exit Sub
            ElseIf KeyCode = vbKeyReturn Then

                If Not frmCantidad.Visible Then
                    Call CompletarEnvioMensajes
                    SendTxt.Visible = True
                    SendTxt.SetFocus
                    Call WriteEscribiendo
                
                End If

            ElseIf KeyCode = vbKeyEscape And Not UserSaliendo Then
                frmCerrar.Show , frmmain
                ' Call WriteQuit
        
            ElseIf KeyCode = 27 And UserSaliendo Then
                Call WriteCancelarExit

                Rem  Call SendData("CU")
            End If

        End If

    Else
        SendTxt.SetFocus

    End If

End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    MouseBoton = Button
    MouseShift = Shift
    
    If frmComerciar.Visible Then
        Unload frmComerciar

    End If
    
    If frmBancoObj.Visible Then
        Unload frmBancoObj

    End If

End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    clicX = x
    clicY = y

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    If prgRun = True Then
        prgRun = False
        Cancel = 1

    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)

    Call DisableURLDetect

End Sub

Private Sub GldLbl_Click()
    Inventario.SelectGold

    If UserGLD > 0 Then
        frmCantidad.Picture = LoadInterface("cantidad.bmp")
        HayFormularioAbierto = True
        frmCantidad.Show , frmmain

    End If

End Sub

Private Sub GlobalIcon_Click()

    If ChatGlobal = 0 Then
        ChatGlobal = 1
        globalIcon.Picture = LoadInterface("globalapretado.bmp")
    Else
        ChatGlobal = 0
        globalIcon.Picture = LoadInterface("global.bmp")

    End If

    Call WriteMacroPos

End Sub

Private Sub hlst_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

    If cmdlanzar.Tag = "1" Then
        cmdlanzar.Picture = Nothing
        cmdlanzar.Tag = "0"

    End If

    If cmdMoverHechi(1).Tag = "1" Then
        cmdMoverHechi(1).Picture = Nothing
        cmdMoverHechi(1).Tag = "0"

    End If
        
    If cmdMoverHechi(0).Tag = "1" Then
        cmdMoverHechi(0).Picture = Nothing
        cmdMoverHechi(0).Tag = "0"

    End If

End Sub

Private Sub HoraFantasiaTimer_Timer()

    HoraFantasia = HoraFantasia + 1

    'frmMain.lblHoraFantasia.Caption = GetTimeFormated(HoraFantasia)
    Select Case HoraFantasia

        Case 0 '0
            frmmain.PicMeteo.Picture = LoadInterface("a29.bmp")

        Case 60 '1
            frmmain.PicMeteo.Picture = LoadInterface("a1.bmp")

        Case 120 '2
            frmmain.PicMeteo.Picture = LoadInterface("a2.bmp")

        Case 180 '3
            frmmain.PicMeteo.Picture = LoadInterface("a3.bmp")

        Case 240 '4
            frmmain.PicMeteo.Picture = LoadInterface("a4.bmp")

        Case 270 '4:30
            frmmain.PicMeteo.Picture = LoadInterface("a5.bmp")

        Case 300 '5
            frmmain.PicMeteo.Picture = LoadInterface("a6.bmp")

        Case 330 '5:30
            frmmain.PicMeteo.Picture = LoadInterface("a7.bmp")

        Case 360 '6:00
            frmmain.PicMeteo.Picture = LoadInterface("a8.bmp")

        Case 420 '7:00
            frmmain.PicMeteo.Picture = LoadInterface("a9.bmp")

        Case 450 '7:30
            frmmain.PicMeteo.Picture = LoadInterface("a10.bmp")

        Case 480 '8:00
            frmmain.PicMeteo.Picture = LoadInterface("a11.bmp")

        Case 510 '8:30
            frmmain.PicMeteo.Picture = LoadInterface("a12.bmp")

        Case 540 '9:00
            frmmain.PicMeteo.Picture = LoadInterface("a13.bmp")

        Case 600 '10:00
            frmmain.PicMeteo.Picture = LoadInterface("a14.bmp")

        Case 660 '11:00
            frmmain.PicMeteo.Picture = LoadInterface("a15.bmp")

        Case 720 '12:00
            frmmain.PicMeteo.Picture = LoadInterface("a16.bmp")

        Case 780 '13:00
            frmmain.PicMeteo.Picture = LoadInterface("a17.bmp")

        Case 840 '14:00
            frmmain.PicMeteo.Picture = LoadInterface("a18.bmp")

        Case 900 '15:00
            frmmain.PicMeteo.Picture = LoadInterface("a19.bmp")

        Case 960 '16:00
            frmmain.PicMeteo.Picture = LoadInterface("a20.bmp")

        Case 1020 '17:00
            frmmain.PicMeteo.Picture = LoadInterface("a21.bmp")

        Case 1080 '18:00
            frmmain.PicMeteo.Picture = LoadInterface("a22.bmp")

        Case 1140 '19:00
            frmmain.PicMeteo.Picture = LoadInterface("a23.bmp")

        Case 1170 '19:30
            frmmain.PicMeteo.Picture = LoadInterface("a24.bmp")

        Case 1200 '20:00
            frmmain.PicMeteo.Picture = LoadInterface("a25.bmp")

        Case 1260 '21:00
            frmmain.PicMeteo.Picture = LoadInterface("a26.bmp")

        Case 1320 '22:00
            frmmain.PicMeteo.Picture = LoadInterface("a27.bmp")

        Case 1380 '23:00
            frmmain.PicMeteo.Picture = LoadInterface("a28.bmp")

    End Select

    If HoraFantasia - 1 = "300" Then
        Call Sound.Sound_Play(FXSound.Gallo_Sound, False, 0, 0)
    ElseIf HoraFantasia - 1 = "1260" Then
        Call Sound.Sound_Play(FXSound.Lobo_Sound, False, 0, 0)
    ElseIf HoraFantasia + 1 = "1440" Then
        HoraFantasia = 0

    End If

End Sub

Private Sub Image2_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)

    If Index = 0 Then
        If OpcionMenu <> 0 Then

            ' Image2(Index).Tag = "1"
            ' Image2(Index).Picture = LoadInterface("botoninventarioapretado.bmp")
            Rem    Image2(1).Picture = LoadInterface("botonconjuros.bmp")
            Rem   Image2(2).Picture = LoadInterface("botonmenu.bmp")
        End If

    End If

    If Index = 1 Then
        If OpcionMenu <> 1 Then

            ' Image2(Index).Tag = "1"
            '  Image2(1).Picture = LoadInterface("botonconjurosapretado.bmp")
            Rem   Image2(2).Picture = LoadInterface("botonmenu.bmp")
            Rem    Image2(0).Picture = LoadInterface("botoninventario.bmp")
        End If

    End If

End Sub

Private Sub Image2_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)

    If OpcionMenu = 1 Then
    
        If Image2(Index).Tag = "0" Then
            Image2(Index).Picture = LoadInterface("botonconjurosmarcado.bmp")
            Image2(Index).Tag = "1"

        End If

    Else

        If Image2(Index).Tag = "0" Then
            Image2(Index).Picture = LoadInterface("botoninventariomarcado.bmp")
            Image2(Index).Tag = "1"

        End If

    End If

End Sub

Private Sub Image3_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    '    Image3.Picture = LoadInterface("elegirchatapretado.bmp")
    frmMensaje.PopupMenuMensaje

End Sub

Private Sub Image3_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

    If Image3.Tag = "0" Then
        Image3.Picture = LoadInterface("elegirchatmarcado.bmp")
        Image3.Tag = "1"

    End If

End Sub

Private Sub Image4_Click(Index As Integer)

    Select Case Index

        Case 0
            Me.WindowState = vbMinimized

        Case 1
            frmCerrar.Show , frmmain

    End Select

End Sub

Private Sub Image4_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)

    Select Case Index

        Case 0
            Image4(0).Picture = LoadInterface("boton-sm-minimizar-off.bmp")

        Case 1
            Image4(1).Picture = LoadInterface("boton-sm-cerrar-off.bmp")

    End Select

End Sub

Private Sub Image4_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)

    Select Case Index

        Case 0

            If Image4(Index).Tag = "0" Then
                Image4(Index).Picture = LoadInterface("boton-sm-minimizar-over.bmp")
                Image4(Index).Tag = "1"
                Image4(1).Picture = Nothing

            End If

            If Image4(1).Tag = "1" Then
                Image4(1).Picture = Nothing
                Image4(1).Tag = "0"

            End If

        Case 1

            If Image4(Index).Tag = "0" Then
                Image4(Index).Picture = LoadInterface("boton-sm-cerrar-over.bmp")
                Image4(Index).Tag = "1"

            End If

            If Image4(0).Tag = "1" Then
                Image4(0).Picture = Nothing
                Image4(0).Tag = "0"

            End If

    End Select
         
End Sub

Private Sub Image5_Click()

    If FrmGrupo.Visible = False Then
        Call WriteRequestGrupo

    End If

End Sub

Private Sub Image5_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

    If Image5.Tag = "0" Then
        Image5.Picture = LoadInterface("grupoover.bmp")
        Image5.Tag = "1"

    End If

End Sub

Private Sub Image6_Click()
    Call WriteSafeToggle

End Sub

Private Sub imgBugReport_Click()
    FrmGmAyuda.Show vbModeless, frmmain

End Sub

Private Sub imgHechizos_Click()

    If hlst.Visible Then Exit Sub
    Panel.Picture = LoadInterface("centrohechizo.bmp")
    picInv.Visible = False
    hlst.Visible = True

    cmdlanzar.Visible = True

    imgSpellInfo.Visible = True

    cmdMoverHechi(0).Visible = True
    cmdMoverHechi(1).Visible = True

    frmmain.imgInvLock(0).Visible = False
    frmmain.imgInvLock(1).Visible = False
    frmmain.imgInvLock(2).Visible = False

End Sub

Private Sub imgHechizos_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    imgHechizos.Picture = LoadInterface("boton-hechizos-ES-off.bmp")
    imgHechizos.Tag = "1"

End Sub

Private Sub imgHechizos_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

    If imgHechizos.Tag = "0" Then
        imgHechizos.Picture = LoadInterface("boton-hechizos-ES-default.bmp")
        imgHechizos.Tag = "1"

    End If

End Sub

Private Sub imgInventario_Click()

    If picInv.Visible Then Exit Sub

    Panel.Picture = LoadInterface("centroinventario.bmp")
    'Call Audio.PlayWave(SND_CLICK)
    picInv.Visible = True
    hlst.Visible = False
    cmdlanzar.Visible = False
    imgSpellInfo.Visible = False

    cmdMoverHechi(0).Visible = False
    cmdMoverHechi(1).Visible = False
    Call Inventario.ReDraw
    frmmain.imgInvLock(0).Visible = True
    frmmain.imgInvLock(1).Visible = True
    frmmain.imgInvLock(2).Visible = True

End Sub

Private Sub imgInventario_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    imgInventario.Picture = LoadInterface("boton-inventory-ES-off.bmp")
    imgInventario.Tag = "1"

End Sub

Private Sub imgInventario_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call Inventario.ReDraw

    If imgInventario.Tag = "0" Then
        imgInventario.Picture = LoadInterface("boton-inventory-ES-over.bmp")
        imgInventario.Tag = "1"

    End If

End Sub

Private Sub LlamaDeclan_Timer()
    frmMapaGrande.llamadadeclan.Visible = False
    frmMapaGrande.Shape2.Visible = False
    LlamaDeclan.Enabled = False

End Sub

Private Sub manualboton_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

    If manualboton.Tag = "0" Then
        manualboton.Picture = LoadInterface("manualover.bmp")
        manualboton.Tag = "1"

    End If

End Sub

Private Sub manualboton_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Manual.Show , frmmain

End Sub

Private Sub OpcionesBoton_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    OpcionesBoton.Picture = LoadInterface("opcionesoverdown.bmp")
    OpcionesBoton.Tag = "1"

End Sub

Private Sub panelGM_Click()
    frmPanelGm.Show , Me
End Sub

Private Sub PicCorreo_Click()
    Call AddtoRichTextBox(frmmain.RecTxt, "Tenes un mensaje, ve al correo local para leerlo.", 255, 255, 255, False, False, False)

End Sub

Private Sub Inventario_ItemDropped(ByVal Drag As Integer, ByVal Drop As Integer, ByVal x As Integer, ByVal y As Integer)

    ' Si soltó un item en un slot válido
    
    If Drop > 0 Then
        ' Muevo el item dentro del iventario
        Call WriteItemMove(Drag, Drop)

    End If

End Sub

Private Sub PicSegClanOff_Click()
    Call WriteSeguroClan

End Sub

Private Sub PicSegClanOn_Click()
    Call WriteSeguroClan

End Sub

Private Sub QuestBoton_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

    If QuestBoton.Tag = "0" Then
        QuestBoton.Picture = LoadInterface("questover.bmp")
        QuestBoton.Tag = "1"

    End If

End Sub

Private Sub QuestBoton_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call WriteQuestListRequest

End Sub

Private Sub RankingBoton_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

    If rankingBoton.Tag = "0" Then
        rankingBoton.Picture = LoadInterface("rankingover.bmp")
        rankingBoton.Tag = "1"

    End If

End Sub

Private Sub RankingBoton_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call WriteTraerRanking

End Sub

Private Sub Label6_Click()
    Inventario.SelectGold

    If UserGLD > 0 Then
        frmCantidad.Picture = LoadInterface("cantidad.bmp")
        HayFormularioAbierto = True
        frmCantidad.Show , frmmain

    End If

End Sub

Private Sub Label7_Click()
    Call AddtoRichTextBox(frmmain.RecTxt, "No tenes mensajes nuevos.", 255, 255, 255, False, False, False)

End Sub

Private Sub lblPorcLvl_Click()
    Call WriteScrollInfo

End Sub

Private Sub MacroLadder_Timer()

    If MainTimer.Check(TimersIndex.Work) Then
        If UserMacro.cantidad > 0 And UserMacro.Activado And UserMinSTA > 0 Then
        
            Select Case UserMacro.TIPO

                Case 1 'Alquimia
                    Call WriteCraftAlquimista(UserMacro.Index)

                Case 2 'Carpinteria
                    Call WriteCraftCarpenter(UserMacro.Index)

                Case 3 'Sasteria
                    Call WriteCraftSastre(UserMacro.Index)

                Case 4 'Herreria
                    Call WriteCraftBlacksmith(UserMacro.Index)

                Case 6
                    Call WriteWorkLeftClick(TargetXMacro, TargetYMacro, UsingSkill)

            End Select

        Else
            Call ResetearUserMacro

        End If
    
        UserMacro.cantidad = UserMacro.cantidad - 1
        
    End If

End Sub

Private Sub macrotrabajo_Timer()
    'If Inventario.SelectedItem = 0 Then
    '   DesactivarMacroTrabajo
    '   Exit Sub
    'End If
    
    'Macros are disabled if not using Argentum!
    If Not Application.IsAppActive() Then
        DesactivarMacroTrabajo
        Exit Sub

    End If
    
    If UsingSkill = FundirMetal Or (UsingSkill = eSkill.Herreria And Not frmHerrero.Visible) Then
        Call WriteWorkLeftClick(TargetXMacro, TargetYMacro, UsingSkill)
        UsingSkill = 0

    End If
    
    'If Inventario.OBJType(Inventario.SelectedItem) = eObjType.otWeapon Then
    If Not (frmCarp.Visible = True) Then Call WriteUseItem(frmmain.Inventario.SelectedItem)

End Sub

Public Sub ActivarMacroTrabajo()
    TargetXMacro = tX
    TargetYMacro = tY
    macrotrabajo.Interval = IntervaloTrabajo
    macrotrabajo.Enabled = True
    Call AddtoRichTextBox(frmmain.RecTxt, "Macro Trabajo ACTIVADO", 0, 200, 200, False, True, False)

End Sub

Public Sub DesactivarMacroTrabajo()
    TargetXMacro = 0
    TargetYMacro = 0
    macrotrabajo.Enabled = False
    MacroBltIndex = 0
    UsingSkill = 0
    MousePointer = vbDefault
    Call AddtoRichTextBox(frmmain.RecTxt, "Macro Trabajo DESACTIVADO", 0, 200, 200, False, True, False)

End Sub

Private Sub MenuOpciones_Click()

End Sub

Private Sub manabar_Click()

    'If UserMinMAN = UserMaxMAN Then Exit Sub
            
    If UserEstado = 1 Then

        With FontTypes(FontTypeNames.FONTTYPE_INFO)
            Call ShowConsoleMsg("¡Estás muerto!", .red, .green, .blue, .bold, .italic)

        End With

        Exit Sub

    End If
    
    Call WriteMeditate

End Sub

Private Sub mapMundo_Click()
    ExpMult = 1
    OroMult = 1
    Call CalcularPosicionMAPA
    frmMapaGrande.Picture = LoadInterface("ventanamapa.bmp")
    HayFormularioAbierto = True
    frmMapaGrande.Show , frmmain

End Sub

Private Sub mapMundo_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

    If mapMundo.Tag = "0" Then
        mapMundo.Picture = LoadInterface("boton-mapa-over.bmp")
        mapMundo.Tag = "1"

    End If

End Sub

Private Sub MiniMap_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

    If Button = vbRightButton Then
        Call ParseUserCommand("/TELEP YO " & UserMap & " " & CByte(x) & " " & CByte(y))
        Exit Sub

    End If
  
    ExpMult = 1
    OroMult = 1
    Call CalcularPosicionMAPA
    frmMapaGrande.Picture = LoadInterface("ventanamapa.bmp")
    HayFormularioAbierto = True
    frmMapaGrande.Show , frmmain
  
End Sub

Private Sub MiniMap_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

    If mapMundo.Tag = "1" Then
        mapMundo.Picture = Nothing
        mapMundo.Tag = "0"

    End If

End Sub

Private Sub mnuEquipar_Click()

    If MainTimer.Check(TimersIndex.UseItemWithU) Then Call WriteEquipItem(Inventario.SelectedItem)

End Sub

Private Sub mnuNPCComerciar_Click()
    Call WriteLeftClick(tX, tY)
    Call WriteCommerceStart

End Sub

Private Sub mnuNpcDesc_Click()
    Call WriteLeftClick(tX, tY)

End Sub

Private Sub mnuTirar_Click()
    Call TirarItem

End Sub

Private Sub mnuUsar_Click()

    If MainTimer.Check(TimersIndex.UseItemWithDblClick) Then Call WriteUseItem(frmmain.Inventario.SelectedItem)

End Sub

Private Sub NameMapa_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

    If mapMundo.Tag = "1" Then
        mapMundo.Picture = Nothing
        mapMundo.Tag = "0"

    End If

End Sub

Private Sub onlines_Click()
    Call WriteOnline

End Sub

Private Sub OpcionesBoton_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

    If OpcionesBoton.Tag = "0" Then
        OpcionesBoton.Picture = LoadInterface("opcionesover.bmp")
        OpcionesBoton.Tag = "1"

    End If

End Sub

Private Sub OpcionesBoton_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call frmOpciones.Init

End Sub

Private Sub Panel_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

    ObjLbl.Visible = False
    
    Select Case OpcionMenu

        Case 0

            If Image2(1).Tag = "1" Then
                Image2(1).Picture = Nothing
                Image2(1).Tag = "0"

            End If
    
        Case 1

            If Image2(1).Tag = "1" Then
                Image2(1).Picture = Nothing
                Image2(1).Tag = "0"

            End If

        Case 2

            If Image2(1).Tag = "1" Then
                Image2(1).Picture = Nothing
                Image2(1).Tag = "0"

            End If

            If Image2(0).Tag = "1" Then
                Image2(0).Picture = Nothing
                Image2(0).Tag = "0"

            End If

    End Select

    If cmdlanzar.Tag = "1" Then
        cmdlanzar.Picture = Nothing
        cmdlanzar.Tag = "0"

    End If
    
    If imgInventario.Tag = "1" Then
        imgInventario.Picture = Nothing
        imgInventario.Tag = "0"

    End If

    If imgHechizos.Tag = "1" Then
        imgHechizos.Picture = Nothing
        imgHechizos.Tag = "0"

    End If

    If cmdMoverHechi(1).Tag = "1" Then
        cmdMoverHechi(1).Picture = Nothing
        cmdMoverHechi(1).Tag = "0"

    End If
    
    If cmdMoverHechi(0).Tag = "1" Then
        cmdMoverHechi(0).Picture = Nothing
        cmdMoverHechi(0).Tag = "0"

    End If

End Sub

Private Sub picInv_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

    Dim slot As Byte

    UsaMacro = False
    
    slot = Inventario.GetSlot(x, y)
    
    If Inventario.Amount(slot) > 0 Then
    
        ObjLbl.Visible = True
        
        Select Case ObjData(Inventario.OBJIndex(slot)).ObjType

            Case eObjType.otWeapon
                ObjLbl = Inventario.ItemName(slot) & " (" & Inventario.Amount(slot) & ")" & vbCrLf & "Daño: " & ObjData(Inventario.OBJIndex(slot)).MinHit & "/" & ObjData(Inventario.OBJIndex(slot)).MaxHit

            Case eObjType.otArmadura
                ObjLbl = Inventario.ItemName(slot) & " (" & Inventario.Amount(slot) & ")" & vbCrLf & "Defensa: " & ObjData(Inventario.OBJIndex(slot)).MinDef & "/" & ObjData(Inventario.OBJIndex(slot)).MaxDef

            Case eObjType.otCASCO
                ObjLbl = Inventario.ItemName(slot) & " (" & Inventario.Amount(slot) & ")" & vbCrLf & "Defensa: " & ObjData(Inventario.OBJIndex(slot)).MinDef & "/" & ObjData(Inventario.OBJIndex(slot)).MaxDef

            Case eObjType.otESCUDO
                ObjLbl = Inventario.ItemName(slot) & " (" & Inventario.Amount(slot) & ")" & vbCrLf & "Defensa: " & ObjData(Inventario.OBJIndex(slot)).MinDef & "/" & ObjData(Inventario.OBJIndex(slot)).MaxDef

            Case Else
                ObjLbl = Inventario.ItemName(slot) & " (" & Inventario.Amount(slot) & ")" & vbCrLf & ObjData(Inventario.OBJIndex(slot)).Texto

        End Select
        
        If Len(ObjLbl.Caption) < 100 Then
            ObjLbl.FontSize = 7
            
        ElseIf Len(ObjLbl.Caption) > 100 And Len(ObjLbl.Caption) < 150 Then
            ObjLbl.FontSize = 6

            '
            ' Else
            '  ObjLbl.FontSize = 5
        End If

    Else
        ObjLbl.Visible = False

    End If

End Sub

Private Sub PicMeteo_Click()

    With FontTypes(FontTypeNames.FONTTYPE_INFOIAO)
        Call ShowConsoleMsg("Servidor> Son las " & Meteo_Engine.Get_Time_String, .red, .green, .blue, .bold, .italic)

    End With

End Sub

Private Sub PicResu_Click()
    Call WriteParyToggle

End Sub

Private Sub PicResuOn_Click()
    Call WriteParyToggle

End Sub

Private Sub PicSeg_Click()
    Call WriteSafeToggle

End Sub

Private Sub CompletarEnvioMensajes()

    Select Case SendingType

        Case 1
            SendTxt.Text = ""

        Case 2
            SendTxt.Text = "-"

        Case 3
            SendTxt.Text = ("\" & sndPrivateTo & " ")

        Case 4
            SendTxt.Text = "/CMSG "

        Case 5
            SendTxt.Text = "/GRUPO "

        Case 6
            SendTxt.Text = "/GRMG "

        Case 7
            SendTxt.Text = ";"

        Case 8
            SendTxt.Text = "/RMSG "

    End Select

    stxtbuffer = SendTxt.Text
    SendTxt.SelStart = Len(SendTxt.Text)

End Sub

Private Sub RecTxt_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

    On Error Resume Next

    If Button = 1 Then

        Dim strBuffer      As String

        Dim lngLength      As Long

        Dim intCurrentLine As Integer
    
        intCurrentLine = RecTxt.GetLineFromChar(RecTxt.SelStart)
        'get line length
        lngLength = SendMessage(RecTxt.hwnd, EM_LINELENGTH, intCurrentLine, 0)
        'resize buffer
        strBuffer = Space(lngLength)
        'get line text
        Call SendMessage(RecTxt.hwnd, EM_GETLINE, intCurrentLine, ByVal strBuffer)

        Dim partea       As String

        Dim destinatario As String
    
        destinatario = SuperMid(strBuffer, "[", "]", False)

        If destinatario <> "A" Then

            destinatario = Replace(destinatario, " ", "+")

            sndPrivateTo = destinatario
            SendTxt.Text = ("\" & sndPrivateTo & " ")

            stxtbuffer = SendTxt.Text
            SendTxt.SelStart = Len(SendTxt.Text)

            If SendTxt.Visible = False Then
                Call WriteEscribiendo

            End If

            SendTxt.Visible = True
            SendTxt.SetFocus

        End If

    End If

End Sub

Private Sub refuerzolanzar_Click()
    Call cmdLanzar_Click

End Sub

Private Sub refuerzolanzar_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    UsaMacro = False
    CnTd = 0

    If cmdlanzar.Tag = "0" Then
        cmdlanzar.Picture = LoadInterface("lanzarmarcado.bmp")
        cmdlanzar.Tag = "1"

    End If

End Sub

Private Sub renderer_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

    'If DropItem Then
    '    frmMain.UsandoDrag = False
    '    DropItem = False
    '    DropIndex = 0
    '    DropActivo = False
    '    Call FormParser.Parse_Form(Me)
    'End If

    clicX = x
    clicY = y
    
    Dim PosX As Integer

    Dim PosY As Integer

    If Pregunta Then
        If x > 395 And x < 410 And y > 233 And y < 250 Then
            If PreguntaLocal Then

                Select Case PreguntaNUM

                    Case 1
                        Pregunta = False
                        DestItemSlot = 0
                        DestItemCant = 0
                        PreguntaLocal = False

                End Select

            Else
                Call WriteResponderPregunta(False)
                Pregunta = False

            End If

        End If
    
        If x > 417 And x < 439 And y > 233 And y < 250 Then
            If PreguntaLocal Then

                Select Case PreguntaNUM

                    Case 1 '¿Destruir item?
                        Call WriteDrop(DestItemSlot, DestItemCant)
                        Pregunta = False
                        PreguntaLocal = False

                End Select

            Else
                Call WriteResponderPregunta(True)
                Pregunta = False

            End If

        End If

    End If
    
End Sub

Private Sub renderer_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    MouseX = x
    MouseY = y
    
    DisableURLDetect

    If cmdlanzar.Tag = "1" Then
        cmdlanzar.Picture = Nothing
        cmdlanzar.Tag = "0"

    End If

    'If DropItem Then

    ' frmMain.UsandoDrag = False
    ' Call ConvertCPtoTP(MouseX, MouseY, tX, tY)
    'Call WriteDropItem(DropIndex, tX, tY, CantidadDrop)
    ' DropItem = False
    ' DropIndex = 0
    ' TimeDrop = 0
    ' DropActivo = False
    ' CantidadDrop = 0
    ' Call FormParser.Parse_Form(frmMain)
    
    ' End If
    
    'engine.Light_Remove (10)
    
    'engine.Light_Create tX, tY, &HFFFFFFF, 1, 10
    'engine.Light_Render_All
    
End Sub

Private Sub renderer_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

    MouseBoton = Button
    MouseShift = Shift
    
    If HayFormularioAbierto Then
        If frmComerciar.Visible Then
            Unload frmComerciar
            HayFormularioAbierto = False
            Exit Sub

        End If
            
        If frmBancoObj.Visible Then
            Unload frmBancoObj
            HayFormularioAbierto = False
            Exit Sub

        End If
        
        If frmEstadisticas.Visible Then
            Unload frmEstadisticas
            HayFormularioAbierto = False
            Exit Sub

        End If
        
        If frmGoliath.Visible Then
            Unload frmGoliath
            HayFormularioAbierto = False
            Exit Sub

        End If
        
        If frmMapaGrande.Visible Then
            Unload frmMapaGrande
            HayFormularioAbierto = False
            Exit Sub

        End If
            
        If FrmViajes.Visible Then
            Unload FrmViajes
            HayFormularioAbierto = False
            Exit Sub

        End If
        
        If frmCantidad.Visible Then
            Unload frmCantidad
            HayFormularioAbierto = False
            Exit Sub

        End If
        
        If FrmGrupo.Visible Then
            Unload FrmGrupo
            HayFormularioAbierto = False
            Exit Sub

        End If
        
        If FrmGmAyuda.Visible Then
            Unload FrmGmAyuda
            HayFormularioAbierto = False
            Exit Sub

        End If
        
        If frmGuildAdm.Visible Then
            Unload frmGuildAdm
            HayFormularioAbierto = False
            Exit Sub

        End If
        
        If FrmShop.Visible Then
            Unload FrmShop
            HayFormularioAbierto = False
            Exit Sub

        End If
        
        If frmHerrero.Visible Then
            Unload frmHerrero
            HayFormularioAbierto = False
            Exit Sub

        End If
        
        If FrmSastre.Visible Then
            Unload FrmSastre
            HayFormularioAbierto = False
            Exit Sub

        End If

        If frmAlqui.Visible Then
            Unload frmAlqui
            HayFormularioAbierto = False
            Exit Sub

        End If

        If frmCarp.Visible Then
            Unload frmCarp
            HayFormularioAbierto = False
            Exit Sub

        End If
        
        If FrmCorreo.Visible Then
            Unload FrmCorreo
            HayFormularioAbierto = False
            Exit Sub

        End If
                
    End If

End Sub

Private Sub renderer_DblClick()
    Call Form_DblClick

End Sub

Private Sub renderer_Click()
    Call Form_Click

End Sub

Private Sub SendTxt_KeyUp(KeyCode As Integer, Shift As Integer)

    On Error Resume Next

    Dim str1 As String

    Dim str2 As String

    'Send text
    If KeyCode = vbKeyReturn Then
        If LenB(stxtbuffer) <> 0 Then
        
            ' If Right$(stxtbuffer, 1) = " " Or left(stxtbuffer, 1) = " " Then
            ' stxtbuffer = Trim(stxtbuffer)
            ' End If
        
            If Left$(stxtbuffer, 1) = "/" Then
                If UCase$(Left$(stxtbuffer, 7)) = "/GRUPO " Then
                    SendingType = 5
                ElseIf UCase$(Left$(stxtbuffer, 6)) = "/CMSG " Then
                    SendingType = 4
                ElseIf UCase$(Left$(stxtbuffer, 6)) = "/GRMG " Then
                    SendingType = 6
                ElseIf UCase$(Left$(stxtbuffer, 6)) = "/RMSG " Then
                    SendingType = 8
                Else
                    SendingType = 1

                End If
            
                If stxtbuffer <> "" Then Call ParseUserCommand(stxtbuffer)
    
                'Shout
            ElseIf Left$(stxtbuffer, 1) = "-" Then

                If Right$(stxtbuffer, Len(stxtbuffer) - 1) <> "" Then Call ParseUserCommand("-" & Right$(stxtbuffer, Len(stxtbuffer) - 1))
                SendingType = 2
            
                'Global
            ElseIf Left$(stxtbuffer, 1) = ";" Then

                If Right$(stxtbuffer, Len(stxtbuffer) - 1) <> "" Then Call ParseUserCommand("/CONSOLA " & Right$(stxtbuffer, Len(stxtbuffer) - 1))
                SendingType = 7
                sndPrivateTo = ""
            
            ElseIf Left$(stxtbuffer, 1) = "/RMSG" Then

                If Right$(stxtbuffer, Len(stxtbuffer) - 1) <> "" Then Call ParseUserCommand("/RMSG " & Right$(stxtbuffer, Len(stxtbuffer) - 1))
                SendingType = 8
                sndPrivateTo = ""

                'Privado
            ElseIf Left$(stxtbuffer, 1) = "\" Then

                Dim mensaje As String
 
                str1 = Right$(stxtbuffer, Len(stxtbuffer) - 1)
                str2 = ReadField(1, str1, 32)
                mensaje = Right$(stxtbuffer, Len(str1) - Len(str2) - 1)
                sndPrivateTo = str2
                SendingType = 3
    
                If str1 <> "" Then Call WriteWhisper(sndPrivateTo, mensaje)
                    
                'Say
            Else

                If stxtbuffer <> "" Then Call ParseUserCommand(stxtbuffer)
                SendingType = 1
                sndPrivateTo = ""

            End If

        Else
            SendingType = 1
            sndPrivateTo = ""

        End If
        
        stxtbuffer = ""
        SendTxt.Text = ""
        KeyCode = 0
        SendTxt.Visible = False
        Call WriteEscribiendo
        
    End If

End Sub

Private Sub Second_Timer()

    If engine.bRunning Then engine.Engine_ActFPS

    Rem   If Not DialogosClanes Is Nothing Then DialogosClanes.PassTimer
End Sub

Private Sub Socket1_Timeout(status As Integer, Response As Integer)
    MsgBox "Se perdio la conexion time out"

End Sub

Private Sub TiendaBoton_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

    If TiendaBoton.Tag = "0" Then
        TiendaBoton.Picture = LoadInterface("tiendaover.bmp")
        TiendaBoton.Tag = "1"

    End If

End Sub

Private Sub TiendaBoton_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call WriteTraerShop

End Sub

Private Sub cerrarcuenta_Timer()
    Unload frmConnect
    Unload frmCrearPersonaje
    cerrarcuenta.Enabled = False

End Sub

Private Sub TimerLluvia_Timer()

    If bRain Then

        If CantPartLLuvia < 250 Then

            CantPartLLuvia = CantPartLLuvia + 1
            engine.Particle_Group_Edit (MeteoIndex)
        Else
            CantPartLLuvia = 250
            TimerLluvia.Enabled = False

        End If

    Else

        If CantPartLLuvia > 0 Then
            CantPartLLuvia = CantPartLLuvia - 1
            engine.Particle_Group_Edit (MeteoIndex)
        Else
    
            engine.Engine_Meteo_Particle_Set (-1)
            CantPartLLuvia = 0
            TimerLluvia.Enabled = False

        End If

    End If

End Sub

Private Sub TimerMusica_Timer()

End Sub

Private Sub TimerNiebla_Timer()

    If bNiebla Then

        If AlphaNiebla < MaxAlphaNiebla Then
            AlphaNiebla = AlphaNiebla + 1
        Else
            AlphaNiebla = MaxAlphaNiebla
            TimerNiebla.Enabled = False

        End If

    Else

        If AlphaNiebla > 0 Then
            AlphaNiebla = AlphaNiebla - 1

        Else
            AlphaNiebla = 0
            MaxAlphaNiebla = 0
            TimerNiebla.Enabled = False
            Call Meteo_Engine.CargarClima

        End If

    End If

End Sub

Private Sub Timerping_Timer()
    Call WritePing

End Sub

Private Sub cmdLanzar_Click()

    If hlst.List(hlst.ListIndex) <> "(Vacio)" And MainTimer.Check(TimersIndex.Work, False) Then
        If UserEstado = 1 Then

            With FontTypes(FontTypeNames.FONTTYPE_INFO)
                Call ShowConsoleMsg("¡¡Estás muerto!!", .red, .green, .blue, .bold, .italic)

            End With

        Else
            Call WriteCastSpell(hlst.ListIndex + 1)
            'Call WriteWork(eSkill.Magia)
            UsaMacro = True
            UsaLanzar = True

        End If

    End If

End Sub

Private Sub CmdLanzar_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    UsaMacro = False
    CnTd = 0

    If cmdlanzar.Tag = "0" Then
        cmdlanzar.Picture = LoadInterface("boton-lanzar-ES-over.bmp")
        cmdlanzar.Tag = "1"

    End If

End Sub

Public Sub Form_Click()

    If MouseBoton = vbLeftButton And ACCION1 = 0 Or MouseBoton = vbRightButton And ACCION2 = 0 Or MouseBoton = 4 And ACCION3 = 0 Then
        If Not Comerciando Then
            ' Fix: game area esta mal
            'If Not InGameArea() Then Exit Sub

            If MouseShift = 0 Then
                If UsingSkill = 0 Then
                    Call WriteLeftClick(tX, tY)
                Else

                    If macrotrabajo.Enabled Then DesactivarMacroTrabajo
                    
                    Dim SendSkill As Boolean
                    
                    If UsingSkill = magia Then
                        If MainTimer.Check(TimersIndex.AttackSpell, False) Then
                            If MainTimer.Check(TimersIndex.CastSpell) Then
                                SendSkill = True
                                Call MainTimer.Restart(TimersIndex.CastAttack)
                            Else

                                With FontTypes(FontTypeNames.FONTTYPE_TALK)
                                    Call ShowConsoleMsg("¡No puedes lanzar hechizos tan rápido!", .red, .green, .blue, .bold, .italic)

                                End With

                            End If

                        Else

                            With FontTypes(FontTypeNames.FONTTYPE_TALK)
                                Call ShowConsoleMsg("¡No puedes lanzar tan rápido después de un golpe!", .red, .green, .blue, .bold, .italic)

                            End With

                        End If

                    End If
                    
                    'Splitted because VB isn't lazy!
                    If UsingSkill = Proyectiles Then
                        If MainTimer.Check(TimersIndex.AttackSpell, False) Then
                            If MainTimer.Check(TimersIndex.CastAttack, False) Then
                                If MainTimer.Check(TimersIndex.Arrows) Then
                                    SendSkill = True
                                    Call MainTimer.Restart(TimersIndex.Attack) ' Prevengo flecha-golpe
                                    Call MainTimer.Restart(TimersIndex.CastSpell) ' flecha-hechizo

                                End If

                            End If

                        End If

                    End If
                
                    'Splitted because VB isn't lazy!
                    If (UsingSkill = Robar Or UsingSkill = Grupo Or UsingSkill = MarcaDeClan Or UsingSkill = MarcaDeGM) Then
                        If MainTimer.Check(TimersIndex.Work) Then
                            If UsingSkill = MarcaDeGM Then

                                Dim Pos As Integer

                                If MapData(tX, tY).charindex <> 0 Then
                                    Pos = InStr(charlist(MapData(tX, tY).charindex).nombre, "<")
                                
                                    If Pos = 0 Then Pos = LenB(charlist(MapData(tX, tY).charindex).nombre) + 2
                                    frmPanelGm.cboListaUsus.Text = Left$(charlist(MapData(tX, tY).charindex).nombre, Pos - 2)

                                End If

                            Else
                                SendSkill = True

                            End If

                        End If

                    End If
                    
                    If (UsingSkill = eSkill.Pescar Or UsingSkill = eSkill.Talar Or UsingSkill = eSkill.Mineria Or UsingSkill = FundirMetal) Then
                        
                        If MainTimer.Check(TimersIndex.Work) Then
                            Call WriteWorkLeftClick(tX, tY, UsingSkill)
                            Call FormParser.Parse_Form(frmmain)

                            If CursoresGraficos = 0 Then
                                frmmain.MousePointer = vbDefault

                            End If

                            Exit Sub

                        End If

                    End If
                   
                    If SendSkill Then
                        Call WriteWorkLeftClick(tX, tY, UsingSkill)

                    End If
                   
                    If OcultarMacro Then
                        OcultarMacro = False

                    End If
                    
                    Call FormParser.Parse_Form(frmmain)

                    If CursoresGraficos = 0 Then
                        frmmain.MousePointer = vbDefault

                    End If
                    
                    UsaLanzar = False
                    UsingSkill = 0

                End If

            Else
                Call WriteWarpChar("YO", UserMap, tX, tY)

            End If
            
            If cartel Then cartel = False
            
        End If
    
    ElseIf MouseBoton = vbLeftButton And ACCION1 = 1 Or MouseBoton = vbRightButton And ACCION2 = 1 Or MouseBoton = 4 And ACCION3 = 1 Then
        Call WriteDoubleClick(tX, tY)
    
    ElseIf MouseBoton = vbLeftButton And ACCION1 = 2 Or MouseBoton = vbRightButton And ACCION2 = 2 Or MouseBoton = 4 And ACCION3 = 2 Then

        If UserDescansar Or UserMeditar Then Exit Sub
        If MainTimer.Check(TimersIndex.CastAttack, False) Then
            If MainTimer.Check(TimersIndex.Attack) Then
                Call MainTimer.Restart(TimersIndex.AttackSpell)
                Call WriteAttack

            End If

        End If
    
    ElseIf MouseBoton = vbLeftButton And ACCION1 = 3 Or MouseBoton = vbRightButton And ACCION2 = 3 Or MouseBoton = 4 And ACCION3 = 3 Then

        If MainTimer.Check(TimersIndex.UseItemWithU) Then Call WriteUseItem(frmmain.Inventario.SelectedItem)
    
    ElseIf MouseBoton = vbLeftButton And ACCION1 = 4 Or MouseBoton = vbRightButton And ACCION2 = 4 Or MouseBoton = 4 And ACCION3 = 4 Then

        If MapData(tX, tY).charindex <> 0 Then
            If charlist(MapData(tX, tY).charindex).nombre <> charlist(MapData(UserPos.x, UserPos.y).charindex).nombre Then
                If charlist(MapData(tX, tY).charindex).EsNpc = False Then
                    SendTxt.Text = "\" & charlist(MapData(tX, tY).charindex).nombre & " "

                    If SendTxt.Visible = False Then
                        Call WriteEscribiendo

                    End If

                    SendTxt.Visible = True
                    SendTxt.SetFocus
                    SendTxt.SelStart = Len(SendTxt.Text)

                End If

            End If

        End If

    End If

End Sub

Private Sub Form_DblClick()

    '**************************************************************
    'Author: Unknown
    'Last Modify Date: 12/27/2007
    '12/28/2007: ByVal - Chequea que la ventana de comercio y boveda no este abierta al hacer doble clic a un comerciante, sobrecarga la lista de items.
    '**************************************************************
    If Not frmComerciar.Visible And Not frmBancoObj.Visible Then
        If MouseBoton = vbLeftButton Then

            'Call WriteDoubleClick(tX, tY)
        End If

    End If

End Sub

Private Sub Form_Load()

    On Error Resume Next

    Call FormParser.Parse_Form(frmmain)
    MenuNivel = 1
    Me.Caption = "Argentum20" 'hay que poner 20 aniversario
    
#If DEBUGGING = 1 Then
    panelGM.Visible = True
    createObj.Visible = True
#End If

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

    ' Disable links checking (not over consola)
    StopCheckingLinks
    
    If PantallaCompleta = 0 Then
        If MoverVentana = 1 Then
            If UserMoving = 0 Then
                moverForm

                'Call Auto_Drag(Me.hwnd)
            End If

        End If

    End If

    MouseX = x - MainViewShp.Left
    MouseY = y - MainViewShp.Top
    ObjLbl.Visible = False
    
    If EstadisticasBoton.Tag = "1" Then
        EstadisticasBoton.Picture = Nothing
        EstadisticasBoton.Tag = "0"

    End If
    
    If cmdlanzar.Tag = "1" Then
        cmdlanzar.Picture = Nothing
        cmdlanzar.Tag = "0"

    End If

    If imgInventario.Tag = "1" Then
        imgInventario.Picture = Nothing
        imgInventario.Tag = "0"

    End If

    If imgHechizos.Tag = "1" Then
        imgHechizos.Picture = Nothing
        imgHechizos.Tag = "0"

    End If
 
    If Image4(0).Tag = "1" Then
        Image4(0).Picture = Nothing
        Image4(0).Tag = "0"

    End If

    If Image4(1).Tag = "1" Then
        Image4(1).Picture = Nothing
        Image4(1).Tag = "0"

    End If

    If Image3.Tag = "1" Then
        Image3.Picture = Nothing
        Image3.Tag = "0"

    End If

    If Image5.Tag = "1" Then
        Image5.Picture = Nothing
        Image5.Tag = "0"

    End If

    If clanimg.Tag = "1" Then
        clanimg.Picture = Nothing
        clanimg.Tag = "0"

    End If

    If mapMundo.Tag = "1" Then
        mapMundo.Picture = Nothing
        mapMundo.Tag = "0"

    End If

    If Image4(1).Tag = "1" Then
        Image4(1).Picture = Nothing
        Image4(1).Tag = "0"

    End If

    If OpcionesBoton.Tag = "1" Then
        OpcionesBoton.Picture = Nothing
        OpcionesBoton.Tag = "0"

    End If

    If QuestBoton.Tag = "1" Then
        QuestBoton.Picture = Nothing
        QuestBoton.Tag = "0"

    End If

    Select Case OpcionMenu

        Case 0

            If Image2(1).Tag = "1" Then
                Image2(1).Picture = Nothing
                Image2(1).Tag = "0"

            End If

            If Image4(1).Tag = "1" Then
                Image4(1).Picture = Nothing
                Image4(1).Tag = "0"

            End If

        Case 1
        
        Case 2

            If Image2(1).Tag = "1" Then
                Image2(1).Picture = Nothing
                Image2(1).Tag = "0"

            End If

            If Image2(0).Tag = "1" Then
                Image2(0).Picture = Nothing
                Image2(0).Tag = "0"

            End If

    End Select

End Sub

Private Sub hlst_KeyDown(KeyCode As Integer, Shift As Integer)
    KeyCode = 0

End Sub

Private Sub hlst_KeyPress(KeyAscii As Integer)
    KeyAscii = 0

End Sub

Private Sub hlst_KeyUp(KeyCode As Integer, Shift As Integer)
    KeyCode = 0

End Sub

Private Sub picInv_DblClick()

    If frmCarp.Visible Or frmHerrero.Visible Or frmComerciar.Visible Or frmBancoObj.Visible Then Exit Sub
    
    If UserMeditar Then Exit Sub
    If Not MainTimer.Check(TimersIndex.UseItemWithDblClick) Then Exit Sub
    
    If macrotrabajo.Enabled Then DesactivarMacroTrabajo
    
    If Inventario.SelectedItem <= 0 Then Exit Sub

    ' Hacemos acción del doble clic correspondiente
    Dim ObjType As Byte

    ObjType = ObjData(Inventario.OBJIndex(Inventario.SelectedItem)).ObjType

    Select Case ObjType

        Case eObjType.otArmadura, eObjType.otESCUDO, eObjType.otmagicos, eObjType.otFlechas, eObjType.otCASCO, eObjType.otNudillos
            Call WriteEquipItem(Inventario.SelectedItem)
            
        Case eObjType.otWeapon

            If ObjData(Inventario.OBJIndex(Inventario.SelectedItem)).proyectil = 1 And Inventario.Equipped(Inventario.SelectedItem) Then
                Call WriteUseItem(Inventario.SelectedItem)
            Else
                Call WriteEquipItem(Inventario.SelectedItem)

            End If
            
        Case eObjType.OtHerramientas

            If Inventario.Equipped(Inventario.SelectedItem) Then
                Call WriteUseItem(Inventario.SelectedItem)
            Else
                Call WriteEquipItem(Inventario.SelectedItem)

            End If
                
        Case Else
            Call WriteUseItem(Inventario.SelectedItem)

    End Select

End Sub

Private Sub RecTxt_Change()
    Exit Sub

    On Error Resume Next  'el .SetFocus causaba errores al salir y volver a entrar

    If Not Application.IsAppActive() Then Exit Sub
    
    If SendTxt.Visible Then
        SendTxt.SetFocus
    ElseIf (Not frmComerciar.Visible) And (Not frmComerciarUsu.Visible) And (Not frmBancoObj.Visible) And (Not frmPanelGm.Visible) And (Not frmEstadisticas.Visible) And (Not frmCantidad.Visible) And (picInv.Visible) And (Not frmGoliath.Visible) And (Not FrmGmAyuda.Visible) Then
        picInv.SetFocus

    End If

End Sub

Private Sub RecTxt_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    StartCheckingLinks

End Sub

Private Sub RecTxt_KeyDown(KeyCode As Integer, Shift As Integer)

    If picInv.Visible Then
        picInv.SetFocus
    Else

        'hlst.SetFocus
    End If

End Sub

Private Sub SendTxt_Change()

    '**************************************************************
    'Author: Unknown
    'Last Modify Date: 3/06/2006
    '3/06/2006: Maraxus - impedí se inserten caractéres no imprimibles
    '**************************************************************
    If Len(SendTxt.Text) > 160 Then
        stxtbuffer = "Soy un cheater, avisenle a un gm"
    Else

        'Make sure only valid chars are inserted (with Shift + Insert they can paste illegal chars)
        Dim i         As Long

        Dim tempstr   As String

        Dim CharAscii As Integer
        
        For i = 1 To Len(SendTxt.Text)
            CharAscii = Asc(mid$(SendTxt.Text, i, 1))

            If CharAscii >= vbKeySpace And CharAscii <= 250 Then
                tempstr = tempstr & Chr$(CharAscii)

            End If

        Next i
        
        If tempstr <> SendTxt.Text Then
            'We only set it if it's different, otherwise the event will be raised
            'constantly and the client will crush
            SendTxt.Text = tempstr

        End If
        
        stxtbuffer = SendTxt.Text

    End If

End Sub

Private Sub SendTxt_KeyPress(KeyAscii As Integer)

    If Not (KeyAscii = vbKeyBack) And Not (KeyAscii >= vbKeySpace And KeyAscii <= 250) Then KeyAscii = 0

End Sub

''''''''''''''''''''''''''''''''''''''
'     SOCKET1                        '
''''''''''''''''''''''''''''''''''''''
#If UsarWrench = 1 Then
    Private Sub Socket1_Connect()
        Socket1.NoDelay = True
    
        'Clean input and output buffers
        Call incomingData.ReadASCIIStringFixed(incomingData.length)
        Call outgoingData.ReadASCIIStringFixed(outgoingData.length)

        'Security.Redundance = 13 'DEFAULT
    
        Second.Enabled = True

        Select Case EstadoLogin

            Case E_MODO.CrearNuevoPj, E_MODO.Normal, E_MODO.Dados
                Call Login
            
            Case E_MODO.CreandoCuenta
                Call WriteNuevaCuenta
          
            Case E_MODO.ActivandoCuenta
                Call WriteValidarCuenta
          
            Case E_MODO.IngresandoConCuenta
                Call WriteIngresandoConCuenta
          
            Case E_MODO.ReValidandoCuenta
                Call WriteReValidarCuenta
        
            Case E_MODO.BorrandoPJ
                Call WriteBorrandoPJ
          
            Case E_MODO.RecuperandoConstraseña
                Call WriteRecuperandoConstraseña
                    
            Case E_MODO.BorrandoCuenta
                Call WriteBorrandoCuenta

        End Select

    End Sub

Private Sub Socket1_Disconnect()

    Dim i As Long
    
    Second.Enabled = False
    Connected = False
    
    Socket1.Cleanup
    
    ' If Not frmCrearPersonaje.Visible And Not frmConnect.Visible Then
    Rem  FrmCuenta.Visible = True
    ' End If

    If LogeoAlgunaVez Then
        frmConnect.MousePointer = vbNormal

        Dim mForm As Form

        For Each mForm In Forms

            Select Case mForm.name

                Case Me.name, frmConnect.name, frmCrearPersonaje.name, frmMensaje.name
            
                Case Else
                    Unload mForm

            End Select

        Next
    
        frmmain.Visible = False

        frmmain.personaje(1).Visible = False
        frmmain.personaje(2).Visible = False
        frmmain.personaje(3).Visible = False
        frmmain.personaje(4).Visible = False
        frmmain.personaje(5).Visible = False
        
        meteo_estado = 0
        UserClase = 0
        UserSexo = 0
        UserRaza = 0
        MiCabeza = 0
        SkillPoints = 0
        UserEstado = 0
        Alocados = 0
    
        For i = 1 To NUMSKILLS
            UserSkills(i) = 0
        Next i

        For i = 1 To NUMATRIBUTOS
            UserAtributos(i) = 0
        Next i
        
        UserParalizado = False
        UserSaliendo = False
        UserInmovilizado = False
        pausa = False
        UserMeditar = False
        UserDescansar = False
        UserNavegando = False
        UserNadando = False
        UserMontado = False
        bRain = False
        AlphaNiebla = 75
        frmmain.TimerNiebla.Enabled = False
        bNiebla = False
        MostrarTrofeo = False
        bNieve = False
        bFogata = False
    
        '  For i = 1 To LastChar + 1
        '      charlist(i).invisible = False
        '      charlist(i).Arma_Aura = 0
        '      charlist(i).Body_Aura = 0
        '      charlist(i).Escudo_Aura = 0
        '      charlist(i).Otra_Aura = 0
        '      charlist(i).Head_Aura = 0
        '      charlist(i).Speeding = 0
        '  Next i

        For i = 1 To LastChar + 1
            charlist(i).dialog = ""
        Next i

        Call RefreshAllChars
    
        macrotrabajo.Enabled = False
        frmmain.Timerping.Enabled = False

        frmConnect.Visible = True
        UserMap = 1
        AlphaNiebla = 25
        EntradaY = 1
        EntradaX = 1
    
        Call SwitchMapIAO(UserMap)
    
        Call engine.Engine_Select_Particle_Set(203)
        ParticleLluviaDorada = General_Particle_Create(208, -1, -1)
    
        frmConnect.txtNombre.Visible = False
        QueRender = 2

        LogeoAlgunaVez = True
    
        'Else
        'General_Set_Connect
    End If

End Sub

Private Sub Socket1_LastError(ErrorCode As Integer, ErrorString As String, Response As Integer)

    '*********************************************
    'Handle socket errors
    '*********************************************
    If ErrorCode = 24036 Then
        frmmain.Socket1.Disconnect
        Debug.Print "ErrorCode = 24036"
        Exit Sub

    End If

    ' Call ComprobarEstado
    
    If frmConnect.Visible Then
        Call TextoAlAsistente("¡No me pude conectar! Te recomiendo verificar el estado de los servidores en www.argentum20.com y asegurarse de estar conectado a internet.")
    Else
        Call MsgBox("Ha ocurrido un error al conectar con el servidor. Le recomendamos verificar el estado de los servidores en www.argentum20.com, y asegurarse de estar conectado directamente a internet", vbApplicationModal + vbInformation + vbOKOnly + vbDefaultButton1, "Error al conectar")
    
        Dim mForm As Form

        For Each mForm In Forms

            Select Case mForm.name

                Case Me.name, frmConnect.name, frmCrearPersonaje.name, frmMensaje.name
                
                Case Else
                    Unload mForm

            End Select

        Next
        
        frmmain.Visible = False
        Call ComprobarEstado
        General_Set_Connect

    End If
    
    frmConnect.MousePointer = 1
    Response = 0
    Second.Enabled = False

    frmmain.Socket1.Disconnect
    LogeoAlgunaVez = False
    
    'General_Set_Connect
    
    'If Not frmCrearPersonaje.Visible Then
    ' General_Set_Connect
    '  Else
    '  frmCrearPersonaje.MousePointer = 0
    'End If
End Sub

Private Sub Socket1_Read(dataLength As Integer, IsUrgent As Integer)

    Dim RD     As String

    Dim Data() As Byte
    
    Call Socket1.Read(RD, dataLength)
    Data = StrConv(RD, vbFromUnicode)
    
    If RD = vbNullString Then Exit Sub
    
    #If SeguridadAlkon Then
        Call DataReceived(Data)
    #End If

    'Put data in the buffer
    Call incomingData.WriteBlock(Data)
    
    'Send buffer to Handle data
    Call HandleIncomingData

End Sub


#End If

'
' -------------------
'    W I N S O C K
' -------------------
'

#If UsarWrench <> 1 Then

Private Sub Winsock1_Close()
    
    Debug.Print "WInsock Close"
    
    Second.Enabled = False
    Connected = False
    
    If Winsock1.State <> sckClosed Then _
        Winsock1.Close
    
    frmConnect.MousePointer = vbNormal
    
    If Not frmCrearPersonaje.Visible And Not FrmCuenta.Visible Then
        General_Set_Connect
    End If

    Dim mForm As Form
    For Each mForm In Forms
        Select Case mForm.name
            Case Me.name, frmConnect.name, frmCrearPersonaje.name, frmMensaje.name
            
            Case Else
                Unload mForm
        End Select
    Next
    
    frmmain.Visible = False
    

frmmain.personaje(1).Visible = False
frmmain.personaje(2).Visible = False
frmmain.personaje(3).Visible = False
frmmain.personaje(4).Visible = False
frmmain.personaje(5).Visible = False
    
 

    pausa = False
    UserMeditar = False

    UserClase = 0
    UserSexo = 0
    UserRaza = 0
    MiCabeza = 0
    
    For i = 1 To NUMSKILLS
        UserSkills(i) = 0
    Next i

    For i = 1 To NUMATRIBUTOS
        UserAtributos(i) = 0
    Next i

    SkillPoints = 0
    Alocados = 0

    Dialogos.CantidadDialogos = 0
End Sub

Private Sub Winsock1_Connect()
    Debug.Print "Winsock Connect"
    
    'Clean input and output buffers
    Call incomingData.ReadASCIIStringFixed(incomingData.length)
    Call outgoingData.ReadASCIIStringFixed(outgoingData.length)
    

    
    Second.Enabled = True
    
    Select Case EstadoLogin
        Case E_MODO.CrearNuevoPj

            Call Login


        Case E_MODO.Normal

            Call Login

        Case E_MODO.Dados

            frmCrearPersonaje.Show vbModal
            
    End Select
End Sub

Private Sub Winsock1_DataArrival(ByVal BytesTotal As Long)
    Dim RD As String
    Dim Data() As Byte
    
    'Socket1.Read RD, DataLength
    Winsock1.GetData RD
    
    Data = StrConv(RD, vbFromUnicode)
    
#If SeguridadAlkon Then
    Call DataReceived(Data)
#End If
    
    'Set data in the buffer
    Call incomingData.WriteBlock(Data)
    
    'Send buffer to Handle data
    Call HandleIncomingData
End Sub

Private Sub Winsock1_Error(ByVal number As Integer, Description As String, ByVal Scode As Long, ByVal source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    '*********************************************
    'Handle socket errors
    '*********************************************
    
    Call MsgBox(Description, vbApplicationModal + vbInformation + vbOKOnly + vbDefaultButton1, "Error")
    frmConnect.MousePointer = 1
    Second.Enabled = False

    If Winsock1.State <> sckClosed Then _
        Winsock1.Close
    


    If Not frmCrearPersonaje.Visible Then
        Rem General_Set_Connect
    Else
        frmCrearPersonaje.MousePointer = 0
    End If
End Sub
#End If

Private Function InGameArea() As Boolean

    If clicX < renderer.Left Or clicX > renderer.Left + (32 * 23) Then Exit Function
    If clicY < renderer.Top Or clicY > renderer.Top + (32 * 17) Then Exit Function
    InGameArea = True

End Function

Private Sub moverForm()

    Dim res As Long

    ReleaseCapture
    res = SendMessage(Me.hwnd, WM_SYSCOMMAND, MOUSE_MOVE, 0)

End Sub

Private Sub imgSpellInfo_Click()

    If hlst.ListIndex <> -1 Then
        Call WriteSpellInfo(hlst.ListIndex + 1)

    End If

End Sub
