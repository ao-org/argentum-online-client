VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmMain 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   13920
   ClientLeft      =   345
   ClientTop       =   240
   ClientWidth     =   19980
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
   ScaleHeight     =   928
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1332
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.PictureBox shapexy 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H000080FF&
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   16920
      ScaleHeight     =   150
      ScaleWidth      =   150
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   7080
      Width           =   180
   End
   Begin VB.Timer dobleclick 
      Left            =   8520
      Top             =   2400
   End
   Begin VB.CommandButton Command1 
      Caption         =   "PRENDER ANTORCHA"
      Height          =   495
      Left            =   11160
      TabIndex        =   43
      Top             =   1680
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox SendTxtCmsg 
      Appearance      =   0  'Flat
      BackColor       =   &H00000040&
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
      TabIndex        =   42
      TabStop         =   0   'False
      ToolTipText     =   "Chat"
      Top             =   1800
      Visible         =   0   'False
      Width           =   8184
   End
   Begin VB.Timer Second 
      Interval        =   1000
      Left            =   7440
      Top             =   2400
   End
   Begin VB.Timer Evento 
      Enabled         =   0   'False
      Interval        =   10000
      Left            =   6360
      Top             =   2400
   End
   Begin VB.Timer UpdateDaytime 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   3480
      Top             =   2400
   End
   Begin VB.Timer Efecto 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   2040
      Top             =   2400
   End
   Begin VB.Timer MacroLadder 
      Enabled         =   0   'False
      Interval        =   1300
      Left            =   1560
      Top             =   2400
   End
   Begin VB.Timer TimerNiebla 
      Interval        =   100
      Left            =   1080
      Top             =   2400
   End
   Begin VB.Timer TimerLluvia 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   600
      Top             =   2400
   End
   Begin VB.Timer UpdateLight 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   3000
      Top             =   2400
   End
   Begin VB.Timer Contadores 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   3960
      Top             =   2400
   End
   Begin VB.Timer cerrarcuenta 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   5400
      Top             =   2400
   End
   Begin VB.Timer LlamaDeclan 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   4920
      Top             =   2400
   End
   Begin VB.PictureBox panel 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4875
      Left            =   11520
      ScaleHeight     =   325
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   247
      TabIndex        =   18
      Top             =   2400
      Width           =   3705
      Begin VB.PictureBox picHechiz 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         CausesValidation=   0   'False
         ClipControls    =   0   'False
         DrawStyle       =   3  'Dash-Dot
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
         Height          =   3405
         Left            =   225
         MousePointer    =   99  'Custom
         ScaleHeight     =   227
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   216
         TabIndex        =   44
         Top             =   600
         Visible         =   0   'False
         Width           =   3240
      End
      Begin VB.PictureBox picInv 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ClipControls    =   0   'False
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
         TabIndex        =   19
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
         Left            =   15
         Tag             =   "0"
         Top             =   15
         Width           =   1830
      End
      Begin VB.Image imgHechizos 
         Height          =   420
         Left            =   1875
         Tag             =   "0"
         Top             =   15
         Width           =   1830
      End
      Begin VB.Image cmdMoverHechi 
         Height          =   285
         Index           =   0
         Left            =   3370
         MouseIcon       =   "frmMain.frx":57E2
         MousePointer    =   99  'Custom
         Tag             =   "0"
         Top             =   4550
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.Image cmdMoverHechi 
         Height          =   285
         Index           =   1
         Left            =   3375
         MouseIcon       =   "frmMain.frx":5934
         MousePointer    =   99  'Custom
         Picture         =   "frmMain.frx":5A86
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
      TabIndex        =   9
      TabStop         =   0   'False
      ToolTipText     =   "Chat"
      Top             =   1800
      Visible         =   0   'False
      Width           =   8184
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
      TabIndex        =   2
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
         Height          =   75
         Index           =   0
         Left            =   450
         Shape           =   3  'Circle
         Top             =   750
         Width           =   75
      End
   End
   Begin VB.Timer macrotrabajo 
      Enabled         =   0   'False
      Left            =   2520
      Top             =   2400
   End
   Begin VB.Timer ShowFPS 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   5880
      Top             =   2400
   End
   Begin RichTextLib.RichTextBox RecTxt 
      Height          =   1275
      Left            =   240
      TabIndex        =   6
      TabStop         =   0   'False
      ToolTipText     =   "Mensajes del servidor"
      Top             =   480
      Width           =   9180
      _ExtentX        =   16193
      _ExtentY        =   2249
      _Version        =   393217
      BackColor       =   0
      BorderStyle     =   0
      HideSelection   =   0   'False
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      DisableNoScroll =   -1  'True
      Appearance      =   0
      TextRTF         =   $"frmMain.frx":6104
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.PictureBox panelInf 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3675
      Left            =   11340
      ScaleHeight     =   245
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   266
      TabIndex        =   8
      Top             =   7815
      Width           =   3990
      Begin VB.Image ImgEstadisticas 
         Height          =   495
         Left            =   2025
         Tag             =   "0"
         ToolTipText     =   "Hogar"
         Top             =   2265
         Visible         =   0   'False
         Width           =   540
      End
      Begin VB.Image Retar 
         Height          =   495
         Left            =   135
         Tag             =   "0"
         ToolTipText     =   "Retos"
         Top             =   2265
         Visible         =   0   'False
         Width           =   540
      End
      Begin VB.Image ImgHogar 
         Height          =   495
         Left            =   2025
         Tag             =   "0"
         ToolTipText     =   "Hogar"
         Top             =   1695
         Visible         =   0   'False
         Width           =   540
      End
      Begin VB.Image ImgSegResu 
         Appearance      =   0  'Flat
         Height          =   408
         Left            =   1680
         ToolTipText     =   "Seguro de resurrección"
         Top             =   3060
         Visible         =   0   'False
         Width           =   408
      End
      Begin VB.Image imgBugReport 
         Height          =   495
         Left            =   2880
         Top             =   3120
         Visible         =   0   'False
         Width           =   1020
      End
      Begin VB.Image mapMundo 
         Height          =   495
         Left            =   135
         Tag             =   "0"
         ToolTipText     =   "Mapa del mundo"
         Top             =   555
         Visible         =   0   'False
         Width           =   540
      End
      Begin VB.Image clanimg 
         Height          =   495
         Left            =   135
         Tag             =   "0"
         ToolTipText     =   "Clanes"
         Top             =   1695
         Visible         =   0   'False
         Width           =   540
      End
      Begin VB.Image Image5 
         Height          =   495
         Left            =   135
         Tag             =   "0"
         ToolTipText     =   "Grupo"
         Top             =   1125
         Visible         =   0   'False
         Width           =   540
      End
      Begin VB.Image cmdLlavero 
         Height          =   495
         Left            =   2025
         Tag             =   "0"
         ToolTipText     =   "Llavero"
         Top             =   555
         Visible         =   0   'False
         Width           =   540
      End
      Begin VB.Image QuestBoton 
         Height          =   495
         Left            =   2040
         Tag             =   "0"
         ToolTipText     =   "Quest"
         Top             =   1125
         Visible         =   0   'False
         Width           =   540
      End
      Begin VB.Image panelinferior 
         Height          =   420
         Index           =   1
         Left            =   1980
         Tag             =   "0"
         Top             =   0
         Width           =   1995
      End
      Begin VB.Image panelinferior 
         Height          =   420
         Index           =   0
         Left            =   0
         Tag             =   "0"
         Top             =   0
         Width           =   1995
      End
      Begin VB.Label stabar 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "999/999"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Left            =   795
         TabIndex        =   35
         Top             =   2055
         Width           =   765
      End
      Begin VB.Label hambar 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "100"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Left            =   3390
         TabIndex        =   34
         Top             =   2055
         Width           =   255
      End
      Begin VB.Label AGUbar 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "100"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Left            =   2445
         TabIndex        =   33
         Top             =   2055
         Width           =   255
      End
      Begin VB.Label HpBar 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "999/999"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   1635
         TabIndex        =   32
         Top             =   1215
         Width           =   945
      End
      Begin VB.Label lbldm 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "+0%"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   240
         Left            =   3075
         TabIndex        =   31
         ToolTipText     =   "Aumento de daño mágico"
         Top             =   3150
         Width           =   570
      End
      Begin VB.Label lblResis 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "+0"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   240
         Left            =   1935
         TabIndex        =   30
         ToolTipText     =   "Resistencia mágica"
         Top             =   3150
         Width           =   330
      End
      Begin VB.Label oxigenolbl 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "15"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   270
         Left            =   2385
         TabIndex        =   29
         ToolTipText     =   "Oxígeno acumulado"
         Top             =   570
         Width           =   225
      End
      Begin VB.Label GldLbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "100.000"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FFFF&
         Height          =   270
         Left            =   480
         TabIndex        =   28
         ToolTipText     =   "Monedas de oro"
         Top             =   585
         Width           =   690
      End
      Begin VB.Label Fuerzalbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "40"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   270
         Left            =   3585
         TabIndex        =   27
         ToolTipText     =   "Fuerza"
         Top             =   570
         Width           =   210
      End
      Begin VB.Label AgilidadLbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "40"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   270
         Left            =   2985
         TabIndex        =   26
         ToolTipText     =   "Agilidad"
         Top             =   570
         Width           =   210
      End
      Begin VB.Label lblWeapon 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0/0"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   240
         Left            =   600
         TabIndex        =   25
         ToolTipText     =   "Daño físico arma"
         Top             =   2700
         Width           =   480
      End
      Begin VB.Label lblShielder 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0/0"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   240
         Left            =   1860
         TabIndex        =   24
         ToolTipText     =   "Defensa escudo"
         Top             =   2700
         Width           =   480
      End
      Begin VB.Label lblHelm 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0/0"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   240
         Left            =   3120
         TabIndex        =   23
         ToolTipText     =   "Defensa casco"
         Top             =   2700
         Width           =   480
      End
      Begin VB.Label lblArmor 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0/0"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   240
         Left            =   600
         TabIndex        =   22
         ToolTipText     =   "Defensa armadura"
         Top             =   3150
         Width           =   480
      End
      Begin VB.Label manabar 
         Alignment       =   2  'Center
         BackColor       =   &H80000006&
         BackStyle       =   0  'Transparent
         Caption         =   "999/999"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   1635
         TabIndex        =   21
         Top             =   1635
         Width           =   945
      End
      Begin VB.Image imgOro 
         Height          =   255
         Left            =   120
         Top             =   600
         Width           =   375
      End
      Begin VB.Image Hpshp 
         Height          =   192
         Left            =   528
         Top             =   1212
         Width           =   2592
      End
      Begin VB.Image shieldBar 
         Height          =   192
         Left            =   528
         Top             =   1212
         Width           =   2592
      End
      Begin VB.Image MANShp 
         Height          =   192
         Left            =   528
         Top             =   1632
         Width           =   2592
      End
      Begin VB.Image STAShp 
         Height          =   108
         Left            =   516
         Top             =   2088
         Width           =   1068
      End
      Begin VB.Image AGUAsp 
         Height          =   108
         Left            =   2340
         Top             =   2088
         Width           =   384
      End
      Begin VB.Image COMIDAsp 
         Height          =   96
         Left            =   3288
         Top             =   2100
         Width           =   384
      End
      Begin VB.Image ImgSegClan 
         Appearance      =   0  'Flat
         Height          =   408
         Left            =   636
         ToolTipText     =   "Seguro de clan"
         Top             =   3060
         Visible         =   0   'False
         Width           =   408
      End
      Begin VB.Image ImgSegParty 
         Height          =   408
         Left            =   108
         ToolTipText     =   "Seguro de grupo"
         Top             =   3060
         Visible         =   0   'False
         Width           =   408
      End
      Begin VB.Image ImgSeg 
         Appearance      =   0  'Flat
         Height          =   408
         Left            =   1152
         ToolTipText     =   "Seguro de ataque"
         Top             =   3060
         Visible         =   0   'False
         Width           =   408
      End
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
      Height          =   9120
      Left            =   120
      ScaleHeight     =   608
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   736
      TabIndex        =   4
      Top             =   2280
      Width           =   11040
   End
   Begin VB.Image imgDeleteItem 
      Height          =   375
      Left            =   11400
      Top             =   7320
      Width           =   615
   End
   Begin VB.Label btnInvisible 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Invisible"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000018&
      Height          =   375
      Left            =   6840
      TabIndex        =   41
      Top             =   75
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label btnSpawn 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Spawn NPC"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000018&
      Height          =   375
      Left            =   5400
      TabIndex        =   40
      Top             =   75
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label createObj 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Crear Obj"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000018&
      Height          =   375
      Left            =   3960
      TabIndex        =   39
      Top             =   75
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label panelGM 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "PanelGM"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000018&
      Height          =   375
      Left            =   2640
      TabIndex        =   38
      Top             =   75
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Label3"
      Height          =   495
      Left            =   9360
      TabIndex        =   37
      Top             =   6720
      Width           =   1215
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   11880
      TabIndex        =   36
      Top             =   0
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.Label lblhora 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "00:00"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   13170
      TabIndex        =   20
      Top             =   2010
      Width           =   495
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
      TabIndex        =   17
      Top             =   7170
      Visible         =   0   'False
      Width           =   3900
   End
   Begin VB.Label ms 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "30 ms"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   210
      Left            =   8580
      TabIndex        =   16
      Top             =   210
      Width           =   465
   End
   Begin VB.Label fps 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fps: 200"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   210
      Left            =   8490
      TabIndex        =   15
      ToolTipText     =   "Numero de usuarios online"
      Top             =   30
      Width           =   645
   End
   Begin VB.Image EstadisticasBoton 
      Height          =   420
      Left            =   14738
      Tag             =   "0"
      Top             =   990
      Width           =   465
   End
   Begin VB.Label Coord 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "000 X:00 Y: 00"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   225
      Left            =   9720
      TabIndex        =   14
      Top             =   210
      Width           =   1215
   End
   Begin VB.Label NombrePJ 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre del pj"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   11400
      TabIndex        =   13
      Top             =   600
      Width           =   3945
   End
   Begin VB.Label lblPorcLvl 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "33.33%"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   11580
      TabIndex        =   11
      Top             =   1545
      Visible         =   0   'False
      Width           =   3540
   End
   Begin VB.Image OpcionesBoton 
      Height          =   315
      Left            =   11431
      Tag             =   "0"
      Top             =   65
      Width           =   315
   End
   Begin VB.Image CombateIcon 
      Height          =   144
      Left            =   8832
      Top             =   1812
      Width           =   444
   End
   Begin VB.Image globalIcon 
      Height          =   144
      Left            =   8832
      Top             =   2004
      Width           =   444
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   10320
      TabIndex        =   7
      ToolTipText     =   "Activar / desactivar chat globales"
      Top             =   1800
      Width           =   855
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
      Caption         =   "Online: 1"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   450
      Left            =   12480
      TabIndex        =   5
      ToolTipText     =   "Numero de usuarios online"
      Top             =   0
      Width           =   1665
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
      TabIndex        =   3
      Top             =   1680
      Width           =   450
   End
   Begin VB.Image PicCorreo 
      Height          =   348
      Left            =   11520
      Top             =   480
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.Label NameMapa 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Mapa Desconocido"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   165
      Left            =   9765
      TabIndex        =   1
      Top             =   45
      Width           =   1125
   End
   Begin VB.Label lblLvl 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Guerrero - Nivel 50"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Left            =   12480
      TabIndex        =   12
      Top             =   1080
      Width           =   1785
   End
   Begin VB.Label exp 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "99999/99999"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   11580
      TabIndex        =   10
      Top             =   1545
      Width           =   3540
   End
   Begin VB.Image ExpBar 
      Height          =   192
      Left            =   11580
      Top             =   1548
      Width           =   2832
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
Attribute VB_Name = "frmMain"
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
'You can contact me at:
'morgolock@speedy.com.ar
'
'Calle 3 número 983 piso 7 dto A
'La Plata - Pcia, Buenos Aires - Republica Argentina
'Código Postal 1900
'Pablo Ignacio Márquez
'Call ParseUserCommand("/CMSG " & stxtbuffercmsg)
Option Explicit

Private Declare Sub svb_shutdown_steam Lib "steam_vb.dll" ()


Public WithEvents Inventario As clsGrapchicalInventory
Attribute Inventario.VB_VarHelpID = -1

Private Const WS_EX_TRANSPARENT = &H20&
Private totalclicks As Integer
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


Private Declare Function ReleaseCapture Lib "user32" () As Long

Public MouseBoton As Long

Public MouseShift As Long

Public IsPlaying  As Byte

Public ShowPercentage As Boolean

Public bmoving    As Boolean

Public dX         As Integer

Public dy         As Integer

Private Const IntervaloEntreClicks As Long = 50

Dim TempTick As Long

Private iClickTick As Long

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
Private cBotonEliminarItem As clsGraphicalButton

Private Sub btnInvisible_Click()
    
    On Error GoTo btnInvisible_Click_Err
    
    Call ParseUserCommand("/INVISIBLE")
    
    Me.SetFocus
    
    Exit Sub

btnInvisible_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmMain.btnInvisible_Click", Erl)
    Resume Next
    
End Sub


Private Sub loadButtons()

    Set cBotonEliminarItem = New clsGraphicalButton
                                                
    Call cBotonEliminarItem.Initialize(imgDeleteItem, "boton-borrar-item-default.bmp", _
                                                "boton-borrar-item-over.bmp", _
                                                "boton-borrar-item-off.bmp", Me)
End Sub

Private Sub btnSpawn_Click()
    
    On Error GoTo btnSpawn_Click_Err
    
    Me.SetFocus
    
    Call WriteSpawnListRequest
    
    Exit Sub

btnSpawn_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmMain.btnSpawn_Click", Erl)
    Resume Next
    
End Sub

Private Sub clanimg_MouseMove(button As Integer, Shift As Integer, x As Single, y As Single)
    
    On Error GoTo clanimg_MouseMove_Err
    

    If clanimg.Tag = "0" Then
        clanimg.Picture = LoadInterface("claniluminado.bmp")
        clanimg.Tag = "1"

    End If

    
    Exit Sub

clanimg_MouseMove_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmMain.clanimg_MouseMove", Erl)
    Resume Next
    
End Sub

Private Sub clanimg_MouseUp(button As Integer, Shift As Integer, x As Single, y As Single)
    
    On Error GoTo clanimg_MouseUp_Err
    
    If pausa Then Exit Sub

    If frmGuildLeader.visible Then Unload frmGuildLeader
    Call WriteRequestGuildLeaderInfo

    
    Exit Sub

clanimg_MouseUp_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmMain.clanimg_MouseUp", Erl)
    Resume Next
    
End Sub

Private Sub cmdlanzar_MouseDown(button As Integer, Shift As Integer, x As Single, y As Single)
    
    On Error GoTo cmdlanzar_MouseDown_Err
    
    
    If ModoHechizos = BloqueoLanzar Then
        If Not MainTimer.Check(TimersIndex.AttackSpell, False) Or Not MainTimer.Check(TimersIndex.CastSpell, False) Then
            Exit Sub
        End If
    End If
    
    cmdlanzar.Picture = LoadInterface("boton-lanzar-off.bmp")
    cmdlanzar.Tag = "1"

    
    Exit Sub

cmdlanzar_MouseDown_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmMain.cmdlanzar_MouseDown", Erl)
    Resume Next
    
End Sub

Private Sub cmdlanzar_MouseUp(button As Integer, Shift As Integer, x As Single, y As Single)
    
    On Error GoTo cmdlanzar_MouseUp_Err
    
    If ModoHechizos = BloqueoLanzar Then
        If Not MainTimer.Check(TimersIndex.AttackSpell, False) Or Not MainTimer.Check(TimersIndex.CastSpell, False) Then
            Exit Sub
        End If
    End If
    
    cmdlanzar.Picture = LoadInterface("boton-lanzar-over.bmp")
    cmdlanzar.Tag = "1"
    
    Exit Sub

cmdlanzar_MouseUp_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmMain.cmdlanzar_MouseUp", Erl)
    Resume Next
    
End Sub

Private Sub cmdLlavero_Click()
    
    On Error GoTo cmdLlavero_Click_Err
    
    If FrmKeyInv.visible Then
        CerrarLlavero
    Else
        FrmKeyInv.Show , Me
        FrmKeyInv.Left = Me.Left + 890 * screen.TwipsPerPixelX - FrmKeyInv.Width \ 2
        FrmKeyInv.Top = Me.Top + 555 * screen.TwipsPerPixelY - FrmKeyInv.Height
        cmdLlavero.Tag = "2"
    End If
    
    Exit Sub

cmdLlavero_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmMain.cmdLlavero_Click", Erl)
    Resume Next
    
End Sub

Public Sub CerrarLlavero()
    
    On Error GoTo CerrarLlavero_Err
    
    FrmKeyInv.visible = False
    cmdLlavero.Picture = LoadInterface("boton-llavero-over.bmp")
    cmdLlavero.Tag = "1"
    
    Exit Sub

CerrarLlavero_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmMain.CerrarLlavero", Erl)
    Resume Next
    
End Sub

Private Sub cmdLlavero_MouseDown(button As Integer, Shift As Integer, x As Single, y As Single)
    
    On Error GoTo cmdLlavero_MouseDown_Err
    
    cmdLlavero.Picture = LoadInterface("boton-llavero-off.bmp")
    cmdLlavero.Tag = "1"
    
    Exit Sub

cmdLlavero_MouseDown_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmMain.cmdLlavero_MouseDown", Erl)
    Resume Next
    
End Sub

Private Sub cmdLlavero_MouseMove(button As Integer, Shift As Integer, x As Single, y As Single)
    
    On Error GoTo cmdLlavero_MouseMove_Err
    
    If cmdLlavero.Tag = "0" Then
        cmdLlavero.Picture = LoadInterface("boton-llavero-over.bmp")
        cmdLlavero.Tag = "1"
    End If
    
    Exit Sub

cmdLlavero_MouseMove_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmMain.cmdLlavero_MouseMove", Erl)
    Resume Next
    
End Sub

Private Sub cmdMoverHechi_Click(Index As Integer)
    
    On Error GoTo cmdMoverHechi_Click_Err
    

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

    
    Exit Sub

cmdMoverHechi_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmMain.cmdMoverHechi_Click", Erl)
    Resume Next
    
End Sub

Public Sub ControlSeguroParty(ByVal Mostrar As Boolean)
    
    On Error GoTo ControlSeguroParty_Err

    If Mostrar Then
        ImgSegParty = LoadInterface("boton-seguro-party-on.bmp")
        SeguroParty = True
    Else
        ImgSegParty = LoadInterface("boton-seguro-party-off.bmp")
        SeguroParty = False
    End If
    
    Exit Sub

ControlSeguroParty_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmMain.ControlSeguroParty", Erl)
    Resume Next
    
End Sub

Public Sub DibujarSeguro()
    
    On Error GoTo DibujarSeguro_Err
    
    ImgSeg = LoadInterface("boton-seguro-ciudadano-on.bmp")
    
    SeguroGame = True

    
    Exit Sub

DibujarSeguro_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmMain.DibujarSeguro", Erl)
    Resume Next
    
End Sub

Public Sub DesDibujarSeguro()
    
    On Error GoTo DesDibujarSeguro_Err
    
    ImgSeg = LoadInterface("boton-seguro-ciudadano-off.bmp")
    
    SeguroGame = False

    
    Exit Sub

DesDibujarSeguro_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmMain.DesDibujarSeguro", Erl)
    Resume Next
    
End Sub

Private Sub cmdMoverHechi_MouseMove(Index As Integer, button As Integer, Shift As Integer, x As Single, y As Single)
    
    On Error GoTo cmdMoverHechi_MouseMove_Err
    

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
    
    
    Exit Sub

cmdMoverHechi_MouseMove_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmMain.cmdMoverHechi_MouseMove", Erl)
    Resume Next
    
End Sub

Private Sub CombateIcon_Click()
    
    On Error GoTo CombateIcon_Click_Err
    

    If ChatCombate = 0 Then
        ChatCombate = 1
        CombateIcon.Picture = LoadInterface("infoapretado.bmp")
    Else
        ChatCombate = 0
        CombateIcon.Picture = LoadInterface("info.bmp")

    End If

    Call WriteMacroPos

    
    Exit Sub

CombateIcon_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmMain.CombateIcon_Click", Erl)
    Resume Next
    
End Sub


Private Sub Command1_Click()
    TieneAntorcha = Not TieneAntorcha
End Sub


Private Sub Contadores_Timer()
    
    On Error GoTo Contadores_Timer_Err
    

    If UserStats.estado = 1 Then Exit Sub
    If InviCounter > 0 Then
        InviCounter = InviCounter - 1

    End If


    If DrogaCounter > 0 Then
        DrogaCounter = DrogaCounter - 1

        If DrogaCounter <= 12 And DrogaCounter > 0 Then
            Call Sound.Sound_Stop(SND_DOPA)
            Call Sound.Sound_Play(SND_DOPA)
            If UserStats.StrState <> eBlink Then
                UserStats.StrState = eBlink
                UserStats.AgiState = eBlink
                If BabelInitialized Then
                    Call UpdateBuffState
                End If
            End If
            If DrogaCounter Mod 2 = 0 Then
                frmMain.Fuerzalbl.ForeColor = vbWhite
                frmMain.AgilidadLbl.ForeColor = vbWhite
            Else
                frmMain.Fuerzalbl.ForeColor = RGB(204, 0, 0)
                frmMain.AgilidadLbl.ForeColor = RGB(204, 0, 0)
            End If
        End If
    End If

    If InviCounter = 0 And DrogaCounter = 0 Then
        Contadores.enabled = False
        If UserStats.AgiState <> eNormal Then
            UserStats.AgiState = eNormal
            UserStats.StrState = eNormal
            If BabelInitialized Then
                Call UpdateBuffState
            End If
        End If
    End If

    
    Exit Sub

Contadores_Timer_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmMain.Contadores_Timer", Erl)
    Resume Next
    
End Sub

Private Sub createObj_Click()
    Call OpenCreateObjectMenu
End Sub

Private Sub dobleclick_Timer()
    Static segundo As Long
    segundo = segundo + 1
    If segundo = 2 And totalclicks > 20 Then
    Call WriteLogMacroClickHechizo(tMacro.dobleclick, totalclicks)
    totalclicks = 0
    segundo = 0
    dobleclick.Interval = 0
    'Label10.Caption = 0
    End If
    If segundo = 2 And totalclicks <= 20 Then
    totalclicks = 0
    segundo = 0
    dobleclick.Interval = 0
    End If
End Sub

Private Sub Efecto_Timer()
    
    On Error GoTo Efecto_Timer_Err
    
    If MapDat.base_light > 0 Then
        Call SetGlobalLight(MapDat.base_light)
    Else
        Call RestaurarLuz
    End If

    Efecto.enabled = False
    EfectoEnproceso = False

    
    Exit Sub

Efecto_Timer_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmMain.Efecto_Timer", Erl)
    Resume Next
    
End Sub

Private Sub hlst_Click()
    

End Sub

Private Sub Image1_Click()

End Sub

Private Sub ImgEstadisticas_MouseDown(button As Integer, Shift As Integer, x As Single, y As Single)
    TempTick = GetTickCount And &H7FFFFFFF
    If TempTick - iClickTick < IntervaloEntreClicks And Not iClickTick = 0 Then Exit Sub
    iClickTick = TempTick
    ImgEstadisticas.Picture = LoadInterface("boton-estadisticas-big-off.bmp")
    ImgEstadisticas.Tag = "1"
End Sub
Private Sub ImgEstadisticas_MouseUp(button As Integer, Shift As Integer, x As Single, y As Single)
    
    If pausa Then Exit Sub
    
    LlegaronAtrib = False
    LlegaronStats = False
    Call WriteRequestAtributes
    Call WriteRequestMiniStats
    
End Sub

Private Sub ImgEstadisticas_MouseMove(button As Integer, Shift As Integer, x As Single, y As Single)
    If ImgEstadisticas.Tag = "0" Then
        ImgEstadisticas.Picture = LoadInterface("boton-estadisticas-big-over.bmp")
        ImgEstadisticas.Tag = "1"
    End If
End Sub
Private Sub EstadisticasBoton_MouseDown(button As Integer, Shift As Integer, x As Single, y As Single)
    
    On Error GoTo EstadisticasBoton_MouseDown_Err
    
    
    
    EstadisticasBoton.Picture = LoadInterface("boton-skills-off.bmp")
    EstadisticasBoton.Tag = "1"
    
    Exit Sub

EstadisticasBoton_MouseDown_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmMain.EstadisticasBoton_MouseDown", Erl)
    Resume Next
    
End Sub

Private Sub EstadisticasBoton_MouseMove(button As Integer, Shift As Integer, x As Single, y As Single)
    
    On Error GoTo EstadisticasBoton_MouseMove_Err
    

    If EstadisticasBoton.Tag = "0" Then
        EstadisticasBoton.Picture = LoadInterface("boton-skills-over.bmp")
        EstadisticasBoton.Tag = "1"
    End If

    
    Exit Sub

EstadisticasBoton_MouseMove_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmMain.EstadisticasBoton_MouseMove", Erl)
    Resume Next
    
End Sub

Private Sub EstadisticasBoton_MouseUp(button As Integer, Shift As Integer, x As Single, y As Single)
On Error GoTo EstadisticasBoton_MouseUp_Err
    Call ModGameplayUI.RequestSkills
    Exit Sub

EstadisticasBoton_MouseUp_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmMain.EstadisticasBoton_MouseUp", Erl)
    Resume Next
    
End Sub

Private Sub Evento_Timer()

    InvasionActual = 0
    
    Evento.enabled = False

End Sub

Private Sub exp_Click()
    
    On Error GoTo exp_Click_Err

    'Call WriteScrollInfo

    ShowPercentage = Not ShowPercentage
    
    Exit Sub

exp_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmMain.exp_Click", Erl)
    Resume Next
    
End Sub

Private Sub exp_MouseMove(button As Integer, Shift As Integer, x As Single, y As Single)

    If Not ShowPercentage Then
        lblPorcLvl.visible = True
        exp.visible = False
    End If
    
End Sub

Private Sub Form_Activate()
    
    On Error GoTo Form_Activate_Err
    
    Exit Sub

Form_Activate_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmMain.Form_Activate", Erl)
    Resume Next
    
End Sub


Private Sub imgHechizos_MouseUp(button As Integer, Shift As Integer, x As Single, y As Single)
    On Error GoTo imgHechizos_MouseMove_Err
    
    imgHechizos.Picture = LoadInterface("boton-hechizos-default.bmp")
    imgHechizos.Tag = "1"

    
    Exit Sub

imgHechizos_MouseMove_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmMain.imgHechizos_MouseMove", Erl)
    Resume Next
End Sub

Private Sub imgInventario_MouseUp(button As Integer, Shift As Integer, x As Single, y As Single)
    On Error GoTo imgInventario_MouseMove_Err

    imgInventario.Picture = LoadInterface("boton-inventory-over.bmp")
    imgInventario.Tag = "1"

    
    Exit Sub

imgInventario_MouseMove_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmMain.imgInventario_MouseMove", Erl)
    Resume Next
    
End Sub

Private Sub picHechiz_MouseDown(button As Integer, Shift As Integer, x As Single, y As Single)
If y < 0 Then y = 0
If y > Int(picHechiz.ScaleHeight / hlst.Pixel_Alto) * hlst.Pixel_Alto - 1 Then y = Int(picHechiz.ScaleHeight / hlst.Pixel_Alto) * hlst.Pixel_Alto - 1
If x < picHechiz.ScaleWidth - 10 Then
    hlst.ListIndex = Int(y / hlst.Pixel_Alto) + hlst.Scroll
    hlst.DownBarrita = 0
    If Seguido = 1 Then
        Call WriteNotifyInventarioHechizos(2, hlst.ListIndex, hlst.Scroll)
    End If
Else
    hlst.DownBarrita = y - hlst.Scroll * (picHechiz.ScaleHeight - hlst.BarraHeight) / (hlst.ListCount - hlst.VisibleCount)
End If
End Sub

Private Sub picHechiz_MouseMove(button As Integer, Shift As Integer, x As Single, y As Single)
If button = 1 Then
    Dim yy As Integer
    yy = y
    If yy < 0 Then yy = 0
    If yy > Int(picHechiz.ScaleHeight / hlst.Pixel_Alto) * hlst.Pixel_Alto - 1 Then yy = Int(picHechiz.ScaleHeight / hlst.Pixel_Alto) * hlst.Pixel_Alto - 1
    If hlst.DownBarrita > 0 Then
        hlst.Scroll = (y - hlst.DownBarrita) * (hlst.ListCount - hlst.VisibleCount) / (picHechiz.ScaleHeight - hlst.BarraHeight)
    Else
        hlst.ListIndex = Int(yy / hlst.Pixel_Alto) + hlst.Scroll
        If Seguido = 1 Then
            Call WriteNotifyInventarioHechizos(2, hlst.ListIndex, hlst.Scroll)
        End If
        If ScrollArrastrar = 0 Then
            If (y < yy) Then hlst.Scroll = hlst.Scroll - 1
            If (y > yy) Then hlst.Scroll = hlst.Scroll + 1
        End If
    End If
ElseIf button = 0 Then
    hlst.ShowBarrita = x > picHechiz.ScaleWidth - hlst.BarraWidth * 2
End If
End Sub

Private Sub picHechiz_MouseUp(button As Integer, Shift As Integer, x As Single, y As Single)
hlst.DownBarrita = 0
End Sub

Private Sub picInv_MouseUp(button As Integer, Shift As Integer, x As Single, y As Single)
    If Not picInv.visible Then Exit Sub
    
    If dobleclick.Interval = 0 Then dobleclick.Interval = 1000
    
    If button = 1 Then
        dobleclick.Interval = 1000
        totalclicks = totalclicks + 1
    End If
End Sub

Private Sub Second_Timer()
    If Not DialogosClanes Is Nothing Then DialogosClanes.PassTimer
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    On Error GoTo Form_KeyUp_Err
        Call HandleKeyUp(KeyCode, Shift)
    Exit Sub
Form_KeyUp_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmMain.Form_KeyUp", Erl)
    Resume Next
End Sub

Private Sub Form_MouseDown(button As Integer, Shift As Integer, x As Single, y As Single)
    
    On Error GoTo Form_MouseDown_Err
    
    
    If SendTxt.visible Then SendTxt.SetFocus
    MouseBoton = button
    MouseShift = Shift
    
    If frmComerciar.visible Then
        Unload frmComerciar

    End If
    
    If frmBancoObj.visible Then
        Unload frmBancoObj

    End If
    
    If FrmQuestInfo.visible Then
        Unload FrmQuestInfo
    End If

    
    Exit Sub

Form_MouseDown_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmMain.Form_MouseDown", Erl)
    Resume Next
    
End Sub

Private Sub Form_MouseUp(button As Integer, Shift As Integer, x As Single, y As Single)
    
    On Error GoTo Form_MouseUp_Err
    
    clicX = x
    clicY = y

    
    Exit Sub

Form_MouseUp_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmMain.Form_MouseUp", Erl)
    Resume Next
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    
    On Error GoTo Form_QueryUnload_Err
    

    If prgRun = True Then
        prgRun = False
        Cancel = 1

    End If

    
    Exit Sub

Form_QueryUnload_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmMain.Form_QueryUnload", Erl)
    Resume Next
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    On Error GoTo Form_Unload_Err
    
    Call svb_shutdown_steam
    Call DisableURLDetect

    
    Exit Sub

Form_Unload_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmMain.Form_Unload", Erl)
    Resume Next
    
End Sub

Private Sub GldLbl_Click()
    
    On Error GoTo GldLbl_Click_Err
    
    Inventario.SelectGold

    If UserStats.GLD > 0 Then
        frmCantidad.Show , frmMain

    End If

    
    Exit Sub

GldLbl_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmMain.GldLbl_Click", Erl)
    Resume Next
    
End Sub

Private Sub GlobalIcon_Click()
    
    On Error GoTo GlobalIcon_Click_Err
    

    If ChatGlobal = 0 Then
        ChatGlobal = 1
        globalIcon.Picture = LoadInterface("globalapretado.bmp")
    Else
        ChatGlobal = 0
        globalIcon.Picture = LoadInterface("global.bmp")

    End If

    Call WriteMacroPos

    
    Exit Sub

GlobalIcon_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmMain.GlobalIcon_Click", Erl)
    Resume Next
    
End Sub

Private Sub Image2_MouseDown(Index As Integer, button As Integer, Shift As Integer, x As Single, y As Single)
    
    On Error GoTo Image2_MouseDown_Err
    

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

    
    Exit Sub

Image2_MouseDown_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmMain.Image2_MouseDown", Erl)
    Resume Next
    
End Sub

Private Sub HpBar_Click()
    Call ParseUserCommand("/PROMEDIO")
End Sub

Private Sub Hpshp_Click()
    HpBar_Click
End Sub

Private Sub Image3_MouseDown(button As Integer, Shift As Integer, x As Single, y As Single)
    '    Image3.Picture = LoadInterface("elegirchatapretado.bmp")
    
    On Error GoTo Image3_MouseDown_Err
    
    frmMensaje.PopupMenuMensaje

    
    Exit Sub

Image3_MouseDown_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmMain.Image3_MouseDown", Erl)
    Resume Next
    
End Sub

Private Sub Image3_MouseMove(button As Integer, Shift As Integer, x As Single, y As Single)
    
    On Error GoTo Image3_MouseMove_Err
    

    If Image3.Tag = "0" Then
        Image3.Picture = LoadInterface("elegirchatmarcado.bmp")
        Image3.Tag = "1"

    End If

    
    Exit Sub

Image3_MouseMove_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmMain.Image3_MouseMove", Erl)
    Resume Next
    
End Sub

Private Sub Image4_Click(Index As Integer)
    
    On Error GoTo Image4_Click_Err
    

    Select Case Index

        Case 0
            Me.WindowState = vbMinimized

        Case 1
            If frmCerrar.visible Then Exit Sub
            Dim mForm As Form
            For Each mForm In Forms
                If mForm.hwnd <> Me.hwnd Then Unload mForm
                Set mForm = Nothing
            Next
            frmCerrar.Show , Me

    End Select

    
    Exit Sub

Image4_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmMain.Image4_Click", Erl)
    Resume Next
    
End Sub

Private Sub Image4_MouseDown(Index As Integer, button As Integer, Shift As Integer, x As Single, y As Single)
    
    On Error GoTo Image4_MouseDown_Err
    

    Select Case Index

        Case 0
            Image4(0).Picture = LoadInterface("boton-sm-minimizar-off.bmp")

        Case 1
            Image4(1).Picture = LoadInterface("boton-sm-cerrar-off.bmp")

    End Select

    
    Exit Sub

Image4_MouseDown_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmMain.Image4_MouseDown", Erl)
    Resume Next
    
End Sub

Private Sub Image4_MouseMove(Index As Integer, button As Integer, Shift As Integer, x As Single, y As Single)
    
    On Error GoTo Image4_MouseMove_Err
    

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
         
    
    Exit Sub

Image4_MouseMove_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmMain.Image4_MouseMove", Erl)
    Resume Next
    
End Sub

Private Sub Image5_Click()
    
    On Error GoTo Image5_Click_Err
    

    If FrmGrupo.visible = False Then
        Call WriteRequestGrupo

    End If

    
    Exit Sub

Image5_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmMain.Image5_Click", Erl)
    Resume Next
    
End Sub

Private Sub Image5_MouseMove(button As Integer, Shift As Integer, x As Single, y As Single)
    
    On Error GoTo Image5_MouseMove_Err
    

    If Image5.Tag = "0" Then
        Image5.Picture = LoadInterface("grupoover.bmp")
        Image5.Tag = "1"

    End If

    
    Exit Sub

Image5_MouseMove_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmMain.Image5_MouseMove", Erl)
    Resume Next
    
End Sub

Private Sub Image6_Click()
    
    On Error GoTo Image6_Click_Err
    
    Call WriteSafeToggle

    
    Exit Sub

Image6_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmMain.Image6_Click", Erl)
    Resume Next
    
End Sub

Private Sub imgBugReport_Click()
    
    On Error GoTo imgBugReport_Click_Err
    
    FrmGmAyuda.Show vbModeless, frmMain

    
    Exit Sub

imgBugReport_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmMain.imgBugReport_Click", Erl)
    Resume Next
    
End Sub

Private Sub imgHechizos_Click()
  Call hechizosClick
End Sub

Public Sub hechizosClick()
  
    On Error GoTo hechizosClick_Err
    

    If picHechiz.visible Then Exit Sub
    
    TempTick = GetTickCount And &H7FFFFFFF
    
    If TempTick - iClickTick < IntervaloEntreClicks And Not iClickTick = 0 And LastMacroButton <> tMacroButton.Hechizos Then
        Call WriteLogMacroClickHechizo(tMacro.Coordenadas)
    End If
    
    iClickTick = TempTick
    
    LastMacroButton = tMacroButton.Hechizos
    
    panel.Picture = LoadInterface("centrohechizo.bmp")
    
    
    If Seguido = 1 Then
        Call WriteNotifyInventarioHechizos(2, hlst.ListIndex, hlst.Scroll)
    End If
    
    picInv.visible = False
    
    picHechiz.visible = True

    cmdlanzar.visible = True

    imgSpellInfo.visible = True

    cmdMoverHechi(0).visible = True
    cmdMoverHechi(1).visible = True

    frmMain.imgInvLock(0).visible = False
    frmMain.imgInvLock(1).visible = False
    frmMain.imgInvLock(2).visible = False
    imgDeleteItem.visible = False
    
    Exit Sub

hechizosClick_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmMain.hechizosClick", Erl)
    Resume Next
End Sub

Private Sub imgHechizos_MouseDown(button As Integer, Shift As Integer, x As Single, y As Single)
    
    On Error GoTo imgHechizos_MouseDown_Err
       
    imgHechizos.Picture = LoadInterface("boton-hechizos-off.bmp")
    imgHechizos.Tag = "2"

    
    Exit Sub

imgHechizos_MouseDown_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmMain.imgHechizos_MouseDown", Erl)
    Resume Next
    
End Sub

Private Sub imgHechizos_MouseMove(button As Integer, Shift As Integer, x As Single, y As Single)
    
    On Error GoTo imgHechizos_MouseMove_Err
    
    
    If imgHechizos.Tag <> "1" And button = 0 Then
        imgHechizos.Picture = LoadInterface("boton-hechizos-default.bmp")
        imgHechizos.Tag = "1"
        imgInventario.Picture = Nothing
        imgInventario.Tag = "0"
    End If

    
    Exit Sub

imgHechizos_MouseMove_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmMain.imgHechizos_MouseMove", Erl)
    Resume Next
    
End Sub

Private Sub ImgHogar_Click()
    Call ParseUserCommand("/HOGAR")
End Sub

Private Sub ImgHogar_MouseDown(button As Integer, Shift As Integer, x As Single, y As Single)
    ImgHogar.Picture = LoadInterface("boton-home-off.bmp")
    ImgHogar.Tag = "1"
End Sub

Private Sub ImgHogar_MouseMove(button As Integer, Shift As Integer, x As Single, y As Single)
    If ImgHogar.Tag = "0" Then
        ImgHogar.Picture = LoadInterface("boton-home-over.bmp")
        ImgHogar.Tag = "1"
    End If
End Sub

Private Sub imgInventario_Click()
    Call inventoryClick
End Sub

Public Sub inventoryClick()
    On Error GoTo inventoryClick_Err
    

    If picInv.visible Then Exit Sub

    TempTick = GetTickCount And &H7FFFFFFF
    
    If TempTick - iClickTick < IntervaloEntreClicks And Not iClickTick = 0 And LastMacroButton <> tMacroButton.Inventario Then
        Call WriteLogMacroClickHechizo(tMacro.Coordenadas)
    End If
    
    iClickTick = TempTick
    
    LastMacroButton = tMacroButton.Inventario

    panel.Picture = LoadInterface("centroinventario.bmp")
    'Call Audio.PlayWave(SND_CLICK)
    picInv.visible = True
    picHechiz.visible = False
    cmdlanzar.visible = False
    imgSpellInfo.visible = False
    If Seguido = 1 Then
        Call WriteNotifyInventarioHechizos(1, hlst.ListIndex, hlst.Scroll)
    End If

    cmdMoverHechi(0).visible = False
    cmdMoverHechi(1).visible = False
   ' Call Inventario.ReDraw
    frmMain.imgInvLock(0).visible = True
    frmMain.imgInvLock(1).visible = True
    frmMain.imgInvLock(2).visible = True
    imgDeleteItem.visible = True

    
    Exit Sub

inventoryClick_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmMain.inventoryClick", Erl)
    Resume Next
End Sub

Private Sub imgInventario_MouseDown(button As Integer, Shift As Integer, x As Single, y As Single)
    

    On Error GoTo imgInventario_MouseDown_Err
    
    imgInventario.Picture = LoadInterface("boton-inventory-off.bmp")
    imgInventario.Tag = "2"

    'Call Inventario.DrawInventory
    
    Exit Sub

imgInventario_MouseDown_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmMain.imgInventario_MouseDown", Erl)
    Resume Next
    
End Sub

Private Sub imgInventario_MouseMove(button As Integer, Shift As Integer, x As Single, y As Single)
    
    On Error GoTo imgInventario_MouseMove_Err

    If imgInventario.Tag <> "1" And button = 0 Then
        imgInventario.Picture = LoadInterface("boton-inventory-over.bmp")
        imgInventario.Tag = "1"
        imgHechizos.Picture = Nothing
        imgHechizos.Tag = "0"
    End If

    
    Exit Sub

imgInventario_MouseMove_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmMain.imgInventario_MouseMove", Erl)
    Resume Next
    
End Sub

Private Sub imgOro_Click()
    
    On Error GoTo imgOro_Click_Err
    
    GldLbl_Click
    
    Exit Sub

imgOro_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmMain.imgOro_Click", Erl)
    Resume Next
    
End Sub



Private Sub ImgSeg_Click()
    Call WriteSafeToggle
End Sub

Private Sub ImgSegClan_Click()
    Call WriteSeguroClan
End Sub

Private Sub ImgSegParty_Click()
    Call WriteParyToggle
End Sub

Private Sub ImgSegResu_Click()
    Call WriteSeguroResu
End Sub

Private Sub Label1_Click()
    frmBancoCuenta.Show , frmMain
End Sub

Private Sub lblPorcLvl_MouseMove(button As Integer, Shift As Integer, x As Single, y As Single)

    If ShowPercentage Then
        lblPorcLvl.visible = False
        exp.visible = True
    End If

End Sub

Private Sub LlamaDeclan_Timer()
    
    On Error GoTo LlamaDeclan_Timer_Err
    
    frmMapaGrande.llamadadeclan.visible = False
    frmMapaGrande.Shape2.visible = False
    LlamaDeclan.enabled = False

    
    Exit Sub

LlamaDeclan_Timer_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmMain.LlamaDeclan_Timer", Erl)
    Resume Next
    
End Sub

Private Sub MANShp_Click()
    
    On Error GoTo MANShp_Click_Err
    
    manabar_Click
    
    Exit Sub

MANShp_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmMain.MANShp_Click", Erl)
    Resume Next
    
End Sub



Private Sub OpcionesBoton_MouseDown(button As Integer, Shift As Integer, x As Single, y As Single)
    
    On Error GoTo OpcionesBoton_MouseDown_Err
    
    OpcionesBoton.Picture = LoadInterface("opcionesoverdown.bmp")
    OpcionesBoton.Tag = "1"

    
    Exit Sub

OpcionesBoton_MouseDown_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmMain.OpcionesBoton_MouseDown", Erl)
    Resume Next
    
End Sub

Private Sub panelGM_Click()
    
    On Error GoTo panelGM_Click_Err
    
    frmPanelgm.Width = 4860
    Call WriteSOSShowList
    Call WriteGMPanel
    
    Me.SetFocus
    
    Exit Sub

panelGM_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmMain.panelGM_Click", Erl)
    Resume Next
    
End Sub

Private Sub panelInf_MouseMove(button As Integer, Shift As Integer, x As Single, y As Single)
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

    If QuestBoton.Tag = "1" Then
        QuestBoton.Picture = Nothing
        QuestBoton.Tag = "0"

    End If
    
    If ImgHogar.Tag = "1" Then
        ImgHogar.Picture = Nothing
        ImgHogar.Tag = "0"

    End If
        
    If ImgEstadisticas.Tag = "1" Then
        ImgEstadisticas.Picture = Nothing
        ImgEstadisticas.Tag = "0"

    End If
    
    If cmdLlavero.Tag = "1" Then
        cmdLlavero.Picture = Nothing
        cmdLlavero.Tag = "0"
    End If
    
    If Retar.Tag = "1" Then
        Retar.Picture = Nothing
        Retar.Tag = "0"
    End If

End Sub

Private Sub panelinferior_Click(Index As Integer)
Select Case Index
    Case 1
        QuePestañaInferior = 1
        GldLbl.visible = False
        'Label6.Visible = False
        stabar.Visible = False
        HpBar.Visible = False
        manabar.Visible = False
        hambar.Visible = False
        AGUbar.Visible = False
        Hpshp.Visible = False
        shieldBar.visible = False
        MANShp.Visible = False
        STAShp.Visible = False
        AGUAsp.Visible = False
        COMIDAsp.Visible = False
        Fuerzalbl.Visible = False
        AgilidadLbl.Visible = False
        oxigenolbl.Visible = False
        QuestBoton.Visible = True
        ImgHogar.Visible = True
        ImgEstadisticas.Visible = True
        lblWeapon.Visible = False
        lblShielder.Visible = False
        lblHelm.Visible = False
        lblArmor.Visible = False
        lblResis.Visible = False
        lbldm.Visible = False
        imgBugReport.Visible = True
        panelinferior(0).Picture = Nothing
        panelinferior(1).Picture = Nothing
        mapMundo.visible = True
        Image5.visible = True
        clanimg.visible = True
        Retar.visible = True
        cmdLlavero.visible = True
        QuestBoton.visible = True
        
        
        
        
        If SeguroGame Then
            ImgSeg = LoadInterface("boton-seguro-ciudadano-on.bmp")
        Else
            ImgSeg = LoadInterface("boton-seguro-ciudadano-off.bmp")
        End If
        
        If SeguroParty Then
            ImgSegParty = LoadInterface("boton-seguro-party-on.bmp")
        Else
            ImgSegParty = LoadInterface("boton-seguro-party-off.bmp")
        End If
        
        If SeguroClanX Then
            ImgSegClan = LoadInterface("boton-seguro-clan-on.bmp")
        Else
            ImgSegClan = LoadInterface("boton-seguro-clan-off.bmp")
        End If
        
        If SeguroResuX Then
            ImgSegResu = LoadInterface("boton-fantasma-on.bmp")
        Else
            ImgSegResu = LoadInterface("boton-fantasma-off.bmp")
        End If
        
        ImgSeg.visible = True
        ImgSegParty.visible = True
        ImgSegClan.visible = True
        ImgSegResu.visible = True
        
        
        panelInf.Picture = LoadInterface("ventanaprincipal_info.bmp")
    Case 0
        QuePestañaInferior = 0
        panelInf.Picture = LoadInterface("ventanaprincipal_stats.bmp")
        stabar.visible = True
        HpBar.visible = True
        If UserStats.maxman <> 0 Then
        manabar.visible = True
        End If
        hambar.visible = True
        AGUbar.visible = True
        Hpshp.visible = True
        MANShp.visible = True
        STAShp.visible = True
        AGUAsp.visible = True
        COMIDAsp.visible = True
        GldLbl.visible = True
       ' Label6.Visible = True
        Fuerzalbl.visible = True
        AgilidadLbl.visible = True
        oxigenolbl.visible = True
        QuestBoton.visible = False
        ImgHogar.visible = False
        ImgEstadisticas.visible = False
        lblWeapon.visible = True
        lblShielder.visible = True
        lblHelm.visible = True
        lblArmor.visible = True
        lblResis.visible = True
        lbldm.visible = True
        imgBugReport.visible = False
        panelinferior(0).Picture = Nothing
        panelinferior(1).Picture = Nothing
        mapMundo.visible = False
        Image5.visible = False
        clanimg.visible = False
        Retar.visible = False
        cmdLlavero.visible = False
        QuestBoton.visible = False
        
        ImgSeg.visible = False
        ImgSegParty.visible = False
        ImgSegClan.visible = False
        ImgSegResu.visible = False
End Select
End Sub

Private Sub Inventario_ItemDropped(ByVal Drag As Integer, ByVal Drop As Integer, ByVal x As Integer, ByVal y As Integer)
    
    On Error GoTo Inventario_ItemDropped_Err
    

    ' Si soltó un item en un slot válido
    
    If Drop > 0 Then
        If Drag <> Drop Then
            ' Muevo el item dentro del iventario
            Call WriteItemMove(Drag, Drop)
        End If
    End If

    
    Exit Sub

Inventario_ItemDropped_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmMain.Inventario_ItemDropped", Erl)
    Resume Next
    
End Sub

Private Sub picInv_Paint()
    
    On Error GoTo picInv_Paint_Err
    
    Inventario.ReDraw
    
    Exit Sub

picInv_Paint_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmMain.picInv_Paint", Erl)
    Resume Next
    
End Sub
Private Sub QuestBoton_MouseMove(button As Integer, Shift As Integer, x As Single, y As Single)
    
    On Error GoTo QuestBoton_MouseMove_Err
    

    If QuestBoton.Tag = "0" Then
        QuestBoton.Picture = LoadInterface("questover.bmp")
        QuestBoton.Tag = "1"

    End If

    
    Exit Sub

QuestBoton_MouseMove_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmMain.QuestBoton_MouseMove", Erl)
    Resume Next
    
End Sub

Private Sub QuestBoton_MouseUp(button As Integer, Shift As Integer, x As Single, y As Single)
    
    On Error GoTo QuestBoton_MouseUp_Err
    
    
    If pausa Then Exit Sub
    
    Call WriteQuestListRequest

    
    Exit Sub

QuestBoton_MouseUp_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmMain.QuestBoton_MouseUp", Erl)
    Resume Next
    
End Sub

Private Sub Label6_Click()
    
    On Error GoTo Label6_Click_Err
    
    Inventario.SelectGold

    If UserStats.GLD > 0 Then
        frmCantidad.Show , frmMain

    End If

    
    Exit Sub

Label6_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmMain.Label6_Click", Erl)
    Resume Next
    
End Sub

Private Sub Label7_Click()
    
    On Error GoTo Label7_Click_Err
    
    Call AddtoRichTextBox(frmMain.RecTxt, "No tenes mensajes nuevos.", 255, 255, 255, False, False, False)

    
    Exit Sub

Label7_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmMain.Label7_Click", Erl)
    Resume Next
    
End Sub

Private Sub lblPorcLvl_Click()
    
    On Error GoTo lblPorcLvl_Click_Err
    
    'Call WriteScrollInfo
    
    ShowPercentage = Not ShowPercentage
    
    Exit Sub

lblPorcLvl_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmMain.lblPorcLvl_Click", Erl)
    Resume Next
    
End Sub

Private Sub MacroLadder_Timer()
    
    
    On Error GoTo MacroLadder_Timer_Err
    
    If pausa Then Exit Sub
    
    If UserMacro.cantidad > 0 And UserMacro.Activado And UserStats.MinSTA > 0 Then
    
        Select Case UserMacro.tipo

            Case 1 'Alquimia
                Call WriteCraftAlquimista(UserMacro.Index)
                UserMacro.cantidad = UserMacro.cantidad - 1

            Case 3 'Sasteria
                Call WriteCraftSastre(UserMacro.Index)
                UserMacro.cantidad = UserMacro.cantidad - 1

            Case 4 'Herreria
                Call WriteCraftBlacksmith(UserMacro.Index)
                UserMacro.cantidad = UserMacro.cantidad - 1

            Case 6
                ' Jopi: Esto se hace desde el servidor
                'Call WriteWorkLeftClick(TargetXMacro, TargetYMacro, UsingSkill)

        End Select

    Else
        Call ResetearUserMacro

    End If

    'UserMacro.cantidad = UserMacro.cantidad - 1
    
    Exit Sub

MacroLadder_Timer_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmMain.MacroLadder_Timer", Erl)
    Resume Next
    
End Sub

Private Sub macrotrabajo_Timer()
    'If Inventario.SelectedItem = 0 Then
    '   DesactivarMacroTrabajo
    '   Exit Sub
    'End If
    
    On Error GoTo macrotrabajo_Timer_Err
    
    
    If pausa Then Exit Sub
    
    
    'Macros are disabled if not using Argentum!
    If Not Application.IsAppActive() Then
        DesactivarMacroTrabajo
        Exit Sub

    End If
    
    If UsingSkill = FundirMetal Or (UsingSkill = eSkill.Herreria And Not frmHerrero.visible) Then
        Call WriteWorkLeftClick(TargetXMacro, TargetYMacro, UsingSkill)
        UsingSkill = 0

    End If
    
    'If Inventario.OBJType(Inventario.SelectedItem) = eObjType.otWeapon Then
    If Not (frmCarp.visible = True) Then
        If frmMain.Inventario.IsItemSelected Then Call WriteUseItem(frmMain.Inventario.SelectedItem)
    End If

    
    Exit Sub

macrotrabajo_Timer_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmMain.macrotrabajo_Timer", Erl)
    Resume Next
    
End Sub

Public Sub ActivarMacroTrabajo()
    
    On Error GoTo ActivarMacroTrabajo_Err
    
    TargetXMacro = tX
    TargetYMacro = tY
    macrotrabajo.Interval = IntervaloTrabajoConstruir
    macrotrabajo.enabled = True
    Call AddtoRichTextBox(frmMain.RecTxt, "Macro Trabajo ACTIVADO", 0, 200, 200, False, True, False)

    
    Exit Sub

ActivarMacroTrabajo_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmMain.ActivarMacroTrabajo", Erl)
    Resume Next
    
End Sub

Public Sub DesactivarMacroTrabajo()
    
    On Error GoTo DesactivarMacroTrabajo_Err
    
    TargetXMacro = 0
    TargetYMacro = 0
    macrotrabajo.enabled = False
    MacroBltIndex = 0
    UsingSkill = 0
    MousePointer = vbDefault
    Call AddtoRichTextBox(frmMain.RecTxt, "Macro Trabajo DESACTIVADO", 0, 200, 200, False, True, False)

    
    Exit Sub

DesactivarMacroTrabajo_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmMain.DesactivarMacroTrabajo", Erl)
    Resume Next
    
End Sub

Private Sub MenuOpciones_Click()

End Sub

Private Sub manabar_Click()
    
    On Error GoTo manabar_Click_Err
    

    If UserStats.minman = UserStats.maxman Then Exit Sub
            
    If UserStats.estado = 1 Then

        With FontTypes(FontTypeNames.FONTTYPE_INFO)
            Call ShowConsoleMsg("¡Estás muerto!", .red, .green, .blue, .bold, .italic)

        End With

        Exit Sub

    End If
    
    Call WriteMeditate

    
    Exit Sub

manabar_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmMain.manabar_Click", Erl)
    Resume Next
    
End Sub

Private Sub mapMundo_Click()
    
    On Error GoTo mapMundo_Click_Err
    
    ExpMult = 1
    OroMult = 1
    Call frmMapaGrande.CalcularPosicionMAPA
    frmMapaGrande.Picture = LoadInterface("ventanamapa.bmp")
    frmMapaGrande.Show , frmMain

    
    Exit Sub

mapMundo_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmMain.mapMundo_Click", Erl)
    Resume Next
    
End Sub

Private Sub mapMundo_MouseMove(button As Integer, Shift As Integer, x As Single, y As Single)
    
    On Error GoTo mapMundo_MouseMove_Err
    

    If mapMundo.Tag = "0" Then
        mapMundo.Picture = LoadInterface("boton-mapa-over.bmp")
        mapMundo.Tag = "1"

    End If

    
    Exit Sub

mapMundo_MouseMove_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmMain.mapMundo_MouseMove", Erl)
    Resume Next
    
End Sub

Private Sub MiniMap_MouseDown(button As Integer, Shift As Integer, x As Single, y As Single)
    
    On Error GoTo MiniMap_MouseDown_Err
    
    Call GetMinimapPosition(x, y)
    

    If button = vbRightButton Then
        Call ParseUserCommand("/TELEP YO " & UserMap & " " & CByte(x) & " " & CByte(y))
        Exit Sub

    End If
  
    ExpMult = 1
    OroMult = 1
    Call frmMapaGrande.CalcularPosicionMAPA
    frmMapaGrande.Picture = LoadInterface("ventanamapa.bmp")
    frmMapaGrande.Show , frmMain
  
    
    Exit Sub

MiniMap_MouseDown_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmMain.MiniMap_MouseDown", Erl)
    Resume Next
    
End Sub

Private Sub MiniMap_MouseMove(button As Integer, Shift As Integer, x As Single, y As Single)
    
    On Error GoTo MiniMap_MouseMove_Err
    

    If mapMundo.Tag = "1" Then
        mapMundo.Picture = Nothing
        mapMundo.Tag = "0"

    End If

    
    Exit Sub

MiniMap_MouseMove_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmMain.MiniMap_MouseMove", Erl)
    Resume Next
    
End Sub

Private Sub mnuEquipar_Click()
    
    On Error GoTo mnuEquipar_Click_Err
    

    If MainTimer.Check(TimersIndex.UseItemWithU) Then Call WriteEquipItem(Inventario.SelectedItem)

    
    Exit Sub

mnuEquipar_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmMain.mnuEquipar_Click", Erl)
    Resume Next
    
End Sub

Private Sub mnuNPCComerciar_Click()
    
    On Error GoTo mnuNPCComerciar_Click_Err
    
    Call WriteLeftClick(tX, tY)
    Call WriteCommerceStart

    
    Exit Sub

mnuNPCComerciar_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmMain.mnuNPCComerciar_Click", Erl)
    Resume Next
    
End Sub

Private Sub mnuNpcDesc_Click()
    
    On Error GoTo mnuNpcDesc_Click_Err
    
    Call WriteLeftClick(tX, tY)

    
    Exit Sub

mnuNpcDesc_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmMain.mnuNpcDesc_Click", Erl)
    Resume Next
    
End Sub

Private Sub mnuTirar_Click()
    
    On Error GoTo mnuTirar_Click_Err
    
    Call TirarItem

    
    Exit Sub

mnuTirar_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmMain.mnuTirar_Click", Erl)
    Resume Next
    
End Sub

Private Sub mnuUsar_Click()
    
    On Error GoTo mnuUsar_Click_Err
    

    If frmMain.Inventario.IsItemSelected Then Call WriteUseItem(frmMain.Inventario.SelectedItem)

    
    Exit Sub

mnuUsar_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmMain.mnuUsar_Click", Erl)
    Resume Next
    
End Sub

Private Sub NameMapa_MouseMove(button As Integer, Shift As Integer, x As Single, y As Single)
    
    On Error GoTo NameMapa_MouseMove_Err
    

    If mapMundo.Tag = "1" Then
        mapMundo.Picture = Nothing
        mapMundo.Tag = "0"

    End If

    
    Exit Sub

NameMapa_MouseMove_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmMain.NameMapa_MouseMove", Erl)
    Resume Next
    
End Sub

Private Sub onlines_Click()
    
    On Error GoTo onlines_Click_Err
    
    Call WriteOnline

    
    Exit Sub

onlines_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmMain.onlines_Click", Erl)
    Resume Next
    
End Sub

Private Sub OpcionesBoton_MouseMove(button As Integer, Shift As Integer, x As Single, y As Single)
    
    On Error GoTo OpcionesBoton_MouseMove_Err
    

    If OpcionesBoton.Tag = "0" Then
        OpcionesBoton.Picture = LoadInterface("opcionesover.bmp")
        OpcionesBoton.Tag = "1"

    End If

    
    Exit Sub

OpcionesBoton_MouseMove_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmMain.OpcionesBoton_MouseMove", Erl)
    Resume Next
    
End Sub

Private Sub OpcionesBoton_MouseUp(button As Integer, Shift As Integer, x As Single, y As Single)
    
    On Error GoTo OpcionesBoton_MouseUp_Err
    
    Call frmOpciones.Init

    
    Exit Sub

OpcionesBoton_MouseUp_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmMain.OpcionesBoton_MouseUp", Erl)
    Resume Next
    
End Sub

Private Sub Panel_MouseMove(button As Integer, Shift As Integer, x As Single, y As Single)
    
    On Error GoTo Panel_MouseMove_Err
    

    ObjLbl.visible = False

    If cmdlanzar.Tag = "1" Then
        cmdlanzar.Picture = Nothing
        cmdlanzar.Tag = "0"

    End If
    
    If imgInventario.Tag <> "0" Then
        imgInventario.Picture = Nothing
        imgInventario.Tag = "0"

    End If

    If imgHechizos.Tag <> "0" Then
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

    
    Exit Sub

Panel_MouseMove_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmMain.Panel_MouseMove", Erl)
    Resume Next
    
End Sub

Private Sub picInv_MouseMove(button As Integer, Shift As Integer, x As Single, y As Single)
    
    On Error GoTo picInv_MouseMove_Err
    

    Dim Slot As Byte

    UsaMacro = False
    
    Slot = Inventario.GetSlot(x, y)
    
    If Slot <= 0 Then
        ObjLbl.visible = False
        Exit Sub
    End If
    
    If Inventario.Amount(Slot) > 0 Then
    
        ObjLbl.visible = True
        
        Select Case ObjData(Inventario.ObjIndex(Slot)).ObjType

            Case eObjType.otWeapon
                ObjLbl = Inventario.ItemName(Slot) & " (" & Inventario.Amount(Slot) & ")" & vbCrLf & "Daño: " & ObjData(Inventario.OBJIndex(Slot)).MinHit & "/" & ObjData(Inventario.OBJIndex(Slot)).MaxHit

            Case eObjType.otArmadura
                ObjLbl = Inventario.ItemName(Slot) & " (" & Inventario.Amount(Slot) & ")" & vbCrLf & "Defensa: " & ObjData(Inventario.ObjIndex(Slot)).MinDef & "/" & ObjData(Inventario.ObjIndex(Slot)).MaxDef

            Case eObjType.otCASCO
                ObjLbl = Inventario.ItemName(Slot) & " (" & Inventario.Amount(Slot) & ")" & vbCrLf & "Defensa: " & ObjData(Inventario.ObjIndex(Slot)).MinDef & "/" & ObjData(Inventario.ObjIndex(Slot)).MaxDef

            Case eObjType.otESCUDO
                ObjLbl = Inventario.ItemName(Slot) & " (" & Inventario.Amount(Slot) & ")" & vbCrLf & "Defensa: " & ObjData(Inventario.ObjIndex(Slot)).MinDef & "/" & ObjData(Inventario.ObjIndex(Slot)).MaxDef

            Case Else
                ObjLbl = Inventario.ItemName(Slot) & " (" & Inventario.Amount(Slot) & ")" & vbCrLf & ObjData(Inventario.ObjIndex(Slot)).Texto

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
        ObjLbl.visible = False

    End If

    
    Exit Sub

picInv_MouseMove_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmMain.picInv_MouseMove", Erl)
    Resume Next
    
End Sub

Public Sub CompletarEnvioMensajes()
    
    On Error GoTo CompletarEnvioMensajes_Err

    Select Case SendingType
        Case 1
            SendTxt.Text = ""
        Case 2
            SendTxt.Text = "-"
        Case 3
            SendTxt.Text = ("\" & sndPrivateTo & " ")
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
    Exit Sub

CompletarEnvioMensajes_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmMain.CompletarEnvioMensajes", Erl)
    Resume Next
    
End Sub

Private Sub RecTxt_MouseUp(button As Integer, Shift As Integer, x As Single, y As Single)
    
    On Error GoTo RecTxt_MouseUp_Err
    

    

    If button = 1 Then

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
            If Not SendTxtCmsg.visible Then
                destinatario = Replace(destinatario, " ", "+")
    
                sndPrivateTo = destinatario
                SendTxt.Text = ("\" & sndPrivateTo & " ")
    
                stxtbuffer = SendTxt.Text
                SendTxt.SelStart = Len(SendTxt.Text)
                    
                SendTxt.visible = True
                SendTxt.SetFocus
            End If

        End If

    End If

    
    Exit Sub

RecTxt_MouseUp_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmMain.RecTxt_MouseUp", Erl)
    Resume Next
    
End Sub

Private Sub refuerzolanzar_Click()
    
    On Error GoTo refuerzolanzar_Click_Err
    
    Call cmdLanzar_Click

    
    Exit Sub

refuerzolanzar_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmMain.refuerzolanzar_Click", Erl)
    Resume Next
    
End Sub

Private Sub refuerzolanzar_MouseMove(button As Integer, Shift As Integer, x As Single, y As Single)
    
    On Error GoTo refuerzolanzar_MouseMove_Err
    
    UsaMacro = False
    CnTd = 0

    If cmdlanzar.Tag = "0" Then
        cmdlanzar.Picture = LoadInterface("lanzarmarcado.bmp")
        cmdlanzar.Tag = "1"

    End If

    
    Exit Sub

refuerzolanzar_MouseMove_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmMain.refuerzolanzar_MouseMove", Erl)
    Resume Next
    
End Sub

Private Sub renderer_MouseUp(button As Integer, Shift As Integer, x As Single, y As Single)
    
    On Error GoTo renderer_MouseUp_Err

    clicX = x
    clicY = y
    If button = vbLeftButton Then
        If HandleMouseInput(x, y) Then
        ElseIf Pregunta Then
            If x >= 419 And x <= 433 And y >= 243 And y <= 260 Then
                If PreguntaLocal Then
    
                    Select Case PreguntaNUM
    
                        Case 1
                            Pregunta = False
                            DestItemSlot = 0
                            DestItemCant = 0
                            PreguntaLocal = False
                            
                        Case 2 ' Denunciar
                            Pregunta = False
                            PreguntaLocal = False
    
                    End Select
    
                Else
                    Call WriteResponderPregunta(False)
                    Pregunta = False
    
                End If
                
                Exit Sub
    
            ElseIf x >= 443 And x <= 458 And y >= 243 And y <= 260 Then
                If PreguntaLocal Then
    
                    Select Case PreguntaNUM
    
                        Case 1 '¿Destruir item?
                            Call WriteDrop(DestItemSlot, DestItemCant)
                            Pregunta = False
                            PreguntaLocal = False
                            
                        Case 2 ' Denunciar
                            Call WriteDenounce(TargetName)
                            Pregunta = False
                            PreguntaLocal = False
    
                    End Select
    
                Else
                    Call WriteResponderPregunta(True)
                    Pregunta = False
    
                End If
                
                Exit Sub
    
            End If
    
        End If
    
    ElseIf button = vbRightButton Then
        
        Dim charindex As Integer
        charindex = MapData(tX, tY).charindex
        
        If charindex = 0 Then
            charindex = MapData(tX, tY + 1).charindex
        End If
        
        If charindex <> 0 And charindex <> UserCharIndex Then
            Dim Frm As Form
            
            Call WriteLeftClick(tX, tY)

            TargetX = tX
            TargetY = tY
        
            If charlist(charindex).EsMascota Then
                Set Frm = MenuNPC
            
            ElseIf Not charlist(charindex).esNpc Then
                
                TargetName = charlist(charindex).nombre
                
                If charlist(UserCharIndex).priv > 0 And Shift = 0 Then
                    Set Frm = MenuGM
                Else
                    Set Frm = MenuUser
                End If
            End If
            
            If Not Frm Is Nothing Then
                Call Frm.Show
            
                Frm.Left = Me.Left + (renderer.Left + x + 1) * screen.TwipsPerPixelX
                    
                If (renderer.Top + y) * screen.TwipsPerPixelY + Frm.Height > Me.Height Then
                    Frm.Top = Me.Top + (renderer.Top + y) * screen.TwipsPerPixelY - Frm.Height
                Else
                    Frm.Top = Me.Top + (renderer.Top + y) * screen.TwipsPerPixelY
                End If
                
                Set Frm = Nothing
            End If
        End If

    End If
    
    
    Exit Sub

renderer_MouseUp_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmMain.renderer_MouseUp", Erl)
    Resume Next
    
End Sub

Private Sub renderer_MouseMove(button As Integer, Shift As Integer, x As Single, y As Single)
    
    On Error GoTo renderer_MouseMove_Err

    DisableURLDetect

    Call Form_MouseMove(button, Shift, renderer.Left + x, renderer.Top + y)

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
    
    'LucesCuadradas.Light_Remove (10)
    
    'LucesCuadradas.Light_Create tX, tY, &HFFFFFFF, 1, 10
    'LucesCuadradas.Light_Render_All
    
    
    Exit Sub

renderer_MouseMove_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmMain.renderer_MouseMove", Erl)
    Resume Next
    
End Sub

Private Sub renderer_MouseDown(button As Integer, Shift As Integer, x As Single, y As Single)
    
    On Error GoTo renderer_MouseDown_Err
    
    If SendTxt.visible Then SendTxt.SetFocus
    MouseBoton = button
    MouseShift = Shift

    If frmComerciar.visible Then Unload frmComerciar
    If frmBancoObj.visible Then Unload frmBancoObj
    If frmEstadisticas.visible Then Unload frmEstadisticas
    If frmGoliath.visible Then Unload frmGoliath
    If frmMapaGrande.visible Then frmMapaGrande.visible = False
    If FrmViajes.visible Then Unload FrmViajes
    If frmCantidad.visible Then Unload frmCantidad
    If FrmGrupo.visible Then Unload FrmGrupo
    If FrmGmAyuda.visible Then Unload FrmGmAyuda
    If frmGuildAdm.visible Then Unload frmGuildAdm
    If frmHerrero.visible Then Unload frmHerrero
    If FrmSastre.visible Then Unload FrmSastre
    If frmAlqui.visible Then Unload frmAlqui
    If frmCarp.visible Then Unload frmCarp
    If MenuUser.visible Then Unload MenuUser
    If MenuGM.visible Then Unload MenuGM
    If MenuNPC.visible Then Unload MenuNPC
    
    Exit Sub

renderer_MouseDown_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmMain.renderer_MouseDown", Erl)
    Resume Next
    
End Sub

Private Sub renderer_DblClick()
    
    On Error GoTo renderer_DblClick_Err
    
    Call Form_DblClick

    
    Exit Sub

renderer_DblClick_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmMain.renderer_DblClick", Erl)
    Resume Next
    
End Sub

Private Sub renderer_Click()
    On Error GoTo renderer_Click_Err
    Call Form_Click
    If SendTxt.visible Then SendTxt.SetFocus
    If SendTxtCmsg.visible Then SendTxtCmsg.SetFocus
    Exit Sub

renderer_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmMain.renderer_Click", Erl)
    Resume Next
    
End Sub

Private Sub Retar_Click()
    Call ParseUserCommand("/RETAR")
End Sub

Private Sub Retar_MouseDown(button As Integer, Shift As Integer, x As Single, y As Single)
    Retar.Picture = LoadInterface("boton-retos-off.bmp")
    Retar.Tag = "1"
End Sub

Private Sub Retar_MouseMove(button As Integer, Shift As Integer, x As Single, y As Single)
    If Retar.Tag = "0" Then
        Retar.Picture = LoadInterface("boton-retos-over.bmp")
        Retar.Tag = "1"
    End If
End Sub

Private Sub Retar_MouseUp(button As Integer, Shift As Integer, x As Single, y As Single)
    Retar.Picture = Nothing
    Retar.Tag = "0"
End Sub

Private Sub SendTxt_KeyUp(KeyCode As Integer, Shift As Integer)
    
    On Error GoTo SendTxt_KeyUp_Err
    'Send text
    If KeyCode = vbKeyReturn Then
        If LenB(stxtbuffer) <> 0 Then
            Call HandleChatMsg(stxtbuffer)
        End If
        stxtbuffer = ""
        SendTxt.Text = ""
        KeyCode = 0
        Dim tiempoTranscurridoCartel As Double
        tiempoTranscurridoCartel = GetTickCount - StartOpenChatTime
        Call computeLastElapsedTimeChat(tiempoTranscurridoCartel)
        SendTxt.visible = False
    End If
    Exit Sub
SendTxt_KeyUp_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmMain.SendTxt_KeyUp", Erl)
    Resume Next
    
End Sub

Private Sub computeLastElapsedTimeChat(ByVal tiempoTranscurridoCartel As Double)
    Dim i As Long
    
    For i = 2 To 6
        LastElapsedTimeChat(i - 1) = LastElapsedTimeChat(i)
    Next i
    
    LastElapsedTimeChat(6) = tiempoTranscurridoCartel
        
    'HarThaoS: Calculo el mínimo y máximo de mis carteleos
    Dim min As Double, max As Double
    
    min = LastElapsedTimeChat(6)
    max = LastElapsedTimeChat(6)
    
    For i = 1 To 6
        If LastElapsedTimeChat(i) > max Then max = LastElapsedTimeChat(i)
        If LastElapsedTimeChat(i) < min Then min = LastElapsedTimeChat(i)
    Next i
    
    If (max - min) > 0 And (max - min) < 12 Then
         Call WriteLogMacroClickHechizo(tMacro.borrarCartel)
    End If
    
    
    
End Sub

Private Sub SendTxtCmsg_KeyUp(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    If SendTxtCmsg.SelStart > 2 Then Call ParseUserCommand("/CMSG " & SendTxtCmsg.Text)
    SendTxtCmsg.visible = False
    SendTxtCmsg.Text = ""
    Call DialogosClanes.toggle_dialogs_visibility(False)
  End If
End Sub

Private Sub ShowFPS_Timer()
    
    On Error GoTo ShowFPS_Timer_Err
    
    fps.Caption = "FPS: " & engine.fps
    
    Exit Sub

ShowFPS_Timer_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmMain.ShowFPS_Timer", Erl)
    Resume Next
    
End Sub

Private Sub cerrarcuenta_Timer()
    
    On Error GoTo cerrarcuenta_Timer_Err
    
    Unload frmConnect
    Unload frmCrearPersonaje
    cerrarcuenta.enabled = False
    Exit Sub

cerrarcuenta_Timer_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmMain.cerrarcuenta_Timer", Erl)
    Resume Next
End Sub

Private Sub TimerLluvia_Timer()
    
    On Error GoTo TimerLluvia_Timer_Err
    

    If bRain Then

        If CantPartLLuvia < 250 Then

            CantPartLLuvia = CantPartLLuvia + 1
            Graficos_Particulas.Particle_Group_Edit (MeteoIndex)
        Else
            CantPartLLuvia = 250
            TimerLluvia.enabled = False

        End If

    Else

        If CantPartLLuvia > 0 Then
            CantPartLLuvia = CantPartLLuvia - 1
            Graficos_Particulas.Particle_Group_Edit (MeteoIndex)
        Else
    
            Call Graficos_Particulas.Engine_MeteoParticle_Set(-1)
            CantPartLLuvia = 0
            TimerLluvia.enabled = False

        End If

    End If

    
    Exit Sub

TimerLluvia_Timer_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmMain.TimerLluvia_Timer", Erl)
    Resume Next
    
End Sub

Private Sub TimerMusica_Timer()

End Sub

Private Sub TimerNiebla_Timer()
    
    On Error GoTo TimerNiebla_Timer_Err
    

    If bNiebla Then

        If AlphaNiebla < MaxAlphaNiebla Then
            AlphaNiebla = AlphaNiebla + 1
        Else
            AlphaNiebla = MaxAlphaNiebla
            TimerNiebla.enabled = False

        End If

    Else

        If AlphaNiebla > 0 Then
            AlphaNiebla = AlphaNiebla - 1

        Else
            AlphaNiebla = 0
            MaxAlphaNiebla = 0
            TimerNiebla.enabled = False
        End If

    End If

    
    Exit Sub

TimerNiebla_Timer_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmMain.TimerNiebla_Timer", Erl)
    Resume Next
    
End Sub


Private Sub cmdLanzar_Click()
On Error GoTo cmdLanzar_Click_Err
    Call UseSpell(hlst.ListIndex + 1, hlst.List(hlst.ListIndex))
    Exit Sub
cmdLanzar_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmMain.cmdLanzar_Click", Erl)
    Resume Next
End Sub

Private Sub CmdLanzar_MouseMove(button As Integer, Shift As Integer, x As Single, y As Single)
    
    On Error GoTo CmdLanzar_MouseMove_Err
    
    UsaMacro = False
    CnTd = 0

    If cmdlanzar.Tag = "0" Then
        cmdlanzar.Picture = LoadInterface("boton-lanzar-over.bmp")
        cmdlanzar.Tag = "1"

    End If

    
    Exit Sub

CmdLanzar_MouseMove_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmMain.CmdLanzar_MouseMove", Erl)
    Resume Next
    
End Sub

Public Sub Form_Click()
    
    Call OnClick(MouseBoton, MouseShift)
    
End Sub

Private Sub Form_DblClick()
    
    On Error GoTo Form_DblClick_Err
    

    '**************************************************************
    'Author: Unknown
    'Last Modify Date: 12/27/2007
    '12/28/2007: ByVal - Chequea que la ventana de comercio y boveda no este abierta al hacer doble clic a un comerciante, sobrecarga la lista de items.
    '**************************************************************
    If Not frmComerciar.visible And Not frmBancoObj.visible Then
        If MouseBoton = vbLeftButton Then

            Call WriteDoubleClick(tX, tY)
        End If

    End If

    
    Exit Sub

Form_DblClick_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmMain.Form_DblClick", Erl)
    Resume Next
    
End Sub

Private Sub Form_Load()
    
    On Error GoTo Form_Load_Err

    Debug.Assert Not FormParser Is Nothing
    
    Call FormParser.Parse_Form(frmMain)
    
    MenuNivel = 1
    Me.Caption = "Argentum20" 'hay que poner 20 aniversario
    
    Set hlst = New clsGraphicalList
    Call hlst.Initialize(Me.picHechiz, RGB(200, 190, 190))

    loadButtons
    
    Exit Sub

Form_Load_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmMain.Form_Load", Erl)
    Resume Next
    
End Sub

Private Sub Form_MouseMove(button As Integer, Shift As Integer, x As Single, y As Single)
    
    On Error GoTo Form_MouseMove_Err

    ' Disable links checking (not over consola)
    StopCheckingLinks
    
    If PantallaCompleta = 0 And button = vbLeftButton Then
        If MoverVentana = 1 Then
            If Not UserMoving Then
                ' Mover form sólo en la parte superior
                If y < 30 Then MoverForm

                'Call Auto_Drag(Me.hwnd)
            End If

        End If

    End If

    MouseX = x - renderer.Left
    MouseY = y - renderer.Top
    ObjLbl.visible = False
    
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
    
    If ImgHogar.Tag = "1" Then
        ImgHogar.Picture = Nothing
        ImgHogar.Tag = "0"

    End If
    
    If cmdLlavero.Tag = "1" Then
        cmdLlavero.Picture = Nothing
        cmdLlavero.Tag = "0"
    End If

    If ShowPercentage Then
        lblPorcLvl.visible = True
        exp.visible = False
    Else
        lblPorcLvl.visible = False
        exp.visible = True
    End If
    
    If Retar.Tag = "1" Then
        Retar.Picture = Nothing
        Retar.Tag = "0"
    End If
    
    MenuUser.LostFocus
    MenuGM.LostFocus
    MenuNPC.LostFocus
    
    Exit Sub

Form_MouseMove_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmMain.Form_MouseMove", Erl)
    Resume Next
    
End Sub




Private Sub picInv_DblClick()
    On Error GoTo picInv_DblClick_Err

    If Not picInv.visible Then Exit Sub
    Call UserItemClick
    Exit Sub
picInv_DblClick_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmMain.picInv_DblClick", Erl)
    Resume Next
End Sub

Private Function countRepts(ByVal packet As Long)
    
End Function

Private Sub RecTxt_MouseMove(button As Integer, Shift As Integer, x As Single, y As Single)
    
    On Error GoTo RecTxt_MouseMove_Err
    
    StartCheckingLinks

    
    Exit Sub

RecTxt_MouseMove_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmMain.RecTxt_MouseMove", Erl)
    Resume Next
    
End Sub

Private Sub RecTxt_KeyDown(KeyCode As Integer, Shift As Integer)
    
    On Error GoTo RecTxt_KeyDown_Err
    

    If picInv.visible Then
        picInv.SetFocus
    Else

        'hlst.SetFocus
    End If

    
    Exit Sub

RecTxt_KeyDown_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmMain.RecTxt_KeyDown", Erl)
    Resume Next
    
End Sub

Private Sub SendTxt_Change()
    
    On Error GoTo SendTxt_Change_Err
    

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

        Dim tempStr   As String

        Dim CharAscii As Integer
        
        For i = 1 To Len(SendTxt.Text)
            CharAscii = Asc(mid$(SendTxt.Text, i, 1))

            If CharAscii >= vbKeySpace And CharAscii <= 250 Then
                tempStr = tempStr & Chr$(CharAscii)

            End If

        Next i
        
        If tempStr <> SendTxt.Text Then
            'We only set it if it's different, otherwise the event will be raised
            'constantly and the client will crush
            SendTxt.Text = tempStr

        End If
        
        stxtbuffer = SendTxt.Text

    End If

    
    Exit Sub

SendTxt_Change_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmMain.SendTxt_Change", Erl)
    Resume Next
    
End Sub

Private Sub SendTxt_KeyPress(KeyAscii As Integer)
    
    On Error GoTo SendTxt_KeyPress_Err
    

    If Not (KeyAscii = vbKeyBack) And Not (KeyAscii >= vbKeySpace And KeyAscii <= 250) Then KeyAscii = 0

    
    Exit Sub

SendTxt_KeyPress_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmMain.SendTxt_KeyPress", Erl)
    Resume Next
    
End Sub

Private Function InGameArea() As Boolean
    
    On Error GoTo InGameArea_Err
    

    If clicX < renderer.Left Or clicX > renderer.Left + (32 * 23) Then Exit Function
    If clicY < renderer.Top Or clicY > renderer.Top + (32 * 17) Then Exit Function
    InGameArea = True

    
    Exit Function

InGameArea_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmMain.InGameArea", Erl)
    Resume Next
    
End Function

Private Sub MoverForm()
    
    On Error GoTo moverForm_Err
    

    Dim res As Long

    ReleaseCapture
    res = SendMessage(Me.hwnd, WM_SYSCOMMAND, MOUSE_MOVE, 0)

    
    Exit Sub

moverForm_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmMain.moverForm", Erl)
    Resume Next
    
End Sub

Private Sub imgSpellInfo_Click()
    
    On Error GoTo imgSpellInfo_Click_Err
    

    If hlst.ListIndex <> -1 Then
        Call WriteSpellInfo(hlst.ListIndex + 1)

    End If

    
    Exit Sub

imgSpellInfo_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmMain.imgSpellInfo_Click", Erl)
    Resume Next
    
End Sub

Private Sub UpdateDaytime_Timer()
    ' Si no hay luz de mapa, usamos la luz ambiental
    
    On Error GoTo UpdateDaytime_Timer_Err
    
    Call RevisarHoraMundo
    
    Exit Sub

UpdateDaytime_Timer_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmMain.UpdateDaytime_Timer", Erl)
    Resume Next
    
End Sub

Private Sub UpdateLight_Timer()
    
    On Error GoTo UpdateLight_Timer_Err
    
    
    If light_transition < 1# Then
        light_transition = light_transition + STEP_LIGHT_TRANSITION * UpdateLight.Interval
        
        If light_transition > 1# Then light_transition = 1#
        
        Call LerpRGBA(global_light, last_light, next_light, light_transition)
        Call MapUpdateGlobalLight
    End If
    
    
    Exit Sub

UpdateLight_Timer_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmMain.UpdateLight_Timer", Erl)
    Resume Next
    
End Sub

Public Sub ShowStats()
    panelinferior_Click 0
End Sub

Public Sub SetMinimapPosition(ByVal Jugador As Integer, ByVal x As Integer, ByVal y As Integer)
    personaje(Jugador).Left = (x - HalfWindowTileWidth - 2) * (100 / (100 - 2 * HalfWindowTileWidth - 4)) - personaje(Jugador).Width \ 2 - 1
    personaje(Jugador).Top = (y - HalfWindowTileHeight - 1) * (100 / (100 - 2 * HalfWindowTileHeight - 2)) - personaje(Jugador).Height \ 2 - 1
End Sub

Public Sub GetMinimapPosition(ByRef x As Single, ByRef y As Single)
    x = x * (100 - 2 * HalfWindowTileWidth - 4) / 100 + HalfWindowTileWidth + 2
    y = y * (100 - 2 * HalfWindowTileHeight - 4) / 100 + HalfWindowTileHeight + 2
End Sub

Private Sub imgDeleteItem_Click()
    If Not frmMain.Inventario.IsItemSelected Then
        Call AddtoRichTextBox(frmMain.RecTxt, "No tienes seleccionado ningún item", 255, 255, 255, False, False, False)
    Else
        If MsgBox("Seguro que desea eliminar el item?", vbYesNo, "Eliminar objeto") = vbYes Then
            Call WriteDeleteItem(frmMain.Inventario.SelectedItem)
        End If
    End If
End Sub

Public Sub UpdateHpBar()
    Dim CurrentHp As Long
    If UserStats.MaxHp > 0 Then
        Dim FullSize As Long
        CurrentHp = UserStats.MinHp + UserStats.HpShield
        FullSize = max(UserStats.MinHp + UserStats.HpShield, UserStats.MaxHp)
        frmMain.Hpshp.Width = UserStats.MinHp / FullSize * 216
        frmMain.HpBar.Caption = CurrentHp & " / " & UserStats.MaxHp
        frmMain.shieldBar.Left = frmMain.Hpshp.Left + frmMain.Hpshp.Width
        frmMain.shieldBar.Width = UserStats.HpShield / FullSize * 216
    Else
        frmMain.Hpshp.Width = 0
        frmMain.shieldBar.Width = 0
    End If
    
    If QuePestañaInferior = 0 Then
        frmMain.Hpshp.visible = (UserStats.MinHp > 0)
        frmMain.shieldBar.visible = UserStats.HpShield > 0
    End If
    If UserStats.HpShield > 0 Then
        frmMain.HpBar.Caption = UserStats.MinHp & " / " & UserStats.MaxHp & " + " & UserStats.HpShield
    Else
        frmMain.HpBar.Caption = UserStats.MinHp & " / " & UserStats.MaxHp
    End If
    
End Sub


Public Sub UpdateStatsLayout()
    If UserStats.PasarNivel > 0 Then
        frmMain.lblPorcLvl.Caption = Round(UserStats.exp * (100 / UserStats.PasarNivel), 2) & "%"
        frmMain.exp.Caption = PonerPuntos(UserStats.exp) & "/" & PonerPuntos(UserStats.PasarNivel)
        frmMain.EXPBAR.Width = UserStats.exp / UserStats.PasarNivel * 235
    Else
        frmMain.EXPBAR.Width = 235
        frmMain.lblPorcLvl.Caption = "¡Nivel máximo!" 'nivel maximo
        frmMain.exp.Caption = "¡Nivel máximo!"
    End If
    
    Call frmMain.UpdateHpBar

    If UserStats.maxman > 0 Then
        frmMain.MANShp.Width = UserStats.minman / UserStats.maxman * 216
        frmMain.manabar.Caption = UserStats.minman & " / " & UserStats.maxman

        If QuePestañaInferior = 0 Then
            frmMain.MANShp.visible = (UserStats.minman > 0)
            frmMain.manabar.visible = True

        End If

    Else
        frmMain.manabar.visible = False
        frmMain.MANShp.Width = 0
        frmMain.MANShp.visible = False

    End If
    
    If UserStats.MaxSTA > 0 Then
        frmMain.STAShp.Width = UserStats.MinSTA / UserStats.MaxSTA * 89
    Else
        frmMain.STAShp.Width = 0
    End If

    frmMain.stabar.Caption = UserStats.MinSTA & " / " & UserStats.MaxSTA
    
    If QuePestañaInferior = 0 Then
        frmMain.STAShp.visible = (UserStats.MinSTA > 0)

    End If
    
    If UserStats.GLD <= 100000 Then
        frmMain.GldLbl.ForeColor = vbRed
    Else
        frmMain.GldLbl.ForeColor = &H80FFFF
    End If

    frmMain.GldLbl.Caption = PonerPuntos(UserStats.GLD)
    frmMain.lblLvl.Caption = ListaClases(UserStats.Clase) & " - Nivel " & UserStats.Lvl
    
    
End Sub

Public Sub UnlockInvslot(ByVal UserInvLevel As Integer)
    Dim i As Integer
    For i = 1 To UserInvLevel
        imgInvLock(i - 1).Picture = LoadInterface("inventoryunlocked.bmp")
    Next i
End Sub

Public Sub UpdateManaBar()
    If UserStats.maxman > 0 Then
        frmMain.MANShp.Width = UserStats.minman / UserStats.maxman * 216
        frmMain.manabar.Caption = UserStats.minman & " / " & UserStats.maxman

        If QuePestañaInferior = 0 Then
            frmMain.MANShp.visible = (UserStats.minman > 0)
            frmMain.manabar.visible = True
        End If
    Else
        frmMain.MANShp.Width = 0
        frmMain.manabar.visible = False
        frmMain.MANShp.visible = False
    End If
End Sub

Public Sub UpdateFoodState()
    frmMain.AGUAsp.Width = UserStats.MinAGU / UserStats.MaxAGU * 32
    frmMain.COMIDAsp.Width = UserStats.MinHAM / UserStats.MaxHAM * 32
    frmMain.AGUbar.Caption = UserStats.MinAGU '& " / " & UserMaxAGU
    frmMain.hambar.Caption = UserStats.MinHAM ' & " / " & UserMaxHAM
    
    If QuePestañaInferior = 0 Then
        frmMain.AGUAsp.visible = (UserStats.MinAGU > 0)
        frmMain.COMIDAsp.visible = (UserStats.MinHAM > 0)
    End If
End Sub

Public Sub UpdateStamina()
    frmMain.STAShp.Width = UserStats.MinSTA / UserStats.MaxSTA * 89
    frmMain.stabar.Caption = UserStats.MinSTA & " / " & UserStats.MaxSTA

    If QuePestañaInferior = 0 Then
        frmMain.STAShp.visible = (UserStats.MinSTA > 0)
    End If
End Sub

Public Sub UpdateGoldState()
    GldLbl.Caption = PonerPuntos(UserStats.GLD)
    'If UserGLD > UserLvl * OroPorNivel Then
    If UserStats.GLD <= 100000 Then
        GldLbl.ForeColor = vbRed
    Else
        GldLbl.ForeColor = &H80FFFF
    End If
End Sub

Public Sub UpdateExpBar()
    If UserStats.PasarNivel > 0 Then
        EXPBAR.Width = UserStats.exp / UserStats.PasarNivel * 235
        lblPorcLvl.Caption = Round(UserStats.exp * (100 / UserStats.PasarNivel), 2) & "%"
        exp.Caption = PonerPuntos(UserStats.exp) & "/" & PonerPuntos(UserStats.PasarNivel)
    Else
        EXPBAR.Width = 235
        lblPorcLvl.Caption = "¡Nivel máximo!"
        exp.Caption = "¡Nivel máximo!"
    End If
End Sub

Public Sub UpdateBuff()
    If UserStats.StrState = eHighBuff Then
        frmMain.Fuerzalbl.ForeColor = RGB(204, 0, 0)
    ElseIf UserStats.StrState = eMinBuff Then
        frmMain.Fuerzalbl.ForeColor = RGB(204, 100, 100)
    Else
        frmMain.Fuerzalbl.ForeColor = vbWhite

    End If
    
    If UserStats.AgiState = eHighBuff Then
        frmMain.AgilidadLbl.ForeColor = RGB(204, 0, 0)
    ElseIf UserStats.AgiState = eMinBuff Then
        frmMain.AgilidadLbl.ForeColor = RGB(204, 100, 100)
    Else
        frmMain.AgilidadLbl.ForeColor = vbWhite
    End If

    frmMain.Fuerzalbl.Caption = UserStats.StrState
    frmMain.AgilidadLbl.Caption = UserStats.AgiState
End Sub
