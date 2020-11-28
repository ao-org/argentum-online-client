VERSION 5.00
Begin VB.Form FrmCuenta 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "RevolucionAO 1.0"
   ClientHeight    =   11520
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   15345
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H0000FF00&
   Icon            =   "FrmCuenta.frx":0000
   KeyPreview      =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   768
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1023
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox rendercuenta 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   11520
      Left            =   0
      ScaleHeight     =   768
      ScaleMode       =   0  'User
      ScaleWidth      =   1024
      TabIndex        =   0
      Top             =   0
      Width           =   15360
      Begin VB.Timer cerrarform 
         Enabled         =   0   'False
         Interval        =   100
         Left            =   13200
         Top             =   120
      End
      Begin VB.Timer Timer1 
         Interval        =   100
         Left            =   840
         Top             =   9960
      End
   End
   Begin VB.Label lblclosed 
      BackStyle       =   0  'Transparent
      Caption         =   "x"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   15000
      TabIndex        =   1
      Top             =   0
      Width           =   255
   End
End
Attribute VB_Name = "FrmCuenta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
