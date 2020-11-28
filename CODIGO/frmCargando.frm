VERSION 5.00
Begin VB.Form frmCargando 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "RevolucionAo 1.2"
   ClientHeight    =   11520
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   15435
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
   Icon            =   "frmCargando.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   768
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1029
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   1080
      Top             =   1320
   End
   Begin VB.PictureBox picLoad 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   11490
      Left            =   9240
      ScaleHeight     =   11490
      ScaleWidth      =   15495
      TabIndex        =   0
      Top             =   960
      Width           =   15495
      Begin VB.Timer Timer2 
         Enabled         =   0   'False
         Interval        =   15
         Left            =   1080
         Top             =   1800
      End
   End
End
Attribute VB_Name = "frmCargando"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Presentacion As Integer

Private Sub Form_Load()
Me.WindowState = vbNormal
Me.Picture = General_Load_Picture_From_Resource("cargandobn.bmp")
picLoad.Picture = General_Load_Picture_From_Resource("cargando.bmp")

End Sub

Private Sub Timer1_Timer()
Timer1.Enabled = False
frmConnect.Show
frmConnect.relampago.Enabled = True



Sound.Sound_Play CStr(SND_LLUVIAIN), True, 0, 0

Rem frmMain.Second.Enabled = True
General_Set_Connect
Unload frmCargando
Rem FrmLogear.Visible = True

FrmLogear.Show , frmConnect


End Sub

Private Sub Timer2_Timer()
Presentacion = Presentacion + 10
frmCargando.picLoad.Width = Presentacion
frmCargando.picLoad.Refresh
If Presentacion > 1200 Then
Timer2.Enabled = False
End If


End Sub
