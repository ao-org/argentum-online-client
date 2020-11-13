VERSION 5.00
Begin VB.Form FrmKeyInv 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   2820
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3600
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
   Picture         =   "FrmKeyInv.frx":0000
   ScaleHeight     =   2820
   ScaleWidth      =   3600
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox interface 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      ClipControls    =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   920
      Left            =   600
      MousePointer    =   99  'Custom
      Picture         =   "FrmKeyInv.frx":216C8
      ScaleHeight     =   61
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   155
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   1440
      Width           =   2325
   End
   Begin VB.Image cmdCerrar 
      Height          =   375
      Left            =   3240
      Top             =   0
      Width           =   375
   End
End
Attribute VB_Name = "FrmKeyInv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public WithEvents InvKeys As clsGrapchicalInventory
Attribute InvKeys.VB_VarHelpID = -1

Private Sub cmdCerrar_Click()
Unload Me
End Sub
