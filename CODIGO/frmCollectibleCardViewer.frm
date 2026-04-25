VERSION 5.00
Begin VB.Form frmCollectibleCardViewer 
   Caption         =   "Form1"
   ClientHeight    =   9045
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11370
   LinkTopic       =   "Form1"
   ScaleHeight     =   9045
   ScaleWidth      =   11370
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Command2"
      Height          =   855
      Left            =   600
      TabIndex        =   2
      Top             =   8040
      Width           =   3375
   End
   Begin VB.CommandButton cmdAccept 
      Caption         =   "Command1"
      Height          =   855
      Left            =   7560
      TabIndex        =   1
      Top             =   8040
      Width           =   3495
   End
   Begin VB.PictureBox CollectibleCardPictureBox 
      Height          =   7815
      Left            =   0
      ScaleHeight     =   517
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   757
      TabIndex        =   0
      Top             =   0
      Width           =   11415
   End
End
Attribute VB_Name = "frmCollectibleCardViewer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAccept_Click()
    frmCollectibleCardViewer.visible = False
End Sub


Private Sub cmdCancel_Click()
    frmCollectibleCardViewer.visible = False
End Sub
