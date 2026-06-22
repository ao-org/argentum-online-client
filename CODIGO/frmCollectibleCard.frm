VERSION 5.00
Begin VB.Form frmCollectibleCard 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Card Viewer"
   ClientHeight    =   8970
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6555
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8970
   ScaleWidth      =   6555
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picCollectibleCard 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   9000
      Left            =   0
      ScaleHeight     =   8940
      ScaleWidth      =   6525
      TabIndex        =   0
      Top             =   0
      Width           =   6585
      Begin VB.CommandButton cmdAccept 
         Caption         =   "Add to collection"
         Height          =   375
         Left            =   2160
         TabIndex        =   1
         Top             =   8520
         Width           =   2175
      End
   End
End
Attribute VB_Name = "frmCollectibleCard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAccept_Click()
    frmCollectibleCard.visible = False
    cmdAccept.visible = False
    Call WriteUseItem(G_LastSelectedSlot)
End Sub

Private Sub cmdCancel_Click()
    frmCollectibleCard.visible = False
    cmdAccept.visible = False
End Sub

Private Sub Form_Load()
    Me.Move (screen.Width - Me.Width) \ 2, (screen.Height - Me.Height) \ 2
    cmdAccept.Caption = JsonLanguage.Item("MSG_ADD_TO_ALBUM_COLLECTION")
    frmCollectibleCard.Caption = JsonLanguage.Item("CAPTION_CARD_VIEWER")
End Sub
