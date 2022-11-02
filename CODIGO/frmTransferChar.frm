VERSION 5.00
Begin VB.Form frmTransferChar 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Character Transfer"
   ClientHeight    =   2340
   ClientLeft      =   36
   ClientTop       =   384
   ClientWidth     =   5088
   ControlBox      =   0   'False
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2340
   ScaleWidth      =   5088
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cancelButton 
      Caption         =   "Cancel"
      Height          =   492
      Left            =   360
      TabIndex        =   3
      Top             =   1320
      Width           =   972
   End
   Begin VB.TextBox textboxTransferEmail 
      Height          =   372
      Left            =   360
      TabIndex        =   1
      Top             =   600
      Width           =   4212
   End
   Begin VB.CommandButton transferButton 
      Caption         =   "Transfer!"
      Height          =   492
      Left            =   3360
      TabIndex        =   0
      Top             =   1320
      Width           =   972
   End
   Begin VB.Label labelTextbox 
      Caption         =   "Email account you wish to transfer the character"
      Height          =   372
      Left            =   360
      TabIndex        =   2
      Top             =   240
      Width           =   3972
   End
End
Attribute VB_Name = "frmTransferChar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cancelButton_Click()
Unload Me
End Sub

Private Sub transferButton_Click()
ModAuth.LoginOperation = e_operation.transfercharacter
Call connectToLoginServer
Unload Me

End Sub
