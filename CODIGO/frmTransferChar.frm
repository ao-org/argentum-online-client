VERSION 5.00
Begin VB.Form frmTransferChar 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Character Transfer"
   ClientHeight    =   2340
   ClientLeft      =   40
   ClientTop       =   380
   ClientWidth     =   5080
   ControlBox      =   0   'False
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2340
   ScaleWidth      =   5080
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cancelButton 
      Caption         =   "Cancel"
      Height          =   492
      Left            =   360
      TabIndex        =   1
      Top             =   1320
      Width           =   972
   End
   Begin VB.TextBox textboxTransferEmail 
      Height          =   372
      Left            =   360
      TabIndex        =   0
      Top             =   600
      Width           =   4212
   End
   Begin VB.CommandButton transferButton 
      Caption         =   "Transfer!"
      Height          =   492
      Left            =   3360
      TabIndex        =   2
      Top             =   1320
      Width           =   972
   End
   Begin VB.Label labelTextbox 
      Caption         =   "Email account you wish to transfer the character"
      Height          =   372
      Left            =   360
      TabIndex        =   3
      Top             =   240
      Width           =   3972
   End
End
Attribute VB_Name = "frmTransferChar"
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
Option Explicit

Private Sub cancelButton_Click()
Unload Me
End Sub

Private Sub transferButton_Click()
TransferCharNewOwner = Me.textboxTransferEmail.Text
ModAuth.LoginOperation = e_operation.transfercharacter
Call connectToLoginServer
Unload Me
End Sub
