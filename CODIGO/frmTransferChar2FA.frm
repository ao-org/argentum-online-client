VERSION 5.00
Begin VB.Form frmTransferChar2FA 
   Caption         =   "Form1"
   ClientHeight    =   2115
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5040
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "2FA"
   ScaleHeight     =   2115
   ScaleWidth      =   5040
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton btnConfirm 
      Caption         =   "Confirm"
      Height          =   360
      Left            =   3720
      TabIndex        =   1
      Top             =   1560
      Width           =   990
   End
   Begin VB.TextBox txt2FACode 
      Height          =   375
      Left            =   480
      TabIndex        =   0
      Top             =   720
      Width           =   4215
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "2FA Code"
      Height          =   195
      Left            =   480
      TabIndex        =   2
      Top             =   360
      Width           =   705
   End
End
Attribute VB_Name = "frmTransferChar2FA"
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

Option Explicit

Private Sub Form_Load()
    Me.Left = (screen.Width - Me.Width) \ 2
    Me.Top = (screen.Height - Me.Height) \ 2
    Me.txt2FACode.Text = ""
    Me.Caption = "2FA Code"
    Me.btnConfirm.Caption = JsonLanguage.Item("MENSAJE_ACEPTAR")
End Sub

Private Sub btnConfirm_Click()
    Me.txt2FACode.Text = Trim(Me.txt2FACode.Text)
    If Me.txt2FACode.Text <> "" Then
        ModAuth.LoginOperation = e_operation.ConfirmTransferChar
        Call connectToLoginServer
        transfer_char_validate_code = Me.txt2FACode.Text
        Call Unload(Me)
    Else
        Call MsgBox(JsonLanguage.Item("MENSAJEBOX_CODIGO_INVALIDO"), vbOKOnly, JsonLanguage.Item("MENSAJEBOX_ERROR"))
    End If
End Sub


