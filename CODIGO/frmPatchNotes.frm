VERSION 5.00
Begin VB.Form frmPatchNotes 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   9255
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11550
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   617
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   770
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Image cmdNext 
      Height          =   420
      Left            =   4560
      Tag             =   "0"
      Top             =   8640
      Width           =   1980
   End
   Begin VB.Image background 
      Height          =   9255
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   11535
   End
End
Attribute VB_Name = "frmPatchNotes"
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
'```
Private cNextButton As clsGraphicalButton


Public Sub SetNotes(ByRef notePath As String)
    Me.Picture = LoadInterface(notePath, False)
    MakeFormTransparent Me, vbBlack    'Set the Form "transparent by color."
End Sub

Private Sub cmdNext_Click()
    Unload Me
    FrmLogear.Show , frmConnect
End Sub

Private Sub Form_Load()
    Set cNextButton = New clsGraphicalButton
    Call cNextButton.Initialize(cmdNext, "boton-aceptar-default.bmp", "boton-aceptar-over.bmp", "boton-aceptar-off.bmp", Me)
End Sub

Private Sub Form_LostFocus()
    Unload Me
    FrmLogear.Show , frmConnect
End Sub
