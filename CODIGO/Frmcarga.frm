VERSION 5.00
Begin VB.Form Frmcarga 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Cargando..."
   ClientHeight    =   1596
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2412
   LinkTopic       =   "Form1"
   ScaleHeight     =   1596
   ScaleWidth      =   2412
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "Frmcarga"
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
Private Sub Form_Load()
    
    On Error GoTo Form_Load_Err
    
    Me.Picture = LoadInterface("VentanaCargando.bmp")
    MakeFormTransparent Me, vbBlack

    Exit Sub

Form_Load_Err:
    Call RegistrarError(Err.number, Err.Description, "Frmcarga.Form_Load", Erl)
    Resume Next
    
End Sub
