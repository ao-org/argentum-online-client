VERSION 5.00
Begin VB.Form frmSellItemsMAO 
   Caption         =   "Form1"
   ClientHeight    =   2475
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7200
   LinkTopic       =   "Form1"
   ScaleHeight     =   2475
   ScaleWidth      =   7200
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtPriceItemInMao 
      Height          =   495
      Left            =   1800
      TabIndex        =   0
      Top             =   840
      Width           =   3615
   End
   Begin VB.Image imgPublishItemMao 
      Height          =   615
      Left            =   2520
      Top             =   1560
      Width           =   2055
   End
End
Attribute VB_Name = "frmSellItemsMAO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Argentum 20 Game Client
'
'    Copyright (C) 2023 Noland Studios LTD
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
'
'    You should have received a copy of the GNU Affero General Public License
'    along with this program.  If not, see <https://www.gnu.org/licenses/>.
'
'    This program was based on Argentum Online 0.11.6
'    Copyright (C) 2002 Márquez Pablo Ignacio
'
'    Argentum Online is based on Baronsoft's VB6 Online RPG
'    You can contact the original creator of ORE at aaron@baronsoft.com
'    for more information about ORE please visit http://www.baronsoft.com/
'
'
'

Private Sub imgPublishItemMao_Click()
    If Val(txtPriceItemInMao.Text <= 0) Then
    ' here we need custom message for invalid item value
        ' Call MsgBox(JsonLanguage.Item("MENSAJE_VALOR_PERSONAJE_INVALIDO"), vbCritical, JsonLanguage.Item("MENSAJE_TITULO_ERROR"))
        Exit Sub
    End If
    
    If MsgBox(JsonLanguage.Item("MENSAJE_PUBLICAR_PERSONAJE") & userName & JsonLanguage.Item("MENSAJE_PUBLICAR_PERSONAJE_VALOR") & txtValor.Text & JsonLanguage.Item("MENSAJE_PUBLICAR_PERSONAJE_COSTO"), vbYesNo + vbQuestion, JsonLanguage.Item("MENSAJE_TITULO_PUBLICAR_PERSONAJE")) = vbYes Then
        Call writePublicarPersonajeMAO(Val(txtValor.Text))
        Call closeForm
        Call WriteDeleteItem(frmMain.Inventario.SelectedItem)
    End If
End Sub
Private Sub closeForm()
    txtPriceItemInMao.Text = ""
    Unload Me
End Sub

Private Sub txtPriceItemInMao_Change()
    textval = txtPriceItemInMao.Text
    If IsNumeric(textval) Then
      numval = textval
    Else
      txtPriceItemInMao.Text = CStr(numval)
    End If
End Sub
