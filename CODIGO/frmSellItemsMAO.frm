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
      Left            =   2760
      TabIndex        =   0
      Top             =   840
      Width           =   3615
   End
   Begin VB.Label Label3 
      Caption         =   "Precio en ARS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   840
      TabIndex        =   3
      Top             =   840
      Width           =   1935
   End
   Begin VB.Label Label2 
      Caption         =   "publish"
      Height          =   255
      Left            =   4200
      TabIndex        =   2
      Top             =   1800
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "Costo de publiacion del item: 1000 ORO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   6735
   End
   Begin VB.Image imgPublishItemMao 
      Height          =   615
      Left            =   3480
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

Public WithEvents InvUser As clsGrapchicalInventory

Private Sub imgPublishItemMao_Click()
    If Val(txtPriceItemInMao.Text <= 0) Then
    ' here we need custom message for invalid item value
        ' Call MsgBox(JsonLanguage.Item("MENSAJE_VALOR_PERSONAJE_INVALIDO"), vbCritical, JsonLanguage.Item("MENSAJE_TITULO_ERROR"))
        Exit Sub
    End If
    
    If (frmMain.Inventario.SelectedItem > 0 And frmMain.Inventario.SelectedItem < MAX_INVENTORY_SLOTS + 1) Then
        Call writePublishItemMAO(Val(txtPriceItemInMao.Text), frmMain.Inventario.SelectedItem)
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
