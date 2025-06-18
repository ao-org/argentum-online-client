VERSION 5.00
Begin VB.Form frmSellItemsMAO 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   6300
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3435
   LinkTopic       =   "Form1"
   ScaleHeight     =   6300
   ScaleWidth      =   3435
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox cantidad 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   620
      TabIndex        =   2
      Text            =   "1"
      Top             =   5835
      Width           =   810
   End
   Begin VB.PictureBox picInv 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   3675
      Left            =   120
      ScaleHeight     =   245
      ScaleMode       =   0  'User
      ScaleWidth      =   210
      TabIndex        =   1
      Top             =   1680
      Width           =   3150
   End
   Begin VB.TextBox txtPriceItemInMao 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000001&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   325
      Left            =   120
      TabIndex        =   0
      Text            =   "0"
      Top             =   1120
      Width           =   3150
   End
   Begin VB.Image cmdCerrar 
      Height          =   375
      Left            =   2950
      Top             =   0
      Width           =   495
   End
   Begin VB.Image cmdMas 
      Height          =   315
      Left            =   1650
      Tag             =   "1"
      Top             =   5760
      Width           =   315
   End
   Begin VB.Image cmdMenos 
      Height          =   315
      Left            =   120
      Tag             =   "1"
      Top             =   5760
      Width           =   315
   End
   Begin VB.Image imgPublishItemMao 
      Height          =   615
      Left            =   2040
      Top             =   5520
      Width           =   1215
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
Private quantity As Integer

Public WithEvents InvUser As clsGrapchicalInventory
Attribute InvUser.VB_VarHelpID = -1

Private Sub loadButtons()
       
    Set cBotonCerrar = New clsGraphicalButton
    Set cBotonMas = New clsGraphicalButton
    Set cBotonMenos = New clsGraphicalButton
                                                
    Call cBotonCerrar.Initialize(cmdCerrar, "boton-cerrar-default.bmp", _
                                                "boton-cerrar-over.bmp", _
                                                "boton-cerrar-off.bmp", Me)
                                                
    Call cBotonMas.Initialize(cmdMas, "boton-sm-mas-default.bmp", _
                                                "boton-sm-mas-over.bmp", _
                                                "boton-sm-mas-off.bmp", Me)
                                                
    Call cBotonMenos.Initialize(cmdMenos, "boton-sm-menos-default.bmp", _
                                                "boton-sm-menos-over.bmp", _
                                                "boton-sm-menos-off.bmp", Me)
End Sub

Private Sub cantidad_Change()
    
    On Error GoTo cantidad_Change_Err
    

    If Val(cantidad.Text) < 1 Then
        cantidad.Text = 1
        quantity = 1
    ElseIf Val(cantidad.Text) > 10000 Then
        cantidad.Text = 10000
        quantity = 10000
    Else
        quantity = Val(cantidad.Text)
    End If
    
    cantidad.SelStart = Len(cantidad.Text)
    
    InvUser.ReDraw

    
    Exit Sub

cantidad_Change_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmComerciar.cantidad_Change", Erl)
    Resume Next
    
End Sub

Private Sub cmdCerrar_Click()
    Unload Me
End Sub

Private Sub cmdMas_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If (Val(cantidad.Text) < 10001) Then
        cantidad.Text = str((Val(cantidad.Text) + 1))
        quantity = quantity + 1
    End If
End Sub

Private Sub cmdMenos_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If (Val(cantidad.Text) > 1) Then
        cantidad.Text = str((Val(cantidad.Text) - 1))
        quantity = quantity - 1
    End If
End Sub

Private Sub Form_Load()
    
    On Error GoTo Form_Load_Err
    
    Call FormParser.Parse_Form(Me)
    cantidad.BackColor = RGB(18, 19, 13)

'we need a new picture for this
    Me.Picture = LoadInterface("sell_items_mao_interface.bmp")
    Call loadButtons
    Exit Sub

Form_Load_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmComerciar.Form_Load", Erl)
    Resume Next
    
End Sub

Private Sub imgPublishItemMao_Click()
    If Val(txtPriceItemInMao.Text <= 0) Then
    ' here we need custom message for invalid item value
        ' Call MsgBox(JsonLanguage.Item("MENSAJE_VALOR_PERSONAJE_INVALIDO"), vbCritical, JsonLanguage.Item("MENSAJE_TITULO_ERROR"))
        Exit Sub
    End If
    
    If (frmSellItemsMAO.InvUser.SelectedItem > 0 And frmSellItemsMAO.InvUser.SelectedItem < MAX_INVENTORY_SLOTS + 1) Then
        Call writePublishItemMAO(Val(txtPriceItemInMao.Text), frmSellItemsMAO.InvUser.SelectedItem)
        Call closeForm
        Call WriteDeleteItem(frmSellItemsMAO.InvUser.SelectedItem)
        Call frmSellItemsMAO.InvUser.ReDraw
    End If
End Sub
Private Sub closeForm()
    txtPriceItemInMao.Text = ""
    Unload Me
End Sub

Private Sub picInv_Paint()
    Call frmSellItemsMAO.InvUser.ReDraw
End Sub

Private Sub txtPriceItemInMao_Change()

    Dim inputValue As Double
    Dim clampedValue As Long

    If Not IsNumeric(txtPriceItemInMao.Text) Then
        clampedValue = 1
        txtPriceItemInMao.Text = "1"
    Else
        inputValue = Val(txtPriceItemInMao.Text)

        If inputValue > 2147483647# Then
            clampedValue = 2147483647
            txtPriceItemInMao.Text = "2147483647"
        ElseIf inputValue < 1 Then
            clampedValue = 1
            txtPriceItemInMao.Text = "1"
        Else
            clampedValue = CLng(inputValue)
        End If
    End If
End Sub
