VERSION 5.00
Begin VB.Form frmSellItemsMAO 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   6630
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7200
   LinkTopic       =   "Form1"
   ScaleHeight     =   6630
   ScaleWidth      =   7200
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtQuantity 
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
      Left            =   5030
      TabIndex        =   2
      Text            =   "1"
      Top             =   2100
      Width           =   940
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
      Left            =   620
      ScaleHeight     =   245
      ScaleMode       =   0  'User
      ScaleWidth      =   210
      TabIndex        =   1
      Top             =   1630
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
      Height          =   300
      Left            =   4560
      TabIndex        =   0
      Text            =   "0"
      Top             =   2890
      Width           =   1935
   End
   Begin VB.Image cmdCancel 
      Appearance      =   0  'Flat
      Height          =   435
      Left            =   2640
      Top             =   5880
      Width           =   1935
   End
   Begin VB.Image cmdCerrar 
      Height          =   375
      Left            =   6795
      Top             =   0
      Width           =   375
   End
   Begin VB.Image cmdMore 
      Height          =   315
      Left            =   6140
      Tag             =   "1"
      Top             =   2040
      Width           =   315
   End
   Begin VB.Image cmdLess 
      Height          =   315
      Left            =   4540
      Tag             =   "1"
      Top             =   2040
      Width           =   315
   End
   Begin VB.Image imgPublishItemMao 
      Height          =   375
      Left            =   4560
      Top             =   3480
      Width           =   1935
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
Attribute InvUser.VB_VarHelpID = -1

Private Sub loadButtons()
       
    Set cBotonCerrar = New clsGraphicalButton
    Set cBotonMas = New clsGraphicalButton
    Set cBotonMenos = New clsGraphicalButton
                                                
    Call cBotonCerrar.Initialize(cmdCerrar, "boton-cerrar-default.bmp", _
                                                "boton-cerrar-over.bmp", _
                                                "boton-cerrar-off.bmp", Me)
                                                
    Call cBotonMas.Initialize(cmdMore, "boton-sm-mas-default.bmp", _
                                                "boton-sm-mas-over.bmp", _
                                                "boton-sm-mas-off.bmp", Me)
                                                
    Call cBotonMenos.Initialize(cmdLess, "boton-sm-menos-default.bmp", _
                                                "boton-sm-menos-over.bmp", _
                                                "boton-sm-menos-off.bmp", Me)
End Sub

Private Sub cmdCancel_Click()
    Unload Me
    Comerciando = False
End Sub

Private Sub txtQuantity_Change()
    
    On Error GoTo txtQuantity_Change_Err
    

    If Val(txtQuantity.Text) < 1 Then
        txtQuantity.Text = 1
    ElseIf Val(txtQuantity.Text) > 10000 Then
        txtQuantity.Text = 10000
    End If
    
    txtQuantity.SelStart = Len(txtQuantity.Text)
    
    InvUser.ReDraw

    
    Exit Sub

txtQuantity_Change_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmComerciar.txtQuantity_Change", Erl)
    Resume Next
    
End Sub

Private Sub txtQuantity_KeyPress(KeyAscii As Integer)
    
    On Error GoTo Form_KeyPress_Err
    

    If (KeyAscii = 27) Then
        Unload Me

    End If

    
    Exit Sub

Form_KeyPress_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmComerciar.Form_KeyPress", Erl)
    Resume Next
    
End Sub

Private Sub cmdCerrar_Click()
    Unload Me
    Comerciando = False
End Sub

Private Sub cmdMore_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If (Val(txtQuantity.Text) < 10001) Then
        txtQuantity.Text = str((Val(txtQuantity.Text) + 1))
    End If
End Sub

Private Sub cmdLess_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If (Val(txtQuantity.Text) > 1) Then
        txtQuantity.Text = str((Val(txtQuantity.Text) - 1))
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    
    On Error GoTo Form_KeyPress_Err
    

    If (KeyAscii = 27) Then
        Unload Me

    End If

    
    Exit Sub

Form_KeyPress_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmComerciar.Form_KeyPress", Erl)
    Resume Next
    
End Sub

Private Sub Form_Load()
    
    On Error GoTo Form_Load_Err
    
    Call FormParser.Parse_Form(Me)
    
    txtQuantity.BackColor = RGB(18, 19, 13)

    Me.Picture = LoadInterface("sell_items_mao_interface.bmp")
    
    Call loadButtons
    
    Exit Sub

Form_Load_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmComerciar.Form_Load", Erl)
    Resume Next
    
End Sub

Private Sub imgPublishItemMao_Click()
    Dim arsInputvalue As Double
    arsInputvalue = Val(txtPriceItemInMao.Text)
    
    Dim itemQuantity As Integer
    itemQuantity = Val(txtQuantity.Text)
    
    If Not frmSellItemsMAO.InvUser.IsItemSelected Then
        Call AddtoRichTextBox(frmMain.RecTxt, JsonLanguage.Item("MENSAJE_NO_TIENE_ITEM_SELECCIONADO"), 255, 255, 255, False, False, False)
        Exit Sub
    ElseIf arsInputvalue < 1 Or arsInputvalue > 2147483647# Then
        Call AddtoRichTextBox(frmMain.RecTxt, JsonLanguage.Item("MENSAJE_VALOR_INTRODUCIDO_INVALIDO"), 255, 0, 32, False, False, False)
        Exit Sub
    End If
    
    If (frmSellItemsMAO.InvUser.SelectedItem > 0 And frmSellItemsMAO.InvUser.SelectedItem < MAX_INVENTORY_SLOTS + 1) Then
        Call writePublishItemMAO(arsInputvalue, frmSellItemsMAO.InvUser.SelectedItem, itemQuantity)
        Call closeForm
    End If
End Sub
Private Sub closeForm()
    txtPriceItemInMao.Text = ""
    Unload Me
    Comerciando = False
End Sub

Private Sub picInv_Paint()
    Call frmSellItemsMAO.InvUser.ReDraw
End Sub

Private Sub txtPriceItemInMao_Change()

    Dim arsInputvalue As Double
    Dim clampedValue As Long

    If Not IsNumeric(txtPriceItemInMao.Text) Then
        clampedValue = 1
        txtPriceItemInMao.Text = "1"
    Else
        arsInputvalue = Val(txtPriceItemInMao.Text)

        If arsInputvalue > 2147483647# Then
            clampedValue = 2147483647
            txtPriceItemInMao.Text = "2147483647"
        ElseIf arsInputvalue < 1 Then
            clampedValue = 1
            txtPriceItemInMao.Text = "1"
        Else
            clampedValue = CLng(arsInputvalue)
        End If
    End If
    InvUser.ReDraw
End Sub

Private Sub txtPriceItemInMao_KeyPress(KeyAscii As Integer)
    
    On Error GoTo Form_KeyPress_Err
    

    If (KeyAscii = 27) Then
        Unload Me

    End If

    
    Exit Sub

Form_KeyPress_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmComerciar.Form_KeyPress", Erl)
    Resume Next
    
End Sub
