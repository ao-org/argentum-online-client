VERSION 5.00
Begin VB.Form frmShopAO20 
   BorderStyle     =   0  'None
   Caption         =   "Tienda AO20"
   ClientHeight    =   7125
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6480
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   475
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   432
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.ListBox lstItemShopFilter 
      Appearance      =   0  'Flat
      BackColor       =   &H80000001&
      ForeColor       =   &H80000000&
      Height          =   3150
      Left            =   360
      TabIndex        =   5
      Top             =   2400
      Width           =   3255
   End
   Begin VB.ListBox lstItemsShop 
      Appearance      =   0  'Flat
      BackColor       =   &H80000001&
      ForeColor       =   &H80000000&
      Height          =   3150
      Left            =   360
      TabIndex        =   3
      Top             =   2400
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.PictureBox PictureItemShop 
      Appearance      =   0  'Flat
      BackColor       =   &H80000001&
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   4770
      ScaleHeight     =   31
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   31
      TabIndex        =   2
      Top             =   2880
      Width           =   495
   End

   Begin VB.PictureBox picUserPreview 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000001&
      ForeColor       =   &H80000008&
      Height          =   1905
      Left            =   4095
      ScaleHeight     =   127
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   124
      TabIndex        =   6
      Top             =   3600
      Width           =   1860
   End
   Begin VB.TextBox txtFindObj 
      BackColor       =   &H80000007&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000000&
      Height          =   255
      Left            =   600
      TabIndex        =   0
      Top             =   1620
      Width           =   2415
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Una vez realizada la transacción, reloguee su personaje por seguridad"
      ForeColor       =   &H000000FF&
      Height          =   735
      Left            =   4095
      TabIndex        =   4
      Top             =   5160
      Width           =   1815
   End
   Begin VB.Image Image4 
      Height          =   375
      Left            =   6000
      Top             =   0
      Width           =   495
   End
   Begin VB.Image Image3 
      Height          =   375
      Left            =   3600
      Top             =   6360
      Width           =   1815
   End
   Begin VB.Image Image2 
      Height          =   375
      Left            =   1080
      Top             =   6360
      Width           =   1935
   End
   Begin VB.Image Image1 
      Height          =   375
      Left            =   3240
      Top             =   1560
      Width           =   375
   End
   Begin VB.Label lblCredits 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   240
      Left            =   5460
      TabIndex        =   1
      Top             =   1560
      Width           =   165
   End
End
Attribute VB_Name = "frmShopAO20"
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

Private Sub Form_Load()
    Me.Picture = LoadInterface("ventanatiendaao20.bmp")
    Label1.Caption = JsonLanguage.Item("MENSAJE_TRANSACCION_RELOGUEO")
    Call DrawUserPreview
End Sub

Private Sub Image2_Click()
    Unload Me
End Sub

Private Sub Image3_Click()
    'Antes de enviar al servidor hago una pre consulta de los créditos en cliente
    Dim obj_to_buy As ObjDatas
    Dim i          As Long
    obj_to_buy = ObjShop(Me.lstItemShopFilter.ListIndex + 1)
    Dim obj_name As String
    obj_name = Split(lstItemShopFilter.text, " (")(0)
    For i = 1 To UBound(ObjShop)
        If obj_name = ObjShop(i).Name Then
            obj_to_buy = ObjData(ObjShop(i).ObjNum)
            obj_to_buy.ObjNum = ObjShop(i).ObjNum
            obj_to_buy.Valor = ObjShop(i).Valor
            Exit For
        End If
    Next i
    If credits_shopAO20 >= obj_to_buy.Valor Then
        Call writeBuyShopItem(obj_to_buy.ObjNum)
    Else
        Call AddtoRichTextBox(frmMain.RecTxt, JsonLanguage.Item("MENSAJE_NO_TIENE_CREDITOS_SUFICIENTES"), 255, 0, 0, True)
    End If
End Sub

Private Sub Image4_Click()
    Unload Me
End Sub

Private Sub lstItemShopFilter_Click()
    Dim Grh      As Long
    Dim i        As Long
    Dim obj_name As String
    obj_name = Split(lstItemShopFilter.text, " (")(0)
    For i = 1 To UBound(ObjShop)
        If obj_name = ObjShop(i).Name Then
            Grh = ObjData(ObjShop(i).ObjNum).GrhIndex
            Exit For
        End If
    Next i
    Call Grh_Render_To_Hdc(PictureItemShop, Grh, 0, 0, False)
End Sub

Private Sub DrawUserPreview()
    On Error GoTo DrawUserPreview_Err

    Call picUserPreview.Cls

    Dim bodyIndex As Integer
    Dim headIndex As Integer

    bodyIndex = UserBody
    headIndex = UserHead

    If UserCharIndex >= LBound(charlist) And UserCharIndex <= UBound(charlist) Then
        With charlist(UserCharIndex)
            If .active <> 0 Then
                If .iBody <> 0 Then bodyIndex = .iBody
                If .IHead <> 0 Then headIndex = .IHead
            End If
        End With
    End If

    If bodyIndex < LBound(BodyData) Or bodyIndex > UBound(BodyData) Then Exit Sub

    If headIndex < LBound(HeadData) Or headIndex > UBound(HeadData) Then
        headIndex = 0
    End If

    Call DibujarNPC(picUserPreview, headIndex, bodyIndex, E_Heading.south)
    Call picUserPreview.Refresh

    Exit Sub

DrawUserPreview_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmShopAO20.DrawUserPreview", Erl)
End Sub

Private Sub txtFindObj_Change()
    lstItemShopFilter.Clear
    Dim i As Long
    For i = 1 To UBound(ObjShop)
        If InStr(1, ObjShop(i).Name, txtFindObj.text, 1) > 0 Then
            Call frmShopAO20.lstItemShopFilter.AddItem(ObjShop(i).Name & " ( " & JsonLanguage.Item("MENSAJE_VALOR") & ObjShop(i).Valor & " )")
        End If
    Next i
End Sub
