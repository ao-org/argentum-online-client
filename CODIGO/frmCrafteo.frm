VERSION 5.00
Begin VB.Form frmCrafteo 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   ClientHeight    =   5076
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   423
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   750
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton BtnAgregarCatalizador 
      Caption         =   "Agregar catalizador"
      Height          =   375
      Left            =   6600
      TabIndex        =   5
      Top             =   4440
      Width           =   1575
   End
   Begin VB.CommandButton BtnQuitarObjeto 
      Caption         =   "Quitar objeto"
      Height          =   375
      Left            =   720
      TabIndex        =   4
      Top             =   4440
      Width           =   1455
   End
   Begin VB.CommandButton BtnAgregarObjeto 
      Caption         =   "Agregar objeto"
      Height          =   375
      Left            =   4920
      TabIndex        =   3
      Top             =   4440
      Width           =   1455
   End
   Begin VB.CommandButton BtnSalir 
      Caption         =   "X"
      Height          =   375
      Left            =   8520
      Picture         =   "frmCrafteo.frx":0000
      TabIndex        =   2
      Top             =   0
      Width           =   375
   End
   Begin VB.CommandButton Craftear 
      Caption         =   "Craftear"
      Height          =   375
      Left            =   3240
      TabIndex        =   1
      Top             =   4440
      Width           =   1455
   End
   Begin VB.PictureBox PicInven 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   3675
      Left            =   360
      ScaleHeight     =   306
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   690
      TabIndex        =   0
      Top             =   600
      Width           =   8280
   End
End
Attribute VB_Name = "frmCrafteo"
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

Public WithEvents InvCraftUser As clsGrapchicalInventory
Attribute InvCraftUser.VB_VarHelpID = -1
Public WithEvents InvCraftItems As clsGrapchicalInventory
Attribute InvCraftItems.VB_VarHelpID = -1
Public WithEvents InvCraftCatalyst As clsGrapchicalInventory
Attribute InvCraftCatalyst.VB_VarHelpID = -1

Public InventoryGrhIndex As Long
Public TipoGrhIndex As Long
Public ResultGrhIndex As Long
Public PorcentajeAcierto As Byte
Public PrecioCrafteo As Long

Private Sub BtnAgregarCatalizador_Click()
    If InvCraftUser.SelectedItem > 0 Then
        Call WriteAddCatalyst(InvCraftUser.SelectedItem)
    End If
End Sub

Private Sub BtnAgregarObjeto_Click()
    If InvCraftUser.SelectedItem > 0 Then
        Call WriteAddItemCrafting(InvCraftUser.SelectedItem, 0)
    End If
End Sub

Private Sub BtnQuitarObjeto_Click()
    If InvCraftCatalyst.SelectedItem > 0 Then
        Call WriteRemoveCatalyst(0)
    ElseIf InvCraftItems.SelectedItem > 0 Then
        Call WriteRemoveItemCrafting(InvCraftItems.SelectedItem, 0)
    End If
End Sub

Private Sub BtnSalir_Click()
    Unload Me
End Sub

Private Sub Craftear_Click()
    Call WriteCraftItem
End Sub

Private Sub Form_Load()
    Call Aplicar_Transparencia(Me.hwnd, 240)
    Call FormParser.Parse_Form(Me)
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call MoverForm(Me.hwnd)
End Sub

Public Sub SetResult(ByVal GrhIndex As Long, ByVal Porcentaje As Byte, ByVal Precio As Long)
    ResultGrhIndex = GrhIndex
    PorcentajeAcierto = Porcentaje
    PrecioCrafteo = Precio
    ' Forzamos el redibujado
    InvCraftUser.ReDraw
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Comerciando Then
        Comerciando = False
        Call WriteCloseCrafting
    End If
End Sub

Private Sub InvCraftCatalyst_ItemDropped(ByVal Drag As Integer, ByVal Drop As Integer, ByVal x As Integer, ByVal y As Integer)
    ' Si soltó fuera del catalizador (drag a otro inventario)
    If Drop = 0 Then
        Drop = InvCraftUser.GetSlot(x, y)

        ' Si lo soltó dentro del inventario
        If Drop > 0 Then
            ' Si ya había un item en ese slot
            If InvCraftUser.OBJIndex(Drop) Then
                ' Y es distinto al que estamos devolviendo
                If InvCraftItems.OBJIndex(Drag) <> InvCraftUser.OBJIndex(Drop) Then
                    ' Lo tiramos en cualquier otro slot del inventario
                    Drop = 0
                End If
            End If

            ' Movemos el catalizador al inventario
            Call WriteRemoveCatalyst(Drop)
        End If

    End If
End Sub

Private Sub InvCraftItems_ItemDropped(ByVal Drag As Integer, ByVal Drop As Integer, ByVal x As Integer, ByVal y As Integer)
    ' Si soltó dentro del mismo inventario
    If Drop > 0 Then
        If Drag <> Drop Then
            ' Movemos el item dentro de los slots de crafteo
            Call WriteMoveCraftItem(Drag, Drop)
        End If
    Else
        Drop = InvCraftUser.GetSlot(x, y)

        ' Si lo soltó dentro del inventario
        If Drop > 0 Then
            ' Si ya había un item en ese slot
            If InvCraftUser.OBJIndex(Drop) Then
                ' Y es distinto al que estamos devolviendo
                If InvCraftItems.OBJIndex(Drag) <> InvCraftUser.OBJIndex(Drop) Then
                    ' Lo tiramos en cualquier otro slot del inventario
                    Drop = 0
                End If
            End If

            ' Sacamos el item al slot que indicó
            Call WriteRemoveItemCrafting(Drag, Drop)
        End If

    End If
End Sub

Private Sub InvCraftUser_ItemDropped(ByVal Drag As Integer, ByVal Drop As Integer, ByVal x As Integer, ByVal y As Integer)
    ' Si soltó dentro del mismo inventario
    If Drop > 0 Then
        If Drag <> Drop Then
            ' Movemos el item dentro del inventario
            Call WriteItemMove(Drag, Drop)
        End If
    Else
        Drop = InvCraftItems.GetSlot(x, y)

        ' Si lo soltó dentro de los slots de crafteo
        If Drop > 0 Then
            ' Si ya había un item en ese slot
            If InvCraftItems.OBJIndex(Drop) Then
                ' Lo devolvemos al inventario
                Call WriteRemoveItemCrafting(Drop, 0)
            End If

            ' Agregamos el item al slot de crafteo
            Call WriteAddItemCrafting(Drag, Drop)

        Else
            Drop = InvCraftCatalyst.GetSlot(x, y)
            
            ' Si lo soltó dentro del slot del catalizador
            If Drop > 0 Then
                ' Si ya había un item en ese slot
                If InvCraftCatalyst.OBJIndex(Drop) Then
                    ' Lo devolvemos al inventario
                    Call WriteRemoveCatalyst(0)
                End If
    
                ' Agregamos el item al slot del catalizador
                Call WriteAddCatalyst(Drag)
            End If
        End If

    End If
End Sub


