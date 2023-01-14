VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsGrapchicalInventory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
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
'*****************************************************
'Graphical inventory for ORE 0.4.1
'Designed and implemented by Juan Martín Sotuyo Dodero (Maraxus)
'juansotuyo@hotmail.com
'*****************************************************

'**************************************************************************
'This program is free software; you can redistribute it and/or modify
'it under the terms of the GNU General Public License as published by
'the Free Software Foundation; either version 2 of the License, or
'(at your option) any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'Affero General Public License for more details.
'
'You should have received a copy of the GNU General Public License
'along with this program; if not, write to the Free Software
'Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA
'**************************************************************************

Option Explicit

Private isInventory            As Boolean ' Para separar al inventario del main, de los otros

Dim InvMaxSlots                As Integer

Dim InvWidth                   As Integer

Dim InvHeight                  As Integer

Dim InvOffsetX                 As Integer

Dim InvOffsetY                 As Integer

Dim InvSelectedItem            As Long             'Currently selected item

Dim TempItem                   As Long

Dim ShowQuantities             As Boolean

Private MAININVENTORY_MIDSPACE As Byte

' Slot item arrastrado (0 si no está arrastrando)
Dim Dragging                   As Integer

' Posición del item arrastrado
Dim XDrag                      As Integer

Dim YDrag                      As Integer

' Flag para saber si se cliqueó dentro del inventario
' al hacer clic sobre el Picture
Dim wasClickedInside           As Boolean

' Flag para saber si hay que volver a dibujar
Dim InvNeedsRedraw             As Boolean

'Dim InvSurface As DirectDrawSurface7            'DD Surface used to render everything

Dim UserInventory()            As Inventory    'User's inventory

Dim WithEvents InventoryWindow As PictureBox    'Placeholder where to render the inventory
Attribute InventoryWindow.VB_VarHelpID = -1

' Evento disparado cuando un item se soltó
Public Event ItemDropped(ByVal Drag As Integer, ByVal Drop As Integer, ByVal x As Integer, ByVal y As Integer)

#If ConMenuesConextuales = 1 Then

    Dim ItemMenu As Menu    'Menu to be shown as pop up

#End If

' Valores constantes que no ameritan que los reciba como parámetro
Const ITEM_SIZE = 32 ' El ancho de los items (nunca debería cambiar

Const LETRA_GRANDE_Y = 20 ' Offset Y de cantidad de items (letra grande)

Const EQUIPPED_X = 22 ' Offset X del caracter de equipado

Const EQUIPPED_Y = -2 ' Offset Y del caracter de equipado

Const EQUIPPED_CHARACTER = "+" ' Caracter de item equipado

Const SELECTION_BOX_W = 34 ' Ancho de la cajita de selección (se centra automáticamente)

Const SELECTION_BOX_H = 34 ' Alto de la cajita de selección (se centra automáticamente)

'Private Const MAININVENTORY_MIDSPACE As Integer = 0 '9 'Separador de inventario básico y extra slots en el inventario principal del usuario

' Valores calculados
Const HALF_ITEM_SIZE = ITEM_SIZE \ 2

Const SELECTION_BOX_X = (ITEM_SIZE - SELECTION_BOX_W) \ 2

Const SELECTION_BOX_Y = (ITEM_SIZE - SELECTION_BOX_H) \ 2

' Valores a calcular cuando se cree el inventario
Dim ITEM_OFFSET_X         As Integer

Dim ITEM_OFFSET_Y         As Integer

Dim ITEMS_PER_ROW         As Integer

Dim HALF_BORDER_X         As Integer

Dim HALF_BORDER_Y         As Integer

Dim HEIGHT_MAIN_INVENTORY As Integer

Private Sub Class_Initialize()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modify Date: 27/07/04
    '
    '***************************************************
End Sub

Private Sub Class_Terminate()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modify Date: 27/07/04
    '
    '***************************************************
End Sub

#If ConMenuesConextuales = 1 Then   'Argument list varies!!!
    Public Sub Initialize(ByRef DirectDraw As DirectDraw7, ByRef InvPic As PictureBox, ByRef PopUpMenu As Menu, Optional ByVal Width As Integer = 0, Optional ByVal Height As Integer = 0, Optional ByVal OffsetX As Integer = 0, Optional ByVal OffsetY As Integer = 0, Optional ByVal SpaceX As Integer = 0, Optional ByVal SpaceY As Integer = 0, Optional thisInventory As Boolean = False, Optional MIDSPACE As Byte = 0)

    End Sub

#Else   'Argument list varies!!!
    Public Sub Initialize(ByRef InvPic As PictureBox, ByVal MaxSlots As Integer, _
                        Optional ByVal Width As Integer = 0, Optional ByVal Height As Integer = 0, Optional ByVal OffsetX As Integer = 0, Optional ByVal OffsetY As Integer = 0, _
                        Optional ByVal SpaceX As Integer = 0, Optional ByVal SpaceY As Integer = 0, Optional thisInventory As Boolean = False, Optional MIDSPACE As Byte = 0)
#End If
    '***************************************************
    'Sets the reference to the picture box on which to render the inventory
    '***************************************************
    Set InventoryWindow = InvPic
    
    isInventory = thisInventory
    
    If Width = 0 Then Width = InvPic.ScaleWidth
    If Height = 0 Then Height = InvPic.ScaleHeight

    InvMaxSlots = MaxSlots
    InvWidth = Width
    InvHeight = Height
    InvOffsetX = OffsetX
    InvOffsetY = OffsetY
    
    ITEM_OFFSET_X = ITEM_SIZE + SpaceX ' Elon Musk, is that u?
    ITEM_OFFSET_Y = ITEM_SIZE + SpaceY
    ITEMS_PER_ROW = InvWidth \ ITEM_OFFSET_X
    
    HALF_BORDER_X = SpaceX \ 2
    HALF_BORDER_Y = SpaceY \ 2
    
    MAININVENTORY_MIDSPACE = MIDSPACE
    
    ShowQuantities = True
    
    If isInventory Then
        Dim aux As Single
        aux = MAX_NORMAL_INVENTORY_SLOTS / ITEMS_PER_ROW
        HEIGHT_MAIN_INVENTORY = IIf(aux <> Int(aux), Int(aux) + 1, aux) * ITEM_OFFSET_Y
    End If

    ReDim UserInventory(1 To MaxSlots) As Inventory

#If ConMenuesConextuales = 1 Then
    Set ItemMenu = PopUpMenu
#End If
End Sub

Public Sub SelectGold()
    
    On Error GoTo SelectGold_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modify Date: 27/07/04
    'Sets the gold as the currently selected item
    '***************************************************
    Dim prevSelectedItem As Long
    
    'Store preivously selected item
    prevSelectedItem = InvSelectedItem
    
    'Select the gold
    InvSelectedItem = FLAGORO
    
    
    Exit Sub

SelectGold_Err:
    Call RegistrarError(Err.Number, Err.Description, "clsGrapchicalInventory.SelectGold", Erl)
    Resume Next
    
End Sub

Public Property Get SelectedItem() As Long
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modify Date: 27/07/04
    'Retrieves the selected item index
    '***************************************************
    
    On Error GoTo SelectedItem_Err
    
    SelectedItem = InvSelectedItem

    
    Exit Property

SelectedItem_Err:
    Call RegistrarError(Err.Number, Err.Description, "clsGrapchicalInventory.SelectedItem", Erl)
    Resume Next
    
End Property

Public Property Get IsItemSelected() As Boolean
    
    On Error GoTo IsItemSelected_Err
    
    If SelectedItem <= 0 Or SelectedItem > MaxSlots Then Exit Sub
    IsItemSelected = (UserInventory(SelectedItem).GrhIndex > 0)
    
    Exit Property

IsItemSelected_Err:
    Call RegistrarError(Err.Number, Err.Description, "clsGrapchicalInventory.IsItemSelected", Erl)
    Resume Next
    
End Property

Public Property Get MaxHit(ByVal Slot As Byte) As Integer
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modify Date: 27/07/04
    'Retrieves the max hit of the selected item
    '***************************************************
    
    On Error GoTo MaxHit_Err
    
    MaxHit = UserInventory(Slot).MaxHit

    
    Exit Property

MaxHit_Err:
    Call RegistrarError(Err.Number, Err.Description, "clsGrapchicalInventory.MaxHit", Erl)
    Resume Next
    
End Property

Public Property Get MinHit(ByVal Slot As Byte) As Integer
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modify Date: 27/07/04
    'Retrieves the min hit of the selected item
    '***************************************************
    
    On Error GoTo MinHit_Err
    
    MinHit = UserInventory(Slot).MinHit

    
    Exit Property

MinHit_Err:
    Call RegistrarError(Err.Number, Err.Description, "clsGrapchicalInventory.MinHit", Erl)
    Resume Next
    
End Property

Public Property Get Def(ByVal Slot As Byte) As Integer
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modify Date: 27/07/04
    'Retrieves the defense of the selected item
    '***************************************************
    
    On Error GoTo Def_Err
    
    Def = UserInventory(Slot).Def

    
    Exit Property

Def_Err:
    Call RegistrarError(Err.Number, Err.Description, "clsGrapchicalInventory.Def", Erl)
    Resume Next
    
End Property

Public Property Get GrhIndex(ByVal Slot As Byte) As Long
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modify Date: 27/07/04
    'Retrieves the grh index of the selected item
    '***************************************************
    
    On Error GoTo GrhIndex_Err
    
    GrhIndex = UserInventory(Slot).GrhIndex

    
    Exit Property

GrhIndex_Err:
    Call RegistrarError(Err.Number, Err.Description, "clsGrapchicalInventory.GrhIndex", Erl)
    Resume Next
    
End Property

Public Property Get Valor(ByVal Slot As Byte) As Single
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modify Date: 27/07/04
    'Retrieves the value of the selected item
    '***************************************************
    
    On Error GoTo Valor_Err
    
    Valor = UserInventory(Slot).Valor

    
    Exit Property

Valor_Err:
    Call RegistrarError(Err.Number, Err.Description, "clsGrapchicalInventory.Valor", Erl)
    Resume Next
    
End Property

Public Property Get Amount(ByVal Slot As Byte) As Long
    
    On Error GoTo Amount_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modify Date: 27/07/04
    'Retrieves the selected item's amount
    '***************************************************
    If Slot = FLAGORO Then
        Amount = UserGLD
    ElseIf Slot >= LBound(UserInventory) And Slot <= UBound(UserInventory) Then
        Amount = UserInventory(Slot).Amount

    End If

    
    Exit Property

Amount_Err:
    Call RegistrarError(Err.Number, Err.Description, "clsGrapchicalInventory.Amount", Erl)
    Resume Next
    
End Property

Public Property Get OBJIndex(ByVal Slot As Byte) As Integer
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modify Date: 27/07/04
    'Retrieves the selected item's object index
    '***************************************************
    
    On Error GoTo OBJIndex_Err
    
    OBJIndex = UserInventory(Slot).OBJIndex

    
    Exit Property

OBJIndex_Err:
    Call RegistrarError(Err.Number, Err.Description, "clsGrapchicalInventory.OBJIndex", Erl)
    Resume Next
    
End Property

Public Property Get ObjType(ByVal Slot As Byte) As Integer
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modify Date: 27/07/04
    'Retrieves the selected item's object type
    '***************************************************
    
    On Error GoTo ObjType_Err
    
    ObjType = UserInventory(Slot).ObjType

    
    Exit Property

ObjType_Err:
    Call RegistrarError(Err.Number, Err.Description, "clsGrapchicalInventory.ObjType", Erl)
    Resume Next
    
End Property

Public Property Get ItemName(ByVal Slot As Byte) As String
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modify Date: 27/07/04
    'Retrieves the selected item's name
    '***************************************************
    
    On Error GoTo ItemName_Err
    
    ItemName = UserInventory(Slot).Name

    
    Exit Property

ItemName_Err:
    Call RegistrarError(Err.Number, Err.Description, "clsGrapchicalInventory.ItemName", Erl)
    Resume Next
    
End Property

Public Property Get Equipped(ByVal Slot As Byte) As Boolean
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modify Date: 27/07/04
    'Retrieves True if the item at the given pos is eqiupped
    '***************************************************
    
    On Error GoTo Equipped_Err
    
    Equipped = UserInventory(Slot).Equipped

    
    Exit Property

Equipped_Err:
    Call RegistrarError(Err.Number, Err.Description, "clsGrapchicalInventory.Equipped", Erl)
    Resume Next
    
End Property

Public Property Get PuedeUsar(ByVal Slot As Byte) As Byte
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modify Date: 27/07/04
    'Retrieves True if the item at the given pos is eqiupped
    '***************************************************
    
    On Error GoTo PuedeUsar_Err
    
    PuedeUsar = UserInventory(Slot).PuedeUsar

    
    Exit Property

PuedeUsar_Err:
    Call RegistrarError(Err.Number, Err.Description, "clsGrapchicalInventory.PuedeUsar", Erl)
    Resume Next
    
End Property

Public Property Get MaxSlots() As Integer
    
    On Error GoTo MaxSlots_Err
    
    MaxSlots = InvMaxSlots

    
    Exit Property

MaxSlots_Err:
    Call RegistrarError(Err.Number, Err.Description, "clsGrapchicalInventory.MaxSlots", Erl)
    Resume Next
    
End Property

Public Property Get Columns() As Integer
    
    On Error GoTo Columns_Err
    
    Columns = ITEMS_PER_ROW

    
    Exit Property

Columns_Err:
    Call RegistrarError(Err.Number, Err.Description, "clsGrapchicalInventory.Columns", Erl)
    Resume Next
    
End Property

Public Property Get SlotDrag() As Byte
    
    On Error GoTo SlotDrag_Err
    
    SlotDrag = Dragging

    
    Exit Property

SlotDrag_Err:
    Call RegistrarError(Err.Number, Err.Description, "clsGrapchicalInventory.SlotDrag", Erl)
    Resume Next
    
End Property

Public Property Get NeedsRedraw() As Boolean
    
    On Error GoTo NeedsRedraw_Err
    
    NeedsRedraw = InvNeedsRedraw

    
    Exit Property

NeedsRedraw_Err:
    Call RegistrarError(Err.Number, Err.Description, "clsGrapchicalInventory.NeedsRedraw", Erl)
    Resume Next
    
End Property

Public Sub ReDraw()
    ' Provoca que el inventario se vuelta a dibujar el próximo frame
    
    On Error GoTo ReDraw_Err
    
    InvNeedsRedraw = True

    
    Exit Sub

ReDraw_Err:
    Call RegistrarError(Err.Number, Err.Description, "clsGrapchicalInventory.ReDraw", Erl)
    Resume Next
    
End Sub

Public Sub SetItem(ByVal Slot As Byte, ByVal eOBJIndex As Integer, ByVal eAmount As Integer, ByVal eEquipped As Byte, ByVal eGrhIndex As Long, ByVal eObjType As Integer, ByVal eMaxHit As Integer, ByVal eMinHit As Integer, ByVal eDef As Integer, ByVal eValor As Single, ByVal eName As String, ByVal PuedeUsar As Byte)
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modify Date: 12/04/06
    'Sets all data for a given inventory slot
    '***************************************************
    
    On Error GoTo SetItem_Err
    

    If Slot < 1 Or Slot > InvMaxSlots Then Exit Sub

    With UserInventory(Slot)
        .Amount = eAmount
        .Def = eDef
        .Equipped = eEquipped
        .GrhIndex = eGrhIndex
        .MaxHit = eMaxHit
        .MinHit = eMinHit
        .Name = eName
        .OBJIndex = eOBJIndex
        .ObjType = eObjType
        .Valor = eValor
        .PuedeUsar = PuedeUsar

    End With

    InvNeedsRedraw = True

    
    Exit Sub

SetItem_Err:
    Call RegistrarError(Err.Number, Err.Description, "clsGrapchicalInventory.SetItem", Erl)
    Resume Next
    
End Sub

Public Sub ClearSlot(ByVal Slot As Byte)
    If Slot < 1 Or Slot > InvMaxSlots Then Exit Sub

    With UserInventory(Slot)
        .Amount = 0
        .Def = 0
        .Equipped = 0
        .GrhIndex = 0
        .MaxHit = 0
        .MinHit = 0
        .Name = vbNullString
        .OBJIndex = 0
        .ObjType = 0
        .Valor = 0
        .PuedeUsar = 0

    End With

    InvNeedsRedraw = True
End Sub

Public Sub SeleccionarItem(ByVal SelectedItem As Byte)
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modify Date: 27/07/04
    'Retrieves the selected item index
    '***************************************************
    
    On Error GoTo SeleccionarItem_Err
    
    InvSelectedItem = SelectedItem
    InvNeedsRedraw = True

    
    Exit Sub

SeleccionarItem_Err:
    Call RegistrarError(Err.Number, Err.Description, "clsGrapchicalInventory.SeleccionarItem", Erl)
    Resume Next
    
End Sub

Public Property Get ClickedInside() As Boolean
    
    On Error GoTo ClickedInside_Err
    
    ClickedInside = wasClickedInside

    
    Exit Property

ClickedInside_Err:
    Call RegistrarError(Err.Number, Err.Description, "clsGrapchicalInventory.ClickedInside", Erl)
    Resume Next
    
End Property

Public Function GetSlot(ByVal x As Integer, ByVal y As Integer) As Byte
    ' Restamos el offset del inventario
    ' (más la mitad del borde, para hacer que los clics en los bordes sean como en el slot)
    
    On Error GoTo GetSlot_Err
    
        
    x = x - InvOffsetX + HALF_BORDER_X
    y = y - InvOffsetY + HALF_BORDER_Y
    
    ' Separación entre inventario principal y slots de mochilas
    If isInventory Then

        ' Si está por debajo del inventari normal
        If y >= HEIGHT_MAIN_INVENTORY Then

            ' Si está en el espacio entre medio, salimos
            If y < HEIGHT_MAIN_INVENTORY + MAININVENTORY_MIDSPACE Then Exit Function
            ' Restamos el offset para los items de abajo
            y = y - MAININVENTORY_MIDSPACE

        End If

    End If
    
    ' ** Cambiado: queda mejor si también se pueden cliquear los bordes
    ' Nos aseguramos que no cliqueó en el borde de un slot
    'If x Mod ITEM_OFFSET_X >= ITEM_SIZE Then Exit Function
    'If y Mod ITEM_OFFSET_Y >= ITEM_SIZE Then Exit Function
    
    ' Dividimos para obtener la coordenada (x, y) del slot
    x = x \ ITEM_OFFSET_X
    y = y \ ITEM_OFFSET_Y
    
    ' Nos aseguramos que no cliqueó muy a izquierda ni muy a la derecha
    If x < 0 Or x >= ITEMS_PER_ROW Then Exit Function
    
    ' Nos aseguramos que no cliqueó muy arriba (después chequeamos por abajo)
    If y < 0 Then Exit Function

    ' Calculamos el índice del slot
    GetSlot = x + y * ITEMS_PER_ROW + 1
    
    ' Nos aseguramos de no pasarnos del límite de slots
    ' (por ej. si hay 24 slots y la grilla es 5x5, al cliquear en el último "slot" este no existe;
    ' o bien, si el Picture es más grande que el inventario)
    If GetSlot > InvMaxSlots Then GetSlot = 0

    
    Exit Function

GetSlot_Err:
    Call RegistrarError(Err.Number, Err.Description, "clsGrapchicalInventory.GetSlot", Erl)
    Resume Next
    
End Function

Public Sub DrawInventory()
    
    On Error GoTo DrawInventory_Err
    

    Dim i         As Byte

    Dim x         As Integer

    Dim y         As Integer
    
    Dim tmpString As String

    Dim Color(3)  As RGBA
    
    InvNeedsRedraw = False

    For i = 1 To UBound(UserInventory)
        x = ((i - 1) Mod ITEMS_PER_ROW) * ITEM_OFFSET_X + InvOffsetX
        y = ((i - 1) \ ITEMS_PER_ROW) * ITEM_OFFSET_Y + InvOffsetY
        
        If isInventory Then
            If (i > MAX_NORMAL_INVENTORY_SLOTS) Then
                y = y + MAININVENTORY_MIDSPACE

                If (i - MAX_NORMAL_INVENTORY_SLOTS) <= (UserInvUnlocked * (InvWidth / ITEM_OFFSET_X)) Then
                    'Slot Desbloqueado
                    Call Draw_GrhIndex(GRH_INVENTORYSLOT_EXTRA, x, y)
                Else
                    'Slot Bloqueado
                    Call Draw_GrhIndex(GRH_INVENTORYSLOT_LOCKED, x, y)

                End If

            Else
                'Slot Default
                Call Draw_GrhIndex(GRH_INVENTORYSLOT, x, y)
            End If
        End If

        If InvSelectedItem = i Then
            Engine_Draw_Box x + SELECTION_BOX_X, y + SELECTION_BOX_Y, SELECTION_BOX_W, SELECTION_BOX_H, RGBA_From_Comp(100, 100, 100, 100)
        End If
            
        If UserInventory(i).GrhIndex Then
            If UserInventory(i).PuedeUsar > 0 Then
                Call RGBAList(Color, 255, 100, 100)
                
                Call Draw_GrhColor(UserInventory(i).GrhIndex, x, y, Color)

            Else
                Call Draw_GrhColor(UserInventory(i).GrhIndex, x, y, COLOR_WHITE())
               
            End If
            
            If ShowQuantities Then
                If UserInventory(i).Amount > 1 Then
                    Call RGBAList(Color, 239, 228, 176)
        
                    If InvSelectedItem <> i And UserInventory(i).Amount > 1000 And NumerosCompletosInventario = 0 Then
                        tmpString = Round(UserInventory(i).Amount * 0.001, 1) & "K"
                    Else
                        tmpString = UserInventory(i).Amount
                    End If
    
                    Engine_Text_Render tmpString, x + HALF_ITEM_SIZE - Engine_Text_Width(tmpString, True) * 0.5, y + LETRA_GRANDE_Y, Color, 1
                End If

                If UserInventory(i).Equipped Then
                    Call RGBAList(Color, 255, 0, 0)
                    tmpString = EQUIPPED_CHARACTER
                    Engine_Text_Render tmpString, x + EQUIPPED_X, y + EQUIPPED_Y, Color, 1
                End If
                Dim cdTime As Single: cdTime = GetLongerCdPercent(i)
                If cdTime > 0 And cdTime < 1 Then
                    Call renderCooldownsInventory(x, y, cdTime)
                    
                End If
            End If
        End If

    Next i
    Exit Sub

DrawInventory_Err:
    Call RegistrarError(Err.Number, Err.Description, "clsGrapchicalInventory.DrawInventory", Erl)
    Resume Next
    
End Sub

Public Sub DrawDraggedItem()
    
    On Error GoTo DrawDraggedItem_Err
    

    Dim Color(3) As RGBA

    If Dragging > 0 Then
        Call RGBAList(Color, 255, 255, 255, 200)
        Call Draw_GrhColor(UserInventory(Dragging).GrhIndex, XDrag - HALF_ITEM_SIZE, YDrag - HALF_ITEM_SIZE, Color)
    End If

    
    Exit Sub

DrawDraggedItem_Err:
    Call RegistrarError(Err.Number, Err.Description, "clsGrapchicalInventory.DrawDraggedItem", Erl)
    Resume Next
    
End Sub

Private Sub InventoryWindow_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    On Error GoTo InventoryWindow_MouseDown_Err
    

    If Button = vbRightButton Then

        ' Si hace clic fuera del borde, salimos
        If x < InvOffsetX Or y < InvOffsetY Or x >= InvOffsetX + InvWidth Or y >= InvOffsetY + InvHeight Then Exit Sub
        
        ' Obtenemos el slot
        Dim Slot As Byte

        Slot = GetSlot(x, y)

        ' Si hay un item, iniciamos el drag
        If Slot > 0 Then
            If UserInventory(Slot).GrhIndex > 0 Then
                Dragging = Slot
                XDrag = x
                YDrag = y
            End If
        End If

    End If

    
    Exit Sub

InventoryWindow_MouseDown_Err:
    Call RegistrarError(Err.Number, Err.Description, "clsGrapchicalInventory.InventoryWindow_MouseDown", Erl)
    Resume Next
    
End Sub

Private Sub InventoryWindow_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    On Error GoTo InventoryWindow_MouseMove_Err
    

    ' Si estamos arrastrando
    If Dragging > 0 Then

        ' Mantenemos el item dentro del Picture
        If x < HALF_ITEM_SIZE Then x = HALF_ITEM_SIZE
        If y < HALF_ITEM_SIZE Then y = HALF_ITEM_SIZE
        If x > InventoryWindow.ScaleWidth - HALF_ITEM_SIZE Then x = InventoryWindow.ScaleWidth - HALF_ITEM_SIZE
        If y > InventoryWindow.ScaleHeight - HALF_ITEM_SIZE Then y = InventoryWindow.ScaleHeight - HALF_ITEM_SIZE

        XDrag = x
        YDrag = y

        InvNeedsRedraw = True

    End If

    
    Exit Sub

InventoryWindow_MouseMove_Err:
    Call RegistrarError(Err.Number, Err.Description, "clsGrapchicalInventory.InventoryWindow_MouseMove", Erl)
    Resume Next
    
End Sub

Private Sub InventoryWindow_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    On Error GoTo InventoryWindow_MouseUp_Err
    
    
    ' Me mantengo dentro de los límites
    If x < HALF_ITEM_SIZE Then x = HALF_ITEM_SIZE
    If y < HALF_ITEM_SIZE Then y = HALF_ITEM_SIZE
    If x > InventoryWindow.ScaleWidth - HALF_ITEM_SIZE Then x = InventoryWindow.ScaleWidth - HALF_ITEM_SIZE
    If y > InventoryWindow.ScaleHeight - HALF_ITEM_SIZE Then y = InventoryWindow.ScaleHeight - HALF_ITEM_SIZE

    ' Calculo el slot
    Dim Slot As Byte

    Slot = GetSlot(x, y)

    ' Si estamos arrastrando y soltamos el clic derecho
    If Dragging > 0 And Button = vbRightButton Then
        ' Disparo un evento para que la lógica del drag&drop sea independiente al inventario
        RaiseEvent ItemDropped(Dragging, Slot, x, y)

        ' Dejo de dibujar el item arrastrado
        Dragging = 0
        InvNeedsRedraw = True

    End If
    
    ' Si es un slot válido
    If Slot > 0 Then

        ' Si no hay un item en el slot
        If UserInventory(Slot).GrhIndex <= 0 Then
            ' Deseleccionamos
            Slot = 0

        End If
        
        wasClickedInside = True
    Else
        wasClickedInside = False

    End If

    ' Si el slot cambio, renderizamos
    If Slot <> InvSelectedItem Then
        InvSelectedItem = Slot
        InvNeedsRedraw = True

    End If

    
    Exit Sub

InventoryWindow_MouseUp_Err:
    Call RegistrarError(Err.Number, Err.Description, "clsGrapchicalInventory.InventoryWindow_MouseUp", Erl)
    Resume Next
    
End Sub

Public Function GetInfo(ByVal Item As Integer) As String
    
    On Error GoTo GetInfo_Err
    

    Select Case ObjData(Item).ObjType

        Case eObjType.otWeapon
            GetInfo = "Golpe: " & ObjData(Item).MinHit & "/" & ObjData(Item).MaxHit

        Case eObjType.otArmadura
            GetInfo = "Defensa: " & ObjData(Item).MinDef & "/" & ObjData(Item).MaxDef

        Case eObjType.otESCUDO
            GetInfo = "Defensa: " & ObjData(Item).MinDef & "/" & ObjData(Item).MaxDef

        Case eObjType.otCASCO
            GetInfo = "Defensa: " & ObjData(Item).MinDef & "/" & ObjData(Item).MaxDef

        Case Else
            GetInfo = ObjData(Item).Texto

    End Select

    
    Exit Function

GetInfo_Err:
    Call RegistrarError(Err.Number, Err.Description, "clsGrapchicalInventory.GetInfo", Erl)
    Resume Next
    
End Function

Property Let MostrarCantidades(ByVal Mostrar As Boolean)
    
    On Error GoTo MostrarCantidades_Err
    
    ShowQuantities = Mostrar
    
    Exit Property

MostrarCantidades_Err:
    Call RegistrarError(Err.Number, Err.Description, "clsGrapchicalInventory.MostrarCantidades", Erl)
    Resume Next
    
End Property

Public Sub UpdateSelectedItemLastUse()
    CdTimes(ObjData(UserInventory(SelectedItem()).ObjIndex).cdType) = GetTickCount()
End Sub

Public Function IsItemOnCd(ByVal Slot As Integer) As Boolean
    If Slot < 0 Or Slot > UBound(UserInventory) Then
        IsItemOnCd = False
        Exit Function
    End If
    Dim ElapsedTime As Long
    ElapsedTime = GetTickCount() - CdTimes(ObjData(UserInventory(Slot).ObjIndex).cdType)
    IsItemOnCd = ElapsedTime < ObjData(UserInventory(Slot).OBJIndex).Cooldown
End Function

Public Function GetActiveWeaponSlot() As Integer
    Dim Slot As Integer
    For Slot = 1 To UBound(UserInventory)
        If UserInventory(Slot).Equipped And UserInventory(Slot).ObjType = 2 Then
            GetActiveWeaponSlot = Slot
            Exit Function
        End If
    Next Slot
    GetActiveWeaponSlot = -1
End Function

Public Function CanUse(ByVal Slot As Integer) As Boolean
    CanUse = UserInventory(Slot).ObjType = eObjType.otWeapon Or UserInventory(Slot).ObjType = eObjType.otPociones Or _
             UserInventory(Slot).ObjType = eObjType.OtHerramientas Or UserInventory(Slot).ObjType = eObjType.otInstrumentos Or _
             UserInventory(Slot).ObjType = eObjType.otNudillos Or UserInventory(Slot).ObjType = eObjType.OtCofre Or _
             (ObjData(UserInventory(Slot).ObjIndex).CDType > 0 And ObjData(UserInventory(Slot).ObjIndex).CDType < e_CdTypes.CDCount)

End Function

Public Function GetLongerCdPercent(ByVal Slot As Integer) As Single
    GetLongerCdPercent = 0
    If Not CanUse(Slot) Then
        Exit Function
    End If
    Dim pendingTime, tempPending As Long
    Dim currTime As Long
    GetLongerCdPercent = 0
    'test stung
    currTime = GetTickCount()
    pendingTime = StunEndTime - currTime
    If pendingTime > 0 Then
        GetLongerCdPercent = 1 - pendingTime / TotalStunTime
    End If
    ' test item cd
    Dim itemCd As Long: itemCd = ObjData(UserInventory(Slot).OBJIndex).Cooldown
    If itemCd > 0 Then
        tempPending = currTime - CdTimes(ObjData(UserInventory(Slot).ObjIndex).cdType)
        If tempPending > pendingTime And tempPending < itemCd Then
            GetLongerCdPercent = tempPending / itemCd
            pendingTime = tempPending
        End If
    End If
    'test global wp cd
    If UserInventory(Slot).ObjType = 2 Or UserInventory(Slot).ObjType = 46 Then
        If cooldown_ataque.totalTime > 0 Then
            tempPending = currTime - cooldown_ataque.initialTime
            If tempPending > pendingTime And tempPending < cooldown_ataque.totalTime Then
                pendingTime = tempPending
                GetLongerCdPercent = pendingTime / cooldown_ataque.totalTime
            End If
        End If
    End If
End Function

