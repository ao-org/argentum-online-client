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


Private isInventory As Boolean ' Para separar al inventario del main, de los otros


Dim InvMaxSlots As Integer
Dim InvWidth As Integer
Dim InvHeight As Integer
Dim InvOffsetX As Integer
Dim InvOffsetY As Integer
Dim InvSelectedItem As Long             'Currently selected item
Dim TempItem As Long


Private MAININVENTORY_MIDSPACE As Byte

' Slot item arrastrado (0 si no está arrastrando)
Dim Dragging As Integer
' Posición del item arrastrado
Dim XDrag As Integer
Dim YDrag As Integer

' Flag para saber si se cliqueó dentro del inventario
' al hacer clic sobre el Picture
Dim wasClickedInside As Boolean

' Flag para saber si hay que volver a dibujar
Dim InvNeedsRedraw As Boolean

'Dim InvSurface As DirectDrawSurface7            'DD Surface used to render everything

Dim UserInventory() As Inventory    'User's inventory

Dim WithEvents InventoryWindow As PictureBox    'Placeholder where to render the inventory
Attribute InventoryWindow.VB_VarHelpID = -1

' Evento disparado cuando un item se soltó
Public Event ItemDropped(ByVal Drag As Integer, ByVal Drop As Integer, ByVal x As Integer, ByVal y As Integer)

#If ConMenuesConextuales = 1 Then
    Dim ItemMenu As Menu    'Menu to be shown as pop up
#End If

' Valores constantes que no ameritan que los reciba como parámetro
Const ITEM_SIZE = 32 ' El ancho de los items (nunca debería cambiar
Const LETRA_CHICA_X = 3 ' Offset X de cantidad de items (letra chica)
Const LETRA_CHICA_Y = 23 ' Offset Y de cantidad de items (letra chica)
Const LETRA_GRANDE_X = 4 ' Offset X de cantidad de items (letra grande)
Const LETRA_GRANDE_Y = 20 ' Offset Y de cantidad de items (letra grande)
Const EQUIPPED_X = 30 ' Offset X del caracter de equipado
Const EQUIPPED_Y = 0 ' Offset Y del caracter de equipado
Const EQUIPPED_CHARACTER = "+" ' Caracter de item equipado
Const SELECTION_BOX_W = 34 ' Ancho de la cajita de selección (se centra automáticamente)
Const SELECTION_BOX_H = 34 ' Alto de la cajita de selección (se centra automáticamente)

'Private Const MAININVENTORY_MIDSPACE As Integer = 0 '9 'Separador de inventario básico y extra slots en el inventario principal del usuario

' Valores calculados
Const HALF_ITEM_SIZE = ITEM_SIZE \ 2
Const SELECTION_BOX_X = (ITEM_SIZE - SELECTION_BOX_W) \ 2
Const SELECTION_BOX_Y = (ITEM_SIZE - SELECTION_BOX_H) \ 2

' Valores a calcular cuando se cree el inventario
Dim ITEM_OFFSET_X As Integer
Dim ITEM_OFFSET_Y As Integer
Dim ITEMS_PER_ROW As Integer
Dim HALF_BORDER_X As Integer
Dim HALF_BORDER_Y As Integer
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
    Public Sub Initialize(ByRef DirectDraw As DirectDraw7, ByRef InvPic As PictureBox, ByRef PopUpMenu As Menu, _
                        Optional ByVal Width As Integer = 0, Optional ByVal Height As Integer = 0, Optional ByVal OffsetX As Integer = 0, Optional ByVal OffsetY As Integer = 0, _
                        Optional ByVal SpaceX As Integer = 0, Optional ByVal SpaceY As Integer = 0, Optional thisInventory As Boolean = False, Optional MIDSPACE As Byte = 0)

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
    
End Sub

Public Property Get SelectedItem() As Long
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 27/07/04
'Retrieves the selected item index
'***************************************************
    SelectedItem = InvSelectedItem
End Property


Public Property Get MaxHit(ByVal slot As Byte) As Integer
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 27/07/04
'Retrieves the max hit of the selected item
'***************************************************
    MaxHit = UserInventory(slot).MaxHit
End Property

Public Property Get MinHit(ByVal slot As Byte) As Integer
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 27/07/04
'Retrieves the min hit of the selected item
'***************************************************
    MinHit = UserInventory(slot).MinHit
End Property

Public Property Get Def(ByVal slot As Byte) As Integer
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 27/07/04
'Retrieves the defense of the selected item
'***************************************************
    Def = UserInventory(slot).Def
End Property

Public Property Get GrhIndex(ByVal slot As Byte) As Long
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 27/07/04
'Retrieves the grh index of the selected item
'***************************************************
    GrhIndex = UserInventory(slot).GrhIndex
End Property

Public Property Get Valor(ByVal slot As Byte) As Single
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 27/07/04
'Retrieves the value of the selected item
'***************************************************
    Valor = UserInventory(slot).Valor
End Property

Public Property Get Amount(ByVal slot As Byte) As Long
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 27/07/04
'Retrieves the selected item's amount
'***************************************************
    If slot = FLAGORO Then
        Amount = UserGLD
    ElseIf slot >= LBound(UserInventory) And slot <= UBound(UserInventory) Then
        Amount = UserInventory(slot).Amount
    End If
End Property

Public Property Get OBJIndex(ByVal slot As Byte) As Integer
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 27/07/04
'Retrieves the selected item's object index
'***************************************************
    OBJIndex = UserInventory(slot).OBJIndex
End Property

Public Property Get ObjType(ByVal slot As Byte) As Integer
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 27/07/04
'Retrieves the selected item's object type
'***************************************************
    ObjType = UserInventory(slot).ObjType
End Property

Public Property Get ItemName(ByVal slot As Byte) As String
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 27/07/04
'Retrieves the selected item's name
'***************************************************
    ItemName = UserInventory(slot).name
End Property

Public Property Get Equipped(ByVal slot As Byte) As Boolean
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 27/07/04
'Retrieves True if the item at the given pos is eqiupped
'***************************************************
    Equipped = UserInventory(slot).Equipped
End Property

Public Property Get PuedeUsar(ByVal slot As Byte) As Byte
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 27/07/04
'Retrieves True if the item at the given pos is eqiupped
'***************************************************
    PuedeUsar = UserInventory(slot).PuedeUsar
End Property

Public Property Get MaxSlots() As Integer
    MaxSlots = InvMaxSlots
End Property

Public Property Get Columns() As Integer
    Columns = ITEMS_PER_ROW
End Property

Public Property Get SlotDrag() As Byte
    SlotDrag = Dragging
End Property

Public Property Get NeedsRedraw() As Boolean
    NeedsRedraw = InvNeedsRedraw
End Property

Public Sub ReDraw()
    ' Provoca que el inventario se vuelta a dibujar el próximo frame
    InvNeedsRedraw = True
End Sub

Public Sub SetItem(ByVal slot As Byte, ByVal eOBJIndex As Integer, ByVal eAmount As Integer, ByVal eEquipped As Byte, _
                        ByVal eGrhIndex As Long, ByVal eObjType As Integer, ByVal eMaxHit As Integer, ByVal eMinHit As Integer, _
                        ByVal eDef As Integer, ByVal eValor As Single, ByVal eName As String, ByVal PuedeUsar As Byte)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 12/04/06
'Sets all data for a given inventory slot
'***************************************************

    If slot < 1 Or slot > InvMaxSlots Then Exit Sub

    With UserInventory(slot)
        .Amount = eAmount
        .Def = eDef
        .Equipped = eEquipped
        .GrhIndex = eGrhIndex
        .MaxHit = eMaxHit
        .MinHit = eMinHit
        .name = eName
        .OBJIndex = eOBJIndex
        .ObjType = eObjType
        .Valor = eValor
        .PuedeUsar = PuedeUsar
    End With

    InvNeedsRedraw = True

End Sub

Public Sub SeleccionarItem(ByVal SelectedItem As Byte)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 27/07/04
'Retrieves the selected item index
'***************************************************
    InvSelectedItem = SelectedItem
    InvNeedsRedraw = True
End Sub

Public Property Get ClickedInside() As Boolean
    ClickedInside = wasClickedInside
End Property

Public Function GetSlot(ByVal x As Integer, ByVal y As Integer) As Byte
    ' Restamos el offset del inventario
    ' (más la mitad del borde, para hacer que los clics en los bordes sean como en el slot)
        
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
End Function

Public Sub DrawInventory()
    Dim i As Byte
    Dim x As Integer
    Dim y As Integer
    
    Dim tmpString As String
    Dim color(3) As Long 'Si le queres dar color a la letra pasa este parametro dsp xD
    
    InvNeedsRedraw = False

    For i = 1 To UBound(UserInventory)
        x = ((i - 1) Mod ITEMS_PER_ROW) * ITEM_OFFSET_X + InvOffsetX
        y = ((i - 1) \ ITEMS_PER_ROW) * ITEM_OFFSET_Y + InvOffsetY
        
        
         If isInventory Then
            If (i > MAX_NORMAL_INVENTORY_SLOTS) Then
                y = y + MAININVENTORY_MIDSPACE
                If (i - MAX_NORMAL_INVENTORY_SLOTS) <= (UserInvUnlocked * (InvWidth / (ITEM_OFFSET_X))) Then
                    'Slot Desbloqueado
                    Call engine.Draw_GrhColor(GRH_INVENTORYSLOT_EXTRA, x, y, COLOR_WHITE)
                Else
                    'Slot Bloqueado
                    Call engine.Draw_GrhColor(GRH_INVENTORYSLOT_LOCKED, x, y, COLOR_WHITE)
                End If
            Else
                'Slot Default
                 Call engine.Draw_GrhColor(GRH_INVENTORYSLOT, x, y, COLOR_WHITE)
            End If
         End If
        

            If InvSelectedItem = i Then
                Engine_Draw_Box x + SELECTION_BOX_X, y + SELECTION_BOX_Y, SELECTION_BOX_W, SELECTION_BOX_H, D3DColorARGB(100, 100, 100, 100)
            End If
            
        If UserInventory(i).GrhIndex Then



            If UserInventory(i).PuedeUsar > 0 Then
                color(0) = D3DColorXRGB(255, 100, 100)
                color(1) = D3DColorXRGB(255, 100, 100)
                color(2) = D3DColorXRGB(255, 100, 100)
                color(3) = D3DColorXRGB(255, 100, 100)
                Call engine.Draw_GrhColor(UserInventory(i).GrhIndex, x, y, color)

            Else
                Call engine.Draw_GrhColor(UserInventory(i).GrhIndex, x, y, COLOR_WHITE)
            End If
                
            color(0) = RGB(239, 228, 176)
            color(1) = RGB(239, 228, 176)
            color(2) = RGB(239, 228, 176)
            color(3) = RGB(239, 228, 176)
            
            tmpString = UserInventory(i).Amount
            If tmpString = 10000 Then
                tmpString = "Max"
            End If
            
            If InvSelectedItem <> i Then
                engine.Engine_Text_Render_LetraChica tmpString, x + LETRA_CHICA_X, y + LETRA_CHICA_Y, color, 4, True
            Else
                engine.Engine_Text_Render tmpString, x + LETRA_GRANDE_X, y + LETRA_GRANDE_Y, color, 1, True
            End If
        
            If UserInventory(i).Equipped Then
                tmpString = EQUIPPED_CHARACTER
                engine.Engine_Text_Render tmpString, x + EQUIPPED_X - engine.Engine_Text_Width(tmpString, True) \ 2, y + EQUIPPED_Y - engine.Engine_Text_Height(tmpString, True), COLOR_WHITE, 1, True
            End If
        End If
    Next i

End Sub

Public Sub DrawDraggedItem()
    Dim color(3) As Long

    If Dragging > 0 Then
        color(0) = D3DColorARGB(200, 255, 255, 255)
        color(1) = D3DColorARGB(200, 255, 255, 255)
        color(2) = D3DColorARGB(200, 255, 255, 255)
        color(3) = D3DColorARGB(200, 255, 255, 255)
        Call engine.Draw_GrhColor(UserInventory(Dragging).GrhIndex, XDrag - HALF_ITEM_SIZE, YDrag - HALF_ITEM_SIZE, color)
    End If
End Sub

Private Sub InventoryWindow_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

    If Button = vbRightButton Then
        ' Si hace clic fuera del borde, salimos
        If x < InvOffsetX Or y < InvOffsetY Or x >= InvOffsetX + InvWidth Or y >= InvOffsetY + InvHeight Then Exit Sub
        
        ' Obtenemos el slot
        Dim slot As Byte
        slot = GetSlot(x, y)

        ' Si hay un item, iniciamos el drag
        If slot > 0 Then
            Dragging = slot
            XDrag = x
            YDrag = y
        End If
    End If

End Sub

Private Sub InventoryWindow_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

    ' Si estamos arrastrando
    If Dragging >= 0 Then
        ' Mantenemos el item dentro del Picture
        If x < HALF_ITEM_SIZE Then x = HALF_ITEM_SIZE
        If y < HALF_ITEM_SIZE Then y = HALF_ITEM_SIZE
        If x > InventoryWindow.ScaleWidth - HALF_ITEM_SIZE Then x = InventoryWindow.ScaleWidth - HALF_ITEM_SIZE
        If y > InventoryWindow.ScaleHeight - HALF_ITEM_SIZE Then y = InventoryWindow.ScaleHeight - HALF_ITEM_SIZE

        XDrag = x
        YDrag = y

        InvNeedsRedraw = True
    End If

End Sub

Private Sub InventoryWindow_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

    ' Calculo el slot
    Dim slot As Byte
    slot = GetSlot(x, y)

    ' Si estamos arrastrando y soltamos el clic derecho
    If Dragging > 0 And Button = vbRightButton Then
        ' Disparo un evento para que la lógica del drag&drop sea independiente al inventario
        RaiseEvent ItemDropped(Dragging, slot, x, y)

        ' Dejo de dibujar el item arrastrado
        Dragging = 0
        InvNeedsRedraw = True
    End If
    
    ' Si es un slot válido
    If slot > 0 Then
        ' Si no hay un item en el slot
        If UserInventory(slot).GrhIndex <= 0 Then
            ' Deseleccionamos
            slot = 0
        End If
        
        wasClickedInside = True
    Else
        wasClickedInside = False
    End If

    ' Si el slot cambio, renderizamos
    If slot <> InvSelectedItem Then
        InvSelectedItem = slot
        InvNeedsRedraw = True
    End If

End Sub

Public Function GetInfo(ByVal Item As Integer) As String

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

End Function
