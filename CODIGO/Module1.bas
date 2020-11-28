Attribute VB_Name = "Module1"

Dim UserInventory(1 To MAX_INVENTORY_SLOTS) As Inventory    'User's inventory


Private Function RenderInvItem(InvIndex As Integer)

Dim ItemX As Single
Dim ItemY As Single

ItemX = ((InvIndex - 1) Mod XCantItems)
ItemY = ((InvIndex - 1) \ XCantItems)

Dim temp_array(3) As Long
temp_array(0) = &HFFFFFF
temp_array(1) = &HFFFFFF
temp_array(2) = &HFFFFFF
temp_array(3) = &HFFFFFF

engine.Grh_Inventory_Render UserInventory(InvIndex).grhindex, ItemX * 32, ItemY * 32, temp_array












   ' Dim i As Byte
    'Dim x As Integer
    'Dim y As Integer
    
    'For i = 1 To UBound(UserInventory)
       ' If UserInventory(i).grhindex Then
           ' x = ((i - 1) Mod (InventoryWindow.Width / 32)) * 32 + 2
           ' y = ((i - 1) \ (InventoryWindow.Width / 32)) * 32 + 2
           ' If InvSelectedItem = i Then
               ' Grh_Render_To_Hdc(frmMain.picInv.hdc, UserInventory(i).grhindex, , 0, True
           '     Call engine.Draw_FilledBox(x, y, 32, 32, D3DColorXRGB(0, 0, 0), D3DColorXRGB(255, 0, 0))
         '   End If
           'Call engine.Draw_GrhIndex(UserInventory(i).grhindex, X, Y)
          '  Call engine.Text_Render_ext(UserInventory(i).Amount, y, x, 40, 40, D3DColorXRGB(255, 255, 255))
            
            
            
           ' If UserInventory(i).Equipped Then
            '    Call engine.Text_Render_ext("+", y + 20, x + 20, 40, 40, D3DColorXRGB(255, 255, 255))
           ' End If
        'End If
    'Next i
End Function
Sub Inventory_Render()
MsgBox ("a")
engine.Inventory_Render_Start
Inventory_Render_All
engine.Inventory_Render_End frmMain.picInv.hwnd
End Sub

Public Function Inventory_Render_All()

Dim i As Integer
Dim x As Single
Dim y As Single
Dim tmp As String
Dim tempito(3) As Long

tempito(0) = &HFFFFFFFF
tempito(1) = &HFFFFFFFF
tempito(2) = &HFFFFFFFF
tempito(3) = &HFFFFFFFF

For i = 1 To MAX_INVENTORY_SLOTS
    If UserInventory(i).Amount > 0 Then
        RenderInvItem i
        x = ((i - 1) Mod XCantItems) * 32
        y = (((i - 1) \ XCantItems) + 0.75) * 32 - 3
        
        If UserInventory(i).Amount = 10000 Then
            tmp = "10000"
        Else
            tmp = str(UserInventory(i).Amount)
        End If

       Call engine.Text_Render_ext(tmp, Int(x - 3), Int(y - 2), 40, 40, tempito(1))
        
    End If
Next i

End Function
