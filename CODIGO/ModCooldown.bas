Attribute VB_Name = "modCooldowns"
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
Private Cooldowns As New Collection


Private colorCooldown As RGBA


Public Sub renderCooldowns(ByVal x As Integer, ByVal y As Integer)



Dim Item As clsCooldown
Call SetRGBA(colorCooldown, 125, 125, 125, 120)


Dim i As Integer
Dim currTime As Long
Dim grh As grh
Dim colores() As RGBA
ReDim colores(3)
Call SetRGBA(colores(0), 255, 255, 255, 255)
Call SetRGBA(colores(1), 255, 255, 255, 255)
Call SetRGBA(colores(2), 255, 255, 255, 255)
Call SetRGBA(colores(3), 255, 255, 255, 255)



i = 1
Do While i <= Cooldowns.count
    Set Item = Cooldowns(i)
    
    currTime = GetTickCount() - Item.initialTime
    If currTime >= Item.totalTime Then
        Cooldowns.Remove (i)
    Else
        Call InitGrh(grh, Item.iconGrh)
        Call Draw_GrhIndex(310, x - 17, y - 17)
        Call Grh_Render_Advance(grh, x - 16, y - 16, 32, 32, colores)
        Call Engine_Draw_Load(x, y, 32, 32, colorCooldown, currTime * 360 / Item.totalTime)
        x = x - 36
        i = i + 1
    End If
Loop

End Sub



Public Sub renderCooldownsInventory(ByVal x As Integer, ByVal y As Integer, ByVal cdProgress As Single)



    Call SetRGBA(colorCooldown, 50, 25, 15, 170)
    
    x = x + 16
    y = y + 16
    
    Dim currTime As Long
    Dim colores() As RGBA
    Dim progress As Single
    ReDim colores(3)
    Call SetRGBA(colores(0), 255, 255, 255, 125)
    Call SetRGBA(colores(1), 255, 255, 255, 125)
    Call SetRGBA(colores(2), 255, 255, 255, 125)
    Call SetRGBA(colores(3), 255, 255, 255, 125)
    
    If cdProgress >= 1 Then
       Set cooldown_ataque = Nothing
    Else
        Call Engine_Draw_Load(x, y, 32, 32, colorCooldown, 360 * cdProgress)
        x = x - 36
        i = i + 1
    End If

End Sub

Public Sub addCooldown(cooldown As clsCooldown)

    Dim cd As clsCooldown
    
    For Each cd In Cooldowns
        If cd Is cooldown Then Exit Sub
    Next cd
    
    Cooldowns.Add cooldown

End Sub



