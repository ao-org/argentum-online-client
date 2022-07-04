Attribute VB_Name = "modCooldowns"

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

Public Sub addCooldown(cooldown As clsCooldown)

    Dim cd As clsCooldown
    
    For Each cd In Cooldowns
        If cd Is cooldown Then Exit Sub
    Next cd
    
    Cooldowns.Add cooldown

End Sub



