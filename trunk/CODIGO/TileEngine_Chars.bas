Attribute VB_Name = "TileEngine_Chars"
Option Explicit

Public Sub ResetCharInfo(ByVal charindex As Integer)

    With charlist(charindex)
    
        .active = 0
        .AlphaPJ = 0
        .Escribiendo = False
        .appear = 0
        .status = 0
        .invisible = False
        .Arma_Aura = vbNullString
        .Body_Aura = vbNullString
        .AuraAngle = 0
        .Head_Aura = vbNullString
        .Speeding = 0
        .Otra_Aura = vbNullString
        .Escudo_Aura = vbNullString
        .Particula = 0
        .ParticulaTime = 0
        .particle_count = 0
        .FxCount = 0
        .CreandoCant = 0
        .Moving = 0
        .MUERTO = False
        .nombre = vbNullString
        .pie = False
        ' .Pos.X = 0
        '.Pos.Y = 0
        
        .MovArmaEscudo = False
        .TimerAct = False
        .TimerM = 128
        .TimerI = 128
        .TimerIAct = False
        .dialog = vbNullString
        .dialogExp = vbNullString
        .dialogEfec = vbNullString
        .dialogOro = vbNullString
        .SubeExp = 0
        .group_index = 0
        .clan_index = 0
        .clan_nivel = 0
        .BarTime = 0
        .BarAccion = 0
        .MaxBarTime = 0
        .UserMaxHp = 0
        .UserMinHp = 0
        
        .FxIndex = 0
        
    End With
    
End Sub


Public Sub EraseChar(ByVal charindex As Integer)
    '*****************************************************************
    'Erases a character from CharList and map
    '*****************************************************************
    
    If charindex = 0 Then Exit Sub
    If charlist(charindex).active = 0 Then Exit Sub

    charlist(charindex).active = 0
    
    'Update lastchar
    If charindex = LastChar Then

        Do Until charlist(LastChar).active = 1
            LastChar = LastChar - 1

            If LastChar = 0 Then Exit Do
        Loop

    End If
    
    MapData(charlist(charindex).Pos.x, charlist(charindex).Pos.y).charindex = 0
    
    'Remove char's dialog
    Call Dialogos.RemoveDialog(charindex)
    
    Call ResetCharInfo(charindex)
    
    'Update NumChars
    NumChars = NumChars - 1

End Sub

Sub MakeChar(ByVal charindex As Integer, ByVal Body As Integer, ByVal Head As Integer, ByVal Heading As Byte, ByVal x As Integer, ByVal y As Integer, ByVal Arma As Integer, ByVal Escudo As Integer, ByVal Casco As Integer, ByVal ParticulaFx As Byte, ByVal appear As Byte)

    On Error Resume Next

    'Apuntamos al ultimo Char
    ' Debug.Print charindex
    If charindex > LastChar Then LastChar = charindex
    
    With charlist(charindex)

        'If the char wasn't allready active (we are rewritting it) don't increase char count
        If .active = 0 Then NumChars = NumChars + 1
        
        If Arma = 0 Then Arma = 2
        If Escudo = 0 Then Escudo = 2
        If Casco = 0 Then Casco = 2
        
        .iHead = Head
        .iBody = Body
        
        .Head = HeadData(Head)
        .Body = BodyData(Body)
        .Arma = WeaponAnimData(Arma)
        
        .Escudo = ShieldAnimData(Escudo)
        .Casco = CascoAnimData(Casco)
        
        .Heading = Heading
        
        'Reset moving stats
        .Moving = 0
        .MoveOffsetX = 0
        .MoveOffsetY = 0
        
        'Update position
        .Pos.x = x
        .Pos.y = y
        
        'Make active
        .active = 1
        
        .AlphaPJ = 255
        
        If BodyData(Body).HeadOffset.y = -26 Then
            .EsEnano = True
        Else
            .EsEnano = False

        End If
        
        If .Particula = ParticulaFx Then
            ParticulaFx = 0

        End If
        
        If ParticulaFx <> 0 Then
            .Particula = ParticulaFx
            Call General_Char_Particle_Create(ParticulaFx, charindex, -1)

        End If
      
    End With
    
    'Plot on map
    MapData(x, y).charindex = charindex

End Sub

Public Sub Char_Move_by_Head(ByVal charindex As Integer, ByVal nHeading As E_Heading)
    '*****************************************************************
    'Starts the movement of a character in nHeading direction
    '*****************************************************************

    If nHeading = 0 Then
        Debug.Print "Heading: " & nHeading

    End If

    On Error Resume Next

    Dim addx As Integer

    Dim addy As Integer

    Dim x    As Integer

    Dim y    As Integer

    Dim nX   As Integer

    Dim nY   As Integer
    
    With charlist(charindex)
        x = .Pos.x
        y = .Pos.y
        
        'Figure out which way to move
        Select Case nHeading

            Case E_Heading.NORTH
                addy = -1
        
            Case E_Heading.EAST
                addx = 1
        
            Case E_Heading.south
                addy = 1
            
            Case E_Heading.WEST
                addx = -1

        End Select
        
        nX = x + addx
        nY = y + addy
        
        MapData(nX, nY).charindex = charindex
        .Pos.x = nX
        .Pos.y = nY
        MapData(x, y).charindex = 0
        
        .MoveOffsetX = -1 * (32 * addx)
        .MoveOffsetY = -1 * (32 * addy)
        
        .Moving = 1
        
        'Attached to ladder ;)
        If MapData(nX, nY).ObjGrh.GrhIndex = 26940 Then
            .Heading = E_Heading.NORTH
        Else
            .Heading = nHeading

        End If
        
        .scrollDirectionX = addx
        .scrollDirectionY = addy

    End With
    
    If UserEstado <> 1 Then Call DoPasosFx(charindex)
    
    'areas viejos
    If (nY < MinLimiteY) Or (nY > MaxLimiteY) Or (nX < MinLimiteX) Or (nX > MaxLimiteX) Then
        Call EraseChar(charindex)

    End If
    
End Sub

Public Sub Char_Move_by_Pos(ByVal charindex As Integer, ByVal nX As Integer, ByVal nY As Integer)

    On Error Resume Next

    Dim x        As Integer

    Dim y        As Integer

    Dim addx     As Integer

    Dim addy     As Integer

    Dim nHeading As E_Heading
    
    With charlist(charindex)
        x = .Pos.x
        y = .Pos.y
        
        MapData(x, y).charindex = 0
        
        addx = nX - x
        addy = nY - y
        
        If Sgn(addx) = 1 Then
            nHeading = E_Heading.EAST

        End If
        
        If Sgn(addx) = -1 Then
            nHeading = E_Heading.WEST

        End If
        
        If Sgn(addy) = -1 Then
            nHeading = E_Heading.NORTH

        End If
        
        If Sgn(addy) = 1 Then
            nHeading = E_Heading.south

        End If
        
        MapData(nX, nY).charindex = charindex
        
        .Pos.x = nX
        .Pos.y = nY
        
        .MoveOffsetX = -1 * (TilePixelWidth * addx)
        .MoveOffsetY = -1 * (TilePixelHeight * addy)
        
        .Moving = 1
        
        If MapData(nX, nY).ObjGrh.GrhIndex = 26940 Then
            .Heading = E_Heading.NORTH
        Else
            .Heading = nHeading

        End If
        
        .scrollDirectionX = Sgn(addx)
        .scrollDirectionY = Sgn(addy)

    End With
    
    If Not EstaPCarea(charindex) Then

        'Call Dialogos.RemoveDialog(CharIndex)
        'Call Hits.RemoveHit(CharIndex)
    End If

    If (nY < MinLimiteY) Or (nY > MaxLimiteY) Or (nX < MinLimiteX) Or (nX > MaxLimiteX) Then
        Call EraseChar(charindex)

    End If

End Sub

Private Function EstaPCarea(ByVal charindex As Integer) As Boolean

    With charlist(charindex).Pos
        EstaPCarea = .x > UserPos.x - MinXBorder And .x < UserPos.x + MinXBorder And .y > UserPos.y - MinYBorder And .y < UserPos.y + MinYBorder

    End With

End Function

Public Function EstaEnArea(ByVal x As Integer, ByVal y As Integer) As Boolean
    EstaEnArea = x > UserPos.x - MinXBorder And x < UserPos.x + MinXBorder And y > UserPos.y - MinYBorder And y < UserPos.y + MinYBorder

End Function

Public Function Char_Check(ByVal char_index As Integer) As Boolean

    '**************************************************************
    'Author: Aaron Perkins - Modified by Juan Mart�n Sotuyo Dodero
    'Last Modify Date: 1/04/2003
    '
    '**************************************************************
    'check char_index
    If char_index > 0 And char_index <= LastChar Then
        Char_Check = (charlist(char_index).Heading > 0)

    End If
    
End Function

Public Function Char_FX_Group_Next_Open(ByVal char_index As Integer) As Integer

    '*****************************************************************
    'Author: Augusto Jos� Rando
    '*****************************************************************
    On Error GoTo ErrorHandler:

    Dim loopc As Long
    
    If charlist(char_index).FxCount = 0 Then
        charlist(char_index).FxCount = 1
        ReDim charlist(char_index).FxList(1 To 1)
        Char_FX_Group_Next_Open = 1
        Exit Function

    End If
    
    loopc = 1

    Do Until charlist(char_index).FxList(loopc).FxIndex = 0

        If loopc = charlist(char_index).FxCount Then
            Char_FX_Group_Next_Open = charlist(char_index).FxCount + 1
            charlist(char_index).FxCount = Char_FX_Group_Next_Open
            ReDim Preserve charlist(char_index).FxList(1 To Char_FX_Group_Next_Open)
            Exit Function

        End If

        loopc = loopc + 1
    Loop

    Char_FX_Group_Next_Open = loopc
    Exit Function

ErrorHandler:
    charlist(char_index).FxCount = 1
    ReDim charlist(char_index).FxList(1 To 1)
    Char_FX_Group_Next_Open = 1

End Function

Public Sub Char_Dialog_Set(ByVal char_index As Integer, ByVal char_dialog As String, ByVal char_dialog_color As Long, ByVal char_dialog_life As Byte, ByVal Sube As Byte, Optional ByVal font_index As Integer = 1)
    
    If Char_Check(char_index) Then
        charlist(char_index).dialog = char_dialog
        charlist(char_index).dialog_color = char_dialog_color
        charlist(char_index).dialog_life = char_dialog_life
        charlist(char_index).dialog_font_index = font_index
        charlist(char_index).dialog_scroll = True
        charlist(char_index).dialog_offset_counter_y = -(IIf(BodyData(charlist(char_index).iBody).HeadOffset.y = 0, -32, BodyData(charlist(char_index).iBody).HeadOffset.y) / 2)
        charlist(char_index).AlphaText = 255

    End If

    Dim slot As Integer

    Dim i    As Long
    
    slot = BinarySearch(char_index)
    
    If slot < 0 Then
        If dialogCount = MAX_DIALOGS Then Exit Sub  'Out of space! Should never happen....
        
        'We need to add it. Get insertion index and move list backwards.
        slot = Not slot
        
        For i = dialogCount To slot + 1 Step -1
            dialogs(i) = dialogs(i - 1)
        Next i
        
        dialogCount = dialogCount + 1

    End If
    
    If char_dialog_life = 250 Then

        With dialogs(slot)
            .startTime = (GetTickCount() And &H7FFFFFFF)
            .lifeTime = MS_ADD_EXTRA + (MS_PER_CHAR * Len(char_dialog))
            .charindex = char_index

        End With

    Else

        With dialogs(slot)
            .startTime = (GetTickCount() And &H7FFFFFFF)
            .lifeTime = (MS_PER_CHAR * Len(char_dialog))
            .charindex = char_index

        End With

    End If
    
End Sub


Public Sub Char_Dialog_Remove(ByVal char_index As Integer, ByVal Index As Integer)

    If char_index = 0 Then Exit Sub

    If charlist(char_index).AlphaText > 0 Then
        charlist(char_index).AlphaText = charlist(char_index).AlphaText - (scroll_dialog_pixels_per_frame * timerTicksPerFrame)
        Exit Sub

    End If

    Dim slot As Integer

    Dim i    As Long
    
    slot = BinarySearch(char_index)
    
    If slot < 0 Then Exit Sub
    
    For i = slot To MAX_DIALOGS - 2
        dialogs(i) = dialogs(i + 1)
    Next i
    
    dialogCount = dialogCount - 1
    
    If Char_Check(char_index) Then
        charlist(char_index).dialog = ""
        charlist(char_index).dialog_color = 0
        charlist(char_index).dialog_life = 0

    End If

End Sub

Public Sub SetCharacterFx(ByVal charindex As Integer, ByVal fX As Integer, ByVal Loops As Integer)

    If fX = 0 Then Exit Sub

    '***************************************************
    'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
    'Last Modify Date: 12/03/04
    'Sets an FX to the character.
    '***************************************************
    Dim indice As Byte

    With charlist(charindex)
    
        indice = Char_FX_Group_Next_Open(charindex)
        
        .FxList(indice).FxIndex = fX
        
        Call InitGrh(.FxList(indice), FxData(fX).Animacion)
        
        .FxList(indice).Loops = Loops
            
    End With

End Sub



Public Function Get_PixelY_Of_Char(ByVal char_index As Integer) As Integer

    '*****************************************************************
    'Author: Pablo Mercavides
    '*****************************************************************
    'Make sure it's a legal char_index
    If Char_Check(char_index) Then
        Get_PixelY_Of_Char = (charlist(char_index).Pos.y - 2 - UserPos.y) * 32 + frmmain.renderer.ScaleWidth / 2
        Get_PixelY_Of_Char = Get_PixelY_Of_Char - 16

    End If

End Function

Public Function Get_Pixelx_Of_Char(ByVal char_index As Integer) As Integer

    '*****************************************************************
    'Author: Pablo Mercavides
    '*****************************************************************
    'Make sure it's a legal char_index
    If Char_Check(char_index) Then
        Get_Pixelx_Of_Char = (charlist(char_index).Pos.x - UserPos.x) * 32 + frmmain.renderer.ScaleWidth / 2
        Get_Pixelx_Of_Char = Get_Pixelx_Of_Char

    End If

End Function

Public Function Get_Pixelx_Of_XY(ByVal x As Byte) As Integer
    '*****************************************************************
    'Author: Pablo Mercavides
    '*****************************************************************
    'Make sure it's a legal char_index
    Get_Pixelx_Of_XY = (x - UserPos.x) * 32 + frmmain.renderer.ScaleWidth / 2
    Get_Pixelx_Of_XY = Get_Pixelx_Of_XY

End Function

Public Function Get_PixelY_Of_XY(ByVal y As Byte) As Integer
    '*****************************************************************
    'Author: Pablo Mercavides
    '*****************************************************************
    'Make sure it's a legal char_index
    Get_PixelY_Of_XY = (y - 2 - UserPos.y) * 32 + frmmain.renderer.ScaleWidth / 2
    Get_PixelY_Of_XY = Get_PixelY_Of_XY - 16

End Function

