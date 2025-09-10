Attribute VB_Name = "TileEngine_Chars"
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

Public Sub ResetCharInfo(ByVal charindex As Integer)
    
    On Error GoTo ResetCharInfo_Err
    

    With charlist(charindex)
    
        .active = 0
        .AlphaPJ = 0
        .appear = 0
        .status = 0
        .Invisible = False
        .Arma_Aura = vbNullString
        .Body_Aura = vbNullString
        .AuraAngle = 0
        .Head_Aura = vbNullString
        .Speeding = 0
        .Otra_Aura = vbNullString
        .Escudo_Aura = vbNullString
        .DM_Aura = vbNullString
        .RM_Aura = vbNullString
        .Particula = 0
        .ParticulaTime = 0
        .particle_count = 0
        .FxCount = 0
        .CreandoCant = 0
        .Moving = False
        .Muerto = False
       ' frmdebug.add_text_tracebox "ResetCharInfo " & .nombre
        .nombre = vbNullString
        .Pie = False
        .simbolo = 0
        .Idle = False
        .Navegando = False
        .LastStep = 0
        ' .Pos.X = 0
        '.Pos.Y = 0
        
        .MovArmaEscudo = False
        .TimerAct = False
        .TimerM = 128
        .TimerI = 128
        .TimerIAct = False
        .dialog = vbNullString
        .group_index = 0
        .clan_index = 0
        .clan_nivel = 0
        .BarTime = 0
        .BarAccion = 0
        .MaxBarTime = 0
        .UserMaxHp = 0
        .UserMinHp = 0
        .Meditating = False
        .ActiveAnimation.PlaybackState = Stopped
        .scrollDirectionX = 0
        .scrollDirectionY = 0
        .MoveOffsetX = 0
        .MoveOffsetY = 0
        
    End With
    
    
    Exit Sub

ResetCharInfo_Err:
    Call RegistrarError(Err.Number, Err.Description, "TileEngine_Chars.ResetCharInfo", Erl)
    Resume Next
    
End Sub


Public Sub EraseChar(ByVal charindex As Integer, Optional ByVal notCancelMe As Boolean = False)

    'Erases a character from CharList and map

    On Error GoTo EraseChar_Err
    
    
    If charindex = 0 Then Exit Sub
    If charlist(charindex).active = 0 Then Exit Sub
    If charindex = UserCharIndex And Not notCancelMe Then Exit Sub
    
    Dim i As Integer
    For i = LBound(Effect) To UBound(Effect)
        If Effect(i).DestinoChar = CharIndex Then
            Effect(i).DestX = charlist(CharIndex).Pos.x
            Effect(i).DesyY = charlist(CharIndex).Pos.y
            Effect(i).DestinoChar = 0
        End If
    Next i
    
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
    If (Not Dialogos Is Nothing) Then
    Call Dialogos.RemoveDialog(charindex)
    End If
    
    Call ResetCharInfo(charindex)
    
    'Update NumChars
    NumChars = NumChars - 1

    
    Exit Sub

EraseChar_Err:
    Call RegistrarError(Err.Number, Err.Description, "TileEngine_Chars.EraseChar", Erl)
    Resume Next
    
End Sub

Sub MakeChar(ByVal CharIndex As Integer, ByVal Body As Integer, ByVal Head As Integer, ByVal Heading As Byte, ByVal x As Integer, ByVal y As Integer, ByVal Arma As Integer, ByVal Escudo As Integer, ByVal Casco As Integer, ByVal CartIndex As Integer, ByVal BackpackIndex As Integer, ByVal ParticulaFx As Byte, ByVal appear As Byte)
    
    On Error GoTo MakeChar_Err

    'Apuntamos al ultimo Char
    ' frmdebug.add_text_tracebox charindex
    If charindex > LastChar Then LastChar = charindex
    
    With charlist(charindex)

        'If the char wasn't allready active (we are rewritting it) don't increase char count
        If .active = 0 Then NumChars = NumChars + 1
        .HasCart = True
        .HasBackpack = True
        If Arma = 0 Or Arma > UBound(WeaponAnimData) Then Arma = 2
        If Escudo = 0 Or Escudo > UBound(ShieldAnimData) Then Escudo = 2
        If Casco = 0 Or Casco > UBound(CascoAnimData) Then Casco = 2
        If CartIndex <= 2 Or CartIndex > UBound(BodyData) Then .HasCart = False
        If BackpackIndex <= 2 Or BackpackIndex > UBound(BodyData) Then .HasBackpack = False
        
        .IHead = Head
        .iBody = Body
     '   If Not charindex = UserCharIndex Then
            .Head = HeadData(Head)
            .Body = BodyData(Body)
            .Arma = WeaponAnimData(Arma)
            .Escudo = ShieldAnimData(Escudo)
            .Casco = CascoAnimData(Casco)
            .Cart = BodyData(CartIndex)
            .Backpack = BodyData(BackpackIndex)
       ' End If
        
        .Heading = Heading
        
        'Reset moving stats
        If Not charindex = UserCharIndex Then
            .Moving = False
            .MoveOffsetX = 0
            .MoveOffsetY = 0
        Else
            UserPos.x = x
            UserPos.y = y
        End If
        
        'Update position
        .Pos.x = x ' - .scrollDirectionX
        .Pos.y = y ' - .scrollDirectionY
        
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
        
        ReDim .DialogEffects(0)
        
        .TimeCreated = FrameTime - RandomNumber(1, 10000)
        mascota.last_time = 0
    End With
    
    'Plot on map
    MapData(x, y).charindex = charindex

    
    Exit Sub

MakeChar_Err:
    Call RegistrarError(Err.Number, Err.Description, "TileEngine_Chars.MakeChar", Erl)
    Resume Next
    
End Sub

Public Sub Char_Move_by_Head(ByVal charindex As Integer, ByVal nHeading As E_Heading)

    On Error GoTo Char_Move_by_Head_Err
       
    Dim addx As Integer
    Dim addy As Integer
    Dim x    As Integer
    Dim y    As Integer
    Dim nX   As Integer
    Dim nY   As Integer
    
    With charlist(charindex)
        x = .Pos.x
        y = .Pos.y
        
        ' Dirección a mover
        Select Case nHeading
            Case E_Heading.NORTH: addy = -1
            Case E_Heading.EAST:  addx = 1
            Case E_Heading.south: addy = 1
            Case E_Heading.WEST:  addx = -1
        End Select
        
        nX = x + addx
        nY = y + addy
        
        MapData(nX, nY).charindex = charindex
        .Pos.x = nX
        .Pos.y = nY
        
        If MapData(x, y).charindex = charindex Then
            MapData(x, y).charindex = 0
        End If
        
        ' ---- Usar tamaño de tile configurable (antes 32 fijo)
        .MoveOffsetX = -1 * (TilePixelWidth * addx)
        .MoveOffsetY = -1 * (TilePixelHeight * addy)
        
        ' Forzar heading en escalera
        Dim newHeading As E_Heading
        If MapData(nX, nY).ObjGrh.GrhIndex = 26940 Or MapData(nX, nY).Trigger = ESCALERA Then
            newHeading = E_Heading.NORTH
        Else
            newHeading = nHeading
        End If
        
        ' Guardamos el heading anterior para conservar fase si cambia
        Dim oldHeading As E_Heading
        oldHeading = .Heading
        .scrollDirectionX = addx
        .scrollDirectionY = addy
        
        .Idle = False

        ' --- Si cambia de dirección mientras ya está moviéndose, preservamos fase
        If .Moving And (newHeading <> oldHeading) Then
            ' BODY
            If .Body.Walk(oldHeading).started > 0 Then
                Dim keepStarted As Long
                keepStarted = SyncGrhPhase(.Body.Walk(oldHeading), .Body.Walk(newHeading).GrhIndex)
                .Body.Walk(newHeading).started = keepStarted
            ElseIf .Body.Walk(newHeading).started = 0 Then
                .Body.Walk(newHeading).started = FrameTime
            End If
            ' WEAPON + SHIELD en fase con el cuerpo
            If .Arma.WeaponWalk(newHeading).started = 0 Then .Arma.WeaponWalk(newHeading).started = .Body.Walk(newHeading).started
            If .Escudo.ShieldWalk(newHeading).started = 0 Then .Escudo.ShieldWalk(newHeading).started = .Body.Walk(newHeading).started
            .Arma.WeaponWalk(newHeading).Loops = INFINITE_LOOPS
            .Escudo.ShieldWalk(newHeading).Loops = INFINITE_LOOPS
        End If

        ' Actualizamos el heading al final para usar el nuevo set arriba
        .Heading = newHeading
        
        If Not .Moving Then
            If .Muerto Then
                .Body = BodyData(CASPER_BODY)
            Else
                If .Body.BodyIndex <> .iBody Then
                    .Body = BodyData(.iBody)
                    .AnimatingBody = 0
                End If
                
                If .BackPack.BodyIndex <> .tmpBackPack Then
                    .BackPack = BodyData(.tmpBackPack)
                End If
            End If

            ' Start animations (solo al empezar a moverse)
            If .Body.Walk(.Heading).Started = 0 Then
                .Body.Walk(.Heading).Started = FrameTime
                .Arma.WeaponWalk(.Heading).Started = FrameTime
                .BackPack.Walk(.Heading).started = FrameTime
                .Escudo.ShieldWalk(.Heading).Started = FrameTime
                .Arma.WeaponWalk(.Heading).Loops = INFINITE_LOOPS
                .Escudo.ShieldWalk(.Heading).Loops = INFINITE_LOOPS
                .BackPack.Walk(.Heading).Loops = INFINITE_LOOPS
            End If
            
            .MovArmaEscudo = False
            .Moving = True
        End If

    End With
    
    If UserStats.Estado <> 1 Then Call DoPasosFx(CharIndex)
    
    If (nY < MinLimiteY) Or (nY > MaxLimiteY) Or (nX < MinLimiteX) Or (nX > MaxLimiteX) Then
        Call EraseChar(charindex)
    End If
    
    Exit Sub

Char_Move_by_Head_Err:
    Call RegistrarError(Err.Number, Err.Description, "TileEngine_Chars.Char_Move_by_Head", Erl)
    Resume Next
End Sub

Public Sub TranslateCharacterToPos(ByVal charindex As Integer, ByVal NewX As Integer, ByVal NewY As Integer, ByVal TranslationTime As Long)
On Error GoTo TranslateCharacterToPos_Err
    Dim TileX, TileY As Integer
    Dim DiffX, DiffY As Integer
100 With charlist(charindex)
102     TileX = .Pos.x
104     TileY = .Pos.y
106     If Not InMapBounds(TileX, TileY) Then Exit Sub
        Debug.Assert MapData(TileX, TileY).charindex = charindex
108     MapData(TileX, TileY).charindex = 0
110     DiffX = NewX - TileX
112     DiffY = NewY - TileY
        .Pos.x = NewX
        .Pos.y = NewY
114     MapData(NewX, NewY).charindex = charindex
116     .MoveOffsetX = -1 * (TilePixelWidth * DiffX)
118     .MoveOffsetY = -1 * (TilePixelHeight * DiffY)
120     .scrollDirectionX = Sgn(DiffX)
122     .scrollDirectionY = Sgn(DiffY)
        
124     If (NewY < MinLimiteY) Or (NewY > MaxLimiteY) Or (NewX < MinLimiteX) Or (NewX > MaxLimiteX) Then
126         Call EraseChar(charindex)
        End If
        .Moving = False
        .TranslationActive = True
        .TranslationTime = TranslationTime
        .TranslationStartTime = FrameTime
    End With
    Exit Sub
TranslateCharacterToPos_Err:
    Call RegistrarError(Err.Number, Err.Description, "TileEngine_Chars.TranslateCharacterToPos", Erl)
End Sub

Public Function IsLadderAt(ByVal TileX As Integer, ByVal TileY As Integer) As Boolean
    IsLadderAt = MapData(TileX, TileY).ObjGrh.GrhIndex = 26940
End Function

Public Sub Char_Move_by_Pos(ByVal charindex As Integer, ByVal nX As Integer, ByVal nY As Integer)
    
    On Error GoTo Char_Move_by_Pos_Err

    Dim x        As Integer
    Dim y        As Integer
    Dim addx     As Integer
    Dim addy     As Integer
    Dim nHeading As E_Heading
    
    With charlist(charindex)
        x = .Pos.x
        y = .Pos.y
        
        If Not InMapBounds(x, y) Then Exit Sub
        
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
        
        If nHeading = 0 Then Exit Sub
        
        .Pos.x = nX
        .Pos.y = nY
        
        .MoveOffsetX = -1 * (TilePixelWidth * addx)
        .MoveOffsetY = -1 * (TilePixelHeight * addy)

        ' === Guardar heading anterior ANTES de cambiarlo ===
        Dim oldHeading As E_Heading
        oldHeading = .Heading

        If IsLadderAt(nX, nY) Then
            .Heading = E_Heading.NORTH
        Else
            .Heading = nHeading
        End If
        
        .scrollDirectionX = Sgn(addx)
        .scrollDirectionY = Sgn(addy)

        .LastStep = FrameTime
        .Idle = False

        If Not .Moving Then
            ' --- Empezó a moverse recién ahora ---
            If .Muerto Then
                .Body = BodyData(CASPER_BODY)
            Else
                If .Body.BodyIndex <> .iBody Then
                    .Body = BodyData(.iBody)
                    .AnimatingBody = 0
                End If
            End If
            
            ' Start animations (solo si no estaban corriendo)
            If .Body.Walk(.Heading).Started = 0 Then
                .Body.Walk(.Heading).Started = FrameTime
                .Arma.WeaponWalk(.Heading).Started = FrameTime
                .Escudo.ShieldWalk(.Heading).Started = FrameTime
                .BackPack.Walk(.Heading).started = FrameTime
                .Arma.WeaponWalk(.Heading).Loops = INFINITE_LOOPS
                .Escudo.ShieldWalk(.Heading).Loops = INFINITE_LOOPS
                .BackPack.Walk(.Heading).Loops = INFINITE_LOOPS
                
            End If

            .MovArmaEscudo = False
            .Moving = True

        ElseIf .Heading <> oldHeading Then
            ' --- Ya venía moviéndose y cambió de dirección: preservar fase ---
            Dim keepStart As Long
            keepStart = SyncGrhPhase(.Body.Walk(oldHeading), .Body.Walk(.Heading).GrhIndex)
            If keepStart > 0 Then
                .Body.Walk(.Heading).started = keepStart
                ' Si necesitás acompasar arma/escudo, descomentá:
                'If .Arma.WeaponWalk(.Heading).started = 0 Then .Arma.WeaponWalk(.Heading).started = keepStart
                'If .Escudo.ShieldWalk(.Heading).started = 0 Then .Escudo.ShieldWalk(.Heading).started = keepStart
            End If
        End If

    End With

    If (nY < MinLimiteY) Or (nY > MaxLimiteY) Or (nX < MinLimiteX) Or (nX > MaxLimiteX) Then
        Call EraseChar(charindex)
    End If

    Exit Sub

Char_Move_by_Pos_Err:
    Call RegistrarError(Err.Number, Err.Description, "TileEngine_Chars.Char_Move_by_Pos", Erl)
    Resume Next
    
End Sub

Public Sub ApplySpeedingToChar(ByVal CharIndex As Integer)
    On Error Resume Next

    Dim rate As Single
    Dim h    As Long
    Dim gi   As Long
    Dim n    As Integer
    Dim total As Long
    Dim base  As Long
    Dim spd   As Long

    With charlist(CharIndex)
        rate = .Speeding
        If rate <= 0! Then rate = 1!

        ' ---- Cuerpo ----
        For h = LBound(.Body.Walk) To UBound(.Body.Walk)
            gi = .Body.Walk(h).GrhIndex
            If gi > 0 Then
                n = GrhData(gi).NumFrames
                total = GrhData(gi).speed
                If n > 0 Then
                    base = total \ n    ' ms por frame base (sin acelerar)
                Else
                    base = total
                End If
                If base <= 0 Then base = 1

                spd = CLng(base / rate) ' a mayor rate => frames más rápidos
                If spd < 40 Then spd = 40   ' clamps opcionales
                If spd > 220 Then spd = 220
                .Body.Walk(h).speed = spd
            End If
        Next h

        ' ---- Arma ----
        For h = LBound(.Arma.WeaponWalk) To UBound(.Arma.WeaponWalk)
            gi = .Arma.WeaponWalk(h).GrhIndex
            If gi > 0 Then
                n = GrhData(gi).NumFrames
                total = GrhData(gi).speed
                If n > 0 Then base = total \ n Else base = total
                If base <= 0 Then base = 1

                spd = CLng(base / rate)
                If spd < 40 Then spd = 40
                If spd > 220 Then spd = 220
                .Arma.WeaponWalk(h).speed = spd
            End If
        Next h

        ' ---- Escudo ----
        For h = LBound(.Escudo.ShieldWalk) To UBound(.Escudo.ShieldWalk)
            gi = .Escudo.ShieldWalk(h).GrhIndex
            If gi > 0 Then
                n = GrhData(gi).NumFrames
                total = GrhData(gi).speed
                If n > 0 Then base = total \ n Else base = total
                If base <= 0 Then base = 1

                spd = CLng(base / rate)
                If spd < 40 Then spd = 40
                If spd > 220 Then spd = 220
                .Escudo.ShieldWalk(h).speed = spd
            End If
        Next h
    End With
End Sub

Public Function EstaPCarea(ByVal CharIndex As Integer) As Boolean
    
    On Error GoTo EstaPCarea_Err
    

    With charlist(charindex).Pos
        EstaPCarea = .x > UserPos.x - MinXBorder And .x < UserPos.x + MinXBorder And .y > UserPos.y - MinYBorder And .y < UserPos.y + MinYBorder

    End With

    
    Exit Function

EstaPCarea_Err:
    Call RegistrarError(Err.Number, Err.Description, "TileEngine_Chars.EstaPCarea", Erl)
    Resume Next
    
End Function

Public Function EstaEnArea(ByVal x As Integer, ByVal y As Integer) As Boolean
    
    On Error GoTo EstaEnArea_Err
    
    EstaEnArea = x > UserPos.x - MinXBorder And x < UserPos.x + MinXBorder And y > UserPos.y - MinYBorder And y < UserPos.y + MinYBorder

    
    Exit Function

EstaEnArea_Err:
    Call RegistrarError(Err.Number, Err.Description, "TileEngine_Chars.EstaEnArea", Erl)
    Resume Next
    
End Function

Public Function Char_Check(ByVal char_index As Integer) As Boolean
    
    On Error GoTo Char_Check_Err

    'check char_index
    If char_index > 0 And char_index <= LastChar Then
        Char_Check = (charlist(char_index).Heading > 0)

    End If
    
    
    Exit Function

Char_Check_Err:
    Call RegistrarError(Err.Number, Err.Description, "TileEngine_Chars.Char_Check", Erl)
    Resume Next
    
End Function

Public Function Char_FX_Group_Next_Open(ByVal char_index As Integer) As Integer

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

Public Sub Char_Dialog_Set(ByVal char_index As Integer, ByVal char_dialog As String, ByVal char_dialog_color As Long, _
                           ByVal char_dialog_life As Byte, ByVal Sube As Byte, Optional ByVal font_index As Integer = 1, _
                           Optional ByVal IsSpell As Boolean = False, Optional ByVal MinChatTime As Integer = 0, Optional ByVal MaxChatTime As Integer = 0)
    
    On Error GoTo Char_Dialog_Set_Err
    
    Dim Slot As Integer
    Dim i    As Long
    Slot = BinarySearch(char_index)
    If Slot < 0 Then
        If dialogCount = MAX_DIALOGS Then Exit Sub  'Out of space! Should never happen....
        'We need to add it. Get insertion index and move list    backwards.
        Slot = Not Slot
        For i = dialogCount To Slot + 1 Step -1
            dialogs(i) = dialogs(i - 1)
        Next i
        dialogCount = dialogCount + 1
    End If
    With dialogs(Slot)
        If Not IsSpell Then
            Dim ElapsedTime As Long
            ElapsedTime = FrameTime - .startTime
            If .MinChatTime > ElapsedTime Then
                Exit Sub
            End If
        End If
        If Char_Check(char_index) Then
            charlist(char_index).dialog = char_dialog
            charlist(char_index).dialog_color = char_dialog_color
            charlist(char_index).dialog_life = char_dialog_life
            charlist(char_index).dialog_font_index = font_index
            charlist(char_index).dialog_scroll = True
            charlist(char_index).dialog_offset_counter_y = -(IIf(BodyData(charlist(char_index).iBody).HeadOffset.y = 0, -32, BodyData(charlist(char_index).iBody).HeadOffset.y) / 2)
            charlist(char_index).AlphaText = 255
        End If
    
        .startTime = FrameTime
        .MinChatTime = MinChatTime
        If IsSpell Then
             If MaxChatTime > 0 Then
                .lifeTime = MaxChatTime
             Else
                .lifeTime = 3500
             End If
        Else
            .lifeTime = 3000 + MS_PER_CHAR * Len(char_dialog)
        End If
        .charindex = char_index
    End With
    Exit Sub

Char_Dialog_Set_Err:
    Call RegistrarError(Err.Number, Err.Description, "TileEngine_Chars.Char_Dialog_Set", Erl)
    Resume Next
    
End Sub


Public Sub Char_Dialog_Remove(ByVal char_index As Integer, ByVal Index As Integer)
    
    On Error GoTo Char_Dialog_Remove_Err
    

    If char_index = 0 Then Exit Sub

    If charlist(char_index).AlphaText > 0 Then
        charlist(char_index).AlphaText = charlist(char_index).AlphaText - (scroll_dialog_pixels_per_frame * timerTicksPerFrame)
        Exit Sub

    End If

    Dim Slot As Integer

    Dim i    As Long
    
    Slot = BinarySearch(char_index)
    
    If Slot < 0 Then Exit Sub
    
    For i = Slot To MAX_DIALOGS - 2
        dialogs(i) = dialogs(i + 1)
    Next i
    
    dialogCount = dialogCount - 1
    
    If Char_Check(char_index) Then
        charlist(char_index).dialog = ""
        charlist(char_index).dialog_color = 0
        charlist(char_index).dialog_life = 0

    End If

    
    Exit Sub

Char_Dialog_Remove_Err:
    Call RegistrarError(Err.Number, Err.Description, "TileEngine_Chars.Char_Dialog_Remove", Erl)
    Resume Next
    
End Sub

Public Sub SetCharacterFx(ByVal charindex As Integer, ByVal fX As Integer, ByVal Loops As Integer)
    
    On Error GoTo SetCharacterFx_Err
    

    If fX = 0 Then Exit Sub

    'Sets an FX to the character.
    
    Dim indice As Byte

    With charlist(charindex)
    
        indice = Char_FX_Group_Next_Open(charindex)
        
        .FxList(indice).FxIndex = fX
        
        Call InitGrh(.FxList(indice), FxData(fX).Animacion, , Loops)
            
    End With

    
    Exit Sub

SetCharacterFx_Err:
    Call RegistrarError(Err.Number, Err.Description, "TileEngine_Chars.SetCharacterFx", Erl)
    Resume Next
    
End Sub

Public Sub SetCharacterDialogFx(ByVal charindex As Integer, ByVal Text As String, Color As RGBA)

    With charlist(charindex)
        
        Dim Index As Integer
        
        If UBound(.DialogEffects) > 0 Then
        
            For Index = 1 To UBound(.DialogEffects)
                If .DialogEffects(Index).Text = vbNullString Then
                    Exit For
                End If
            Next
            
            If Index > UBound(.DialogEffects) Then
                ReDim .DialogEffects(1 To UBound(.DialogEffects) + 1)
            End If
        
        Else
            ReDim .DialogEffects(1)
            
            Index = 1
        End If
        
        With .DialogEffects(Index)
        
            .Color = Color
            .Start = FrameTime
            .Text = Text
        
        End With
        
    End With
    
End Sub

Public Function Get_PixelY_Of_Char(ByVal char_index As Integer) As Integer
    
    On Error GoTo Get_PixelY_Of_Char_Err
    'Make sure it's a legal char_index
    If Char_Check(char_index) Then
        Get_PixelY_Of_Char = (charlist(char_index).Pos.y - 2 - UserPos.y) * 32 + frmMain.renderer.ScaleWidth / 2
        Get_PixelY_Of_Char = Get_PixelY_Of_Char - 16 + gameplay_render_offset.y

    End If

    
    Exit Function

Get_PixelY_Of_Char_Err:
    Call RegistrarError(Err.Number, Err.Description, "TileEngine_Chars.Get_PixelY_Of_Char", Erl)
    Resume Next
    
End Function

Public Function Get_Pixelx_Of_Char(ByVal char_index As Integer) As Integer
    
    On Error GoTo Get_Pixelx_Of_Char_Err
    'Make sure it's a legal char_index
    If Char_Check(char_index) Then
        Get_Pixelx_Of_Char = (charlist(char_index).Pos.x - UserPos.x) * 32 + frmMain.renderer.ScaleWidth / 2
        Get_Pixelx_Of_Char = Get_Pixelx_Of_Char + gameplay_render_offset.x

    End If

    
    Exit Function

Get_Pixelx_Of_Char_Err:
    Call RegistrarError(Err.Number, Err.Description, "TileEngine_Chars.Get_Pixelx_Of_Char", Erl)
    Resume Next
    
End Function

Public Function Get_Pixelx_Of_XY(ByVal x As Byte) As Integer
    'Make sure it's a legal char_index
    
    On Error GoTo Get_Pixelx_Of_XY_Err
    
    Get_Pixelx_Of_XY = (x - UserPos.x) * 32 + frmMain.renderer.ScaleWidth / 2
    Get_Pixelx_Of_XY = Get_Pixelx_Of_XY

    
    Exit Function

Get_Pixelx_Of_XY_Err:
    Call RegistrarError(Err.Number, Err.Description, "TileEngine_Chars.Get_Pixelx_Of_XY", Erl)
    Resume Next
    
End Function

Public Function Get_PixelY_Of_XY(ByVal y As Byte) As Integer
    'Make sure it's a legal char_index
    
    On Error GoTo Get_PixelY_Of_XY_Err
    
    Get_PixelY_Of_XY = (y - 2 - UserPos.y) * 32 + frmMain.renderer.ScaleWidth / 2
    Get_PixelY_Of_XY = Get_PixelY_Of_XY - 16

    
    Exit Function

Get_PixelY_Of_XY_Err:
    Call RegistrarError(Err.Number, Err.Description, "TileEngine_Chars.Get_PixelY_Of_XY", Erl)
    Resume Next
    
End Function



Public Function SyncGrhPhase(ByRef Grh As Grh, ByVal newGrhIndex As Long) As Long
    Dim oldNum As Long, elapsed As Long, phase As Long
    If Grh.started <= 0 Then SyncGrhPhase = FrameTime: Exit Function
    oldNum = GrhData(Grh.GrhIndex).NumFrames
    If oldNum <= 0 Then SyncGrhPhase = FrameTime: Exit Function
    elapsed = Fix((FrameTime - Grh.started) / Grh.speed)
    phase = elapsed Mod oldNum
    SyncGrhPhase = FrameTime - (phase * Grh.speed)
End Function

