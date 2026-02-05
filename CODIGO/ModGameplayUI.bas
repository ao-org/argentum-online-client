Attribute VB_Name = "ModGameplayUI"
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

Public Sub SetupGameplayUI()
    frmMain.shapexy.Left = 1200
    frmMain.shapexy.Top = 1200
    frmMain.shapexy.BackColor = RGB(170, 0, 0)
    frmMain.NombrePJ.Caption = userName
    ' Detect links in console
    Call EnableURLDetect(frmMain.RecTxt.hWnd, frmMain.hWnd)
    Call Make_Transparent_Richtext(frmMain.RecTxt.hWnd)
    ' Removemos la barra de titulo pero conservando el caption para la barra de tareas
    Call Form_RemoveTitleBar(frmMain)
    frmMain.panel.Picture = LoadInterface("centroinventario.bmp")
    frmMain.picInv.visible = True
    frmMain.picHechiz.visible = False
    frmMain.cmdlanzar.visible = False
    frmMain.imgSpellInfo.visible = False
    frmMain.cmdMoverHechi(0).visible = False
    frmMain.cmdMoverHechi(1).visible = False
    Call frmMain.Inventario.ReDraw
    frmMain.Left = 0
    frmMain.Top = 0
    frmMain.Width = D3DWindow.BackBufferWidth * screen.TwipsPerPixelX
    frmMain.Height = D3DWindow.BackBufferHeight * screen.TwipsPerPixelY
    frmMain.visible = True
    ActiveInventoryTab = eInventory
    Call LoadHotkeys
End Sub

Public Sub OnClick(ByVal MouseButton As Long, ByVal MouseShift As Long)
    On Error GoTo OnClick_Err
    If pausa Then Exit Sub
    If IsGameDialogOpen Then Exit Sub
    If mascota.visible Then
        If Sqr((mouseX - mascota.PosX) ^ 2 + (mouseY - mascota.PosY) ^ 2) < 30 Then
            mascota.dialog = ""
        End If
    End If
    If cartel_visible Then
        If mouseX > 50 And mouseY > 478 And mouseX < 671 And mouseY < 585 Then
            If tutorial_index > 0 Then
                Call nextCartel
            Else
                Call cerrarCartel
            End If
        End If
    End If
    Dim MouseAction As e_MouseAction
    Select Case MouseButton
        Case vbLeftButton
            MouseAction = ACCION1
        Case vbRightButton
            MouseAction = ACCION2
        Case vbMiddleButton
            MouseAction = ACCION3
        Case Else
            Exit Sub
    End Select
    If MouseAction = e_MouseAction.eThrowOrLook Then
        If Not Comerciando Then
            If MouseShift = 0 Then
                If UsingSkill = 0 Or frmMain.MacroLadder.enabled Then
                    Call CountPacketIterations(packetControl(ClientPacketID.eLeftClick), 150)
                    Call WriteLeftClick(tX, tY)
                Else
                    Dim SendSkill As Boolean
                    If UsingSkill = magia Then
                        If ModoHechizos = BloqueoLanzar Then
                            SendSkill = IIf((mouseX >= frmMain.renderer.ScaleLeft And mouseX <= 736 + frmMain.renderer.ScaleLeft And mouseY >= frmMain.renderer.ScaleTop And _
                                    mouseY <= frmMain.renderer.ScaleTop + 608), True, False)
                            If Not SendSkill Then
                                Exit Sub
                            End If
                            Call MainTimer.Restart(TimersIndex.CastAttack)
                            Call MainTimer.Restart(TimersIndex.CastSpell)
                        Else
                            If MainTimer.Check(TimersIndex.AttackSpell, False) Then
                                If MainTimer.Check(TimersIndex.CastSpell) Then
                                    SendSkill = IIf((mouseX >= frmMain.renderer.ScaleLeft And mouseX <= 736 + frmMain.renderer.ScaleLeft And mouseY >= frmMain.renderer.ScaleTop _
                                            And mouseY <= frmMain.renderer.ScaleTop + 608), True, False)
                                    If Not SendSkill Then
                                        Exit Sub
                                    End If
                                    Call MainTimer.Restart(TimersIndex.CastAttack)
                                ElseIf ModoHechizos = SinBloqueo Then
                                    SendSkill = IIf((mouseX >= frmMain.renderer.ScaleLeft And mouseX <= 736 + frmMain.renderer.ScaleLeft And mouseY >= frmMain.renderer.ScaleTop _
                                            And mouseY <= frmMain.renderer.ScaleTop + 608), True, False)
                                    If Not SendSkill Then
                                        Exit Sub
                                    End If
                                    With FontTypes(FontTypeNames.FONTTYPE_TALK)
                                        Call ShowConsoleMsg(JsonLanguage.Item("MENSAJE_LANZAMIENTO_RAPIDO"), .red, .green, .blue, .bold, .italic) ' MENSAJE_LANZAMIENTO_RAPIDO=No puedes lanzar hechizos tan rápido.
                                    End With
                                Else
                                    Exit Sub
                                End If
                            ElseIf ModoHechizos = SinBloqueo Then
                                SendSkill = IIf((mouseX >= frmMain.renderer.ScaleLeft And mouseX <= 736 + frmMain.renderer.ScaleLeft And mouseY >= frmMain.renderer.ScaleTop And _
                                        mouseY <= frmMain.renderer.ScaleTop + 608), True, False)
                                If Not SendSkill Then
                                    Exit Sub
                                End If
                                With FontTypes(FontTypeNames.FONTTYPE_TALK)
                                    Call ShowConsoleMsg(JsonLanguage.Item("MENSAJE_ATAQUE_RAPIDO_GOLPE"), .red, .green, .blue, .bold, .italic) ' MENSAJE_ATAQUE_RAPIDO_GOLPE=No puedes lanzar tan rápido después de un golpe.
                                End With
                            Else
                                Exit Sub
                            End If
                        End If
                    End If
                    'Splitted because VB isn't lazy!
                    If UsingSkill = Proyectiles Then
                        If MainTimer.Check(TimersIndex.AttackSpell, False) Then
                            If MainTimer.Check(TimersIndex.CastAttack, False) Then
                                If MainTimer.Check(TimersIndex.Arrows) Then
                                    SendSkill = True
                                    Call MainTimer.Restart(TimersIndex.Attack) ' Prevengo flecha-golpe
                                    Call MainTimer.Restart(TimersIndex.CastSpell) ' flecha-hechizo
                                End If
                            End If
                        End If
                    End If
                    'Splitted because VB isn't lazy!
                    If (UsingSkill = Robar Or UsingSkill = Domar Or UsingSkill = Grupo Or UsingSkill = MarcaDeClan Or UsingSkill = MarcaDeGM) Then
                        If MainTimer.Check(TimersIndex.CastSpell) Then
                            If UsingSkill = MarcaDeGM Then
                                Dim Pos As Integer
                                If MapData(tX, tY).charindex <> 0 Then
                                    Pos = InStr(charlist(MapData(tX, tY).charindex).nombre, "<")
                                    If Pos = 0 Then Pos = LenB(charlist(MapData(tX, tY).charindex).nombre) + 2
                                    frmPanelgm.cboListaUsus.text = Left$(charlist(MapData(tX, tY).charindex).nombre, Pos - 2)
                                End If
                            Else
                                SendSkill = True
                            End If
                        End If
                    End If
                    If (UsingSkill = eSkill.Alquimia Or UsingSkill = eSkill.TargetableItem) Then
                        If MainTimer.Check(TimersIndex.CastSpell) Then
                            Call WriteWorkLeftClick(tX, tY, UsingSkill)
                            Call FormParser.Parse_Form(GetGameplayForm)
                            If CursoresGraficos = 0 Then
                                GetGameplayForm.MousePointer = vbDefault
                            End If
                        End If
                    Else
                        If UsingSkill = eSkill.Talar Or UsingSkill = eSkill.Mineria Or UsingSkill = eSkill.Pescar Or UsingSkill = eSkill.Smelting Then
                            Call WriteStartAutomatedAction(tX, tY, UsingSkill)
                        End If
                    End If
                    If SendSkill Then
                        If UsingSkill = eSkill.magia Then
                            If ComprobarPosibleMacro(mouseX, mouseY) Then
                                Call WriteWorkLeftClick(tX + RandomNumber(-2, 2), tY + RandomNumber(-2, 2), UsingSkill)
                            Else
                                Call WriteWorkLeftClick(tX, tY, UsingSkill)
                            End If
                        Else
                            Call WriteWorkLeftClick(tX, tY, UsingSkill)
                        End If
                    End If
                    Call FormParser.Parse_Form(GetGameplayForm)
                    If CursoresGraficos = 0 Then
                        GetGameplayForm.MousePointer = vbDefault
                    End If
                    UsaLanzar = False
                    UsingSkill = 0
                End If
            Else
                Call WriteWarpChar("YO", UserMap, tX, tY)
            End If
            If cartel Then cartel = False
        End If
    ElseIf MouseAction = e_MouseAction.eInteract Then
        Call WriteDoubleClick(tX, tY)
    ElseIf MouseAction = e_MouseAction.eAttack Then
        If UserDescansar Or UserMeditar Then Exit Sub
        If MainTimer.Check(TimersIndex.CastAttack, False) Then
            If MainTimer.Check(TimersIndex.Attack) Then
                Call MainTimer.Restart(TimersIndex.AttackSpell)
                Call WriteAttack
            End If
        End If
    ElseIf MouseAction = e_MouseAction.eWhisper Then
        Dim charindex As Integer
        charindex = MapData(tX, tY).charindex
        If charindex = 0 And tY < YMaxMapSize Then
            charindex = MapData(tX, tY + 1).charindex
        End If
        If charindex <> 0 Then
            If charlist(charindex).nombre <> charlist(UserCharIndex).nombre Then
                If charlist(charindex).EsNpc = False Then
                    frmMain.SendTxt.text = "\" & charlist(charindex).nombre & " "
                    If frmMain.SendTxtCmsg.visible = False Then
                        frmMain.SendTxt.visible = True
                        frmMain.SendTxt.SetFocus
                        frmMain.SendTxt.SelStart = Len(frmMain.SendTxt.text)
                    End If
                End If
            End If
        End If
    End If
    Exit Sub
OnClick_Err:
    Call RegistrarError(Err.Number, Err.Description, "ModGameplayUi.OnClick", Erl)
    Resume Next
End Sub

Public Sub HandleQuestionResponse(ByVal Result As Boolean)
    If PreguntaLocal Then
        If Result Then
            Select Case PreguntaNUM
                Case 1 '¿Destruir item?
                    Call WriteDrop(DestItemSlot, DestItemCant)
                Case 2 ' Denunciar
                    Call WriteDenounce(targetName)
            End Select
        Else
            Select Case PreguntaNUM
                Case 1
                    DestItemSlot = 0
                    DestItemCant = 0
            End Select
        End If
    Else
        Call WriteResponderPregunta(Result)
    End If
    Pregunta = False
    PreguntaLocal = False
End Sub

Public Sub HandleGameplayAreaMouseUp(ByVal Button As Integer, _
                                     ByVal x As Integer, _
                                     ByVal y As Integer, _
                                     ByVal FormTop As Long, _
                                     ByVal FormLeft As Long, _
                                     ByVal FormHeight As Long, _
                                     ByRef GameplayArea As Rect)
    clicX = x
    clicY = y
    
    Dim MouseAction As e_MouseAction
    Select Case Button
        Case vbLeftButton:  MouseAction = ACCION1
        Case vbRightButton: MouseAction = ACCION2
        Case vbMiddleButton: MouseAction = ACCION3
        Case Else: Exit Sub
    End Select
    
    Select Case MouseAction
    
        Case e_MouseAction.eThrowOrLook
            If HandleMouseInput(x, y) Then
            ElseIf Pregunta Then
                If x >= 419 And x <= 433 And y >= 243 And y <= 260 Then
                    Call HandleQuestionResponse(False)
                    Exit Sub
                ElseIf x >= 443 And x <= 458 And y >= 243 And y <= 260 Then
                    Call HandleQuestionResponse(True)
                    Exit Sub
                End If
            End If
        Case e_MouseAction.eInteract
            Call ShowInteractionMenu(FormTop, FormLeft, FormHeight, x, y, GameplayArea)
    End Select
End Sub

Public Sub HandleChatMsg(ByVal InputText As String)
    Dim str2 As String
    Dim str1 As String
    If LenB(InputText) <> 0 Then
        If Left$(InputText, 1) = "/" Then
            If UCase$(Left$(InputText, 7)) = "/GRUPO " Then
                SendingType = 5
            ElseIf UCase$(Left$(InputText, 6)) = "/CMSG " Then
                SendingType = 4
            ElseIf UCase$(Left$(InputText, 6)) = "/GRMG " Then
                SendingType = 6
            ElseIf UCase$(Left$(InputText, 6)) = "/RMSG " Then
                SendingType = 8
            Else
                SendingType = 1
            End If
            If InputText <> "" Then Call ParseUserCommand(InputText)
            'Shout
        ElseIf Left$(InputText, 1) = "-" Then
            If Right$(InputText, Len(InputText) - 1) <> "" Then Call ParseUserCommand("-" & Right$(InputText, Len(InputText) - 1))
            SendingType = 2
            'Global
        ElseIf Left$(InputText, 1) = ";" Then
            If Right$(InputText, Len(InputText) - 1) <> "" Then Call ParseUserCommand("/CONSOLA " & Right$(InputText, Len(InputText) - 1))
            sndPrivateTo = ""
        ElseIf Left$(InputText, 1) = "/RMSG" Then
            If Right$(InputText, Len(InputText) - 1) <> "" Then Call ParseUserCommand("/RMSG " & Right$(InputText, Len(InputText) - 1))
            SendingType = 8
            sndPrivateTo = ""
            'Faccion
        ElseIf Left$(InputText, 1) = "/FMSG" Then
            If Right$(InputText, Len(InputText) - 1) <> "" Then Call ParseUserCommand("/FMSG " & Right$(InputText, Len(InputText) - 1))
            SendingType = 9
            'Privado
        ElseIf Left$(InputText, 1) = "\" Then
            Dim mensaje As String
            str1 = Right$(InputText, Len(InputText) - 1)
            str2 = ReadField(1, str1, 32)
            mensaje = Right$(InputText, Len(str1) - Len(str2) - 1)
            sndPrivateTo = str2
            SendingType = 3
            If str1 <> "" Then Call WriteWhisper(sndPrivateTo, mensaje)
            'Say
        Else
            If InputText <> "" Then Call ParseUserCommand(InputText)
            SendingType = 1
            sndPrivateTo = ""
        End If
    Else
        SendingType = 1
        sndPrivateTo = ""
    End If
End Sub

Public Sub UseSelectInvItem()
End Sub

Public Sub SetInvItem(ByVal Slot As Byte, _
                      ByVal ObjIndex As Integer, _
                      ByVal Amount As Integer, _
                      ByVal Equipped As Byte, _
                      ByVal GrhIndex As Long, _
                      ByVal ObjType As Integer, _
                      ByVal MaxHit As Integer, _
                      ByVal MinHit As Integer, _
                      ByVal Def As Integer, _
                      ByVal value As Single, _
                      ByVal Name As String, _
                      ByVal CanUse As Byte, _
                      ByVal ElementalTags As Long, _
                      ByVal IsBindable As Byte)
    If Slot < 1 Or Slot > UBound(UserInventory.Slots) Then Exit Sub
    With UserInventory.Slots(Slot)
        .Amount = Amount
        .Def = Def
        .Equipped = Equipped
        .GrhIndex = GrhIndex
        .MaxHit = MaxHit
        .MinHit = MinHit
        .Name = Name
        .ObjIndex = ObjIndex
        .ObjType = ObjType
        .Valor = value
        .PuedeUsar = CanUse
        .ElementalTags = ElementalTags
        .IsBindable = IsBindable > 0
    End With
    Call frmMain.Inventario.SetItem(Slot, ObjIndex, Amount, Equipped, GrhIndex, ObjType, MaxHit, MinHit, Def, value, Name, ElementalTags, CanUse)
End Sub

Public Sub SelectItemSlot(ByVal Slot As Integer)
    UserInventory.SelectedSlot = Slot
End Sub

Public Function GetSelectedItemSlot() As Integer
    GetSelectedItemSlot = frmMain.Inventario.SelectedItem
End Function

Public Function IsItemSelected() As Boolean
    IsItemSelected = frmMain.Inventario.IsItemSelected
End Function

Public Sub UseItemKey()
    If Not MainTimer.Check(TimersIndex.AttackUse, False) Then Exit Sub
    Call CountPacketIterations(packetControl(ClientPacketID.eUseItemU), 100)
    If frmMain.Inventario.IsItemSelected Then
        Call WriteUseItemU(frmMain.Inventario.SelectedItem)
    End If
End Sub

Public Sub UserItemClick()
    If frmCarp.visible Or frmHerrero.visible Or frmComerciar.visible Or frmBancoObj.visible Then Exit Sub
    If pausa Then Exit Sub
    If UserMeditar Then Exit Sub
    If frmMain.macrotrabajo.enabled Then frmMain.DesactivarMacroTrabajo
    If Not IsItemSelected Then Exit Sub
    Call UserOrEquipItem(frmMain.Inventario.SelectedItem, frmMain.Inventario.Equipped(frmMain.Inventario.SelectedItem), frmMain.Inventario.ObjIndex( _
            frmMain.Inventario.SelectedItem))
End Sub

Public Sub UserOrEquipItem(ByVal Slot As Integer, ByVal Equipped As Boolean, ByVal ObjIndex As Integer)
    Dim ObjType As Byte
    ObjType = ObjData(ObjIndex).ObjType
    Select Case ObjType
        Case eObjType.otArmadura, eObjType.otESCUDO, eObjType.otmagicos, eObjType.otFlechas, eObjType.otCASCO, eObjType.otAnillos, eObjType.otManchas
            If Not Equipped Then
                Call WriteEquipItem(Slot)
            Else
                Call WriteUseItem(Slot)
            End If
        Case eObjType.otWeapon
            If ObjData(ObjIndex).proyectil = 1 And Equipped Then
                Call WriteUseItem(Slot)
            Else
                If Not Equipped Then
                    Call WriteEquipItem(Slot)
                End If
            End If
        Case eObjType.OtHerramientas
            If Equipped Then
                Call WriteUseItem(Slot)
            Else
                If Not Equipped Then
                    Call WriteEquipItem(Slot)
                End If
            End If
        Case eObjType.OtDonador
            If Not Equipped Then
                Call WriteEquipItem(Slot)
            End If
        Case Else
            Call CountPacketIterations(packetControl(ClientPacketID.eUseItem), 180)
            Call WriteUseItem(Slot)
    End Select
End Sub

Public Sub HandleKeyUp(KeyCode As Integer, Shift As Integer)
    If Not IsInputFocus Then
        If Not IsDialogOpen Then
            If Accionar(KeyCode) Then
                Exit Sub
            ElseIf KeyCode = BindKeys(e_KeyAction.eSendText).KeyCode Then
                Call OpenChatInput
            ElseIf KeyCode = vbKeyDelete Then
                Call OpenAndFocusClanChat
            ElseIf KeyCode = vbKeyEscape And Not UserSaliendo Then
                Call HandleEsc
            ElseIf KeyCode = 27 And UserSaliendo Then
                Call WriteCancelarExit
            ElseIf KeyCode = 80 And PescandoEspecial Then
                Call IntentarObtenerPezEspecial
            ElseIf KeyCode = vbKeyF1 Then
                Call ParseUserCommand("/SM")
                Call ParseUserCommand("/IRA " & targetName)
            End If
        End If
    Else
        Call FocusInput
    End If
End Sub

Public Sub HandleEsc()
    frmCerrar.Show , frmMain
End Sub

Public Function IsDialogOpen() As Boolean
    IsDialogOpen = pausa Or frmComerciar.visible Or frmComerciarUsu.visible Or frmBancoObj.visible Or frmGoliath.visible Or IsGameDialogOpen
End Function

Public Function IsInputFocus() As Boolean
    IsInputFocus = frmMain.SendTxt.visible Or frmMain.SendTxtCmsg.visible
End Function

Public Sub OpenAndFocusClanChat()
    If Not frmMain.SendTxt.visible Then
        frmMain.SendTxtCmsg.visible = True
        frmMain.SendTxtCmsg.SetFocus
    End If
    Call DialogosClanes.toggle_dialogs_visibility(True)
End Sub

Public Sub OpenChatInput()
    If Not frmCantidad.visible Then
        Call frmMain.CompletarEnvioMensajes
        StartOpenChatTime = GetTickCount
        frmMain.SendTxt.visible = True
        frmMain.SendTxt.SetFocus
    End If
End Sub

Public Sub FocusInput()
    If frmMain.SendTxt.visible Then
        frmMain.SendTxt.SetFocus
    End If
    If frmMain.SendTxtCmsg.visible Then
        frmMain.SendTxtCmsg.SetFocus
    End If
End Sub

Public Function GetGameplayForm() As Form
    Set GetGameplayForm = frmMain
End Function

Public Sub UseSpell(ByVal SpellSlot As Byte, ByVal SpellName As String)
    If pausa Then Exit Sub
    TempTick = GetTickCount And &H7FFFFFFF
    If TempTick - iClickTick < IntervaloEntreClicks And Not iClickTick = 0 And LastMacroButton <> tMacroButton.Lanzar Then
        Call WriteLogMacroClickHechizo(tMacro.Coordenadas)
    End If
    iClickTick = TempTick
    LastMacroButton = tMacroButton.Lanzar
    If SpellName <> "(Vacío)" Then
        If UserStats.estado = 1 Then
            With FontTypes(FontTypeNames.FONTTYPE_INFO)
                Call ShowConsoleMsg(JsonLanguage.Item("MENSAJE_ESTAS_MUERTO"), .red, .green, .blue, .bold, .italic)
            End With
        Else
            If ModoHechizos = BloqueoLanzar Then
                If Not MainTimer.Check(TimersIndex.AttackSpell, False) Or Not MainTimer.Check(TimersIndex.CastSpell, False) Then
                    Exit Sub
                End If
            End If
            Call WriteCastSpell(SpellSlot)
            UsaMacro = True
            UsaLanzar = True
        End If
    End If
End Sub

Public Sub UpdateMapPos()
    Call frmMain.SetMinimapPosition(0, UserPos.x, UserPos.y)
    frmMain.Coord.Caption = UserMap & "-" & UserPos.x & "-" & UserPos.y
    If frmMapaGrande.visible Then
        Call frmMapaGrande.ActualizarPosicionMapa
    End If
End Sub

Public Sub RequestSkills()
    If pausa Or tutorial_index > 0 Then Exit Sub
    If MostrarTutorial And tutorial_index <= 0 Then
        If tutorial(4).Activo = 1 Then
            tutorial_index = e_tutorialIndex.TUTORIAL_SkillPoints
            'TUTORIAL MAPA INSEGURO
            Call mostrarCartel(tutorial(tutorial_index).titulo, tutorial(tutorial_index).textos(1), tutorial(tutorial_index).Grh, -1, &H164B8A, , , False, 100, 479, 100, 535, _
                    640, 530, 64, 64)
            Exit Sub
        End If
    End If
    LlegaronSkills = True
    Call WriteRequestSkills
End Sub

Public Function IsUsableItem(ByRef ItemData As ObjDatas) As Boolean
    IsUsableItem = ItemData.ObjType = eObjType.otWeapon Or ItemData.ObjType = eObjType.otPociones Or ItemData.ObjType = eObjType.OtHerramientas Or ItemData.ObjType = _
            eObjType.otInstrumentos Or ItemData.ObjType = eObjType.OtCofre
End Function

Public Sub EquipSelectedItem()
    If frmMain.Inventario.IsItemSelected Then Call WriteEquipItem(frmMain.Inventario.SelectedItem)
End Sub

Public Sub OpenCreateObjectMenu()
    On Error GoTo createObj_Click_Err
    Dim i As Long
    For i = 1 To NumOBJs
        If ObjData(i).Name <> "" Then
            Dim subelemento As ListItem
            Set subelemento = FrmObjetos.ListView1.ListItems.Add(, , ObjData(i).Name)
            subelemento.SubItems(1) = i
        End If
    Next i
    GetGameplayForm().SetFocus
    FrmObjetos.Show , GetGameplayForm
    Exit Sub
createObj_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmMain.createObj_Click", Erl)
    Resume Next
End Sub

Public Sub SelectInventoryTab()
    ActiveInventoryTab = eInventory
    TempTick = GetTickCount And &H7FFFFFFF
    If TempTick - iClickTick < IntervaloEntreClicks And Not iClickTick = 0 And LastMacroButton <> tMacroButton.Inventario Then
        Call WriteLogMacroClickHechizo(tMacro.Coordenadas)
    End If
    iClickTick = TempTick
    LastMacroButton = tMacroButton.Inventario
    If Seguido = 1 Then
        Call WriteNotifyInventarioHechizos(1, hlst.ListIndex, hlst.Scroll)
    End If
End Sub

Public Sub SelectSpellTab()
    ActiveInventoryTab = eSpellList
    TempTick = GetTickCount And &H7FFFFFFF
    If TempTick - iClickTick < IntervaloEntreClicks And Not iClickTick = 0 And LastMacroButton <> tMacroButton.Hechizos Then
        Call WriteLogMacroClickHechizo(tMacro.Coordenadas)
    End If
    iClickTick = TempTick
    LastMacroButton = tMacroButton.Hechizos
    If Seguido = 1 Then
        Call WriteNotifyInventarioHechizos(2, hlst.ListIndex, hlst.Scroll)
    End If
End Sub

Public Sub GetMinimapPosition(ByRef x As Single, ByRef y As Single)
    x = x * (100 - 2 * HalfWindowTileWidth - 4) / 100 + HalfWindowTileWidth + 2
    y = y * (100 - 2 * HalfWindowTileHeight - 4) / 100 + HalfWindowTileHeight + 2
End Sub

Public Sub RequestMeditate()
    If UserStats.minman = UserStats.maxman Then Exit Sub
    If UserStats.estado = 1 Then
        With FontTypes(FontTypeNames.FONTTYPE_INFO)
            Call ShowConsoleMsg(JsonLanguage.Item("MENSAJE_ESTAS_MUERTO"), .red, .green, .blue, .bold, .italic) ' MENSAJE_ESTAS_MUERTO=¡Estás muerto!
        End With
        Exit Sub
    End If
    Call WriteMeditate
End Sub

Public Sub SetHotkey(ByVal Index As Integer, ByVal LastKnownSlot As Integer, ByVal HotkeyType As e_HotkeyType, ByVal HotkeySlot As Integer)
    HotkeyList(HotkeySlot).Index = Index
    HotkeyList(HotkeySlot).LastKnownSlot = LastKnownSlot
    HotkeyList(HotkeySlot).Type = HotkeyType
    Call SaveHotkey(Index, LastKnownSlot, HotkeyType, HotkeySlot)
    Call WriteSetHotkeySlot(HotkeySlot, Index, LastKnownSlot, HotkeyType)
End Sub

Public Sub ClearHotkeys()
    Dim i As Integer
    For i = 0 To HotKeyCount - 1
        HotkeyList(i).Index = -1
        HotkeyList(i).LastKnownSlot = -1
        HotkeyList(i).Type = Unknown
    Next i
End Sub

Public Sub ShowInteractionMenu(ByVal FormTop As Long, _
                                ByVal FormLeft As Long, _
                                ByVal FormHeight As Long, _
                                ByVal x As Integer, _
                                ByVal y As Integer, _
                                ByRef GameplayArea As RECT)
    On Error GoTo ShowInteractionMenu_Err
    Dim charindex As Integer
    charindex = MapData(tX, tY).charindex
    If charindex = 0 Then charindex = MapData(tX, tY + 1).charindex
    If charindex = 0 Or charindex = UserCharIndex Then Exit Sub

    Dim Frm As Form
    Call WriteLeftClick(tX, tY)
    TargetX = tX
    TargetY = tY

    If charlist(charindex).EsMascota Then
        Set Frm = MenuNPC
    ElseIf Not charlist(charindex).EsNpc Then
        targetName = charlist(charindex).nombre
        If charlist(UserCharIndex).priv > 0 And Shift = 0 Then
            Set Frm = MenuGM
        Else
            Set Frm = MenuUser
        End If
    End If

    If Not Frm Is Nothing Then
        Frm.Show
        Frm.Left = FormLeft + (GameplayArea.Left + x + 1) * screen.TwipsPerPixelX
        If (GameplayArea.Top + y) * screen.TwipsPerPixelY + Frm.Height > FormHeight Then
            Frm.Top = FormTop + (GameplayArea.Top + y) * screen.TwipsPerPixelY - Frm.Height
        Else
            Frm.Top = FormTop + (GameplayArea.Top + y) * screen.TwipsPerPixelY
        End If
        Set Frm = Nothing
    End If
    Exit Sub
ShowInteractionMenu_Err:
    Call RegistrarError(Err.Number, Err.Description, "ModGameplayUi.ShowInteractionMenu", Erl)
    Resume Next
End Sub


