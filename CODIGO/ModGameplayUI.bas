Attribute VB_Name = "ModGameplayUI"
Public Sub SetupGameplayUI()
    If BabelInitialized Then
        Call BabelUI.SetActiveScreen("gameplay")
        Call BabelUI.SetUserName(username)
    Else
        frmMain.shapexy.Left = 1200
        frmMain.shapexy.Top = 1200
        frmMain.shapexy.BackColor = RGB(170, 0, 0)
        frmMain.NombrePJ.Caption = username
        ' Detect links in console
        Call EnableURLDetect(frmMain.RecTxt.hwnd, frmMain.hwnd)
        Call Make_Transparent_Richtext(frmMain.RecTxt.hwnd)
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
    End If
End Sub

Public Sub OnClick(ByVal MouseButton As Long, ByVal MouseShift As Long)
On Error GoTo OnClick_Err

    If pausa Then Exit Sub
    
    If mascota.visible Then
        If Sqr((MouseX - mascota.PosX) ^ 2 + (MouseY - mascota.PosY) ^ 2) < 30 Then
            mascota.dialog = ""
        End If
    End If
    
    If cartel_visible Then
        If MouseX > 50 And MouseY > 478 And MouseX < 671 And MouseY < 585 Then
            If tutorial_index > 0 Then
                Call nextCartel
            Else
                Call cerrarCartel
            End If
        End If
    End If
    If MouseButton = vbLeftButton And ACCION1 = 0 Or MouseButton = vbRightButton And ACCION2 = 0 Or MouseButton = 4 And ACCION3 = 0 Then
        If Not Comerciando Then
            If MouseShift = 0 Then
                If UsingSkill = 0 Or frmMain.MacroLadder.enabled Then
                    Call CountPacketIterations(packetControl(ClientPacketID.LeftClick), 150)
                    Call WriteLeftClick(tX, tY)
                Else
                    Dim SendSkill As Boolean
                    If UsingSkill = magia Then
                        If ModoHechizos = BloqueoLanzar Then
                            SendSkill = IIf((MouseX >= frmMain.renderer.ScaleLeft And MouseX <= 736 + frmMain.renderer.ScaleLeft And MouseY >= frmMain.renderer.ScaleTop And MouseY <= frmMain.renderer.ScaleTop + 608), True, False)
                            If Not SendSkill Then
                                Exit Sub
                            End If
                            
                            Call MainTimer.Restart(TimersIndex.CastAttack)
                            Call MainTimer.Restart(TimersIndex.CastSpell)
                        Else
                            If MainTimer.Check(TimersIndex.AttackSpell, False) Then
                                If MainTimer.Check(TimersIndex.CastSpell) Then
                                    SendSkill = IIf((MouseX >= frmMain.renderer.ScaleLeft And MouseX <= 736 + frmMain.renderer.ScaleLeft And MouseY >= frmMain.renderer.ScaleTop And MouseY <= frmMain.renderer.ScaleTop + 608), True, False)
                                    If Not SendSkill Then
                                        Exit Sub
                                    End If
                                    Call MainTimer.Restart(TimersIndex.CastAttack)
                                
                                ElseIf ModoHechizos = SinBloqueo Then
                                    SendSkill = IIf((MouseX >= frmMain.renderer.ScaleLeft And MouseX <= 736 + frmMain.renderer.ScaleLeft And MouseY >= frmMain.renderer.ScaleTop And MouseY <= frmMain.renderer.ScaleTop + 608), True, False)
                                    
                                    If Not SendSkill Then
                                        Exit Sub
                                    End If
                                
                                    With FontTypes(FontTypeNames.FONTTYPE_TALK)
                                        Call ShowConsoleMsg("No puedes lanzar hechizos tan rápido.", .red, .green, .blue, .bold, .italic)
                                    End With
                                Else
                                    Exit Sub
                                End If
                                
                            ElseIf ModoHechizos = SinBloqueo Then
                                SendSkill = IIf((MouseX >= frmMain.renderer.ScaleLeft And MouseX <= 736 + frmMain.renderer.ScaleLeft And MouseY >= frmMain.renderer.ScaleTop And MouseY <= frmMain.renderer.ScaleTop + 608), True, False)
                                If Not SendSkill Then
                                    Exit Sub
                                End If
                                With FontTypes(FontTypeNames.FONTTYPE_TALK)
                                    Call ShowConsoleMsg("No puedes lanzar tan rápido después de un golpe.", .red, .green, .blue, .bold, .italic)
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
                                If MapData(tX, tY).CharIndex <> 0 Then
                                    Pos = InStr(charlist(MapData(tX, tY).CharIndex).nombre, "<")
                                    If Pos = 0 Then Pos = LenB(charlist(MapData(tX, tY).CharIndex).nombre) + 2
                                    frmPanelgm.cboListaUsus.Text = Left$(charlist(MapData(tX, tY).CharIndex).nombre, Pos - 2)
                                End If
                            Else
                                SendSkill = True
                            End If
                        End If
                    End If
                    
                    If (UsingSkill = eSkill.Pescar Or UsingSkill = eSkill.Talar Or UsingSkill = eSkill.Mineria Or _
                        UsingSkill = FundirMetal Or UsingSkill = eSkill.TargetableItem) Then
                        If MainTimer.Check(TimersIndex.CastSpell) Then
                            Call WriteWorkLeftClick(tX, tY, UsingSkill)
                            Call FormParser.Parse_Form(GetGameplayForm)

                            If CursoresGraficos = 0 Then
                                GetGameplayForm.MousePointer = vbDefault
                            End If
                        End If
                    End If
                   
                    If SendSkill Then
                        If UsingSkill = eSkill.magia Then
                            If ComprobarPosibleMacro(MouseX, MouseY) Then
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
    ElseIf MouseButton = vbLeftButton And ACCION1 = 2 Or MouseButton = vbRightButton And ACCION2 = 2 Or MouseButton = 4 And ACCION3 = 2 Then
        If UserDescansar Or UserMeditar Then Exit Sub
        If MainTimer.Check(TimersIndex.CastAttack, False) Then
            If MainTimer.Check(TimersIndex.Attack) Then
                Call MainTimer.Restart(TimersIndex.AttackSpell)
                Call WriteAttack
            End If
        End If
    
    ElseIf MouseButton = vbLeftButton And ACCION1 = 3 Or MouseButton = vbRightButton And ACCION2 = 3 Or MouseButton = 4 And ACCION3 = 3 Then
            If frmMain.Inventario.IsItemSelected Then Call WriteUseItem(frmMain.Inventario.SelectedItem)
    ElseIf MouseButton = vbLeftButton And ACCION1 = 4 Or MouseButton = vbRightButton And ACCION2 = 4 Or MouseButton = 4 And ACCION3 = 4 Then
        If MapData(tX, tY).CharIndex <> 0 Then
            If charlist(MapData(tX, tY).CharIndex).nombre <> charlist(MapData(UserPos.x, UserPos.y).CharIndex).nombre Then
                If charlist(MapData(tX, tY).CharIndex).esNpc = False Then
                    SendTxt.Text = "\" & charlist(MapData(tX, tY).CharIndex).nombre & " "
                    If SendTxtCmsg.visible = False Then
                        SendTxt.visible = True
                        SendTxt.SetFocus
                        SendTxt.SelStart = Len(SendTxt.Text)
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

Public Sub HandleGameplayAreaMouseUp(ByVal button As Integer, ByVal x As Integer, ByVal y As Integer, ByVal FormTop As Long, _
                                     ByVal FormLeft As Long, ByVal FormHeight As Long, ByRef GameplayArea As RECT)
    clicX = x
    clicY = y
    If button = vbLeftButton Then
        If HandleMouseInput(x, y) Then
        ElseIf Pregunta Then
            If x >= 419 And x <= 433 And y >= 243 And y <= 260 Then
                If PreguntaLocal Then
                    Select Case PreguntaNUM
                        Case 1
                            Pregunta = False
                            DestItemSlot = 0
                            DestItemCant = 0
                            PreguntaLocal = False
                        Case 2 ' Denunciar
                            Pregunta = False
                            PreguntaLocal = False
                    End Select
                Else
                    Call WriteResponderPregunta(False)
                    Pregunta = False
                End If
                Exit Sub
            ElseIf x >= 443 And x <= 458 And y >= 243 And y <= 260 Then
                If PreguntaLocal Then
                    Select Case PreguntaNUM
                        Case 1 '¿Destruir item?
                            Call WriteDrop(DestItemSlot, DestItemCant)
                            Pregunta = False
                            PreguntaLocal = False
                        Case 2 ' Denunciar
                            Call WriteDenounce(TargetName)
                            Pregunta = False
                            PreguntaLocal = False
                    End Select
                Else
                    Call WriteResponderPregunta(True)
                    Pregunta = False
                End If
                Exit Sub
            End If
        End If
    
    ElseIf button = vbRightButton Then
        Dim CharIndex As Integer
        CharIndex = MapData(tX, tY).CharIndex
        If CharIndex = 0 Then
            CharIndex = MapData(tX, tY + 1).CharIndex
        End If
        If CharIndex <> 0 And CharIndex <> UserCharIndex Then
            Dim Frm As Form
            Call WriteLeftClick(tX, tY)
            TargetX = tX
            TargetY = tY
            If charlist(CharIndex).EsMascota Then
                Set Frm = MenuNPC
            ElseIf Not charlist(CharIndex).esNpc Then
                TargetName = charlist(CharIndex).nombre
                If charlist(UserCharIndex).priv > 0 And Shift = 0 Then
                    Set Frm = MenuGM
                Else
                    Set Frm = MenuUser
                End If
            End If
            
            If Not Frm Is Nothing Then
                Call Frm.Show
                Frm.Left = FormLeft + (GameplayArea.Left + x + 1) * screen.TwipsPerPixelX
                If (GameplayArea.Top + y) * screen.TwipsPerPixelY + Frm.Height > FormHeight Then
                    Frm.Top = FormTop + (GameplayArea.Top + y) * screen.TwipsPerPixelY - Frm.Height
                Else
                    Frm.Top = FormTop + (GameplayArea.Top + y) * screen.TwipsPerPixelY
                End If
                Set Frm = Nothing
            End If
        End If
    End If
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

Public Sub SetInvItem(ByVal Slot As Byte, ByVal ObjIndex As Integer, ByVal Amount As Integer, ByVal Equipped As Byte, _
                      ByVal GrhIndex As Long, ByVal ObjType As Integer, ByVal MaxHit As Integer, ByVal MinHit As Integer, _
                      ByVal Def As Integer, ByVal Value As Single, ByVal Name As String, ByVal CanUse As Byte)

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
        .Valor = Valor
        .PuedeUsar = PuedeUsar
    End With
    If BabelInitialized Then
        Dim SlotInfo As t_InvItem
        SlotInfo.Amount = Amount
        SlotInfo.CanUse = podrausarlo
        SlotInfo.Equiped = Equipped
        SlotInfo.GrhIndex = GrhIndex
        SlotInfo.MaxDef = MaxDef
        SlotInfo.MinDef = MinDef
        SlotInfo.MaxHit = MaxHit
        SlotInfo.MinHit = MinHit
        SlotInfo.Name = Name
        SlotInfo.Slot = Slot
        SlotInfo.ObjIndex = ObjIndex
        SlotInfo.ObjType = ObjType
        Call SetInvSlot(SlotInfo)
    Else
        Call frmMain.Inventario.SetItem(Slot, ObjIndex, Amount, Equipped, GrhIndex, ObjType, MaxHit, MinHit, Def, Value, Name, CanUse)
    End If
End Sub

Public Sub SelectItemSlot(ByVal Slot As Integer)
    UserInventory.SelectedSlot = Slot
End Sub

Public Function GetSelectedItemSlot() As Integer
    If BabelInitialized Then
        GetSelectedItemSlot = UserInventory.SelectedSlot
    Else
        GetSelectedItemSlot = frmMain.Inventario.SelectedItem
    End If
End Function

Public Function IsItemSelected() As Boolean
    If BabelInitialized Then
        If UserInventory.SelectedSlot <= 0 Or UserInventory.SelectedSlot > UBound(UserInventory.Slots) Then Exit Function
        IsItemSelected = (UserInventory.Slots(UserInventory.SelectedSlot).GrhIndex > 0)
    Else
        IsItemSelected = frmMain.Inventario.IsItemSelected
    End If
End Function

Public Sub UseItemKey()
    If Not MainTimer.Check(TimersIndex.AttackUse, False) Then Exit Sub
        Call CountPacketIterations(packetControl(ClientPacketID.UseItemU), 100)
        If BabelInitialized Then
            If UserInventory.SelectedSlot > 0 And UserInventory.SelectedSlot <= UBound(UserInventory.Slots) Then
                Call WriteUseItemU(UserInventory.SelectedSlot)
            End If
        Else
            If frmMain.Inventario.IsItemSelected Then
                Call WriteUseItemU(frmMain.Inventario.SelectedItem)
            End If
        End If
        
End Sub

Public Sub UserItemClick()
    If frmCarp.visible Or frmHerrero.visible Or frmComerciar.visible Or frmBancoObj.visible Then Exit Sub
    If pausa Then Exit Sub
    If UserMeditar Then Exit Sub
    If frmMain.macrotrabajo.enabled Then frmMain.DesactivarMacroTrabajo
    If Not IsItemSelected Then Exit Sub

    ' Hacemos acción del doble clic correspondiente
    If BabelInitialized Then
        Call UserOrEquipItem(UserInventory.SelectedSlot, UserInventory.Slots(UserInventory.SelectedSlot).Equipped, UserInventory.Slots(UserInventory.SelectedSlot).ObjIndex)
    Else
        Call UserOrEquipItem(frmMain.Inventario.SelectedItem, frmMain.Inventario.Equipped(frmMain.Inventario.SelectedItem), frmMain.Inventario.ObjIndex(frmMain.Inventario.SelectedItem))
    End If
End Sub

Public Sub UserOrEquipItem(ByVal Slot As Integer, ByVal Equipped As Boolean, ByVal ObjIndex As Integer)
    Dim ObjType As Byte
    ObjType = ObjData(ObjIndex).ObjType
    Select Case ObjType
        Case eObjType.otArmadura, eObjType.otESCUDO, eObjType.otmagicos, eObjType.otFlechas, eObjType.otCASCO, eObjType.otNudillos, eObjType.otAnillos, eObjType.otManchas
            If Not Equipped Then
                Call WriteEquipItem(Slot)
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
            Call CountPacketIterations(packetControl(ClientPacketID.UseItem), 180)
            Call WriteUseItem(Slot)
    End Select
End Sub

Public Sub HandleKeyUp(KeyCode As Integer, Shift As Integer)
    If Not BabelUI.InputFocus Then
        If Not IsDialogOpen Then
            If Accionar(KeyCode) Then
                Exit Sub
            ElseIf KeyCode = vbKeyReturn Then
                Call OpenChatInput
            ElseIf KeyCode = vbKeyDelete Then
                Call OpenAndFocusClanChat
            ElseIf KeyCode = vbKeyEscape And Not UserSaliendo Then
                Call HandleEsc
            ElseIf KeyCode = 27 And UserSaliendo Then
                Call WriteCancelarExit
            ElseIf KeyCode = 80 And PescandoEspecial Then
                Call IntentarObtenerPezEspecial
            End If
        End If
    ElseIf Not BabelInitialized Then
        Call FocusInput
    End If
End Sub

Public Sub HandleEsc()
    If BabelInitialized Then
        frmCerrar.Show , frmBabelUI
    Else
        frmCerrar.Show , frmMain
    End If
End Sub
Public Function IsDialogOpen() As Boolean
    IsDialogOpen = pausa Or frmComerciar.visible Or frmComerciarUsu.visible Or frmBancoObj.visible Or frmGoliath.visible
End Function

Public Function IsInputFocus() As Boolean
    If BabelInitialized Then
        IsInputFocus = BabelUI.InputFocus
        Exit Function
    Else
        IsInputFocus = frmMain.SendTxt.visible Or frmMain.SendTxtCmsg.visible
    End If
End Function

Public Sub OpenAndFocusClanChat()
    If BabelInitialized Then
        Call OpenChat(e_ChatMode.ClanChat)
    Else
        If Not frmMain.SendTxt.visible Then
            frmMain.SendTxtCmsg.visible = True
            frmMain.SendTxtCmsg.SetFocus
        End If
    End If
    Call DialogosClanes.toggle_dialogs_visibility(True)
End Sub

Public Sub OpenChatInput()
    If BabelInitialized Then
        Call OpenChat(e_ChatMode.NormalChat)
    Else
        If Not frmCantidad.visible Then
            Call frmMain.CompletarEnvioMensajes
            StartOpenChatTime = GetTickCount
            frmMain.SendTxt.visible = True
            frmMain.SendTxt.SetFocus
        End If
    End If
End Sub

Public Sub FocusInput()
    If BabelInitialized Then
        Call OpenChat(e_ChatMode.NormalChat)
    Else
        If frmMain.SendTxt.visible Then
            frmMain.SendTxt.SetFocus
        End If
        If frmMain.SendTxtCmsg.visible Then
            frmMain.SendTxtCmsg.SetFocus
        End If
    End If
End Sub

Public Function GetGameplayForm() As Form
    If BabelInitialized Then
        Set GetGameplayForm = frmBabelUI
    Else
        Set GetGameplayForm = frmMain
    End If
End Function

Public Sub UseSpell(ByVal SpellSlot As Byte, ByVal SpellName As String)
If pausa Then Exit Sub

    TempTick = GetTickCount And &H7FFFFFFF
    If TempTick - iClickTick < IntervaloEntreClicks And Not iClickTick = 0 And _
       LastMacroButton <> tMacroButton.Lanzar Then
        Call WriteLogMacroClickHechizo(tMacro.Coordenadas)
    End If
    
    iClickTick = TempTick
    LastMacroButton = tMacroButton.Lanzar
    If SpellName <> "(Vacío)" Then
        If UserStats.estado = 1 Then
            With FontTypes(FontTypeNames.FONTTYPE_INFO)
                Call ShowConsoleMsg("¡¡Estás muerto!!", .red, .green, .blue, .bold, .italic)
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
    If BabelInitialized Then
        Dim Pos As t_Position
        Pos.x = UserPos.x
        Pos.y = UserPos.y
        Call ConvertToMinimapPosition(Pos.x, Pos.y, 2, 2)
        Call BabelUI.UpdateUserPos(UserPos.x, UserPos.y, Pos)
    Else
        Call frmMain.SetMinimapPosition(0, UserPos.x, UserPos.y)
        frmMain.Coord.Caption = UserMap & "-" & UserPos.x & "-" & UserPos.y
        If frmMapaGrande.visible Then
            Call frmMapaGrande.ActualizarPosicionMapa
        End If
    End If
End Sub

Public Sub RequestSkills()
    If pausa Or tutorial_index > 0 Then Exit Sub
    If MostrarTutorial And tutorial_index <= 0 Then
        If tutorial(4).Activo = 1 Then
            tutorial_index = e_tutorialIndex.TUTORIAL_SkillPoints
            'TUTORIAL MAPA INSEGURO
            Call mostrarCartel(tutorial(tutorial_index).titulo, tutorial(tutorial_index).textos(1), tutorial(tutorial_index).Grh, -1, &H164B8A, , , False, 100, 479, 100, 535, 640, 530, 64, 64)
            Exit Sub
        End If
    End If
    
    LlegaronSkills = True
    Call WriteRequestSkills
End Sub
