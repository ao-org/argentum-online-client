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
    On Error Goto SetupGameplayUI_Err
   
        frmMain.shapexy.Left = 1200
        frmMain.shapexy.Top = 1200
        frmMain.shapexy.BackColor = RGB(170, 0, 0)
        frmMain.NombrePJ.Caption = userName
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
        ActiveInventoryTab = eInventory

    Call LoadHotkeys
    Exit Sub
SetupGameplayUI_Err:
    Call TraceError(Err.Number, Err.Description, "ModGameplayUI.SetupGameplayUI", Erl)
End Sub

Public Sub OnClick(ByVal MouseButton As Long, ByVal MouseShift As Long)
    On Error Goto OnClick_Err
On Error GoTo OnClick_Err

    If pausa Then Exit Sub
    If IsGameDialogOpen Then Exit Sub
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
                                        Call ShowConsoleMsg(JsonLanguage.Item("MENSAJE_LANZAMIENTO_RAPIDO"), .red, .green, .blue, .bold, .italic) ' MENSAJE_LANZAMIENTO_RAPIDO=No puedes lanzar hechizos tan rápido.
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
                    
                    If (UsingSkill = eSkill.Pescar Or UsingSkill = eSkill.Talar Or UsingSkill = eSkill.Alquimia Or UsingSkill = eSkill.Mineria Or _
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
        Dim CharIndex As Integer
        CharIndex = MapData(tX, tY).CharIndex
        If CharIndex = 0 And tY < YMaxMapSize Then
            CharIndex = MapData(tX, tY + 1).CharIndex
        End If
        If CharIndex <> 0 Then
            If charlist(CharIndex).nombre <> charlist(UserCharIndex).nombre Then
                If charlist(CharIndex).EsNpc = False Then
                    frmMain.SendTxt.Text = "\" & charlist(CharIndex).nombre & " "
                    If frmMain.SendTxtCmsg.visible = False Then
                        frmMain.SendTxt.visible = True
                        frmMain.SendTxt.SetFocus
                        frmMain.SendTxt.SelStart = Len(frmMain.SendTxt.Text)
                    End If
                End If
            End If
        End If
    End If
    Exit Sub

OnClick_Err:
    Call RegistrarError(Err.Number, Err.Description, "ModGameplayUi.OnClick", Erl)
    Resume Next
    Exit Sub
OnClick_Err:
    Call TraceError(Err.Number, Err.Description, "ModGameplayUI.OnClick", Erl)
End Sub

Public Sub HandleQuestionResponse(ByVal Result As Boolean)
    On Error Goto HandleQuestionResponse_Err
    If PreguntaLocal Then
        If Result Then
            Select Case PreguntaNUM
                Case 1 '¿Destruir item?
                    Call WriteDrop(DestItemSlot, DestItemCant)
                Case 2 ' Denunciar
                    Call WriteDenounce(TargetName)
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
    Exit Sub
HandleQuestionResponse_Err:
    Call TraceError(Err.Number, Err.Description, "ModGameplayUI.HandleQuestionResponse", Erl)
End Sub

Public Sub HandleGameplayAreaMouseUp(ByVal Button As Integer, ByVal x As Integer, ByVal y As Integer, ByVal FormTop As Long, _
    On Error Goto HandleGameplayAreaMouseUp_Err
                                     ByVal FormLeft As Long, ByVal FormHeight As Long, ByRef GameplayArea As RECT)
    clicX = x
    clicY = y
    If Button = vbLeftButton Then
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
    
    ElseIf Button = vbRightButton Then
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
            ElseIf Not charlist(CharIndex).EsNpc Then
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
    Exit Sub
HandleGameplayAreaMouseUp_Err:
    Call TraceError(Err.Number, Err.Description, "ModGameplayUI.HandleGameplayAreaMouseUp", Erl)
End Sub

Public Sub HandleChatMsg(ByVal InputText As String)
    On Error Goto HandleChatMsg_Err
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
    Exit Sub
HandleChatMsg_Err:
    Call TraceError(Err.Number, Err.Description, "ModGameplayUI.HandleChatMsg", Erl)
End Sub

Public Sub UseSelectInvItem()
    On Error Goto UseSelectInvItem_Err
    
    Exit Sub
UseSelectInvItem_Err:
    Call TraceError(Err.Number, Err.Description, "ModGameplayUI.UseSelectInvItem", Erl)
End Sub

Public Sub SetInvItem(ByVal Slot As Byte, ByVal ObjIndex As Integer, ByVal Amount As Integer, ByVal Equipped As Byte, _
    On Error Goto SetInvItem_Err
                      ByVal GrhIndex As Long, ByVal ObjType As Integer, ByVal MaxHit As Integer, ByVal MinHit As Integer, _
                      ByVal Def As Integer, ByVal Value As Single, ByVal Name As String, ByVal CanUse As Byte, ByVal ElementalTags As Long, ByVal IsBindable As Byte)

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
        .Valor = Value
        .PuedeUsar = CanUse
        .ElementalTags = ElementalTags
        .IsBindable = IsBindable > 0
    End With
    Call frmMain.Inventario.SetItem(Slot, ObjIndex, Amount, Equipped, GrhIndex, ObjType, MaxHit, MinHit, Def, Value, Name, ElementalTags, CanUse)
    Exit Sub
SetInvItem_Err:
    Call TraceError(Err.Number, Err.Description, "ModGameplayUI.SetInvItem", Erl)
End Sub

Public Sub SelectItemSlot(ByVal Slot As Integer)
    On Error Goto SelectItemSlot_Err
    UserInventory.SelectedSlot = Slot
    Exit Sub
SelectItemSlot_Err:
    Call TraceError(Err.Number, Err.Description, "ModGameplayUI.SelectItemSlot", Erl)
End Sub

Public Function GetSelectedItemSlot() As Integer
    On Error Goto GetSelectedItemSlot_Err
    GetSelectedItemSlot = frmMain.Inventario.SelectedItem
    Exit Function
GetSelectedItemSlot_Err:
    Call TraceError(Err.Number, Err.Description, "ModGameplayUI.GetSelectedItemSlot", Erl)
End Function

Public Function IsItemSelected() As Boolean
    On Error Goto IsItemSelected_Err
    IsItemSelected = frmMain.Inventario.IsItemSelected

    Exit Function
IsItemSelected_Err:
    Call TraceError(Err.Number, Err.Description, "ModGameplayUI.IsItemSelected", Erl)
End Function

Public Sub UseItemKey()
    On Error Goto UseItemKey_Err
    If Not MainTimer.Check(TimersIndex.AttackUse, False) Then Exit Sub
        Call CountPacketIterations(packetControl(ClientPacketID.eUseItemU), 100)
        If frmMain.Inventario.IsItemSelected Then
                Call WriteUseItemU(frmMain.Inventario.SelectedItem)
            End If
    Exit Sub
UseItemKey_Err:
    Call TraceError(Err.Number, Err.Description, "ModGameplayUI.UseItemKey", Erl)
End Sub

Public Sub UserItemClick()
    On Error Goto UserItemClick_Err
    If frmCarp.visible Or frmHerrero.visible Or frmComerciar.visible Or frmBancoObj.visible Then Exit Sub
    If pausa Then Exit Sub
    If UserMeditar Then Exit Sub
    If frmMain.macrotrabajo.enabled Then frmMain.DesactivarMacroTrabajo
    If Not IsItemSelected Then Exit Sub

    Call UserOrEquipItem(frmMain.Inventario.SelectedItem, frmMain.Inventario.Equipped(frmMain.Inventario.SelectedItem), frmMain.Inventario.ObjIndex(frmMain.Inventario.SelectedItem))

    Exit Sub
UserItemClick_Err:
    Call TraceError(Err.Number, Err.Description, "ModGameplayUI.UserItemClick", Erl)
End Sub

Public Sub UserOrEquipItem(ByVal Slot As Integer, ByVal Equipped As Boolean, ByVal ObjIndex As Integer)
    On Error Goto UserOrEquipItem_Err
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
    Exit Sub
UserOrEquipItem_Err:
    Call TraceError(Err.Number, Err.Description, "ModGameplayUI.UserOrEquipItem", Erl)
End Sub

Public Sub HandleKeyUp(KeyCode As Integer, Shift As Integer)
    On Error Goto HandleKeyUp_Err
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
                Call ParseUserCommand("/IRA " & TargetName)
            End If
        End If
    Else
        Call FocusInput
    End If
    Exit Sub
HandleKeyUp_Err:
    Call TraceError(Err.Number, Err.Description, "ModGameplayUI.HandleKeyUp", Erl)
End Sub

Public Sub HandleEsc()
    On Error Goto HandleEsc_Err

        frmCerrar.Show , frmMain

    Exit Sub
HandleEsc_Err:
    Call TraceError(Err.Number, Err.Description, "ModGameplayUI.HandleEsc", Erl)
End Sub
Public Function IsDialogOpen() As Boolean
    On Error Goto IsDialogOpen_Err
    IsDialogOpen = pausa Or frmComerciar.visible Or frmComerciarUsu.visible Or frmBancoObj.visible Or frmGoliath.visible Or IsGameDialogOpen
    Exit Function
IsDialogOpen_Err:
    Call TraceError(Err.Number, Err.Description, "ModGameplayUI.IsDialogOpen", Erl)
End Function

Public Function IsInputFocus() As Boolean
    On Error Goto IsInputFocus_Err
    IsInputFocus = frmMain.SendTxt.visible Or frmMain.SendTxtCmsg.visible

    Exit Function
IsInputFocus_Err:
    Call TraceError(Err.Number, Err.Description, "ModGameplayUI.IsInputFocus", Erl)
End Function

Public Sub OpenAndFocusClanChat()
    On Error Goto OpenAndFocusClanChat_Err
       If Not frmMain.SendTxt.visible Then
            frmMain.SendTxtCmsg.visible = True
            frmMain.SendTxtCmsg.SetFocus
        End If

    Call DialogosClanes.toggle_dialogs_visibility(True)
    Exit Sub
OpenAndFocusClanChat_Err:
    Call TraceError(Err.Number, Err.Description, "ModGameplayUI.OpenAndFocusClanChat", Erl)
End Sub

Public Sub OpenChatInput()
    On Error Goto OpenChatInput_Err
        If Not frmCantidad.visible Then
            Call frmMain.CompletarEnvioMensajes
            StartOpenChatTime = GetTickCount
            frmMain.SendTxt.visible = True
            frmMain.SendTxt.SetFocus
        End If

    Exit Sub
OpenChatInput_Err:
    Call TraceError(Err.Number, Err.Description, "ModGameplayUI.OpenChatInput", Erl)
End Sub

Public Sub FocusInput()
    On Error Goto FocusInput_Err

        If frmMain.SendTxt.visible Then
            frmMain.SendTxt.SetFocus
        End If
        If frmMain.SendTxtCmsg.visible Then
            frmMain.SendTxtCmsg.SetFocus
        End If

    Exit Sub
FocusInput_Err:
    Call TraceError(Err.Number, Err.Description, "ModGameplayUI.FocusInput", Erl)
End Sub

Public Function GetGameplayForm() As Form
    On Error Goto GetGameplayForm_Err

        Set GetGameplayForm = frmMain

    Exit Function
GetGameplayForm_Err:
    Call TraceError(Err.Number, Err.Description, "ModGameplayUI.GetGameplayForm", Erl)
End Function

Public Sub UseSpell(ByVal SpellSlot As Byte, ByVal SpellName As String)
    On Error Goto UseSpell_Err
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
    Exit Sub
UseSpell_Err:
    Call TraceError(Err.Number, Err.Description, "ModGameplayUI.UseSpell", Erl)
End Sub

Public Sub UpdateMapPos()
    On Error Goto UpdateMapPos_Err

        Call frmMain.SetMinimapPosition(0, UserPos.x, UserPos.y)
        frmMain.Coord.Caption = UserMap & "-" & UserPos.x & "-" & UserPos.y
        If frmMapaGrande.visible Then
            Call frmMapaGrande.ActualizarPosicionMapa
        End If

    Exit Sub
UpdateMapPos_Err:
    Call TraceError(Err.Number, Err.Description, "ModGameplayUI.UpdateMapPos", Erl)
End Sub

Public Sub RequestSkills()
    On Error Goto RequestSkills_Err
    If pausa Or tutorial_index > 0 Then Exit Sub
    If MostrarTutorial And tutorial_index <= 0 Then
        If tutorial(4).Activo = 1 Then
            tutorial_index = e_tutorialIndex.TUTORIAL_SkillPoints
            'TUTORIAL MAPA INSEGURO
            Call mostrarCartel(tutorial(tutorial_index).titulo, tutorial(tutorial_index).textos(1), tutorial(tutorial_index).grh, -1, &H164B8A, , , False, 100, 479, 100, 535, 640, 530, 64, 64)
            Exit Sub
        End If
    End If
    
    LlegaronSkills = True
    Call WriteRequestSkills
    Exit Sub
RequestSkills_Err:
    Call TraceError(Err.Number, Err.Description, "ModGameplayUI.RequestSkills", Erl)
End Sub

Public Function IsUsableItem(ByRef ItemData As ObjDatas) As Boolean
    On Error Goto IsUsableItem_Err
    IsUsableItem = ItemData.ObjType = eObjType.otWeapon Or ItemData.ObjType = eObjType.otPociones Or _
             ItemData.ObjType = eObjType.OtHerramientas Or ItemData.ObjType = eObjType.otInstrumentos Or _
             ItemData.ObjType = eObjType.OtCofre
    Exit Function
IsUsableItem_Err:
    Call TraceError(Err.Number, Err.Description, "ModGameplayUI.IsUsableItem", Erl)
End Function

Public Sub EquipSelectedItem()
    On Error Goto EquipSelectedItem_Err
        If frmMain.Inventario.IsItemSelected Then Call WriteEquipItem(frmMain.Inventario.SelectedItem)

    Exit Sub
EquipSelectedItem_Err:
    Call TraceError(Err.Number, Err.Description, "ModGameplayUI.EquipSelectedItem", Erl)
End Sub

Public Sub OpenCreateObjectMenu()
    On Error Goto OpenCreateObjectMenu_Err
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
    Exit Sub
OpenCreateObjectMenu_Err:
    Call TraceError(Err.Number, Err.Description, "ModGameplayUI.OpenCreateObjectMenu", Erl)
End Sub

Public Sub SelectInventoryTab()
    On Error Goto SelectInventoryTab_Err
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
    Exit Sub
SelectInventoryTab_Err:
    Call TraceError(Err.Number, Err.Description, "ModGameplayUI.SelectInventoryTab", Erl)
End Sub

Public Sub SelectSpellTab()
    On Error Goto SelectSpellTab_Err
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
    Exit Sub
SelectSpellTab_Err:
    Call TraceError(Err.Number, Err.Description, "ModGameplayUI.SelectSpellTab", Erl)
End Sub

Public Sub GetMinimapPosition(ByRef x As Single, ByRef y As Single)
    On Error Goto GetMinimapPosition_Err
    x = x * (100 - 2 * HalfWindowTileWidth - 4) / 100 + HalfWindowTileWidth + 2
    y = y * (100 - 2 * HalfWindowTileHeight - 4) / 100 + HalfWindowTileHeight + 2
    Exit Sub
GetMinimapPosition_Err:
    Call TraceError(Err.Number, Err.Description, "ModGameplayUI.GetMinimapPosition", Erl)
End Sub

Public Sub RequestMeditate()
    On Error Goto RequestMeditate_Err
    If UserStats.minman = UserStats.maxman Then Exit Sub
    If UserStats.estado = 1 Then
        With FontTypes(FontTypeNames.FONTTYPE_INFO)
            Call ShowConsoleMsg(JsonLanguage.Item("MENSAJE_ESTAS_MUERTO"), .red, .green, .blue, .bold, .italic) ' MENSAJE_ESTAS_MUERTO=¡Estás muerto!
        End With
        Exit Sub
    End If
    Call WriteMeditate
    Exit Sub
RequestMeditate_Err:
    Call TraceError(Err.Number, Err.Description, "ModGameplayUI.RequestMeditate", Erl)
End Sub

Public Sub SetHotkey(ByVal Index As Integer, ByVal LastKnownSlot As Integer, ByVal HotkeyType As e_HotkeyType, ByVal HotkeySlot As Integer)
    On Error Goto SetHotkey_Err
    HotkeyList(HotkeySlot).Index = Index
    HotkeyList(HotkeySlot).LastKnownSlot = LastKnownSlot
    HotkeyList(HotkeySlot).Type = HotkeyType
    Call SaveHotkey(Index, LastKnownSlot, HotkeyType, HotkeySlot)
    Call WriteSetHotkeySlot(HotkeySlot, Index, LastKnownSlot, HotkeyType)
    Exit Sub
SetHotkey_Err:
    Call TraceError(Err.Number, Err.Description, "ModGameplayUI.SetHotkey", Erl)
End Sub

Public Sub ClearHotkeys()
    On Error Goto ClearHotkeys_Err
    Dim i As Integer
    For i = 0 To HotKeyCount - 1
        HotkeyList(i).Index = -1
        HotkeyList(i).LastKnownSlot = -1
        HotkeyList(i).Type = Unknown
    Next i
    Exit Sub
ClearHotkeys_Err:
    Call TraceError(Err.Number, Err.Description, "ModGameplayUI.ClearHotkeys", Erl)
End Sub
