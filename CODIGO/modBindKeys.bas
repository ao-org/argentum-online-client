Attribute VB_Name = "modBindKeys"
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

Type tBoton

    TipoAccion As Integer
    SendString As String
    hlist As Integer
    invslot As Integer

End Type

Type tBindedKey

    KeyCode As Integer
    Name As String

End Type

Public ServerIndex      As String

Public NUMBINDS         As Integer

Public ACCION1          As e_MouseAction

Public ACCION2          As e_MouseAction

Public ACCION3          As e_MouseAction

Public BindKeys()       As tBindedKey

Public BotonElegido     As Integer

Public MacroTipoElegido As Byte

Public Enum e_KeyAction
    eAttack = 1
    eLoot = 2
    eDrop = 3
    eUseItem = 4
    eEquipItem = 5
    eChangeSafe = 6
    eHideNames = 7
    ePartyToggle = 8
    eSteal = 9
    eRefreshPosition = 10
    eHide = 11
    eShowOnline = 12
    eScreenShoot = 13
    eMoveUp = 14
    eMoveDown = 15
    eMoveLeft = 16
    eMoveRight = 17
    eClanCall = 18
    eGameStats = 19
    eClanMark = 20
    eMeditate = 21
    eExitGame = 22
    eTaming = 23
    eHKey1 = 24
    eHKey2 = 25
    eHKey3 = 26
    eHKey4 = 27
    eHKey5 = 28
    eHKey6 = 29
    eHKey7 = 30
    eHKey8 = 31
    eHKey9 = 32
    eHKey10 = 33
    eSendText = 34
    
    [eMaxBinds]
End Enum

Public Enum e_MouseAction
    eThrowOrLook
    eInteract
    eAttack
    eWhisper
    
    eUnknown
End Enum

Const CustomKeyMappingFile As String = "\..\Recursos\OUTPUT\Teclas.ini"
Const DefaultKeyMappingFile As String = "\..\Recursos\OUTPUT\DefaultKey.ini"

Public Sub LoadBindedKeys()
    
    On Error GoTo LoadBindedKeys_Err

    If Not FileExist(App.path & DefaultKeyMappingFile, vbArchive) Then
        MsgBox JsonLanguage.Item("MENSAJE_ARCHIVO_REQUERIDO") & App.path & DefaultKeyMappingFile, vbCritical + vbOKOnly, JsonLanguage.Item("TITULO_ERROR")
        End
    End If

    ' Si no existe el Teclas.ini lo creamos como copia del DefaultKey.ini
    If Not FileExist(App.path & CustomKeyMappingFile, vbArchive) Then
        Call FileSystem.FileCopy(App.path & DefaultKeyMappingFile, App.path & CustomKeyMappingFile)
    End If
    
    Dim DefaultBinds As New clsIniManager
    Call DefaultBinds.Initialize(App.path & DefaultKeyMappingFile)
    
    Dim UserBinds As New clsIniManager
    Call UserBinds.Initialize(App.path & CustomKeyMappingFile)

    NUMBINDS = eMaxBinds - 1

    ACCION1 = GetAction(DefaultBinds, UserBinds, 1)
    ACCION2 = GetAction(DefaultBinds, UserBinds, 2)
    ACCION3 = GetAction(DefaultBinds, UserBinds, 3)

    ReDim Preserve BindKeys(1 To NUMBINDS) As tBindedKey

    Dim Index As Integer
    Dim Bind As String

    For Index = 1 To NUMBINDS
        Bind = GetBind(DefaultBinds, UserBinds, CStr(Index))
        BindKeys(Index).KeyCode = Val(General_Field_Read(1, Bind, ","))
        BindKeys(Index).Name = General_Field_Read(2, Bind, ",")
    Next Index

    Set DefaultBinds = Nothing
    Set UserBinds = Nothing
    
    Exit Sub

LoadBindedKeys_Err:
    Call RegistrarError(Err.Number, Err.Description, "ModUtils.LoadBindedKeys", Erl)
    Resume Next
    
End Sub

Private Function GetAction(ByRef DefaultBinds As clsIniManager, ByRef UserBinds As clsIniManager, ByVal Index As Integer) As e_MouseAction
    Dim Temp As String
    Temp = UserBinds.GetValue("INIT", "ACCION" & Index)
    GetAction = ParseMouseAction(Trim(Temp))
    
    If GetAction = e_MouseAction.eUnknown Then
        GetAction = ParseOldMouseAction(Temp)
        
        If GetAction = e_MouseAction.eUnknown Then
            Temp = DefaultBinds.GetValue("INIT", "ACCION" & Index)
            GetAction = ParseMouseAction(Trim(Temp))
        End If
    End If
End Function

Private Function GetBind(ByRef DefaultBinds As clsIniManager, ByRef UserBinds As clsIniManager, ByVal key As String) As String
    GetBind = UserBinds.GetValue("USER", key)
    If GetBind = vbNullString Then
        GetBind = DefaultBinds.GetValue("DEFAULTS", key)
    End If
End Function

Public Sub SaveBindedKeys()
    
    On Error GoTo SaveBindedKeys_Err
    

    Dim lC As Integer, Arch As String

    Arch = App.path & "\..\Recursos\OUTPUT\" & "Teclas.ini"

    Call General_Var_Write(Arch, "INIT", "NUMBINDS", Int(NUMBINDS))

    Call General_Var_Write(Arch, "INIT", "ACCION1", MouseActionToString(ACCION1))
    Call General_Var_Write(Arch, "INIT", "ACCION2", MouseActionToString(ACCION2))
    Call General_Var_Write(Arch, "INIT", "ACCION3", MouseActionToString(ACCION3))

    For lC = 1 To NUMBINDS
        Call General_Var_Write(Arch, "USER", CStr(lC), CStr(BindKeys(lC).KeyCode) & "," & BindKeys(lC).Name)
    Next lC

    lC = 0

    
    Exit Sub

SaveBindedKeys_Err:
    Call RegistrarError(Err.Number, Err.Description, "ModUtils.SaveBindedKeys", Erl)
    Resume Next
    
End Sub

Sub LoadDefaultBinds()
    
    On Error GoTo LoadDefaultBinds_Err

    Dim DefaultBinds As New clsIniManager
    Call DefaultBinds.Initialize(App.path & DefaultKeyMappingFile)

    Dim Index As Integer, Temp As String
    For Index = 1 To NUMBINDS
        Temp = DefaultBinds.GetValue("DEFAULTS", CStr(Index))
        BindKeys(Index).KeyCode = Val(General_Field_Read(1, Temp, ","))
        BindKeys(Index).Name = General_Field_Read(2, Temp, ",")
    Next Index

    Set DefaultBinds = Nothing
    
    Exit Sub

LoadDefaultBinds_Err:
    Call RegistrarError(Err.Number, Err.Description, "modBindKeys.LoadDefaultBinds", Erl)
    Resume Next
    
End Sub

Sub LoadDefaultBinds2()
    
    On Error GoTo LoadDefaultBinds2_Err

    Dim DefaultBinds As New clsIniManager
    Call DefaultBinds.Initialize(App.path & DefaultKeyMappingFile)

    Dim Index As Integer, Temp As String
    For Index = 1 To NUMBINDS
        Temp = DefaultBinds.GetValue("DEFAULTSMODERN", CStr(Index))
        BindKeys(Index).KeyCode = Val(General_Field_Read(1, Temp, ","))
        BindKeys(Index).Name = General_Field_Read(2, Temp, ",")
    Next Index

    Set DefaultBinds = Nothing
    
    Exit Sub

LoadDefaultBinds2_Err:
    Call RegistrarError(Err.Number, Err.Description, "modBindKeys.LoadDefaultBinds2", Erl)
    Resume Next
    
End Sub

Public Function Accionar(ByVal KeyCode As Integer) As Boolean
    
    On Error GoTo Accionar_Err
    
    
    Select Case KeyCode
        Case BindKeys(1).KeyCode
            If UserStats.estado = 1 Then
    
                With FontTypes(FontTypeNames.FONTTYPE_INFO)
                    Call ShowConsoleMsg("¡Estás muerto!", .red, .green, .blue, .bold, .italic)
    
                End With
    
                Exit Function
    
            End If
    
            If UserDescansar Then Exit Function
            If frmMain.Inventario.IsItemOnCd(frmMain.Inventario.GetActiveWeaponSlot) Then Exit Function
            If MainTimer.Check(TimersIndex.CastAttack, False) Then
                If MainTimer.Check(TimersIndex.Attack) Then
                    Call MainTimer.Restart(TimersIndex.AttackSpell)
                    Call MainTimer.Restart(TimersIndex.AttackUse)
                    Set cooldown_ataque = New clsCooldown
                    Call cooldown_ataque.Cooldown_Initialize(gIntervals.Hit, 36602)
                    Call WriteAttack
                End If
    
            End If
    
        Case BindKeys(2).KeyCode
    
            If UserStats.estado = 1 Then
    
                With FontTypes(FontTypeNames.FONTTYPE_INFO)
                    Call ShowConsoleMsg("¡Estás muerto!", .red, .green, .blue, .bold, .italic)
    
                End With
    
                Exit Function
    
            End If
    
            If Not Comerciando Then
                Call AgarrarItem
            Else
                Call AddtoRichTextBox(frmMain.RecTxt, JsonLanguage.Item("MENSAJE_NO_PODES_AGARRAR_OBJETOS_MIENTRAS_COMERCIAS"), 255, 0, 32, False, False, False)
    
            End If
    
        Case BindKeys(3).KeyCode
    
            If UserStats.estado = 1 Then
    
                With FontTypes(FontTypeNames.FONTTYPE_INFO)
                    Call ShowConsoleMsg("¡Estás muerto!", .red, .green, .blue, .bold, .italic)
    
                End With
    
                Exit Function
    
            End If
    
            If Not Comerciando Then
                Call TirarItem
            Else
                Call AddtoRichTextBox(frmMain.RecTxt, JsonLanguage.Item("MENSAJE_NO_PODES_TIRAR_OBJETOS_MIENTRAS_COMERCIAS"), 255, 0, 32, False, False, False)
    
            End If
    
        Case BindKeys(6).KeyCode
            If SeguroGame Then
                Call AddtoRichTextBox(frmMain.RecTxt, JsonLanguage.Item("MENSAJE_DESACTIVAR_SEGURO_CON_SEG"), 255, 0, 0, True, False, False)
            Else
                Call WriteSafeToggle
            End If
        Case BindKeys(12).KeyCode
            MostrarOnline = Not MostrarOnline
        Case BindKeys(7).KeyCode
            Nombres = Not Nombres
        Case BindKeys(8).KeyCode
            Call WriteParyToggle
        Case BindKeys(9).KeyCode
    
            If UserStats.estado = 1 Then
    
                With FontTypes(FontTypeNames.FONTTYPE_INFO)
                    Call ShowConsoleMsg("¡Estás muerto!", .red, .green, .blue, .bold, .italic)
    
                End With
    
                Exit Function
    
            End If
    
            Call WriteWork(eSkill.Robar)
            
        Case BindKeys(18).KeyCode
    
            If IntervaloPermiteLLamadaClan Then Call WriteLlamadadeClan
            
        Case BindKeys(20).KeyCode
    
            If IntervaloPermiteLLamadaClan Then Call WriteMarcaDeClan
        
        Case BindKeys(5).KeyCode
            If UserStats.estado = 1 Then
                With FontTypes(FontTypeNames.FONTTYPE_INFO)
                    Call ShowConsoleMsg("¡Estás muerto!", .red, .green, .blue, .bold, .italic)
                End With
                Exit Function
            End If
                Call EquipSelectedItem
        
        Case BindKeys(4).KeyCode
            Call UseItemKey
        
        Case BindKeys(10).KeyCode
    
            If MainTimer.Check(TimersIndex.SendRPU) Then
                Call WriteRequestPositionUpdate
                Beep
    
            End If
        
        Case BindKeys(11).KeyCode
    
            If UserStats.estado = 1 Then
    
                With FontTypes(FontTypeNames.FONTTYPE_INFO)
                    Call ShowConsoleMsg("¡Estás muerto!", .red, .green, .blue, .bold, .italic)
    
                End With
    
                Exit Function
    
            End If
    
            Call WriteWork(eSkill.Ocultarse)
                
        Case BindKeys(19).KeyCode
            Call SaveSetting("OPCIONES", "FPSFLAG", FPSFLAG)
            
        Case BindKeys(21).KeyCode
            If UserStats.minman = UserStats.maxman Then Exit Function
            If UserStats.estado = 1 Then
                With FontTypes(FontTypeNames.FONTTYPE_INFO)
                    Call ShowConsoleMsg("¡Estás muerto!", .red, .green, .blue, .bold, .italic)
                End With
                Exit Function
            End If
            Call WriteMeditate
            
        Case BindKeys(22).KeyCode
            Call WriteQuit
    
        Case BindKeys(23).KeyCode
            If UserStats.estado = 1 Then
                With FontTypes(FontTypeNames.FONTTYPE_INFO)
                    Call ShowConsoleMsg("¡Estás muerto!", .red, .green, .blue, .bold, .italic)
                End With
            Else
                Call WriteWork(eSkill.Domar)
            End If
        Case BindKeys(e_KeyAction.eHKey1).KeyCode
            Call DoHotKey(0)
        Case BindKeys(e_KeyAction.eHKey2).KeyCode
            Call DoHotKey(1)
        Case BindKeys(e_KeyAction.eHKey3).KeyCode
            Call DoHotKey(2)
        Case BindKeys(e_KeyAction.eHKey4).KeyCode
            Call DoHotKey(3)
        Case BindKeys(e_KeyAction.eHKey5).KeyCode
            Call DoHotKey(4)
        Case BindKeys(e_KeyAction.eHKey6).KeyCode
            Call DoHotKey(5)
        Case BindKeys(e_KeyAction.eHKey7).KeyCode
            Call DoHotKey(6)
        Case BindKeys(e_KeyAction.eHKey8).KeyCode
            Call DoHotKey(7)
        Case BindKeys(e_KeyAction.eHKey9).KeyCode
            Call DoHotKey(8)
        Case BindKeys(e_KeyAction.eHKey10).KeyCode
            Call DoHotKey(9)
        Case Else
            Accionar = False
            Exit Function

    End Select

    Accionar = True

    
    Exit Function

Accionar_Err:
    Call RegistrarError(Err.Number, Err.Description, "modBindKeys.Accionar", Erl)
    Resume Next
    
End Function

Public Sub DoHotKey(ByVal HkSlot As Byte)
    If UserStats.estado = 1 Then
        With FontTypes(FontTypeNames.FONTTYPE_INFO)
            Call ShowConsoleMsg("¡Estás muerto!", .red, .green, .blue, .bold, .italic)
        End With
    Else
        If IsSet(FeatureToggles, eEnableHotkeys) Then
            If HotkeyList(HkSlot).Index > 0 Then
                Call FormParser.Parse_Form(GetGameplayForm)
                    UsaLanzar = False
                    UsingSkill = 0
                    If CursoresGraficos = 0 Then
                        GetGameplayForm.MousePointer = vbDefault
                    End If
            End If
            Call WriteUseHKeySlot(HkSlot)
        End If
    End If
End Sub

Public Sub TirarItem()
    On Error GoTo TirarItem_Err
    
        If (frmMain.Inventario.SelectedItem > 0 And frmMain.Inventario.SelectedItem < MAX_INVENTORY_SLOTS + 1) Or (frmMain.Inventario.SelectedItem = FLAGORO) Then
            If frmMain.Inventario.Amount(frmMain.Inventario.SelectedItem) = 1 Then
                If ObjData(frmMain.Inventario.ObjIndex(frmMain.Inventario.SelectedItem)).Destruye = 0 Then
                    Call WriteDrop(frmMain.Inventario.SelectedItem, 1)
                Else
                    PreguntaScreen = "El item se destruira al tirarlo ¿Esta seguro?"
                    Pregunta = True

                    DestItemSlot = frmMain.Inventario.SelectedItem
                    DestItemCant = 1
                    PreguntaLocal = True
                    PreguntaNUM = 1
                End If
            Else
                If frmMain.Inventario.Amount(frmMain.Inventario.SelectedItem) > 1 Then
                    frmCantidad.Picture = LoadInterface("cantidad.bmp")
                    frmCantidad.Show , frmMain
                End If
            End If
        End If
    
    Exit Sub

TirarItem_Err:
    Call RegistrarError(Err.Number, Err.Description, "modBindKeys.TirarItem", Erl)
    Resume Next
    
End Sub

Public Sub AgarrarItem()
    
    On Error GoTo AgarrarItem_Err
    
    Call WritePickUp

    
    Exit Sub

AgarrarItem_Err:
    Call RegistrarError(Err.Number, Err.Description, "modBindKeys.AgarrarItem", Erl)
    Resume Next
    
End Sub

Public Function BuscarObjEnInv(ObjIndex) As Byte
    
    On Error GoTo BuscarObjEnInv_Err
    

    'Devuelve el slot del inventario donde se encuentra el obj
    'Creaado por Ladder 25/09/2014
    Dim i As Byte

    For i = 1 To 42

        If frmMain.Inventario.ObjIndex(i) = ObjIndex Then
            BuscarObjEnInv = i
            Exit Function

        End If

    Next i

    BuscarObjEnInv = 0

    
    Exit Function

BuscarObjEnInv_Err:
    Call RegistrarError(Err.Number, Err.Description, "modBindKeys.BuscarObjEnInv", Erl)
    Resume Next
    
End Function

Private Function MouseActionToString(ByVal Action As e_MouseAction) As String
    Select Case Action
        Case e_MouseAction.eThrowOrLook
            MouseActionToString = "THROW_LOOK"
        Case e_MouseAction.eInteract
            MouseActionToString = "INTERACT"
        Case e_MouseAction.eAttack
            MouseActionToString = "ATTACK"
        Case e_MouseAction.eWhisper
            MouseActionToString = "WHISPER"
    End Select
End Function

Private Function ParseMouseAction(ByVal str As String) As e_MouseAction
    Select Case str
        Case "THROW_LOOK"
            ParseMouseAction = e_MouseAction.eThrowOrLook
        Case "INTERACT"
            ParseMouseAction = e_MouseAction.eInteract
        Case "ATTACK"
            ParseMouseAction = e_MouseAction.eAttack
        Case "WHISPER"
            ParseMouseAction = e_MouseAction.eWhisper
        Case Else
            ParseMouseAction = e_MouseAction.eUnknown
    End Select
End Function

Private Function ParseOldMouseAction(ByVal str As String) As e_MouseAction
    If str = vbNullString Then
        ParseOldMouseAction = e_MouseAction.eUnknown
    End If

    Dim Value As Integer
    Value = Val(str)

    Select Case Value
        Case 0
            ParseOldMouseAction = e_MouseAction.eThrowOrLook
        Case 1
            ParseOldMouseAction = e_MouseAction.eInteract
        Case 2
            ParseOldMouseAction = e_MouseAction.eAttack
        Case 4
            ParseOldMouseAction = e_MouseAction.eWhisper
        Case Else
            ParseOldMouseAction = e_MouseAction.eUnknown
    End Select
End Function
