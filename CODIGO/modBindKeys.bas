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

Public ACCION1          As Byte

Public ACCION2          As Byte

Public ACCION3          As Byte

Public BindKeys()       As tBindedKey

Public BotonElegido     As Integer

Public MacroTipoElegido As Byte

Private Enum e_KeyAction
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
    
    [eMaxBinds]
End Enum
Public Sub LoadImpAoInit()
    
    On Error GoTo LoadImpAoInit_Err
    

    Windows_Temp_Dir = General_Get_Temp_Dir

    Dim File As String
    Call InitializeKeyMapping
    File = App.Path & "\..\Recursos\OUTPUT\" & "Teclas.ini"

    Dim lC As Integer, tmpStr As String

    NUMBINDS = eMaxBinds - 1

    ACCION1 = 0 'Val(GetVar(File, "INIT", "ACCION1"))
    ACCION2 = 1 'Val(GetVar(File, "INIT", "ACCION2"))
    ACCION3 = 4 'Val(GetVar(File, "INIT", "ACCION3"))

    ReDim Preserve BindKeys(1 To NUMBINDS) As tBindedKey

    lC = 0

    For lC = 1 To NUMBINDS
        tmpStr = General_Var_Get(File, "USER", str(lC))
        BindKeys(lC).KeyCode = Val(General_Field_Read(1, tmpStr, ","))
        BindKeys(lC).Name = General_Field_Read(2, tmpStr, ",")
    Next lC

    
    Exit Sub

LoadImpAoInit_Err:
    Call RegistrarError(Err.Number, Err.Description, "ModUtils.LoadImpAoInit", Erl)
    Resume Next
    
End Sub

Public Sub SaveRAOInit()
    
    On Error GoTo SaveRAOInit_Err
    

    Dim lC As Integer, Arch As String

    Arch = App.Path & "\..\Recursos\OUTPUT\" & "Teclas.ini"

    Call General_Var_Write(Arch, "INIT", "NUMBINDS", Int(NUMBINDS))

    Call General_Var_Write(Arch, "INIT", "ACCION1", ACCION1)
    Call General_Var_Write(Arch, "INIT", "ACCION2", ACCION2)
    Call General_Var_Write(Arch, "INIT", "ACCION3", ACCION3)

    For lC = 1 To NUMBINDS
        Call General_Var_Write(Arch, "User", str(lC), str(BindKeys(lC).KeyCode) & "," & BindKeys(lC).Name)
    Next lC

    lC = 0

    
    Exit Sub

SaveRAOInit_Err:
    Call RegistrarError(Err.Number, Err.Description, "ModUtils.SaveRAOInit", Erl)
    Resume Next
    
End Sub

Sub LoadDefaultBinds()
    
    On Error GoTo LoadDefaultBinds_Err
    

    Dim Arch As String, lC As Integer

    Arch = App.Path & "\..\Recursos\OUTPUT\" & "Teclas.ini"

    NUMBINDS = Val(General_Var_Get(Arch, "INIT", "NumBinds"))
    ReDim Preserve BindKeys(1 To NUMBINDS) As tBindedKey

    For lC = 1 To NUMBINDS
        BindKeys(lC).KeyCode = Val(General_Field_Read(1, General_Var_Get(Arch, "DEFAULTS", str(lC)), ","))
        BindKeys(lC).Name = General_Field_Read(2, General_Var_Get(Arch, "DEFAULTS", str(lC)), ",")
    Next lC

    
    Exit Sub

LoadDefaultBinds_Err:
    Call RegistrarError(Err.Number, Err.Description, "modBindKeys.LoadDefaultBinds", Erl)
    Resume Next
    
End Sub

Sub LoadDefaultBinds2()
    
    On Error GoTo LoadDefaultBinds2_Err
    

    Dim Arch As String, lC As Integer

    Arch = App.Path & "\..\Recursos\OUTPUT\" & "Teclas.ini"

    NUMBINDS = Val(General_Var_Get(Arch, "INIT", "NumBinds"))
    ReDim Preserve BindKeys(1 To NUMBINDS) As tBindedKey

    For lC = 1 To NUMBINDS
        BindKeys(lC).KeyCode = Val(General_Field_Read(1, General_Var_Get(Arch, "DEFAULTSMODERN", str(lC)), ","))
        BindKeys(lC).Name = General_Field_Read(2, General_Var_Get(Arch, "DEFAULTSMODERN", str(lC)), ",")
    Next lC

    
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
                    If BabelInitialized Then Call BabelUI.ActivateInterval(e_CdTypes.e_Melee)
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
                Call AddtoRichTextBox(frmMain.RecTxt, "No podes agarrar objetos mientras comercias", 255, 0, 32, False, False, False)
    
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
                Call AddtoRichTextBox(frmMain.RecTxt, "No podes tirar objetos mientras comercias", 255, 0, 32, False, False, False)
    
            End If
    
        Case BindKeys(6).KeyCode
            If SeguroGame Then
                Call AddtoRichTextBox(frmMain.RecTxt, "Para desactivar el seguro escribe /SEG o usa el botón en la pestaña MENU en la esquina inferior derecha.", 255, 0, 0, True, False, False)
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
    If BabelInitialized Then
        If UserInventory.SelectedSlot > 0 And UserInventory.SelectedSlot <= MAX_INVENTORY_SLOTS Then
            With UserInventory.Slots(UserInventory.SelectedSlot)
                If .Amount = 1 Then
                    If ObjData(.ObjIndex).Destruye = 0 Then
                        Call WriteDrop(UserInventory.SelectedSlot, 1)
                    Else
                        If BabelInitialized Then
                            Call ShowQuestion("El item se destruira al tirarlo ¿Esta seguro?")
                        Else
                            PreguntaScreen = "El item se destruira al tirarlo ¿Esta seguro?"
                            Pregunta = True
                        End If
                        DestItemSlot = UserInventory.SelectedSlot
                        DestItemCant = 1
                        PreguntaLocal = True
                        PreguntaNUM = 1
                    End If
                ElseIf .Amount > 1 Then
                    frmCantidad.Picture = LoadInterface("cantidad.bmp")
                    frmCantidad.Show , frmBabelUI
                End If
            End With
        ElseIf UserInventory.SelectedSlot = FLAGORO Then
            frmCantidad.Picture = LoadInterface("cantidad.bmp")
                    frmCantidad.Show , frmBabelUI
        End If
    Else
        If (frmMain.Inventario.SelectedItem > 0 And frmMain.Inventario.SelectedItem < MAX_INVENTORY_SLOTS + 1) Or (frmMain.Inventario.SelectedItem = FLAGORO) Then
            If frmMain.Inventario.Amount(frmMain.Inventario.SelectedItem) = 1 Then
                If ObjData(frmMain.Inventario.ObjIndex(frmMain.Inventario.SelectedItem)).Destruye = 0 Then
                    Call WriteDrop(frmMain.Inventario.SelectedItem, 1)
                Else
                    If BabelInitialized Then
                        Call ShowQuestion("El item se destruira al tirarlo ¿Esta seguro?")
                    Else
                        PreguntaScreen = "El item se destruira al tirarlo ¿Esta seguro?"
                        Pregunta = True
                    End If
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

Public Function BuscarObjEnInv(OBJIndex) As Byte
    
    On Error GoTo BuscarObjEnInv_Err
    

    'Devuelve el slot del inventario donde se encuentra el obj
    'Creaado por Ladder 25/09/2014
    Dim i As Byte

    For i = 1 To 42

        If frmMain.Inventario.OBJIndex(i) = OBJIndex Then
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

