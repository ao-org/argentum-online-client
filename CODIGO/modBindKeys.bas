Attribute VB_Name = "modBindKeys"
'*****************************************************************
'modBindKeys - ImperiumAO - v1.3.0
'
'User input functions.
'
'*****************************************************************
'RevolucionAo 1.0
'Pablo Mercavides
'*****************************************************************
'Augusto José Rando (barrin@imperiumao.com.ar)
'   - First Relase
'*****************************************************************

Option Explicit

Type tBoton

    TipoAccion As Integer
    SendString As String
    hlist As Integer
    invslot As Integer

End Type

Type tBindedKey

    KeyCode As Integer
    name As String

End Type

Public ServerIndex      As Integer

Public NUMBINDS         As Integer

Public ACCION1          As Byte

Public ACCION2          As Byte

Public ACCION3          As Byte

Public BindKeys()       As tBindedKey

Public BotonElegido     As Integer

Public MacroTipoElegido As Byte

Sub LoadDefaultBinds()

    Dim Arch As String, lC As Integer

    Arch = App.Path & "\..\Recursos\OUTPUT\" & "raoinit.ini"

    NUMBINDS = Val(General_Var_Get(Arch, "INIT", "NumBinds"))
    ReDim Preserve BindKeys(1 To NUMBINDS) As tBindedKey

    For lC = 1 To NUMBINDS
        BindKeys(lC).KeyCode = Val(General_Field_Read(1, General_Var_Get(Arch, "DEFAULTS", str(lC)), ","))
        BindKeys(lC).name = General_Field_Read(2, General_Var_Get(Arch, "DEFAULTS", str(lC)), ",")
    Next lC

End Sub

Sub LoadDefaultBinds2()

    Dim Arch As String, lC As Integer

    Arch = App.Path & "\..\Recursos\OUTPUT\" & "raoinit.ini"

    NUMBINDS = Val(General_Var_Get(Arch, "INIT", "NumBinds"))
    ReDim Preserve BindKeys(1 To NUMBINDS) As tBindedKey

    For lC = 1 To NUMBINDS
        BindKeys(lC).KeyCode = Val(General_Field_Read(1, General_Var_Get(Arch, "DEFAULTSMODERN", str(lC)), ","))
        BindKeys(lC).name = General_Field_Read(2, General_Var_Get(Arch, "DEFAULTSMODERN", str(lC)), ",")
    Next lC

End Sub

Public Function Accionar(ByVal KeyCode As Integer) As Boolean
    
    Select Case KeyCode
        Case BindKeys(1).KeyCode
            If UserEstado = 1 Then
    
                With FontTypes(FontTypeNames.FONTTYPE_INFO)
                    Call ShowConsoleMsg("¡Estás muerto!", .red, .green, .blue, .bold, .italic)
    
                End With
    
                Exit Function
    
            End If
    
            If UserDescansar Then Exit Function
            If MainTimer.Check(TimersIndex.CastAttack, False) Then
                If MainTimer.Check(TimersIndex.Attack) Then
                    Call MainTimer.Restart(TimersIndex.AttackSpell)
                    Call MainTimer.Restart(TimersIndex.AttackUse)
                    Call WriteAttack
                End If
    
            End If
    
        Case BindKeys(2).KeyCode
    
            If UserEstado = 1 Then
    
                With FontTypes(FontTypeNames.FONTTYPE_INFO)
                    Call ShowConsoleMsg("¡Estás muerto!", .red, .green, .blue, .bold, .italic)
    
                End With
    
                Exit Function
    
            End If
    
            If Not Comerciando Then
                Call AgarrarItem
            Else
                Call AddtoRichTextBox(frmmain.RecTxt, "No podes agarrar objetos mientras comercias", 255, 0, 32, False, False, False)
    
            End If
    
        Case BindKeys(3).KeyCode
    
            If UserEstado = 1 Then
    
                With FontTypes(FontTypeNames.FONTTYPE_INFO)
                    Call ShowConsoleMsg("¡Estás muerto!", .red, .green, .blue, .bold, .italic)
    
                End With
    
                Exit Function
    
            End If
    
            If Not Comerciando Then
                Call TirarItem
            Else
                Call AddtoRichTextBox(frmmain.RecTxt, "No podes tirar objetos mientras comercias", 255, 0, 32, False, False, False)
    
            End If
    
        Case BindKeys(6).KeyCode
            Call WriteSafeToggle
        Case BindKeys(12).KeyCode
            MostrarOnline = Not MostrarOnline
        Case BindKeys(7).KeyCode
            Nombres = Not Nombres
        Case BindKeys(8).KeyCode
            Call WriteParyToggle
        Case BindKeys(9).KeyCode
    
            If UserEstado = 1 Then
    
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
    
            If UserEstado = 1 Then
    
                With FontTypes(FontTypeNames.FONTTYPE_INFO)
                    Call ShowConsoleMsg("¡Estás muerto!", .red, .green, .blue, .bold, .italic)
    
                End With
    
                Exit Function
    
            End If
    
            If MainTimer.Check(TimersIndex.UseItemWithU) Then
                If frmmain.Inventario.IsItemSelected Then Call WriteEquipItem(frmmain.Inventario.SelectedItem)
            End If
        
        Case BindKeys(4).KeyCode
    
            If Not MainTimer.Check(TimersIndex.UseItemWithU) Then Exit Function
            If Not MainTimer.Check(TimersIndex.AttackUse, False) Then Exit Function
            If frmmain.Inventario.IsItemSelected Then Call WriteUseItem(frmmain.Inventario.SelectedItem)
        
        Case BindKeys(10).KeyCode
    
            If MainTimer.Check(TimersIndex.SendRPU) Then
                Call WriteRequestPositionUpdate
                Beep
    
            End If
        
        Case BindKeys(11).KeyCode
    
            If UserEstado = 1 Then
    
                With FontTypes(FontTypeNames.FONTTYPE_INFO)
                    Call ShowConsoleMsg("¡Estás muerto!", .red, .green, .blue, .bold, .italic)
    
                End With
    
                Exit Function
    
            End If
    
            Call WriteWork(eSkill.Ocultarse)
            
        Case BindKeys(13).KeyCode
        
            Call ScreenCapture
    
        Case BindKeys(12).KeyCode
            'If ShowMacros = 0 Then
            '  ShowMacros = 1
            ' frmMain.picmacroOn.Visible = True
            'frmMain.picmacroOff.Visible = False
            'Else
            '   frmMain.picmacroOn.Visible = False
            '  frmMain.picmacroOff.Visible = True
            ' ShowMacros = 0
            'End If
        Case BindKeys(19).KeyCode
            'FPSFLAG = Not FPSFLAG
            
            ' If FPSFLAG Then
            ' frmMain.Timerping.Enabled = True
            ' Else
            '  frmMain.Timerping.Enabled = False
            'End If
            
            Dim Arch As String
    
            Arch = App.Path & "\..\Recursos\OUTPUT\" & "raoinit.ini"
            Call WriteVar(Arch, "OPCIONES", "FPSFLAG", FPSFLAG)
            
        Case BindKeys(21).KeyCode
    
            'If UserMinMAN = UserMaxMAN Then Exit Function
                
            If UserEstado = 1 Then
    
                With FontTypes(FontTypeNames.FONTTYPE_INFO)
                    Call ShowConsoleMsg("¡Estás muerto!", .red, .green, .blue, .bold, .italic)
    
                End With
    
                Exit Function
    
            End If
            
            Call WriteMeditate
            
        Case BindKeys(22).KeyCode
            Call WriteQuit
    
        Case BindKeys(23).KeyCode
            If UserEstado = 1 Then
                With FontTypes(FontTypeNames.FONTTYPE_INFO)
                    'Call ShowConsoleMsg("¡Estás muerto!", .red, .green, .blue, .bold, .italic)
                End With
            Else
                'Call WriteWork(eSkill.Domar)
            End If
    
        Case Else
            Accionar = False
            Exit Function

    End Select

    Accionar = True

End Function

Public Sub TirarItem()

    If (frmmain.Inventario.SelectedItem > 0 And frmmain.Inventario.SelectedItem < MAX_INVENTORY_SLOTS + 1) Or (frmmain.Inventario.SelectedItem = FLAGORO) Then
        If frmmain.Inventario.Amount(frmmain.Inventario.SelectedItem) = 1 Then
        
            If ObjData(frmmain.Inventario.OBJIndex(frmmain.Inventario.SelectedItem)).Destruye = 0 Then
                Call WriteDrop(frmmain.Inventario.SelectedItem, 1)
            Else
                PreguntaScreen = "El item se destruira al tirarlo ¿Esta seguro?"
                Pregunta = True
                DestItemSlot = frmmain.Inventario.SelectedItem
                DestItemCant = 1
                PreguntaLocal = True
                PreguntaNUM = 1

            End If

        Else

            If frmmain.Inventario.Amount(frmmain.Inventario.SelectedItem) > 1 Then
                frmCantidad.Picture = LoadInterface("cantidad.bmp")
                HayFormularioAbierto = True
                frmCantidad.Show , frmmain

            End If

        End If

    End If

End Sub

Public Sub AgarrarItem()
    Call WritePickUp

End Sub

Public Function BuscarObjEnInv(OBJIndex) As Byte

    'Devuelve el slot del inventario donde se encuentra el obj
    'Creaado por Ladder 25/09/2014
    Dim i As Byte

    For i = 1 To 42

        If frmmain.Inventario.OBJIndex(i) = OBJIndex Then
            BuscarObjEnInv = i
            Exit Function

        End If

    Next i

    BuscarObjEnInv = 0

End Function

