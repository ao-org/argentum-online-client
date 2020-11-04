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

Public ServerIndex As Integer
Public NUMBINDS As Integer

Public ACCION1 As Byte
Public ACCION2 As Byte
Public ACCION3 As Byte

Public BindKeys() As tBindedKey
Public BotonElegido As Integer
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

    
    If KeyCode = BindKeys(1).KeyCode Then
        If UserEstado = 1 Then
            With FontTypes(FontTypeNames.FONTTYPE_INFO)
                Call ShowConsoleMsg("¡Estás muerto!", .red, .green, .blue, .bold, .italic)
            End With
            Exit Function
        End If
        If UserDescansar Or UserMeditar Then Exit Function
        If Not IntervaloPermiteComboMagiaGolpe(False) Then Exit Function
        If Not IntervaloPermiteAtacar Then Exit Function
        Call WriteAttack

    ElseIf KeyCode = BindKeys(2).KeyCode Then
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
    ElseIf KeyCode = BindKeys(3).KeyCode Then
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
    ElseIf KeyCode = BindKeys(6).KeyCode Then
       Call WriteSafeToggle
    ElseIf KeyCode = BindKeys(12).KeyCode Then
        MostrarOnline = Not MostrarOnline
    ElseIf KeyCode = BindKeys(7).KeyCode Then
       Nombres = Not Nombres
    ElseIf KeyCode = BindKeys(8).KeyCode Then
        Call WriteParyToggle
    ElseIf KeyCode = BindKeys(9).KeyCode Then
        If UserEstado = 1 Then
            With FontTypes(FontTypeNames.FONTTYPE_INFO)
                Call ShowConsoleMsg("¡Estás muerto!", .red, .green, .blue, .bold, .italic)
            End With
            Exit Function
        End If
        Call WriteWork(eSkill.Robar)
        
    ElseIf KeyCode = BindKeys(18).KeyCode Then
        If IntervaloPermiteLLamadaClan Then Call WriteLlamadadeClan
        
    ElseIf KeyCode = BindKeys(20).KeyCode Then
        If IntervaloPermiteLLamadaClan Then Call WriteMarcaDeClan
    
    ElseIf KeyCode = BindKeys(5).KeyCode Then
        If UserEstado = 1 Then
            With FontTypes(FontTypeNames.FONTTYPE_INFO)
                Call ShowConsoleMsg("¡Estás muerto!", .red, .green, .blue, .bold, .italic)
            End With
            Exit Function
        End If
        If IntervaloPermiteUsar Then Call EquiparItem
    
    ElseIf KeyCode = BindKeys(4).KeyCode Then
        If Not MainTimer.Check(TimersIndex.UseItemWithU) Then Exit Function
        Call UsarItem
    
    ElseIf KeyCode = BindKeys(10).KeyCode Then
        If MainTimer.Check(TimersIndex.SendRPU) Then
                Call WriteRequestPositionUpdate
                Beep
        End If
    
    ElseIf KeyCode = BindKeys(11).KeyCode Then
        If UserEstado = 1 Then
            With FontTypes(FontTypeNames.FONTTYPE_INFO)
                Call ShowConsoleMsg("¡Estás muerto!", .red, .green, .blue, .bold, .italic)
            End With
            Exit Function
        End If
        Call WriteWork(eSkill.Ocultarse)
        
    ElseIf KeyCode = BindKeys(13).KeyCode Then
    
        Call ScreenCapture

     ElseIf KeyCode = BindKeys(12).KeyCode Then
     'If ShowMacros = 0 Then
      '  ShowMacros = 1
       ' frmMain.picmacroOn.Visible = True
        'frmMain.picmacroOff.Visible = False
    'Else
     '   frmMain.picmacroOn.Visible = False
      '  frmMain.picmacroOff.Visible = True
       ' ShowMacros = 0
    'End If
    ElseIf KeyCode = BindKeys(19).KeyCode Then
        'FPSFLAG = Not FPSFLAG
        
       ' If FPSFLAG Then
           ' frmMain.Timerping.Enabled = True
       ' Else
          '  frmMain.Timerping.Enabled = False
        'End If
        
        
        
        Dim Arch As String
        Arch = App.Path & "\..\Recursos\OUTPUT\" & "raoinit.ini"
        Call WriteVar(Arch, "OPCIONES", "FPSFLAG", FPSFLAG)
        
    ElseIf KeyCode = BindKeys(21).KeyCode Then
        If UserMinMAN = UserMaxMAN Then Exit Function
            
        If UserEstado = 1 Then
            With FontTypes(FontTypeNames.FONTTYPE_INFO)
                Call ShowConsoleMsg("¡Estás muerto!", .red, .green, .blue, .bold, .italic)
            End With
            Exit Function
        End If
        
        Call WriteMeditate
        
    ElseIf KeyCode = BindKeys(22).KeyCode Then
        Call WriteQuit

    Else
        Accionar = False
        Exit Function
    End If

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

Public Sub UsarItem()

    'If Not frmComerciar.Visible Then
        If (frmmain.Inventario.SelectedItem > 0) And (frmmain.Inventario.SelectedItem < MAX_INVENTORY_SLOTS + 1) Then _
            Call WriteUseItem(frmmain.Inventario.SelectedItem)
   ' Else
   ' If (InvComUsu.SelectedItem > 0) And (InvComUsu.SelectedItem < MAX_INVENTORY_SLOTS + 1) Then _
      '      Call WriteUseItem(InvComUsu.SelectedItem)
    'End If
        
End Sub

Public Sub EquiparItem()
    'If Not frmComerciar.Visible Then
        If (frmmain.Inventario.SelectedItem > 0) And (frmmain.Inventario.SelectedItem < MAX_INVENTORY_SLOTS + 1) Then _
            Call WriteEquipItem(frmmain.Inventario.SelectedItem)
  ' Else
    'If (InvComUsu.SelectedItem > 0) And (InvComUsu.SelectedItem < MAX_INVENTORY_SLOTS + 1) Then _
    '        Call WriteEquipItem(InvComUsu.SelectedItem)
    'End If
    
End Sub
Public Sub Bind_Accion(ByVal FNUM As Integer)

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



