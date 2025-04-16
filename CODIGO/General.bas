Attribute VB_Name = "Mod_General"
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

Private Declare Function SetDllDirectory Lib "kernel32" Alias "SetDllDirectoryA" (ByVal path As String) As Long
Private Declare Function svb_init_steam Lib "steam_vb.dll" (ByVal appid As Long) As Long
Private Declare Sub svb_run_callbacks Lib "steam_vb.dll" ()
Private Declare Function svb_retlong Lib "steam_vb.dll" (ByVal Number As Long) As Long
Public Declare Function svb_unlock_achivement Lib "steam_vb.dll" (ByVal Name As String) As Long

Private Type Position

    x As Integer
    y As Integer

End Type

'Item type
Private Type tItem

    ObjIndex As Integer
    Amount As Integer

End Type


Private Type tWorldPos

    map As Integer
    x As Integer
    y As Integer

End Type

Private Type grh

    GrhIndex As Long
    framecounter As Single
    speed As Single
    started As Long
    alpha_blend As Boolean
    angle As Single

End Type

Private Type GrhData

    sX As Integer
    sY As Integer
    FileNum As Integer
    pixelWidth As Integer
    pixelHeight As Integer
    TileWidth As Single
    TileHeight As Single
    NumFrames As Integer
    Frames() As Integer
    speed As Integer
    mini_map_color As Long

End Type

Private Declare Sub InitCommonControls Lib "comctl32" ()

Public bFogata As Boolean

Public ServerIpCount As Integer
'Very percise counter 64bit system counter
Public Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Currency) As Long

Public Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Currency) As Long
'debemos mostrar la animacion de la lluvia

Private lFrameTimer              As Long

'Scroll de richtbox
Public Type SCROLLINFO
    cbSize As Long
    fMask As Long
    nMin As Long
    nMax As Long
    nPage As Long
    nPos As Long
    nTrackPos As Long
End Type


Public Const EM_GETTHUMB = &HBE
Public Const SB_THUMBPOSITION = &H4
Public Const WM_VSCROLL = &H115
Public Const SB_VERT As Integer = &H1
Public Const SIF_RANGE As Integer = &H1
Public Const SIF_PAGE As Integer = &H2
Public Const SIF_POS As Integer = &H4

Public Const SIF_DISABLENOSCROLL = &H8
Public Const SIF_TRACKPOS = &H10
Public Const SIF_ALL = (SIF_RANGE Or SIF_PAGE Or SIF_POS Or SIF_TRACKPOS)

Private tSI As SCROLLINFO

Public Declare Function GetScrollInfo Lib "user32" (ByVal hwnd As Long, ByVal n As Long, ByRef lpScrollInfo As SCROLLINFO) As Long

Public Declare Function GetScrollPos Lib "user32" (ByVal hwnd As Long, ByVal nBar As Long) As Long

'Api SendMessage
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Public Const MAX_TAB_STOPS = 32&

Public Type PARAFORMAT2
    'Los primeros campos coinciden con PARAFORMAT y se usan igual
    cbSize As Integer
    wPad1 As Integer
    dwMask As Long
    wNumbering As Integer
    wEffects As Integer 'No usado en PARAFORMAT
    dxStartIndent As Long
    dxRightIndent As Long
    dxOffset As Long
    wAlignment As Integer
    cTabCount As Integer
    lTabStops(0 To MAX_TAB_STOPS - 1) As Long
    ' Desde aquí lo añadido por PARAFORMAT2
    dySpaceBefore As Long '/* Vertical spacing before para */
    dySpaceAfter As Long '/* Vertical spacing after para */
    dyLineSpacing As Long '/* Line spacing depending on Rule */
    sStyle As Integer ' /* Style handle */
    bLineSpacingRule As Byte '/* Rule for line spacing (see tom.doc) */
    bOutlineLevel As Byte '/* Outline Level*/'antes bCRC As Byte
    wShadingWeight As Integer '/* Shading in hundredths of a per cent */
    wShadingStyle As Integer '/* Byte 0: style, nib 2: cfpat, 3: cbpat*/
    wNumberingStart As Integer '/* Starting value for numbering */
    wNumberingStyle As Integer ' /* Alignment, Roman/Arabic, (), ), ., etc.*/
    wNumberingTab As Integer '/* Space bet 1st indent and 1st-line text*/
    wBorderSpace As Integer ' /* Border-text spaces (nbl/bdr in pts) */
    wBorderWidth As Integer '/* Pen widths (nbl/bdr in half twips) */
    wBorders As Integer '/* Border styles (nibble/border) */
End Type

Public Const EM_LINEFROMCHAR = &HC9
Public Const EM_LINEINDEX = &HBB

Public Const EM_SETPARAFORMAT = &H447
Public Const PFM_LINESPACING = &H100&

Public hlst As clsGraphicalList
Public SelectedSpellSlot As Integer
Public FirstSpellInListToRender As Integer

Public Function DirGraficos() As String
    
    On Error GoTo DirGraficos_Err
    
    DirGraficos = App.path & "\..\Recursos\Graficos\"

    
    Exit Function

DirGraficos_Err:
    Call RegistrarError(Err.Number, Err.Description, "Mod_General.DirGraficos", Erl)
    Resume Next
    
End Function

Public Function DirSound() As String
    
    On Error GoTo DirSound_Err
    
    DirSound = App.path & "\..\Recursos\wav\"

    
    Exit Function

DirSound_Err:
    Call RegistrarError(Err.Number, Err.Description, "Mod_General.DirSound", Erl)
    Resume Next
    
End Function

Public Function DirMidi() As String
    
    On Error GoTo DirMidi_Err
    
    DirMidi = App.path & "\..\Recursos\midi\"

    
    Exit Function

DirMidi_Err:
    Call RegistrarError(Err.Number, Err.Description, "Mod_General.DirMidi", Erl)
    Resume Next
    
End Function

Public Function DirMapas() As String
    
    On Error GoTo DirMapas_Err
    
    DirMapas = App.path & "\..\Recursos\mapas\"

    
    Exit Function

DirMapas_Err:
    Call RegistrarError(Err.Number, Err.Description, "Mod_General.DirMapas", Erl)
    Resume Next
    
End Function


Public Function RandomNumber(ByVal LowerBound As Long, ByVal UpperBound As Long) As Long
    'Initialize randomizer
    
    On Error GoTo RandomNumber_Err
    
    Randomize Timer
    
    'Generate random number
    RandomNumber = (UpperBound - LowerBound) * Rnd + LowerBound

    
    Exit Function

RandomNumber_Err:
    Call RegistrarError(Err.Number, Err.Description, "Mod_General.RandomNumber", Erl)
    Resume Next
    
End Function

Sub AddtoRichTextBox2(ByRef RichTextBox As RichTextBox, ByVal Text As String, Optional ByVal red As Integer = -1, Optional ByVal green As Integer, Optional ByVal blue As Integer, Optional ByVal bold As Boolean = False, Optional ByVal italic As Boolean = False, Optional ByVal bCrLf As Boolean = True, Optional ByVal Alignment As Byte = rtfLeft)
    
    On Error GoTo AddtoRichTextBox2_Err
    
    
    '****************************************************
    'Adds text to a Richtext box at the bottom.
    'Automatically scrolls to new text.
    'Text box MUST be multiline and have a 3D apperance!
    '****************************************************
    'Pablo (ToxicWaste) 01/26/2007 : Now the list refeshes properly.
    'Juan Martin Sotuyo Dodero (Maraxus) 03/29/2007 : Replaced ToxicWaste's code for extra performance.
    'Jopi 17/08/2019 : Consola transparente.
    'Jopi 17/08/2019 : Ahora podes especificar el alineamiento del texto.
    'Ladder 17/12/20 : agrego que la barra no se nos baje si estamos haciedno scroll. Gracias barrin tkm
    '****************************************************
    
        Dim bUrl As Boolean
        Dim sMax As Long
        Dim sPos As Long
        Dim Pos As Long
        Dim ret As Long
        
        Dim bHoldBar As Boolean

    Call EnableURLDetect(frmMain.RecTxt.hwnd, frmMain.hwnd)

    With RichTextBox
        
        If Len(.Text) > 20000 Then
        
            'Get rid of first line
            .Text = vbNullString
            .SelStart = InStr(1, .Text, vbCrLf) + 1
            .SelLength = Len(.Text) - .SelStart + 2
            .TextRTF = .SelRTF

        End If
        
        tSI.cbSize = Len(tSI)
        tSI.fMask = SIF_TRACKPOS Or SIF_RANGE Or SIF_PAGE
        ret = GetScrollInfo(.hwnd, SB_VERT, tSI)
        sMax = tSI.nMax - tSI.nPage + 1
        Pos = tSI.nTrackPos
        Call GetScrollInfo(.hwnd, SB_VERT, tSI)
        bHoldBar = ((((tSI.nMax) - tSI.nPage) > tSI.nTrackPos) And tSI.nPage > 0)
        
        .SelStart = Len(.Text)
        .SelLength = 0
        .SelBold = bold
        .SelItalic = italic
        
        ' 0 = Left
        ' 1 = Center
        ' 2 = Right
        .SelAlignment = Alignment

        If Not red = -1 Then .SelColor = RGB(red, green, blue)
        
        If bCrLf And Len(.Text) > 0 Then Text = vbCrLf & Text
        
        .SelText = Text

        ' Esto arregla el bug de las letras superponiendose la consola del frmMain
        If Not (RichTextBox = frmMain.RecTxt) Then
            RichTextBox.Refresh
        End If
        
        If bHoldBar Then
            Call SendMessage(.hwnd, WM_VSCROLL, SB_THUMBPOSITION + &H10000 * tSI.nTrackPos, Nothing)
        End If

    End With
    
    Exit Sub

AddtoRichTextBox2_Err:
    Call RegistrarError(Err.Number, Err.Description, "Mod_General.AddtoRichTextBox2", Erl)
    Resume Next
    
End Sub

Sub AddtoRichTextBox(ByRef RichTextBox As RichTextBox, ByVal Text As String, Optional ByVal red As Integer = -1, Optional ByVal green As Integer, _
                     Optional ByVal blue As Integer, Optional ByVal bold As Boolean = False, Optional ByVal italic As Boolean = False, _
                     Optional ByVal bCrLf As Boolean = False)
    
    On Error GoTo AddtoRichTextBox_Err
    Dim bUrl As Boolean
    Dim sMax As Long
    Dim sPos As Long
    Dim Pos As Long
    Dim ret As Long
    
    Dim bHoldBar As Boolean
    Call EnableURLDetect(frmMain.RecTxt.hwnd, frmMain.hwnd)

    With RichTextBox

        If Len(.Text) > 20000 Then
            .Text = vbNullString
            .SelStart = InStr(1, .Text, vbCrLf) + 1
            .SelLength = Len(.Text) - .SelStart + 2
            .TextRTF = .SelRTF
        End If
        
        tSI.cbSize = Len(tSI)
        tSI.fMask = SIF_TRACKPOS Or SIF_RANGE Or SIF_PAGE
        ret = GetScrollInfo(.hwnd, SB_VERT, tSI)
        sMax = tSI.nMax - tSI.nPage + 1
        Pos = tSI.nTrackPos
        Call GetScrollInfo(.hwnd, SB_VERT, tSI)
         bHoldBar = ((((tSI.nMax) - tSI.nPage) > tSI.nTrackPos) And tSI.nPage > 0)
        .SelStart = Len(.Text)
        .SelLength = 0
        .SelBold = bold
        .SelItalic = italic
        
        If Not red = -1 Then .SelColor = RGB(red, green, blue)
        bCrLf = True
        
        If bCrLf And Len(.Text) > 0 Then Text = vbCrLf & Text
        .SelText = Text
        
        If bHoldBar Then
            Call SendMessage(.hwnd, WM_VSCROLL, SB_THUMBPOSITION + &H10000 * tSI.nTrackPos, Nothing)
        End If
    End With
    
   ' If bUrl Then DisableUrlDetect
    
    Exit Sub

AddtoRichTextBox_Err:
    Call RegistrarError(Err.Number, Err.Description, "Mod_General.AddtoRichTextBox", Erl)
    Resume Next
    
End Sub

' WyroX: Copiado desde https://www.vbforums.com/showthread.php?727119-RESOLVED-VB2010-richtextbox-paragraph-space-width-seleted-and-RichTextBoxStreamType
Public Sub SelLineSpacing(rtbTarget As RichTextBox, ByVal SpacingRule As Long, Optional ByVal LineSpacing As Long = 20)
    ' SpacingRule
    ' Type of line spacing. To use this member, set the PFM_SPACEAFTER flag in the dwMask member. This member can be one of the following values.
    ' 0 - Single spacing. The dyLineSpacing member is ignored.
    ' 1 - One-and-a-half spacing. The dyLineSpacing member is ignored.
    ' 2 - Double spacing. The dyLineSpacing member is ignored.
    ' 3 - The dyLineSpacing member specifies the spacingfrom one line to the next, in twips. However, if dyLineSpacing specifies a value that is less than single spacing, the control displays single-spaced text.
    ' 4 - The dyLineSpacing member specifies the spacing from one line to the next, in twips. The control uses the exact spacing specified, even if dyLineSpacing specifies a value that is less than single spacing.
    ' 5 - The value of dyLineSpacing / 20 is the spacing, in lines, from one line to the next. Thus, setting dyLineSpacing to 20 produces single-spaced text, 40 is double spaced, 60 is triple spaced, and so on.

    Dim Para As PARAFORMAT2

    With Para
        .cbSize = Len(Para)
        .dwMask = PFM_LINESPACING
        .bLineSpacingRule = SpacingRule
        .dyLineSpacing = LineSpacing
    End With

    Dim ret As Long
    ret = SendMessage(rtbTarget.hwnd, EM_SETPARAFORMAT, 0&, Para)
    
    If ret = 0 Then frmDebug.add_text_tracebox "Error al setear el espaciado entre líneas del RichTextBox."
End Sub

Public Sub RefreshAllChars()
        On Error GoTo RefreshAllChars_Err
        'Goes through the charlist and replots all the characters on the map
        'Used to make sure everyone is visible
        Dim loopC As Long
100     For loopC = 1 To LastChar
102         If charlist(loopC).active = 1 Then
104            If charlist(loopC).Invisible Then
106                 If Not ((charlist(UserCharIndex).clan_nivel < 6 Or charlist(loopC).clan_index = 0 Or charlist(loopC).clan_index <> charlist(UserCharIndex).clan_index) And Not charlist(loopC).Navegando) And _
                        Not (General_Distance_Get(charlist(loopC).Pos.x, charlist(loopC).Pos.y, UserPos.x, UserPos.y) > DISTANCIA_ENVIO_DATOS And _
                        charlist(loopC).dialog_life = 0 And charlist(loopC).FxCount = 0 And charlist(loopC).particle_count = 0) Then

108                     MapData(charlist(loopC).Pos.x, charlist(loopC).Pos.y).CharIndex = loopC

                    End If
                Else
110                 MapData(charlist(loopC).Pos.x, charlist(loopC).Pos.y).CharIndex = loopC
                End If
            End If
112     Next loopC
        Exit Sub
RefreshAllChars_Err:
114     Call RegistrarError(Err.Number, Err.Description, "Mod_General.RefreshAllChars", Erl)
116     Resume Next
End Sub

Function AsciiValidos(ByVal cad As String) As Boolean
    
    On Error GoTo AsciiValidos_Err
    

    Dim car As Byte

    Dim i   As Long
    
    cad = LCase$(cad)
    
    For i = 1 To Len(cad)
        car = Asc(mid$(cad, i, 1))
        
        If ((car < 97 Or car > 122) Or car = Asc("º")) And (car <> 255) And (car <> 32) Then
            Exit Function

        End If

    Next i
    
    AsciiValidos = True

    
    Exit Function

AsciiValidos_Err:
    Call RegistrarError(Err.Number, Err.Description, "Mod_General.AsciiValidos", Erl)
    Resume Next
    
End Function

Function ValidDescriptionCharacters(ByVal cad As String) As Boolean

    On Error GoTo ValidDescriptionCharacters_Err

    Dim car As Byte
    Dim i   As Integer

    For i = 1 To Len(cad)
        car = Asc(mid$(cad, i, 1))
        ' If character is not printable
        If (car < 32 Or car >= 127) And (car < 160) Then
            Exit Function
        End If
    Next i

    ValidDescriptionCharacters = True

    Exit Function

ValidDescriptionCharacters_Err:
    Call RegistrarError(Err.Number, Err.Description, "Mod_General.ValidDescriptionCharacters", Erl)
    Resume Next

End Function

Function CheckUserDataLoged() As Boolean
    'Validamos los datos del user
    
    On Error GoTo CheckUserDataLoged_Err
    
    
    If CuentaEmail = "" Or Not CheckMailString(CuentaEmail) Then
        Call TextoAlAsistente(JsonLanguage.Item("MENSAJEBOX_EMAIL_INVALIDO"), False, False)
        Exit Function

    End If
    
       
    If CuentaPassword = "" Then
        Call TextoAlAsistente(JsonLanguage.Item("MENSAJEBOX_INGRESE_CONTRASENA"), False, False)
        Exit Function
    End If
    
    CheckUserDataLoged = True

    
    Exit Function

CheckUserDataLoged_Err:
    Call RegistrarError(Err.Number, Err.Description, "Mod_General.CheckUserDataLoged", Erl)
    Resume Next
    
End Function

Function CheckUserData(ByVal checkemail As Boolean) As Boolean
    
    On Error GoTo CheckUserData_Err
    

    'Validamos los datos del user
    Dim loopC     As Long

    Dim CharAscii As Integer
    
    If CuentaEmail = "" Or Not CheckMailString(CuentaEmail) Then
        Call TextoAlAsistente(JsonLanguage.Item("MENSAJEBOX_EMAIL_INVALIDO"), False, False)
        Exit Function

    End If
    
    If CuentaPassword = "" Then
        MsgBox (JsonLanguage.Item("MENSAJEBOX_INGRESE_PASSWORD"))
        Exit Function

    End If
    
    For loopC = 1 To Len(CuentaPassword)
        CharAscii = Asc(mid$(CuentaPassword, loopC, 1))

        If Not LegalCharacter(CharAscii) Then
            MsgBox (JsonLanguage.Item("MENSAJEBOX_PASSWORD_INVALIDO") & Chr$(CharAscii) & JsonLanguage.Item("MENSAJEBOX_NO_PERMITIDO"))
            Exit Function

        End If

    Next loopC
    
    CheckUserData = True

    
    Exit Function

CheckUserData_Err:
    Call RegistrarError(Err.Number, Err.Description, "Mod_General.CheckUserData", Erl)
    Resume Next
    
End Function

Sub UnloadAllForms()
    
    On Error GoTo UnloadAllForms_Err
    

    

    Dim mifrm As Form
    
    For Each mifrm In Forms

        Unload mifrm
    Next
    
    
    Exit Sub

UnloadAllForms_Err:
    Call RegistrarError(Err.Number, Err.Description, "Mod_General.UnloadAllForms", Erl)
    Resume Next
    
End Sub

Function LegalCharacter(ByVal KeyAscii As Integer) As Boolean
    
    On Error GoTo LegalCharacter_Err
    

    '*****************************************************************
    'Only allow characters that are Win 95 filename compatible
    '*****************************************************************
    'if backspace allow
    If KeyAscii = 8 Then
        LegalCharacter = True
        Exit Function

    End If
    
    'Only allow space, numbers, letters and special characters
    If KeyAscii < 32 Or KeyAscii = 44 Then
        Exit Function

    End If
    
    If KeyAscii > 126 Then
        Exit Function

    End If
    
    'Check for bad special characters in between
    If KeyAscii = 34 Or KeyAscii = 42 Or KeyAscii = 47 Or KeyAscii = 58 Or KeyAscii = 60 Or KeyAscii = 62 Or KeyAscii = 63 Or KeyAscii = 92 Or KeyAscii = 124 Then
        Exit Function

    End If
    
    'else everything is cool
    LegalCharacter = True

    
    Exit Function

LegalCharacter_Err:
    Call RegistrarError(Err.Number, Err.Description, "Mod_General.LegalCharacter", Erl)
    Resume Next
    
End Function

Sub SetConnected()
    '*****************************************************************
    'Sets the client to "Connect" mode
    '*****************************************************************
    'Set Connected
    
    On Error GoTo SetConnected_Err
    
    Connected = True
    Call frmConnect.AuthSocket.Close
    Call ModGameplayUI.SetupGameplayUI
    
    Seguido = False
    CharindexSeguido = 0
    OffsetLimitScreen = 32
    AlphaNiebla = 0

    'Vaciamos la cola de movimiento
    Call keysMovementPressedQueue.Clear
    frmMain.UpdateDaytime.enabled = True
    light_transition = 1#
    COLOR_AZUL = RGB(0, 0, 0)
    OpcionMenu = 0
    Call ResetContadores
    frmMain.cerrarcuenta.enabled = True
    engine.FadeInAlpha = 255
    isLogged = True
  
    If newUser Then
         If MostrarTutorial And tutorial_index <= 0 Then
            If tutorial(e_tutorialIndex.TUTORIAL_NUEVO_USER).Activo = 1 Then
                tutorial_index = e_tutorialIndex.TUTORIAL_NUEVO_USER
                Call mostrarCartel(tutorial(tutorial_index).titulo, tutorial(tutorial_index).textos(1), tutorial(tutorial_index).grh, -1, &H164B8A, , , False, 100, 479, 100, 535, 640, 490, 50, 100)
            End If
        End If
    End If
    Exit Sub

SetConnected_Err:
    Call RegistrarError(Err.Number, Err.Description, "Mod_General.SetConnected", Erl)
    Resume Next
    
End Sub
Sub ResetContadores()
    packetCounters.TS_CastSpell = 0
    packetCounters.TS_WorkLeftClick = 0
    packetCounters.TS_LeftClick = 0
    packetCounters.TS_UseItem = 0
    packetCounters.TS_UseItemU = 0
    packetCounters.TS_Walk = 0
    packetCounters.TS_Talk = 0
    packetCounters.TS_Attack = 0
    packetCounters.TS_Drop = 0
    packetCounters.TS_Work = 0
    packetCounters.TS_EquipItem = 0
    packetCounters.TS_GuildMessage = 0
    packetCounters.TS_QuestionGM = 0
    packetCounters.TS_ChangeHeading = 0
   
End Sub

Sub MoveTo(ByVal Heading As E_Heading, ByVal Dumb As Boolean)
    On Error GoTo MoveTo_Err
    If Dumb Then
        If RandomNumber(1, 100) < 50 Then
            Dim newHeading As E_Heading
            Do
                newHeading = RandomNumber(E_Heading.NORTH, E_Heading.WEST)
            Loop Until newHeading <> Heading
            Heading = newHeading
        End If
    End If
    Dim LegalOk As Boolean
    
    If cartel Then cartel = False
    
    Select Case Heading

        Case E_Heading.NORTH
            LegalOk = LegalPos(UserPos.x, UserPos.y - 1, Heading)

        Case E_Heading.EAST
            LegalOk = LegalPos(UserPos.x + 1, UserPos.y, Heading)

        Case E_Heading.south
            LegalOk = LegalPos(UserPos.x, UserPos.y + 1, Heading)

        Case E_Heading.WEST
            LegalOk = LegalPos(UserPos.x - 1, UserPos.y, Heading)

    End Select

    If LegalOk And CanMove() Then
        If Not UserDescansar Then
            If UserMacro.Activado Then
                Call ResetearUserMacro
            End If

            Moviendose = True
            Call MainTimer.Restart(TimersIndex.Walk)
            
            If PescandoEspecial Then
                Call AddtoRichTextBox(frmMain.RecTxt, JsonLanguage.Item("MENSAJE_PEZ_ROMPIO_LINEA_PESCA"), 255, 0, 0, 1, 0)
                Call WriteRomperCania
                PescandoEspecial = False
            End If
           
            If EstaSiguiendo Then Exit Sub
            
            Call WriteWalk(Heading) 'We only walk if we are not meditating or resting

            Call Char_Move_by_Head(UserCharIndex, Heading)
            Call MoveScreen(Heading)
            Call checkTutorial
            
            Dim i As Integer
            For i = 1 To LastChar
                If charlist(i).Invisible And Not EsGM And Not charlist(i).Meditating Then
                    If MapData(charlist(i).Pos.x, charlist(i).Pos.y).CharIndex = i And (charlist(UserCharIndex).clan_nivel < 6 Or charlist(i).clan_index = 0 Or charlist(i).clan_index <> charlist(UserCharIndex).clan_index) And Not charlist(i).Navegando Then
                        If General_Distance_Get(charlist(i).Pos.x, charlist(i).Pos.y, UserPos.x, UserPos.y) > DISTANCIA_ENVIO_DATOS And charlist(i).dialog_life = 0 And charlist(i).FxCount = 0 And charlist(i).particle_count = 0 Then
                            MapData(charlist(i).Pos.x, charlist(i).Pos.y).CharIndex = 0
                        End If
                    End If
                End If
            Next i
        Else

            If Not UserAvisado Then
                If UserDescansar Then
                    WriteRest 'Stop resting (we do NOT have the 1 step enforcing anymore) sono como un tema de los guns.
                End If

                UserAvisado = True

            End If

        End If

    Else

        If charlist(UserCharIndex).Heading <> Heading Then
            If IntervaloPermiteHeading(True) Then
                Call WriteChangeHeading(Heading)
            End If
        End If

    End If
    
    Call UpdateMapPos
    
    ' Update 3D sounds!
    ' Call Audio.MoveListener(UserPos.x, UserPos.y)
    If frmMain.macrotrabajo.enabled Then frmMain.DesactivarMacroTrabajo
    
    
    Exit Sub

MoveTo_Err:
    Call RegistrarError(Err.Number, Err.Description, "Mod_General.MoveTo", Erl)
    Resume Next
    
End Sub
Public Function EstaSiguiendo() As Boolean
      If CharindexSeguido > 0 Then
            'Call AddtoRichTextBox(frmMain.RecTxt, "No puedes moverte mientras estás revisando a un usuario.", 255, 0, 0, 1)
            EstaSiguiendo = True
            Exit Function
        End If
End Function

Public Sub AddMovementToKeysMovementPressedQueue()
    
    On Error GoTo AddMovementToKeysMovementPressedQueue_Err
    
    If BindKeys(14).KeyCode <> 0 And GetKeyState(BindKeys(14).KeyCode) < 0 Then
        If keysMovementPressedQueue.itemExist(BindKeys(14).KeyCode) = False Then keysMovementPressedQueue.Add (BindKeys(14).KeyCode) ' Agrega la tecla al arraylist
    Else

        If keysMovementPressedQueue.itemExist(BindKeys(14).KeyCode) Then keysMovementPressedQueue.Remove (BindKeys(14).KeyCode) ' Remueve la tecla que teniamos presionada

    End If

    If BindKeys(15).KeyCode <> 0 And GetKeyState(BindKeys(15).KeyCode) < 0 Then
        If keysMovementPressedQueue.itemExist(BindKeys(15).KeyCode) = False Then keysMovementPressedQueue.Add (BindKeys(15).KeyCode) ' Agrega la tecla al arraylist
    Else

        If keysMovementPressedQueue.itemExist(BindKeys(15).KeyCode) Then keysMovementPressedQueue.Remove (BindKeys(15).KeyCode) ' Remueve la tecla que teniamos presionada

    End If

    If BindKeys(16).KeyCode <> 0 And GetKeyState(BindKeys(16).KeyCode) < 0 Then
        If keysMovementPressedQueue.itemExist(BindKeys(16).KeyCode) = False Then keysMovementPressedQueue.Add (BindKeys(16).KeyCode) ' Agrega la tecla al arraylist
    Else

        If keysMovementPressedQueue.itemExist(BindKeys(16).KeyCode) Then keysMovementPressedQueue.Remove (BindKeys(16).KeyCode) ' Remueve la tecla que teniamos presionada

    End If

    If BindKeys(17).KeyCode <> 0 And GetKeyState(BindKeys(17).KeyCode) < 0 Then
        If keysMovementPressedQueue.itemExist(BindKeys(17).KeyCode) = False Then keysMovementPressedQueue.Add (BindKeys(17).KeyCode) ' Agrega la tecla al arraylist
    Else

        If keysMovementPressedQueue.itemExist(BindKeys(17).KeyCode) Then keysMovementPressedQueue.Remove (BindKeys(17).KeyCode) ' Remueve la tecla que teniamos presionada

    End If

    
    Exit Sub

AddMovementToKeysMovementPressedQueue_Err:
    Call RegistrarError(Err.Number, Err.Description, "Mod_General.AddMovementToKeysMovementPressedQueue", Erl)
    Resume Next
    
End Sub

Sub Check_Keys()
    
    On Error GoTo Check_Keys_Err
    

    

    Static lastMovement As Long

    Dim direccion As E_Heading
    'Debug.Assert UserCharIndex > 0
    
    direccion = charlist(UserCharIndex).Heading

    If Not Application.IsAppActive() Then Exit Sub
    

    If Not pausa And _
        g_game_state.State = e_state_gameplay_screen And _
        Not frmComerciarUsu.visible And _
        Not frmBancoObj.visible And _
        Not frmOpciones.visible And _
        Not frmComerciar.visible And _
        Not frmGoliath.visible And _
        Not frmEstadisticas.visible And _
        Not frmStatistics.visible And _
        Not frmAlqui.visible And _
        Not frmCarp.visible And _
        Not frmHerrero.visible And _
        Not FrmGrupo.visible And _
        Not FrmSastre.visible And _
        Not FrmGmAyuda.visible And _
        Not frmCrafteo.visible And _
        Not IsGameDialogOpen Then
 
        If IsInputFocus() And PermitirMoverse = 0 Then Exit Sub
 
        If Not UserMoving Then
            Call AddMovementToKeysMovementPressedQueue
            Select Case keysMovementPressedQueue.GetLastItem()
                ' Prevenimos teclas sin asignar... Te deja moviendo para siempre
                Case 0: Exit Sub
                'Move Up
                Case BindKeys(14).KeyCode
                    Call MoveTo(E_Heading.NORTH, UserEstupido)
                'Move Right
                Case BindKeys(17).KeyCode
                    Call MoveTo(E_Heading.EAST, UserEstupido)
                'Move down
                Case BindKeys(15).KeyCode
                    Call MoveTo(E_Heading.south, UserEstupido)
                'Move left
                Case BindKeys(16).KeyCode
                    Call MoveTo(E_Heading.WEST, UserEstupido)
            End Select
        End If
    End If
    Exit Sub

Check_Keys_Err:
    Call RegistrarError(Err.Number, Err.Description, "Mod_General.Check_Keys", Erl)
    Resume Next
    
End Sub

Function ReadField(ByVal Pos As Integer, ByRef Text As String, ByVal SepASCII As Byte) As String
    
    On Error GoTo ReadField_Err
    

    '*****************************************************************
    'Gets a field from a delimited string
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modify Date: 11/15/2004
    '*****************************************************************
    Dim i          As Long

    Dim LastPos    As Long

    Dim CurrentPos As Long

    Dim delimiter  As String * 1
    
    delimiter = Chr$(SepASCII)
    
    For i = 1 To Pos
        LastPos = CurrentPos
        CurrentPos = InStr(LastPos + 1, Text, delimiter, vbBinaryCompare)
    Next i
    
    If CurrentPos = 0 Then
        ReadField = mid$(Text, LastPos + 1, Len(Text) - LastPos)
    Else
        ReadField = mid$(Text, LastPos + 1, CurrentPos - LastPos - 1)

    End If

    
    Exit Function

ReadField_Err:
    Call RegistrarError(Err.Number, Err.Description, "Mod_General.ReadField", Erl)
    Resume Next
    
End Function

Function FieldCount(ByRef Text As String, ByVal SepASCII As Byte) As Long
    
    On Error GoTo FieldCount_Err
    

    '*****************************************************************
    'Gets the number of fields in a delimited string
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modify Date: 07/29/2007
    '*****************************************************************
    Dim count     As Long

    Dim curPos    As Long

    Dim delimiter As String * 1
    
    If LenB(Text) = 0 Then Exit Function
    
    delimiter = Chr$(SepASCII)
    
    curPos = 0
    
    Do
        curPos = InStr(curPos + 1, Text, delimiter)
        count = count + 1
    Loop While curPos <> 0
    
    FieldCount = count

    
    Exit Function

FieldCount_Err:
    Call RegistrarError(Err.Number, Err.Description, "Mod_General.FieldCount", Erl)
    Resume Next
    
End Function

Function FileExist(ByVal File As String, ByVal FileType As VbFileAttribute) As Boolean
    
    On Error GoTo FileExist_Err
    
    FileExist = (dir$(File, FileType) <> "")

    
    Exit Function

FileExist_Err:
    Call RegistrarError(Err.Number, Err.Description, "Mod_General.FileExist", Erl)
    Resume Next
    
End Function

    
Public Sub SaveStringInFile(ByVal Cadena As String, ByVal nombreArchivo As String)
On Error GoTo ErrorHandler
    Dim fileNumber As Integer
    fileNumber = FreeFile
    Open nombreArchivo For Append As fileNumber
    Print #fileNumber, Now & " " & Cadena ' O usa vbNewLine en lugar de vbCrLf si lo prefieres
    Close #fileNumber
    Exit Sub
ErrorHandler:
End Sub

Sub parse_cmd_line_args()

#If REMOTE_CLOSE = 1 Then
    Call Application.DeleteFile("remote_debug.txt")
    IPdelServidorLogin = "127.0.0.1"
    PuertoDelServidorLogin = 4000
    IPdelServidor = IPdelServidorLogin
    PuertoDelServidor = 6501
    CuentaEmail = "some@yahoo.com.ar"
    CuentaPassword = "secret"
    CharacterRemote = "rolo"
    Dim sArgs() As String
    Dim iLoop As Integer
    sArgs = Split(command$, " ")
    For iLoop = 0 To UBound(sArgs)
        frmDebug.add_text_tracebox sArgs(iLoop)
        Dim Value() As String
        Value = Split(sArgs(iLoop), "=")
        
        If Value(0) = "account" Then
              CuentaEmail = Value(1)
        ElseIf Value(0) = "password" Then
            CuentaPassword = Value(1)
        ElseIf Value(0) = "serverip" Then
            IPdelServidorLogin = Value(1)
            IPdelServidor = Value(1)
        ElseIf Value(0) = "lport" Then
            PuertoDelServidorLogin = Value(1)
        ElseIf Value(0) = "gport" Then
            PuertoDelServidor = Value(1)
        ElseIf Value(0) = "pc" Then
            CharacterRemote = Value(1)
        End If
        
    Next
    Call SaveStringInFile("Using IPdelServidorLogin: " & IPdelServidorLogin, "remote_debug.txt")
    Call SaveStringInFile("Using PuertoDelServidorLogin: " & PuertoDelServidorLogin, "remote_debug.txt")
    Call SaveStringInFile("Using IPdelServidor: " & IPdelServidor, "remote_debug.txt")
    Call SaveStringInFile("Using PuertoDelServidor: " & PuertoDelServidor, "remote_debug.txt")
    Call SaveStringInFile("Using CuentaEmail: " & CuentaEmail, "remote_debug.txt")
    Call SaveStringInFile("Using CuentaPassword: " & CuentaPassword, "remote_debug.txt")
    Call SaveStringInFile("Using CharacterRemote: " & CharacterRemote, "remote_debug.txt")
#End If

End Sub


Sub Main()

On Error GoTo Main_Err

    Call parse_cmd_line_args
    
#If REMOTE_CLOSE Then
    Call DoLogin("", "", False)
    Call bot_main_loop
    End
#End If

    Call Application.DeleteFile(ao20config.GetErrorLogFilename())
    Call LoadConfig
    Call SetLanguageApplication
    Call Frmcarga.Show
    Set FormParser = New clsCursor
    Call FormParser.Init
    Call CheckResources
    If Not ValidateResources Then
        Call MsgBox(JsonLanguage.Item("MENSAJEBOX_RECURSOS_INVALIDOS"), vbApplicationModal + vbInformation + vbOKOnly, JsonLanguage.Item("MENSAJEBOX_TITULO_RECURSOS_INVALIDOS"))
        End
    End If
    If PantallaCompleta Then
        Call Resolution.SetResolution
    End If
#If EXPERIMENTAL_RENDERER Then
    Call new_engine_init(ao20rendering.renderer)
#Else
    Call engine_init 'initializes DX
    Debug.Assert Not DirectX Is Nothing
    Call ao20audio.CreateAudioEngine(frmConnect.hwnd, DirectX, ao20audio.AudioEngine)
#End If
    Call InitCommonControls

    #If DEBUGGING = 0 Or ENABLE_ANTICHEAT = 1 Then
        SetDllDirectory App.path
        Dim steam_init_result As Long
        steam_init_result = svb_init_steam(1956740)
        frmDebug.add_text_tracebox "Init Steam " & steam_init_result
        If Not RunningInVB Then
            If FindPreviousInstance Then
                Call MsgBox(JsonLanguage.Item("MENSAJEBOX_ERROR_EJECUCION"), vbApplicationModal + vbInformation + vbOKOnly, "Error")
                End
            End If
 
        End If
 
    #End If


    Call initPacketControl
    
    Call SetNpcsRenderText
    Call cargarTutoriales
    Call InitializeEffectArrays
    
    #If DEBUGGING = 0 Then
        CheckMD5 = GetMd5
    #Else
        CheckMD5 = "NotNeededinDebug"
    #End If
    
    SessionOpened = False
    
    Call Load(frmConnect)
    Call Load(FrmLogear)
    
    Windows_Temp_Dir = General_Get_Temp_Dir

    Call SetDefaultServer
    Call ComprobarEstado
    Call CargarLst
    Call InicializarNombres
    Call InitializeInventory
    Call Init_TileEngine
    Call CargarRecursos
    Call LoadFonts
    Call initMascotaTutorial
    Call LoadProjectiles
    Call LoadBuffResources
    Call InitilializeProjectiles
    Call InitializeTeamColors
    Call InitializeAntiCheat
    FrameTime = GetTickCount()
    UserMap = 1
    AlphaNiebla = 75
    EntradaY = 10
    EntradaX = 10
    UpdateLights = False
    LastOffset2X = 0
    LastOffset2Y = 0
    Call SwitchMap(UserMap)
    
    Dialogos.font = frmMain.font
    DialogosClanes.font = frmMain.font
    
    prgRun = True
    pausa = False

    Call Unload(Frmcarga)
    Call General_Set_Connect
    Call engine.GetElapsedTime
    Call Start

    Set AudioEngine = Nothing

    
    Exit Sub

Main_Err:
    If Err.Number = 339 Then
        RegisterCom
    End If
    
    Call RegistrarError(Err.Number, Err.Description, "Mod_General.Main", Erl)
    Resume Next
    
End Sub

Public Sub RegisterCom()
    On Error GoTo Com_Err:
    If MsgBox(JsonLanguage.Item("MENSAJEBOX_COMPONENTES_FALTANTES"), vbYesNo) = vbYes Then
            If System.ShellExecuteEx("regcom.bat", App.path) Then
                Call MsgBox(JsonLanguage.Item("MENSAJEBOX_ARCHIVOS_COM_REGISTRADOS"), vbOKOnly, "Info")
            Else
                Call MsgBox(JsonLanguage.Item("MENSAJEBOX_ARCHIVOS_COM_NO_REGISTRADOS"), vbOKOnly, "Error")
            End If
        End If
        End
Com_Err:
    Call RegistrarError(Err.Number, Err.Description, "Mod_General.RegisterCom", Erl)
    Resume Next
End Sub

Public Function SetDefaultServer()
On Error GoTo SetDefaultServer_Err

#If PYMMO = 1 And Developer = 1 Then
    IPdelServidorLogin = "127.0.0.1"
    PuertoDelServidorLogin = 4000
    IPdelServidor = IPdelServidorLogin
    PuertoDelServidor = 7667
#Else
    
    Call SetActiveEnvironment("Production")
#End If
    Exit Function
SetDefaultServer_Err:
    Call RegistrarError(Err.Number, Err.Description, "Mod_General.WriteVar", Erl)
End Function

Public Function randomMap() As Integer
    Select Case RandomNumber(1, 8)
        Case 1 ' ulla 45-43
            randomMap = 1
        Case 2 ' nix 22-75
            randomMap = 34
        Case 3 ' bander 49-43
            randomMap = 59
        Case 4 ' Arghal 38-41
            randomMap = 151
        Case 5 ' Lindos 63-40
            randomMap = 62
        Case 6 ' Arkhein 64-32
            randomMap = 195
        Case 7 ' Esperanza 50-45
            randomMap = 112
        Case 8 ' Polo 78-66
            randomMap = 354
    End Select
End Function

Sub WriteVar(ByVal File As String, ByVal Main As String, ByVal Var As String, ByVal Value As String)
    '*****************************************************************
    'Writes a var to a text file
    '*****************************************************************
    
    On Error GoTo WriteVar_Err
    
    writeprivateprofilestring Main, Var, Value, File

    
    Exit Sub

WriteVar_Err:
    Call RegistrarError(Err.Number, Err.Description, "Mod_General.WriteVar", Erl)
    Resume Next
    
End Sub

Function GetVar(ByVal File As String, ByVal Main As String, ByVal Var As String) As String
    
    On Error GoTo GetVar_Err
    

    '*****************************************************************
    'Gets a Var from a text file
    '*****************************************************************
    Dim sSpaces As String ' This will hold the input that the program will retrieve
    
    sSpaces = Space$(100) ' This tells the computer how long the longest string can be. If you want, you can change the number 100 to any number you wish
    
    getprivateprofilestring Main, Var, vbNullString, sSpaces, Len(sSpaces), File
    
    GetVar = RTrim$(sSpaces)
    GetVar = Left$(GetVar, Len(GetVar) - 1)

    
    Exit Function

GetVar_Err:
    Call RegistrarError(Err.Number, Err.Description, "Mod_General.GetVar", Erl)
    Resume Next
    
End Function

Function GetVarOrDefault(ByVal File As String, ByVal Main As String, ByVal Var As String, ByVal DefaultValue As String) As String
    
    On Error GoTo GetVarOrDefault_Err
    

    '*****************************************************************
    'Gets a Var from a text file and if empty returns default value
    '*****************************************************************

    GetVarOrDefault = GetVar(File, Main, Var)
    If GetVarOrDefault = vbNullString Then
        GetVarOrDefault = DefaultValue
    End If
    
    Exit Function

GetVarOrDefault_Err:
    Call RegistrarError(Err.Number, Err.Description, "Mod_General.GetVarOrDefault", Erl)
    Resume Next
    
End Function

'[CODE 002]:MatuX
'
'  Función para chequear el email
'
'  Corregida por Maraxus para que reconozca como válidas casillas con puntos antes de la arroba y evitar un chequeo innecesario
Public Function CheckMailString(ByVal sString As String) As Boolean

    On Error GoTo errHnd

    Dim lPos As Long

    Dim lX   As Long

    Dim iAsc As Integer
    
    '1er test: Busca un simbolo @
    lPos = InStr(sString, "@")

    If (lPos <> 0) Then

        '2do test: Busca un simbolo . después de @ + 1
        If Not (InStr(lPos, sString, ".", vbBinaryCompare) > lPos + 1) Then Exit Function
        
        '3er test: Recorre todos los caracteres y los valída
        For lX = 0 To Len(sString) - 1

            If Not (lX = (lPos - 1)) Then   'No chequeamos la '@'
                iAsc = Asc(mid$(sString, (lX + 1), 1))

                If Not CMSValidateChar_(iAsc) Then Exit Function

            End If

        Next lX
        
        'Finale
        CheckMailString = True

    End If

errHnd:

End Function

'  Corregida por Maraxus para que reconozca como válidas casillas con puntos antes de la arroba
Private Function CMSValidateChar_(ByVal iAsc As Integer) As Boolean
    
    On Error GoTo CMSValidateChar__Err
    
    CMSValidateChar_ = (iAsc >= 48 And iAsc <= 57) Or (iAsc >= 65 And iAsc <= 90) Or (iAsc >= 97 And iAsc <= 122) Or (iAsc = 95) Or (iAsc = 45) Or (iAsc = 46)

    
    Exit Function

CMSValidateChar__Err:
    Call RegistrarError(Err.Number, Err.Description, "Mod_General.CMSValidateChar_", Erl)
    Resume Next
    
End Function

Public Sub LeerLineaComandos()
    
    On Error GoTo LeerLineaComandos_Err
    

    Dim t() As String

    Dim i   As Long
    
    'Parseo los comandos
    t = Split(command, " ")

    For i = LBound(t) To UBound(t)

        Select Case UCase$(t(i))

            Case "/LAUNCHER" 'no cambiar la resolucion
                Launcher = True
        
            Case "/NORES" 'no cambiar la resolucion
                NoRes = True

        End Select

    Next i

    
    Exit Sub

LeerLineaComandos_Err:
    Call RegistrarError(Err.Number, Err.Description, "Mod_General.LeerLineaComandos", Erl)
    Resume Next
    
End Sub

Private Sub InicializarNombres()
    '**************************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modify Date: 11/27/2005
    'Inicializa los nombres de razas, ciudades, clases, skills, atributos, etc.
    '**************************************************************
    
    On Error GoTo InicializarNombres_Err
    

    ListaRazas(eRaza.Humano) = "Humano"
    ListaRazas(eRaza.Elfo) = "Elfo"
    ListaRazas(eRaza.ElfoOscuro) = "Elfo Oscuro"
    ListaRazas(eRaza.Gnomo) = "Gnomo"
    ListaRazas(eRaza.Enano) = "Enano"
    ListaRazas(eRaza.Orco) = "Orco"
        
    ListaCiudades(eCiudad.cUllathorpe) = "Ullathorpe"
    ListaCiudades(eCiudad.cNix) = "Nix"
    ListaCiudades(eCiudad.cBanderbill) = "Banderbill"
    ListaCiudades(eCiudad.cLindos) = "Lindos"
    ListaCiudades(eCiudad.cArghal) = "Arghal"
    ListaCiudades(eCiudad.cForgat) = "Forgat"

    ListaClases(eClass.Mage) = "Mago"
    ListaClases(eClass.Cleric) = "Clérigo"
    ListaClases(eClass.Warrior) = "Guerrero"
    ListaClases(eClass.Assasin) = "Asesino"
    ListaClases(eClass.Bard) = "Bardo"
    ListaClases(eClass.Druid) = "Druida"
    ListaClases(eClass.paladin) = "Paladín"
    ListaClases(eClass.Hunter) = "Cazador"
    ListaClases(eClass.Trabajador) = "Trabajador"
    ListaClases(eClass.Pirat) = "Pirata"
    ListaClases(eClass.Thief) = "Ladrón"
    ListaClases(eClass.Bandit) = "Bandido"

    SkillsNames(eSkill.magia) = "Magia"
    SkillsNames(eSkill.Robar) = "Robar"
    SkillsNames(eSkill.Tacticas) = "Destreza en combate"
    SkillsNames(eSkill.Armas) = "Combate con armas"
    SkillsNames(eSkill.Meditar) = "Meditar"
    SkillsNames(eSkill.Apuñalar) = "Apuñalar"
    SkillsNames(eSkill.Ocultarse) = "Ocultarse"
    SkillsNames(eSkill.Supervivencia) = "Supervivencia"
    SkillsNames(eSkill.Comerciar) = "Comercio"
    SkillsNames(eSkill.Defensa) = "Defensa con escudo"
    SkillsNames(eSkill.Liderazgo) = "Liderazgo"
    SkillsNames(eSkill.Proyectiles) = "Armas a distancia"
    SkillsNames(eSkill.Wrestling) = "Combate sin armas"
    SkillsNames(eSkill.Navegacion) = "Navegación"
    SkillsNames(eSkill.equitacion) = "Equitación"
    SkillsNames(eSkill.Resistencia) = "Resistencia mágica"
    SkillsNames(eSkill.Talar) = "Tala"
    SkillsNames(eSkill.Pescar) = "Pesca"
    SkillsNames(eSkill.Mineria) = "Minería"
    SkillsNames(eSkill.Herreria) = "Herrería"
    SkillsNames(eSkill.Carpinteria) = "Carpintería"
    SkillsNames(eSkill.Alquimia) = "Alquimia"
    SkillsNames(eSkill.Sastreria) = "Sastrería"
    SkillsNames(eSkill.Domar) = "Domar"

    SkillsDesc(eSkill.magia) = "Los hechizos requieren un cierto número de puntos mágicos para ser usados. Sube lanzando cualquier hechizo."
    SkillsDesc(eSkill.Robar) = "Aumenta las posibilidades de conseguir objetos u oro mientras robas. Se sube robando. Solo el ladrón puede robar objetos, las otras clases solo pueden robar oro."
    SkillsDesc(eSkill.Tacticas) = "Aumenta la posibilidad de esquivar ataques. Cuantos más puntos tengas, mejor será tu evasión. Sube mientras peleas cuerpo a cuerpo."
    SkillsDesc(eSkill.Armas) = "Aumenta las posibilidades de golpear al enemigo con un arma.Subes peleando cuerpo a cuerpo usando cualquier arma."
    SkillsDesc(eSkill.Meditar) = "Aumenta la cantidad de mana que recuperamos al meditar. Se sube meditando. Al aumentar los puntos de esta habilidad, aumenta la mana que se recupera."
    SkillsDesc(eSkill.Apuñalar) = "Aumenta la probabilidad de apuñalar. Se sube luchando cuerpo a cuerpo con dagas. Mientras mas skill tengas, mas posibilidad de apuñalar."
    SkillsDesc(eSkill.Ocultarse) = "Esta habilidad es responsable de aumentar las posibilidades de esconderse. Se sube tratando de esconderse. Mientras mas skills, mas tiempo oculto. "
    SkillsDesc(eSkill.Supervivencia) = "La supervivencia nos permitirá tomar agua de ríos, comer de los árboles y ver la vida de los NPCs Hostiles. También aumenta la velocidad que recuperamos energía o sanamos. Con 30 puntos podemos beber de los rios, con 40 puntos podemos comer de los arboles, con 50 puntos vemos el estado de los demas personajes y el tiempo exacto que le queda de paralizis a una criatura, con 75 puntos vemos la vida exacta de los npcs. Se sube combatiendo con las criaturas o prendiendo fogatas."
    SkillsDesc(eSkill.Comerciar) = "Cuanto más puntos en comerciar tengas más baratas te saldrán las cosas en las tiendas. Sube tanto al comprar como al vender items a NPCs."
    SkillsDesc(eSkill.Defensa) = "Aumenta las chances de defenderte con un escudo, mientras más puntos tengas, hay más probabilidad de rechazar el golpe del adversario."
    SkillsDesc(eSkill.Liderazgo) = "Es la habilidad necesaria para crear un clan. Se sube manualmente."
    SkillsDesc(eSkill.Proyectiles) = "Aumenta las probabilidades de pegarle al enemigo con un arco."
    SkillsDesc(eSkill.Wrestling) = "Aumenta las probabilidades de impactar al enemigo en la lucha sin armas, estupidizar o paralizar."
    SkillsDesc(eSkill.Navegacion) = "Necesaria para poder utilizar traje de baño, barcas, galeras o galeones."
    SkillsDesc(eSkill.equitacion) = " Necesaria para equipar una montura."
    SkillsDesc(eSkill.Resistencia) = "Sirve para que los hechizos no te peguen tan fuerte, mientras más puntos tengas, menos es el daño mágico que recibes. Se sube cuando un NPC o una persona te ataca con hechizos."
    SkillsDesc(eSkill.Talar) = "Aumenta la velocidad a la que recoletas madera de los árboles."
    SkillsDesc(eSkill.Pescar) = "Aumenta la velocidad a la que capturas peces."
    SkillsDesc(eSkill.Mineria) = "Aumenta la velocidad a la que extraes minerales de los yacimientos."
    SkillsDesc(eSkill.Herreria) = "Te permite construir mejores objetos de herrería."
    SkillsDesc(eSkill.Carpinteria) = "Te permite construir mejores objetos de carpintería."
    SkillsDesc(eSkill.Alquimia) = "Te permite crear pociones más poderosas."
    SkillsDesc(eSkill.Sastreria) = "Te permite confeccionar mejores vestimentas."
    SkillsDesc(eSkill.Domar) = "Aumenta tu habilidad para domar animales."
    
    AtributosNames(eAtributos.Fuerza) = "Fuerza"
    AtributosNames(eAtributos.Agilidad) = "Agilidad"
    AtributosNames(eAtributos.Inteligencia) = "Inteligencia"
    AtributosNames(eAtributos.Constitucion) = "Constitucion"
    AtributosNames(eAtributos.Carisma) = "Carisma"

    
    Exit Sub

InicializarNombres_Err:
    Call RegistrarError(Err.Number, Err.Description, "Mod_General.InicializarNombres", Erl)
    Resume Next
    
End Sub

''
' Removes all text from the console and dialogs

Public Sub CleanDialogs()
    '**************************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modify Date: 11/27/2005
    'Removes all text from the console and dialogs
    '**************************************************************
    'Clean console and dialogs
    'frmMain.RecTxt.Text = vbNullString
    
    On Error GoTo CleanDialogs_Err
    If (Not DialogosClanes Is Nothing) Then
    Call DialogosClanes.RemoveDialogs
    End If
    If (Not Dialogos Is Nothing) Then
    Call Dialogos.RemoveAllDialogs
    End If
    
    
    Exit Sub

CleanDialogs_Err:
    Call RegistrarError(Err.Number, Err.Description, "Mod_General.CleanDialogs", Erl)
    Resume Next
    
End Sub

Public Sub CloseClient()
    '**************************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modify Date: 8/14/2007
    'Frees all used resources, cleans up and leaves
    '**************************************************************
    ' Allow new instances of the client to be opened
    
    On Error GoTo CloseClient_Err
    UserSaliendo = True

    Call SaveConfig
    Call UnloadAntiCheat
    Call PrevInstance.ReleaseInstance
  
    
    ao20audio.StopAllPlayback
    EngineRun = False
    
    Call General_Set_Mouse_Speed(SensibilidadMouseOriginal)
    Call Resolution.ResetResolution

    Set SurfaceDB = Nothing
    Set Dialogos = Nothing
    Set DialogosClanes = Nothing
    ' Set Audio = Nothing
    Set MainTimer = Nothing

    Set FormParser = Nothing
    Call EndGame(True)
    
    ' Destruyo los inventarios
    Set frmMain.Inventario = Nothing
    Set frmComerciar.InvComNpc = Nothing
    Set frmComerciar.InvComUsu = Nothing
    Set frmBancoObj.InvBankUsu = Nothing
    Set frmBancoObj.InvBoveda = Nothing
    Set frmComerciarUsu.InvUser = Nothing
    
    
    Set frmBancoCuenta.InvBankUsuCuenta = Nothing
    Set frmBancoCuenta.InvBovedaCuenta = Nothing
    
    Set FrmKeyInv.InvKeys = Nothing
      Call Client_UnInitialize_DirectX_Objects
   Exit Sub

CloseClient_Err:
    Call RegistrarError(Err.Number, Err.Description, "Mod_General.CloseClient", Erl)
    Resume Next
    
End Sub

Public Function General_Field_Read(ByVal field_pos As Long, ByVal Text As String, ByVal delimiter As String) As String
    
    On Error GoTo General_Field_Read_Err
    

    '*****************************************************************
    'Author: Juan Martín Sotuyo Dodero
    'Last Modify Date: 11/15/2004
    'Gets a field from a delimited string
    '*****************************************************************
    Dim i          As Long

    Dim LastPos    As Long

    Dim CurrentPos As Long
    
    LastPos = 0
    CurrentPos = 0
    
    For i = 1 To field_pos
        LastPos = CurrentPos
        CurrentPos = InStr(LastPos + 1, Text, delimiter, vbBinaryCompare)
    Next i
    
    If CurrentPos = 0 Then
        General_Field_Read = mid$(Text, LastPos + 1, Len(Text) - LastPos)
    Else
        General_Field_Read = mid$(Text, LastPos + 1, CurrentPos - LastPos - 1)

    End If

    
    Exit Function

General_Field_Read_Err:
    Call RegistrarError(Err.Number, Err.Description, "Mod_General.General_Field_Read", Erl)
    Resume Next
    
End Function

Public Function General_Field_Count(ByVal Text As String, ByVal delimiter As Byte) As Long
    
    On Error GoTo General_Field_Count_Err
    

    '*****************************************************************
    'Author: Aaron Perkins
    'Last Modify Date: 10/07/2002
    'Count the number of fields in a delimited string
    '*****************************************************************
    'If string is empty there aren't any fields
    If Len(Text) = 0 Then
        Exit Function

    End If

    Dim i        As Long

    Dim FieldNum As Long

    FieldNum = 0

    For i = 1 To Len(Text)

        If delimiter = CByte(Asc(mid$(Text, i, 1))) Then
            FieldNum = FieldNum + 1

        End If

    Next i

    General_Field_Count = FieldNum + 1

    
    Exit Function

General_Field_Count_Err:
    Call RegistrarError(Err.Number, Err.Description, "Mod_General.General_Field_Count", Erl)
    Resume Next
    
End Function


Public Function General_Get_Elapsed_Time() As Single
    
    On Error GoTo General_Get_Elapsed_Time_Err
    

    '**************************************************************
    'Author: Aaron Perkins
    'Last Modify Date: 10/07/2002
    'Gets the time that past since the last call
    '**************************************************************
    Dim Start_Time    As Currency

    Static end_time   As Currency

    Static timer_freq As Currency

    'Get the timer frequency
    If timer_freq = 0 Then
        QueryPerformanceFrequency timer_freq

    End If

    'Get current time
    QueryPerformanceCounter Start_Time
    
    'Calculate elapsed time
    General_Get_Elapsed_Time = (Start_Time - end_time) / timer_freq * 1000
    
    'Get next end time
    QueryPerformanceCounter end_time

    
    Exit Function

General_Get_Elapsed_Time_Err:
    Call RegistrarError(Err.Number, Err.Description, "Mod_General.General_Get_Elapsed_Time", Erl)
    Resume Next
    
End Function


Public Function max(ByVal a As Variant, ByVal b As Variant) As Variant
    
    On Error GoTo max_Err
    

    If a > b Then
        max = a
    Else
        max = b

    End If

    
    Exit Function

max_Err:
    Call RegistrarError(Err.Number, Err.Description, "Mod_General.max", Erl)
    Resume Next
    
End Function

Public Function min(ByVal a As Double, ByVal b As Double) As Variant
    
    On Error GoTo min_Err
    

    If a < b Then
        min = a
    Else
        min = b

    End If

    
    Exit Function

min_Err:
    Call RegistrarError(Err.Number, Err.Description, "Mod_General.min", Erl)
    Resume Next
    
End Function

Public Function Clamp(ByVal a As Variant, ByVal min As Variant, ByVal max As Variant) As Variant
    
    On Error GoTo min_Err
    

    If a < min Then
        Clamp = min
    
    ElseIf a > max Then
        Clamp = max

    Else
        Clamp = a
    End If

    
    Exit Function

min_Err:
    Call RegistrarError(Err.Number, Err.Description, "Mod_General.min", Erl)
    Resume Next
    
End Function


Public Function LoadInterface(FileName As String, Optional localize As Boolean = True) As IPicture

On Error GoTo errhandler
    
    If localize Then
        Select Case language
            Case e_language.English
                FileName = "en_" & FileName
            Case e_language.Spanish
                FileName = "es_" & FileName
            Case Else
                FileName = "en_" & FileName
        End Select
    End If
    If FileName <> "" Then
        #If Compresion = 1 Then
            Set LoadInterface = General_Load_Picture_From_Resource_Ex(LCase$(FileName), ResourcesPassword)
        #Else
            Set LoadInterface = LoadPicture(App.path & "/../Recursos/interface/" & LCase$(FileName))
        #End If
    End If
Exit Function

errhandler:
    MsgBox "Error al cargar la interface: " & FileName

End Function

Public Function LoadMinimap(ByVal map As Integer) As IPicture

On Error GoTo errhandler

    #If Compresion = 1 Then
        Set LoadMinimap = General_Load_Minimap_From_Resource_Ex("mapa" & map & ".bmp", ResourcesPassword)
    #Else
        Set LoadMinimap = LoadPicture(App.path & "/../Recursos/Minimapas/Mapa" & map & ".bmp")
    #End If
    
Exit Function

errhandler:
    MsgBox "Error al cargar minimapa: Mapa" & map & ".bmp"

End Function

Public Function Tilde(ByRef Data As String) As String
    
    On Error GoTo Tilde_Err
    

    Tilde = UCase$(Data)
 
    Tilde = Replace$(Tilde, "Á", "A")
    Tilde = Replace$(Tilde, "É", "E")
    Tilde = Replace$(Tilde, "Í", "I")
    Tilde = Replace$(Tilde, "Ó", "O")
    Tilde = Replace$(Tilde, "Ú", "U")
        
    
    Exit Function

Tilde_Err:
    Call RegistrarError(Err.Number, Err.Description, "Mod_General.Tilde", Erl)
    Resume Next
    
End Function

' Copiado de https://www.vbforums.com/showthread.php?231468-VB-Detect-if-you-are-running-in-the-IDE
Function RunningInVB() As Boolean
    'Returns whether we are running in vb(true), or compiled (false)
    
    On Error GoTo RunningInVB_Err
    
 
    Static counter As Variant

    If IsEmpty(counter) Then
        counter = 1
        Debug.Assert RunningInVB() Or True
        counter = counter - 1
    ElseIf counter = 1 Then
        counter = 0

    End If

    RunningInVB = counter
 
    
    Exit Function

RunningInVB_Err:
    Call RegistrarError(Err.Number, Err.Description, "Mod_General.RunningInVB", Erl)
    Resume Next
    
End Function

Function GetTimeFromString(str As String) As Long
    
    On Error GoTo GetTimeFromString_Err
    
    If Len(str) = 0 Then Exit Function

    Dim Splitted() As String
    Splitted = Split(str, ":")
    
    Dim Hour As Long, min As Long
    Hour = Val(Splitted(0))

    If Hour < 0 Then Hour = 0
    If Hour > 23 Then Hour = 23
    
    GetTimeFromString = Hour * 60
    
    If UBound(Splitted) > 0 Then
        min = Val(Splitted(1))
        
        If min < 0 Then min = 0
        If min > 59 Then min = 59
        
        GetTimeFromString = GetTimeFromString + min
    End If

    GetTimeFromString = GetTimeFromString * (DuracionDia / 1440)

    
    Exit Function

GetTimeFromString_Err:
    Call RegistrarError(Err.Number, Err.Description, "Mod_General.GetTimeFromString", Erl)
    Resume Next
    
End Function

Public Function GetMd5() As String

On Error GoTo Handler

    GetMd5 = MD5File(App.path & "\Argentum.exe")
    
    Exit Function
    
Handler:
    Call MsgBox(JsonLanguage.Item("MENSAJEBOX_ERROR_CLIENTE_COMPROBAR"), vbOKOnly, JsonLanguage.Item("MENSAJEBOX_TITULO_CLIENTE_CORROMPIDO"))
    End

End Function

Public Sub CheckResources()

    Dim Data(1 To 200) As Byte
    
    Dim handle As Integer
    handle = FreeFile

    Open App.path & "/../Recursos/OUTPUT/AO.bin" For Binary Access Read As #handle
    
    Get #handle, , Data
    
    Close #handle
    
    Dim Length As Integer
    Length = Data(UBound(Data)) + Data(UBound(Data) - 1) * 256

    Dim i As Integer
    
    For i = 1 To Length
        ResourcesPassword = ResourcesPassword & Chr(Data(i * 3 - 1) Xor 37)
    Next

End Sub

Function ValidarNombre(nombre As String, Error As String) As Boolean

    If Len(nombre) < 3 Or Len(nombre) > 18 Then
        Error = JsonLanguage.Item("ERROR_NOMBRE_LONGITUD_INVALIDA")
        Exit Function
    End If
    
    Dim Temp As String
    Temp = UCase$(nombre)
    
    Dim i As Long, Char As Integer, LastChar As Integer
    For i = 1 To Len(Temp)
        Char = Asc(mid$(Temp, i, 1))
        
        If (Char < 65 Or Char > 90) And Char <> 32 Then
            Error = JsonLanguage.Item("ERROR_CARACTERES_INVALIDOS")
            Exit Function
        
        ElseIf Char = 32 And LastChar = 32 Then
            Error = JsonLanguage.Item("ERROR_ESPACIOS_CONSECUTIVOS")
            Exit Function
        End If
        
        LastChar = Char
    Next

    If Asc(mid$(Temp, 1, 1)) = 32 Or Asc(mid$(Temp, Len(Temp), 1)) = 32 Then
        Error = JsonLanguage.Item("ERROR_ESPACIOS_INICIO_FIN")
        Exit Function
    End If
    
    ValidarNombre = True

End Function

Function BeautifyBigNumber(ByVal Number As Long) As String

    If Number > 1000000000 Then
        BeautifyBigNumber = Round(Number * 0.000000001, 3) & "KKK"
    ElseIf Number > 10000000 Then
        BeautifyBigNumber = Round(Number * 0.000001, 2) & "KK"
    ElseIf Number > 10000& Then
        BeautifyBigNumber = Round(Number * 0.001, 1) & "K"
    Else
        BeautifyBigNumber = Number
    End If

End Function

Public Function IntentarObtenerPezEspecial()
    
    Dim acierto As Byte
    
    frmDebug.add_text_tracebox "Aciertos: " & ContadorIntentosPescaEspecial_Acertados & "Posicion barra : " & PosicionBarra
        'El + y -10 es por inputLag (Margen de error)
    If PuedeIntentar Then
        If PosicionBarra >= (90 - 15) And PosicionBarra <= (111 + 15) Then
            ContadorIntentosPescaEspecial_Acertados = ContadorIntentosPescaEspecial_Acertados + 1
            acierto = 1
        Else
            ContadorIntentosPescaEspecial_Fallados = ContadorIntentosPescaEspecial_Fallados + 1
            acierto = 2
        End If
        
        PuedeIntentar = False
        
        If acierto = 1 Then
            intentosPesca(ContadorIntentosPescaEspecial_Fallados + ContadorIntentosPescaEspecial_Acertados) = 1
        ElseIf acierto = 2 Then
            intentosPesca(ContadorIntentosPescaEspecial_Fallados + ContadorIntentosPescaEspecial_Acertados) = 2
        End If
    
        If ContadorIntentosPescaEspecial_Fallados + ContadorIntentosPescaEspecial_Acertados >= 5 Or ContadorIntentosPescaEspecial_Acertados >= 3 Then
            PescandoEspecial = False
            Call WriteFinalizarPescaEspecial
        ElseIf ContadorIntentosPescaEspecial_Acertados >= 3 Then
            PescandoEspecial = False
            Call WriteFinalizarPescaEspecial
        ElseIf ContadorIntentosPescaEspecial_Fallados >= 3 Then
            PescandoEspecial = False
            Call AddtoRichTextBox(frmMain.RecTxt, JsonLanguage.Item("MENSAJE_PEZ_ROMPIO_LINEA_PESCA"), 255, 0, 0, 1, 0)
            Call WriteRomperCania
        End If
    End If
    
    
    
End Function


Public Function isValidEmail(Email As String) As Boolean
    Dim At As Integer
    Dim oneDot As Integer
    Dim twoDots As Integer
 
    isValidEmail = True
    At = InStr(1, Email, "@", vbTextCompare)
    oneDot = InStr(At + 2, Email, ".", vbTextCompare)
    twoDots = InStr(At + 2, Email, "..", vbTextCompare)
    If At = 0 Or oneDot = 0 Or Not twoDots = 0 Or Right(Email, 1) = "." Then isValidEmail = False
End Function


Public Sub SetNpcsRenderText()

    '************************************************************************************.
    ' Carga el JSON con las traducciones en un objeto para su uso a lo largo del proyecto
    '************************************************************************************
    Dim render_text As String
    render_text = GetSetting("OPCIONES", "NpcsEnRender")
    
    ' Si no se especifica el idioma en el archivo de configuracion, se le pregunta si quiere usar castellano
    ' y escribimos el archivo de configuracion con el idioma seleccionado
    If LenB(render_text) = 0 Then
        npcs_en_render = 1
        Call SaveSetting("OPCIONES", "NpcsEnRender", npcs_en_render)
    Else
       npcs_en_render = Val(render_text)
    End If

End Sub

Public Sub deleteCharIndexs()
    Dim i As Long
    For i = 1 To LastChar
        If charlist(i).EsNpc = False And i <> UserCharIndex Then
            Call EraseChar(i)
        End If
    Next i
End Sub
