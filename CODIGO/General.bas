Attribute VB_Name = "Mod_General"

'RevolucionAo 1.0
'Pablo Mercavides

Option Explicit

Private Type Position

    x As Integer
    y As Integer

End Type

'Item type
Private Type tItem

    OBJIndex As Integer
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
    Started As Long
    alpha_blend As Boolean
    Angle As Single

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

'Very percise counter 64bit system counter
Private Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Currency) As Long

Private Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Currency) As Long
'debemos mostrar la animacion de la lluvia

Private lFrameTimer              As Long

'Scroll de richtbox
Private Type SCROLLINFO
    cbSize As Long
    fMask As Long
    nMin As Long
    nMax As Long
    nPage As Long
    nPos As Long
    nTrackPos As Long
End Type


Private Const EM_GETTHUMB = &HBE
Private Const SB_THUMBPOSITION = &H4
Private Const WM_VSCROLL = &H115
Private Const SB_VERT As Integer = &H1
Private Const SIF_RANGE As Integer = &H1
Private Const SIF_PAGE As Integer = &H2
Private Const SIF_POS As Integer = &H4

Private Const SIF_DISABLENOSCROLL = &H8
Private Const SIF_TRACKPOS = &H10
Private Const SIF_ALL = (SIF_RANGE Or SIF_PAGE Or SIF_POS Or SIF_TRACKPOS)
Private tSI As SCROLLINFO

Public Declare Function GetScrollInfo Lib "user32" (ByVal hWnd As Long, ByVal N As Long, ByRef lpScrollInfo As SCROLLINFO) As Long

Public Declare Function GetScrollPos Lib "user32" (ByVal hWnd As Long, ByVal nBar As Long) As Long

'Api SendMessage
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

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


Public Function DirGraficos() As String
    
    On Error GoTo DirGraficos_Err
    
    DirGraficos = App.Path & "\..\Recursos\Graficos\"

    
    Exit Function

DirGraficos_Err:
    Call RegistrarError(Err.Number, Err.Description, "Mod_General.DirGraficos", Erl)
    Resume Next
    
End Function

Public Function DirSound() As String
    
    On Error GoTo DirSound_Err
    
    DirSound = App.Path & "\..\Recursos\wav\"

    
    Exit Function

DirSound_Err:
    Call RegistrarError(Err.Number, Err.Description, "Mod_General.DirSound", Erl)
    Resume Next
    
End Function

Public Function DirMidi() As String
    
    On Error GoTo DirMidi_Err
    
    DirMidi = App.Path & "\..\Recursos\midi\"

    
    Exit Function

DirMidi_Err:
    Call RegistrarError(Err.Number, Err.Description, "Mod_General.DirMidi", Erl)
    Resume Next
    
End Function

Public Function DirMapas() As String
    
    On Error GoTo DirMapas_Err
    
    DirMapas = App.Path & "\..\Recursos\mapas\"

    
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
        Dim Ret As Long
        
        Dim bHoldBar As Boolean

    Call EnableURLDetect(frmMain.RecTxt.hWnd, frmMain.hWnd)

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
        Ret = GetScrollInfo(.hWnd, SB_VERT, tSI)
        sMax = tSI.nMax - tSI.nPage + 1
        Pos = tSI.nTrackPos
        Call GetScrollInfo(.hWnd, SB_VERT, tSI)
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
            Call SendMessage(.hWnd, WM_VSCROLL, SB_THUMBPOSITION + &H10000 * tSI.nTrackPos, Nothing)
        End If

    End With
    
    Exit Sub

AddtoRichTextBox2_Err:
    Call RegistrarError(Err.Number, Err.Description, "Mod_General.AddtoRichTextBox2", Erl)
    Resume Next
    
End Sub

Sub AddtoRichTextBox(ByRef RichTextBox As RichTextBox, ByVal Text As String, Optional ByVal red As Integer = -1, Optional ByVal green As Integer, Optional ByVal blue As Integer, Optional ByVal bold As Boolean = False, Optional ByVal italic As Boolean = False, Optional ByVal bCrLf As Boolean = False, Optional ByVal FontTypeIndex As Byte = 0)
    
    On Error GoTo AddtoRichTextBox_Err
    

    '******************************************
    'Adds text to a Richtext box at the bottom.
    'Automatically scrolls to new text.
    'Text box MUST be multiline and have a 3D
    'apperance!
    'Pablo (ToxicWaste) 01/26/2007 : Now the list refeshes properly.
    'Juan Martín Sotuyo Dodero (Maraxus) 03/29/2007 : Replaced ToxicWaste's code for extra performance.
    'Ladder 17/12/20 agrego que la barra no se nos baje si estamos haciedno scroll. Gracias barrin tkm
    '******************************************r
    Dim bUrl As Boolean
    Dim sMax As Long
    Dim sPos As Long
    Dim Pos As Long
    Dim Ret As Long
    
    Dim bHoldBar As Boolean
    Call EnableURLDetect(frmMain.RecTxt.hWnd, frmMain.hWnd)

    With RichTextBox

        If Len(.Text) > 20000 Then
            .Text = vbNullString
            .SelStart = InStr(1, .Text, vbCrLf) + 1
            .SelLength = Len(.Text) - .SelStart + 2
            .TextRTF = .SelRTF
        End If
        
        tSI.cbSize = Len(tSI)
        tSI.fMask = SIF_TRACKPOS Or SIF_RANGE Or SIF_PAGE
        Ret = GetScrollInfo(.hWnd, SB_VERT, tSI)
        sMax = tSI.nMax - tSI.nPage + 1
        Pos = tSI.nTrackPos
        Call GetScrollInfo(.hWnd, SB_VERT, tSI)
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
            Call SendMessage(.hWnd, WM_VSCROLL, SB_THUMBPOSITION + &H10000 * tSI.nTrackPos, Nothing)
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

    Dim Ret As Long
    Ret = SendMessage(rtbTarget.hWnd, EM_SETPARAFORMAT, 0&, Para)
    
    If Ret = 0 Then Debug.Print "Error al setear el espaciado entre líneas del RichTextBox."
End Sub

'TODO : Never was sure this is really necessary....
'TODO : 08/03/2006 - (AlejoLp) Esto hay que volarlo...
Public Sub RefreshAllChars()
    
    On Error GoTo RefreshAllChars_Err
    

    '*****************************************************************
    'Goes through the charlist and replots all the characters on the map
    'Used to make sure everyone is visible
    '*****************************************************************
    Dim loopc As Long
    
    For loopc = 1 To LastChar
    
        If charlist(loopc).active = 1 Then
            MapData(charlist(loopc).Pos.x, charlist(loopc).Pos.y).charindex = loopc

        End If

    Next loopc

    
    Exit Sub

RefreshAllChars_Err:
    Call RegistrarError(Err.Number, Err.Description, "Mod_General.RefreshAllChars", Erl)
    Resume Next
    
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

Function CheckUserDataLoged() As Boolean
    'Validamos los datos del user
    
    On Error GoTo CheckUserDataLoged_Err
    
    
    If CuentaEmail = "" Or Not CheckMailString(CuentaEmail) Then
        Call TextoAlAsistente("El email es inválido.")
        Exit Function

    End If
    
    ' If Len(UserCuenta) > 30 Then
    '   Call TextoAlAsistente("El nombre debe tener menos de 30 letras.")
    '  frmMensaje.Show vbModal
    '  Exit Function
    '  End If
    
    '  For loopc = 1 To Len(UserCuenta)
    '   CharAscii = Asc(mid$(UserCuenta, loopc, 1))
    ' If Not LegalCharacter(CharAscii) Then
    ' Call TextoAlAsistente("Nombre inválido. El caractér " & Chr$(CharAscii) & " no está permitido.")
    '    Exit Function
    '  End If
    ' Next loopc
    
    If CuentaPassword = "" Then
        Call TextoAlAsistente("Ingrese la contraseña de la cuenta.")
        'frmMensaje.msg.Caption = "Ingrese un password."
        ' frmMensaje.Show vbModal
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
    Dim loopc     As Long

    Dim CharAscii As Integer
    
    If CuentaEmail = "" Or Not CheckMailString(CuentaEmail) Then
        Call TextoAlAsistente("El email es inválido.")
        Exit Function

    End If
    
    If CuentaPassword = "" Then
        MsgBox ("Ingrese un password.")
        Exit Function

    End If
    
    For loopc = 1 To Len(CuentaPassword)
        CharAscii = Asc(mid$(CuentaPassword, loopc, 1))

        If Not LegalCharacter(CharAscii) Then
            MsgBox ("Password inválido. El caractér " & Chr$(CharAscii) & " no está permitido.")
            Exit Function

        End If

    Next loopc
    
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
    
    'Unload the connect form
    'FrmCuenta.Visible = False

    frmMain.NombrePJ.Caption = UserName
    LogeoAlgunaVez = True
    
    ' bTecho = False
    AlphaNiebla = 0

    'Vaciamos la cola de movimiento
    Call keysMovementPressedQueue.Clear

    If FPSFLAG Then
        frmMain.Timerping.Enabled = True
    Else
        frmMain.Timerping.Enabled = False
    End If
    
    frmMain.UpdateLight.Enabled = True
    frmMain.UpdateDaytime.Enabled = True
    light_transition = 1#

    COLOR_AZUL = RGB(0, 0, 0)
    
    ' establece el borde al listbox
    Call Establecer_Borde(frmMain.hlst, frmMain, COLOR_AZUL, 0, 0)

    Call Make_Transparent_Richtext(frmMain.RecTxt.hWnd)
   
    ' Detect links in console
    Call EnableURLDetect(frmMain.RecTxt.hWnd, frmMain.hWnd)
        
    ' Removemos la barra de titulo pero conservando el caption para la barra de tareas
    Call Form_RemoveTitleBar(frmMain)

    OpcionMenu = 0
    frmMain.panel.Picture = LoadInterface("centroinventario.bmp")
    'Image2(0).Visible = False
    'Image2(1).Visible = True

    frmMain.picInv.Visible = True
     
    frmMain.hlst.Visible = False

    frmMain.cmdlanzar.Visible = False
    frmMain.imgSpellInfo.Visible = False

    frmMain.cmdMoverHechi(0).Visible = False
    frmMain.cmdMoverHechi(1).Visible = False
    
    Call frmMain.Inventario.ReDraw
    
    frmMain.Left = 0
    frmMain.Top = 0
    frmMain.Width = 1024 * Screen.TwipsPerPixelX
    frmMain.Height = 768 * Screen.TwipsPerPixelY

    frmMain.Visible = True
    frmMain.cerrarcuenta.Enabled = True

    
    Exit Sub

SetConnected_Err:
    Call RegistrarError(Err.Number, Err.Description, "Mod_General.SetConnected", Erl)
    Resume Next
    
End Sub

Sub MoveTo(ByVal Direccion As E_Heading)
    
    On Error GoTo MoveTo_Err
    

    '***************************************************
    'Author: Alejandro Santos (AlejoLp)
    'Last Modify Date: 06/28/2008
    'Last Modified By: Lucas Tavolaro Ortiz (Tavo)
    ' 06/03/2006: AlejoLp - Elimine las funciones Move[NSWE] y las converti a esta
    ' 12/08/2007: Tavo    - Si el usuario esta paralizado no se puede mover.
    ' 06/28/2008: NicoNZ - Saqué lo que impedía que si el usuario estaba paralizado se ejecute el sub.
    '***************************************************
    Dim LegalOk As Boolean
    
    If cartel Then cartel = False
    
    Select Case Direccion

        Case E_Heading.NORTH
            LegalOk = LegalPos(UserPos.x, UserPos.y - 1, Direccion)

        Case E_Heading.EAST
            LegalOk = LegalPos(UserPos.x + 1, UserPos.y, Direccion)

        Case E_Heading.south
            LegalOk = LegalPos(UserPos.x, UserPos.y + 1, Direccion)

        Case E_Heading.WEST
            LegalOk = LegalPos(UserPos.x - 1, UserPos.y, Direccion)

    End Select
    
    If LegalOk And Not UserParalizado And Not UserInmovilizado And Not UserStopped Then
        If Not UserDescansar Then
            If UserMacro.Activado Then
                Call ResetearUserMacro
            End If

            Moviendose = True
            Call MainTimer.Restart(TimersIndex.Walk)
            Call WriteWalk(Direccion) 'We only walk if we are not meditating or resting
            Call Char_Move_by_Head(UserCharIndex, Direccion)
            Call MoveScreen(Direccion)
            
        Else

            If Not UserAvisado Then
                If UserDescansar Then
                    WriteRest 'Stop resting (we do NOT have the 1 step enforcing anymore) sono como un tema de los guns.
                End If

                UserAvisado = True

            End If

        End If

    Else

        If charlist(UserCharIndex).Heading <> Direccion Then
            If IntervaloPermiteHeading(True) Then
                Call WriteChangeHeading(Direccion)
            End If
        End If

    End If
    
    Call frmMain.SetMinimapPosition(0, UserPos.x, UserPos.y)
    
    frmMain.Coord.Caption = UserMap & "-" & UserPos.x & "-" & UserPos.y

    If frmMapaGrande.Visible Then
        Call CalcularPosicionMAPA
    End If
    
    ' Update 3D sounds!
    ' Call Audio.MoveListener(UserPos.x, UserPos.y)
    If frmMain.macrotrabajo.Enabled Then frmMain.DesactivarMacroTrabajo
    
    
    Exit Sub

MoveTo_Err:
    Call RegistrarError(Err.Number, Err.Description, "Mod_General.MoveTo", Erl)
    Resume Next
    
End Sub

Sub RandomMove()
    '***************************************************
    'Author: Alejandro Santos (AlejoLp)
    'Last Modify Date: 06/03/2006
    ' 06/03/2006: AlejoLp - Ahora utiliza la funcion MoveTo
    '***************************************************
    
    On Error GoTo RandomMove_Err
    
    Call MoveTo(RandomNumber(E_Heading.NORTH, E_Heading.WEST))

    
    Exit Sub

RandomMove_Err:
    Call RegistrarError(Err.Number, Err.Description, "Mod_General.RandomMove", Erl)
    Resume Next
    
End Sub

Private Sub AddMovementToKeysMovementPressedQueue()
    
    On Error GoTo AddMovementToKeysMovementPressedQueue_Err
    

    If GetKeyState(BindKeys(14).KeyCode) < 0 Then
        If keysMovementPressedQueue.itemExist(BindKeys(14).KeyCode) = False Then keysMovementPressedQueue.Add (BindKeys(14).KeyCode) ' Agrega la tecla al arraylist
    Else

        If keysMovementPressedQueue.itemExist(BindKeys(14).KeyCode) Then keysMovementPressedQueue.Remove (BindKeys(14).KeyCode) ' Remueve la tecla que teniamos presionada

    End If

    If GetKeyState(BindKeys(15).KeyCode) < 0 Then
        If keysMovementPressedQueue.itemExist(BindKeys(15).KeyCode) = False Then keysMovementPressedQueue.Add (BindKeys(15).KeyCode) ' Agrega la tecla al arraylist
    Else

        If keysMovementPressedQueue.itemExist(BindKeys(15).KeyCode) Then keysMovementPressedQueue.Remove (BindKeys(15).KeyCode) ' Remueve la tecla que teniamos presionada

    End If

    If GetKeyState(BindKeys(16).KeyCode) < 0 Then
        If keysMovementPressedQueue.itemExist(BindKeys(16).KeyCode) = False Then keysMovementPressedQueue.Add (BindKeys(16).KeyCode) ' Agrega la tecla al arraylist
    Else

        If keysMovementPressedQueue.itemExist(BindKeys(16).KeyCode) Then keysMovementPressedQueue.Remove (BindKeys(16).KeyCode) ' Remueve la tecla que teniamos presionada

    End If

    If GetKeyState(BindKeys(17).KeyCode) < 0 Then
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

    Dim Direccion As E_Heading

    Direccion = charlist(UserCharIndex).Heading

    If Not Application.IsAppActive() Then Exit Sub

    If Not pausa And _
        frmMain.Visible And _
        Not frmComerciarUsu.Visible And _
        Not frmBancoObj.Visible And _
        Not frmOpciones.Visible And _
        Not frmComerciar.Visible And _
        Not frmCantidad.Visible And _
        Not frmGoliath.Visible And _
        Not FrmCorreo.Visible And _
        Not frmEstadisticas.Visible And _
        Not frmAlqui.Visible And _
        Not frmCarp.Visible And _
        Not frmHerrero.Visible And _
        Not FrmGrupo.Visible And _
        Not FrmShop.Visible And _
        Not FrmSastre.Visible And _
        Not FrmCorreo.Visible And _
        Not FrmGmAyuda.Visible Then
 
        If frmMain.SendTxt.Visible And PermitirMoverse = 0 Then Exit Sub
 
        If Not UserMoving Then
            If Not UserEstupido Then
                If Not MainTimer.Check(TimersIndex.Walk, False) Then Exit Sub

                Call AddMovementToKeysMovementPressedQueue
                
                Select Case keysMovementPressedQueue.GetLastItem()
                    
                    'Move Up
                    Case BindKeys(14).KeyCode
                        Call MoveTo(E_Heading.NORTH)
                    
                    'Move Right
                    Case BindKeys(17).KeyCode
                        Call MoveTo(E_Heading.EAST)
                        
                    'Move down
                    Case BindKeys(15).KeyCode
                        Call MoveTo(E_Heading.south)
                        
                    'Move left
                    Case BindKeys(16).KeyCode
                        Call MoveTo(E_Heading.WEST)
                        
                End Select

            Else

                Dim kp As Boolean
                    kp = (GetKeyState(BindKeys(14).KeyCode) < 0) Or GetKeyState(BindKeys(17).KeyCode) < 0 Or GetKeyState(BindKeys(15).KeyCode) < 0 Or GetKeyState(BindKeys(16).KeyCode) < 0
            
                If kp Then Call RandomMove

            End If

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
    
    FileExist = (Dir$(File, FileType) <> "")

    
    Exit Function

FileExist_Err:
    Call RegistrarError(Err.Number, Err.Description, "Mod_General.FileExist", Erl)
    Resume Next
    
End Function

Sub Main()
    On Error GoTo Main_Err

    Call InitCommonControls

    #If DEBUGGING = 0 Then
    
        'If Not RunningInVB Then
        
           ' If FindPreviousInstance Then
               ' Call MsgBox("¡Argentum Online ya esta corriendo! No es posible correr otra instancia del juego. Haga clic en Aceptar para salir.", vbApplicationModal + vbInformation + vbOKOnly, "Error al ejecutar")
               ' End
            'End If
    
        'End If
        
    #End If
    CanMoveX = True
    'If Not Launcher Then
    '  Call MsgBox("¡El Juego debe ser abierto desde el Launcher! El Cliente ahora se cerrara.", vbApplicationModal + vbInformation + vbOKOnly, "Error al ejecutar")
    ' End
    ' End If
    
    Call CargarOpciones
    
    If FileExist(App.Path & "\..\..\Launcher\LauncherAO20.ex_", vbNormal) Then
        Kill App.Path & "\..\..\Launcher\LauncherAO20.exe"
        Name App.Path & "\..\..\Launcher\LauncherAO20.ex_" As App.Path & "\..\..\Launcher\LauncherAO20.exe"
    End If

    'Cursores******
    Set FormParser = New clsCursor
    Call FormParser.Init
    'Cursores******
    
    ' Tolerancia 0
    If ComprobarTolerancia0 Then End

    ' Security
    MacAdress = GetMacAddress
    HDserial = GetDriveSerialNumber
    CheckMD5 = GetMd5
    
    ' Leer contraseña de recursos
    Call CheckResources

    Call Load(frmConnect)
    Call Load(FrmLogear)
        
    'If MsgBox("¿Desea jugar en pantalla completa?", vbYesNo, "¡Atención!") = vbYes Then
    
    If PantallaCompleta Then
        Call Resolution.SetResolution
        PantallaCompleta = 1
    End If
    
    Call Frmcarga.Show
 
    
    If Sonido Then
    
        If Sound.Initialize_Engine(frmConnect.hWnd, App.Path & "\..\Recursos", App.Path & "\MP3\", App.Path & "\..\Recursos", False, True, True, VolFX, VolMusic, InvertirSonido) Then
            Call Sound.Ambient_Volume_Set(VolAmbient)
        
        Else

            Call MsgBox("¡No se ha logrado iniciar el engine de DirectSound! Reinstale los últimos controladores de DirectX desde ao20.com.ar", vbCritical, "Saliendo")
            
            Call CloseClient

        End If

    End If

    RawServersList = "190.245.160.106:7667:Horacio;45.235.98.165:7667:InetG;186.139.27.206:7667:ReyarB;191.97.254.154:7667:Martin;127.0.0.1:7667:Localhost"

    Call ComprobarEstado
    Call CargarLst
    
    Call InicializarNombres
    
    'Inicializamos el motor grafico.
    Call Engine_Init
    
    'Inicializamos el inventario
    Call InitializeInventory
    
    'Iniciamos el motor de tiles
    Call Init_TileEngine
    
    'Cargamos todos los init
    Call CargarRecursos
    
    'Cargar fuentes
    Call LoadFonts
    
    FrameTime = timeGetTime And &H7FFFFFFF
    
    UserMap = 1
    AlphaNiebla = 75
    EntradaY = 10
    EntradaX = 10
    
    Call SwitchMap(UserMap)

    'Inicializamos el socket
    Call frmMain.Socket1.Startup
    
    'Set the dialog's font
    Dialogos.font = frmMain.font
    
    ' Load the form for screenshots
    Call Load(frmScreenshots)

    prgRun = True
    pausa = False

    Call Unload(Frmcarga)
    
    Call General_Set_Connect
    
    Call Start
 
    
    Exit Sub

Main_Err:
    Call RegistrarError(Err.Number, Err.Description, "Mod_General.Main", Erl)
    Resume Next
    
End Sub

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

Public Sub ShowSendTxt()
    
    On Error GoTo ShowSendTxt_Err
    

    If Not frmCantidad.Visible Then

        '   Call CompletarEnvioMensajes
        'SendTxt.Visible = True
        'SendTxt.SetFocus
    End If

    
    Exit Sub

ShowSendTxt_Err:
    Call RegistrarError(Err.Number, Err.Description, "Mod_General.ShowSendTxt", Erl)
    Resume Next
    
End Sub

Public Sub LeerLineaComandos()
    
    On Error GoTo LeerLineaComandos_Err
    

    Dim t() As String

    Dim i   As Long
    
    'Parseo los comandos
    t = Split(Command, " ")

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
        
    ListaCiudades(eCiudad.cUllathorpe) = "Ullathorpe"
    ListaCiudades(eCiudad.cNix) = "Nix"
    ListaCiudades(eCiudad.cBanderbill) = "Banderbill"
    ListaCiudades(eCiudad.cLindos) = "Lindos"
    ListaCiudades(eCiudad.cArghal) = "Arghal"
   ' ListaCiudades(eCiudad.cHillidan) = "Hillidan"

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
    
    
    Call Dialogos.RemoveAllDialogs

    
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
    
    Call GuardarOpciones
    
    Call PrevInstance.ReleaseInstance
    'StopURLDetect

    Call Client_UnInitialize_DirectX_Objects
    Sound.Music_Stop
    
    Sound.Engine_DeInitialize
    EngineRun = False
    
    Call General_Set_Mouse_Speed(SensibilidadMouseOriginal)
    
    Rem frmCargando.Show
    
    ' Call Resolution.ResetResolution
    'Stop tile engine
    'Engine_Deinit
    'Stop tile engine
    'Call DeinitTileEngine
    'Engine_Deinit
    
    'Destruimos los objetos públicos creados
    Set CustomKeys = Nothing
    Set SurfaceDB = Nothing
    Set Dialogos = Nothing
    ' Set Audio = Nothing
    Set MainTimer = Nothing
    Set incomingData = Nothing
    Set outgoingData = Nothing
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
    
    ' Call UnloadAllForms
    End

    
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

Public Sub InitServersList(ByVal Lst As String)
    
    On Error GoTo InitServersList_Err
    

    

    Dim NumServers As Integer

    Dim i          As Integer, Cont As Integer

    Cont = General_Field_Count(RawServersList, Asc(";"))

    ReDim ServersLst(1 To Cont) As tServerInfo
    CantServer = Cont

    For i = 1 To Cont

        Dim cur$

        cur$ = General_Field_Read(i, RawServersList, ";")
        ServersLst(i).IP = General_Field_Read(1, cur$, ":")
        ServersLst(i).puerto = Val(General_Field_Read(2, cur$, ":"))
        ServersLst(i).desc = General_Field_Read(3, cur$, ":")
    Next i

    CurServer = 1

    
    Exit Sub

InitServersList_Err:
    Call RegistrarError(Err.Number, Err.Description, "Mod_General.InitServersList", Erl)
    Resume Next
    
End Sub

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


Public Function max(ByVal A As Variant, ByVal B As Variant) As Variant
    
    On Error GoTo max_Err
    

    If A > B Then
        max = A
    Else
        max = B

    End If

    
    Exit Function

max_Err:
    Call RegistrarError(Err.Number, Err.Description, "Mod_General.max", Erl)
    Resume Next
    
End Function

Public Function min(ByVal A As Double, ByVal B As Double) As Variant
    
    On Error GoTo min_Err
    

    If A < B Then
        min = A
    Else
        min = B

    End If

    
    Exit Function

min_Err:
    Call RegistrarError(Err.Number, Err.Description, "Mod_General.min", Erl)
    Resume Next
    
End Function

Public Function Clamp(ByVal A As Variant, ByVal min As Variant, ByVal max As Variant) As Variant
    
    On Error GoTo min_Err
    

    If A < min Then
        Clamp = min
    
    ElseIf A > max Then
        Clamp = max

    Else
        Clamp = A
    End If

    
    Exit Function

min_Err:
    Call RegistrarError(Err.Number, Err.Description, "Mod_General.min", Erl)
    Resume Next
    
End Function


Public Function LoadInterface(FileName As String) As IPicture

On Error GoTo errhandler

    #If Compresion = 1 Then
        Set LoadInterface = General_Load_Picture_From_Resource_Ex(LCase$(FileName), ResourcesPassword)
    #Else
        Set LoadInterface = LoadPicture(App.Path & "/../Recursos/interface/" & LCase$(FileName))
    #End If
    
Exit Function

errhandler:
    MsgBox "Error al cargar la interface: " & FileName

End Function

Public Function LoadMinimap(ByVal map As Integer) As IPicture

On Error GoTo errhandler

    #If Compresion = 1 Then
        Set LoadMinimap = General_Load_Minimap_From_Resource_Ex("mapa" & map & ".bmp", ResourcesPassword)
    #Else
        Set LoadMinimap = LoadPicture(App.Path & "/../Recursos/Minimapas/Mapa" & map & ".bmp")
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

    GetMd5 = MD5File(App.Path & "\Argentum.exe")
    
    Exit Function
    
Handler:
    Call MsgBox("Error al comprobar el cliente del juego, por favor reinstale y vuelva a intentar.", vbOKOnly, "Cliente corrompido")
    End

End Function

Public Function WriteStringToRegistry(Hkey As Long, strPath As String, strValue As String, strdata As String) As Boolean
    
    Dim bAns As Boolean
    
    On Error GoTo ErrorHandler
    
       Dim keyhand As Long
       Dim r As Long
       r = RegCreateKey(Hkey, strPath, keyhand)
       If r = 0 Then
            r = RegSetValueEx(keyhand, strValue, 0, _
               1, ByVal strdata, Len(strdata))
            r = RegCloseKey(keyhand)
        End If
        
       WriteStringToRegistry = (r = 0)
    
    Exit Function
ErrorHandler:
    WriteStringToRegistry = False
    Exit Function
    
End Function

Public Function ReadRegistryKey(Hkey As Long, strPath As String, strValue As String) As String

    Dim bAns As Boolean
    
    On Error GoTo ErrorHandler
    
    Dim keyhand As Long
    Dim r As Long
    Dim Data As String
    Dim LenValue As Long
       
    r = RegOpenKey(Hkey, strPath, keyhand)
    If r = 0 Then
        r = RegQueryValueEx(keyhand, strValue, 0, 1, vbNullString, LenValue)
        
        Data = Space(LenValue)
        
        r = RegQueryValueEx(keyhand, strValue, 0, 1, ByVal Data, Len(Data))
        r = RegCloseKey(keyhand)
        
        ReadRegistryKey = Left$(Data, Len(Data) - 1)
    End If
    
    Exit Function
ErrorHandler:
    ReadRegistryKey = vbNullString
    Exit Function

End Function

Public Function ComprobarTolerancia0() As Boolean

    If ReadRegistryKey(&H80000002, "Software\Temp", "e14a3ff5b5e67ede599cac94358e1028") = "rekcahnuyos" Then
        ComprobarTolerancia0 = True
    End If

End Function

Public Sub CheckResources()

    Dim Data(1 To 200) As Byte
    
    Dim Handle As Integer
    Handle = FreeFile

    Open App.Path & "/../Recursos/OUTPUT/AO.bin" For Binary Access Read As #Handle
    
    Get #Handle, , Data
    
    Close #Handle
    
    Dim length As Integer
    length = Data(UBound(Data)) + Data(UBound(Data) - 1) * 256

    Dim i As Integer
    
    For i = 1 To length
        ResourcesPassword = ResourcesPassword & Chr(Data(i * 3 - 1) Xor 37)
    Next

End Sub

Function ValidarNombre(nombre As String, Error As String) As Boolean
    
    If Len(nombre) < 1 Then
        Error = "Ingrese algún nombre."
        Exit Function
    End If
    
    If Len(nombre) > 18 Then
        Error = "Nombre demasiado largo."
        Exit Function
    End If
    
    Dim temp As String
    temp = UCase$(nombre)
    
    Dim i As Long, Char As Integer, LastChar As Integer
    For i = 1 To Len(temp)
        Char = Asc(mid$(temp, i, 1))
        
        If (Char < 65 Or Char > 90) And Char <> 32 Then
            Error = "Sólo se permites letras y espacios."
            Exit Function
        
        ElseIf Char = 32 And LastChar = 32 Then
            Error = "No se permiten espacios consecutivos."
            Exit Function
        End If
        
        LastChar = Char
    Next

    If Asc(mid$(temp, 1, 1)) = 32 Or Asc(mid$(temp, Len(temp), 1)) = 32 Then
        Error = "No se permiten espacios al inicio o al final."
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
