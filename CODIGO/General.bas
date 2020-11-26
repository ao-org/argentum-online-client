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
    Started As Byte
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

'Very percise counter 64bit system counter
Private Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Currency) As Long

Private Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Currency) As Long
'debemos mostrar la animacion de la lluvia

Private lFrameTimer              As Long

Public Function DirGraficos() As String
    DirGraficos = App.Path & "\..\Recursos\Graficos\"

End Function

Public Function DirSound() As String
    DirSound = App.Path & "\..\Recursos\wav\"

End Function

Public Function DirMidi() As String
    DirMidi = App.Path & "\..\Recursos\midi\"

End Function

Public Function DirMapas() As String
    DirMapas = App.Path & "\..\Recursos\mapas\"

End Function

Public Function RandomNumber(ByVal LowerBound As Long, ByVal UpperBound As Long) As Long
    'Initialize randomizer
    Randomize Timer
    
    'Generate random number
    RandomNumber = (UpperBound - LowerBound) * Rnd + LowerBound

End Function

#If SeguridadAlkon Then
Sub InitMI()
    Dim alternativos As Integer
    Dim CualMITemp As Integer
    
    alternativos = RandomNumber(1, 7368)
    CualMITemp = RandomNumber(1, 1233)
    

    Set MI(CualMITemp) = New clsManagerInvisibles
    Call MI(CualMITemp).Inicializar(alternativos, 10000)
    
    If CualMI <> 0 Then
        Call MI(CualMITemp).CopyFrom(MI(CualMI))
        Set MI(CualMI) = Nothing
    End If
    CualMI = CualMITemp
End Sub
#End If

Sub AddtoRichTextBox2(ByRef RichTextBox As RichTextBox, ByVal Text As String, Optional ByVal red As Integer = -1, Optional ByVal green As Integer, Optional ByVal blue As Integer, Optional ByVal bold As Boolean = False, Optional ByVal italic As Boolean = False, Optional ByVal bCrLf As Boolean = True, Optional ByVal Alignment As Byte = rtfLeft)
    
    '****************************************************
    'Adds text to a Richtext box at the bottom.
    'Automatically scrolls to new text.
    'Text box MUST be multiline and have a 3D apperance!
    '****************************************************
    'Pablo (ToxicWaste) 01/26/2007 : Now the list refeshes properly.
    'Juan Martin Sotuyo Dodero (Maraxus) 03/29/2007 : Replaced ToxicWaste's code for extra performance.
    'Jopi 17/08/2019 : Consola transparente.
    'Jopi 17/08/2019 : Ahora podes especificar el alineamiento del texto.
    '****************************************************

    Call EnableURLDetect(frmmain.RecTxt.hwnd, frmmain.hwnd)

    With RichTextBox
        
        If Len(.Text) > 20000 Then
        
            'Get rid of first line
            .Text = vbNullString
            .SelStart = InStr(1, .Text, vbCrLf) + 1
            .SelLength = Len(.Text) - .SelStart + 2
            .TextRTF = .SelRTF

        End If
        
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
        If Not (RichTextBox = frmmain.RecTxt) Then
            RichTextBox.Refresh

        End If

    End With
    
End Sub

Sub AddtoRichTextBox(ByRef RichTextBox As RichTextBox, ByVal Text As String, Optional ByVal red As Integer = -1, Optional ByVal green As Integer, Optional ByVal blue As Integer, Optional ByVal bold As Boolean = False, Optional ByVal italic As Boolean = False, Optional ByVal bCrLf As Boolean = False, Optional ByVal FontTypeIndex As Byte = 0)

    '******************************************
    'Adds text to a Richtext box at the bottom.
    'Automatically scrolls to new text.
    'Text box MUST be multiline and have a 3D
    'apperance!
    'Pablo (ToxicWaste) 01/26/2007 : Now the list refeshes properly.
    'Juan Martín Sotuyo Dodero (Maraxus) 03/29/2007 : Replaced ToxicWaste's code for extra performance.
    '******************************************r
    Dim bUrl As Boolean

    With RichTextBox

        If Len(.Text) > 20000 Then
            .Text = vbNullString
            .SelStart = InStr(1, .Text, vbCrLf) + 1
            .SelLength = Len(.Text) - .SelStart + 2
            .TextRTF = .SelRTF

        End If
        
        .SelStart = Len(.Text)
        .SelLength = 0
        .SelBold = bold
        .SelItalic = italic
        
        If Not red = -1 Then .SelColor = RGB(red, green, blue)
        bCrLf = True
        
        If bCrLf And Len(.Text) > 0 Then Text = vbCrLf & Text
        .SelText = Text

    End With
    
    ' If bUrl Then DisableUrlDetect

    Dim i As Byte
 
    For i = 2 To MaxLineas
        Con(i - 1).t = Con(i).t
        'Con(i - 1).Color = Con(i).Color
        Con(i - 1).b = Con(i).b
        Con(i - 1).g = Con(i).g
        Con(i - 1).r = Con(i).r
    Next i
 
    Con(MaxLineas).t = Text
    Con(MaxLineas).b = blue
    Con(MaxLineas).g = green
    Con(MaxLineas).r = red
    OffSetConsola = 16
 
    UltimaLineavisible = False
    
End Sub

'TODO : Never was sure this is really necessary....
'TODO : 08/03/2006 - (AlejoLp) Esto hay que volarlo...
Public Sub RefreshAllChars()

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

End Sub

Function AsciiValidos(ByVal cad As String) As Boolean

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

End Function

Function CheckUserDataLoged() As Boolean
    'Validamos los datos del user
    
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

End Function

Function CheckUserData(ByVal checkemail As Boolean) As Boolean

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

End Function

Sub UnloadAllForms()

    On Error Resume Next

    Dim mifrm As Form
    
    For Each mifrm In Forms

        Unload mifrm
    Next
    
End Sub

Function LegalCharacter(ByVal KeyAscii As Integer) As Boolean

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

End Function

Sub SetConnected()
    '*****************************************************************
    'Sets the client to "Connect" mode
    '*****************************************************************
    'Set Connected
    Connected = True
    
    'Unload the connect form
    'FrmCuenta.Visible = False

    frmmain.Label8.Caption = UserName
    LogeoAlgunaVez = True
    
    ' bTecho = False
    AlphaNiebla = 0

    'Vaciamos la cola de movimiento
    keysMovementPressedQueue.Clear

    If FPSFLAG Then
        frmmain.Timerping.Enabled = True
    Else
        frmmain.Timerping.Enabled = False

    End If

    COLOR_AZUL = RGB(0, 0, 0)
    ' establece el borde al listbox
    Call Establecer_Borde(frmmain.hlst, frmmain, COLOR_AZUL, 0, 0)

    Call Make_Transparent_Richtext(frmmain.RecTxt.hwnd)
   
    ' Detect links in console
    Call EnableURLDetect(frmmain.RecTxt.hwnd, frmmain.hwnd)
        
    ' Removemos la barra de titulo pero conservando el caption para la barra de tareas
    Call Form_RemoveTitleBar(frmmain)
   
    frmmain.Image2(1).Tag = "0"
    OpcionMenu = 0
    frmmain.Image2(1).Picture = Nothing
    frmmain.panel.Picture = LoadInterface("centroinventario.bmp")
    '            Image2(0).Visible = False
    ' Image2(1).Visible = True

    frmmain.picInv.Visible = True
     
    frmmain.hlst.Visible = False

    frmmain.cmdlanzar.Visible = False
    frmmain.imgSpellInfo.Visible = False

    frmmain.cmdMoverHechi(0).Visible = False
    frmmain.cmdMoverHechi(1).Visible = False
    Call frmmain.Inventario.ReDraw
    
    frmmain.Left = 0
    frmmain.Top = 0
    frmmain.Width = 1024 * Screen.TwipsPerPixelX
    frmmain.Height = 768 * Screen.TwipsPerPixelY

    frmmain.Visible = True
    frmmain.cerrarcuenta.Enabled = True

End Sub

Sub MoveTo(ByVal Direccion As E_Heading)

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
    
    If LegalOk And Not UserParalizado And Not UserInmovilizado Then
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
    
    frmmain.personaje(0).Left = UserPos.x - 5
    frmmain.personaje(0).Top = UserPos.y - 4
    
    frmmain.Coord.Caption = UserMap & "-" & UserPos.x & "-" & UserPos.y

    If frmMapaGrande.Visible Then

        Dim x As Long

        Dim y As Long
            
        x = (idmap - 1) Mod 16
        y = Int((idmap - 1) / 16)

        frmMapaGrande.lblAllies.Top = y * 27
        frmMapaGrande.lblAllies.Left = x * 27

        frmMapaGrande.Shape1.Top = y * 27 + (UserPos.y / 4.5)
        frmMapaGrande.Shape1.Left = x * 27 + (UserPos.x / 4.5)

    End If
    
    ' Update 3D sounds!
    ' Call Audio.MoveListener(UserPos.x, UserPos.y)
    If frmmain.macrotrabajo.Enabled Then frmmain.DesactivarMacroTrabajo
    
End Sub

Sub RandomMove()
    '***************************************************
    'Author: Alejandro Santos (AlejoLp)
    'Last Modify Date: 06/03/2006
    ' 06/03/2006: AlejoLp - Ahora utiliza la funcion MoveTo
    '***************************************************
    Call MoveTo(RandomNumber(E_Heading.NORTH, E_Heading.WEST))

End Sub

Private Sub AddMovementToKeysMovementPressedQueue()

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

End Sub

Sub Check_Keys()

    On Error Resume Next

    Static lastMovement As Long

    Dim Direccion As E_Heading

    Direccion = charlist(UserCharIndex).Heading

    If Not Application.IsAppActive() Then Exit Sub

    If Not pausa And _
        frmmain.Visible And _
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
 
        If frmmain.SendTxt.Visible And PermitirMoverse = 0 Then Exit Sub
 
        If UserMoving = 0 Then
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

End Sub

'TODO : Si bien nunca estuvo allí, el mapa es algo independiente o a lo sumo dependiente del engine, no va acá!!!
Sub SwitchMapIAO(ByVal map As Integer)
    
    'Cargamos el mapa.
    Call Recursos.CargarMapa(map)
    
    Dim Rojo As Byte, Verde As Byte, Azul As Byte

    If MapDat.base_light = 16777215 Then
    
        Map_light_base = D3DColorARGB(255, 255, 255, 255)
        ColorAmbiente.r = 255
        ColorAmbiente.b = 255
        ColorAmbiente.g = 255
        ColorAmbiente.a = 255
        
        Call Map_Base_Light_Set(D3DColorARGB(255, 255, 255, 255))
        
    Else
        
        Call Obtener_RGB(MapDat.base_light, Rojo, Verde, Azul)
        
        ColorAmbiente.r = Rojo
        ColorAmbiente.b = Azul
        ColorAmbiente.g = Verde
        ColorAmbiente.a = 255
        Map_light_base = D3DColorARGB(255, ColorAmbiente.r, ColorAmbiente.g, ColorAmbiente.b)
        
        Call Map_Base_Light_Set(Map_light_base)

    End If

    Map_light_baseBackup = Map_light_base

    Call LucesRedondas.LightRenderAll
    Call LucesCuadradas.Light_Render_All
    
    Call DibujarMiniMapa
    
    CurMap = map
    
    If Musica Then
        
        If MapDat.music_numberLow > 0 Then
        
            If Sound.MusicActual <> MapDat.music_numberLow Then
                Sound.NextMusic = MapDat.music_numberLow
                Sound.Fading = 200
            End If

        Else

            If MapDat.music_numberHi > 0 Then
                
                If Sound.MusicActual <> MapDat.music_numberHi Then
                    Sound.NextMusic = MapDat.music_numberHi
                    Sound.Fading = 100
                End If

                Call ReproducirMp3(MapDat.music_numberHi)
                
                Call Sound.Music_Load(MapDat.music_numberHi, 0, 0)
                
                Call Sound.Music_Play

            End If

        End If

    End If

    If bRain And MapDat.LLUVIA Then
        Call Graficos_Particulas.Engine_Meteo_Particle_Set(Particula_Lluvia)
    
    ElseIf bNieve And MapDat.NIEVE Then
        Call Graficos_Particulas.Engine_Meteo_Particle_Set(Particula_Nieve)

    End If
    
    If AmbientalActivated = 1 Then
        Call AmbientarAudio(map)
    End If

    Call NameMapa(map)

End Sub

Function ReadField(ByVal Pos As Integer, ByRef Text As String, ByVal SepASCII As Byte) As String

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

End Function

Function FieldCount(ByRef Text As String, ByVal SepASCII As Byte) As Long

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

End Function

Function FileExist(ByVal File As String, ByVal FileType As VbFileAttribute) As Boolean
    FileExist = (Dir$(File, FileType) <> "")

End Function

Sub Main()

    On Error Resume Next

    Call InitCommonControls
    
    #If DEBUGGING = 0 Then
    
        If Not RunningInVB Then
        
            If FindPreviousInstance Then
                Call MsgBox("¡Argentum Online ya esta corriendo! No es posible correr otra instancia del juego. Haga clic en Aceptar para salir.", vbApplicationModal + vbInformation + vbOKOnly, "Error al ejecutar")
                End
            End If
    
        End If
        
    #End If

    'If Not Launcher Then
    '  Call MsgBox("¡El Juego debe ser abierto desde el Launcher! El Cliente ahora se cerrara.", vbApplicationModal + vbInformation + vbOKOnly, "Error al ejecutar")
    ' End
    ' End If
    
    Call CargarOpciones
    
    If FileExist(App.Path & "\..\LauncherAO20.ex_", vbNormal) Then
        Call Sleep(2)
        Delete_File App.Path & "\..\LauncherAO20.exe"
        Name App.Path & "\..\LauncherAO20.ex_" As App.Path & "\..\LauncherAO20.exe"

    End If

    'Cursores******
    Set FormParser = New clsCursor
    Call FormParser.Init
    'Cursores******

    MacAdress = GetMacAddress
    HDserial = GetDriveSerialNumber
    
    Call Load(frmConnect)
    Call Load(FrmLogear)
        
    'If MsgBox("¿Desea jugar en pantalla completa?", vbYesNo, "¡Atención!") = vbYes Then
    
    If PantallaCompleta Then
        Call Resolution.SetResolution
        PantallaCompleta = 1
    End If
    
    Call Frmcarga.Show
 
    
    If Sonido Then
    
        If Sound.Initialize_Engine(frmConnect.hwnd, App.Path & "\..\Recursos", App.Path & "\MP3\", App.Path & "\..\Recursos", False, True, True, VolFX, VolMusic, InvertirSonido) Then
            Call Sound.Ambient_Volume_Set(VolAmbient)
        
        Else

            Call MsgBox("¡No se ha logrado iniciar el engine de DirectSound! Reinstale los últimos controladores de DirectX desde www.argentum20.com", vbCritical, "Saliendo")
            
            Call CloseClient

        End If

    End If

    RawServersList = "190.245.145.3:7667:Horacio;190.210.83.155:7667:Iplan;186.139.103.88:7667:ReyarB;45.235.99.105:7500:Pablo;127.0.0.1:7667:Localhost"

    Call ComprobarEstado
    Call CargarLst
    
    Call InicializarNombres
    
    'Inicializamos el motor grafico.
    Call Engine_Init
    
    'Iniciamos el motor de tiles
    Call Init_TileEngine
    
    'Cargamos todos los init
    Call CargarRecursos
    
    UserMap = 1
    AlphaNiebla = 75
    EntradaY = 10
    EntradaX = 10
    
    Call SwitchMapIAO(UserMap)

    'Inicializamos el socket
    Call frmmain.Socket1.Startup
    
    'Set the dialog's font
    Dialogos.font = frmmain.font
    
    ' Load the form for screenshots
    Call Load(frmScreenshots)

    prgRun = True
    pausa = False

    Call Unload(Frmcarga)
    
    Call General_Set_Connect
    
    Call Start
 
End Sub

Sub WriteVar(ByVal File As String, ByVal Main As String, ByVal Var As String, ByVal Value As String)
    '*****************************************************************
    'Writes a var to a text file
    '*****************************************************************
    writeprivateprofilestring Main, Var, Value, File

End Sub

Function GetVar(ByVal File As String, ByVal Main As String, ByVal Var As String) As String

    '*****************************************************************
    'Gets a Var from a text file
    '*****************************************************************
    Dim sSpaces As String ' This will hold the input that the program will retrieve
    
    sSpaces = Space$(100) ' This tells the computer how long the longest string can be. If you want, you can change the number 100 to any number you wish
    
    getprivateprofilestring Main, Var, vbNullString, sSpaces, Len(sSpaces), File
    
    GetVar = RTrim$(sSpaces)
    GetVar = Left$(GetVar, Len(GetVar) - 1)

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
    CMSValidateChar_ = (iAsc >= 48 And iAsc <= 57) Or (iAsc >= 65 And iAsc <= 90) Or (iAsc >= 97 And iAsc <= 122) Or (iAsc = 95) Or (iAsc = 45) Or (iAsc = 46)

End Function

'TODO : como todo lo relativo a mapas, no tiene nada que hacer acá....
Function HayAgua(ByVal x As Integer, ByVal y As Integer) As Boolean
    HayAgua = ((MapData(x, y).Graphic(1).GrhIndex >= 1505 And MapData(x, y).Graphic(1).GrhIndex <= 1520) Or (MapData(x, y).Graphic(1).GrhIndex >= 24223 And MapData(x, y).Graphic(1).GrhIndex <= 24238) Or (MapData(x, y).Graphic(1).GrhIndex >= 24143 And MapData(x, y).Graphic(1).GrhIndex <= 24158) Or (MapData(x, y).Graphic(1).GrhIndex >= 468 And MapData(x, y).Graphic(1).GrhIndex <= 483) Or (MapData(x, y).Graphic(1).GrhIndex >= 44668 And MapData(x, y).Graphic(1).GrhIndex <= 44939) Or (MapData(x, y).Graphic(1).GrhIndex >= 24303 And MapData(x, y).Graphic(1).GrhIndex <= 24318))
                
    'If MapData(x, y).Trigger = 8 Then
    ' HayAgua = True
    ' End If
                
End Function

Function EsArbol(ByVal GrhIndex As Long) As Boolean
    EsArbol = GrhIndex = 7000 Or GrhIndex = 7001 Or GrhIndex = 7002 Or GrhIndex = 641 Or GrhIndex = 26075 Or GrhIndex = 643 Or GrhIndex = 644 Or _
       GrhIndex = 647 Or GrhIndex = 26076 Or GrhIndex = 7222 Or GrhIndex = 7223 Or GrhIndex = 7224 Or GrhIndex = 7225 Or GrhIndex = 7226 Or _
       GrhIndex = 26077 Or GrhIndex = 26079 Or GrhIndex = 735 Or GrhIndex = 32343 Or GrhIndex = 32344 Or GrhIndex = 26080 Or GrhIndex = 26081 Or _
       GrhIndex = 32345 Or GrhIndex = 32346 Or GrhIndex = 32347 Or GrhIndex = 32348 Or GrhIndex = 32349 Or GrhIndex = 32350 Or GrhIndex = 32351 Or _
       GrhIndex = 32352 Or GrhIndex = 14961 Or GrhIndex = 14950 Or GrhIndex = 14951 Or GrhIndex = 14952 Or GrhIndex = 14953 Or GrhIndex = 14954 Or _
       GrhIndex = 14955 Or GrhIndex = 14956 Or GrhIndex = 14957 Or GrhIndex = 14958 Or GrhIndex = 14959 Or GrhIndex = 14962 Or GrhIndex = 14963 Or _
       GrhIndex = 14964 Or GrhIndex = 14967 Or GrhIndex = 14968 Or GrhIndex = 14969 Or GrhIndex = 14970 Or GrhIndex = 14971 Or GrhIndex = 14972 Or _
       GrhIndex = 14973 Or GrhIndex = 14974 Or GrhIndex = 14975 Or GrhIndex = 14976 Or GrhIndex = 14978 Or GrhIndex = 14980 Or GrhIndex = 14982 Or _
       GrhIndex = 14983 Or GrhIndex = 14984 Or GrhIndex = 14985 Or GrhIndex = 14987 Or GrhIndex = 14988 Or GrhIndex = 26078 Or GrhIndex = 26192

End Function

Public Function EsObjetoFijo(ByVal x As Integer, ByVal y As Integer) As Boolean
    Dim OBJIndex As Integer
    OBJIndex = MapData(x, y).OBJInfo.OBJIndex
    
    Dim ObjType As eObjType
    ObjType = ObjData(OBJIndex).ObjType
    
    EsObjetoFijo = ObjType = eObjType.otForos Or ObjType = eObjType.otCarteles Or ObjType = eObjType.otArboles Or ObjType = eObjType.otYacimiento Or ObjType = eObjType.OtDecoraciones

End Function

Public Sub ShowSendTxt()

    If Not frmCantidad.Visible Then

        '   Call CompletarEnvioMensajes
        'SendTxt.Visible = True
        'SendTxt.SetFocus
    End If

End Sub

Public Sub LeerLineaComandos()

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

End Sub

Private Sub InicializarNombres()
    '**************************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modify Date: 11/27/2005
    'Inicializa los nombres de razas, ciudades, clases, skills, atributos, etc.
    '**************************************************************

    ListaRazas(eRaza.Humano) = "Humano"
    ListaRazas(eRaza.Elfo) = "Elfo"
    ListaRazas(eRaza.ElfoOscuro) = "Elfo Drow"
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

    SkillsNames(eSkill.magia) = "Magia"
    SkillsNames(eSkill.Robar) = "Robar"
    SkillsNames(eSkill.Tacticas) = "Destreza en combate"
    SkillsNames(eSkill.Armas) = "Combate con armas"
    SkillsNames(eSkill.Meditar) = "Meditar"
    SkillsNames(eSkill.Apuñalar) = "Apuñalar"
    SkillsNames(eSkill.Ocultarse) = "Ocultarse"
    SkillsNames(eSkill.Supervivencia) = "Supervivencia"
    SkillsNames(eSkill.Comerciar) = "Comercio"
    SkillsNames(eSkill.Defensa) = "Defensa con escudos"
    SkillsNames(eSkill.Liderazgo) = "Liderazgo"
    SkillsNames(eSkill.Proyectiles) = "Armas a distancia"
    SkillsNames(eSkill.Wrestling) = "Combate sin armas"
    SkillsNames(eSkill.Navegacion) = "Navegacion"
    SkillsNames(eSkill.equitacion) = "Equitación"
    SkillsNames(eSkill.Resistencia) = "Resistencia Mágica"
    SkillsNames(eSkill.Talar) = "Tala"
    SkillsNames(eSkill.Pescar) = "Pesca"
    SkillsNames(eSkill.Mineria) = "Minería"
    SkillsNames(eSkill.Herreria) = "Herrería"
    SkillsNames(eSkill.Carpinteria) = "Carpinteriía"
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
    
    Call Dialogos.RemoveAllDialogs

End Sub

Public Sub CloseClient()
    '**************************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modify Date: 8/14/2007
    'Frees all used resources, cleans up and leaves
    '**************************************************************
    ' Allow new instances of the client to be opened
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
    Set frmmain.Inventario = Nothing
    Set frmComerciar.InvComNpc = Nothing
    Set frmComerciar.InvComUsu = Nothing
    Set frmBancoObj.InvBankUsu = Nothing
    Set frmBancoObj.InvBoveda = Nothing
    Set FrmKeyInv.InvKeys = Nothing
    
    ' Call UnloadAllForms
    End

End Sub

Public Function General_Field_Read(ByVal field_pos As Long, ByVal Text As String, ByVal delimiter As String) As String

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

End Function

Public Function General_Field_Count(ByVal Text As String, ByVal delimiter As Byte) As Long

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

End Function

Public Sub InitServersList(ByVal Lst As String)

    On Error Resume Next

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

End Sub

Public Function General_Get_Elapsed_Time() As Single

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

End Function


Public Function max(ByVal a As Double, ByVal b As Double) As Double

    If a > b Then
        max = a
    Else
        max = b

    End If

End Function

Public Function min(ByVal a As Double, ByVal b As Double) As Double

    If a < b Then
        min = a
    Else
        min = b

    End If

End Function

Public Function LoadInterface(FileName As String) As IPicture

On Error GoTo errhandler

    #If Compresion = 1 Then
        Set LoadInterface = General_Load_Picture_From_Resource_Ex(LCase$(FileName))
    
    #Else
        Set LoadInterface = LoadPicture(App.Path & "/../Recursos/interface/" & LCase$(FileName))
    #End If
    
Exit Function

errhandler:
    MsgBox "Error al cargar la interface: " & FileName

End Function

Public Function Tilde(ByRef Data As String) As String

    Tilde = UCase$(Data)
 
    Tilde = Replace$(Tilde, "Á", "A")
    Tilde = Replace$(Tilde, "É", "E")
    Tilde = Replace$(Tilde, "Í", "I")
    Tilde = Replace$(Tilde, "Ó", "O")
    Tilde = Replace$(Tilde, "Ú", "U")
        
End Function

' Copiado de https://www.vbforums.com/showthread.php?231468-VB-Detect-if-you-are-running-in-the-IDE
Function RunningInVB() As Boolean
    'Returns whether we are running in vb(true), or compiled (false)
 
    Static counter As Variant

    If IsEmpty(counter) Then
        counter = 1
        Debug.Assert RunningInVB() Or True
        counter = counter - 1
    ElseIf counter = 1 Then
        counter = 0

    End If

    RunningInVB = counter
 
End Function
