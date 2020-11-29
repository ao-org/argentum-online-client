Attribute VB_Name = "Recursos"
Option Explicit

' *********************************************************
' FUENTES
' *********************************************************
Private Type tFont
    red As Byte
    green As Byte
    blue As Byte
    bold As Boolean
    italic As Boolean
End Type

Public Enum FontTypeNames

    FONTTYPE_TALK
    FONTTYPE_FIGHT
    FONTTYPE_WARNING
    FONTTYPE_INFO
    FONTTYPE_INFOBOLD
    FONTTYPE_EJECUCION
    FONTTYPE_PARTY
    FONTTYPE_VENENO
    FONTTYPE_GUILD
    FONTTYPE_SERVER
    FONTTYPE_GUILDMSG
    FONTTYPE_CONSEJO
    FONTTYPE_CONSEJOCAOS
    FONTTYPE_CONSEJOVesA
    FONTTYPE_CONSEJOCAOSVesA
    FONTTYPE_CENTINELA
    FONTTYPE_GMMSG
    FONTTYPE_GM
    FONTTYPE_DIOS
    FONTTYPE_CITIZEN
    FONTTYPE_CRIMINAL
    FONTTYPE_EXP
    FONTTYPE_SUBASTA
    FONTTYPE_GLOBAL
    FONTTYPE_MP
    FONTTYPE_ROSA
    FONTTYPE_VIOLETA
    FONTTYPE_INFOIAO
    FONTTYPE_New_Amarillo_Oscuro
    FONTTYPE_New_Verde_Oscuro
    FONTTYPE_New_Naranja
    FONTTYPE_New_Celeste
    FONTTYPE_New_Amarillo_Verdoso
    FONTTYPE_New_Gris
    FONTTYPE_New_Blanco
    FONTTYPE_New_Rojo_Salmon
    FONTTYPE_New_DONADOR
    FONTTYPE_New_GRUPO
    FONTTYPE_New_Eventos '39

End Enum

Public FontTypes(39) As tFont
' *********************************************************
' FIN - FUENTES
' *********************************************************

' *********************************************************
' CARGA DE MAPAS
' Sinuhe - Map format .CSM
' *********************************************************
'The only current map



Private Type tMapHeader

    NumeroBloqueados As Long
    NumeroLayers(1 To 4) As Long
    NumeroTriggers As Long
    NumeroLuces As Long
    NumeroParticulas As Long
    NumeroNPCs As Long
    NumeroOBJs As Long
    NumeroTE As Long

End Type

Private Type tDatosBloqueados

    X As Integer
    Y As Integer
    lados As Byte

End Type

Private Type tDatosGrh

    X As Integer
    Y As Integer
    GrhIndex As Long

End Type

Private Type tDatosTrigger

    X As Integer
    Y As Integer
    Trigger As Integer

End Type

Private Type tDatosLuces

    X As Integer
    Y As Integer
    Color As RGBA
    Rango As Byte

End Type

Private Type tDatosParticulas

    X As Integer
    Y As Integer
    Particula As Long

End Type

Public Type tDatosNPC

    X As Integer
    Y As Integer
    NpcIndex As Integer

End Type

Private Type tDatosObjs

    X As Integer
    Y As Integer
    OBJIndex As Integer
    ObjAmmount As Integer

End Type

Private Type tDatosTE

    X As Integer
    Y As Integer
    DestM As Integer
    DestX As Integer
    DestY As Integer

End Type

Private Type tMapSize

    XMax As Integer
    XMin As Integer
    YMax As Integer
    YMin As Integer

End Type

Private Type tMapDat

    map_name As String
    backup_mode As Byte
    restrict_mode As String
    music_numberHi As Long
    music_numberLow As Long
    Seguro As Byte
    zone As String
    terrain As String
    ambient As String
    base_light As Long
    letter_grh As Long
    extra1 As Long
    extra2 As Long
    extra3 As String
    LLUVIA As Byte
    NIEVE As Byte
    niebla As Byte

End Type

Private MapSize As tMapSize
Public MapDat   As tMapDat
Public iplst    As String
' *********************************************************
'   FIN - CARGA DE MAPAS
' *********************************************************

Public Sub CargarRecursos()
    
    If UtilizarPreCarga = 1 Then
        Call PreloadGraphics
    End If
    
    Call CargarParticulasBinary
    Call CargarIndicesOBJ
    Call Cargarmapsworlddata
    Call InitFonts
    
    Call LoadGrhData
    Call CargarCabezas
    Call CargarCascos
    Call CargarCuerpos
    Call CargarFxs
    Call CargarMiniMap
    Call CargarPasos
    Call CargarAnimArmas
    Call CargarAnimEscudos
    Call CargarColores

End Sub

''
' Initializes the fonts array

Public Sub InitFonts()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    With FontTypes(FontTypeNames.FONTTYPE_TALK)
        .red = 255
        .green = 255
        .blue = 255

    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_FIGHT)
        .red = 255
        .bold = 1

    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_WARNING)
        .red = 32
        .green = 51
        .blue = 223
        .bold = 1
        .italic = 1

    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_INFO)
        .red = 65
        .green = 190
        .blue = 156

    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_INFOBOLD)
        .red = 65
        .green = 190
        .blue = 156
        .bold = 1

    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_EJECUCION)
        .red = 130
        .green = 130
        .blue = 130
        .bold = 1

    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_PARTY)
        .red = 255
        .green = 180
        .blue = 250

    End With
    
    FontTypes(FontTypeNames.FONTTYPE_VENENO).green = 255
    
    With FontTypes(FontTypeNames.FONTTYPE_GUILD)
        .red = 255
        .green = 255
        .blue = 255
        .bold = 1

    End With
    
    FontTypes(FontTypeNames.FONTTYPE_SERVER).green = 185
    
    With FontTypes(FontTypeNames.FONTTYPE_GUILDMSG)
        .red = 228
        .green = 199
        .blue = 27

    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_CONSEJO)
        .red = 31
        .green = 139
        .blue = 139
        .bold = 1

    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_CONSEJOCAOS)
        .red = 179
        .green = 0
        .blue = 4
        .bold = 1

    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_CONSEJOVesA)
        .red = 31
        .green = 139
        .blue = 139
        .bold = 1

    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_CONSEJOCAOSVesA)
        .red = 189
        .green = 0
        .bold = 1

    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_CENTINELA)
        .green = 255
        .bold = 1

    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_GMMSG)
        .red = 255
        .green = 255
        .blue = 255
        .italic = 1

    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_GM)
        .red = 2
        .green = 161
        .blue = 38
        .bold = 1

    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_DIOS)
        .red = 217
        .green = 164
        .blue = 32
        .bold = 1

    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_CITIZEN)
        .red = 0
        .green = 128
        .blue = 255
        .bold = 1

    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_CRIMINAL)
        .red = 255
        .bold = 1
    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_EXP)
        .red = 42
        .green = 169
        .blue = 222
        .bold = 1

    End With
        
    With FontTypes(FontTypeNames.FONTTYPE_SUBASTA)
        .red = 188
        .green = 192
        .blue = 103
        .bold = 0

    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_GLOBAL)
        .red = 0
        .green = 176
        .blue = 176
        .bold = 0
        .italic = True

    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_MP)
        .red = 157
        .green = 226
        .blue = 20
        .bold = 0

    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_ROSA)
        .red = 255
        .green = 0
        .blue = 128
        .bold = 0

    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_VIOLETA)
        .red = 99
        .green = 0
        .blue = 198
        .bold = 0

    End With
        
    With FontTypes(FontTypeNames.FONTTYPE_INFOIAO)
        .red = 204
        .green = 193
        .blue = 115
        .bold = 0
        .italic = 1

    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_New_Amarillo_Oscuro)
        .red = 150
        .green = 100
        .blue = 20
        .bold = 0

    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_New_Verde_Oscuro)
        .red = 0
        .green = 120
        .blue = 70
        .bold = 0

    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_New_Naranja)
        .red = 255
        .green = 80
        .blue = 0
        .bold = 0

    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_New_Celeste)
        .red = 0
        .green = 200
        .blue = 255
        .bold = 0

    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_New_Amarillo_Verdoso)
        .red = 150
        .green = 150
        .blue = 0
        .bold = 0

    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_New_Gris)
        .red = 128
        .green = 128
        .blue = 128
        .bold = 1

    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_New_Blanco)
        .red = 255
        .green = 255
        .blue = 255
        .bold = 0

    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_New_Rojo_Salmon)
        .red = 200
        .green = 50
        .blue = 50
        .bold = 0

    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_New_DONADOR)
        .red = 100
        .green = 180
        .blue = 200
        .bold = 1

    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_New_GRUPO)
        .red = 250
        .green = 200
        .blue = 0
        .bold = 0
        .italic = 1

    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_New_Eventos)
        .red = 0
        .green = 200
        .blue = 250
        .bold = 1
        .italic = 1

    End With

End Sub

Public Sub CargarPasos()

    ReDim Pasos(1 To NUM_PASOS) As tPaso

    Pasos(CONST_BOSQUE).CantPasos = 2  ' OK
    ReDim Pasos(CONST_BOSQUE).wav(1 To Pasos(CONST_BOSQUE).CantPasos) As Integer
    Pasos(CONST_BOSQUE).wav(1) = 201
    Pasos(CONST_BOSQUE).wav(2) = 69

    Pasos(CONST_NIEVE).CantPasos = 2 ' OK
    ReDim Pasos(CONST_NIEVE).wav(1 To Pasos(CONST_NIEVE).CantPasos) As Integer
    Pasos(CONST_NIEVE).wav(1) = 199
    Pasos(CONST_NIEVE).wav(2) = 200

    Pasos(CONST_CABALLO).CantPasos = 2
    ReDim Pasos(CONST_CABALLO).wav(1 To Pasos(CONST_CABALLO).CantPasos) As Integer
    Pasos(CONST_CABALLO).wav(1) = 70
    Pasos(CONST_CABALLO).wav(2) = 71

    Pasos(CONST_DUNGEON).CantPasos = 2 '
    ReDim Pasos(CONST_DUNGEON).wav(1 To Pasos(CONST_DUNGEON).CantPasos) As Integer
    Pasos(CONST_DUNGEON).wav(1) = 23
    Pasos(CONST_DUNGEON).wav(2) = 24

    Pasos(CONST_DESIERTO).CantPasos = 2 ' OK
    ReDim Pasos(CONST_DESIERTO).wav(1 To Pasos(CONST_DESIERTO).CantPasos) As Integer
    Pasos(CONST_DESIERTO).wav(1) = 197
    Pasos(CONST_DESIERTO).wav(2) = 198

    Pasos(CONST_PISO).CantPasos = 2 ' OK
    ReDim Pasos(CONST_PISO).wav(1 To Pasos(CONST_PISO).CantPasos) As Integer
    Pasos(CONST_PISO).wav(1) = 23
    Pasos(CONST_PISO).wav(2) = 24

End Sub

Sub CargarDatosMapa(ByVal map As Integer)

    If Len(NameMaps(map).desc) <> 0 Then
        frmMapaGrande.Label1.Caption = NameMaps(map).desc
    Else
        frmMapaGrande.Label1.Caption = "Sin información relevante."

    End If

    '**************************************************************
    'Formato de mapas optimizado para reducir el espacio que ocupan.
    'Diseñado y creado por Juan Martín Sotuyo Dodero (Maraxus) (juansotuyo@hotmail.com)
    '**************************************************************

    Dim fh           As Integer
    Dim MH           As tMapHeader

    Dim Blqs()       As tDatosBloqueados

    Dim MapRoute     As String

    Dim L1()         As tDatosGrh
    Dim L2()         As tDatosGrh
    Dim L3()         As tDatosGrh
    Dim L4()         As tDatosGrh

    Dim Triggers()   As tDatosTrigger
    Dim Luces()      As tDatosLuces
    Dim Particulas() As tDatosParticulas
    Dim Objetos()    As tDatosObjs
    Dim NPCs()       As tDatosNPC

    Dim i            As Long
    Dim j            As Long
    
    Dim X            As Long
    Dim Y            As Long
    
    #If Compresion = 1 Then

        If Not Extract_File(Maps, App.Path & "\..\Recursos\OUTPUT\", "mapa" & map & ".csm", Windows_Temp_Dir, False) Then
            Err.Description = "¡No se puede cargar el archivo de mapas! El juego se cerrara."
            MsgBox Err.Description
            End

        End If

        MapRoute = Windows_Temp_Dir & "mapa" & map & ".csm"
    #Else
        MapRoute = App.Path & "\..\Recursos\Mapas\mapa" & map & ".csm"
    #End If

    fh = FreeFile
    Open MapRoute For Binary As fh
    Get #fh, , MH
    Get #fh, , MapSize
    Get #fh, , MapDat

    With MapSize
    
        ' Get #fh, , L1
        With MH

            'Cargamos Bloqueos
        
            If .NumeroBloqueados > 0 Then
                ReDim Blqs(1 To .NumeroBloqueados)
                Get #fh, , Blqs

                For i = 1 To .NumeroBloqueados
            
                    'MapData(Blqs(i).X, Blqs(i).Y).Blocked = 1
                Next i

            End If
    
            'Cargamos Layer 1
        
            If .NumeroLayers(1) > 0 Then
        
                ReDim L1(1 To .NumeroLayers(1))
                Get #fh, , L1

                For i = 1 To .NumeroLayers(1)
            
                    ' MapData(L1(i).X, L1(i).Y).Graphic(1).grhindex = L1(i).grhindex
            
                    '  InitGrh MapData(L1(i).X, L1(i).Y).Graphic(1), MapData(L1(i).X, L1(i).Y).Graphic(1).grhindex
                    ' Call Map_Grh_Set(L2(i).x, L2(i).y, L2(i).GrhIndex, 2)
                Next i

            End If
    
            If .NumeroLayers(2) > 0 Then

                ReDim L2(1 To .NumeroLayers(2))
                Get #fh, , L2

                For i = 1 To .NumeroLayers(2)
            
                    '   MapData(L2(i).X, L2(i).Y).Graphic(2).grhindex = L2(i).grhindex
            
                    '  InitGrh MapData(L2(i).X, L2(i).Y).Graphic(2), MapData(L2(i).X, L2(i).Y).Graphic(2).grhindex
                Next i

            End If
    
            If .NumeroLayers(3) > 0 Then
                ReDim L3(1 To .NumeroLayers(3))
                Get #fh, , L3

                For i = 1 To .NumeroLayers(3)
            
                    '  MapData(L3(i).X, L3(i).Y).Graphic(3).grhindex = L3(i).grhindex
            
                    ' InitGrh MapData(L3(i).X, L3(i).Y).Graphic(3), MapData(L3(i).X, L3(i).Y).Graphic(3).grhindex
                Next i

            End If
    
            If .NumeroLayers(4) > 0 Then

                ReDim L4(1 To .NumeroLayers(4))
                Get #fh, , L4

                For i = 1 To .NumeroLayers(4)
            
                    '   MapData(L4(i).X, L4(i).Y).Graphic(4).grhindex = L4(i).grhindex
                    '   MapData(L4(i).X, L4(i).Y).GrhBlend = 255
                    '   InitGrh MapData(L4(i).X, L4(i).Y).Graphic(4), MapData(L4(i).X, L4(i).Y).Graphic(4).grhindex
                Next i

            End If
    
            If .NumeroTriggers > 0 Then
                ReDim Triggers(1 To .NumeroTriggers)
                Get #fh, , Triggers
            
                For i = 1 To .NumeroTriggers
            
                    Rem   If Triggers(i).Trigger > 8 Then Triggers(i).Trigger = 1
                    '      MapData(Triggers(i).X, Triggers(i).Y).Trigger = Triggers(i).Trigger
                Next i

            End If
    
            If .NumeroParticulas > 0 Then
                ReDim Particulas(1 To .NumeroParticulas)
                Get #fh, , Particulas

                For i = 1 To .NumeroParticulas
            
                    '   MapData(Particulas(i).X, Particulas(i).Y).particle_Index = Particulas(i).Particula
                    '   General_Particle_Create MapData(Particulas(i).X, Particulas(i).Y).particle_Index, Particulas(i).X, Particulas(i).Y

                Next i

            End If

            If .NumeroLuces > 0 Then
                ReDim Luces(1 To .NumeroLuces)
                Get #fh, , Luces

                For i = 1 To .NumeroLuces
                    '     MapData(Luces(i).X, Luces(i).Y).luz.color = Luces(i).color
                    '    MapData(Luces(i).X, Luces(i).Y).luz.Rango = Luces(i).Rango
                    '    If MapData(Luces(i).X, Luces(i).Y).luz.Rango <> 0 Then
                    '  LightRound.Create_Light_To_Map Luces(I).x, Luces(I).y, CByte(MapData(Luces(I).x, Luces(I).y).luz.Rango), MapData(Luces(I).x, Luces(I).y).luz.color
                    'LucesCuadradas.Light_Create Luces(i).X, Luces(i).Y, MapData(Luces(i).X, Luces(i).Y).luz.color, MapData(Luces(i).X, Luces(i).Y).luz.Rango, Luces(i).X & Luces(i).Y
                    'LightRound.Render_All_Light
                    'LucesRedondas.Create_Light_To_Map Luces(i).X, Luces(i).Y, CByte(MapData(Luces(i).X, Luces(i).Y).luz.Rango), 255, 255, 255
                    '   LucesCuadradas.Light_Create Luces(i).X, Luces(i).Y, MapData(Luces(i).X, Luces(i).Y).luz.color, MapData(Luces(i).X, Luces(i).Y).luz.Rango, Luces(i).X & Luces(i).Y
               
                Next i

            End If

            If .NumeroOBJs > 0 Then
                ReDim Objetos(1 To .NumeroOBJs)
                Get #fh, , Objetos

                For i = 1 To .NumeroOBJs
                    '                 MapData(Objetos(i).X, Objetos(i).Y).OBJInfo.OBJIndex = Objetos(i).OBJIndex
                    '   MapData(Objetos(i).X, Objetos(i).Y).OBJInfo.Amount = Objetos(i).ObjAmmount
                    '    MapData(Objetos(i).X, Objetos(i).Y).ObjGrh.grhindex = ObjData(Objetos(i).OBJIndex).grhindex
       
                    '    Call InitGrh(MapData(Objetos(i).X, Objetos(i).Y).ObjGrh, MapData(Objetos(i).X, Objetos(i).Y).ObjGrh.grhindex)

                Next i

            End If
        
            frmMapaGrande.ListView1.ListItems.Clear

            frmMapaGrande.listdrop.ListItems.Clear

            If .NumeroNPCs > 0 Then
                CantNpcWorld = .NumeroNPCs
        
                ReDim NPCs(1 To .NumeroNPCs)
                Get #fh, , NPCs

                Dim c As Long
                
                For c = 1 To 1000
                    NpcWorlds(c) = 0
                Next c

                ' frmMapaGrande.ListView1.ListItems.Clear
                For i = 1 To .NumeroNPCs
                
                    NpcWorlds(NPCs(i).NpcIndex) = NpcWorlds(NPCs(i).NpcIndex) + 1

                Next i
               
                For c = 1 To 1000

                    If NpcWorlds(c) > 0 Then

                        If c > 399 And c < 450 Or c > 499 Then

                            Dim subelemento As ListItem

                            Set subelemento = frmMapaGrande.ListView1.ListItems.Add(, , NpcData(c).Name)

                            subelemento.SubItems(1) = NpcWorlds(c)
                            subelemento.SubItems(2) = c

                        End If

                    End If

                Next c
                
            End If

        End With

        Close fh

    End With
    
    #If Compresion = 1 Then
        Delete_File Windows_Temp_Dir & "mapa" & map & ".csm"
    #End If

End Sub

Public Sub CargarMapa(ByVal map As Integer)

    '**************************************************************
    'Formato de mapas optimizado para reducir el espacio que ocupan.
    'Diseñado y creado por Juan Martín Sotuyo Dodero (Maraxus) (juansotuyo@hotmail.com)
    '**************************************************************
    Dim fh           As Integer

    Dim MH           As tMapHeader

    Dim Blqs()       As tDatosBloqueados

    Dim MapRoute     As String

    Dim L1()         As tDatosGrh
    Dim L2()         As tDatosGrh
    Dim L3()         As tDatosGrh
    Dim L4()         As tDatosGrh

    Dim Triggers()   As tDatosTrigger
    Dim Luces()      As tDatosLuces
    Dim Particulas() As tDatosParticulas
    Dim Objetos()    As tDatosObjs
    
    Dim LBoundRoof As Integer, UBoundRoof As Integer

    Dim i            As Long
    Dim j            As Long

    Dim X            As Long
    Dim Y            As Long

    Dim demora       As Long
    Dim demorafinal  As Long

    demora = timeGetTime
    
    #If Compresion = 1 Then

        If Not Extract_File(Maps, App.Path & "\..\Recursos\OUTPUT\", "mapa" & map & ".csm", Windows_Temp_Dir, False) Then
            Err.Description = "¡No se puede cargar el archivo de mapas! El juego se cerrara."
            Call MsgBox(Err.Description)
            End

        End If

        MapRoute = Windows_Temp_Dir & "mapa" & map & ".csm"
    #Else
    
        MapRoute = App.Path & "\..\Recursos\Mapas\mapa" & map & ".csm"
        
    #End If
    
    'Limpiamos los efectos remantentes del mapa.
    Call LucesCuadradas.Light_Remove_All
    Call LucesRedondas.Delete_All_LigthRound(False)
    Call Graficos_Particulas.Particle_Group_Remove_All

    HayLayer4 = False
    
    For i = 1 To LastChar
        'If charlist(i).active = 1 Then
        Call EraseChar(i)
        ' End If
    Next i

    fh = FreeFile
    Open MapRoute For Binary As fh
    Get #fh, , MH
    Get #fh, , MapSize
    Get #fh, , MapDat
    
    ReDim MapData(1 To 100, 1 To 100)
    
    If MapDat.base_light = 0 Then
        Call RestaurarLuz
        
    Else
        Call SetGlobalLight(MapDat.base_light)
    End If
        
    For X = 1 To 100
        For Y = 1 To 100
            With MapData(X, Y)

                .light_value(0) = global_light
                .light_value(1) = global_light
                .light_value(2) = global_light
                .light_value(3) = global_light
                
            End With
        Next Y
    Next X
    
    ' Get #fh, , L1
    With MH

        'Cargamos Bloqueos
        
        If .NumeroBloqueados > 0 Then
            ReDim Blqs(1 To .NumeroBloqueados)
            Get #fh, , Blqs

            For i = 1 To .NumeroBloqueados
                MapData(Blqs(i).X, Blqs(i).Y).Blocked = Blqs(i).lados
            Next i
        End If
    
        'Cargamos Layer 1
        
        If .NumeroLayers(1) > 0 Then
        
            ReDim L1(1 To .NumeroLayers(1))
            Get #fh, , L1

            For i = 1 To .NumeroLayers(1)
            
                X = L1(i).X
                Y = L1(i).Y
                
                With MapData(X, Y)
            
                    .Graphic(1).GrhIndex = L1(i).GrhIndex
                    
                    ' Precalculate position
                    .Graphic(1).X = X * TilePixelWidth
                    .Graphic(1).Y = Y * TilePixelHeight
                    ' *********************
                
                    InitGrh .Graphic(1), .Graphic(1).GrhIndex
                    
                    If HayAgua(X, Y) Then
                        .Blocked = .Blocked Or FLAG_AGUA
                    End If
                    
                End With
            Next i

        End If
    
        'Cargamos Layer 2
        If .NumeroLayers(2) > 0 Then
            ReDim L2(1 To .NumeroLayers(2))
            Get #fh, , L2

            For i = 1 To .NumeroLayers(2)
                
                X = L2(i).X
                Y = L2(i).Y

                MapData(X, Y).Graphic(2).GrhIndex = L2(i).GrhIndex
                
                InitGrh MapData(X, Y).Graphic(2), MapData(X, Y).Graphic(2).GrhIndex
                
                MapData(X, Y).Blocked = MapData(X, Y).Blocked Or FLAG_COSTA
                
            Next i

        End If
    
        If .NumeroLayers(3) > 0 Then
            ReDim L3(1 To .NumeroLayers(3))
            Get #fh, , L3

            For i = 1 To .NumeroLayers(3)
            
                X = L3(i).X
                Y = L3(i).Y
            
                MapData(X, Y).Graphic(3).GrhIndex = L3(i).GrhIndex
            
                InitGrh MapData(X, Y).Graphic(3), MapData(X, Y).Graphic(3).GrhIndex
                
                If EsArbol(L3(i).GrhIndex) Then
                    MapData(X, Y).Blocked = MapData(X, Y).Blocked Or FLAG_ARBOL
                End If
            Next i

        End If
    
        If .NumeroLayers(4) > 0 Then
            HayLayer4 = True
            ReDim L4(1 To .NumeroLayers(4))
            Get #fh, , L4

            For i = 1 To .NumeroLayers(4)
                MapData(L4(i).X, L4(i).Y).Graphic(4).GrhIndex = L4(i).GrhIndex
                InitGrh MapData(L4(i).X, L4(i).Y).Graphic(4), MapData(L4(i).X, L4(i).Y).Graphic(4).GrhIndex
            Next i

        End If
    
        If .NumeroTriggers > 0 Then
            ReDim Triggers(1 To .NumeroTriggers)
            Get #fh, , Triggers
            
            For i = 1 To .NumeroTriggers
                MapData(Triggers(i).X, Triggers(i).Y).Trigger = Triggers(i).Trigger
                
                ' Transparencia de techos
                If Triggers(i).Trigger >= PRIMER_TRIGGER_TECHO Then
                    ' Array con todos los distintos tipos de triggers para techo
                    If Triggers(i).Trigger < LBoundRoof Then
                        LBoundRoof = Triggers(i).Trigger
                        ReDim Preserve RoofsLight(LBoundRoof To UBoundRoof)

                    ElseIf Triggers(i).Trigger > UBoundRoof Then
                        UBoundRoof = Triggers(i).Trigger
                        ReDim Preserve RoofsLight(LBoundRoof To UBoundRoof)
                    End If
                    
                    RoofsLight(Triggers(i).Trigger) = 255
                End If
            Next i

        End If
    
        If .NumeroParticulas > 0 Then
            ReDim Particulas(1 To .NumeroParticulas)
            Get #fh, , Particulas

            For i = 1 To .NumeroParticulas
            
                MapData(Particulas(i).X, Particulas(i).Y).particle_Index = Particulas(i).Particula
                General_Particle_Create MapData(Particulas(i).X, Particulas(i).Y).particle_Index, Particulas(i).X, Particulas(i).Y

            Next i

        End If

        If .NumeroLuces > 0 Then
            ReDim Luces(1 To .NumeroLuces)
            Get #fh, , Luces

            For i = 1 To .NumeroLuces
                MapData(Luces(i).X, Luces(i).Y).luz.Color = Luces(i).Color
                MapData(Luces(i).X, Luces(i).Y).luz.Rango = Luces(i).Rango

                If MapData(Luces(i).X, Luces(i).Y).luz.Rango <> 0 Then
                    If MapData(Luces(i).X, Luces(i).Y).luz.Rango < 100 Then
                        LucesCuadradas.Light_Create Luces(i).X, Luces(i).Y, Luces(i).Color, Luces(i).Rango, Luces(i).X & Luces(i).Y
                    Else
                        LucesRedondas.Create_Light_To_Map Luces(i).X, Luces(i).Y, Luces(i).Color, Luces(i).Rango - 99
                    End If

                End If
               
            Next i

        End If

        If .NumeroOBJs > 0 Then
            ReDim Objetos(1 To .NumeroOBJs)
            Get #fh, , Objetos

            For i = 1 To .NumeroOBJs
                MapData(Objetos(i).X, Objetos(i).Y).OBJInfo.OBJIndex = Objetos(i).OBJIndex
                MapData(Objetos(i).X, Objetos(i).Y).OBJInfo.Amount = Objetos(i).ObjAmmount
                MapData(Objetos(i).X, Objetos(i).Y).ObjGrh.GrhIndex = ObjData(Objetos(i).OBJIndex).GrhIndex
                Call InitGrh(MapData(Objetos(i).X, Objetos(i).Y).ObjGrh, MapData(Objetos(i).X, Objetos(i).Y).ObjGrh.GrhIndex)

            Next i

        End If

    End With

    Close fh
    
    #If Compresion = 1 Then
        Delete_File Windows_Temp_Dir & "mapa" & map & ".csm"
    #End If


End Sub

Public Sub CargarParticulas()

    '*************************************
    'Coded by OneZero (onezero_ss@hotmail.com)
    'Last Modified: 6/4/03
    'Loads the Particles.ini file to the ComboBox
    'Edited by Juan Martín Sotuyo Dodero to add speed and life
    '*************************************
    Dim loopc      As Long
    Dim i          As Long
    Dim GrhListing As String
    Dim TempSet    As String
    Dim ColorSet   As Long
    Dim temp       As Integer

    #If Compresion = 1 Then

        If Not Extract_File(Scripts, App.Path & "\..\Recursos\OUTPUT\", "particles.ini", Windows_Temp_Dir, False) Then
            Err.Description = "¡No se puede cargar el archivo de particles.ini!"
            MsgBox Err.Description

        End If

        StreamFile = Windows_Temp_Dir & "particles.ini"
    #Else
        StreamFile = App.Path & "\..\Recursos\init\particles.ini"
    #End If

    ParticulasTotales = Val(General_Var_Get(StreamFile, "INIT", "Total"))
    
    'resize StreamData array
    ReDim StreamData(1 To ParticulasTotales) As Stream
    
    'fill StreamData array with info from Particles.ini
    For loopc = 1 To ParticulasTotales
        StreamData(loopc).Name = General_Var_Get(StreamFile, Val(loopc), "Name")
        StreamData(loopc).NumOfParticles = General_Var_Get(StreamFile, Val(loopc), "NumOfParticles")
        StreamData(loopc).x1 = General_Var_Get(StreamFile, Val(loopc), "X1")
        StreamData(loopc).y1 = General_Var_Get(StreamFile, Val(loopc), "Y1")
        StreamData(loopc).x2 = General_Var_Get(StreamFile, Val(loopc), "X2")
        StreamData(loopc).y2 = General_Var_Get(StreamFile, Val(loopc), "Y2")
        StreamData(loopc).Angle = General_Var_Get(StreamFile, Val(loopc), "Angle")
        StreamData(loopc).vecx1 = General_Var_Get(StreamFile, Val(loopc), "VecX1")
        StreamData(loopc).vecx2 = General_Var_Get(StreamFile, Val(loopc), "VecX2")
        StreamData(loopc).vecy1 = General_Var_Get(StreamFile, Val(loopc), "VecY1")
        StreamData(loopc).vecy2 = General_Var_Get(StreamFile, Val(loopc), "VecY2")
        StreamData(loopc).life1 = General_Var_Get(StreamFile, Val(loopc), "Life1")
        StreamData(loopc).life2 = General_Var_Get(StreamFile, Val(loopc), "Life2")
        StreamData(loopc).friction = General_Var_Get(StreamFile, Val(loopc), "Friction")
        StreamData(loopc).spin = General_Var_Get(StreamFile, Val(loopc), "Spin")
        StreamData(loopc).spin_speedL = General_Var_Get(StreamFile, Val(loopc), "Spin_SpeedL")
        StreamData(loopc).spin_speedH = General_Var_Get(StreamFile, Val(loopc), "Spin_SpeedH")
        StreamData(loopc).AlphaBlend = General_Var_Get(StreamFile, Val(loopc), "AlphaBlend")
        StreamData(loopc).gravity = General_Var_Get(StreamFile, Val(loopc), "Gravity")
        StreamData(loopc).grav_strength = General_Var_Get(StreamFile, Val(loopc), "Grav_Strength")
        StreamData(loopc).bounce_strength = General_Var_Get(StreamFile, Val(loopc), "Bounce_Strength")
        StreamData(loopc).XMove = General_Var_Get(StreamFile, Val(loopc), "XMove")
        StreamData(loopc).YMove = General_Var_Get(StreamFile, Val(loopc), "YMove")
        StreamData(loopc).move_x1 = General_Var_Get(StreamFile, Val(loopc), "move_x1")
        StreamData(loopc).move_x2 = General_Var_Get(StreamFile, Val(loopc), "move_x2")
        StreamData(loopc).move_y1 = General_Var_Get(StreamFile, Val(loopc), "move_y1")
        StreamData(loopc).move_y2 = General_Var_Get(StreamFile, Val(loopc), "move_y2")
        StreamData(loopc).life_counter = General_Var_Get(StreamFile, Val(loopc), "life_counter")
        StreamData(loopc).speed = Val(General_Var_Get(StreamFile, Val(loopc), "Speed"))
        temp = General_Var_Get(StreamFile, Val(loopc), "resize")
        StreamData(loopc).grh_resize = IIf((temp = -1), True, False)
        StreamData(loopc).grh_resizex = General_Var_Get(StreamFile, Val(loopc), "rx")
        StreamData(loopc).grh_resizey = General_Var_Get(StreamFile, Val(loopc), "ry")
        
        StreamData(loopc).NumGrhs = General_Var_Get(StreamFile, Val(loopc), "NumGrhs")
        ReDim StreamData(loopc).grh_list(1 To StreamData(loopc).NumGrhs)
        GrhListing = General_Var_Get(StreamFile, Val(loopc), "Grh_List")
        
        For i = 1 To StreamData(loopc).NumGrhs
            StreamData(loopc).grh_list(i) = General_Field_Read(str(i), GrhListing, ",")
        Next i

        StreamData(loopc).grh_list(i - 1) = StreamData(loopc).grh_list(i - 1)
        
        For ColorSet = 1 To 4
            TempSet = General_Var_Get(StreamFile, Val(loopc), "ColorSet" & ColorSet)
            StreamData(loopc).colortint(ColorSet - 1).R = General_Field_Read(1, TempSet, ",")
            StreamData(loopc).colortint(ColorSet - 1).G = General_Field_Read(2, TempSet, ",")
            StreamData(loopc).colortint(ColorSet - 1).B = General_Field_Read(3, TempSet, ",")
        Next ColorSet
        
    Next loopc
        
    #If Compresion = 1 Then
        Delete_File Windows_Temp_Dir & "particles.ini"
    #End If

End Sub

Public Sub CargarParticulasBinary()

    '*************************************
    'Coded by OneZero (onezero_ss@hotmail.com)
    'Last Modified: 6/4/03
    'Loads the Particles.ini file to the ComboBox
    'Edited by Juan Martín Sotuyo Dodero to add speed and life
    '*************************************
    Dim loopc      As Long
    Dim i          As Long
    Dim GrhListing As String
    Dim TempSet    As String
    Dim ColorSet   As Long
    Dim temp       As Integer

    Dim handle     As Integer

    'Open files
    handle = FreeFile()

    #If Compresion = 1 Then

        If Not Extract_File(Scripts, App.Path & "\..\Recursos\OUTPUT\", "particles.ind", Windows_Temp_Dir, False) Then
            Err.Description = "¡No se puede cargar el archivo de particles.ind!"
            MsgBox Err.Description

        End If

        StreamFile = Windows_Temp_Dir & "particles.ind"
    #Else
        StreamFile = App.Path & "\..\Recursos\init\particles.ind"
    #End If

    Dim N As Integer
    
    N = FreeFile()

    Open StreamFile For Binary Access Read As #N
    'num de cabezas
    Get #N, , ParticulasTotales

    ReDim StreamData(1 To ParticulasTotales) As Stream

    'fill StreamData array with info from Particles.ini
    For loopc = 1 To ParticulasTotales
        Get #N, , StreamData(loopc)
    Next loopc
    
    Close #N

    Exit Sub
    ParticulasTotales = Val(General_Var_Get(StreamFile, "INIT", "Total"))
    
    'resize StreamData array
    
    'fill StreamData array with info from Particles.ini
    For loopc = 1 To ParticulasTotales

        temp = General_Var_Get(StreamFile, Val(loopc), "resize")
        StreamData(loopc).grh_resize = IIf((temp = -1), True, False)
        StreamData(loopc).grh_resizex = General_Var_Get(StreamFile, Val(loopc), "rx")
        StreamData(loopc).grh_resizey = General_Var_Get(StreamFile, Val(loopc), "ry")
        
        StreamData(loopc).NumGrhs = General_Var_Get(StreamFile, Val(loopc), "NumGrhs")
        ReDim StreamData(loopc).grh_list(1 To StreamData(loopc).NumGrhs)
        GrhListing = General_Var_Get(StreamFile, Val(loopc), "Grh_List")
        
        For i = 1 To StreamData(loopc).NumGrhs
            StreamData(loopc).grh_list(i) = General_Field_Read(str(i), GrhListing, ",")
        Next i

        StreamData(loopc).grh_list(i - 1) = StreamData(loopc).grh_list(i - 1)
        
        For ColorSet = 1 To 4
            TempSet = General_Var_Get(StreamFile, Val(loopc), "ColorSet" & ColorSet)
            StreamData(loopc).colortint(ColorSet - 1).R = General_Field_Read(1, TempSet, ",")
            StreamData(loopc).colortint(ColorSet - 1).G = General_Field_Read(2, TempSet, ",")
            StreamData(loopc).colortint(ColorSet - 1).B = General_Field_Read(3, TempSet, ",")
        Next ColorSet
        
    Next loopc
        
    #If Compresion = 1 Then
        Delete_File Windows_Temp_Dir & "particles.ini"
    #End If

End Sub


Public Sub CargarIndicesOBJBinary()

    Dim Obj       As Integer
    Dim Npc       As Integer
    Dim Hechizo   As Integer
    Dim i         As Integer
    Dim SearchVar As String

    #If Compresion = 1 Then

        If Not Extract_File(Scripts, App.Path & "\..\Recursos\OUTPUT\", "localindex.ind", Windows_Temp_Dir, False) Then
            Err.Description = "¡No se puede cargar el archivo de localindex.ind!"
            MsgBox Err.Description

        End If
    
        ObjFile = Windows_Temp_Dir & "localindex.ind"
        
    #Else
    
        ObjFile = App.Path & "\..\Recursos\init\localindex.ind"
        
    #End If

    Dim handle As Integer

    'Open files
    handle = FreeFile()

    Dim N As Integer
    
    N = FreeFile()

    Open ObjFile For Binary Access Read As #N
    'num de cabezas
    Get #N, , NumOBJs

    ReDim ObjData(0 To NumOBJs) As ObjDatas
    
    'ReDim NpcData(0 To NumNpcs) As NpcDatas
    ' ReDim HechizoData(0 To NumHechizos) As HechizoDatas
    'ReDim Locale_SMG(0 To NumLocaleMsg) As String

    For Obj = 1 To NumOBJs
        DoEvents
        Get #N, , ObjData(Obj)
    Next Obj

    Get #N, , NumNpcs

    ReDim NpcData(0 To NumNpcs) As NpcDatas
    
    For Npc = 1 To NumNpcs
        Get #N, , NpcData(Npc)
    Next Npc

    Get #N, , NumHechizos
    
    ReDim HechizoData(0 To NumHechizos) As HechizoDatas
    
    For Hechizo = 1 To NumHechizos
        DoEvents
        Get #N, , HechizoData(Npc)
    Next Hechizo

    Get #N, , NumLocaleMsg
    
    ReDim Locale_SMG(0 To NumLocaleMsg) As String
    
    For i = 1 To NumLocaleMsg
        Get #N, , Locale_SMG(i)
    Next i

    'Modificadores de Raza
    For i = 1 To NUMRAZAS

        With ModRaza(i)
            Get #N, , .Fuerza
            Get #N, , .Agilidad
            Get #N, , .Inteligencia
            Get #N, , .Constitucion
            Get #N, , .Carisma
        End With

    Next i
    
    #If Compresion = 1 Then
        Delete_File Windows_Temp_Dir & "localindex.ind"
    #End If

    Close #N

    Exit Sub

End Sub

Public Sub CargarIndicesOBJ()

    Dim Obj     As Integer

    Dim Npc     As Integer

    Dim Hechizo As Integer

    Dim i       As Integer
    
    #If Compresion = 1 Then

        If Not Extract_File(Scripts, App.Path & "\..\Recursos\OUTPUT\", "localindex.dat", Windows_Temp_Dir, False) Then
            Err.Description = "¡No se puede cargar el archivo de localindex.dat!"
            MsgBox Err.Description

        End If
    
        ObjFile = Windows_Temp_Dir & "localindex.dat"
    #Else
        ObjFile = App.Path & "\..\Recursos\init\localindex.dat"
    #End If
            
    Dim Leer As New clsIniManager
    Call Leer.Initialize(ObjFile)

    NumOBJs = Val(Leer.GetValue("INIT", "NumObjs"))
    NumNpcs = Val(Leer.GetValue("INIT", "NumNpcs"))
    NumHechizos = Val(Leer.GetValue("INIT", "NumeroHechizo"))
    NumHechizos = Val(Leer.GetValue("INIT", "NumeroHechizo"))
    NumLocaleMsg = Val(Leer.GetValue("INIT", "NumLocaleMsg"))
    
    NumQuest = Val(Leer.GetValue("INIT", "NUMQUESTS"))
    
    NumSug = Val(Leer.GetValue("INIT", "NUMSUGERENCIAS"))
    
    ReDim ObjData(0 To NumOBJs) As ObjDatas
    ReDim NpcData(0 To NumNpcs) As NpcDatas
    ReDim HechizoData(0 To NumHechizos) As HechizoDatas
    ReDim Locale_SMG(0 To NumLocaleMsg) As String
    
    
    ReDim QuestList(1 To NumQuest)

    
    ReDim Sugerencia(1 To NumSug) As String

    
    ReDim PosMap(1 To NumQuest) As Integer

    For Obj = 1 To NumOBJs
        DoEvents
        ObjData(Obj).GrhIndex = Val(Leer.GetValue("OBJ" & Obj, "grhindex"))
        ObjData(Obj).Name = Leer.GetValue("OBJ" & Obj, "Name")
        ObjData(Obj).MinDef = Val(Leer.GetValue("OBJ" & Obj, "MinDef"))
        ObjData(Obj).MaxDef = Val(Leer.GetValue("OBJ" & Obj, "MaxDef"))
        ObjData(Obj).MinHit = Val(Leer.GetValue("OBJ" & Obj, "MinHit"))
        ObjData(Obj).MaxHit = Val(Leer.GetValue("OBJ" & Obj, "MaxHit"))
        ObjData(Obj).ObjType = Val(Leer.GetValue("OBJ" & Obj, "ObjType"))
        ObjData(Obj).info = Leer.GetValue("OBJ" & Obj, "Info")
        ObjData(Obj).Texto = Leer.GetValue("OBJ" & Obj, "Texto")
        ObjData(Obj).CreaGRH = Leer.GetValue("OBJ" & Obj, "CreaGRH")
        ObjData(Obj).CreaLuz = Leer.GetValue("OBJ" & Obj, "CreaLuz")
        ObjData(Obj).CreaParticulaPiso = Val(Leer.GetValue("OBJ" & Obj, "CreaParticulaPiso"))
        ObjData(Obj).proyectil = Val(Leer.GetValue("OBJ" & Obj, "proyectil"))
        ObjData(Obj).Raices = Val(Leer.GetValue("OBJ" & Obj, "Raices"))
        ObjData(Obj).Madera = Val(Leer.GetValue("OBJ" & Obj, "Madera"))
        ObjData(Obj).PielLobo = Val(Leer.GetValue("OBJ" & Obj, "PielLobo"))
        ObjData(Obj).PielOsoPardo = Val(Leer.GetValue("OBJ" & Obj, "PielOsoPardo"))
        ObjData(Obj).PielOsoPolar = Val(Leer.GetValue("OBJ" & Obj, "PielOsoPolar"))
        ObjData(Obj).LingH = Val(Leer.GetValue("OBJ" & Obj, "LingH"))
        ObjData(Obj).LingP = Val(Leer.GetValue("OBJ" & Obj, "LingP"))
        ObjData(Obj).LingO = Val(Leer.GetValue("OBJ" & Obj, "LingO"))
        ObjData(Obj).Destruye = Val(Leer.GetValue("OBJ" & Obj, "Destruye"))
        ObjData(Obj).SkHerreria = Val(Leer.GetValue("OBJ" & Obj, "SkHerreria"))
        ObjData(Obj).SkPociones = Val(Leer.GetValue("OBJ" & Obj, "SkPociones"))
        ObjData(Obj).Sksastreria = Val(Leer.GetValue("OBJ" & Obj, "Sksastreria"))
        ObjData(Obj).Valor = Val(Leer.GetValue("OBJ" & Obj, "Valor"))
        
    Next Obj
    
    Dim aux   As String

    Dim loopc As Byte
    
    For Npc = 1 To NumNpcs
        DoEvents
        
        NpcData(Npc).Name = Leer.GetValue("npc" & Npc, "Name")

        If NpcData(Npc).Name = "" Then
            NpcData(Npc).Name = "Vacio"

        End If

        NpcData(Npc).desc = Leer.GetValue("npc" & Npc, "desc")
        NpcData(Npc).Body = Val(Leer.GetValue("npc" & Npc, "Body"))
        NpcData(Npc).exp = Val(Leer.GetValue("npc" & Npc, "exp"))
        NpcData(Npc).Head = Val(Leer.GetValue("npc" & Npc, "Head"))
        NpcData(Npc).Hp = Val(Leer.GetValue("npc" & Npc, "Hp"))
        NpcData(Npc).MaxHit = Val(Leer.GetValue("npc" & Npc, "MaxHit"))
        NpcData(Npc).MinHit = Val(Leer.GetValue("npc" & Npc, "MinHit"))
        NpcData(Npc).oro = Val(Leer.GetValue("npc" & Npc, "oro"))
        
        NpcData(Npc).ExpClan = Val(Leer.GetValue("npc" & Npc, "GiveEXPClan"))
       
        aux = Val(Leer.GetValue("npc" & Npc, "NumQuiza"))

        If aux = 0 Then
            NpcData(Npc).NumQuiza = 0
        Else
            NpcData(Npc).NumQuiza = Val(aux)
            ReDim NpcData(Npc).QuizaDropea(1 To NpcData(Npc).NumQuiza) As Integer

            For loopc = 1 To NpcData(Npc).NumQuiza
               
                NpcData(Npc).QuizaDropea(loopc) = Val(Leer.GetValue("npc" & Npc, "QuizaDropea" & loopc))
                ' Debug.Print NpcData(Npc).QuizaDropea(loopc)
            Next loopc

        End If

    Next Npc
    
    For Hechizo = 1 To NumHechizos
        DoEvents
        HechizoData(Hechizo).nombre = Leer.GetValue("Hechizo" & Hechizo, "Nombre")
        HechizoData(Hechizo).desc = Leer.GetValue("Hechizo" & Hechizo, "desc")
        HechizoData(Hechizo).PalabrasMagicas = Leer.GetValue("Hechizo" & Hechizo, "PalabrasMagicas")
        HechizoData(Hechizo).HechizeroMsg = Leer.GetValue("Hechizo" & Hechizo, "HechizeroMsg")
        HechizoData(Hechizo).TargetMsg = Leer.GetValue("Hechizo" & Hechizo, "TargetMsg")
        HechizoData(Hechizo).PropioMsg = Leer.GetValue("Hechizo" & Hechizo, "PropioMsg")
        HechizoData(Hechizo).ManaRequerido = Val(Leer.GetValue("Hechizo" & Hechizo, "ManaRequerido"))
        HechizoData(Hechizo).MinSkill = Val(Leer.GetValue("Hechizo" & Hechizo, "MinSkill"))
        HechizoData(Hechizo).StaRequerido = Val(Leer.GetValue("Hechizo" & Hechizo, "StaRequerido"))
        HechizoData(Hechizo).IconoIndex = Val(Leer.GetValue("Hechizo" & Hechizo, "IconoIndex"))
        'HechizoData(Hechizo).IconoIndex = 35696
    Next Hechizo
    
    Hechizo = 1
    
    For Hechizo = 1 To 350
        DoEvents
        NameMaps(Hechizo).Name = Leer.GetValue("NameMapa", "Mapa" & Hechizo)
        NameMaps(Hechizo).desc = Leer.GetValue("NameMapa", "Mapa" & Hechizo & "Desc")
    Next Hechizo
    
    For Hechizo = 1 To NumQuest
        DoEvents
        
        QuestList(Hechizo).nombre = Leer.GetValue("QUEST" & Hechizo, "NOMBRE")
        
        QuestList(Hechizo).desc = Leer.GetValue("QUEST" & Hechizo, "DESC")
        QuestList(Hechizo).NextQuest = Leer.GetValue("QUEST" & Hechizo, "NEXTQUEST")
        QuestList(Hechizo).DescFinal = Leer.GetValue("QUEST" & Hechizo, "DESCFINAL")
        QuestList(Hechizo).RequiredLevel = Leer.GetValue("QUEST" & Hechizo, "RequiredLevel")
        PosMap(Hechizo) = Leer.GetValue("QUEST" & Hechizo, "PosMap")
    Next Hechizo
    
    For Hechizo = 1 To NumSug
        DoEvents
        Sugerencia(Hechizo) = Leer.GetValue("SUGERENCIAS", "SUGERENCIA" & Hechizo)
    Next Hechizo
    
    For i = 1 To NumLocaleMsg
        DoEvents
        Locale_SMG(i) = Leer.GetValue("msg", "Msg" & i)
    Next i
    
    Dim SearchVar As String
    
    'Modificadores de Raza
    For i = 1 To NUMRAZAS

        With ModRaza(i)
            SearchVar = Replace(ListaRazas(i), " ", vbNullString)
            
            .Fuerza = Val(Leer.GetValue("MODRAZA", SearchVar + "Fuerza"))
            .Agilidad = Val(Leer.GetValue("MODRAZA", SearchVar + "Agilidad"))
            .Inteligencia = Val(Leer.GetValue("MODRAZA", SearchVar + "Inteligencia"))
            .Constitucion = Val(Leer.GetValue("MODRAZA", SearchVar + "Constitucion"))
            .Carisma = Val(Leer.GetValue("MODRAZA", SearchVar + "Carisma"))

        End With

    Next i
    
    #If Compresion = 1 Then
        Delete_File Windows_Temp_Dir & "localindex.dat"
    #End If

End Sub

Public Sub Cargarmapsworlddata()

    'Ladder
    Dim MapFile As String

    Dim i       As Integer

    #If Compresion = 1 Then

        If Not Extract_File(Scripts, App.Path & "\..\Recursos\OUTPUT\", "mapsworlddata.dat", Windows_Temp_Dir, False) Then
            Err.Description = "¡No se puede cargar el archivo de mapsworlddata.dat!"
            MsgBox Err.Description

        End If
    
        MapFile = Windows_Temp_Dir & "mapsworlddata.dat"
    #Else
        MapFile = App.Path & "\..\Recursos\init\mapsworlddata.dat"
    #End If

    Dim Leer As New clsIniManager
    Call Leer.Initialize(MapFile)

    WordMapaNum = Val(Leer.GetValue("WORLDMAP", "NumMap"))
    DungeonDataNum = Val(Leer.GetValue("DUNGEON", "NumMap"))
     
    ReDim WordMapa(1 To WordMapaNum) As String
    ReDim DungeonData(1 To DungeonDataNum) As String
    
    For i = 1 To WordMapaNum
        WordMapa(i) = Val(Leer.GetValue("WORLDMAP", i))
    Next i
    
    For i = 1 To DungeonDataNum
        DungeonData(i) = Val(Leer.GetValue("DUNGEON", i))
    Next i

    #If Compresion = 1 Then
        Delete_File Windows_Temp_Dir & "mapsworlddata.dat"
    #End If

End Sub

Sub CargarCabezas()

    Dim N            As Integer

    Dim i            As Long

    Dim Numheads     As Integer

    Dim Miscabezas() As tIndiceCabeza
    
    N = FreeFile()
    
    #If Compresion = 1 Then

        If Not Extract_File(Scripts, App.Path & "\..\Recursos\OUTPUT\", "cabezas.ind", Windows_Temp_Dir, False) Then
            Err.Description = "¡No se puede cargar el archivo de Cabezas.ind!"
            MsgBox Err.Description

        End If

        Open Windows_Temp_Dir & "cabezas.ind" For Binary Access Read As #N
    #Else
        Open App.Path & "\..\Recursos\init\cabezas.ind" For Binary Access Read As #N
    #End If

    'cabecera
    Get #N, , MiCabecera
    
    'num de cabezas
    Get #N, , Numheads
    
    'Resize array
    ReDim HeadData(0 To Numheads) As HeadData
    ReDim Miscabezas(0 To Numheads) As tIndiceCabeza
    
    For i = 1 To Numheads
        Get #N, , Miscabezas(i)
        
        If Miscabezas(i).Head(1) Then
            Call InitGrh(HeadData(i).Head(1), Miscabezas(i).Head(1), 0)
            Call InitGrh(HeadData(i).Head(2), Miscabezas(i).Head(2), 0)
            Call InitGrh(HeadData(i).Head(3), Miscabezas(i).Head(3), 0)
            Call InitGrh(HeadData(i).Head(4), Miscabezas(i).Head(4), 0)

        End If

    Next i
    
    Close #N
    
    #If Compresion = 1 Then
        Delete_File Windows_Temp_Dir & "cabezas.ind"
    #End If
    
End Sub

Sub CargarCascos()

    Dim N            As Integer

    Dim i            As Long

    Dim NumCascos    As Integer

    Dim Miscabezas() As tIndiceCabeza
    
    N = FreeFile()
  
    #If Compresion = 1 Then

        If Not Extract_File(Scripts, App.Path & "\..\Recursos\OUTPUT\", "cascos.ind", Windows_Temp_Dir, False) Then
            Err.Description = "¡No se puede cargar el archivo de Cabezas.ind!"
            MsgBox Err.Description

        End If

        Open Windows_Temp_Dir & "cascos.ind" For Binary Access Read As #N
    #Else
        Open App.Path & "\..\Recursos\init\cascos.ind" For Binary Access Read As #N
    #End If
    
    'cabecera
    Get #N, , MiCabecera
    
    'num de cabezas
    Get #N, , NumCascos
    
    'Resize array
    ReDim CascoAnimData(0 To NumCascos) As HeadData
    ReDim Miscabezas(0 To NumCascos) As tIndiceCabeza
    
    For i = 1 To NumCascos
        Get #N, , Miscabezas(i)
        
        If Miscabezas(i).Head(1) Then
            Call InitGrh(CascoAnimData(i).Head(1), Miscabezas(i).Head(1), 0)
            Call InitGrh(CascoAnimData(i).Head(2), Miscabezas(i).Head(2), 0)
            Call InitGrh(CascoAnimData(i).Head(3), Miscabezas(i).Head(3), 0)
            Call InitGrh(CascoAnimData(i).Head(4), Miscabezas(i).Head(4), 0)

        End If

    Next i
    
    Close #N
    
    #If Compresion = 1 Then
        Delete_File Windows_Temp_Dir & "cascos.ind"
    #End If

End Sub

Sub CargarCuerpos()

    Dim N            As Integer

    Dim i            As Long

    Dim NumCuerpos   As Integer

    Dim MisCuerpos() As tIndiceCuerpo
    
    N = FreeFile()
    
    #If Compresion = 1 Then

        If Not Extract_File(Scripts, App.Path & "\..\Recursos\OUTPUT\", "personajes.ind", Windows_Temp_Dir, False) Then
            Err.Description = "¡No se puede cargar el archivo de personajes.ind!"
            MsgBox Err.Description

        End If

        Open Windows_Temp_Dir & "personajes.ind" For Binary Access Read As #N
    #Else
        Open App.Path & "\..\Recursos\init\personajes.ind" For Binary Access Read As #N
    #End If
    
    'cabecera
    Get #N, , MiCabecera
    
    'num de cabezas
    Get #N, , NumCuerpos
    
    'Resize array
    ReDim BodyData(0 To NumCuerpos) As BodyData
    ReDim MisCuerpos(0 To NumCuerpos) As tIndiceCuerpo
    
    For i = 1 To NumCuerpos
        Get #N, , MisCuerpos(i)
        
        If MisCuerpos(i).Body(1) Then
            InitGrh BodyData(i).Walk(1), MisCuerpos(i).Body(1), 0
            InitGrh BodyData(i).Walk(2), MisCuerpos(i).Body(2), 0
            InitGrh BodyData(i).Walk(3), MisCuerpos(i).Body(3), 0
            InitGrh BodyData(i).Walk(4), MisCuerpos(i).Body(4), 0
            
            BodyData(i).HeadOffset.X = MisCuerpos(i).HeadOffsetX
            BodyData(i).HeadOffset.Y = MisCuerpos(i).HeadOffsetY

        End If

    Next i
    
    Close #N
    #If Compresion = 1 Then
        Delete_File Windows_Temp_Dir & "personajes.ind"
    #End If

End Sub

Sub CargarFxs()

    Dim N      As Integer

    Dim i      As Long

    Dim NumFxs As Integer
    
    N = FreeFile()

    #If Compresion = 1 Then

        If Not Extract_File(Scripts, App.Path & "\..\Recursos\OUTPUT\", "fxs.ind", Windows_Temp_Dir, False) Then
            Err.Description = "¡No se puede cargar el archivo de fxs.ind!"
            MsgBox Err.Description

        End If

        Open Windows_Temp_Dir & "fxs.ind" For Binary Access Read As #N
    #Else
        Open App.Path & "\..\Recursos\init\fxs.ind" For Binary Access Read As #N
    #End If
    
    'cabecera
    Get #N, , MiCabecera
    
    'num de cabezas
    Get #N, , NumFxs
    
    'Resize array
    ReDim FxData(1 To NumFxs) As tIndiceFx
    
    For i = 1 To NumFxs
        Get #N, , FxData(i)
    Next i
    
    Close #N
    
    #If Compresion = 1 Then
        Delete_File Windows_Temp_Dir & "fxs.ind"
    #End If

End Sub

Public Function LoadGrhData() As Boolean

    On Error GoTo ErrorHandler

    Dim grh         As Long
    Dim Frame       As Long
    Dim grhCount    As Long
    Dim handle      As Integer
    Dim fileVersion As Long
    
    'Open files
    handle = FreeFile()
    
    #If Compresion = 1 Then

        If Not Extract_File(Scripts, App.Path & "\..\Recursos\OUTPUT\", "graficos.ind", Windows_Temp_Dir, False) Then
            Err.Description = "¡No se puede cargar el archivo de recurso!"
            GoTo ErrorHandler

        End If
    
        Open Windows_Temp_Dir & "graficos.ind" For Binary Access Read As #handle
    #Else
        Open App.Path & "\..\Recursos\init\graficos.ind" For Binary Access Read As #handle
    #End If
    
    'Get file version
    Get #handle, , fileVersion
    
    'Get number of grhs
    Get #handle, , grhCount
    
    'Resize arrays
    ReDim GrhData(1 To grhCount) As GrhData
    
    MaxGrh = grhCount

    Dim Fin As Boolean

    Fin = False

    While Not EOF(handle) And Fin = False

        Get #handle, , grh

        With GrhData(grh)
        
            GrhData(grh).active = True
            'Get number of frames
            Get #handle, , .NumFrames

            If .NumFrames <= 0 Then GoTo ErrorHandler
            
            ReDim .Frames(1 To GrhData(grh).NumFrames)
            
            If .NumFrames > 1 Then

                'Read a animation GRH set
                For Frame = 1 To .NumFrames
                    Get #handle, , .Frames(Frame)

                    If .Frames(Frame) <= 0 Or .Frames(Frame) > grhCount Then
                        GoTo ErrorHandler

                    End If

                Next Frame
                
                Get #handle, , GrhData(grh).speed
                
                If .speed <= 0 Then GoTo ErrorHandler
                
                'Compute width and height
                .pixelWidth = GrhData(.Frames(1)).pixelWidth

                If .pixelWidth <= 0 Then GoTo ErrorHandler
                
                .pixelHeight = GrhData(.Frames(1)).pixelHeight

                If .pixelHeight <= 0 Then GoTo ErrorHandler
                                                
                .TileWidth = GrhData(.Frames(1)).TileWidth

                If .TileWidth <= 0 Then GoTo ErrorHandler

                .TileHeight = GrhData(.Frames(1)).TileHeight

                If .TileHeight <= 0 Then GoTo ErrorHandler
            Else
                'Read in normal GRH data
                Get #handle, , .FileNum

                If .FileNum <= 0 Then GoTo ErrorHandler
                                
                Get #handle, , GrhData(grh).sX

                If .sX < 0 Then GoTo ErrorHandler
                
                Get #handle, , GrhData(grh).sY

                If .sY < 0 Then GoTo ErrorHandler
                
                Get #handle, , GrhData(grh).pixelWidth

                If .pixelWidth <= 0 Then GoTo ErrorHandler
                
                Get #handle, , GrhData(grh).pixelHeight

                If .pixelHeight <= 0 Then GoTo ErrorHandler
                
                'Compute width and height
                .TileWidth = .pixelWidth / TilePixelHeight
                .TileHeight = .pixelHeight / TilePixelWidth

                .Frames(1) = grh

            End If

        End With

        If grh = MaxGrh Then Fin = True
    Wend

    Close #handle
    
    LoadGrhData = True
    
    #If Compresion = 1 Then
        Delete_File Windows_Temp_Dir & "graficos.ind"
    #End If

    Exit Function

ErrorHandler:
    LoadGrhData = False
    MsgBox "Error " & Err.Description & " durante la carga de Grh.dat! La carga se ha detenido en GRH: " & grh
    
End Function

Public Function CargarMiniMap()

    Dim count  As Long

    Dim handle As Integer

    handle = FreeFile
    
    #If Compresion = 1 Then

        If Not Extract_File(Scripts, App.Path & "\..\Recursos\OUTPUT\", "minimap.bin", Windows_Temp_Dir, False) Then
            Err.Description = "¡No se puede cargar el archivo de recurso!"
            GoTo ErrorHandler

        End If
    
        Open Windows_Temp_Dir & "minimap.bin" For Binary Access Read As #handle
    #Else
    
        Open App.Path & "\..\Recursos\init\minimap.bin" For Binary Access Read As #handle
        
    #End If

    For count = 1 To MaxGrh

        If GrhData(count).active Then
            Get #handle, , GrhData(count).MiniMap_color

        End If

    Next count
    
    Close #handle
    
    #If Compresion = 1 Then
        Delete_File Windows_Temp_Dir & "minimap.bin"
    #End If
    
    Exit Function

ErrorHandler:
    CargarMiniMap = False
    MsgBox "Error " & Err.Description & " durante la carga de Grh.dat! La carga se ha detenido en GRH: " & count

End Function


Sub CargarAnimArmas()

    On Error Resume Next

    Dim loopc As Long

    Dim Arch  As String
    
    #If Compresion = 1 Then

        If Not Extract_File(Scripts, App.Path & "\..\Recursos\OUTPUT\", "armas.dat", Windows_Temp_Dir, False) Then
            Err.Description = "¡No se puede cargar el archivo de armas.dat!"
            MsgBox Err.Description

        End If

        Arch = Windows_Temp_Dir & "armas.dat"
    #Else
        Arch = App.Path & "\..\Recursos\init\armas.dat"
    #End If
    
    NumWeaponAnims = Val(GetVar(Arch, "INIT", "NumArmas"))
    
    ReDim WeaponAnimData(1 To NumWeaponAnims) As WeaponAnimData

    For loopc = 1 To NumWeaponAnims
        InitGrh WeaponAnimData(loopc).WeaponWalk(1), Val(GetVar(Arch, "ARMA" & loopc, "Dir1")), 0
        InitGrh WeaponAnimData(loopc).WeaponWalk(2), Val(GetVar(Arch, "ARMA" & loopc, "Dir2")), 0
        InitGrh WeaponAnimData(loopc).WeaponWalk(3), Val(GetVar(Arch, "ARMA" & loopc, "Dir3")), 0
        InitGrh WeaponAnimData(loopc).WeaponWalk(4), Val(GetVar(Arch, "ARMA" & loopc, "Dir4")), 0
    Next loopc
    
    #If Compresion = 1 Then
        Delete_File Windows_Temp_Dir & "armas.dat"
    #End If

End Sub

Sub CargarColores()

    On Error Resume Next

    Dim archivoC As String

    #If Compresion = 1 Then

        If Not Extract_File(Scripts, App.Path & "\..\Recursos\OUTPUT\", "colores.dat", Windows_Temp_Dir, False) Then
            Err.Description = "¡No se puede cargar el archivo de colores.dat!"
            MsgBox Err.Description

        End If

        archivoC = Windows_Temp_Dir & "colores.dat"
    #Else
        archivoC = App.Path & "\..\Recursos\init\colores.dat"
    #End If
    
    If Not FileExist(archivoC, vbArchive) Then
        'TODO : Si hay que reinstalar, porque no cierra???
        Call MsgBox("ERROR: no se ha podido cargar los colores. Falta el archivo colores.dat, reinstale el juego", vbCritical + vbOKOnly)
        Exit Sub

    End If
    
    Dim i As Long
    
    For i = 0 To 47 '49 y 50 reservados para ciudadano y criminal
        ColoresPJ(i).R = CByte(GetVar(archivoC, CStr(i), "R"))
        ColoresPJ(i).G = CByte(GetVar(archivoC, CStr(i), "G"))
        ColoresPJ(i).B = CByte(GetVar(archivoC, CStr(i), "B"))
    Next i
    
    ColoresPJ(50).R = CByte(GetVar(archivoC, "CR", "R"))
    ColoresPJ(50).G = CByte(GetVar(archivoC, "CR", "G"))
    ColoresPJ(50).B = CByte(GetVar(archivoC, "CR", "B"))
    
    ColoresPJ(49).R = CByte(GetVar(archivoC, "CI", "R"))
    ColoresPJ(49).G = CByte(GetVar(archivoC, "CI", "G"))
    ColoresPJ(49).B = CByte(GetVar(archivoC, "CI", "B"))
    
    ColoresPJ(48).R = CByte(GetVar(archivoC, "NE", "R"))
    ColoresPJ(48).G = CByte(GetVar(archivoC, "NE", "G"))
    ColoresPJ(48).B = CByte(GetVar(archivoC, "NE", "B"))
    
    #If Compresion = 1 Then
        Delete_File Windows_Temp_Dir & "colores.dat"
    #End If

End Sub

Sub CargarAnimEscudos()

    Dim loopc As Long

    Dim Arch  As String
    
    #If Compresion = 1 Then

        If Not Extract_File(Scripts, App.Path & "\..\Recursos\OUTPUT\", "escudos.dat", Windows_Temp_Dir, False) Then
            Err.Description = "¡No se puede cargar el archivo de escudos.dat!"
            MsgBox Err.Description

        End If

        Arch = Windows_Temp_Dir & "escudos.dat"
    #Else
        Arch = App.Path & "\..\Recursos\init\escudos.dat"
    #End If
    
    NumEscudosAnims = Val(GetVar(Arch, "INIT", "NumEscudos"))
    
    ReDim ShieldAnimData(1 To NumEscudosAnims) As ShieldAnimData
    
    For loopc = 1 To NumEscudosAnims
        InitGrh ShieldAnimData(loopc).ShieldWalk(1), Val(GetVar(Arch, "ESC" & loopc, "Dir1")), 0
        InitGrh ShieldAnimData(loopc).ShieldWalk(2), Val(GetVar(Arch, "ESC" & loopc, "Dir2")), 0
        InitGrh ShieldAnimData(loopc).ShieldWalk(3), Val(GetVar(Arch, "ESC" & loopc, "Dir3")), 0
        InitGrh ShieldAnimData(loopc).ShieldWalk(4), Val(GetVar(Arch, "ESC" & loopc, "Dir4")), 0
    Next loopc
    
    #If Compresion = 1 Then
        Delete_File Windows_Temp_Dir & "escudos.dat"
    #End If

End Sub

