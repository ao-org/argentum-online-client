Attribute VB_Name = "Recursos"
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
    FONTTYPE_CITIZEN_ARMADA
    FONTTYPE_CRIMINAL
    FONTTYPE_CRIMINAL_CAOS
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
    
    FONTTYPE_PROMEDIO_IGUAL
    FONTTYPE_PROMEDIO_MENOR
    FONTTYPE_PROMEDIO_MAYOR

    [FONTTYPE_MAX]
End Enum

Public FontTypes([FONTTYPE_MAX] - 1) As tFont
' *********************************************************
' FIN - FUENTES
' *********************************************************

' *********************************************************
' CARGA DE MAPAS
' Sinuhe - Map format .CSM
' *********************************************************
'The only current map

Public ResourcesPassword As String

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

    x As Integer
    y As Integer
    lados As Byte

End Type

Private Type tDatosGrh

    x As Integer
    y As Integer
    GrhIndex As Long

End Type

Private Type tDatosTrigger

    x As Integer
    y As Integer
    Trigger As Integer

End Type

Private Type tDatosLuces

    x As Integer
    y As Integer
    Color As RGBA
    Rango As Byte

End Type

Private Type tDatosParticulas

    x As Integer
    y As Integer
    Particula As Long

End Type

Public Type tDatosNPC

    x As Integer
    y As Integer
    NpcIndex As Integer

End Type

Private Type tDatosObjs

    x As Integer
    y As Integer
    ObjIndex As Integer
    ObjAmmount As Integer

End Type

Private Type tDatosTE

    x As Integer
    y As Integer
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

Public MapSize As tMapSize
Public MapDat   As tMapDat
Public iplst    As String
' *********************************************************
'   FIN - CARGA DE MAPAS
' *********************************************************
' *********************************************************
'   BEGIN - COMPOSED ANIMATIONS
' *********************************************************
Public Enum ePlaybackType
    Stopped
    Forward
    Pause
    Backward
End Enum
Public Type tAnimationClip
    Fx As Long 'Fx number
    LoopCount As Long 'number of loops for the animation, -1 for infintite
    Playback As ePlaybackType ' direction of the playback
    ClipTime As Long 'Calculated time for this clip
End Type

Public Type tComposedAnimation
    Clips() As tAnimationClip
End Type
' *********************************************************
'   END - COMPOSED ANIMATIONS
' *********************************************************

''''''''''''''' CARGA DE NPCS DATA MAP QUEST QCYO VIEJA nO ME IMPORTA NADA '''''''''''''''''''''''''''''''
Public Type t_Projectile
    speed As Single
    OffsetRotation As Integer
    RotationSpeed As Single
    grh As Long
    RigthGrh As Long
End Type

Public Type t_Position
    x As Single
    y As Single
End Type

Public Type t_QuestNPCMapData
    Position As t_Position
    NPCNumber As Integer
    State As Integer
End Type

Public Type t_MapNpc
    NpcList() As t_QuestNPCMapData
    NpcCount As Integer
End Type

Public ListNPCMapData() As t_MapNpc
Public Const MAX_QUESTNPCS_VISIBLE As Long = 100 'leerlo desde Quest.Dat [INIT] NumQuests =
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Type tMoldeCuerpo
    x As Long
    y As Long
    Width As Long
    Height As Long
    DirCount(1 To 4) As Byte
    TotalGrhs As Long
End Type

Private MoldesBodies() As tMoldeCuerpo
Private BodiesHeading(1 To 4) As E_Heading

Public Sub CargarRecursos()
    
    On Error GoTo CargarRecursos_Err
    
    
    If UtilizarPreCarga = 1 Then
        Call PreloadGraphics
    End If
    
    Call CargarNPCsMapData
    Call CargarParticulasBinary
    Call CargarIndicesOBJ
    Call Cargarmapsworlddata
    Call InitFontTypes
    Call CargarZonas

    'Call LoadGrhData
    Call LoadGrhIni
    Call CargarMoldes
    Call CargarCabezas
    Call CargarCascos
    Call CargarCuerpos
    Call CargarFxs
    Call LoadComposedFx
    Call CargarPasos
    Call CargarAnimArmas
    Call CargarAnimEscudos
    Call CargarColores
    Call CargarCrafteo
    
    
    
    Exit Sub

CargarRecursos_Err:
    Call RegistrarError(Err.Number, Err.Description, "Recursos.CargarRecursos", Erl)
    Resume Next
    
End Sub

''
' Initializes the fonts array

Public Sub InitFontTypes()
    
    On Error GoTo InitFonts_Err
    

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
        .red = 255
        .green = 0
        .blue = 0
        .bold = 0
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
        .red = 66
        .green = 201
        .blue = 255
        .bold = 1

    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_CONSEJOCAOS)
        .red = 255
        .green = 102
        .blue = 102
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
        .red = 6
        .green = 128
        .blue = 255
        .bold = 1

    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_CITIZEN_ARMADA)
        .red = 60
        .green = 163
        .blue = 255
        .bold = 1

    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_CRIMINAL)
        .red = 255
        .bold = 1
    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_CRIMINAL_CAOS)
        .red = 255
        .green = 51
        .blue = 51
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
    
    With FontTypes(FontTypeNames.FONTTYPE_PROMEDIO_IGUAL)
        .red = 255
        .green = 255
        .blue = 0
        .bold = 1
    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_PROMEDIO_MENOR)
        .red = 255
        .green = 0
        .blue = 0
        .bold = 1
    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_PROMEDIO_MAYOR)
        .red = 0
        .green = 255
        .blue = 0
        .bold = 1
    End With
    
    Exit Sub

InitFonts_Err:
    Call RegistrarError(Err.Number, Err.Description, "Recursos.InitFonts", Erl)
    Resume Next
    
End Sub

Public Sub CargarPasos()
    
    On Error GoTo CargarPasos_Err
    

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
    
    Pasos(CONST_AGUA).CantPasos = 2 ' OK
    ReDim Pasos(CONST_AGUA).wav(1 To 2) As Integer
    Pasos(CONST_AGUA).wav(1) = SND_NAVEGANDO
    Pasos(CONST_AGUA).wav(2) = SND_NAVEGANDO

    
    Exit Sub

CargarPasos_Err:
    Call RegistrarError(Err.Number, Err.Description, "Recursos.CargarPasos", Erl)
    Resume Next
    
End Sub

Sub CargarDatosMapa(ByVal map As Integer)
    
    On Error GoTo CargarDatosMapa_Err
    

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
    
    Dim x            As Long
    Dim y            As Long
    
    #If Compresion = 1 Then

        If Not Extract_File(Maps, App.path & "\..\Recursos\OUTPUT\", "mapa" & map & ".csm", Windows_Temp_Dir, ResourcesPassword, False) Then
            Debug.Print "Error al cargar datos del mapa " & map
            Exit Sub
        End If

        MapRoute = Windows_Temp_Dir & "mapa" & map & ".csm"
    #Else
        MapRoute = App.path & "\..\Recursos\Mapas\mapa" & map & ".csm"
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
                
                For c = 1 To UBound(NpcWorlds)
                    NpcWorlds(c) = 0
                Next c

                ' frmMapaGrande.ListView1.ListItems.Clear
                For i = 1 To .NumeroNPCs
                
                    NpcWorlds(NPCs(i).NpcIndex) = NpcWorlds(NPCs(i).NpcIndex) + 1

                Next i
               
                For c = 1 To UBound(NpcWorlds)

                    If NpcWorlds(c) > 0 Then

                        If c > 399 And c < 450 Or c > 499 Then

                            Dim subelemento As ListItem

                            Set subelemento = frmMapaGrande.ListView1.ListItems.Add(, , NpcData(c).name)

                            subelemento.SubItems(1) = NpcWorlds(c)
                            subelemento.SubItems(2) = c
                            subelemento.EnsureVisible

                        End If

                    End If

                Next c
                
            End If

        End With

    End With
    
    Close #fh
    
    #If Compresion = 1 Then
        Delete_File Windows_Temp_Dir & "mapa" & map & ".csm"
    #End If

    frmMapaGrande.ListView1.ColumnHeaders(2).Alignment = lvwColumnRight
    frmMapaGrande.ListView1.ColumnHeaders(2).Width = 30
    
    If frmMapaGrande.ListView1.ListItems.count <= 4 Then
        frmMapaGrande.ListView1.ColumnHeaders(1).Width = frmMapaGrande.ListView1.Width - 30
    Else
        frmMapaGrande.ListView1.ColumnHeaders(1).Width = frmMapaGrande.ListView1.Width - 50
    End If
    
    If frmMapaGrande.ListView1.ListItems.count > 0 Then
        Call frmMapaGrande.ListView1.ListItems.Item(1).EnsureVisible
    End If
    
    Exit Sub

CargarDatosMapa_Err:
    Call RegistrarError(Err.Number, Err.Description, "Recursos.CargarDatosMapa", Erl)
    
End Sub

Public Sub CargarMapa(ByVal map As Integer)
    
    On Error GoTo CargarMapa_Err
    

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

    Dim x            As Long
    Dim y            As Long

    Dim demora       As Long
    Dim demorafinal  As Long

    demora = GetTickCount()

    Debug.Assert map <> 0

    #If Compresion = 1 Then

        If Not Extract_File(Maps, App.path & "\..\Recursos\OUTPUT\", "mapa" & map & ".csm", Windows_Temp_Dir, ResourcesPassword, False) Then
            Err.Description = "No se pudo cargar el mapa " & map & ", el juego se cerrará. Si su personaje se encuentra en un mapa inválido, por favor, avise a un GM."
            Call MsgBox(Err.Description)
            End

        End If

        MapRoute = Windows_Temp_Dir & "mapa" & map & ".csm"
    #Else
    
        MapRoute = App.path & "\..\Recursos\Mapas\mapa" & map & ".csm"
        
    #End If
    
    'Limpiamos los efectos remantentes del mapa.
    For x = 1 To 100
        For y = 1 To 100
            With MapData(x, y)

                Call SetRGBA(.light_value(0), 0, 0, 0)
                Call SetRGBA(.light_value(1), 0, 0, 0)
                Call SetRGBA(.light_value(2), 0, 0, 0)
                Call SetRGBA(.light_value(3), 0, 0, 0)

            End With
        Next y
    Next x
    
    Call LucesCuadradas.Light_Remove_All
    Call LucesRedondas.Delete_All_LigthRound(False)
    Call Graficos_Particulas.Particle_Group_Remove_All

    HayLayer4 = False
    
    If UserPos.x = 0 Then UserPos.x = 10
    If UserPos.y = 0 Then UserPos.y = 10
    MapData(UserPos.x, UserPos.y).charindex = 0
    
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
    
    UpdateLights = True
    
    For x = 1 To 100
        For y = 1 To 100
            With MapData(x, y)

                .light_value(0) = global_light
                .light_value(1) = global_light
                .light_value(2) = global_light
                .light_value(3) = global_light

                ReDim .DialogEffects(0)

            End With
        Next y
    Next x
    
    ' Get #fh, , L1
    With MH

        'Cargamos Bloqueos
        
        If .NumeroBloqueados > 0 Then
            ReDim Blqs(1 To .NumeroBloqueados)
            Get #fh, , Blqs

            For i = 1 To .NumeroBloqueados
                MapData(Blqs(i).x, Blqs(i).y).Blocked = Blqs(i).lados
            Next i
        End If
    
        'Cargamos Layer 1
        
        If .NumeroLayers(1) > 0 Then
        
            ReDim L1(1 To .NumeroLayers(1))
            Get #fh, , L1

            For i = 1 To .NumeroLayers(1)
            
                x = L1(i).x
                y = L1(i).y
                
                With MapData(x, y)
            
                    .Graphic(1).GrhIndex = L1(i).GrhIndex
                    
                    ' Precalculate position
                    .Graphic(1).x = x * TilePixelWidth
                    .Graphic(1).y = y * TilePixelHeight
                    ' *********************
                
                    InitGrh .Graphic(1), .Graphic(1).GrhIndex
                    
                    If HayAgua(x, y) Then
                        .Blocked = .Blocked Or FLAG_AGUA
                        
                    ElseIf HayLava(x, y) Then
                        .Blocked = .Blocked Or FLAG_LAVA
                    End If
                    
                End With
            Next i

        End If
    
        'Cargamos Layer 2
        If .NumeroLayers(2) > 0 Then
            ReDim L2(1 To .NumeroLayers(2))
            Get #fh, , L2

            For i = 1 To .NumeroLayers(2)
                
                x = L2(i).x
                y = L2(i).y

                MapData(x, y).Graphic(2).GrhIndex = L2(i).GrhIndex
                
                InitGrh MapData(x, y).Graphic(2), MapData(x, y).Graphic(2).GrhIndex
                
                MapData(x, y).Blocked = MapData(x, y).Blocked Or FLAG_COSTA
                
            Next i

        End If
    
        If .NumeroLayers(3) > 0 Then
            ReDim L3(1 To .NumeroLayers(3))
            Get #fh, , L3

            For i = 1 To .NumeroLayers(3)
            
                x = L3(i).x
                y = L3(i).y
            
                MapData(x, y).Graphic(3).GrhIndex = L3(i).GrhIndex
            
                InitGrh MapData(x, y).Graphic(3), MapData(x, y).Graphic(3).GrhIndex

                If EsArbol(L3(i).GrhIndex) Then
                    MapData(x, y).Blocked = MapData(x, y).Blocked Or FLAG_ARBOL
                End If
            Next i

        End If
    
        If .NumeroLayers(4) > 0 Then
            HayLayer4 = True
            ReDim L4(1 To .NumeroLayers(4))
            Get #fh, , L4

            For i = 1 To .NumeroLayers(4)
                MapData(L4(i).x, L4(i).y).Graphic(4).GrhIndex = L4(i).GrhIndex
                InitGrh MapData(L4(i).x, L4(i).y).Graphic(4), MapData(L4(i).x, L4(i).y).Graphic(4).GrhIndex
            Next i

        End If
    
        If .NumeroTriggers > 0 Then
            ReDim Triggers(1 To .NumeroTriggers)
            Get #fh, , Triggers
            
            For i = 1 To .NumeroTriggers
                MapData(Triggers(i).x, Triggers(i).y).Trigger = Triggers(i).Trigger
                
                ' Transparencia de techos
                If HayTecho(Triggers(i).x, Triggers(i).y) Then
                    ' Array con todos los distintos tipos de triggers para techo
                    If Triggers(i).Trigger < LBoundRoof Then
                        LBoundRoof = Triggers(i).Trigger
                        ReDim Preserve RoofsLight(LBoundRoof To UBoundRoof)

                    ElseIf Triggers(i).Trigger > UBoundRoof Then
                        UBoundRoof = Triggers(i).Trigger
                        ReDim Preserve RoofsLight(LBoundRoof To UBoundRoof)
                    End If
                    
                    RoofsLight(Triggers(i).Trigger) = 255
                    
                ' Trigger detalles en agua
                ElseIf Triggers(i).Trigger = eTrigger.DETALLEAGUA Or Triggers(i).Trigger = eTrigger.VALIDONADO Or Triggers(i).Trigger = eTrigger.NADOCOMBINADO Or Triggers(i).Trigger = eTrigger.NADOBAJOTECHO Then
                    ' Borro flag de costa
                    MapData(Triggers(i).x, Triggers(i).y).Blocked = MapData(Triggers(i).x, Triggers(i).y).Blocked And Not FLAG_COSTA
                End If
            Next i

        End If
    
        If .NumeroParticulas > 0 Then
            ReDim Particulas(1 To .NumeroParticulas)
            Get #fh, , Particulas

            For i = 1 To .NumeroParticulas
            
                MapData(Particulas(i).x, Particulas(i).y).particle_Index = Particulas(i).Particula
                General_Particle_Create MapData(Particulas(i).x, Particulas(i).y).particle_Index, Particulas(i).x, Particulas(i).y

            Next i

        End If

        If .NumeroLuces > 0 Then
            ReDim Luces(1 To .NumeroLuces)
            Get #fh, , Luces

            For i = 1 To .NumeroLuces
                MapData(Luces(i).x, Luces(i).y).luz.Color = Luces(i).Color
                MapData(Luces(i).x, Luces(i).y).luz.Rango = Luces(i).Rango

                If MapData(Luces(i).x, Luces(i).y).luz.Rango <> 0 Then
                    If MapData(Luces(i).x, Luces(i).y).luz.Rango < 100 Then
                        LucesCuadradas.Light_Create Luces(i).x, Luces(i).y, Luces(i).Color, Luces(i).Rango, Luces(i).x & Luces(i).y
                    Else
                        LucesRedondas.Create_Light_To_Map Luces(i).x, Luces(i).y, Luces(i).Color, Luces(i).Rango - 99
                    End If

                End If
               
            Next i

        End If

        If .NumeroOBJs > 0 Then
            ReDim Objetos(1 To .NumeroOBJs)
            Get #fh, , Objetos

            For i = 1 To .NumeroOBJs
                MapData(Objetos(i).x, Objetos(i).y).OBJInfo.ObjIndex = Objetos(i).ObjIndex
                MapData(Objetos(i).x, Objetos(i).y).OBJInfo.Amount = Objetos(i).ObjAmmount
                MapData(Objetos(i).x, Objetos(i).y).ObjGrh.GrhIndex = ObjData(Objetos(i).ObjIndex).GrhIndex
                Call InitGrh(MapData(Objetos(i).x, Objetos(i).y).ObjGrh, MapData(Objetos(i).x, Objetos(i).y).ObjGrh.GrhIndex)

            Next i

        End If

    End With

    Close fh
    
    '
    'Creo un array de zonas provisorio de ese mapa
    ReDim Temp_zone(1 To 1) As MapZone
    Dim UpperB As Integer
    Dim ZonasEnMapa As Integer: ZonasEnMapa = 0
    For i = 1 To UBound(Zonas)
        'Me fijo se la zona pertenece al mapa, de serlo agrego la zona al array
        If Zonas(i).NumMapa = map Then
            ZonasEnMapa = ZonasEnMapa + 1
            ReDim Preserve Temp_zone(1 To ZonasEnMapa) As MapZone
            
            
            Temp_zone(ZonasEnMapa).NumMapa = Zonas(i).NumMapa
            Temp_zone(ZonasEnMapa).Musica = Zonas(i).Musica
            Temp_zone(ZonasEnMapa).OcultarNombre = Zonas(i).OcultarNombre
            Temp_zone(ZonasEnMapa).x1 = Zonas(i).x1
            Temp_zone(ZonasEnMapa).x2 = Zonas(i).x2
            Temp_zone(ZonasEnMapa).y1 = Zonas(i).y1
            Temp_zone(ZonasEnMapa).y2 = Zonas(i).y2
        End If
    Next i
        
    If ZonasEnMapa > 0 Then
        For i = 1 To (UBound(Temp_zone))
            For x = Temp_zone(i).x1 To Temp_zone(i).x2
                For y = Temp_zone(i).y1 To Temp_zone(i).y2
                    MapData(x, y).zone.OcultarNombre = Temp_zone(i).OcultarNombre
                    MapData(x, y).zone.Musica = Temp_zone(i).Musica
                Next y
            Next x
        Next i
    End If
    
    #If Compresion = 1 Then
        Delete_File Windows_Temp_Dir & "mapa" & map & ".csm"
    #End If

    Exit Sub

CargarMapa_Err:
    Call RegistrarError(Err.Number, Err.Description, "Recursos.CargarMapa", Erl)
    Resume Next
    
End Sub

Public Sub CargarParticulas()
    
    On Error GoTo CargarParticulas_Err
    

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

        If Not Extract_File(Scripts, App.path & "\..\Recursos\OUTPUT\", "particles.ini", Windows_Temp_Dir, ResourcesPassword, False) Then
            Err.Description = "¡No se puede cargar el archivo de particles.ini!"
            MsgBox Err.Description

        End If

        StreamFile = Windows_Temp_Dir & "particles.ini"
    #Else
        StreamFile = App.path & "\..\Recursos\init\particles.ini"
    #End If

    ParticulasTotales = Val(General_Var_Get(StreamFile, "INIT", "Total"))
    
    'resize StreamData array
    ReDim StreamData(1 To ParticulasTotales) As Stream
    
    'fill StreamData array with info from Particles.ini
    For loopc = 1 To ParticulasTotales
        StreamData(loopc).name = General_Var_Get(StreamFile, Val(loopc), "Name")
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
            StreamData(loopc).colortint(ColorSet - 1).r = General_Field_Read(1, TempSet, ",")
            StreamData(loopc).colortint(ColorSet - 1).G = General_Field_Read(2, TempSet, ",")
            StreamData(loopc).colortint(ColorSet - 1).B = General_Field_Read(3, TempSet, ",")
        Next ColorSet
        
    Next loopc
        
    #If Compresion = 1 Then
        Delete_File Windows_Temp_Dir & "particles.ini"
    #End If

    
    Exit Sub

CargarParticulas_Err:
    Call RegistrarError(Err.Number, Err.Description, "Recursos.CargarParticulas", Erl)
    Resume Next
    
End Sub

Public Sub CargarParticulasBinary()
    
    On Error GoTo CargarParticulasBinary_Err
    

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

    Dim Handle     As Integer

    'Open files
    Handle = FreeFile()

    #If Compresion = 1 Then

        If Not Extract_File(Scripts, App.path & "\..\Recursos\OUTPUT\", "particles.ind", Windows_Temp_Dir, ResourcesPassword, False) Then
            Err.Description = "¡No se puede cargar el archivo de particles.ind!"
            MsgBox Err.Description

        End If

        StreamFile = Windows_Temp_Dir & "particles.ind"
    #Else
        StreamFile = App.path & "\..\Recursos\init\particles.ind"
    #End If

    Dim n As Integer
    
    n = FreeFile()

    Open StreamFile For Binary Access Read As #n
    'num de cabezas
    Get #n, , ParticulasTotales

    ReDim StreamData(1 To ParticulasTotales) As Stream

    'fill StreamData array with info from Particles.ini
    For loopc = 1 To ParticulasTotales
        Get #n, , StreamData(loopc)
    Next loopc
    
    Close #n

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
            StreamData(loopc).colortint(ColorSet - 1).r = General_Field_Read(1, TempSet, ",")
            StreamData(loopc).colortint(ColorSet - 1).G = General_Field_Read(2, TempSet, ",")
            StreamData(loopc).colortint(ColorSet - 1).B = General_Field_Read(3, TempSet, ",")
        Next ColorSet
        
    Next loopc
        
    #If Compresion = 1 Then
        Delete_File Windows_Temp_Dir & "particles.ini"
    #End If

    
    Exit Sub

CargarParticulasBinary_Err:
    'Call RegistrarError(Err.Number, Err.Description, "Recursos.CargarParticulasBinary", Erl)
    Resume Next
    
End Sub

Public Sub LoadProjectiles()
    Dim ProjectN As Integer
    #If Compresion = 1 Then
        If Not Extract_File(Scripts, App.path & "\..\Recursos\OUTPUT\", "ProjectileDef.dat", Windows_Temp_Dir, ResourcesPassword, False) Then
            Err.Description = "¡No se puede cargar el archivo de ProjectileDef.dat!"
            MsgBox Err.Description

        End If
        ObjFile = Windows_Temp_Dir & "ProjectileDef.dat"
    #Else
        ObjFile = App.path & "\..\Recursos\init\ProjectileDef.dat"
    #End If
    Dim IniReader As New clsIniManager
    Debug.Assert FileExist(ObjFile, vbNormal)
    Call IniReader.Initialize(ObjFile)

    ProjectN = Val(IniReader.GetValue("INIT", "NumProjectile"))
    ReDim ProjectileData(1 To ProjectN) As t_Projectile
    Dim Prj As Integer
    For Prj = 1 To ProjectN
        ProjectileData(Prj).grh = Val(IniReader.GetValue("PROJECTILE" & Prj, "GRH"))
        ProjectileData(Prj).RigthGrh = Val(IniReader.GetValue("PROJECTILE" & Prj, "GRHRigth"))
        ProjectileData(Prj).speed = Val(IniReader.GetValue("PROJECTILE" & Prj, "SPEED")) / 1000
        ProjectileData(Prj).OffsetRotation = Val(IniReader.GetValue("PROJECTILE" & Prj, "OFFSETROTATION"))
        ProjectileData(Prj).RotationSpeed = Val(IniReader.GetValue("PROJECTILE" & Prj, "ROTATIONSPEED"))
    Next Prj
End Sub

Public Sub LoadBuffResources()
    Dim EffectCount As Integer
    #If Compresion = 1 Then
        If Not Extract_File(Scripts, App.path & "\..\Recursos\OUTPUT\", "Effects.ini", Windows_Temp_Dir, ResourcesPassword, False) Then
            Err.Description = "¡No se puede cargar el archivo de ProjectileDef.dat!"
            MsgBox Err.Description

        End If
        ObjFile = Windows_Temp_Dir & "Effects.ini"
    #Else
        ObjFile = App.path & "\..\Recursos\init\Effects.ini"
    #End If
    Dim IniReader As New clsIniManager
    Debug.Assert FileExist(ObjFile, vbNormal)
    Call IniReader.Initialize(ObjFile)

    EffectCount = Val(IniReader.GetValue("INIT", "EffectCount"))
    ReDim EffectResources(1 To EffectCount) As e_effectResource
    Dim Prj As Integer
    For Prj = 1 To EffectCount
        EffectResources(Prj).GrhId = Val(IniReader.GetValue("Effect" & Prj, "GRH"))
    Next Prj
End Sub

Public Function GetPatchNotes() As String
On Error GoTo GetPatchNotes_Err
    Dim PatchDate As Long
    Dim LastDisplayPatch As Long
    Dim PatchFile As String
    #If Compresion = 1 Then
        If Not Extract_File(Scripts, App.path & "\..\Recursos\OUTPUT\", "PatchNotes.dat", Windows_Temp_Dir, ResourcesPassword, False) Then
            GetPatchNotes = ""
            Exit Function
        End If
        PatchFile = Windows_Temp_Dir & "PatchNotes.dat"
    #Else
        PatchFile = App.path & "\..\Recursos\init\PatchNotes.dat"
    #End If
    Dim IniReader As New clsIniManager
    If Not FileExist(PatchFile, vbNormal) Then
        GetPatchNotes = ""
        Exit Function
    End If

    Call IniReader.Initialize(PatchFile)
    
    PatchDate = Val(IniReader.GetValue("INIT", "Date"))
    If PatchDate = 0 Then
        GetPatchNotes = ""
        Exit Function
    End If
    LastDisplayPatch = Val(GetSetting("OPCIONES", "LastPatch"))
    
    If PatchDate > LastDisplayPatch Then
        GetPatchNotes = IniReader.GetValue("INIT", "FileName")
        Call SaveSetting("OPCIONES", "LastPatch", PatchDate)
    Else
        GetPatchNotes = ""
    End If
    Exit Function
GetPatchNotes_Err:
    Call RegistrarError(Err.Number, Err.Description, "Recursos.GetPatchNotes", Erl)
    GetPatchNotes = ""
End Function

Public Sub CargarIndicesOBJ()
    
    On Error GoTo CargarIndicesOBJ_Err
    

    Dim Obj     As Integer

    Dim Npc     As Integer

    Dim Hechizo As Integer

    Dim i       As Integer
    
    #If Compresion = 1 Then
        If Not Extract_File(Scripts, App.path & "\..\Recursos\OUTPUT\", "localindex.dat", Windows_Temp_Dir, ResourcesPassword, False) Then
            Err.Description = "¡No se puede cargar el archivo de localindex.dat!"
            MsgBox Err.Description

        End If
        ObjFile = Windows_Temp_Dir & "localindex.dat"
    #Else
        ObjFile = App.path & "\..\Recursos\init\localindex.dat"
    #End If
    
            
    Dim Leer As New clsIniManager
    Debug.Assert FileExist(ObjFile, vbNormal)
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
    ReDim ObjShop(1 To 1) As ObjDatas
    
    Debug.Assert NumQuest > 0
    Debug.Assert NumSug > 0
    
    ReDim QuestList(1 To NumQuest)
    ReDim PosMap(1 To NumQuest) As Integer
    ReDim Sugerencia(1 To NumSug) As String

    For Obj = 1 To NumOBJs
        DoEvents
        ObjData(Obj).GrhIndex = Val(Leer.GetValue("OBJ" & Obj, "grhindex"))
        If Obj = 403 Then
            Debug.Print "asd"
        End If
        
        Select Case language
            Case e_language.English
                ObjData(Obj).name = IIf(Leer.GetValue("OBJ" & Obj, "en_Name") <> vbNullString, Leer.GetValue("OBJ" & Obj, "en_Name"), Leer.GetValue("OBJ" & Obj, "Name"))
                ObjData(Obj).info = IIf(Leer.GetValue("OBJ" & Obj, "en_Info") <> vbNullString, Leer.GetValue("OBJ" & Obj, "en_Info"), Leer.GetValue("OBJ" & Obj, "Info"))
                ObjData(Obj).Texto = IIf(Leer.GetValue("OBJ" & Obj, "en_Texto") <> vbNullString, Leer.GetValue("OBJ" & Obj, "en_Texto"), Leer.GetValue("OBJ" & Obj, "Texto"))
            Case e_language.Spanish
                ObjData(Obj).name = Leer.GetValue("OBJ" & Obj, "Name")
                ObjData(Obj).info = Leer.GetValue("OBJ" & Obj, "Info")
                ObjData(Obj).Texto = Leer.GetValue("OBJ" & Obj, "Texto")
        End Select
        
        ObjData(Obj).MinDef = Val(Leer.GetValue("OBJ" & Obj, "MinDef"))
        ObjData(Obj).MaxDef = Val(Leer.GetValue("OBJ" & Obj, "MaxDef"))
        ObjData(Obj).MinHit = Val(Leer.GetValue("OBJ" & Obj, "MinHit"))
        ObjData(Obj).MaxHit = Val(Leer.GetValue("OBJ" & Obj, "MaxHit"))
        ObjData(Obj).ObjType = Val(Leer.GetValue("OBJ" & Obj, "ObjType"))
        ObjData(Obj).Cooldown = Val(Leer.GetValue("OBJ" & Obj, "CD"))
        ObjData(Obj).cdType = Val(Leer.GetValue("OBJ" & Obj, "CDType"))
        ObjData(Obj).CreaGRH = Leer.GetValue("OBJ" & Obj, "CreaGRH")
        ObjData(Obj).CreaLuz = Leer.GetValue("OBJ" & Obj, "CreaLuz")
        ObjData(Obj).CreaParticulaPiso = Val(Leer.GetValue("OBJ" & Obj, "CreaParticulaPiso"))
        ObjData(Obj).proyectil = Val(Leer.GetValue("OBJ" & Obj, "proyectil"))
        ObjData(Obj).Amunition = Val(Leer.GetValue("OBJ" & Obj, "MUNICIONES"))
        ObjData(Obj).Hechizo = Val(Leer.GetValue("OBJ" & Obj, "Hechizo"))
        ObjData(Obj).Raices = Val(Leer.GetValue("OBJ" & Obj, "Raices"))
        ObjData(Obj).Cuchara = Val(Leer.GetValue("OBJ" & Obj, "Cuchara"))
        ObjData(Obj).Botella = Val(Leer.GetValue("OBJ" & Obj, "Botella"))
        ObjData(Obj).Mortero = Val(Leer.GetValue("OBJ" & Obj, "Mortero"))
        ObjData(Obj).FrascoAlq = Val(Leer.GetValue("OBJ" & Obj, "FrascoAlq"))
        ObjData(Obj).FrascoElixir = Val(Leer.GetValue("OBJ" & Obj, "FrascoElixir"))
        ObjData(Obj).Dosificador = Val(Leer.GetValue("OBJ" & Obj, "Dosificador"))
        ObjData(Obj).Orquidea = Val(Leer.GetValue("OBJ" & Obj, "Orquidea"))
        ObjData(Obj).Carmesi = Val(Leer.GetValue("OBJ" & Obj, "Carmesi"))
        ObjData(Obj).HongoDeLuz = Val(Leer.GetValue("OBJ" & Obj, "HongoDeLuz"))
        ObjData(Obj).Esporas = Val(Leer.GetValue("OBJ" & Obj, "Esporas"))
        ObjData(Obj).Tuna = Val(Leer.GetValue("OBJ" & Obj, "Tuna"))
        ObjData(Obj).Cala = Val(Leer.GetValue("OBJ" & Obj, "Cala"))
        ObjData(Obj).ColaDeZorro = Val(Leer.GetValue("OBJ" & Obj, "ColaDeZorro"))
        ObjData(Obj).FlorOceano = Val(Leer.GetValue("OBJ" & Obj, "FlorOceano"))
        ObjData(Obj).FlorRoja = Val(Leer.GetValue("OBJ" & Obj, "FlorRoja"))
        ObjData(Obj).Hierva = Val(Leer.GetValue("OBJ" & Obj, "Hierva"))
        ObjData(Obj).HojasDeRin = Val(Leer.GetValue("OBJ" & Obj, "HojasDeRin"))
        ObjData(Obj).HojasRojas = Val(Leer.GetValue("OBJ" & Obj, "HojasRojas"))
        ObjData(Obj).SemillasPros = Val(Leer.GetValue("OBJ" & Obj, "SemillasPros"))
        ObjData(Obj).Pimiento = Val(Leer.GetValue("OBJ" & Obj, "Pimiento"))
        ObjData(Obj).Madera = Val(Leer.GetValue("OBJ" & Obj, "Madera"))
        ObjData(Obj).MaderaElfica = Val(Leer.GetValue("OBJ" & Obj, "MaderaElfica"))
        ObjData(Obj).PielLobo = Val(Leer.GetValue("OBJ" & Obj, "PielLobo"))
        ObjData(Obj).PielOsoPardo = Val(Leer.GetValue("OBJ" & Obj, "PielOsoPardo"))
        ObjData(Obj).PielOsoPolar = Val(Leer.GetValue("OBJ" & Obj, "PielOsoPolar"))
        ObjData(Obj).PielLoboNegro = Val(Leer.GetValue("OBJ" & Obj, "PielLoboNegro"))
        ObjData(Obj).PielTigre = Val(Leer.GetValue("OBJ" & Obj, "PielTigre"))
        ObjData(Obj).PielTigreBengala = Val(Leer.GetValue("OBJ" & Obj, "PielTigreBengala"))
        ObjData(Obj).LingH = Val(Leer.GetValue("OBJ" & Obj, "LingH"))
        ObjData(Obj).LingP = Val(Leer.GetValue("OBJ" & Obj, "LingP"))
        ObjData(Obj).LingO = Val(Leer.GetValue("OBJ" & Obj, "LingO"))
        ObjData(Obj).Coal = Val(Leer.GetValue("OBJ" & Obj, "Coal"))
        ObjData(Obj).Destruye = Val(Leer.GetValue("OBJ" & Obj, "Destruye"))
        ObjData(Obj).SkHerreria = Val(Leer.GetValue("OBJ" & Obj, "SkHerreria"))
        ObjData(Obj).SkPociones = Val(Leer.GetValue("OBJ" & Obj, "SkPociones"))
        ObjData(Obj).Sksastreria = Val(Leer.GetValue("OBJ" & Obj, "Sksastreria"))
        ObjData(Obj).Valor = Val(Leer.GetValue("OBJ" & Obj, "Valor"))
        ObjData(Obj).Agarrable = Val(Leer.GetValue("OBJ" & Obj, "Agarrable"))
        ObjData(Obj).Llave = Val(Leer.GetValue("OBJ" & Obj, "Llave"))
            
        If Val(Leer.GetValue("OBJ" & Obj, "NFT")) = 1 Then
            ObjShop(i).name = Leer.GetValue("OBJ" & Obj, "Name")
            ObjShop(i).Valor = Val(Leer.GetValue("OBJ" & Obj, "Valor"))
            ObjShop(i).objNum = Obj
            ReDim Preserve ObjShop(1 To (UBound(ObjShop) + 1)) As ObjDatas
        End If
        
    Next Obj
    
    Dim aux   As String

    Dim loopc As Byte
    
    For Npc = 1 To NumNpcs
        DoEvents

        If (CBool(Val(Leer.GetValue("npc" & Npc, "NoMapInfo")))) Then
            GoTo Continue
        End If
        
        Select Case language
            Case e_language.English
                NpcData(Npc).name = IIf(Leer.GetValue("npc" & Npc, "en_Name") <> vbNullString, Leer.GetValue("npc" & Npc, "en_Name"), Leer.GetValue("npc" & Npc, "Name"))
                NpcData(Npc).desc = IIf(Leer.GetValue("npc" & Npc, "en_desc") <> vbNullString, Leer.GetValue("npc" & Npc, "en_desc"), Leer.GetValue("npc" & Npc, "desc"))
            Case e_language.Spanish
                NpcData(Npc).name = Leer.GetValue("npc" & Npc, "Name")
                NpcData(Npc).desc = Leer.GetValue("npc" & Npc, "desc")
        End Select
        

        If NpcData(Npc).name = "" Then
            NpcData(Npc).Name = "Vacío"

        End If

        NpcData(Npc).Body = Val(Leer.GetValue("npc" & Npc, "Body"))
        NpcData(Npc).exp = Val(Leer.GetValue("npc" & Npc, "exp"))
        NpcData(Npc).Head = Val(Leer.GetValue("npc" & Npc, "Head"))
        NpcData(Npc).Hp = Val(Leer.GetValue("npc" & Npc, "Hp"))
        NpcData(Npc).MaxHit = Val(Leer.GetValue("npc" & Npc, "MaxHit"))
        NpcData(Npc).MinHit = Val(Leer.GetValue("npc" & Npc, "MinHit"))
        NpcData(Npc).oro = Val(Leer.GetValue("npc" & Npc, "oro"))
        
        NpcData(Npc).ExpClan = Val(Leer.GetValue("npc" & Npc, "GiveEXPClan"))
        
        NpcData(Npc).PuedeInvocar = Val(Leer.GetValue("npc" & Npc, "PuedeInvocar"))
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
    Continue:
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
        NameMaps(Hechizo).name = Leer.GetValue("NameMapa", "Mapa" & Hechizo)
        NameMaps(Hechizo).desc = Leer.GetValue("NameMapa", "Mapa" & Hechizo & "Desc")
    Next Hechizo
    
    For Hechizo = 1 To NumQuest
        DoEvents
        
        QuestList(Hechizo).nombre = Leer.GetValue("QUEST" & Hechizo, "NOMBRE")
        
        QuestList(Hechizo).desc = Leer.GetValue("QUEST" & Hechizo, "DESC")
        QuestList(Hechizo).NextQuest = Leer.GetValue("QUEST" & Hechizo, "NEXTQUEST")
        QuestList(Hechizo).DescFinal = Leer.GetValue("QUEST" & Hechizo, "DESCFINAL")
        QuestList(Hechizo).RequiredLevel = Leer.GetValue("QUEST" & Hechizo, "RequiredLevel")
        QuestList(Hechizo).Repetible = Val(Leer.GetValue("QUEST" & Hechizo, "Repetible"))
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

    
    Exit Sub

CargarIndicesOBJ_Err:
    Call RegistrarError(Err.Number, Err.Description, "Recursos.CargarIndicesOBJ", Erl)
    Resume Next
    
End Sub

Public Sub Cargarmapsworlddata()
    
    On Error GoTo Cargarmapsworlddata_Err
    

    'Ladder
    Dim MapFile As String

    Dim i       As Integer
    Dim j       As Byte

    #If Compresion = 1 Then

        If Not Extract_File(Scripts, App.path & "\..\Recursos\OUTPUT\", "mapsworlddata.dat", Windows_Temp_Dir, ResourcesPassword, False) Then
            Err.Description = "¡No se puede cargar el archivo de mapsworlddata.dat!"
            MsgBox Err.Description

        End If
    
        MapFile = Windows_Temp_Dir & "mapsworlddata.dat"
    #Else
        MapFile = App.path & "\..\Recursos\init\mapsworlddata.dat"
    #End If

    Dim Leer As New clsIniManager
    Call Leer.Initialize(MapFile)

    
    TotalWorlds = Val(Leer.GetValue("INIT", "TotalWorlds"))
       
    ReDim Mundo(1 To TotalWorlds) As WorldMap
   
    For j = 1 To TotalWorlds
        Mundo(j).Alto = Val(Leer.GetValue("WORLDMAP" & j, "Alto"))
        Mundo(j).Ancho = Val(Leer.GetValue("WORLDMAP" & j, "Ancho"))

        ReDim Mundo(j).MapIndice(1 To Mundo(j).Alto * Mundo(j).Ancho) As Integer
         
         For i = 1 To Mundo(j).Alto * Mundo(j).Ancho
             Mundo(j).MapIndice(i) = Val(Leer.GetValue("WORLDMAP" & j, i))
         Next i
         
     Next j
    
    

    #If Compresion = 1 Then
        Delete_File Windows_Temp_Dir & "mapsworlddata.dat"
    #End If

    
    Exit Sub

Cargarmapsworlddata_Err:
    Call RegistrarError(Err.Number, Err.Description, "Recursos.Cargarmapsworlddata", Erl)
    Resume Next
    
End Sub

Sub CargarMoldes()

    BodiesHeading(1) = E_Heading.south
    BodiesHeading(2) = E_Heading.NORTH
    BodiesHeading(3) = E_Heading.WEST
    BodiesHeading(4) = E_Heading.EAST
    
    Dim Loader As clsIniManager
    Set Loader = New clsIniManager
    
    #If Compresion = 1 Then

        If Not Extract_File(Scripts, App.path & "\..\Recursos\OUTPUT\", "moldes.ini", Windows_Temp_Dir, ResourcesPassword, False) Then
            Err.Description = "¡No se puede cargar el archivo de moldes.ini!"
            MsgBox Err.Description

        End If

        Call Loader.Initialize(Windows_Temp_Dir & "moldes.ini")
    #Else
        Call Loader.Initialize(App.path & "\..\Recursos\init\moldes.ini")
    #End If
    
    Dim NumMoldes As Integer
    NumMoldes = Val(Loader.GetValue("INIT", "Moldes"))

    ReDim MoldesBodies(1 To NumMoldes)
    
    Dim i As Integer, MoldeKey As String
    
    For i = 1 To NumMoldes
        MoldeKey = "Molde" & i
    
        With MoldesBodies(i)
            .x = Val(Loader.GetValue(MoldeKey, "X"))
            .y = Val(Loader.GetValue(MoldeKey, "Y"))
            .Width = Val(Loader.GetValue(MoldeKey, "Width"))
            .Height = Val(Loader.GetValue(MoldeKey, "Height"))
            .DirCount(1) = Val(Loader.GetValue(MoldeKey, "Dir1"))
            .DirCount(2) = Val(Loader.GetValue(MoldeKey, "Dir2"))
            .DirCount(3) = Val(Loader.GetValue(MoldeKey, "Dir3"))
            .DirCount(4) = Val(Loader.GetValue(MoldeKey, "Dir4"))
            .TotalGrhs = .DirCount(1) + .DirCount(2) + .DirCount(3) + .DirCount(4) + 4
        End With
    Next
    
    Set Loader = Nothing
    
    #If Compresion = 1 Then
        Delete_File Windows_Temp_Dir & "moldes.ini"
    #End If

End Sub
Sub CargarZonas()
    Dim Reader As clsIniManager
    Dim cantidadZonas As Integer
    Dim i As Integer
    Set Reader = New clsIniManager
    
    Call Reader.Initialize(App.path & "\..\Recursos\Dat\zonas.dat")
    
    cantidadZonas = Reader.GetValue("Config", "Cantidad")
    
    ReDim Zonas(1 To cantidadZonas) As MapZone
    
    For i = 1 To cantidadZonas
        Zonas(i).Musica = Val(Reader.GetValue("Zona" & i, "Musica"))
        Zonas(i).OcultarNombre = Val(Reader.GetValue("Zona" & i, "OcultarNombre"))
        Zonas(i).NumMapa = Val(Reader.GetValue("Zona" & i, "Mapa"))
        Zonas(i).x1 = Val(Reader.GetValue("Zona" & i, "X1"))
        Zonas(i).x2 = Val(Reader.GetValue("Zona" & i, "X2"))
        Zonas(i).y1 = Val(Reader.GetValue("Zona" & i, "Y1"))
        Zonas(i).y2 = Val(Reader.GetValue("Zona" & i, "Y2"))
    Next i
    
    Set Reader = Nothing
End Sub

Sub CargarCabezas()
    
    On Error GoTo CargarCabezas_Err
    

    Dim n            As Integer

    Dim i            As Long

    Dim Numheads     As Integer

    Dim Miscabezas() As tIndiceCabeza
    
    n = FreeFile()
    
    #If Compresion = 1 Then

        If Not Extract_File(Scripts, App.path & "\..\Recursos\OUTPUT\", "cabezas.ind", Windows_Temp_Dir, ResourcesPassword, False) Then
            Err.Description = "¡No se puede cargar el archivo de Cabezas.ind!"
            MsgBox Err.Description

        End If

        Open Windows_Temp_Dir & "cabezas.ind" For Binary Access Read As #n
    #Else
        Open App.path & "\..\Recursos\init\cabezas.ind" For Binary Access Read As #n
    #End If

    
    'cabecera
    Get #n, , MiCabecera
    
    'num de cabezas
    Get #n, , Numheads
    
    'Resize array
    ReDim HeadData(0 To Numheads) As HeadData
    ReDim Miscabezas(0 To Numheads) As tIndiceCabeza
    
    For i = 1 To Numheads
        Get #n, , Miscabezas(i)
        
        If Miscabezas(i).Head(1) Then
            Call InitGrh(HeadData(i).Head(1), Miscabezas(i).Head(1), 0)
            Call InitGrh(HeadData(i).Head(2), Miscabezas(i).Head(2), 0)
            Call InitGrh(HeadData(i).Head(3), Miscabezas(i).Head(3), 0)
            Call InitGrh(HeadData(i).Head(4), Miscabezas(i).Head(4), 0)

        End If

    Next i
    
    Close #n
    
    #If Compresion = 1 Then
        Delete_File Windows_Temp_Dir & "cabezas.ind"
    #End If
    
    
    Exit Sub

CargarCabezas_Err:
    Call RegistrarError(Err.Number, Err.Description, "Recursos.CargarCabezas", Erl)
    Resume Next
    
End Sub

Sub CargarCascos()
    
    On Error GoTo CargarCascos_Err
    

    Dim n            As Integer

    Dim i            As Long

    Dim NumCascos    As Integer

    Dim Miscabezas() As tIndiceCabeza
    
    n = FreeFile()
  
    #If Compresion = 1 Then

        If Not Extract_File(Scripts, App.path & "\..\Recursos\OUTPUT\", "cascos.ind", Windows_Temp_Dir, ResourcesPassword, False) Then
            Err.Description = "¡No se puede cargar el archivo de Cabezas.ind!"
            MsgBox Err.Description

        End If

        Open Windows_Temp_Dir & "cascos.ind" For Binary Access Read As #n
    #Else
        Open App.path & "\..\Recursos\init\cascos.ind" For Binary Access Read As #n
    #End If
        
       
    
    'cabecera
    Get #n, , MiCabecera
      
    'num de cabezas
    Get #n, , NumCascos
    
    'Resize array
    ReDim CascoAnimData(0 To NumCascos) As HeadData
    ReDim Miscabezas(0 To NumCascos) As tIndiceCabeza
    
    For i = 1 To NumCascos
        Get #n, , Miscabezas(i)
        
        If Miscabezas(i).Head(1) Then
            Call InitGrh(CascoAnimData(i).Head(1), Miscabezas(i).Head(1), 0)
            Call InitGrh(CascoAnimData(i).Head(2), Miscabezas(i).Head(2), 0)
            Call InitGrh(CascoAnimData(i).Head(3), Miscabezas(i).Head(3), 0)
            Call InitGrh(CascoAnimData(i).Head(4), Miscabezas(i).Head(4), 0)

        End If

    Next i
    
    Close #n
    
    #If Compresion = 1 Then
        Delete_File Windows_Temp_Dir & "cascos.ind"
    #End If

    
    Exit Sub

CargarCascos_Err:
    Call RegistrarError(Err.Number, Err.Description, "Recursos.CargarCascos", Erl)
    Resume Next
    
End Sub


Sub CargarCuerpos()
    
    On Error GoTo CargarCuerpos_Err
    
    Dim Loader       As clsIniManager

    Dim i            As Long
    
    Dim j            As Byte
    
    Dim k            As Integer
    
    Dim Heading      As Byte
    
    Dim BodyKey      As String
    
    Dim Std          As Byte

    Dim NumCuerpos   As Integer
    
    Dim LastGrh      As Long
    
    Dim AnimStart    As Long
    
    Dim x            As Long
    
    Dim y            As Long
    
    Dim FileNum      As Long
    
    Dim AnimSpeed    As Single
    
    Set Loader = New clsIniManager
    
    #If Compresion = 1 Then

        If Not Extract_File(Scripts, App.path & "\..\Recursos\OUTPUT\", "cuerpos.dat", Windows_Temp_Dir, ResourcesPassword, False) Then
            Err.Description = "¡No se puede cargar el archivo de cuerpos.dat!"
            MsgBox Err.Description

        End If

        Call Loader.Initialize(Windows_Temp_Dir & "cuerpos.dat")
    #Else
        Call Loader.Initialize(App.path & "\..\Recursos\init\cuerpos.dat")
    #End If
    
    NumCuerpos = Val(Loader.GetValue("INIT", "NumBodies"))
    
    'Resize array
    ReDim Preserve BodyData(0 To NumCuerpos)

    For i = 1 To NumCuerpos
        BodyKey = "BODY" & i
    
        Std = Val(Loader.GetValue(BodyKey, "Std"))
        
        With BodyData(i)
            .BodyOffset.x = Val(Loader.GetValue(BodyKey, "BodyOffsetX"))
            .BodyOffset.y = Val(Loader.GetValue(BodyKey, "BodyOffsetY"))
            .HeadOffset.x = Val(Loader.GetValue(BodyKey, "HeadOffsetX")) + .BodyOffset.x
            .HeadOffset.y = Val(Loader.GetValue(BodyKey, "HeadOffsetY")) + .BodyOffset.y
            .BodyIndex = i
            .IdleBody = Val(Loader.GetValue(BodyKey, "IdleBody"))
            .AnimateOnIdle = Val(Loader.GetValue(BodyKey, "AnimateOnIdle"))
        End With

        If Std = 0 Then
            InitGrh BodyData(i).Walk(1), Val(Loader.GetValue(BodyKey, "Walk1")), 0
            InitGrh BodyData(i).Walk(2), Val(Loader.GetValue(BodyKey, "Walk2")), 0
            InitGrh BodyData(i).Walk(3), Val(Loader.GetValue(BodyKey, "Walk3")), 0
            InitGrh BodyData(i).Walk(4), Val(Loader.GetValue(BodyKey, "Walk4")), 0
            
        Else
            FileNum = Val(Loader.GetValue(BodyKey, "FileNum"))
            
            AnimSpeed = Val(Loader.GetValue(BodyKey, "Speed"))
            
            If AnimSpeed = 0 Then
                AnimSpeed = 1
            End If

            AnimSpeed = 1 / AnimSpeed / 0.018

            LastGrh = UBound(GrhData)

            ' Agrego espacio para meter el body en GrhData
            ReDim Preserve GrhData(1 To LastGrh + MoldesBodies(Std).TotalGrhs)
            
            MaxGrh = UBound(GrhData)
            
            LastGrh = LastGrh + 1
            x = MoldesBodies(Std).x
            y = MoldesBodies(Std).y
            
            For j = 1 To 4
                AnimStart = LastGrh
            
                For k = 1 To MoldesBodies(Std).DirCount(j)
                    With GrhData(LastGrh)
                        .FileNum = FileNum
                        .NumFrames = 1
                        .sX = x
                        .sY = y
                        .pixelWidth = MoldesBodies(Std).Width
                        .pixelHeight = MoldesBodies(Std).Height
                        
                        .TileWidth = .pixelWidth / TilePixelHeight
                        .TileHeight = .pixelHeight / TilePixelWidth
        
                        ReDim .Frames(1)
                        .Frames(1) = LastGrh
                    End With
                    
                    LastGrh = LastGrh + 1
                    x = x + MoldesBodies(Std).Width
                Next
                
                x = MoldesBodies(Std).x
                y = y + MoldesBodies(Std).Height
                
                Heading = BodiesHeading(j)
                
                With GrhData(LastGrh)
                    .NumFrames = MoldesBodies(Std).DirCount(j)
                    .speed = .NumFrames * AnimSpeed
                    
                    ReDim .Frames(1 To MoldesBodies(Std).DirCount(j))
                    
                    For k = 1 To MoldesBodies(Std).DirCount(j)
                        .Frames(k) = AnimStart + k - 1
                    Next
                    
                    .pixelWidth = GrhData(.Frames(1)).pixelWidth
                    .pixelHeight = GrhData(.Frames(1)).pixelHeight
                    .TileWidth = GrhData(.Frames(1)).TileWidth
                    .TileHeight = GrhData(.Frames(1)).TileHeight
                End With
                
                InitGrh BodyData(i).Walk(Heading), LastGrh, 0
                
                LastGrh = LastGrh + 1
            Next

        End If

    Next i

    #If Compresion = 1 Then
        Delete_File Windows_Temp_Dir & "cuerpos.dat"
    #End If

    Set Loader = Nothing
    
    Exit Sub

CargarCuerpos_Err:
    Call RegistrarError(Err.Number, Err.Description, "Recursos.CargarCuerpos", Erl)
    Resume Next
    
End Sub

Sub CargarFxs()
    
    On Error GoTo CargarFxs_Err
    

    Dim n      As Integer

    Dim i      As Long

    Dim NumFxs As Integer
    
    n = FreeFile()

    #If Compresion = 1 Then

        If Not Extract_File(Scripts, App.path & "\..\Recursos\OUTPUT\", "fxs.ind", Windows_Temp_Dir, ResourcesPassword, False) Then
            Err.Description = "¡No se puede cargar el archivo de fxs.ind!"
            MsgBox Err.Description

        End If

        Open Windows_Temp_Dir & "fxs.ind" For Binary Access Read As #n
    #Else
        Open App.path & "\..\Recursos\init\fxs.ind" For Binary Access Read As #n
    #End If
       
    
    'cabecera
    Get #n, , MiCabecera
       
    'num de cabezas
    Get #n, , NumFxs
    
    'Resize array
    ReDim FxData(1 To NumFxs) As tIndiceFx
    ReDim FxToAnimationMap(1 To NumFxs)
    
    For i = 1 To NumFxs
        Get #n, , FxData(i)
    Next i
    
    Close #n
    
    #If Compresion = 1 Then
        Delete_File Windows_Temp_Dir & "fxs.ind"
    #End If

    
    Exit Sub

CargarFxs_Err:
    Call RegistrarError(Err.Number, Err.Description, "Recursos.CargarFxs", Erl)
    Resume Next
    
End Sub

Public Sub CalculateCliptime(ByRef clip As tAnimationClip)
    clip.ClipTime = GrhData(FxData(clip.fX).Animacion).speed
End Sub

Public Sub CalculateClipsTime(ByRef animData As tComposedAnimation)
    Dim i As Integer
    For i = 1 To UBound(animData.Clips())
        Call CalculateCliptime(animData.Clips(i))
    Next i
End Sub

Public Sub AddComposedMetitation(ByVal index As Long, ByVal startFx As Long, ByVal loopFx As Long)
    ReDim ComposedFxData(index).Clips(3)
    ComposedFxData(index).Clips(1).fX = startFx
    ComposedFxData(index).Clips(1).LoopCount = 0
    ComposedFxData(index).Clips(2).fX = loopFx
    ComposedFxData(index).Clips(2).LoopCount = -1
    ComposedFxData(index).Clips(3).fX = startFx
    ComposedFxData(index).Clips(3).LoopCount = 0
    ComposedFxData(index).Clips(3).Playback = Backward
    Call CalculateClipsTime(ComposedFxData(index))
    ComposedFxData(index).Clips(3).ClipTime = ComposedFxData(index).Clips(3).ClipTime / 2
    FxToAnimationMap(StartFx) = Index
End Sub


Public Sub LoadComposedFx()
    ReDim ComposedFxData(1 To 21) As tComposedAnimation
    
    ReDim ComposedFxData(1).Clips(1)
    ComposedFxData(1).Clips(1).Fx = 115
    ComposedFxData(1).Clips(1).LoopCount = -1
    Call CalculateCliptime(ComposedFxData(1).Clips(1))
    FxToAnimationMap(115) = 1
    
    ReDim ComposedFxData(2).Clips(1)
    ComposedFxData(2).Clips(1).fX = 116
    ComposedFxData(2).Clips(1).LoopCount = -1
    Call CalculateCliptime(ComposedFxData(2).Clips(1))
    FxToAnimationMap(116) = 2
    
    ReDim ComposedFxData(3).Clips(1)
    ComposedFxData(3).Clips(1).Fx = 117
    ComposedFxData(3).Clips(1).LoopCount = -1
    Call CalculateCliptime(ComposedFxData(3).Clips(1))
    FxToAnimationMap(117) = 3
    
    ReDim ComposedFxData(4).Clips(1)
    ComposedFxData(4).Clips(1).Fx = 118
    ComposedFxData(4).Clips(1).LoopCount = -1
    Call CalculateCliptime(ComposedFxData(4).Clips(1))
    FxToAnimationMap(118) = 4
    
    ReDim ComposedFxData(5).Clips(1)
    ComposedFxData(5).Clips(1).Fx = 119
    ComposedFxData(5).Clips(1).LoopCount = -1
    Call CalculateCliptime(ComposedFxData(5).Clips(1))
    FxToAnimationMap(119) = 5
    
    ReDim ComposedFxData(6).Clips(1)
    ComposedFxData(6).Clips(1).Fx = 120
    ComposedFxData(6).Clips(1).LoopCount = -1
    Call CalculateCliptime(ComposedFxData(6).Clips(1))
    FxToAnimationMap(120) = 6
    
    Call AddComposedMetitation(7, 122, 126)
    Call AddComposedMetitation(8, 123, 130)
    Call AddComposedMetitation(9, 124, 134)
    Call AddComposedMetitation(10, 127, 126)
    Call AddComposedMetitation(11, 128, 126)
    Call AddComposedMetitation(12, 129, 126)
    Call AddComposedMetitation(13, 131, 130)
    Call AddComposedMetitation(14, 132, 130)
    Call AddComposedMetitation(15, 133, 130)
    Call AddComposedMetitation(16, 135, 134)
    Call AddComposedMetitation(17, 136, 134)
    Call AddComposedMetitation(18, 137, 134)
    Call AddComposedMetitation(19, 139, 138)
    Call AddComposedMetitation(20, 140, 138)
    Call AddComposedMetitation(21, 141, 138)
    
End Sub

Public Function LoadGrhData() As Boolean

    On Error GoTo ErrorHandler

    Dim grh         As Long
    Dim Frame       As Long
    Dim grhCount    As Long
    Dim Handle      As Integer
    Dim fileVersion As Long
    
    'Open files
    Handle = FreeFile()
    
    #If Compresion = 1 Then

        If Not Extract_File(Scripts, App.path & "\..\Recursos\OUTPUT\", "graficos.ind", Windows_Temp_Dir, ResourcesPassword, False) Then
            Err.Description = "¡No se puede cargar el archivo de recurso!"
            GoTo ErrorHandler

        End If
    
        Open Windows_Temp_Dir & "graficos.ind" For Binary Access Read As #Handle
    #Else
        Open App.path & "\..\Recursos\init\graficos.ind" For Binary Access Read As #Handle
    #End If
    
    'Get file version
    Get #Handle, , fileVersion
    
    'Get number of grhs
    Get #Handle, , grhCount
    
    'Resize arrays
    ReDim GrhData(1 To grhCount) As GrhData
    
    MaxGrh = grhCount

    Dim Fin As Boolean

    Fin = False

    While Not EOF(Handle) And Fin = False

        Get #Handle, , grh

        With GrhData(grh)
        
            GrhData(grh).active = True
            'Get number of frames
            Get #Handle, , .NumFrames

            If .NumFrames <= 0 Then GoTo ErrorHandler
            
            ReDim .Frames(1 To GrhData(grh).NumFrames)
            
            If .NumFrames > 1 Then

                'Read a animation GRH set
                For Frame = 1 To .NumFrames
                    Get #Handle, , .Frames(Frame)

                    If .Frames(Frame) <= 0 Or .Frames(Frame) > grhCount Then
                        GoTo ErrorHandler

                    End If

                Next Frame
                
                Get #Handle, , GrhData(grh).speed
                
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
                Get #Handle, , .FileNum

                If .FileNum <= 0 Then GoTo ErrorHandler
                                
                Get #Handle, , GrhData(grh).sX

                If .sX < 0 Then GoTo ErrorHandler
                
                Get #Handle, , GrhData(grh).sY

                If .sY < 0 Then GoTo ErrorHandler
                
                Get #Handle, , GrhData(grh).pixelWidth

                If .pixelWidth <= 0 Then GoTo ErrorHandler
                
                Get #Handle, , GrhData(grh).pixelHeight

                If .pixelHeight <= 0 Then GoTo ErrorHandler
                
                'Compute width and height
                .TileWidth = .pixelWidth / TilePixelHeight
                .TileHeight = .pixelHeight / TilePixelWidth

                .Frames(1) = grh

            End If

        End With

        If grh = MaxGrh Then Fin = True
    Wend

    Close #Handle
    
    LoadGrhData = True
    
    #If Compresion = 1 Then
        Delete_File Windows_Temp_Dir & "graficos.ind"
    #End If

    Exit Function

ErrorHandler:
    LoadGrhData = False
    MsgBox "Error " & Err.Description & " durante la carga de Grh.dat! La carga se ha detenido en GRH: " & grh
    
End Function

Public Sub LoadGrhIni()
    On Error GoTo hErr

    Dim FileHandle     As Integer
    Dim grh            As Long
    Dim Frame          As Long
    Dim SeparadorClave As String
    Dim SeparadorGrh   As String
    Dim CurrentLine    As String
    Dim Fields()       As String
    
    ' Guardo el separador en una variable asi no lo busco en cada bucle.
    SeparadorClave = "="
    SeparadorGrh = "-"

    ' Abrimos el archivo. No uso FileManager porque obliga a cargar todo el archivo en memoria
    ' y es demasiado grande. En cambio leo linea por linea y procesamos de a una.
    FileHandle = FreeFile()

    #If Compresion = 1 Then

        If Not Extract_File(Scripts, App.path & "\..\Recursos\OUTPUT\", "Graficos.ini", Windows_Temp_Dir, ResourcesPassword, False) Then
            Err.Description = "¡No se puede cargar el archivo de recurso!"
            GoTo hErr
        End If
    
        Open Windows_Temp_Dir & "Graficos.ini" For Input As #FileHandle
    #Else
        Open App.path & "\..\Recursos\init\Graficos.ini" For Input As #FileHandle
    #End If

    ' Leemos el total de Grhs
    Do While Not EOF(FileHandle)
        ' Leemos la linea actual
        Line Input #FileHandle, CurrentLine

        Fields = Split(CurrentLine, SeparadorClave)
            
        ' Buscamos la clave "NumGrh"
        If Fields(0) = "NumGrh" Then
            ' Asignamos el tamano al array de Grhs
            MaxGrh = Val(Fields(1))

            ReDim GrhData(1 To MaxGrh) As GrhData
                
            Exit Do
        End If
    Loop
        
    ' Chequeamos si pudimos leer la cantidad de Grhs
    If UBound(GrhData) <= 0 Then GoTo hErr
        
    ' Buscamos la posicion del primer Grh
    Do While Not EOF(FileHandle)
        ' Leemos la linea actual
        Line Input #FileHandle, CurrentLine
            
        ' Buscamos el nodo "[Graphics]"
        If UCase$(CurrentLine) = "[GRAPHICS]" Then
            ' Ya lo tenemos, salimos
            Exit Do
        End If
    Loop
        
    ' Recorremos todos los Grhs
    Do While Not EOF(FileHandle)
        ' Leemos la linea actual
        Line Input #FileHandle, CurrentLine
            
        ' Ignoramos lineas vacias
        If CurrentLine <> vbNullString Then
            
            ' Divimos por el "="
            Fields = Split(CurrentLine, SeparadorClave)
                
            ' Leemos el numero de Grh (el numero a la derecha de la palabra "Grh")
            grh = Right(Fields(0), Len(Fields(0)) - 3)
            
            ' Leemos los campos de datos del Grh
            Fields = Split(Fields(1), SeparadorGrh)
                
            With GrhData(grh)
                    
                ' Primer lugar: cantidad de frames.
                .NumFrames = Val(Fields(0))
    
                ReDim .Frames(1 To .NumFrames)
                    
                ' Tiene mas de un frame entonces es una animacion
                If .NumFrames > 1 Then
                    
                    ' Segundo lugar: Leemos los numeros de grh de la animacion
                    For Frame = 1 To .NumFrames
                        .Frames(Frame) = Val(Fields(Frame))
                        If .Frames(Frame) <= LBound(GrhData) Or .Frames(Frame) > UBound(GrhData) Then GoTo hErr
                    Next
                        
                    ' Tercer lugar: leemos la velocidad de la animacion
                    .speed = Val(Fields(Frame))
                    If .speed <= 0 Then GoTo hErr
                        
                    ' Por ultimo, copiamos las dimensiones del primer frame
                    .pixelHeight = GrhData(.Frames(1)).pixelHeight
                    .pixelWidth = GrhData(.Frames(1)).pixelWidth
                    .TileWidth = GrhData(.Frames(1)).TileWidth
                    .TileHeight = GrhData(.Frames(1)).TileHeight
        
                ElseIf .NumFrames = 1 Then
                    
                    ' Si es un solo frame lo asignamos a si mismo
                    .Frames(1) = grh
                        
                    ' Segundo lugar: NumeroDelGrafico.bmp, pero sin el ".bmp"
                    .FileNum = Val(Fields(1))
                    If .FileNum <= 0 Then GoTo hErr
                            
                    ' Tercer Lugar: La coordenada X del grafico
                    .sX = Val(Fields(2))
                    If .sX < 0 Then GoTo hErr
                            
                    ' Cuarto Lugar: La coordenada Y del grafico
                    .sY = Val(Fields(3))
                    If .sY < 0 Then GoTo hErr
                            
                    ' Quinto lugar: El ancho del grafico
                    .pixelWidth = Val(Fields(4))
                    If .pixelWidth <= 0 Then GoTo hErr
                            
                    ' Sexto lugar: La altura del grafico
                    .pixelHeight = Val(Fields(5))
                    If .pixelHeight <= 0 Then GoTo hErr
                        
                    ' Calculamos el ancho y alto en tiles
                    .TileWidth = .pixelWidth / TilePixelHeight
                    .TileHeight = .pixelHeight / TilePixelWidth
                        
                Else
                    ' 0 frames o negativo? Error
                    GoTo hErr
                End If
        
            End With
        End If
    Loop
    
hErr:
    Close FileHandle
    
    If Err.Number <> 0 Then
        
        If Err.Number = 53 Then
            Call MsgBox("El archivo Graficos.ini no existe. Por favor, reinstale el juego.", , "Argentum 20")
        
        ElseIf grh > 0 Then
            Call MsgBox("Hay un error en Graficos.ini con el Grh" & grh & ".", , "Argentum 20")
        
        Else
            Call MsgBox("Hay un error en Graficos.ini. Por favor, reinstale el juego.", , "Argentum 20")
        End If
        
        Call CloseClient
        
    End If
    
    #If Compresion = 1 Then
        Delete_File Windows_Temp_Dir & "graficos.ini"
    #End If
    
    Exit Sub

End Sub


Sub CargarAnimArmas()
    
    On Error GoTo CargarAnimArmas_Err
    
    
    Dim Loader       As clsIniManager

    Dim i            As Long
    
    Dim j            As Byte
    
    Dim k            As Integer
    
    Dim Heading      As Byte
    
    Dim ArmaKey      As String
    
    Dim Std          As Byte

    Dim NumCuerpos   As Integer
    
    Dim LastGrh      As Long
    
    Dim AnimStart    As Long
    
    Dim x            As Long
    
    Dim y            As Long
    
    Dim FileNum      As Long
    
    Set Loader = New clsIniManager

    Dim loopc As Long

    Dim Arch  As String
    
    
    #If Compresion = 1 Then

        If Not Extract_File(Scripts, App.path & "\..\Recursos\OUTPUT\", "armas.dat", Windows_Temp_Dir, ResourcesPassword, False) Then
            Err.Description = "¡No se puede cargar el archivo de armas.dat!"
            MsgBox Err.Description

        End If
        
        Call Loader.Initialize(Windows_Temp_Dir & "armas.dat")
    #Else
        Call Loader.Initialize(App.path & "\..\Recursos\init\armas.dat")
    #End If
    
    NumWeaponAnims = Val(Loader.GetValue("INIT", "NumArmas"))
    
    
    
    ReDim WeaponAnimData(1 To NumWeaponAnims) As WeaponAnimData

    For loopc = 1 To NumWeaponAnims
        ArmaKey = "ARMA" & loopc
        Std = Val(Loader.GetValue(ArmaKey, "Std"))
        
        If Std = 0 Then
            
            InitGrh WeaponAnimData(loopc).WeaponWalk(1), Val(Loader.GetValue(ArmaKey, "Dir1")), 0
            InitGrh WeaponAnimData(loopc).WeaponWalk(2), Val(Loader.GetValue(ArmaKey, "Dir2")), 0
            InitGrh WeaponAnimData(loopc).WeaponWalk(3), Val(Loader.GetValue(ArmaKey, "Dir3")), 0
            InitGrh WeaponAnimData(loopc).WeaponWalk(4), Val(Loader.GetValue(ArmaKey, "Dir4")), 0
            
        Else
        
        
            FileNum = Val(Loader.GetValue(ArmaKey, "FileNum"))
        
            LastGrh = UBound(GrhData)

            ' Agrego espacio para meter el body en GrhData
            ReDim Preserve GrhData(1 To LastGrh + MoldesBodies(Std).TotalGrhs)
            
            MaxGrh = UBound(GrhData)
            
            LastGrh = LastGrh + 1
            x = MoldesBodies(Std).x
            y = MoldesBodies(Std).y
            
            For j = 1 To 4
                AnimStart = LastGrh
            
                For k = 1 To MoldesBodies(Std).DirCount(j)
                    With GrhData(LastGrh)
                        .FileNum = FileNum
                        .NumFrames = 1
                        .sX = x
                        .sY = y
                        .pixelWidth = MoldesBodies(Std).Width
                        .pixelHeight = MoldesBodies(Std).Height
                        
                        .TileWidth = .pixelWidth / TilePixelHeight
                        .TileHeight = .pixelHeight / TilePixelWidth
        
                        ReDim .Frames(1)
                        .Frames(1) = LastGrh
                    End With
                    
                    LastGrh = LastGrh + 1
                    x = x + MoldesBodies(Std).Width
                Next
                
                x = MoldesBodies(Std).x
                y = y + MoldesBodies(Std).Height
                
                Heading = BodiesHeading(j)
                
                With GrhData(LastGrh)
                    .NumFrames = MoldesBodies(Std).DirCount(j)
                    .speed = .NumFrames / 0.018
                    
                    ReDim .Frames(1 To MoldesBodies(Std).DirCount(j))
                    
                    For k = 1 To MoldesBodies(Std).DirCount(j)
                        .Frames(k) = AnimStart + k - 1
                    Next
                    
                    .pixelWidth = GrhData(.Frames(1)).pixelWidth
                    .pixelHeight = GrhData(.Frames(1)).pixelHeight
                    .TileWidth = GrhData(.Frames(1)).TileWidth
                    .TileHeight = GrhData(.Frames(1)).TileHeight
                End With
                
                InitGrh WeaponAnimData(loopc).WeaponWalk(Heading), LastGrh, 0
                
                
                LastGrh = LastGrh + 1
            Next
        
        
        End If
    Next loopc
    
    #If Compresion = 1 Then
        Delete_File Windows_Temp_Dir & "armas.dat"
    #End If

    
    Exit Sub

CargarAnimArmas_Err:
    Call RegistrarError(Err.Number, Err.Description, "Recursos.CargarAnimArmas", Erl)
    Resume Next
    
End Sub

Sub CargarColores()
    
    On Error GoTo CargarColores_Err
    

    

    Dim archivoC As String

    #If Compresion = 1 Then

        If Not Extract_File(Scripts, App.path & "\..\Recursos\OUTPUT\", "colores.dat", Windows_Temp_Dir, ResourcesPassword, False) Then
            Err.Description = "¡No se puede cargar el archivo de colores.dat!"
            MsgBox Err.Description

        End If

        archivoC = Windows_Temp_Dir & "colores.dat"
    #Else
        archivoC = App.path & "\..\Recursos\init\colores.dat"
    #End If
    
    If Not FileExist(archivoC, vbArchive) Then
        'TODO : Si hay que reinstalar, porque no cierra???
        Call MsgBox("ERROR: no se ha podido cargar los colores. Falta el archivo colores.dat, reinstale el juego", vbCritical + vbOKOnly)
        Exit Sub

    End If
    
    Dim i As Long
    
    For i = 0 To 47 '49 y 50 reservados para ciudadano y criminal
        ColoresPJ(i).r = CByte(GetVar(archivoC, CStr(i), "R"))
        ColoresPJ(i).G = CByte(GetVar(archivoC, CStr(i), "G"))
        ColoresPJ(i).B = CByte(GetVar(archivoC, CStr(i), "B"))
    Next i
    
    ColoresPJ(50).r = CByte(GetVar(archivoC, "CR", "R"))
    ColoresPJ(50).G = CByte(GetVar(archivoC, "CR", "G"))
    ColoresPJ(50).B = CByte(GetVar(archivoC, "CR", "B"))
    
    ColoresPJ(49).r = CByte(GetVar(archivoC, "CI", "R"))
    ColoresPJ(49).G = CByte(GetVar(archivoC, "CI", "G"))
    ColoresPJ(49).B = CByte(GetVar(archivoC, "CI", "B"))
    
    ColoresPJ(48).r = CByte(GetVar(archivoC, "NE", "R"))
    ColoresPJ(48).G = CByte(GetVar(archivoC, "NE", "G"))
    ColoresPJ(48).B = CByte(GetVar(archivoC, "NE", "B"))
    
    #If Compresion = 1 Then
        Delete_File Windows_Temp_Dir & "colores.dat"
    #End If

    
    Exit Sub

CargarColores_Err:
    Call RegistrarError(Err.Number, Err.Description, "Recursos.CargarColores", Erl)
    Resume Next
    
End Sub

Sub CargarCrafteo()
    Dim FileName As String

    #If Compresion = 1 Then
        If Not Extract_File(Scripts, App.path & "\..\Recursos\OUTPUT\", "crafteo.ini", Windows_Temp_Dir, ResourcesPassword, False) Then
            Err.Description = "No se puede cargar el archivo de crafteo.ini"
            MsgBox Err.Description
        End If
        
        FileName = Windows_Temp_Dir & "crafteo.ini"
    #Else
        FileName = App.path & "\..\Recursos\init\crafteo.ini"
    #End If
    
    Dim Reader As clsIniManager
    Set Reader = New clsIniManager
    
    Call Reader.Initialize(FileName)
    
    ReDim TipoCrafteo(1 To Reader.NodesCount)
    
    Dim i As Byte
    For i = 0 To Reader.NodesCount - 1
        Dim nombre As String
        nombre = Reader.GetNode(i)
        
        Dim id As Byte
        id = Val(Reader.GetValue(nombre, "ID"))
        
        With TipoCrafteo(id)
            .nombre = nombre
            .Ventana = Reader.GetValue(nombre, "Ventana")
            .Inventario = Val(Reader.GetValue(nombre, "Inventario"))
            .Icono = Val(Reader.GetValue(nombre, "Icono"))
        End With
    Next

    Set Reader = Nothing
    
    #If Compresion = 1 Then
        Delete_File FileName
    #End If
End Sub

Sub CargarAnimEscudos()
    
    On Error GoTo CargarAnimEscudos_Err
    
    Dim Loader       As clsIniManager

    Dim i            As Long
    
    Dim j            As Byte
    
    Dim k            As Integer
    
    Dim Heading      As Byte
    
    Dim EscudoKey      As String
    
    Dim Std          As Byte

    Dim NumCuerpos   As Integer
    
    Dim LastGrh      As Long
    
    Dim AnimStart    As Long
    
    Dim x            As Long
    
    Dim y            As Long
    
    Dim FileNum      As Long
    
    Set Loader = New clsIniManager

    Dim loopc As Long

    Dim Arch  As String
    
    
    
    #If Compresion = 1 Then

        If Not Extract_File(Scripts, App.path & "\..\Recursos\OUTPUT\", "escudos.dat", Windows_Temp_Dir, ResourcesPassword, False) Then
            Err.Description = "¡No se puede cargar el archivo de escudos.dat!"
            MsgBox Err.Description

        End If

        Call Loader.Initialize(Windows_Temp_Dir & "escudos.dat")
    #Else
        Call Loader.Initialize(App.path & "\..\Recursos\init\escudos.dat")
    #End If
    
    
    NumEscudosAnims = Val(Loader.GetValue("INIT", "NumEscudos"))
    
    ReDim ShieldAnimData(1 To NumEscudosAnims) As ShieldAnimData
    
    For loopc = 1 To NumEscudosAnims
    
        EscudoKey = "ESC" & loopc
        Std = Val(Loader.GetValue(EscudoKey, "Std"))
        
        If Std = 0 Then
            InitGrh ShieldAnimData(loopc).ShieldWalk(1), Val(Loader.GetValue(EscudoKey, "Dir1")), 0
            InitGrh ShieldAnimData(loopc).ShieldWalk(2), Val(Loader.GetValue(EscudoKey, "Dir2")), 0
            InitGrh ShieldAnimData(loopc).ShieldWalk(3), Val(Loader.GetValue(EscudoKey, "Dir3")), 0
            InitGrh ShieldAnimData(loopc).ShieldWalk(4), Val(Loader.GetValue(EscudoKey, "Dir4")), 0
        Else
        
        
            FileNum = Val(Loader.GetValue(EscudoKey, "FileNum"))
        
            LastGrh = UBound(GrhData)

            ' Agrego espacio para meter el body en GrhData
            ReDim Preserve GrhData(1 To LastGrh + MoldesBodies(Std).TotalGrhs)
            
            MaxGrh = UBound(GrhData)
            
            LastGrh = LastGrh + 1
            x = MoldesBodies(Std).x
            y = MoldesBodies(Std).y
            
            For j = 1 To 4
                AnimStart = LastGrh
            
                For k = 1 To MoldesBodies(Std).DirCount(j)
                    With GrhData(LastGrh)
                        .FileNum = FileNum
                        .NumFrames = 1
                        .sX = x
                        .sY = y
                        .pixelWidth = MoldesBodies(Std).Width
                        .pixelHeight = MoldesBodies(Std).Height
                        
                        .TileWidth = .pixelWidth / TilePixelHeight
                        .TileHeight = .pixelHeight / TilePixelWidth
        
                        ReDim .Frames(1)
                        .Frames(1) = LastGrh
                    End With
                    
                    LastGrh = LastGrh + 1
                    x = x + MoldesBodies(Std).Width
                Next
                
                x = MoldesBodies(Std).x
                y = y + MoldesBodies(Std).Height
                
                Heading = BodiesHeading(j)
                
                With GrhData(LastGrh)
                    .NumFrames = MoldesBodies(Std).DirCount(j)
                    .speed = .NumFrames / 0.018
                    
                    ReDim .Frames(1 To MoldesBodies(Std).DirCount(j))
                    
                    For k = 1 To MoldesBodies(Std).DirCount(j)
                        .Frames(k) = AnimStart + k - 1
                    Next
                    
                    .pixelWidth = GrhData(.Frames(1)).pixelWidth
                    .pixelHeight = GrhData(.Frames(1)).pixelHeight
                    .TileWidth = GrhData(.Frames(1)).TileWidth
                    .TileHeight = GrhData(.Frames(1)).TileHeight
                End With
                
                InitGrh ShieldAnimData(loopc).ShieldWalk(Heading), LastGrh, 0
                
                
                LastGrh = LastGrh + 1
            Next
        
        
        End If
    Next loopc
    
    #If Compresion = 1 Then
        Delete_File Windows_Temp_Dir & "escudos.dat"
    #End If

    
    Exit Sub

CargarAnimEscudos_Err:
    Call RegistrarError(Err.Number, Err.Description, "Recursos.CargarAnimEscudos", Erl)
    Resume Next
    
End Sub


Sub LoadFonts()
    If LoadFont("Cardo.ttf") Then
        frmMain.NombrePJ.font.name = "Cardo"
    End If

    If LoadFont("Alegreya Sans AO.ttf") Then
        Dim CurControl As Control
        Dim Middle As Integer
    
        For Each CurControl In frmMain.Controls
            If CurControl.name <> "NombrePJ" Then
                Select Case TypeName(CurControl)
                    Case "Label"
                        CurControl.font.name = "Alegreya Sans AO"

                        ' Centrar texto verticalmente
                        If Not CurControl.AutoSize Then
                            Middle = Fix(CurControl.Top + CurControl.Height * 0.5)
                            CurControl.AutoSize = True
                            CurControl.Top = Fix(Middle - CurControl.Height * 0.5)
                        End If
                        
                    Case "RichTextBox", "ListBox"
                        CurControl.font.name = "Alegreya Sans AO"
                End Select
            End If
        Next

        Call SelLineSpacing(frmMain.RecTxt, 5, 22)
    End If
    '#If PYMMO = 1 Then
        Dim arr() As Byte
        
        ReDim arr(1 To 16) As Byte
        
    arr(15) = 1
    arr(16) = 62
    arr(4) = 7
    arr(3) = 2
    arr(13) = 56
    arr(5) = 22
    arr(14) = 9
    arr(7) = 21
    arr(10) = 52
    arr(9) = 23
    arr(12) = 28
    arr(11) = 19
    arr(8) = 38
    arr(6) = 22
    arr(1) = 11
    arr(2) = 64
        MapInfoEspeciales = estaInmovilizado(arr)
    '#End If
    
    #If DEBUGGING = 1 Then
        Debug.Print MapInfoEspeciales
    #Else
    
    #End If
End Sub

Function LoadFont(name As String) As Boolean
    Static YaMostreError As Boolean
    LoadFont = AddFontResourceEx(App.path & "\..\Recursos\OUTPUT\" & name, FR_PRIVATE, 0&) <> 0

    If Not YaMostreError And Not LoadFont Then
        Call MsgBox("No se pudieron cargar algunas fuentes, reinstale el juego para repararlas.", vbOKOnly, "Error al cargar - Argentum20")
        YaMostreError = True
    End If
End Function

Public Sub CargarNPCsMapData()
    Dim fh      As Integer
    Dim NumMaps As Integer
    
    fh = FreeFile

    NumMaps = Val(GetVar(App.path & "\..\Recursos\Dat\zonas.dat", "Mapas", "Cantidad"))
        
    Open App.path & "\..\Recursos\OUTPUT\QuestNPCsMapData.bin" For Binary As fh
       
    ReDim ListNPCMapData(1 To NumMaps) As t_MapNpc
    Dim x As Single
    Dim y As Single
    Do While Not EOF(fh)
        Dim map As Integer
        Get fh, , map
        
        If map > 0 Then
            ReDim ListNPCMapData(map).NpcList(1 To MAX_QUESTNPCS_VISIBLE) As t_QuestNPCMapData
            Dim i As Long
            For i = 1 To MAX_QUESTNPCS_VISIBLE
                Dim TempInt As Integer
                Get #fh, , TempInt
                'Debug.Assert map > 0
                ListNPCMapData(map).NpcList(i).NPCNumber = TempInt
                If TempInt > 0 Then
                    ListNPCMapData(map).NpcCount = ListNPCMapData(map).NpcCount + 1
                End If
                Get #fh, , TempInt
                x = TempInt
                
                Get #fh, , TempInt
                y = TempInt
                
                Get #fh, , TempInt
                ListNPCMapData(map).NpcList(i).state = TempInt
                Call ConvertToMinimapPosition(x, y, 2, 0)
                ListNPCMapData(map).NpcList(i).Position.x = x
                ListNPCMapData(map).NpcList(i).Position.y = y
            Next i
        End If
        DoEvents
    Loop
    Close fh
End Sub
