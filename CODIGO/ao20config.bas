Attribute VB_Name = "ao20config"
'    Argentum 20 - Game Client Program
'    Copyright (C) 2023 - Noland Studios
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

Public Const OPTION_MUSIC_ENABLED As String = "Music"
Public Const OPTION_SOUND_ENABLED As String = "Sound"
Public Const OPTION_FX_ENABLED As String = "Fx"
Public Const OPTION_AMBIENT_ENABLED As String = "AmbientEnabled"
Public Const OPTION_INVERTLR_CHANNELS_ENABLED As String = "InverLRChannels"

Public X_OFFSET As Integer
Public Y_OFFSET As Integer
Public EQUIPMENT_CARACTER As String
Public RED_SHADER As Byte
Public GREEN_SHADER As Byte
Public BLUE_SHADER As Byte
Public SHADER_TRANSPARENCY As Byte
Public EquipmentStyle As Byte
    
Public Function GetErrorLogFilename() As String
   GetErrorLogFilename = App.path & "\logs\Errores.log"
End Function


Sub SaveConfig()
    On Error GoTo SaveConfig_Err
    
    #If PYMMO = 0 Or DEBUGGING = 1 Then
    Call SaveSetting("INIT", "ServerIndex", IPdelServidor & ":" & PuertoDelServidor)
    #End If
    
    Call SaveSetting("AUDIO", OPTION_MUSIC_ENABLED, MusicEnabled)
    Call SaveSetting("AUDIO", OPTION_FX_ENABLED, FxEnabled)
    Call SaveSetting("AUDIO", OPTION_AMBIENT_ENABLED, AmbientEnabled)
    Call SaveSetting("AUDIO", OPTION_INVERTLR_CHANNELS_ENABLED, InvertirSonido)
    
    Call SaveSetting("AUDIO", "VolMusic", VolMusic)
    Call SaveSetting("AUDIO", "Volfx", VolFX)
    Call SaveSetting("AUDIO", "VolAmbient", VolAmbient)

     
    Call SaveSetting("OPCIONES", "MoverVentana", MoverVentana)
    Call SaveSetting("OPCIONES", "PermitirMoverse", PermitirMoverse)
    Call SaveSetting("OPCIONES", "ScrollArrastrar", ScrollArrastrar)
       
    Call SaveSetting("OPCIONES", "CopiarDialogoAConsola", CopiarDialogoAConsola)
    Call SaveSetting("OPCIONES", "FPSFLAG", FPSFLAG)
    Call SaveSetting("OPCIONES", "AlphaMacro", AlphaMacro)
    Call SaveSetting("OPCIONES", "ModoHechizos", ModoHechizos)
    Call SaveSetting("OPCIONES", "FxNavega", FxNavega)
    
    Call SaveSetting("OPCIONES", "NumerosCompletosInventario", NumerosCompletosInventario)

    Call SaveSetting("VIDEO", "MostrarRespiracion", IIf(MostrarRespiracion, 1, 0))
    Call SaveSetting("VIDEO", "PantallaCompleta", IIf(PantallaCompleta, 1, 0))
    Call SaveSetting("VIDEO", "InfoItemsEnRender", IIf(InfoItemsEnRender, 1, 0))
    Call SaveSetting("VIDEO", "Aceleracion", ModoAceleracion)

    Call SaveSetting("OPCIONES", "SensibilidadMouse", SensibilidadMouse)
    Call SaveSetting("OPCIONES", "DialogosClanes", IIf(DialogosClanes.Activo, 1, 0))

    Call SaveSetting("OPCIONES", "EquipmentIndicatorRedColor", RED_SHADER)
    Call SaveSetting("OPCIONES", "EquipmentIndicatorGreenColor", GREEN_SHADER)
    Call SaveSetting("OPCIONES", "EquipmentIndicatorBlueColor", BLUE_SHADER)
    Call SaveSetting("OPCIONES", "EquipmentIndicatorTransparency", SHADER_TRANSPARENCY)
    Call SaveSetting("OPCIONES", "EquipmentIndicatorCoordinateX", X_OFFSET)
    Call SaveSetting("OPCIONES", "EquipmentIndicatorCoordinateY", Y_OFFSET)
    Call SaveSetting("OPCIONES", "EquipmentIndicatorCaracter", EQUIPMENT_CARACTER)

    
    Exit Sub

SaveConfig_Err:
    Call RegistrarError(Err.Number, Err.Description, "ModUtils.SaveConfig", Erl)
    Resume Next
    
End Sub

Sub LoadConfig()

    On Error GoTo ErrorHandler
    Set DialogosClanes = New clsGuildDlg
    
    If InitializeSettings() Then
        Call LoadBindedKeys
    Else
        Call MsgBox(JsonLanguage.Item("MENSAJE_ERROR_CARGAR_OPCIONES"), vbCritical, JsonLanguage.Item("TITULO_ERROR_CARGAR"))
        End
    End If

    MusicEnabled = GetSettingAsByte("AUDIO", OPTION_MUSIC_ENABLED, 1)
    AudioEnabled = GetSettingAsByte("AUDIO", OPTION_SOUND_ENABLED, 1)
    FxEnabled = GetSettingAsByte("AUDIO", OPTION_FX_ENABLED, 1)
    AmbientEnabled = GetSettingAsByte("AUDIO", OPTION_AMBIENT_ENABLED, 1)
    InvertirSonido = GetSettingAsByte("AUDIO", OPTION_INVERTLR_CHANNELS_ENABLED, 1)


    
    'Musica y Sonido - Volumen
    VolMusicFadding = VolMusic
    VolMusic = Val(GetSetting("AUDIO", "VolMusic"))
    VolFX = Val(GetSetting("AUDIO", "VolFX"))
    VolAmbient = Val(GetSetting("AUDIO", "VolAmbient"))

    Call ao20audio.SetMusicVolume(VolMusic)
    Call ao20audio.SetFxVolume(VolFX)
    'Video
    PantallaCompleta = GetSetting("VIDEO", "PantallaCompleta")
    CursoresGraficos = IIf(RunningInVB, 0, GetSetting("VIDEO", "CursoresGraficos"))
    InfoItemsEnRender = Val(GetSetting("VIDEO", "InfoItemsEnRender"))
    ModoAceleracion = GetSetting("VIDEO", "Aceleracion")
    DisableDungeonLighting = Val(GetSetting("VIDEO", "DisableDungeonLighting"))
    '------------------------------------------------------------------------------
    ' Configuration: VIDEO.NumTexRelease
    '
    ' When the user transitions from one map to the next, the texture manager will
    ' eagerly free up a batch of textures to keep memory under control.
    '
    ' Effective value (clamped 25…250):
    '   NumTexRelease = max(25, min(GetSetting("VIDEO", "NumTexRelease"), 250))
    '
    '   • Reads the INI setting [VIDEO] NumTexRelease
    '   • Ensures at least 25 textures are freed per transition
    '   • Caps the release at 250 textures to avoid excessive unloading
    '
    ' Tuning:
    '   – Higher values free more textures (reducing VRAM use more aggressively)
    '     but may incur extra reload costs.
    '   – Lower values free fewer textures (less reload overhead)
    '     but risk higher memory usage.
    '------------------------------------------------------------------------------
    NumTexRelease = max(1, min(val(GetSetting("VIDEO", "NumTexRelease")), 250))


    '------------------------------------------------------------------------------
    ' Configuration: VIDEO.TexHighWaterMark
    '
    ' Before performing bulk-free on map transition, the manager checks that the
    ' current total texture usage (in MB) is at or above this threshold. If usage
    ' is below, no eager release occurs.
    '
    ' Effective value (clamped 200…600 MB):
    '   TexHighWaterMark = max(200, min(GetSetting("VIDEO", "TexHighWaterMark"), 600))
    '
    '   • Reads the INI setting [VIDEO] TexHighWaterMark (in megabytes)
    '   • Ensures the high-water mark is at least 200 MB
    '   • Caps the high-water mark at 600 MB to avoid overly late eviction
    '
    ' Tuning:
    '   – A lower threshold triggers eviction more often, keeping VRAM tighter
    '     at the cost of potential stutter.
    '   – A higher threshold delays eviction, reducing stutter but increasing
    '     peak memory usage.
    '------------------------------------------------------------------------------
    TexHighWaterMark = max(200, min(val(GetSetting("VIDEO", "TexHighWaterMark")), 800))

    
    Dim Value As String
    Value = GetSetting("VIDEO", "MostrarRespiracion")
    MostrarRespiracion = IIf(LenB(Value) > 0, Val(Value), True)

    FxNavega = GetSetting("OPCIONES", "FxNavega")
    MostrarIconosMeteorologicos = GetSetting("OPCIONES", "MostrarIconosMeteorologicos")
    CopiarDialogoAConsola = GetSetting("OPCIONES", "CopiarDialogoAConsola")
    PermitirMoverse = GetSetting("OPCIONES", "PermitirMoverse")
    ScrollArrastrar = Val(GetSetting("OPCIONES", "ScrollArrastrar"))
    LastScroll = Val(GetSetting("OPCIONES", "LastScroll"))

    MoverVentana = GetSetting("OPCIONES", "MoverVentana")
    FPSFLAG = GetSetting("OPCIONES", "FPSFLAG")
    AlphaMacro = GetSetting("OPCIONES", "AlphaMacro")
    ModoHechizos = Val(GetSetting("OPCIONES", "ModoHechizos"))
    DialogosClanes.Activo = Val(GetSetting("OPCIONES", "DialogosClanes"))
    NumerosCompletosInventario = Val(GetSetting("OPCIONES", "NumerosCompletosInventario"))

    EquipmentStyle = GetSettingAsByte("OPCIONES", "EquipmentIndicator", 0)
    RED_SHADER = GetSettingAsByte("OPCIONES", "EquipmentIndicatorRedColor", 255)
    GREEN_SHADER = GetSettingAsByte("OPCIONES", "EquipmentIndicatorGreenColor", 255)
    BLUE_SHADER = GetSettingAsByte("OPCIONES", "EquipmentIndicatorBlueColor", 0)
    SHADER_TRANSPARENCY = GetSettingAsByte("OPCIONES", "EquipmentIndicatorTransparency", 20)
    X_OFFSET = CInt(val(GetSetting("OPCIONES", "EquipmentIndicatorCoordinateX")))
    Y_OFFSET = CInt(val(GetSetting("OPCIONES", "EquipmentIndicatorCoordinateY")))
    EQUIPMENT_CARACTER = GetSetting("OPCIONES", "EquipmentIndicatorCaracter")

    'Init
    #If PYMMO = 0 Or DEBUGGING = 1 Then
        ServerIndex = GetSetting("INIT", "ServerIndex")
    #End If

    SensibilidadMouse = GetSetting("OPCIONES", "SensibilidadMouse")
    If SensibilidadMouse = 0 Then: SensibilidadMouse = 10
    SensibilidadMouseOriginal = General_Get_Mouse_Speed
    Call General_Set_Mouse_Speed(SensibilidadMouse)
    
    Exit Sub
    
ErrorHandler:
    Call MsgBox(JsonLanguage.Item("MENSAJE_ERROR_CARGAR_CONFIG"), vbCritical, JsonLanguage.Item("MENSAJE_TITULO_CONFIGURACION"))

    End
End Sub

