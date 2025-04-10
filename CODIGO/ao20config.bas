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

