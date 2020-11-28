VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMeteorologic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'*****************************************************************
'clsMeteorologic - ImperiumAO - v1.3.0
'
'Meteorologic engine.
'
'*****************************************************************
'Respective portions copyrighted by contributors listed below.
'
'This library is free software; you can redistribute it and/or
'modify it under the terms of the GNU Lesser General Public
'License as published by the Free Software Foundation version 2.1 of
'the License
'
'This library is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU
'Lesser General Public License for more details.
'
'You should have received a copy of the GNU Lesser General Public
'License along with this library; if not, write to the Free Software
'Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA
'*****************************************************************

'*****************************************************************
'Sinuhe (sinuhe@imperiumao.com.ar)
'Augusto José Rando (barrin@imperiumao.com.ar)
'   - First Relase
'*****************************************************************

Option Explicit

Private Const DURACION_TRANS = 5000

Private Out_Color           As Long

'Main Status
Private Const DIA           As Byte = 1

Private Const ATARDECER     As Byte = 2

Private Const NOCHE         As Byte = 3

Private Const AMANECER      As Byte = 4

Private Const MUERTO        As Byte = 5

Private Const FIN_ANIMACION As Byte = &HFF

'Secondary Status
Private Const Normal        As Byte = 0

Private Const NUBLADO       As Byte = 1

Private Const LLUVIA        As Byte = 2

Private Const NIEVE         As Byte = 3

Private Const TORMENTA      As Byte = 4

Private Type SYSTEMTIME

    wYear As Integer
    wMonth As Integer
    wDayOfWeek As Integer
    wDay As Integer
    wHour As Integer
    wMinute As Integer
    wSecond As Integer
    wMilliseconds As Integer

End Type

Private Type MeteoStat

    EstadoActual As Byte
    EstadoProximo As Byte

End Type

Private Type Efectos

    MainStatus As MeteoStat
    SecondaryStatus As MeteoStat

End Type

Private Type D3DCOLORVALUE

    a As Long
    r As Long
    g As Long
    b As Long

End Type

Private Meteo              As Efectos

Private m_Color            As D3DCOLORVALUE

Private TimeElapsed        As Single

Private m_Color_Dia        As D3DCOLORVALUE

Private m_Color_Noche      As D3DCOLORVALUE

Private m_Color_Tarde      As D3DCOLORVALUE

Private m_Color_Manana     As D3DCOLORVALUE

Private m_Color_Muerto     As D3DCOLORVALUE

Private m_Hora_Actual      As Long

Private m_Last_Hora_Actual As Long

Private Meteo_Visible      As Boolean

Private Declare Sub GetSystemTime Lib "Kernel32.dll" (lpSystemTime As SYSTEMTIME)

Private theTime As SYSTEMTIME

'Inicializa colores de los diferentes estados
Private Sub Class_Initialize()

    'No animo nada al iniciar
    Meteo.MainStatus.EstadoProximo = FIN_ANIMACION
    Meteo.MainStatus.EstadoActual = DIA

    With m_Color_Dia
        .a = 255
        .b = 255
        .r = 255
        .g = 255

    End With

    With m_Color_Tarde
        .a = 255
        .b = 251
        .r = 170
        .g = 172

    End With

    With m_Color_Noche
        .a = 255
        .b = 160
        .r = 130
        .g = 130

    End With

    With m_Color_Manana
        .a = 255
        .b = 190
        .r = 230
        .g = 200

    End With

    With m_Color_Muerto
        .a = 255
        .b = 120
        .r = 120
        .g = 120

    End With

    m_Last_Hora_Actual = -1

End Sub

Public Sub Meteo_Logic()

    Dim Diferencia    As Integer

    Dim ParticleIndex As Long

    Dim CambioValor   As Single

    Dim C1            As D3DCOLORVALUE

    Dim C2            As D3DCOLORVALUE

    If Meteo.MainStatus.EstadoProximo <> FIN_ANIMACION Then
        TimeElapsed = TimeElapsed + General_Get_Elapsed_Time
        CambioValor = TimeElapsed / DURACION_TRANS

        If CambioValor > 1 Then
            Meteo.MainStatus.EstadoActual = Meteo.MainStatus.EstadoProximo

            Rem  Meteo.MainStatus.EstadoProximo = FIN_ANIMACION
        End If

    End If

    Select Case Meteo.MainStatus.EstadoActual

        Case DIA

            If Meteo.MainStatus.EstadoProximo = ATARDECER Then
                C1 = m_Color_Dia
                C2 = m_Color_Tarde
                Diferencia = C1.r - C2.r

                If Diferencia > 255 Then Diferencia = 255
                m_Color.r = C1.r - Diferencia * CambioValor
            
                Diferencia = C1.g - C2.g

                If Diferencia > 255 Then Diferencia = 255
                m_Color.g = C1.g - Diferencia * CambioValor
            
                Diferencia = C1.b - C2.b

                If Diferencia > 255 Then Diferencia = 255
                m_Color.b = C1.b - Diferencia * CambioValor
            ElseIf Meteo.MainStatus.EstadoProximo = MUERTO Then
                C1 = m_Color_Dia
                C2 = m_Color_Muerto
                Diferencia = C1.r - C2.r

                If Diferencia > 255 Then Diferencia = 255
                m_Color.r = C1.r - Diferencia * CambioValor
            
                Diferencia = C1.g - C2.g

                If Diferencia > 255 Then Diferencia = 255
                m_Color.g = C1.g - Diferencia * CambioValor
            
                Diferencia = C1.b - C2.b

                If Diferencia > 255 Then Diferencia = 255
                m_Color.b = C1.b - Diferencia * CambioValor
        
            Else
                m_Color = m_Color_Dia

            End If

        Case NOCHE

            If Meteo.MainStatus.EstadoProximo = AMANECER Then
                C1 = m_Color_Noche
                C2 = m_Color_Manana
                Diferencia = C1.r - C2.r

                If Diferencia > 255 Then Diferencia = 255
                m_Color.r = C1.r - Diferencia * CambioValor
                
                Diferencia = C1.g - C2.g

                If Diferencia > 255 Then Diferencia = 255
                m_Color.g = C1.g - Diferencia * CambioValor
                
                Diferencia = C1.b - C2.b

                If Diferencia > 255 Then Diferencia = 255
                m_Color.b = C1.b - Diferencia * CambioValor
            ElseIf Meteo.MainStatus.EstadoProximo = MUERTO Then
                C1 = m_Color_Noche
                C2 = m_Color_Muerto
                Diferencia = C1.r - C2.r

                If Diferencia > 255 Then Diferencia = 255
                m_Color.r = C1.r - Diferencia * CambioValor
            
                Diferencia = C1.g - C2.g

                If Diferencia > 255 Then Diferencia = 255
                m_Color.g = C1.g - Diferencia * CambioValor
            
                Diferencia = C1.b - C2.b

                If Diferencia > 255 Then Diferencia = 255
                m_Color.b = C1.b - Diferencia * CambioValor
            Else
                m_Color = m_Color_Noche

            End If

        Case AMANECER

            If Meteo.MainStatus.EstadoProximo = DIA Then
                C1 = m_Color_Manana
                C2 = m_Color_Dia
                Diferencia = C1.r - C2.r

                If Diferencia > 255 Then Diferencia = 255
                m_Color.r = C1.r - Diferencia * CambioValor
                
                Diferencia = C1.g - C2.g

                If Diferencia > 255 Then Diferencia = 255
                m_Color.g = C1.g - Diferencia * CambioValor
                
                Diferencia = C1.b - C2.b

                If Diferencia > 255 Then Diferencia = 255
                m_Color.b = C1.b - Diferencia * CambioValor
            ElseIf Meteo.MainStatus.EstadoProximo = MUERTO Then
                C1 = m_Color_Manana
                C2 = m_Color_Muerto
                Diferencia = C1.r - C2.r

                If Diferencia > 255 Then Diferencia = 255
                m_Color.r = C1.r - Diferencia * CambioValor
            
                Diferencia = C1.g - C2.g

                If Diferencia > 255 Then Diferencia = 255
                m_Color.g = C1.g - Diferencia * CambioValor
            
                Diferencia = C1.b - C2.b

                If Diferencia > 255 Then Diferencia = 255
                m_Color.b = C1.b - Diferencia * CambioValor
            Else
                m_Color = m_Color_Manana

            End If

        Case ATARDECER

            If Meteo.MainStatus.EstadoProximo = NOCHE Then
                C1 = m_Color_Tarde
                C2 = m_Color_Noche
                Diferencia = C1.r - C2.r

                If Diferencia > 255 Then Diferencia = 255
                m_Color.r = C1.r - Diferencia * CambioValor
                
                Diferencia = C1.g - C2.g

                If Diferencia > 255 Then Diferencia = 255
                m_Color.g = C1.g - Diferencia * CambioValor
                
                Diferencia = C1.b - C2.b

                If Diferencia > 255 Then Diferencia = 255
                m_Color.b = C1.b - Diferencia * CambioValor
            ElseIf Meteo.MainStatus.EstadoProximo = MUERTO Then
                C1 = m_Color_Tarde
                C2 = m_Color_Muerto
                Diferencia = C1.r - C2.r

                If Diferencia > 255 Then Diferencia = 255
                m_Color.r = C1.r - Diferencia * CambioValor
            
                Diferencia = C1.g - C2.g

                If Diferencia > 255 Then Diferencia = 255
                m_Color.g = C1.g - Diferencia * CambioValor
            
                Diferencia = C1.b - C2.b

                If Diferencia > 255 Then Diferencia = 255
                m_Color.b = C1.b - Diferencia * CambioValor
            Else
                m_Color = m_Color_Tarde

            End If
        
        Case MUERTO

            Dim Puesto As Boolean
        
            '     If Meteo.MainStatus.EstadoProximo = NOCHE Then
            C1 = m_Color_Muerto

            If Meteo.MainStatus.EstadoProximo = 1 Then
                C2 = m_Color_Dia
                Puesto = False
            ElseIf Meteo.MainStatus.EstadoProximo = 2 Then
                C2 = m_Color_Tarde
                Puesto = False
            ElseIf Meteo.MainStatus.EstadoProximo = 3 Then
                C2 = m_Color_Noche
                Puesto = False
            ElseIf Meteo.MainStatus.EstadoProximo = 4 Then
                C2 = m_Color_Manana
                Puesto = False
            Else
                m_Color = m_Color_Muerto
                Puesto = True

            End If
            
            If Not Puesto Then
                Diferencia = C1.r - C2.r

                If Diferencia > 255 Then Diferencia = 255
                m_Color.r = C1.r - Diferencia * CambioValor
                
                Diferencia = C1.g - C2.g

                If Diferencia > 255 Then Diferencia = 255
                m_Color.g = C1.g - Diferencia * CambioValor
                
                Diferencia = C1.b - C2.b

                If Diferencia > 255 Then Diferencia = 255
                m_Color.b = C1.b - Diferencia * CambioValor

            End If
        
    End Select

    Out_Color = D3DColorARGB(m_Color.a, m_Color.r, m_Color.g, m_Color.b)
        
    If ColorAmbiente.r <> m_Color.r Or ColorAmbiente.g <> m_Color.g Or ColorAmbiente.b <> m_Color.b Then
        ColorAmbiente.a = m_Color.a
        ColorAmbiente.r = m_Color.r
        ColorAmbiente.g = m_Color.g
        ColorAmbiente.b = m_Color.b
        LucesRedondas.LightRenderAll

    End If

    ColorAmbiente.a = m_Color.a
    ColorAmbiente.r = m_Color.r
    ColorAmbiente.g = m_Color.g
    ColorAmbiente.b = m_Color.b
        
    Call Map_Base_Light_Set(Out_Color)

End Sub

Public Sub Time_Logic()

    GetSystemTime theTime
    m_Hora_Actual = format((theTime.wHour - 3), "00")

    If Val(m_Hora_Actual) = -3 Then m_Hora_Actual = "21"
    If Val(m_Hora_Actual) = -2 Then m_Hora_Actual = "22"
    If Val(m_Hora_Actual) = -1 Then m_Hora_Actual = "23"

    If m_Hora_Actual <> m_Last_Hora_Actual Then
        Rem  frmMain.imgHora.Picture = LoadInterface(m_Hora_Actual & ".bmp")
        m_Last_Hora_Actual = m_Hora_Actual

    End If

End Sub

Public Function Get_Time_String() As String

    Dim HoraActual   As String

    Dim HoraYminutos As String

    HoraActual = GetHora(HoraFantasia)

    HoraYminutos = GetTimeFormated(HoraFantasia)

    Get_Time_String = HoraYminutos & "... "

    Select Case HoraActual

        Case 5, 6, 7
            Get_Time_String = Get_Time_String & "el sol se asoma lentamente en el horizonte..."

        Case 8, 9, 10, 11, 12, 13, 14, 15, 16, 17
            Get_Time_String = Get_Time_String & "¡No pierdas el tiempo!"

        Case 18, 19
            Get_Time_String = Get_Time_String & "lentamente el día termina..."

        Case Else
            Get_Time_String = Get_Time_String & "¿Despierto a estas horas? ¡No olvides visitar El Mesón Hostigado!"

    End Select

End Function

Public Sub SetNuevoEstado(ByVal estado As Byte)

    TimeElapsed = 0

    If Meteo.MainStatus.EstadoProximo <> FIN_ANIMACION Then
        'Como no termino la animación anterior, al menos le pongo que la que estaba haciendo sea desde donde sale
        Meteo.MainStatus.EstadoActual = Meteo.MainStatus.EstadoProximo

    End If

    Meteo.MainStatus.EstadoProximo = estado

    If bNiebla And Not bRain Then
        '   frmMain.PicMeteo.Picture = LoadInterface("clima_nublado.bmp")
    ElseIf bRain Then
        '  frmMain.PicMeteo.Picture = LoadInterface("clima_lluvia.bmp")
    ElseIf estado = 1 Then
        ' frmMain.PicMeteo.Picture = LoadInterface("clima_mediodia.bmp")
    ElseIf estado = 2 Then
        ' frmMain.PicMeteo.Picture = LoadInterface("clima_atardecer.bmp")
    ElseIf estado = 3 Then
        ' frmMain.PicMeteo.Picture = LoadInterface("clima_noche.bmp")
    ElseIf estado = 4 Then

        'frmMain.PicMeteo.Picture = LoadInterface("clima_amanecer.bmp")
    End If

End Sub

Public Sub CargarClima()

    If bNiebla And Not bRain Then
        '   frmMain.PicMeteo.Picture = LoadInterface("clima_nublado.bmp")
    ElseIf bRain Then
        ' frmMain.PicMeteo.Picture = LoadInterface("clima_lluvia.bmp")
    ElseIf meteo_estado = 1 Then
        ' frmMain.PicMeteo.Picture = LoadInterface("clima_mediodia.bmp")
    ElseIf meteo_estado = 2 Then
        ' frmMain.PicMeteo.Picture = LoadInterface("clima_atardecer.bmp")
    ElseIf meteo_estado = 3 Then
        ' frmMain.PicMeteo.Picture = LoadInterface("clima_noche.bmp")
    ElseIf meteo_estado = 4 Then

        ' frmMain.PicMeteo.Picture = LoadInterface("clima_amanecer.bmp")
    End If

End Sub

Public Property Get SecondaryStatus() As Byte
    SecondaryStatus = Meteo.SecondaryStatus.EstadoActual

End Property

Public Property Let SecondaryStatus(ByVal vNewValue As Byte)
    Meteo.SecondaryStatus.EstadoActual = vNewValue

End Property

Public Property Get PrimaryStatus() As Byte
    PrimaryStatus = Meteo.MainStatus.EstadoActual

End Property

Public Property Let PrimaryStatus(ByVal vNewValue As Byte)
    Meteo.MainStatus.EstadoProximo = vNewValue

End Property

Public Sub ForzarEstado(ByVal estado As Byte)
    TimeElapsed = 0
    Meteo.MainStatus.EstadoActual = estado
    Meteo.MainStatus.EstadoProximo = FIN_ANIMACION

    If bNiebla And Not bRain Then
        '  frmMain.PicMeteo.Picture = LoadInterface("clima_nublado.bmp")
    ElseIf bRain Then
        ' frmMain.PicMeteo.Picture = LoadInterface("clima_lluvia.bmp")
    ElseIf meteo_estado = 1 Then
        '  frmMain.PicMeteo.Picture = LoadInterface("clima_mediodia.bmp")
    ElseIf meteo_estado = 2 Then
        '  frmMain.PicMeteo.Picture = LoadInterface("clima_atardecer.bmp")
    ElseIf meteo_estado = 3 Then
        ' frmMain.PicMeteo.Picture = LoadInterface("clima_noche.bmp")
    ElseIf meteo_estado = 4 Then

        ' frmMain.PicMeteo.Picture = LoadInterface("clima_amanecer.bmp")
    End If

End Sub

Public Sub NextEstado()

    TimeElapsed = 0

    If Meteo.MainStatus.EstadoProximo <> FIN_ANIMACION Then
        'Como no termino la animación anterior, al menos le pongo que la que estaba haciendo sea desde donde sale
        Meteo.MainStatus.EstadoActual = Meteo.MainStatus.EstadoProximo

    End If

    Meteo.MainStatus.EstadoProximo = Meteo.MainStatus.EstadoActual + 1

    If Meteo.MainStatus.EstadoProximo > AMANECER Then Meteo.MainStatus.EstadoProximo = DIA

End Sub

