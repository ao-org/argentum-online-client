VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CBufferMan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
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
'*****************************************************************
'CBufferMan - ImperiumAO - v1.3.0
'
'Sound buffer manager based on Maraxus's texture manager.
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
'Maraxus (Juan Martín Sotuyo Dodero - juansotuyo@hotmail.com)
'   - First Relase (as CTextureMan)
'Barrin (Augusto José Rando - barrin@imperiumao.com.ar)
'   - Modified to use sound buffers
'   - Modified to copy buffers when already playing
'*****************************************************************

Option Explicit

Private Const MAX_BUFFERS As Integer = 3

Private Type WAV_DB_ENTRY

    FileName As Integer
    UltimoAcceso As Long
    buffer(1 To MAX_BUFFERS) As DirectSoundSecondaryBuffer8

End Type

Private mBuffers()       As WAV_DB_ENTRY

Private mDS              As DirectSound8

Private mMaxEntries      As Integer

Private mCantidadBuffers As Integer

Private WavPath          As String

Private Declare Sub CopyMem Lib "kernel32" Alias "RtlCopyMemory" (destination As Any, source As Any, ByVal Length As Long)

Private Sub Class_Initialize()
    '**************************************************************
    'Author: Juan Martín Sotuyo Dodero
    'Modified by Augusto José Rando
    'Last Modify Date: 5/04/2005
    '
    '**************************************************************

End Sub

Private Sub Class_Terminate()
    
    On Error GoTo Class_Terminate_Err
    

    '**************************************************************
    'Author: Juan Martín Sotuyo Dodero
    'Last Modify Date: 5/04/2005
    '
    '**************************************************************
    

    Dim i As Long

    Dim j As Long
    
    For i = 1 To mCantidadBuffers
        For j = 1 To MAX_BUFFERS
            Set mBuffers(i).buffer(j) = Nothing
        Next j
    Next i
    
    Erase mBuffers
    
    
    Exit Sub

Class_Terminate_Err:
    Call RegistrarError(Err.number, Err.Description, "CBufferMan.Class_Terminate", Erl)
    Resume Next
    
End Sub

Public Property Let Path(ByVal inString As String)
    '**************************************************************
    'Author: Juan Martín Sotuyo Dodero
    'Last Modify Date: 5/04/2005
    '
    '**************************************************************
    
    On Error GoTo Path_Err
    
    WavPath = inString

    
    Exit Property

Path_Err:
    Call RegistrarError(Err.number, Err.Description, "CBufferMan.Path", Erl)
    Resume Next
    
End Property

Private Function GetBuffer(ByVal FileName As Integer) As DirectSoundSecondaryBuffer8
    '**************************************************************
    'Author: Juan Martín Sotuyo Dodero
    'Modified by Augusto José Rando
    'Last Modify Date: 9/05/2005
    '
    '**************************************************************
    
    On Error GoTo GetBuffer_Err
    

    Dim IndiceObtenido As Integer, i As Integer

    IndiceObtenido = ObtenerIndice(FileName)
    i = 1
    
    If IndiceObtenido > 0 Then
        
        'Do While ((Not mBuffers(IndiceObtenido).buffer(i) Is Nothing) And (mBuffers(IndiceObtenido).buffer(i).GetStatus <> DSBSTATUS_PLAYING And mBuffers(IndiceObtenido).buffer(i).GetStatus <> DSBSTATUS_LOOPING + DSBSTATUS_PLAYING)) Or (i >= MAX_BUFFERS)
        '    If mBuffers(IndiceObtenido).buffer(i) Is Nothing Then
        '        If i = 1 Then
        '            'No debería entrar acá pero bueno, por las dudas
        '            If CreateBufferFromFile(FileName, IndiceObtenido, 1) Then
        '                Set GetBuffer = mBuffers(IndiceObtenido).buffer(1)
        '            End If
        '        Else
        '            Set mBuffers(IndiceObtenido).buffer(i) = mDS.DuplicateSoundBuffer(ByVal mBuffers(IndiceObtenido).buffer(1))
        '            mBuffers(IndiceObtenido).buffer(i).Stop
        '            mBuffers(IndiceObtenido).buffer(i).SetCurrentPosition 0
        '        End If
        '    Else
        '        i = i + 1
        '    End If
        'Loop
    
        If Not mBuffers(IndiceObtenido).buffer(1) Is Nothing Then
            If mBuffers(IndiceObtenido).buffer(1).GetStatus = DSBSTATUS_PLAYING Then
                If mBuffers(IndiceObtenido).buffer(2) Is Nothing Then
                    Set mBuffers(IndiceObtenido).buffer(2) = mDS.DuplicateSoundBuffer(ByVal mBuffers(IndiceObtenido).buffer(1))
                    Set GetBuffer = mBuffers(IndiceObtenido).buffer(2)
                Else

                    If mBuffers(IndiceObtenido).buffer(2).GetStatus = DSBSTATUS_PLAYING Then
                        If mBuffers(IndiceObtenido).buffer(3) Is Nothing Then
                            Set mBuffers(IndiceObtenido).buffer(3) = mDS.DuplicateSoundBuffer(ByVal mBuffers(IndiceObtenido).buffer(1))
                            Set GetBuffer = mBuffers(IndiceObtenido).buffer(3)
                        Else
                            Set GetBuffer = mBuffers(IndiceObtenido).buffer(3)

                        End If

                    Else
                        Set GetBuffer = mBuffers(IndiceObtenido).buffer(2)

                    End If

                End If

            Else
                'Devuelvo un buffer con el sonido cargado
                Set GetBuffer = mBuffers(IndiceObtenido).buffer(1)

            End If

        Else

            If CreateBufferFromFile(FileName, IndiceObtenido, 1) Then
                Set GetBuffer = mBuffers(IndiceObtenido).buffer(1)

            End If

        End If
        
        'Devolvemos el buffer SIEMPRE, aunque hayamos llegado al máximo de buffers simultáneos
        'Set GetBuffer = mBuffers(IndiceObtenido).buffer(i)
        
        'Ultimo acceso
        mBuffers(IndiceObtenido).UltimoAcceso = GetTickCount()
            
    Else    'Sonido no cargado
        GoTo CrearNuevoBuffer

    End If
    
    Exit Function
    
CrearNuevoBuffer:

    'Vemos si puedo agregar uno a la lista
    If mMaxEntries = mCantidadBuffers Then
        'Sacamos el que hace más que no usamos, y utilizamos el slot
        IndiceObtenido = CrearBuffer(FileName, BorraMenosUsado())
        Set GetBuffer = mBuffers(IndiceObtenido).buffer(1)
    Else
        'Agrego un buffer nueva a la lista
        IndiceObtenido = CrearBuffer(FileName)
        Set GetBuffer = mBuffers(IndiceObtenido).buffer(1)

    End If
    
    
    Exit Function

GetBuffer_Err:
    Call RegistrarError(Err.number, Err.Description, "CBufferMan.GetBuffer", Erl)
    Resume Next
    
End Function

Private Function ObtenerIndice(ByVal FileName As Integer) As Integer
    
    On Error GoTo ObtenerIndice_Err
    

    '**************************************************************
    'Author: Juan Martín Sotuyo Dodero
    'Last Modify Date: 5/04/2005
    'Busqueda binaria para hallar el buffer deseado
    '**************************************************************
    Dim max As Integer  'Max index

    Dim min As Integer  'Min index

    Dim mid As Integer  'Middle index
    
    min = 1
    max = mCantidadBuffers

    Do While min <= max
        mid = (min + max) / 2

        If FileName < mBuffers(mid).FileName Then

            'El índice no existe
            If max = mid Then
                max = max - 1
            Else
                max = mid

            End If

        ElseIf FileName > mBuffers(mid).FileName Then

            'El índice no existe
            If min = mid Then
                min = min + 1
            Else
                min = mid

            End If

        Else
            ObtenerIndice = mid
            Exit Function

        End If

    Loop

    
    Exit Function

ObtenerIndice_Err:
    Call RegistrarError(Err.number, Err.Description, "CBufferMan.ObtenerIndice", Erl)
    Resume Next
    
End Function

Public Function Init(ByRef DS8 As DirectSound8, ByVal MaxEntries As Integer) As Boolean
    '**************************************************************
    'Author: Juan Martín Sotuyo Dodero
    'Last Modify Date: 5/04/2005
    '
    '**************************************************************
    
    On Error GoTo Init_Err
    
    
    mMaxEntries = MaxEntries
    
    If mMaxEntries < 1 Then 'por lo menos 1 sonido
        Exit Function

    End If
    
    mCantidadBuffers = 0
    
    'Seteamos el objeto
    Set mDS = DS8
    
    Init = True

    
    Exit Function

Init_Err:
    Call RegistrarError(Err.number, Err.Description, "CBufferMan.Init", Erl)
    Resume Next
    
End Function

Public Sub BorrarTodo()
    
    On Error GoTo BorrarTodo_Err
    

    '**************************************************************
    'Author: Juan Martín Sotuyo Dodero
    'Last Modify Date: 5/04/2005
    'Vacia la lista de sonidos
    '**************************************************************
    Dim i As Long, j As Long
    
    For i = 1 To mCantidadBuffers
        For j = 1 To MAX_BUFFERS
            Set mBuffers(i).buffer(j) = Nothing
        Next j
    Next i
    
    ReDim mBuffers(0)
    mCantidadBuffers = 0

    
    Exit Sub

BorrarTodo_Err:
    Call RegistrarError(Err.number, Err.Description, "CBufferMan.BorrarTodo", Erl)
    Resume Next
    
End Sub

Public Sub DetenerTodo()
    
    On Error GoTo DetenerTodo_Err
    

    

    '**************************************************************
    'Author: Augusto José Rando
    'Last Modify Date: 5/04/2005
    'Detiene todos los sonidos
    '**************************************************************
    Dim i As Long

    Dim j As Long
    
    For i = 1 To mCantidadBuffers
        For j = 1 To MAX_BUFFERS

            If Not mBuffers(i).buffer(j) Is Nothing Then
                If (mBuffers(i).buffer(j).GetStatus = DSBSTATUS_PLAYING + DSBSTATUS_LOOPING) Or (mBuffers(i).buffer(j).GetStatus = DSBSTATUS_PLAYING) Then
                    mBuffers(i).buffer(j).Stop
                    mBuffers(i).buffer(j).SetCurrentPosition 0

                End If

            End If

        Next j
    Next i
    
    
    Exit Sub

DetenerTodo_Err:
    Call RegistrarError(Err.number, Err.Description, "CBufferMan.DetenerTodo", Erl)
    Resume Next
    
End Sub

Public Sub ModificarVolumenTodo(ByVal Volume As Long)
    
    On Error GoTo ModificarVolumenTodo_Err
    

    '**************************************************************
    'Author: Augusto José Rando
    'Last Modify Date: 5/04/2005
    'Detiene todos los sonidos
    '**************************************************************
    

    Dim i As Long

    Dim j As Long
    
    For i = 1 To mCantidadBuffers
        For j = 1 To MAX_BUFFERS

            If Not mBuffers(i).buffer(j) Is Nothing Then
                If (mBuffers(i).buffer(j).GetStatus = DSBSTATUS_LOOPING + DSBSTATUS_PLAYING) Or (mBuffers(i).buffer(j).GetStatus = DSBSTATUS_PLAYING) Then
                    mBuffers(i).buffer(j).SetVolume Volume

                End If

            End If

        Next j
    Next i
    
    
    Exit Sub

ModificarVolumenTodo_Err:
    Call RegistrarError(Err.number, Err.Description, "CBufferMan.ModificarVolumenTodo", Erl)
    Resume Next
    
End Sub

Public Function Borrar(ByVal FileName As Integer) As Integer
    
    On Error GoTo Borrar_Err
    

    '**************************************************************
    'Author: Juan Martín Sotuyo Dodero
    'Last Modify Date: 5/04/2005
    'Borra un sonido
    '**************************************************************
    Dim indice As Long
    
    'Obtenemos el ìndice
    indice = ObtenerIndice(FileName)
    
    'Lo eliminamos
    With mBuffers(indice)
        Erase .buffer

    End With
    
    'Actualizamos el número de sonidos
    mCantidadBuffers = mCantidadBuffers - 1
    
    'Movemos para atrás el resto de la lista
    For indice = indice To mCantidadBuffers
        mBuffers(indice) = mBuffers(indice + 1)
    Next indice
    
    'Redimencionamos la lista
    ReDim Preserve mBuffers(1 To mCantidadBuffers)

    
    Exit Function

Borrar_Err:
    Call RegistrarError(Err.number, Err.Description, "CBufferMan.Borrar", Erl)
    Resume Next
    
End Function

Private Function CrearBuffer(ByVal Archivo As Integer, Optional ByVal Index As Integer = -1, Optional ByVal Extraido As Boolean = False) As Integer

    '**************************************************************
    'Author: Juan Martín Sotuyo Dodero
    'Last Modify Date: 5/04/2005
    '
    '**************************************************************
    On Error GoTo ErrHandler
    
    If Index = -1 Then
        'Agregamos al final de la lista
        Index = mCantidadBuffers + 1
        ReDim Preserve mBuffers(1 To Index) As WAV_DB_ENTRY

    End If
                
    With mBuffers(Index)
        'Nombre
        .FileName = Archivo
        
        'Ultimo acceso
        .UltimoAcceso = GetTickCount()
        
        Call CreateBufferFromFile(Archivo, Index, 1, Extraido)

    End With
    
    'Aumentamos la cantidad de sonidos
    mCantidadBuffers = mCantidadBuffers + 1
    
    'Ordenamos la lista (al estar agregando de a un item no es mucho trabajo, y se compensa con la velocidad de busqueda)
    Call OrdenarBuffers(1, mCantidadBuffers)
    
    'Devolvemos el ìndice en que lo cargamos
    CrearBuffer = ObtenerIndice(Archivo)
    Exit Function

ErrHandler:
    
End Function

Private Function CreateBufferFromFile(ByVal file_name As Integer, ByVal buffer_index, ByVal buffer_sub_index As Integer, Optional ByVal Extraido As Boolean = False) As Boolean
    
    On Error GoTo CreateBufferFromFile_Err
    

    Dim dsbd As DSBUFFERDESC

    dsbd.lFlags = DSBCAPS_CTRLFREQUENCY Or DSBCAPS_CTRLPAN Or DSBCAPS_CTRLVOLUME
    dsbd.fxFormat.nFormatTag = WAVE_FORMAT_PCM
    dsbd.fxFormat.nChannels = 2
    dsbd.fxFormat.lSamplesPerSec = 22050
    dsbd.fxFormat.nBitsPerSample = 16
    dsbd.fxFormat.nBlockAlign = dsbd.fxFormat.nBitsPerSample / 8 * dsbd.fxFormat.nChannels
    dsbd.fxFormat.lAvgBytesPerSec = dsbd.fxFormat.lSamplesPerSec * dsbd.fxFormat.nBlockAlign
    
    #If Compresion Then

        If Extraido = False Then
            If Extract_File(wav, App.Path & "\..\Recursos\OUTPUT\", file_name & ".wav", Windows_Temp_Dir, ResourcesPassword, False) Then
                Set mBuffers(buffer_index).buffer(buffer_sub_index) = mDS.CreateSoundBufferFromFile(Windows_Temp_Dir & file_name & ".wav", dsbd)
                Delete_File Windows_Temp_Dir & file_name & ".wav"
                CreateBufferFromFile = True

            End If

        Else
            Set mBuffers(buffer_index).buffer(buffer_sub_index) = mDS.CreateSoundBufferFromFile(Windows_Temp_Dir & file_name & ".wav", dsbd)
            Delete_File Windows_Temp_Dir & file_name & ".wav"
            CreateBufferFromFile = True

        End If

    #Else

        Dim PathToFile As String

        PathToFile = App.Path & "\..\Recursos\WAV\" & file_name & ".wav"
    
        If FileExist(PathToFile, vbArchive) Then
            Set mBuffers(buffer_index).buffer(buffer_sub_index) = mDS.CreateSoundBufferFromFile(PathToFile, dsbd)
            CreateBufferFromFile = True

        End If

    #End If

ErrHandler:
    'Ocurrió un error, la causa más probable es que el archivo no exista
    ' If General_File_Exists(Windows_Temp_Dir & file_name & ".wav", vbNormal) Then
    '    Delete_File Windows_Temp_Dir & file_name & ".wav"
    'End If

    
    Exit Function

CreateBufferFromFile_Err:
    Call RegistrarError(Err.number, Err.Description, "CBufferMan.CreateBufferFromFile", Erl)
    Resume Next
    
End Function

Private Sub OrdenarBuffers(ByVal primero As Integer, ByVal ultimo As Integer)
    
    On Error GoTo OrdenarBuffers_Err
    

    '**************************************************************
    'Author: Juan Martín Sotuyo Dodero
    'Last Modify Date: 5/04/2005
    'Ordenamos los sonidos por nombre usando QuickSort
    '**************************************************************
    Dim min  As Integer      'Primer item de la lista

    Dim max  As Integer      'Ultimo item de la lista

    Dim comp As Integer     'Item usado para comparar

    Dim temp As WAV_DB_ENTRY
    
    min = primero
    max = ultimo
    
    comp = mBuffers((min + max) / 2).FileName
    
    Do While min <= max
        Do While mBuffers(min).FileName < comp And min < ultimo
            min = min + 1
        Loop

        Do While mBuffers(max).FileName > comp And max > primero
            max = max - 1
        Loop

        If min <= max Then
            temp = mBuffers(min)
            mBuffers(min) = mBuffers(max)
            mBuffers(max) = temp
            min = min + 1
            max = max - 1

        End If

    Loop

    If primero < max Then OrdenarBuffers primero, max
    If min < ultimo Then OrdenarBuffers min, ultimo

    
    Exit Sub

OrdenarBuffers_Err:
    Call RegistrarError(Err.number, Err.Description, "CBufferMan.OrdenarBuffers", Erl)
    Resume Next
    
End Sub

Private Function BorraMenosUsado() As Integer
    
    On Error GoTo BorraMenosUsado_Err
    

    '**************************************************************
    'Author: Juan Martín Sotuyo Dodero
    'Last Modify Date: 5/04/2005
    'Borra el sonido menos usado. Devuelve el index del borrado para ser reutilizado
    'NO redimenciona la lista, estamos forzando a sacar uno para meter otro en su lugar
    '**************************************************************
    Dim Valor As Long

    Dim i     As Long
    
    'Inicializamos todo
    Valor = mBuffers(1).UltimoAcceso
    BorraMenosUsado = 1
    
    'Buscamos cual es el que lleva más tiempo sin ser utilizado
    For i = 1 To mCantidadBuffers

        If mBuffers(i).UltimoAcceso < Valor And (Not PlayingCheck(0, i)) Then
            Valor = mBuffers(i).UltimoAcceso
            BorraMenosUsado = i

        End If

    Next i
    
    'Disminuimos el contador
    mCantidadBuffers = mCantidadBuffers - 1
    
    'Borramos los buffers
    Erase mBuffers(BorraMenosUsado).buffer
    
    
    Exit Function

BorraMenosUsado_Err:
    Call RegistrarError(Err.number, Err.Description, "CBufferMan.BorraMenosUsado", Erl)
    Resume Next
    
End Function

Public Property Get MaxEntries() As Integer
    '**************************************************************
    'Author: Juan Martín Sotuyo Dodero
    'Last Modify Date: 5/04/2005
    '
    '**************************************************************
    
    On Error GoTo MaxEntries_Err
    
    MaxEntries = mMaxEntries

    
    Exit Property

MaxEntries_Err:
    Call RegistrarError(Err.number, Err.Description, "CBufferMan.MaxEntries", Erl)
    Resume Next
    
End Property

Public Property Let MaxEntries(ByVal vNewValue As Integer)
    '**************************************************************
    'Author: Juan Martín Sotuyo Dodero
    'Last Modify Date: 5/04/2005
    '
    '**************************************************************
    
    On Error GoTo MaxEntries_Err
    
    mMaxEntries = vNewValue

    
    Exit Property

MaxEntries_Err:
    Call RegistrarError(Err.number, Err.Description, "CBufferMan.MaxEntries", Erl)
    Resume Next
    
End Property

Public Property Get CantidadBuffers() As Integer
    '**************************************************************
    'Author: Juan Martín Sotuyo Dodero
    'Last Modify Date: 5/04/2005
    '
    '**************************************************************
    
    On Error GoTo CantidadBuffers_Err
    
    CantidadBuffers = mCantidadBuffers

    
    Exit Property

CantidadBuffers_Err:
    Call RegistrarError(Err.number, Err.Description, "CBufferMan.CantidadBuffers", Erl)
    Resume Next
    
End Property

Public Sub BorraTimerProc()
    
    On Error GoTo BorraTimerProc_Err
    

    '**************************************************************
    'Author: Juan Martín Sotuyo Dodero
    'Last Modify Date: 5/04/2005
    'Cada un minuto revisamos que buffers se pueden eliminar
    '**************************************************************
    
    
    Dim LoopC        As Long

    Dim offset       As Integer

    Dim TiempoActual As Long
    
    TiempoActual = GetTickCount()
    
    For LoopC = 1 To mCantidadBuffers

        'Revisar cuales usamos
        If mBuffers(LoopC).UltimoAcceso > TiempoActual - 60000 Then

            'Si el offset está seteado, movemos para atrás la lista
            If offset Then
                mBuffers(LoopC - offset) = mBuffers(LoopC)

            End If

        Else

            If PlayingCheck(0, LoopC) = False Then
                'Eliminar de memoria
                Erase mBuffers(LoopC).buffer
                offset = offset + 1

            End If

        End If

    Next LoopC
    
    'Actualizar el número de sonidos
    mCantidadBuffers = mCantidadBuffers - offset

    
    Exit Sub

BorraTimerProc_Err:
    Call RegistrarError(Err.number, Err.Description, "CBufferMan.BorraTimerProc", Erl)
    Resume Next
    
End Sub

Public Function CargarBuffer(ByVal FileName As Integer) As Boolean
    '**************************************************************
    'Author: Augusto José Rando
    'Last Modify Date: 17/08/2005
    '
    '**************************************************************
    
    On Error GoTo CargarBuffer_Err
    

    Dim Index As Integer

    If ObtenerIndice(FileName) <> 0 Then
        CargarBuffer = True
    Else    'Sonido no cargado

        'Vemos si puedo agregar uno a la lista
        If mMaxEntries = mCantidadBuffers Then
            'Sacamos el que hace más que no usamos, y utilizamos el slot
            Call CrearBuffer(FileName, BorraMenosUsado())
            CargarBuffer = True
        Else
            'Agrego un buffer nueva a la lista
            Call CrearBuffer(FileName)
            CargarBuffer = True

        End If

    End If

    
    Exit Function

CargarBuffer_Err:
    Call RegistrarError(Err.number, Err.Description, "CBufferMan.CargarBuffer", Erl)
    Resume Next
    
End Function

Public Function Reproducir(ByVal file_name As Integer, play_format As CONST_DSBPLAYFLAGS, ByVal Volume As Long, ByVal pan As Long) As Boolean
    '**************************************************************
    'Author: Augusto José Rando
    'Last Modify Date: 17/08/2005
    '
    '**************************************************************

    Dim Sound_Buffer As DirectSoundSecondaryBuffer8

    On Error GoTo ErrorHandler

    Set Sound_Buffer = GetBuffer(file_name)

    If Sound_Buffer Is Nothing Then Exit Function
    
    'Esto solo debería pasar si ya hay MAX_BUFFERS en reproducción
    If Sound_Buffer.GetStatus = DSBSTATUS_PLAYING Then
        Sound_Buffer.Stop
        Sound_Buffer.SetCurrentPosition 0

    End If
    
    Sound_Buffer.SetPan pan
    Sound_Buffer.SetVolume Volume
    Sound_Buffer.Play play_format
    Reproducir = True

    Exit Function

ErrorHandler:
    Reproducir = False

End Function

Public Function Detener(ByVal file_name As Integer) As Boolean
    '**************************************************************
    'Author: Augusto José Rando
    'Last Modify Date: 17/08/2005
    '
    '**************************************************************

    On Error GoTo ErrorHandler

    Dim IndiceObtenido As Integer, i As Long

    IndiceObtenido = ObtenerIndice(file_name)

    If IndiceObtenido <= 0 Then Exit Function
    
    For i = 1 To MAX_BUFFERS

        If Not mBuffers(IndiceObtenido).buffer(i) Is Nothing Then
            mBuffers(IndiceObtenido).buffer(i).Stop
            mBuffers(IndiceObtenido).buffer(i).SetCurrentPosition 0

        End If

    Next i
    
    Detener = True

    Exit Function

ErrorHandler:
    Detener = False

End Function

Public Function CambiarVolumen(ByVal file_name As Integer, ByVal volumen As Long) As Boolean
    '**************************************************************
    'Author: Augusto José Rando
    'Last Modify Date: 17/08/2005
    'Returns true if successfully changed volume of a playing buffer
    '**************************************************************

    On Error GoTo ErrorHandler

    Dim IndiceObtenido As Integer, i As Long, bFlag As Boolean

    IndiceObtenido = ObtenerIndice(file_name)

    If IndiceObtenido <= 0 Then Exit Function
    
    For i = 1 To MAX_BUFFERS

        If Not mBuffers(IndiceObtenido).buffer(i) Is Nothing Then
            mBuffers(IndiceObtenido).buffer(i).SetVolume volumen

            If (mBuffers(IndiceObtenido).buffer(i).GetStatus = DSBSTATUS_PLAYING Or mBuffers(IndiceObtenido).buffer(i).GetStatus = DSBSTATUS_PLAYING + DSBSTATUS_LOOPING) Then bFlag = True

        End If

    Next i
    
    CambiarVolumen = bFlag

    Exit Function

ErrorHandler:
    CambiarVolumen = False

End Function

Public Function CambiarPan(ByVal file_name As Integer, ByVal pan As Long) As Boolean
    '**************************************************************
    'Author: Augusto José Rando
    'Last Modify Date: 17/08/2005
    'Returns true if successfully changed pan of a playing buffer
    '**************************************************************

    On Error GoTo ErrorHandler

    Dim IndiceObtenido As Integer, i As Long, bFlag As Boolean

    IndiceObtenido = ObtenerIndice(file_name)

    If IndiceObtenido <= 0 Then Exit Function
    
    For i = 1 To MAX_BUFFERS

        If Not mBuffers(IndiceObtenido).buffer(i) Is Nothing Then
            mBuffers(IndiceObtenido).buffer(i).SetPan pan

            If (mBuffers(IndiceObtenido).buffer(i).GetStatus = DSBSTATUS_PLAYING Or mBuffers(IndiceObtenido).buffer(i).GetStatus = DSBSTATUS_PLAYING + DSBSTATUS_LOOPING) Then bFlag = True

        End If

    Next i
    
    CambiarPan = bFlag

    Exit Function

ErrorHandler:
    CambiarPan = False

End Function

Public Function PlayingCheck(ByVal file_name As Integer, Optional ByVal IndiceObtenido As Integer = -1) As Boolean

    '**************************************************************
    'Author: Augusto José Rando
    'Last Modify Date: 17/08/2005
    'Returns true if playing
    '**************************************************************

    On Error GoTo ErrorHandler

    Dim i As Long, bFlag As Boolean

    If IndiceObtenido = -1 Then
        IndiceObtenido = ObtenerIndice(file_name)

        If IndiceObtenido <= 0 Then Exit Function

    End If
    
    For i = 1 To MAX_BUFFERS

        If Not mBuffers(IndiceObtenido).buffer(i) Is Nothing Then
            If (mBuffers(IndiceObtenido).buffer(i).GetStatus = DSBSTATUS_PLAYING + DSBSTATUS_LOOPING) Or (mBuffers(IndiceObtenido).buffer(i).GetStatus = DSBSTATUS_PLAYING) Then
                PlayingCheck = True
                Exit Function

            End If

        End If

    Next i

    Exit Function

ErrorHandler:
    PlayingCheck = False

End Function

