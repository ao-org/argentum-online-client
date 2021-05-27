VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsByteQueue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : clsByteQueue
' Author    : Mateo Miccino (Kuruk), con ayuda de Brian Sabatier [/About]
' Date      : 23/07/2020
' Purpose   : El protocolo ahora evita descoordinaciones de protocolo (se puede monitorear, activando TEST = 1 y BYTE_QUEUE_DEBUG = 1) tanto en el server
'             como el storage manager y el cliente (los tres deben tener las mismas opciones activadas).
'             Este mismo módulo debe estar coordinado siempre en Servidor, SQlManager y Cliente.
'             Hay un sistema de validación CRC por cada paquete, evitando paquetes basura, ataques, sniffers y etc.
'---------------------------------------------------------------------------------------

'About: clsByteQueue de 0.4.0 Mateo.
Option Explicit
Option Base 0                  'It's the default, but we make it explicit just in case...

' The error number thrown when there is not enough data in

' the buffer to read the specified data type.
' It's 9 (subscript out of range) + the object error constant
Private Const NOT_ENOUGH_DATA  As Long = vbObjectError + 9

' The error number thrown when there is not enough space in
' the buffer to write.
Private Const NOT_ENOUGH_SPACE As Long = vbObjectError + 10

' Default size of a data buffer (10 Kbs)

' @see Class_Initialize
Private Const DATA_BUFFER      As Long = 358400    '358400    '350kb
Private CUSTOM_DATA_BUFFER     As Long

Private PacketStartPosition    As Long
Private LastLength             As Long
Private packetLength           As Long

' The byte data
Public errNumber               As Long

Dim DATA()                     As Byte

'Posicion de lectura
Private queuePos               As Long

' How big the data array is
Dim queueCapacity              As Long

' How far into the data array have we written
Dim queueLength                As Long

#If BYTE_QUEUE_DEBUG = 1 Then

    Public Enum eWriteTypes

        wrBoolean
        wrByte
        wrInteger
        wrLong
        wrASCIIString
        wrASCIStringFixed
        wrID
        wrSingle
        wrDouble

    End Enum

#End If

' CopyMemory is the fastest way to copy memory blocks, so we abuse of it
' @param destination Where the data will be copied.
' @param source The data to be copied.
' @param length Number of bytes to be copied.

Private Declare Sub CopyMemory _
                Lib "kernel32" _
                Alias "RtlMoveMemory" (ByRef destination As Any, _
                                       ByRef source As Any, _
                                       ByVal Length As Long)

' Initializes the queue with the default queueCapacity
' @see DATA_BUFFER

Private Sub Class_Initialize()
    '***************************************************
    'Autor: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 04/27/06
    'Initializes the queue with the default queueCapacity
    '***************************************************
    ReDim DATA(DATA_BUFFER - 1) As Byte

    queueCapacity = DATA_BUFFER
    CUSTOM_DATA_BUFFER = DATA_BUFFER
    PacketStartPosition = -1

End Sub

' Clean up and release resources

Private Sub Class_Terminate()
    '***************************************************
    'Autor: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 04/27/06
    'Clean up
    '***************************************************
    Erase DATA

End Sub

' Copies another ByteQueue's data into this object.
'
' @param source The ByteQueue whose buffer will eb copied.
' @remarks  This method will resize the ByteQueue's buffer to match
'           the source. All previous data on this object will be lost.

Public Sub CopyBuffer(ByRef source As clsByteQueue)

    '***************************************************
    'Autor: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 04/27/06
    'A Visual Basic equivalent of a Copy Contructor
    '***************************************************
    If source.Length = 0 Then
        'Clear the list and exit
        Call RemoveData(Length)

        Exit Sub

    End If

    ' Set capacity and resize array - make sure all data is lost
    queueCapacity = source.Capacity
    CUSTOM_DATA_BUFFER = queueCapacity

    ReDim DATA(queueCapacity - 1) As Byte

    ' Read buffer
    Dim buf() As Byte

    ReDim buf(source.Length - 1) As Byte

    Call source.PeekBlock(buf, source.Length)

    queueLength = 0
    queuePos = 0
    ' Write buffer
    Call WriteBlock(buf, source.Length)

End Sub

' Returns the smaller of Value1 and Value2
'
' @param Value1 First value to compare
' @param Value2 Second Value to compare
' @return   The smaller of Value1 and Value2
' @remarks  This method is faster than Iif() and cleaner, therefore it's used instead of it

Private Function Min(ByVal Value1 As Long, ByVal Value2 As Long) As Long

    '***************************************************
    'Autor: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 04/27/06
    'It's faster than iif and I like it better
    '***************************************************
    If Value1 < Value2 Then
        Min = Value1
    Else
        Min = Value2

    End If

End Function

' Writes a byte array at the end of the byte queue if there is enough space.
' Otherwise it throws NOT_ENOUGH_DATA.
'
' @param buf Byte array containing the data to be copied. MUST have 0 as the first index.
' @param datalength Total number of elements in the array
' @return   The actual number of bytes copied
' @remarks  buf MUST be Base 0
' @see RemoveData
' @see ReadData
' @see NOT_ENOUGH_DATA

Private Function WriteData(ByRef buf() As Byte, ByVal DataLength As Long) As Long
    '***************************************************
    'Autor: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 04/27/06
    'If the queueCapacity allows it copyes a byte buffer to the queue, if not it throws NOT_ENOUGH_DATA
    '***************************************************

    'Check if there is enough free space
    If queueCapacity - queueLength - DataLength < 0 Then
        Debug.Print "TCP Error NOT_ENOUGH_SPACE"
        Call ErrRaise(NOT_ENOUGH_SPACE)
        Call Clear

        Exit Function

    End If

    'Copy data from buffer
    Call CopyMemory(DATA(queueLength), buf(0), DataLength)

    'Update length of data
    queueLength = queueLength + DataLength
    WriteData = DataLength

End Function

' Reads a byte array from the beginning of the byte queue if there is enough data available.
' Otherwise it throws NOT_ENOUGH_DATA.
'
' @param buf Byte array where to copy the data. MUST have 0 as the first index and already be sized properly.
' @param datalength Total number of elements in the array
' @return   The actual number of bytes copied
' @remarks  buf MUST be Base 0 and be already resized to be able to contain the requested bytes.
' This method performs no checks of such things as being a private method it's supposed that the consistency of the module is to be kept.
' If there is not enough data available it will read all available data.
' @see WriteData
' @see RemoveData
' @see NOT_ENOUGH_DATA

Private Function ReadData(ByRef buf() As Byte, ByVal DataLength As Long) As Long

    ''A on error resume next
    '***************************************************
    'Autor: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 04/27/06
    'If enough memory is available, it copies the requested number of bytes to the buffer
    '***************************************************
    'Check if we can read the number of bytes requested

    'Debug.Print "REadData Length: " & dataLength
    If DataLength > queueLength Then
        Call ErrRaise(NOT_ENOUGH_DATA)

        Exit Function

    End If

    'Copy data to buffer
    Call CopyMemory(buf(0), DATA(queuePos), DataLength)

    ReadData = DataLength

End Function

' Removes a given number of bytes from the beginning of the byte queue.
' If there is less data available than the requested amount it removes all data.
'
' @param datalength Total number of bytes to remove
' @return   The actual number of bytes removed
' @see WriteData
' @see ReadData

Private Function RemoveData(ByVal DataLength As Long) As Long
    '***************************************************
    'Autor: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 04/27/06
    'Moves forward the queue overwriting the first dataLength bytes
    '***************************************************
    RemoveData = Min(DataLength, queueLength)

    queuePos = queuePos + DataLength

End Function

''
' Writes a single byte at the end of the queue
'
' @param value The value to be written
' @return   The number of bytes written
' @see PeekByte
' @see ReadByte
#If BYTE_QUEUE_DEBUG = 1 Then
Public Sub WriteType(ByVal value As eWriteTypes)
'***************************************************
'Autor: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 04/27/06
'Writes a byte to the queue
'***************************************************

'Debug.Print "Bytes: " & value

    If queueCapacity - queueLength - 1 < 0 Then
        Call ErrRaise(NOT_ENOUGH_SPACE)
        Exit Sub
    End If
    DATA(queueLength) = value
    queueLength = queueLength + 1

End Sub

#End If
Public Function WriteByte(ByVal value As Byte) As Long
    '***************************************************
    'Autor: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 04/27/06
    'Writes a byte to the queue
    '***************************************************

    'Debug.Print "Bytes: " & value
    #If BYTE_QUEUE_DEBUG = 1 Then
        Call WriteType(wrByte)
    #End If

    If queueCapacity - queueLength - 1 < 0 Then
        Call ErrRaise(NOT_ENOUGH_SPACE)

        Exit Function

    End If

    DATA(queueLength) = value
    queueLength = queueLength + 1
    WriteByte = 1

End Function

''
' Writes an integer at the end of the queue
'
' @param value The value to be written
' @return   The number of bytes written
' @see PeekInteger
' @see ReadInteger

Public Function WriteInteger(ByVal value As Integer) As Long
    '***************************************************
    'Autor: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 04/27/06
    'Writes an integer to the queue
    '***************************************************
    'Debug.Print "Integers: " & value
    #If BYTE_QUEUE_DEBUG = 1 Then
        Call WriteType(wrInteger)
    #End If

    If queueCapacity - queueLength - 2 < 0 Then
        Call ErrRaise(NOT_ENOUGH_SPACE)

        Exit Function

    End If

    CopyMemory DATA(queueLength), value, 2
    queueLength = queueLength + 2
    WriteInteger = 2

End Function

''
' Writes a long at the end of the queue
'
' @param value The value to be written
' @return   The number of bytes written
' @see PeekLong
' @see ReadLong

Public Function WriteLong(ByVal value As Long) As Long
    '***************************************************
    'Autor: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 04/27/06
    'Writes a long to the queue
    '***************************************************
    #If BYTE_QUEUE_DEBUG = 1 Then
        Call WriteType(wrLong)
    #End If

    Dim buf(3) As Byte

    'Debug.Print "Long: " & value
    'Copy data to temp buffer
    Call CopyMemory(buf(0), value, 4)

    WriteLong = WriteData(buf, 4)

End Function

Public Sub WriteCollection(ByRef value As Collection)

    Dim i As Long

    Call WriteLong(value.Count)

    For i = 1 To value.Count
        Call WriteASCIIString(value.Item(i))
    Next i

End Sub

Public Function ReadCollection() As Collection

    Dim i    As Long

    Dim Max  As Long

    Dim List As New Collection

    Max = ReadLong()

    For i = 1 To Max
        List.Add ReadASCIIString()
    Next i

    Set ReadCollection = List

End Function

''
' Writes a single at the end of the queue
'
' @param value The value to be written
' @return   The number of bytes written
' @see PeekSingle
' @see ReadSingle

Public Function WriteSingle(ByVal value As Single) As Long
    '***************************************************
    'Autor: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 10/28/07
    'Writes a single to the queue
    '***************************************************
    #If BYTE_QUEUE_DEBUG = 1 Then
        Call WriteType(wrSingle)
    #End If

    Dim buf(3) As Byte

    'Copy data to temp buffer
    Call CopyMemory(buf(0), value, 4)

    WriteSingle = WriteData(buf, 4)

End Function

''
' Writes a double at the end of the queue
'
' @param value The value to be written
' @return   The number of bytes written
' @see PeekDouble
' @see ReadDouble

Public Function WriteDouble(ByVal value As Double) As Long
    '***************************************************
    'Autor: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 10/28/07
    'Writes a double to the queue
    '***************************************************
    #If BYTE_QUEUE_DEBUG = 1 Then
        Call WriteType(wrDouble)
    #End If

    Dim buf(7) As Byte

    'Copy data to temp buffer
    Call CopyMemory(buf(0), value, 8)

    WriteDouble = WriteData(buf, 8)

End Function

''
' Writes a boolean value at the end of the queue
'
' @param value The value to be written
' @return   The number of bytes written
' @see PeekBoolean
' @see ReadBoolean

Public Function WriteBoolean(ByVal value As Boolean) As Long
    '***************************************************
    'Autor: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 04/27/06
    'Writes a byte to the queue
    '***************************************************
    #If BYTE_QUEUE_DEBUG = 1 Then
        Call WriteType(wrBoolean)
    #End If

    Dim buf(0) As Byte

    If value Then buf(0) = 1

    WriteBoolean = WriteData(buf, 1)

End Function

''
' Writes a fixed length ASCII string at the end of the queue
'
' @param value The string to be written
' @return   The number of bytes written
' @see PeekASCIIStringFixed
' @see ReadASCIIStringFixed

Public Function WriteASCIIStringFixed(ByVal value As String) As Long
    '***************************************************
    'Autor: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 04/27/06
    'Writes a fixed length ASCII string to the queue
    '***************************************************
    #If BYTE_QUEUE_DEBUG = 1 Then
        Call WriteType(wrASCIStringFixed)
    #End If

    Dim buf() As Byte

    Dim i     As Integer

    i = Len(value)
    ReDim buf(i - 1) As Byte
    'Copy data to temp buffer
    Call CopyMemory(buf(0), ByVal StrPtr(StrConv(value, vbFromUnicode)), i)

    WriteASCIIStringFixed = WriteData(buf, i)

End Function

''
' Writes a fixed length unicode string at the end of the queue
'
' @param value The string to be written
' @return   The number of bytes written
' @see PeekUnicodeStringFixed
' @see ReadUnicodeStringFixed

Public Function WriteUnicodeStringFixed(ByVal value As String) As Long

    '***************************************************
    'Autor: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 04/27/06
    'Writes a fixed length UNICODE string to the queue
    '***************************************************
    Dim buf() As Byte

    Dim i     As Integer

    i = LenB(value)
    ReDim buf(i) As Byte

    'Copy data to temp buffer
    Call CopyMemory(buf(0), ByVal StrPtr(value), i)

    WriteUnicodeStringFixed = WriteData(buf, i)

End Function

''
' Writes a variable length ASCII string at the end of the queue
'
' @param value The string to be written
' @return   The number of bytes written
' @see PeekASCIIString
' @see ReadASCIIString

Public Function WriteASCIIString(ByVal value As String) As Long
    '***************************************************
    'Autor: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 04/27/06
    'Writes a variable length ASCII string to the queue
    '***************************************************
    #If BYTE_QUEUE_DEBUG = 1 Then
        Call WriteType(wrASCIIString)
    #End If

    If queueCapacity - queueLength - 1 < 0 Then
        Call ErrRaise(NOT_ENOUGH_SPACE)

        Exit Function

    End If

    Dim i As Integer

    i = Len(value)
    'Copy length to temp buffer
    'value = Encriptar(value, "cryptbao")
    Call CopyMemory(DATA(queueLength), i, 2)

    If i > 0 Then
        'Copy data to temp buffer
        Call CopyMemory(DATA(queueLength + 2), ByVal StrPtr(StrConv(value, vbFromUnicode)), i)

    End If

    queueLength = queueLength + i + 2
    WriteASCIIString = i + 2

End Function

''
' Writes a variable length unicode string at the end of the queue
'
' @param value The string to be written
' @return   The number of bytes written
' @see PeekUnicodeString
' @see ReadUnicodeString

Public Function WriteUnicodeString(ByVal value As String) As Long

    '***************************************************
    'Autor: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 04/27/06
    'Writes a variable length UNICODE string to the queue
    '***************************************************
    Dim buf() As Byte

    ReDim buf(LenB(value) + 1) As Byte

    'Copy length to temp buffer
    Call CopyMemory(buf(0), CInt(Len(value)), 2)

    'If Len(value) > 0 Then
    If Not StrComp(value, vbNullString) = 0 Then    'About
        'Copy data to temp buffer
        Call CopyMemory(buf(2), ByVal StrPtr(value), LenB(value))

    End If

    WriteUnicodeString = WriteData(buf, LenB(value) + 2)

End Function

''
' Writes a byte array at the end of the queue
'
' @param value The byte array to be written. MUST be Base 0.
' @param length The number of elements to copy from the byte array. If less than 0 it will copy the whole array.
' @return   The number of bytes written
' @remarks  value() MUST be Base 0.
' @see PeekBlock
' @see ReadBlock

Public Function WriteBlock(ByRef value() As Byte, _
                           Optional ByVal Length As Long = -1) As Long

    '***************************************************
    'Autor: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 04/27/06
    'Writes a byte array to the queue
    '***************************************************
    'Prevent from copying memory outside the array
    If Length = 0 Then

        Exit Function

    End If

    If Length > UBound(value()) + 1 Or Length < 0 Then Length = UBound(value()) + 1
    WriteBlock = WriteData(value, Length)

End Function

''
' Reads a single byte from the begining of the queue and removes it
'
' @return   The read value
' @remarks  Read methods removes the data from the queue.
' Data removed can't be recovered by the queue in any way
' @see PeekByte
' @see WriteByte
#If BYTE_QUEUE_DEBUG = 1 Then
Public Sub ReadCheckType(ByVal ExpectedType As eWriteTypes)
'***************************************************
'Autor: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 04/27/06
'Reads a byte from the queue and removes it
'***************************************************
Dim buf(0)                      As Byte
    'Debug.Print "ReadByte: " & ReadByte
    'Read the data and remove it
    Call RemoveData(ReadData(buf, 1))

    'If buf(0) <> ExpectedType Then Stop
End Sub
Public Sub RestoreCheckType()
    queuePos = queuePos - 1
End Sub

#End If
Public Function ReadByte() As Byte
    '***************************************************
    'Autor: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 04/27/06
    'Reads a byte from the queue and removes it
    '***************************************************
    #If BYTE_QUEUE_DEBUG = 1 Then
        Call ReadCheckType(wrByte)
    #End If

    Dim buf(0) As Byte

    'Debug.Print "ReadByte: " & ReadByte
    'Read the data and remove it
    Call RemoveData(ReadData(buf, 1))

    ReadByte = buf(0)

End Function

''
' Reads an integer from the begining of the queue and removes it
'
' @return   The read value
' @remarks  Read methods removes the data from the queue.
' Data removed can't be recovered by the queue in any way
' @see PeekInteger
' @see WriteInteger

Public Function ReadInteger() As Integer
    '***************************************************
    'Autor: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 04/27/06
    'Reads an integer from the queue and removes it
    '***************************************************
    #If BYTE_QUEUE_DEBUG = 1 Then
        Call ReadCheckType(wrInteger)
    #End If

    Dim buf(1) As Byte

    'Debug.Print "ReadInteger: " & ReadInteger
    'Read the data and remove it
    Call RemoveData(ReadData(buf, 2))

    'Copy data to temp buffer
    Call CopyMemory(ReadInteger, buf(0), 2)

End Function

''
' Reads a long from the begining of the queue and removes it
'
' @return   The read value
' @remarks  Read methods removes the data from the queue.
' Data removed can't be recovered by the queue in any way
' @see PeekLong
' @see WriteLong

Public Function ReadLong() As Long
    '***************************************************
    'Autor: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 04/27/06
    'Reads a long from the queue and removes it
    '***************************************************
    #If BYTE_QUEUE_DEBUG = 1 Then
        Call ReadCheckType(wrLong)
    #End If

    Dim buf(3) As Byte

    'Debug.Print "ReadLong: " & ReadInteger
    'Read the data and remove it
    Call RemoveData(ReadData(buf, 4))

    'Copy data to temp buffer
    Call CopyMemory(ReadLong, buf(0), 4)

End Function

''
' Reads a single from the begining of the queue and removes it
'
' @return   The read value
' @remarks  Read methods removes the data from the queue.
' Data removed can't be recovered by the queue in any way
' @see PeekSingle
' @see WriteSingle

Public Function ReadSingle() As Single
    '***************************************************
    'Autor: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 10/28/07
    'Reads a single from the queue and removes it
    '***************************************************
    #If BYTE_QUEUE_DEBUG = 1 Then
        Call ReadCheckType(wrSingle)
    #End If

    Dim buf(3) As Byte

    'Read the data and remove it
    Call RemoveData(ReadData(buf, 4))

    'Copy data to temp buffer
    Call CopyMemory(ReadSingle, buf(0), 4)

End Function

''
' Reads a double from the begining of the queue and removes it
'
' @return   The read value
' @remarks  Read methods removes the data from the queue.
' Data removed can't be recovered by the queue in any way
' @see PeekDouble
' @see WriteDouble

Public Function ReadDouble() As Double
    '***************************************************
    'Autor: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 10/28/07
    'Reads a double from the queue and removes it
    '***************************************************
    #If BYTE_QUEUE_DEBUG = 1 Then
        Call ReadCheckType(wrDouble)
    #End If

    Dim buf(7) As Byte

    'Read the data and remove it
    Call RemoveData(ReadData(buf, 8))

    'Copy data to temp buffer
    Call CopyMemory(ReadDouble, buf(0), 8)

End Function

''
' Reads a Boolean from the begining of the queue and removes it
'
' @return   The read value
' @remarks  Read methods removes the data from the queue.
' Data removed can't be recovered by the queue in any way
' @see PeekBoolean
' @see WriteBoolean

Public Function ReadBoolean() As Boolean
    '***************************************************
    'Autor: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 04/27/06
    'Reads a Boolean from the queue and removes it
    '***************************************************
    #If BYTE_QUEUE_DEBUG = 1 Then
        Call ReadCheckType(wrBoolean)
    #End If

    Dim buf(0) As Byte

    'Read the data and remove it
    Call RemoveData(ReadData(buf, 1))

    If buf(0) = 1 Then ReadBoolean = True

End Function

''
' Reads a fixed length ASCII string from the begining of the queue and removes it
'
' @param length The length of the string to be read
' @return   The read string
' @remarks  Read methods removes the data from the queue.
' Data removed can't be recovered by the queue in any way
' If there is not enough data to read the complete string then nothing is removed and an empty string is returned
' @see PeekASCIIStringFixed
' @see WriteUnicodeStringFixed

Public Function ReadASCIIStringFixed(ByVal Length As Long) As String
    '***************************************************
    'Autor: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 04/27/06
    'Reads a fixed length ASCII string from the queue and removes it
    '***************************************************
    #If BYTE_QUEUE_DEBUG = 1 Then
        Call ReadCheckType(wrASCIStringFixed)
    #End If

    If Length <= 0 Then Exit Function

    If queueLength >= Length Then

        Dim buf() As Byte

        ReDim buf(Length - 1) As Byte

        'Read the data and remove it
        Call RemoveData(ReadData(buf, Length))

        ReadASCIIStringFixed = StrConv(buf, vbUnicode)
    Else
        Call ErrRaise(NOT_ENOUGH_DATA)

    End If

End Function

''
' Reads a fixed length unicode string from the begining of the queue and removes it
'
' @param length The length of the string to be read.
' @return   The read string if enough data is available, an empty string otherwise.
' @remarks  Read methods removes the data from the queue.
' Data removed can't be recovered by the queue in any way.
' If there is not enough data to read the complete string then nothing is removed and an empty string is returned
' @see PeekUnicodeStringFixed
' @see WriteUnicodeStringFixed

Public Function ReadUnicodeStringFixed(ByVal Length As Long) As String

    '***************************************************
    'Autor: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 04/27/06
    'Reads a fixed length UNICODE string from the queue and removes it
    '***************************************************
    If Length <= 0 Then Exit Function

    If queueLength >= Length * 2 Then

        Dim buf() As Byte

        ReDim buf(Length * 2 - 1) As Byte

        'Read the data and remove it
        Call RemoveData(ReadData(buf, Length * 2))

        ReadUnicodeStringFixed = buf
    Else
        Call ErrRaise(NOT_ENOUGH_DATA)

    End If

End Function

''
' Reads a variable length ASCII string from the begining of the queue and removes it
'
' @return   The read string
' @remarks  Read methods removes the data from the queue.
' Data removed can't be recovered by the queue in any way
' If there is not enough data to read the complete string then nothing is removed and an empty string is returned
' @see PeekASCIIString
' @see WriteASCIIString

Public Function ReadASCIIString() As String

    '***************************************************
    'Autor: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 04/27/06
    'Reads a variable length ASCII string from the queue and removes it
    '***************************************************
    Dim buf(1) As Byte

    Dim Length As Integer

    #If BYTE_QUEUE_DEBUG = 1 Then
        Call ReadCheckType(wrASCIIString)
    #End If

    'Make sure we can read a valid length
    If queueLength > 1 Then
        'Read the length
        Call ReadData(buf, 2)
        Call CopyMemory(Length, buf(0), 2)

        'Make sure there are enough bytes
        If queueLength >= Length + 2 Then
            'Remove the length
            Call RemoveData(2)

            If Length > 0 Then

                Dim buf2() As Byte

                ReDim buf2(Length - 1) As Byte

                'Read the data and remove it
                Call RemoveData(ReadData(buf2, Length))

                ReadASCIIString = StrConv(buf2, vbUnicode)

            End If

        Else
            Call ErrRaise(NOT_ENOUGH_DATA)

        End If

    Else
        Call ErrRaise(NOT_ENOUGH_DATA)

    End If

End Function

''
' Reads a variable length unicode string from the begining of the queue and removes it
'
' @return   The read string if enough data is available, an empty string otherwise.
' @remarks  Read methods removes the data from the queue.
' Data removed can't be recovered by the queue in any way
' If there is not enough data to read the complete string then nothing is removed and an empty string is returned
' @see PeekUnicodeString
' @see WriteUnicodeString

Public Function ReadUnicodeString() As String

    '***************************************************
    'Autor: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 04/27/06
    'Reads a variable length UNICODE string from the queue and removes it
    '***************************************************
    Dim buf(1) As Byte

    Dim Length As Integer

    'Make sure we can read a valid length
    If queueLength > 1 Then
        'Read the length
        Call ReadData(buf, 2)
        Call CopyMemory(Length, buf(0), 2)

        'Make sure there are enough bytes
        If queueLength >= Length * 2 + 2 Then
            'Remove the length
            Call RemoveData(2)

            Dim buf2() As Byte

            ReDim buf2(Length * 2 - 1) As Byte

            'Read the data and remove it
            Call RemoveData(ReadData(buf2, Length * 2))

            ReadUnicodeString = buf2
        Else
            Call ErrRaise(NOT_ENOUGH_DATA)

        End If

    Else
        Call ErrRaise(NOT_ENOUGH_DATA)

    End If

End Function

''
' Reads a byte array from the begining of the queue and removes it
'
' @param block Byte array which will contain the read data. MUST be Base 0 and previously resized to contain the requested amount of bytes.
' @param dataLength Number of bytes to retrieve from the queue.
' @return   The number of read bytes.
' @remarks  The block() array MUST be Base 0 and previously resized to be able to contain the requested bytes.
' Read methods removes the data from the queue.
' Data removed can't be recovered by the queue in any way
' @see PeekBlock
' @see WriteBlock

Public Function ReadBlock(ByRef Block() As Byte, ByVal DataLength As Long) As Long

    '***************************************************
    'Autor: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 04/27/06
    'Reads a byte array from the queue and removes it
    '***************************************************
    'Read the data and remove it
    If DataLength > 0 Then ReadBlock = RemoveData(ReadData(Block(), DataLength))

End Function

''
' Reads a single byte from the begining of the queue but DOES NOT remove it.
'
' @return   The read value.
' @remarks  Peek methods, unlike Read methods, don't remove the data from the queue.
' @see ReadByte
' @see WriteByte
Public Sub PeekID(ByVal ID As Byte, ByVal NumID As Byte)

    If ID <> NumID Then

    End If

End Sub

Public Function PeekByte() As Byte
    '***************************************************
    'Autor: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 04/27/06
    'Reads a byte from the queue but doesn't removes it
    '***************************************************
    #If BYTE_QUEUE_DEBUG = 1 Then
        Call ReadCheckType(wrByte)
    #End If

    If 1 > queueLength Then
        Call ErrRaise(NOT_ENOUGH_DATA)

        Exit Function

    End If

    PeekByte = DATA(queuePos)

    #If BYTE_QUEUE_DEBUG = 1 Then
        Call RestoreCheckType
    #End If

End Function

''
' Reads an integer from the begining of the queue but DOES NOT remove it.
'
' @return   The read value.
' @remarks  Peek methods, unlike Read methods, don't remove the data from the queue.
' @see ReadInteger
' @see WriteInteger

Public Function PeekInteger() As Integer

    '***************************************************
    'Autor: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 04/27/06
    'Reads an integer from the queue but doesn't removes it
    '***************************************************
    Dim buf(1) As Byte

    #If BYTE_QUEUE_DEBUG = 1 Then
        Call ReadCheckType(wrInteger)
    #End If
    'Read the data and remove it
    Call ReadData(buf, 2)

    'Copy data to temp buffer
    Call CopyMemory(PeekInteger, buf(0), 2)
    #If BYTE_QUEUE_DEBUG = 1 Then
        Call RestoreCheckType
    #End If

End Function

Public Function PeekLength() As Integer
    #If BYTE_QUEUE_DEBUG = 1 Then
        ' En Debug escribimos 1byte antes del PacketID y otro Byte antes del PacketLength
        CopyMemory PeekLength, DATA(queuePos + 3), 2
    #Else
        CopyMemory PeekLength, DATA(queuePos + 1), 2
    #End If

End Function

''
' Reads a long from the begining of the queue but DOES NOT remove it.
'
' @return   The read value.
' @remarks  Peek methods, unlike Read methods, don't remove the data from the queue.
' @see ReadLong
' @see WriteLong

Public Function PeekLong() As Long
    '***************************************************
    'Autor: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 04/27/06
    'Reads a long from the queue but doesn't removes it
    '***************************************************
    #If BYTE_QUEUE_DEBUG = 1 Then
        Call ReadCheckType(wrLong)
    #End If

    Dim buf(3) As Byte

    'Read the data and remove it
    Call ReadData(buf, 4)

    'Copy data to temp buffer
    Call CopyMemory(PeekLong, buf(0), 4)
    #If BYTE_QUEUE_DEBUG = 1 Then
        Call RestoreCheckType
    #End If

End Function

''
' Reads a single from the begining of the queue but DOES NOT remove it.
'
' @return   The read value.
' @remarks  Peek methods, unlike Read methods, don't remove the data from the queue.
' @see ReadSingle
' @see WriteSingle

Public Function PeekSingle() As Single
    '***************************************************
    'Autor: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 10/28/07
    'Reads a single from the queue but doesn't removes it
    '***************************************************
    #If BYTE_QUEUE_DEBUG = 1 Then
        Call ReadCheckType(wrSingle)
    #End If

    Dim buf(3) As Byte

    'Read the data and remove it
    Call ReadData(buf, 4)

    'Copy data to temp buffer
    Call CopyMemory(PeekSingle, buf(0), 4)
    #If BYTE_QUEUE_DEBUG = 1 Then
        Call RestoreCheckType
    #End If

End Function

''
' Reads a double from the begining of the queue but DOES NOT remove it.
'
' @return   The read value.
' @remarks  Peek methods, unlike Read methods, don't remove the data from the queue.
' @see ReadDouble
' @see WriteDouble

Public Function PeekDouble() As Double
    '***************************************************
    'Autor: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 10/28/07
    'Reads a double from the queue but doesn't removes it
    '***************************************************
    #If BYTE_QUEUE_DEBUG = 1 Then
        Call ReadCheckType(wrDouble)
    #End If

    Dim buf(7) As Byte

    'Read the data and remove it
    Call ReadData(buf, 8)

    'Copy data to temp buffer
    Call CopyMemory(PeekDouble, buf(0), 8)
    #If BYTE_QUEUE_DEBUG = 1 Then
        Call RestoreCheckType
    #End If

End Function

''
' Reads a Bollean from the begining of the queue but DOES NOT remove it.
'
' @return   The read value.
' @remarks  Peek methods, unlike Read methods, don't remove the data from the queue.
' @see ReadBoolean
' @see WriteBoolean

Public Function PeekBoolean() As Boolean
    '***************************************************
    'Autor: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 04/27/06
    'Reads a Boolean from the queue but doesn't removes it
    '***************************************************
    #If BYTE_QUEUE_DEBUG = 1 Then
        Call ReadCheckType(wrBoolean)
    #End If

    Dim buf(0) As Byte

    'Read the data and remove it
    Call ReadData(buf, 1)

    If buf(0) = 1 Then PeekBoolean = True
    #If BYTE_QUEUE_DEBUG = 1 Then
        Call RestoreCheckType
    #End If

End Function

''
' Reads a fixed length ASCII string from the begining of the queue but DOES NOT remove it.
'
' @param length The length of the string to be read
' @return   The read string if enough data is available, an empty string otherwise.
' @remarks  Peek methods, unlike Read methods, don't remove the data from the queue.
' If there is not enough data to read the complete string then an empty string is returned
' @see ReadASCIIStringFixed
' @see WriteASCIIStringFixed

Public Function PeekASCIIStringFixed(ByVal Length As Long) As String
    '***************************************************
    'Autor: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 04/27/06
    'Reads a fixed length ASCII string from the queue but doesn't removes it
    '***************************************************
    #If BYTE_QUEUE_DEBUG = 1 Then
        Call ReadCheckType(wrASCIStringFixed)
    #End If

    If Length <= 0 Then
        #If BYTE_QUEUE_DEBUG = 1 Then
            Call RestoreCheckType
        #End If

        Exit Function

    End If

    If queueLength >= Length Then

        Dim buf() As Byte

        ReDim buf(Length - 1) As Byte

        'Read the data and remove it
        Call ReadData(buf, Length)

        PeekASCIIStringFixed = StrConv(buf, vbUnicode)
    Else
        Call ErrRaise(NOT_ENOUGH_DATA)

    End If

    #If BYTE_QUEUE_DEBUG = 1 Then
        Call RestoreCheckType
    #End If

End Function

''
' Reads a fixed length unicode string from the begining of the queue but DOES NOT remove it.
'
' @param length The length of the string to be read
' @return   The read string if enough data is available, an empty string otherwise.
' @remarks  Peek methods, unlike Read methods, don't remove the data from the queue.
' If there is not enough data to read the complete string then an empty string is returned
' @see ReadUnicodeStringFixed
' @see WriteUnicodeStringFixed

Public Function PeekUnicodeStringFixed(ByVal Length As Long) As String

    '***************************************************
    'Autor: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 04/27/06
    'Reads a fixed length UNICODE string from the queue but doesn't removes it
    '***************************************************
    If Length <= 0 Then Exit Function

    Dim i As Integer

    i = Length + Length

    If queueLength >= i Then

        Dim buf() As Byte

        ReDim buf(i - 1) As Byte

        'Read the data and remove it
        Call ReadData(buf, i)

        PeekUnicodeStringFixed = buf
    Else
        Call ErrRaise(NOT_ENOUGH_DATA)

    End If

End Function

''
' Reads a variable length ASCII string from the begining of the queue but DOES NOT remove it.
'
' @return   The read string if enough data is available, an empty string otherwise.
' @remarks  Peek methods, unlike Read methods, don't remove the data from the queue.
' If there is not enough data to read the complete string then an empty string is returned
' @see ReadASCIIString
' @see WriteASCIIString

Public Function PeekASCIIString() As String

    '***************************************************
    'Autor: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 04/27/06
    'Reads a variable length ASCII string from the queue but doesn't removes it
    '***************************************************
    Dim buf(1) As Byte

    Dim Length As Integer

    #If BYTE_QUEUE_DEBUG = 1 Then
        Call ReadCheckType(wrASCIIString)
    #End If

    'Make sure we can read a valid length
    If queueLength > 1 Then
        'Read the length
        Call ReadData(buf, 2)
        Call CopyMemory(Length, buf(0), 2)

        'Make sure there are enough bytes
        If queueLength >= Length + 2 Then

            Dim buf2() As Byte

            ReDim buf2(Length + 1) As Byte

            'Read the data (we have to read the length again)
            Call ReadData(buf2, Length + 2)

            If Length > 0 Then

                'Get rid of the length
                Dim buf3() As Byte

                ReDim buf3(Length - 1) As Byte
                Call CopyMemory(buf3(0), buf2(2), Length)

                PeekASCIIString = StrConv(buf3, vbUnicode)

            End If

        Else
            Call ErrRaise(NOT_ENOUGH_DATA)

        End If

    Else
        Call ErrRaise(NOT_ENOUGH_DATA)

    End If

    #If BYTE_QUEUE_DEBUG = 1 Then
        Call RestoreCheckType
    #End If

End Function

''
' Reads a variable length unicode string from the begining of the queue but DOES NOT remove it.
'
' @return   The read string if enough data is available, an empty string otherwise.
' @remarks  Peek methods, unlike Read methods, don't remove the data from the queue.
' If there is not enough data to read the complete string then an empty string is returned
' @see ReadUnicodeString
' @see WriteUnicodeString

Public Function PeekUnicodeString() As String

    '***************************************************
    'Autor: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 04/27/06
    'Reads a variable length UNICODE string from the queue but doesn't removes it
    '***************************************************
    Dim buf(1) As Byte

    Dim Length As Integer

    'Make sure we can read a valid length
    If queueLength > 1 Then
        'Read the length
        Call ReadData(buf, 2)
        Call CopyMemory(Length, buf(0), 2)

        'Make sure there are enough bytes
        If queueLength >= Length * 2 + 2 Then

            Dim buf2() As Byte

            ReDim buf2(Length * 2 + 1) As Byte

            'Read the data (we need to read the length again)
            Call ReadData(buf2, Length * 2 + 2)

            'Get rid of the length bytes
            Dim buf3() As Byte

            ReDim buf3(Length * 2 - 1) As Byte
            Call CopyMemory(buf3(0), buf2(2), Length * 2)

            PeekUnicodeString = buf3
        Else
            Call ErrRaise(NOT_ENOUGH_DATA)

        End If

    Else
        Call ErrRaise(NOT_ENOUGH_DATA)

    End If

End Function

' Reads a byte array from the begining of the queue but DOES NOT remove it.
'
' @param block() Byte array that will contain the read data. MUST be Base 0 and previously resized to contain the requested amount of bytes.
' @param dataLength Number of bytes to be read
' @return   The actual number of read bytes.
' @remarks  Peek methods, unlike Read methods, don't remove the data from the queue.
' @see ReadBlock
' @see WriteBlock

Public Function PeekBlock(ByRef Block() As Byte, ByVal DataLength As Long) As Long

    'Read the data
    If DataLength > 0 Then
        PeekBlock = ReadData(Block(), DataLength)

    End If

End Function

' Retrieves the current capacity of the queue.
'
' @return   The current capacity of the queue.

Public Property Get Capacity() As Long

    Capacity = queueCapacity

End Property

' Sets the capacity of the queue.
'
' @param value The new capacity of the queue.
' @remarks If the new capacity is smaller than the current Length, all exceeding data is lost.
' @see Length

Public Property Let SetCapacity(ByVal value As Long)

    'Upate capacity
    queueCapacity = value
    CUSTOM_DATA_BUFFER = queueCapacity

    'All extra data is lost
    If Length > value Then queueLength = value

    'Resize the queue
    ReDim Preserve DATA(queueCapacity - 1) As Byte

End Property

''
' Retrieves the length of the total data in the queue.
'
' @return   The length of the total data in the queue.

Public Property Get Length() As Long

    Length = queueLength

End Property

' Retrieves the NOT_ENOUGH_DATA error code.
'
' @return   NOT_ENOUGH_DATA.

Public Property Get NotEnoughDataErrCode() As Long

    NotEnoughDataErrCode = NOT_ENOUGH_DATA

End Property

' Retrieves the NOT_ENOUGH_SPACE error code.
'
' @return   NOT_ENOUGH_SPACE.

Public Property Get NotEnoughSpaceErrCode() As Long

    NotEnoughSpaceErrCode = NOT_ENOUGH_SPACE

End Property

Public Function PeekAll() As Byte()
    PeekAll = DATA

End Function

Public Sub Clean()
    queueLength = 0
    queuePos = 0

End Sub

Public Function ReadAll() As Byte()
    Call EndPacket

    If queueLength > 0 Then

        Dim tData() As Byte

        ReDim tData(queueLength - 1) As Byte

        CopyMemory tData(0), DATA(0), queueLength

        ReadAll = tData

        queuePos = 0
        queueLength = 0

    End If

End Function

Public Sub ReadAllFixed(ByRef tData() As Byte)

    ReDim tData(queueLength - 1) As Byte

    CopyMemory tData(0), DATA(0), queueLength

    queuePos = 0
    queueLength = 0

End Sub

' Limpia la data del buffer del queue array
Public Sub Clear()
    ReDim DATA(CUSTOM_DATA_BUFFER - 1) As Byte
    queueCapacity = CUSTOM_DATA_BUFFER
    queueLength = 0
    queuePos = 0

End Sub

Public Function WriteID(ByVal value As Byte) As Long

    If queueCapacity - queueLength - GetHeaderSize < 0 Then

        Exit Function

    End If

    Call EndPacket
    #If BYTE_QUEUE_DEBUG = 1 Then
        'Debug.Print "WriteID: " & value
    #End If

    PacketStartPosition = queueLength

    Call WriteInteger(value)
    Call WriteInteger(0)

    WriteID = 4

End Function

Public Function GetCRCSize() As Byte
    #If NO_CRC = 1 Then
        GetCRCSize = 0
    #Else
        #If BYTE_QUEUE_DEBUG = 1 Then
            GetCRCSize = 2
        #Else
            GetCRCSize = 1
        #End If
    #End If

End Function

Public Function GetHeaderSize() As Byte
    #If BYTE_QUEUE_DEBUG = 1 Then
        GetHeaderSize = 5
    #Else
        GetHeaderSize = 3
    #End If

End Function

Public Function ValidCRC() As Boolean

    #If NO_CRC = 1 Then
        ValidCRC = True
    #Else

        Dim packetLength As Integer

        Dim i            As Long

        Dim crc          As Byte

        crc = 47
        packetLength = PeekLength()

        For i = queuePos To queuePos + packetLength - GetCRCSize - 1
            crc = crc Xor DATA(i)
        Next i

        ValidCRC = crc = PeekCRC(queuePos + packetLength - 1)
    #End If

End Function

Public Function ReadID() As Byte

    If GetHeaderSize > queueLength Then
        Call ErrRaise(NOT_ENOUGH_DATA)

        Exit Function

    End If

    Dim PacketID As Byte

    PacketID = ReadInteger()
    LastLength = ReadInteger()

    ReadID = PacketID

End Function

Public Function PeekCRC(ByVal pos As Integer) As Byte

    #If BYTE_QUEUE_DEBUG = 1 Then
        Call ReadCheckType(wrByte)
    #End If

    If 1 > queueLength Then
        Call ErrRaise(NOT_ENOUGH_DATA)
        Exit Function

    End If

    PeekCRC = DATA(pos)

    #If BYTE_QUEUE_DEBUG = 1 Then
        Call RestoreCheckType
    #End If

End Function

Public Function SafeClearPacket()    ' SI O SI HAY QUE USARLO DESPUES DE UN READID

    If queuePos = 0 Then
        Call SkipPacket
    Else
        Call RemoveData(LastLength - queuePos - GetCRCSize)
        Call ReadNewPacket

    End If

End Function

Public Function CheckLength() As Boolean

    'Call ReadNewPacket

    'CopyMemory packetLength, data(queuePos + 1), 2
    If queueLength < GetHeaderSize Then
        CheckLength = False
    Else

        Dim packetLength As Integer

        packetLength = PeekLength

        If packetLength < GetHeaderSize Then
            CheckLength = False
        ElseIf queueLength < packetLength Then
            CheckLength = False
        Else
            CheckLength = True

        End If

    End If

End Function

Public Sub EndPacket()

    If PacketStartPosition >= 0 Then
        packetLength = queueLength - PacketStartPosition + GetCRCSize

        #If BYTE_QUEUE_DEBUG = 1 Then
            CopyMemory DATA(PacketStartPosition + 3), packetLength, 2
        #Else
            CopyMemory DATA(PacketStartPosition + 1), packetLength, 2
        #End If

        #If NO_CRC = 0 Then

            Dim i   As Long

            Dim crc As Byte

            crc = 47

            For i = PacketStartPosition To queueLength - 1
                crc = crc Xor DATA(i)
            Next i

            Call WriteByte(crc)
        #End If

        PacketStartPosition = -1
        
        #If BYTE_QUEUE_DEBUG = 1 Then
        
        Else
            'Stop ' Hice EndPacket y no estoy escribiendo ningun paquete...
        #End If
        
    End If

End Sub

Public Function BufferOver() As Boolean

    If queueLength <= queuePos Then
        Call Clean
        BufferOver = True
    Else
        BufferOver = False

    End If

End Function

Public Function ReadNewPacket()

    If queuePos > 0 Then
        #If NO_CRC = 0 Then
            Call ReadByte      ' Borramos el CRC!!
        #End If

        If queuePos > queueLength Then
            Call ErrRaise(NOT_ENOUGH_DATA)

            Exit Function

        End If

        If queueCapacity - queueLength - queuePos < 0 Then
            Call ErrRaise(NOT_ENOUGH_DATA)

            Exit Function

        End If

        #If BYTE_QUEUE_DEBUG = 1 Then
            Call CheckCopyMemory(queuePos, queueLength - queuePos)
        #End If
        CopyMemory DATA(0), DATA(queuePos), queueLength - queuePos
        queueLength = queueLength - queuePos
        queuePos = 0
        LastLength = 0

    End If

End Function

Public Function SkipLength(ByVal Length As Integer)

    Call EndPacket

    Dim packet_length As Integer

    packet_length = Length

    If packet_length > queueLength Then
        Call ErrRaise(NOT_ENOUGH_DATA)

        Exit Function

    End If

    #If BYTE_QUEUE_DEBUG = 1 Then
        Call CheckCopyMemory(packet_length, queueLength - packet_length)
    #End If

    If queueLength - packet_length > 0 Then
        CopyMemory DATA(0), DATA(packet_length), queueLength - packet_length
        queueLength = queueLength - packet_length
        queuePos = 0
    Else
        queueLength = 0
        queuePos = 0

    End If

    'LastLength = 0

End Function

Private Function SkipPacket()

    Dim packet_length As Integer

    packet_length = PeekLength

    If packet_length > queueLength Then
        Call ErrRaise(NOT_ENOUGH_DATA)

        Exit Function

    End If

    #If BYTE_QUEUE_DEBUG = 1 Then
        Call CheckCopyMemory(packet_length, queueLength - packet_length)
    #End If

    CopyMemory DATA(0), DATA(packet_length), queueLength - packet_length
    queueLength = queueLength - packet_length
    queuePos = 0
    LastLength = 0

End Function

Public Function ErrRaise(ByVal Error As Long)
    errNumber = Error

End Function

Private Function CheckCopyMemory(ByVal pos As Long, ByVal Length As Long) As Boolean

    If pos + Length > queueLength Then Stop

End Function

