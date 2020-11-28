Attribute VB_Name = "ModCompresion"
'*****************************************************************
'modCompression.bas - v1.0.0
'
'All methods to handle resource files
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
'Contributors History
'   When releasing modifications to this source file please add your
'   date of release, name, email, and any info to the top of this list.
'   Follow this template:
'    XX/XX/200X - Your Name Here (Your Email Here)
'       - Your Description Here
'       Sub Release Contributors:
'           XX/XX/2003 - Sub Contributor Name Here (SC Email Here)
'               - SC Description Here
'*****************************************************************
'
'Juan Martín Sotuyo Dodero (juansotuyo@hotmail.com) - 10/13/2004
'   - First Release
'*****************************************************************
Option Explicit
Private Declare Function GetDiskFreeSpace Lib "kernel32" Alias "GetDiskFreeSpaceExA" (ByVal lpRootPathName As String, FreeBytesToCaller As Currency, BytesTotal As Currency, FreeBytesTotal As Currency) As Long
'This structure will describe our binary file's
'size and number of contained files
Public Type FILEHEADER
    lngFileSize As Long                 'How big is this file? (Used to check integrity)
    intNumFiles As Integer              'How many files are inside?
End Type

'This structure will describe each file contained
'in our binary file
Public Type INFOHEADER
    lngFileStart As Long            'Where does the chunk start?
    lngFileSize As Long             'How big is this chunk of stored data?
    strFileName As String * 16      'What's the name of the file this data came from?
    lngFileSizeUncompressed As Long 'How big is the file compressed
End Type

Public Enum resource_file_type
    Grh
    MIDI
    MP3
    WAV
    Scripts
    Patch
    Interfaz
End Enum
Private Const CarpetaRecursos As String = "\Recursos"
Private Const GRAPHIC_PATH As String = "\Graficos\"
Private Const MIDI_PATH As String = "\Midi\"
Private Const MP3_PATH As String = "\Mp3\"
Private Const WAV_PATH As String = "\Wavs\"
Private Const SCRIPT_PATH As String = "\Init\"
Private Const PATCH_PATH As String = "\Patches\"
Private Const INTERFAZ_PATH As String = "\Interfaz\"
Private Const OUTPUT_PATH As String = "\Output\"

Private Declare Function compress Lib "zlib.dll" (dest As Any, destlen As Any, src As Any, ByVal srclen As Long) As Long
Private Declare Function uncompress Lib "zlib.dll" (dest As Any, destlen As Any, src As Any, ByVal srclen As Long) As Long

Public Sub Compress_Data(ByRef data() As Byte)
'*****************************************************************
'Author: Juan Martín Dotuyo Dodero
'Last Modify Date: 10/13/2004
'Compresses binary data avoiding data loses
'*****************************************************************
    Dim Dimensions As Long
    Dim DimBuffer As Long
    Dim BufTemp() As Byte
    Dim BufTemp2() As Byte
    Static compression_rate As Single
    Dim loopc As Long
    
    If compression_rate = 0 Then compression_rate = 0.05
    
    Dimensions = UBound(data)
    
    DimBuffer = Dimensions * compression_rate
    
    ReDim BufTemp(DimBuffer)
    
    compress BufTemp(0), DimBuffer, data(0), Dimensions
    
    'Check if there was data loss
    ReDim BufTemp2(Dimensions)
    
    uncompress BufTemp2(0), Dimensions, BufTemp(0), UBound(BufTemp) + 1
    
    For loopc = 0 To UBound(data)
        If data(loopc) <> BufTemp2(loopc) Then
            'Clear memory
            Erase BufTemp
            Erase BufTemp2
            
            'If we have reached 1, then just copy the data
            If compression_rate < 1 Then
                'Increase compression rate
                compression_rate = compression_rate + 0.05
                'Try again
                Compress_Data data
            End If
            
            'Reset compression rate and exit
            compression_rate = 0.05
            Exit Sub
        End If
    Next loopc
    
    Erase data
    
    ReDim data(DimBuffer - 1)
    
    data = BufTemp
    
    Erase BufTemp
    Erase BufTemp2
    
    'Encrypt the first byte of the compressed data for extra security
    data(0) = data(0) Xor 12
End Sub

Public Sub Decompress_Data(ByRef data() As Byte, ByVal OrigSize As Long)
'*****************************************************************
'Author: Juan Martín Dotuyo Dodero
'Last Modify Date: 10/13/2004
'Decompresses binary data
'*****************************************************************
    Dim BufTemp() As Byte
    
    ReDim BufTemp(OrigSize - 1)
    
    'Des-encrypt the first byte of the compressed data
    data(0) = data(0) Xor 12
    
    uncompress BufTemp(0), OrigSize, data(0), UBound(data) + 1
    
    ReDim data(OrigSize - 1)
    
    data = BufTemp
    
    Erase BufTemp
End Sub

Public Sub Encrypt_File_Header(ByRef FileHead As FILEHEADER)
'*****************************************************************
'Author: Juan Martín Dotuyo Dodero
'Last Modify Date: 10/13/2004
'Encrypts normal data or turns encrypted data back to normal
'*****************************************************************
    'Each different variable is encrypted with a different key for extra security
    With FileHead
        .intNumFiles = .intNumFiles Xor 12345
        .lngFileSize = .lngFileSize Xor 1234567890
    End With
End Sub

Public Sub Encrypt_Info_Header(ByRef InfoHead As INFOHEADER)
'*****************************************************************
'Author: Juan Martín Dotuyo Dodero
'Last Modify Date: 10/13/2004
'Encrypts normal data or turns encrypted data back to normal
'*****************************************************************
    Dim EncryptedFileName As String
    Dim loopc As Long
    
    For loopc = 1 To Len(InfoHead.strFileName)
        If loopc Mod 2 = 0 Then
            EncryptedFileName = EncryptedFileName & Chr(Asc(mid(InfoHead.strFileName, loopc, 1)) Xor 123)
        Else
            EncryptedFileName = EncryptedFileName & Chr(Asc(mid(InfoHead.strFileName, loopc, 1)) Xor 12)
        End If
    Next loopc
    
    'Each different variable is encrypted with a different key for extra security
    With InfoHead
        .lngFileSize = .lngFileSize Xor 1234567890
        .lngFileSizeUncompressed = .lngFileSizeUncompressed Xor 1234567890
        .lngFileStart = .lngFileStart Xor 123456789
        .strFileName = EncryptedFileName
    End With
End Sub

Public Function Extract_Files(ByVal file_type As resource_file_type, ByVal resource_path As String, ByRef ResourcePrgbar As ProgressBar, ByRef GeneralPrgBar As ProgressBar, ByRef GeneralLbl As Label, Optional ByVal UseOutputFolder As Boolean = False) As Boolean
'*****************************************************************
'Author: Juan Martín Dotuyo Dodero
'Last Modify Date: 10/13/2004
'Extracts all files from a resource file
'*****************************************************************
    Dim loopc As Long
    Dim SourceFilePath As String
    Dim OutputFilePath As String
    Dim SourceFile As Integer
    Dim SourceData() As Byte
    Dim FileHead As FILEHEADER
    Dim InfoHead() As INFOHEADER
    Dim RequiredSpace As Currency

    
'Set up the error handler
On Local Error GoTo ErrHandler
    
    Select Case file_type
        Case Grh
            If UseOutputFolder Then
                SourceFilePath = resource_path & OUTPUT_PATH & "Graphics.ORE"
            Else
                SourceFilePath = resource_path & "\Graphics.ORE"
            End If
            OutputFilePath = resource_path & GRAPHIC_PATH
            
        Case MIDI
            If UseOutputFolder Then
                SourceFilePath = resource_path & OUTPUT_PATH & "Low-Def Music.ORE"
            Else
                SourceFilePath = resource_path & "\Low-Def Music.ORE"
            End If
            OutputFilePath = resource_path & MIDI_PATH
        
        Case MP3
            If UseOutputFolder Then
                SourceFilePath = resource_path & OUTPUT_PATH & "Hi-Def Music.ORE"
            Else
                SourceFilePath = resource_path & "\Hi-Def Music.ORE"
            End If
            OutputFilePath = resource_path & MP3_PATH
        
        Case WAV
            If UseOutputFolder Then
                SourceFilePath = resource_path & OUTPUT_PATH & "Sounds.ORE"
            Else
                SourceFilePath = resource_path & "\Sounds.ORE"
            End If
            OutputFilePath = resource_path & WAV_PATH
        
        Case Scripts
            If UseOutputFolder Then
                SourceFilePath = resource_path & OUTPUT_PATH & "init.ore"
                MsgBox ("Archivo init.ore encontrado")
            Else
                SourceFilePath = resource_path & "\init.ore"
            End If
            OutputFilePath = resource_path & CarpetaRecursos & SCRIPT_PATH
            MsgBox resource_path & CarpetaRecursos & SCRIPT_PATH
        
        Case Else
            Exit Function
    End Select
    
    'Open the binary file
    SourceFile = FreeFile
    Open SourceFilePath For Binary Access Read Lock Write As SourceFile
    
    'Extract the FILEHEADER
    Get SourceFile, 1, FileHead
    
    'Desencrypt File Header
    Encrypt_File_Header FileHead
    
    'Check the file for validity
    If LOF(SourceFile) <> FileHead.lngFileSize Then
        MsgBox "Resource file " & SourceFilePath & " seems to be corrupted.", , "Error"
        Close SourceFile
        Exit Function
    End If
    
    'Size the InfoHead array
    ReDim InfoHead(FileHead.intNumFiles - 1)
    
    'Extract the INFOHEADER
    Get SourceFile, , InfoHead
    
    'Check if there is enough hard drive space to extract all files
    For loopc = 0 To UBound(InfoHead)
        'Desencrypt each Info Header before accessing the data
        Encrypt_Info_Header InfoHead(loopc)
        RequiredSpace = RequiredSpace + InfoHead(loopc).lngFileSizeUncompressed
    Next loopc
    
    If RequiredSpace >= General_Drive_Get_Free_Bytes(Left(App.path, 3)) Then
        Erase InfoHead
        Close SourceFile
        MsgBox "There is not enough drive space to extract the compressed files.", , "Error"
        Exit Function
    End If
    
    'Size the file handler array
    ReDim file_handler_list(UBound(InfoHead))
    
    If Not ResourcePrgbar Is Nothing Then ResourcePrgbar.max = FileHead.intNumFiles
    
    'Extract all of the files from the binary file
    For loopc = 0 To UBound(InfoHead)
        'Resize the byte data array
        ReDim SourceData(InfoHead(loopc).lngFileSize - 1)
        
        'Get the data
        Get SourceFile, InfoHead(loopc).lngFileStart, SourceData
        
        'Decompress all data
        If InfoHead(loopc).lngFileSize < InfoHead(loopc).lngFileSizeUncompressed Then
            Decompress_Data SourceData, InfoHead(loopc).lngFileSizeUncompressed
        End If
        
        'Get a free handler
        file_handler_list(loopc).handle = FreeFile
        
        Open OutputFilePath & InfoHead(loopc).strFileName For Binary As file_handler_list(loopc).handle
        
        file_handler_list(loopc).file_name = InfoHead(loopc).strFileName
        
        Put file_handler_list(loopc).handle, , SourceData
        
        'We leave the files open so they are locked. To unlock them call Close_All
        
        Erase SourceData
        
        'Update progress bars
        If Not ResourcePrgbar Is Nothing Then ResourcePrgbar.value = ResourcePrgbar.value + 1
        If Not GeneralPrgBar Is Nothing Then GeneralPrgBar.value = GeneralPrgBar.value + 1
        If Not GeneralLbl Is Nothing Then GeneralLbl.Caption = "Loading File " & loopc & " of " & FileHead.intNumFiles & "..."
        DoEvents
    Next loopc
    
    'Close the binary file
    Close SourceFile
    
    'Copy handler list
    Select Case file_type
        Case Grh
            ReDim GRH_Handles(UBound(file_handler_list()))
            GRH_Handles = file_handler_list
            
        Case MIDI
            ReDim MIDI_Handles(UBound(file_handler_list()))
            MIDI_Handles = file_handler_list
        
        Case MP3
            ReDim MP3_Handles(UBound(file_handler_list()))
            MP3_Handles = file_handler_list
        
        Case WAV
            ReDim WAV_Handles(UBound(file_handler_list()))
            WAV_Handles = file_handler_list
        
        Case Scripts
            ReDim Scripts_Handles(UBound(file_handler_list()))
            Scripts_Handles = file_handler_list
    End Select
    
    Erase InfoHead
    
    Extract_Files = True
Exit Function

ErrHandler:
    Close SourceFile
    Erase SourceData
    Erase InfoHead
    'Display an error message if it didn't work
    MsgBox "Unable to decode binary file. Reason: " & Err.number & " : " & Err.Description, vbOKOnly, "Error"
End Function

Public Function Extract_Patch(ByVal resource_path As String, ByVal file_name As String) As Boolean
'*****************************************************************
'Author: Juan Martín Dotuyo Dodero
'Last Modify Date: 10/13/2004
'Comrpesses all files to a resource file
'*****************************************************************
    Dim loopc As Long
    Dim LoopC2 As Long
    Dim LoopC3 As Long
    Dim OutputFilePath As String
    Dim OutputFile As Integer
    Dim UpdatedFile As Integer
    Dim SourceFilePath As String
    Dim SourceFile As Integer
    Dim SourceData() As Byte
    Dim ResFileHead As FILEHEADER
    Dim ResInfoHead() As INFOHEADER
    Dim UpdatedInfoHead As INFOHEADER
    Dim FileHead As FILEHEADER
    Dim InfoHead() As INFOHEADER
    Dim RequiredSpace As Currency
    Dim FileExtension As String
    Dim DataOffset As Long
    
    'Done flags
    Dim bmp_done As Boolean
    Dim wav_done As Boolean
    Dim mid_done As Boolean
    Dim mp3_done As Boolean
    
    '************************************************************************************************
    'This is similar to Extract, but has some small differences to make sure what is being updated
    '************************************************************************************************
'Set up the error handler
'On Local Error GoTo ErrHandler
    
    'Open the binary file
    SourceFile = FreeFile
    SourceFilePath = resource_path & "\" & file_name
    Open SourceFilePath For Binary Access Read Lock Write As SourceFile
    
    'Extract the FILEHEADER
    Get SourceFile, 1, FileHead
    
    'Desencrypt File Header
    Encrypt_File_Header FileHead
    
    'Check the file for validity
    If LOF(SourceFile) <> FileHead.lngFileSize Then
        MsgBox "Resource file " & SourceFilePath & " seems to be corrupted.", , "Error"
        Exit Function
    End If
    
    'Size the InfoHead array
    ReDim InfoHead(FileHead.intNumFiles - 1)
    
    'Extract the INFOHEADER
    Get SourceFile, , InfoHead
    
    'Check if there is enough hard drive space to extract all files
    For loopc = 0 To UBound(InfoHead)
        'Desencrypt each Info Header befWO accessing the data
        Encrypt_Info_Header InfoHead(loopc)
        RequiredSpace = RequiredSpace + InfoHead(loopc).lngFileSizeUncompressed
    Next loopc
    
    If RequiredSpace >= General_Drive_Get_Free_Bytes(Left(App.path, 3)) Then
        Erase InfoHead
        MsgBox "There is not enough drive space to extract the compressed files.", , "Error"
        Exit Function
    End If
    
    'Extract all of the files from the binary file
    For loopc = 0 To UBound(InfoHead())
        'Check the extension of the file
        Select Case LCase(Right(Trim(InfoHead(loopc).strFileName), 3))
            Case Is = "bmp"
                If bmp_done Then GoTo EndMainLoop
                FileExtension = "bmp"
                OutputFilePath = resource_path & "\Graphics.WO"
                bmp_done = True
            Case Is = "mid"
                If mid_done Then GoTo EndMainLoop
                FileExtension = "mid"
                OutputFilePath = resource_path & "\Low-Def Music.WO"
                mid_done = True
            Case Is = "mp3"
                If mp3_done Then GoTo EndMainLoop
                FileExtension = "mp3"
                OutputFilePath = resource_path & "\Hi-Def Music.WO"
                mp3_done = True
            Case Is = "wav"
                If wav_done Then GoTo EndMainLoop
                FileExtension = "wav"
                OutputFilePath = resource_path & "\Sounds.WO"
                wav_done = True
            Case Else
                MsgBox "Unkown file extension detected: " & LCase(Right(Trim(InfoHead(loopc).strFileName), 4))
                Exit Function
        End Select
        
        OutputFile = FreeFile
        Open OutputFilePath For Binary Access Read Lock Write As OutputFile
        
        'Get file header
        Get OutputFile, 1, ResFileHead
        
        'Desencrypt file header
        Encrypt_File_Header ResFileHead
        
        'Resize the Info Header array
        ReDim ResInfoHead(ResFileHead.intNumFiles - 1)
        
        'Load the info header
        Get OutputFile, , ResInfoHead
        
        'Desencrypt all Info Headers
        For LoopC2 = 0 To UBound(ResInfoHead())
            Encrypt_Info_Header ResInfoHead(LoopC2)
        Next LoopC2
        
        'Check how many of the files are new, and how many are replacements
        For LoopC2 = loopc To UBound(InfoHead())
            If LCase(Right(Trim(InfoHead(LoopC2).strFileName), 3)) = FileExtension Then
                'Look for same name in the resource file
                For LoopC3 = 0 To UBound(ResInfoHead())
                    If ResInfoHead(LoopC3).strFileName = InfoHead(LoopC2).strFileName Then
                        Exit For
                    End If
                Next LoopC3
                
                'Update the File Head
                If LoopC3 > UBound(ResInfoHead()) Then
                    'Update number of files and size
                    ResFileHead.intNumFiles = ResFileHead.intNumFiles + 1
                    ResFileHead.lngFileSize = ResFileHead.lngFileSize + Len(InfoHead(0)) + InfoHead(LoopC2).lngFileSize
                Else
                    'We substract the size of the old file and add the one of the new one
                    ResFileHead.lngFileSize = ResFileHead.lngFileSize - ResInfoHead(LoopC3).lngFileSize + InfoHead(LoopC2).lngFileSize
                End If
            End If
        Next LoopC2
        
        'Get the offset of the compressed data
        DataOffset = ResFileHead.intNumFiles * Len(ResInfoHead(0)) + Len(FileHead) + 1
        
        'Encrypt file Header
        Encrypt_File_Header ResFileHead
        
        'Now we start saving the updated file
        UpdatedFile = FreeFile
        Open OutputFilePath & "2" For Binary Access Write Lock Read As UpdatedFile
        
        'StWO the filehead
        Put UpdatedFile, 1, ResFileHead
        
        'Start storing the Info Heads
        LoopC2 = loopc
        For LoopC3 = 0 To UBound(ResInfoHead())
            Do While LoopC2 <= UBound(InfoHead())
                If LCase(ResInfoHead(LoopC3).strFileName) < LCase(InfoHead(LoopC2).strFileName) Then Exit Do
                If LCase(Right(Trim(InfoHead(LoopC2).strFileName), 3)) = FileExtension Then
                    'Copy the info head data
                    UpdatedInfoHead = InfoHead(LoopC2)
                    
                    'Set the file start pos and update the offset
                    UpdatedInfoHead.lngFileStart = DataOffset
                    DataOffset = DataOffset + UpdatedInfoHead.lngFileSize
                    
                    'Encrypt the info header and save it
                    Encrypt_Info_Header UpdatedInfoHead
                    
                    Put UpdatedFile, , UpdatedInfoHead
                End If
                LoopC2 = LoopC2 + 1
            Loop
            
            'If the file was replaced in the patch, we skip it
            If LoopC2 Then
                If LCase(ResInfoHead(LoopC3).strFileName) <= LCase(InfoHead(LoopC2 - 1).strFileName) Then GoTo EndLoop
            End If
            
            'Copy the info head data
            UpdatedInfoHead = ResInfoHead(LoopC3)
            
            'Set the file start pos and update the offset
            UpdatedInfoHead.lngFileStart = DataOffset
            DataOffset = DataOffset + UpdatedInfoHead.lngFileSize
            
            'Encrypt the info header and save it
            Encrypt_Info_Header UpdatedInfoHead
            
            Put UpdatedFile, , UpdatedInfoHead
EndLoop:
        Next LoopC3
        
        'If there was any file in the patch that would go in the bottom of the list we put it now
        For LoopC2 = LoopC2 To UBound(InfoHead())
            If LCase(Right(Trim(InfoHead(LoopC2).strFileName), 3)) = FileExtension Then
                'Copy the info head data
                UpdatedInfoHead = InfoHead(LoopC2)
                
                'Set the file start pos and update the offset
                UpdatedInfoHead.lngFileStart = DataOffset
                DataOffset = DataOffset + UpdatedInfoHead.lngFileSize
                
                'Encrypt the info header and save it
                Encrypt_Info_Header UpdatedInfoHead
                
                Put UpdatedFile, , UpdatedInfoHead
            End If
        Next LoopC2
        
        'Now we start adding the compressed data
        LoopC2 = loopc
        For LoopC3 = 0 To UBound(ResInfoHead())
            Do While LoopC2 <= UBound(InfoHead())
                If LCase(ResInfoHead(LoopC3).strFileName) < LCase(InfoHead(LoopC2).strFileName) Then Exit Do
                If LCase(Right(Trim(InfoHead(LoopC2).strFileName), 3)) = FileExtension Then
                    'Get the compressed data
                    ReDim SourceData(InfoHead(LoopC2).lngFileSize - 1)
                    
                    Get SourceFile, InfoHead(LoopC2).lngFileStart, SourceData
                    
                    Put UpdatedFile, , SourceData
                End If
                LoopC2 = LoopC2 + 1
            Loop
            
            'If the file was replaced in the patch, we skip it
            If LoopC2 Then
                If LCase(ResInfoHead(LoopC3).strFileName) <= LCase(InfoHead(LoopC2 - 1).strFileName) Then GoTo EndLoop2
            End If
            
            'Get the compressed data
            ReDim SourceData(ResInfoHead(LoopC3).lngFileSize - 1)
            
            Get OutputFile, ResInfoHead(LoopC3).lngFileStart, SourceData
            
            Put UpdatedFile, , SourceData
EndLoop2:
        Next LoopC3
        
        'If there was any file in the patch that would go in the bottom of the lsit we put it now
        For LoopC2 = LoopC2 To UBound(InfoHead())
            If LCase(Right(Trim(InfoHead(LoopC2).strFileName), 3)) = FileExtension Then
                'Get the compressed data
                ReDim SourceData(InfoHead(LoopC2).lngFileSize - 1)
                
                Get SourceFile, InfoHead(LoopC2).lngFileStart, SourceData
                
                Put UpdatedFile, , SourceData
            End If
        Next LoopC2
        
        'We are done updating the file
        Close UpdatedFile
        
        'Close and delete the old resource file
        Close OutputFile
        Kill OutputFilePath
        
        'Rename the new one
        Name OutputFilePath & "2" As OutputFilePath
        
        'Deallocate the memory used by the data array
        Erase SourceData
EndMainLoop:
    Next loopc
    
    'Close the binary file
    Close SourceFile
    
    Erase InfoHead
    Erase ResInfoHead
    
    Extract_Patch = True
Exit Function

ErrHandler:
    Erase SourceData
    Erase InfoHead
    'Display an error message if it didn't work
    MsgBox "Unable to decode binary file. Reason: " & Err.number & " : " & Err.Description, vbOKOnly, "Error"
End Function

Public Function Compress_Files(ByVal file_type As resource_file_type, ByVal resource_path As String, ByVal dest_path As String, ByRef GeneralPrgBar As ProgressBar) As Boolean
'*****************************************************************
'Author: Juan Martín Dotuyo Dodero
'Last Modify Date: 10/13/2004
'Comrpesses all files to a resource file
'*****************************************************************
    Dim SourceFilePath As String
    Dim SourceFileExtension As String
    Dim OutputFilePath As String
    Dim SourceFile As Long
    Dim OutputFile As Long
    Dim SourceFileName As String
    Dim SourceData() As Byte
    Dim FileHead As FILEHEADER
    Dim InfoHead() As INFOHEADER
    Dim FileNames() As String
    Dim lngFileStart As Long
    Dim loopc As Long
    
'Set up the error handler
On Local Error GoTo ErrHandler
    
    Select Case file_type
        Case Grh
            SourceFilePath = resource_path & GRAPHIC_PATH
            SourceFileExtension = ".bmp"
            OutputFilePath = dest_path & "Graficos.ORE"
        
        Case MIDI
            SourceFilePath = resource_path & MIDI_PATH
            SourceFileExtension = ".mid"
            OutputFilePath = dest_path & "Midi.ore"
        
        Case MP3
            SourceFilePath = resource_path & MP3_PATH
            SourceFileExtension = ".mp3"
            OutputFilePath = dest_path & "mp3.ore"
        
        Case WAV
            SourceFilePath = resource_path & WAV_PATH
            SourceFileExtension = ".wav"
            OutputFilePath = dest_path & "Sonidos.ore"
        
        Case Scripts
            SourceFilePath = resource_path & SCRIPT_PATH
            SourceFileExtension = ".*"
            OutputFilePath = dest_path & "init.ore"
        
        Case Patch
            SourceFilePath = resource_path & PATCH_PATH
            SourceFileExtension = ".*"
            OutputFilePath = dest_path & "Patch.WO"
            
        Case Interfaz
            SourceFilePath = resource_path & INTERFAZ_PATH
            SourceFileExtension = ".*"
            OutputFilePath = dest_path & "Interfaz.ore"
    End Select
    
    'Get first file in the directoy
    SourceFileName = Dir(SourceFilePath & "*" & SourceFileExtension, vbNormal)
    
    SourceFile = FreeFile
    
    'Get all other files i nthe directory
    While SourceFileName <> ""
        FileHead.intNumFiles = FileHead.intNumFiles + 1
        
        ReDim Preserve FileNames(FileHead.intNumFiles - 1)
        FileNames(FileHead.intNumFiles - 1) = LCase$(SourceFileName)
        
        'Search new file
        SourceFileName = Dir
    Wend
    
    'If we found none, be can't compress a thing, so we exit
    If FileHead.intNumFiles = 0 Then
        MsgBox "There are no files of extension " & SourceFileExtension & " in " & SourceFilePath & ".", , "Error"
        Exit Function
    End If
    
    'Sort file names alphabetically (this will make patching much easier).

    
    'Resize InfoHead array
    ReDim InfoHead(FileHead.intNumFiles - 1)
    
    GeneralPrgBar.max = FileHead.intNumFiles
    GeneralPrgBar.value = 0
    
    'Destroy file if it previuosly existed
    If Dir(OutputFilePath, vbNormal) <> "" Then
        Kill OutputFilePath
    End If
    
    'Open a new file
    OutputFile = FreeFile
    Open OutputFilePath For Binary Access Read Write As OutputFile
    
    For loopc = 0 To FileHead.intNumFiles - 1
        'Find a free file number to use and open the file
        SourceFile = FreeFile
        Open SourceFilePath & FileNames(loopc) For Binary Access Read Lock Write As SourceFile
        
        'StWO file name
        InfoHead(loopc).strFileName = FileNames(loopc)
        
        'Find out how large the file is and resize the data array appropriately
        ReDim SourceData(LOF(SourceFile) - 1)
        
        'StWO the value so we can decompress it later on
        InfoHead(loopc).lngFileSizeUncompressed = LOF(SourceFile)
        
        'Get the data from the file
        Get SourceFile, , SourceData
        
        'Compress it
        Compress_Data SourceData
        
        'Save it to a temp file
        Put OutputFile, , SourceData
        
        'Set up the file header
        FileHead.lngFileSize = FileHead.lngFileSize + UBound(SourceData) + 1
        
        'Set up the info headers
        InfoHead(loopc).lngFileSize = UBound(SourceData) + 1
        
        Erase SourceData
        
        'Close temp file
        Close SourceFile
        
        'Update progress bar
        If Not GeneralPrgBar Is Nothing Then GeneralPrgBar.value = GeneralPrgBar.value + 1
        DoEvents
    Next loopc
    
    'Finish setting the FileHeader data
    FileHead.lngFileSize = FileHead.lngFileSize + CLng(FileHead.intNumFiles) * Len(InfoHead(0)) + Len(FileHead)
    
    'Set InfoHead data
    lngFileStart = Len(FileHead) + CLng(FileHead.intNumFiles) * Len(InfoHead(0)) + 1
    For loopc = 0 To FileHead.intNumFiles - 1
        InfoHead(loopc).lngFileStart = lngFileStart
        lngFileStart = lngFileStart + InfoHead(loopc).lngFileSize
        'Once an InfoHead index is ready, we encrypt it
        Encrypt_Info_Header InfoHead(loopc)
    Next loopc
    
    'Encrypt the FileHeader
    Encrypt_File_Header FileHead
    
    '************ Write Data
    
    'Get all data stWOd so far
    ReDim SourceData(LOF(OutputFile) - 1)
    Seek OutputFile, 1
    Get OutputFile, , SourceData
    
    Seek OutputFile, 1
    
    'StWO the data in the file
    Put OutputFile, , FileHead
    Put OutputFile, , InfoHead
    Put OutputFile, , SourceData
    
    'Close the file
    Close OutputFile
    
    Erase InfoHead
    Erase SourceData
Exit Function

ErrHandler:
    Erase SourceData
    Erase InfoHead
    'Display an error message if it didn't work
    MsgBox "Unable to create binary file. Reason: " & Err.number & " : " & Err.Description, vbOKOnly, "Error"
End Function
Public Sub Delete_Resources(ByVal resource_path As String)
'*****************************************************************
'Author: Juan Martín Dotuyo Dodero
'Last Modify Date: 10/13/2004
'Deletes all resource files
'*****************************************************************
On Local Error Resume Next

    
    Kill resource_path & GRAPHIC_PATH & "*.bmp"
    Kill resource_path & MP3_PATH & "*.mp3"
    Kill resource_path & MIDI_PATH & "*.mid"
    Kill resource_path & SCRIPT_PATH & "*.*"
End Sub
Public Function Extract_File(ByVal file_type As resource_file_type, ByVal resource_path As String, ByVal file_name As String, Optional ByVal UseOutputFolder As Boolean = False) As Boolean
'*****************************************************************
'Author: Juan Martín Dotuyo Dodero
'Last Modify Date: 10/13/2004
'Extracts all files from a resource file
'*****************************************************************
    Dim loopc As Long
    Dim SourceFilePath As String
    Dim OutputFilePath As String
    Dim SourceFile As Integer
    Dim SourceData() As Byte
    Dim FileHead As FILEHEADER
    Dim InfoHead() As INFOHEADER
    Dim handle As Integer
    
'Set up the error handler
On Local Error GoTo ErrHandler
    
    Select Case file_type
        Case Interfaz
            If UseOutputFolder Then
                SourceFilePath = resource_path & "Interfaz.ore"
            Else
                SourceFilePath = resource_path & "\Interfaz.ore"
            End If
            OutputFilePath = resource_path
            MsgBox OutputFilePath
         
        Case Else
            Exit Function
    End Select
    
    'Make sure it's lower case
    file_name = LCase$(file_name)
    
    'Open the binary file
    SourceFile = FreeFile
    Open SourceFilePath For Binary Access Read Lock Write As SourceFile
    
    'Extract the FILEHEADER
    Get SourceFile, 1, FileHead
    
    'Desencrypt File Header
    Encrypt_File_Header FileHead
    
    'Check the file for validity
    If LOF(SourceFile) <> FileHead.lngFileSize Then
        MsgBox "Resource file " & SourceFilePath & " seems to be corrupted.", , "Error"
        Close SourceFile
        Exit Function
    End If
    
    'Size the InfoHead array
    ReDim InfoHead(FileHead.intNumFiles - 1)
    
    'Extract the INFOHEADER
    Get SourceFile, , InfoHead
    
    'Get the file position within the compressed resource file
    For loopc = 0 To UBound(InfoHead)
        'Desencrypt each Info Header befWO accessing the data
        Encrypt_Info_Header InfoHead(loopc)
        If Left$(InfoHead(loopc).strFileName, Len(file_name)) = file_name Then
            Exit For
        End If
    Next loopc
    
    'Make sure index is valid
    If loopc > UBound(InfoHead) Then
        Erase InfoHead
        Close SourceFile
        Exit Function
    End If
    
    'Make sure there is enough space in the HD
    If InfoHead(loopc).lngFileSizeUncompressed > General_Drive_Get_Free_Bytes(Left(App.path, 3)) Then
        Erase InfoHead
        Close SourceFile
        MsgBox "There is not enough drive space to extract the compressed file.", , "Error"
        Exit Function
    End If
    
    'Extract file from the binary file
    
    'Resize the byte data array
    ReDim SourceData(InfoHead(loopc).lngFileSize - 1)
    
    'Get the data
    Get SourceFile, InfoHead(loopc).lngFileStart, SourceData
    
    'Decompress all data
    If InfoHead(loopc).lngFileSize < InfoHead(loopc).lngFileSizeUncompressed Then
        Decompress_Data SourceData, InfoHead(loopc).lngFileSizeUncompressed
    End If
    
    'Get a free handler
    handle = FreeFile
    
    Open OutputFilePath & InfoHead(loopc).strFileName For Binary As handle
    
    Put handle, , SourceData
    
    Close handle
    
    Erase SourceData
    
    'Close the binary file
    Close SourceFile
        
    Erase InfoHead
    
    Extract_File = True
Exit Function

ErrHandler:
    Close SourceFile
    Erase SourceData
    Erase InfoHead
    'Display an error message if it didn't work
    MsgBox "Unable to decode binary file. Reason: " & Err.number & " : " & Err.Description, vbOKOnly, "Error"
End Function



Public Function General_Load_Picture_From_Resource(ByVal picture_file_name As String) As IPicture
'**************************************************************
'Author: Augusto José Rando
'Last Modify Date: 6/11/2005
'Loads a picture from a resource file and returns it
'**************************************************************

'On Error GoTo ErrorHandler
MsgBox picture_file_name
If Extract_File(6, App.path & "\Recursos\", picture_file_name, True) Then
    Set General_Load_Picture_From_Resource = LoadPicture(App.path & "\Recursos\" & picture_file_name)
    'Call Delete_File(App.path & "\Recursos\" & picture_file_name)

Else
    Set General_Load_Picture_From_Resource = Nothing
End If

Exit Function



End Function
Public Sub Delete_File(ByVal file_path As String)
'*****************************************************************
'Author: Juan Martín Dotuyo Dodero
'Last Modify Date: 3/03/2005
'Deletes a resource files
'*****************************************************************
    Dim handle As Integer
    Dim data() As Byte
    
    On Error GoTo Error_Handler
    
    'We open the file to delete
    handle = FreeFile
    Open file_path For Binary Access Write Lock Read As handle
    
    'We replace all the bytes in it with 0s
    ReDim data(LOF(handle) - 1)
    Put handle, 1, data
    
    'We close the file
    Close handle
    
    'Now we delete it, knowing that if they retrieve it (some antivirus may create backup copies of deleted files), it will be useless
    Kill file_path
    
    Exit Sub
    
Error_Handler:
    Kill file_path
        
End Sub
Public Function General_Drive_Get_Free_Bytes(ByVal DriveName As String) As Currency
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 6/07/2004
'
'**************************************************************
    Dim RetVal As Long
    Dim FB As Currency
    Dim BT As Currency
    Dim FBT As Currency
    
    RetVal = GetDiskFreeSpace(Left(DriveName, 2), FB, BT, FBT)
    
    General_Drive_Get_Free_Bytes = FB * 10000 'convert result to actual size in bytes
End Function
