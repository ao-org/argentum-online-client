Attribute VB_Name = "modDX8FIFO"
'RevolucionAo 1.0
'Pablo Mercavides
Option Explicit

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
            
            BodyData(i).HeadOffset.x = MisCuerpos(i).HeadOffsetX
            BodyData(i).HeadOffset.y = MisCuerpos(i).HeadOffsetY

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

        If Not Extract_File(Scripts, App.Path & "\..\Recursos\OUTPUT\", "minimap.dat", Windows_Temp_Dir, False) Then
            Err.Description = "¡No se puede cargar el archivo de recurso!"
            GoTo ErrorHandler

        End If
    
        Open Windows_Temp_Dir & "minimap.dat" For Binary Access Read As #handle
    #Else
        Open App.Path & "\..\Recursos\init\minimap.dat" For Binary Access Read As #handle
    #End If

    For count = 1 To MaxGrh

        If GrhData(count).active Then
            Get #handle, , GrhData(count).MiniMap_color

        End If

    Next count
    
    Close #handle
    
    #If Compresion = 1 Then
        Delete_File Windows_Temp_Dir & "minimap.dat"
    #End If
    
    Exit Function

ErrorHandler:
    CargarMiniMap = False
    MsgBox "Error " & Err.Description & " durante la carga de Grh.dat! La carga se ha detenido en GRH: " & count

End Function
