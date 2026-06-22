Attribute VB_Name = "ModMinimap"
Option Explicit

' DirectX-rendered dot per slot (0=player, 1-5=allies) used when CenteredMinimap <> 0
Public Type MinimapDotState
    visible  As Boolean
    screenX  As Long
    screenY  As Long
    dotColor(0 To 3) As RGBA
End Type
Public MinimapDots(0 To 5) As MinimapDotState

Public Sub DrawCenteredMinimap()
    Call RenderMinimapCentered(UserMap, UserPos.x, UserPos.y, CenteredMinimapZoom, CenteredMinimapZoom)
End Sub

Public Sub RenderMinimapCentered(ByVal currentMap As Integer, ByVal tileX As Integer, ByVal tileY As Integer, Optional ByVal viewDeltaW As Long = 0, Optional ByVal viewDeltaH As Long = 0)
    On Error GoTo RenderMinimap_Err
    Static colorsInit As Boolean
    If Not colorsInit Then
        Call InitMinimapDotColors
        colorsInit = True
    End If
    Dim i        As Integer
    Dim j        As Byte
    Dim idmap    As Integer
    Dim worldNum As Byte
    Dim mapGridX As Long, mapGridY As Long
    ' Find which world and grid position the current map is in
    worldNum = 0
    idmap = 0
    For j = 1 To TotalWorlds
        For i = 1 To Mundo(j).Ancho * Mundo(j).Alto
            If Mundo(j).MapIndice(i) = currentMap Then
                idmap = i
                worldNum = j
                Exit For
            End If
        Next i
        If idmap > 0 Then Exit For
    Next j
    If idmap = 0 Then Exit Sub
    ' Ensure destination units are pixels
    frmMain.MiniMap.ScaleMode = vbPixels
    ' Load/cached world texture into the DirectX surface manager
    Static lastWorld    As Byte
    Static worldFileNum As Integer
    Static bmpPxW       As Long
    Static bmpPxH       As Long
        Dim minimapFile As String
        Select Case worldNum
            Case 1: minimapFile = "mapa1_200x200.bmp"
            Case 2: minimapFile = "mapa2_200x200.bmp"
            Case Else: minimapFile = ""
        End Select
        If minimapFile = "" Then Exit Sub
        worldFileNum = -CInt(worldNum)
        Call SurfaceDB.GetInterfaceTexture(minimapFile, bmpPxW, bmpPxH)
        If bmpPxW = 0 Or bmpPxH = 0 Then
            worldFileNum = 0
            Exit Sub
        End If
        lastWorld = worldNum
    If worldFileNum = 0 Then Exit Sub
    ' Grid of maps in the world image
    Dim mapCellsX As Long, mapCellsY As Long
    mapCellsX = Mundo(worldNum).Ancho   ' e.g., 100
    mapCellsY = Mundo(worldNum).Alto    ' e.g., 100
    ' Size of one map cell in pixels on the world image
    Dim mapCellPxW As Double: mapCellPxW = bmpPxW / mapCellsX
    Dim mapCellPxH As Double: mapCellPxH = bmpPxH / mapCellsY
    ' Current map's grid coordinates on the world image
    mapGridX = (idmap - 1) Mod mapCellsX
    mapGridY = (idmap - 1) \ mapCellsX
    ' Usable tile ranges inside a map
    ' Usable tile ranges inside a map (using module-level constants)
    Dim tileCountX   As Long, tileCountY As Long
    tileCountX = MINIMAP_TILE_COUNT_X ' 74 tiles
    tileCountY = MINIMAP_TILE_COUNT_Y ' 80 tiles
    ' Clamp incoming tile to valid range, just in case
    If tileX < MINIMAP_MIN_TILE_X Then tileX = MINIMAP_MIN_TILE_X
    If tileX > MINIMAP_MAX_TILE_X Then tileX = MINIMAP_MAX_TILE_X
    If tileY < MINIMAP_MIN_TILE_Y Then tileY = MINIMAP_MIN_TILE_Y
    If tileY > MINIMAP_MAX_TILE_Y Then tileY = MINIMAP_MAX_TILE_Y
    ' Player pixel center on the world image (tile center = offset + 0.5)
    Dim centerPxX As Long: centerPxX = CLng((mapGridX + (tileX - MINIMAP_MIN_TILE_X + 0.5) / tileCountX) * mapCellPxW)
    Dim centerPxY As Long: centerPxY = CLng((mapGridY + (tileY - MINIMAP_MIN_TILE_Y + 0.5) / tileCountY) * mapCellPxH)
    ' Destination size (control size)
    Dim destW As Long, destH As Long
    destW = frmMain.MiniMap.ScaleWidth
    destH = frmMain.MiniMap.ScaleHeight
    ' Clamp the configurable deltas to [-50, +50]
    If viewDeltaW < -50 Then viewDeltaW = -50
    If viewDeltaW > 50 Then viewDeltaW = 50
    If viewDeltaH < -50 Then viewDeltaH = -50
    If viewDeltaH > 50 Then viewDeltaH = 50
    ' Source crop size (zoom). Start from dest size and adjust by deltas.
    ' Smaller src => zoom in; larger src => zoom out. Keep sane minimum.
    Dim srcW As Long, srcH As Long
    srcW = destW + viewDeltaW
    srcH = destH + viewDeltaH
    ' Ensure positive and not exceeding bitmap
    If srcW < 16 Then srcW = 16           ' minimum crop width
    If srcH < 16 Then srcH = 16           ' minimum crop height
    If srcW > bmpPxW Then srcW = bmpPxW
    If srcH > bmpPxH Then srcH = bmpPxH
    ' Source top-left so that the player is centered in the source crop
    Dim srcX As Long, srcY As Long
    srcX = centerPxX - srcW \ 2
    srcY = centerPxY - srcH \ 2
    ' Clamp to bitmap bounds based on source crop size
    If srcX < 0 Then srcX = 0
    If srcY < 0 Then srcY = 0
    If srcX > (bmpPxW - srcW) Then srcX = bmpPxW - srcW
    If srcY > (bmpPxH - srcH) Then srcY = bmpPxH - srcH
    ' Draw: use DirectX 8 rendering instead of slow PaintPicture
    
    Dim DestRect As RECT
    
    
    DestRect.Bottom = destH
    DestRect.Right = destW
    DestRect.Left = 0
    DestRect.Top = 0
    
    MinimapVP_SrcX = srcX
    MinimapVP_SrcY = srcY
    MinimapVP_SrcW = srcW
    MinimapVP_SrcH = srcH
    MinimapVP_DestW = destW
    MinimapVP_DestH = destH
    MinimapVP_MapGridX = mapGridX
    MinimapVP_MapGridY = mapGridY
    MinimapVP_CellPxW = mapCellPxW
    MinimapVP_CellPxH = mapCellPxH
    
    Call Engine_BeginScene
    
    Call Batch_Textured_Box_File(0, 0, srcW, srcH, srcX, srcY, minimapFile, COLOR_WHITE, False, 0, CSng(destW) / CSng(srcW), CSng(destH) / CSng(srcH))
       
    Call RenderMinimapDots
       
    Call Engine_EndScene(DestRect, frmMain.MiniMap.hWnd)
    
    Exit Sub
RenderMinimap_Err:
    Call RegistrarError(Err.Number, Err.Description, "ModMinimap.RenderMinimapCentered", Erl)
End Sub

Public Sub DibujarMiniMapa()
    On Error GoTo DibujarMiniMapa_Err
    If CenteredMinimap = 0 Then
        ' Old system: load individual 100x100 map images
        frmMain.MiniMap.Picture = LoadMinimap(ResourceMap)
    End If
    ' Paint NPCs on minimap
    If ListNPCMapData(ResourceMap).NpcCount > 0 And CenteredMinimap = 0 Then
        Dim i As Long
        For i = 1 To MAX_QUESTNPCS_VISIBLE
            Dim PosX As Long
            Dim PosY As Long
            PosX = ListNPCMapData(ResourceMap).NpcList(i).Position.x
            PosY = ListNPCMapData(ResourceMap).NpcList(i).Position.y
            Dim Color As Long
            Select Case ListNPCMapData(ResourceMap).NpcList(i).state
                Case 1
                    Color = RGB(0, 198, 254)
                Case 2
                    Color = RGB(255, 201, 14)
                Case Else
                    Color = RGB(255, 201, 14)
            End Select
            Call SetPixel(frmMain.MiniMap.hdc, PosX + 1, PosY, Color)
            Call SetPixel(frmMain.MiniMap.hdc, PosX, PosY + 1, Color)
            Call SetPixel(frmMain.MiniMap.hdc, PosX + 1, PosY + 1, Color)
            Call SetPixel(frmMain.MiniMap.hdc, PosX, PosY, Color)
            Call SetPixel(frmMain.MiniMap.hdc, PosX, PosY - 1, &H808080)
            Call SetPixel(frmMain.MiniMap.hdc, PosX + 1, PosY - 1, &H808080)
            Call SetPixel(frmMain.MiniMap.hdc, PosX + 2, PosY, &H808080)
            Call SetPixel(frmMain.MiniMap.hdc, PosX + 2, PosY + 1, &H808080)
            Call SetPixel(frmMain.MiniMap.hdc, PosX + 1, PosY + 2, &H808080)
            Call SetPixel(frmMain.MiniMap.hdc, PosX, PosY + 2, &H808080)
            Call SetPixel(frmMain.MiniMap.hdc, PosX - 1, PosY + 1, &H808080)
            Call SetPixel(frmMain.MiniMap.hdc, PosX - 1, PosY, &H808080)
        Next i
        frmMain.MiniMap.Refresh
    End If
    Exit Sub
DibujarMiniMapa_Err:
    Call RegistrarError(Err.Number, Err.Description, "ModMinimap.DibujarMiniMapa", Erl)
End Sub

Public Sub RenderMinimapDots()
    On Error GoTo RenderMinimapDots_Err
    
    Dim i As Long
    Dim squareSize As Long
    squareSize = 6  ' Size of each square in pixels (adjust as needed)
    
    ' Loop through all minimap dots
    For i = 0 To 5
        If MinimapDots(i).visible Then
            Dim Col As RGBA
            ' Draw the square using the dot's position and color
            Call Batch_Textured_Box_File( _
                MinimapDots(i).screenX - (squareSize \ 2), _
                MinimapDots(i).screenY - (squareSize \ 2), _
                1, 1, 0, 0, _
                "white_pixel.png", _
                MinimapDots(i).dotColor, _
                False, 0, _
                squareSize, squareSize)
        End If
    Next i
    
    Exit Sub
    
RenderMinimapDots_Err:
    Call RegistrarError(Err.Number, Err.Description, "ModUtils.RenderMinimapDots", Erl)
End Sub


Public Sub InitMinimapDotColors()
    SetRGBA MinimapDots(0).dotColor(0), 255, 255, 255     ' player  : white
    SetRGBA MinimapDots(0).dotColor(1), 255, 255, 255
    SetRGBA MinimapDots(0).dotColor(2), 255, 255, 255
    SetRGBA MinimapDots(0).dotColor(3), 255, 255, 255
    
    SetRGBA MinimapDots(1).dotColor(0), 255, 255, 0      ' ally 1  : yellow
    SetRGBA MinimapDots(1).dotColor(1), 255, 255, 0
    SetRGBA MinimapDots(1).dotColor(2), 255, 255, 0
    SetRGBA MinimapDots(1).dotColor(3), 255, 255, 0
    
    SetRGBA MinimapDots(2).dotColor(0), 0, 192, 0        ' ally 2  : green
    SetRGBA MinimapDots(2).dotColor(1), 0, 192, 0
    SetRGBA MinimapDots(2).dotColor(2), 0, 192, 0
    SetRGBA MinimapDots(2).dotColor(3), 255, 255, 0
    
    SetRGBA MinimapDots(3).dotColor(0), 255, 128, 0      ' ally 3  : orange
    SetRGBA MinimapDots(3).dotColor(1), 255, 128, 0
    SetRGBA MinimapDots(3).dotColor(2), 255, 128, 0
    SetRGBA MinimapDots(3).dotColor(3), 255, 255, 0
    
    SetRGBA MinimapDots(4).dotColor(0), 255, 0, 255      ' ally 4  : magenta
    SetRGBA MinimapDots(4).dotColor(1), 255, 0, 255
    SetRGBA MinimapDots(4).dotColor(2), 255, 0, 255
    SetRGBA MinimapDots(4).dotColor(3), 255, 255, 0
    
    SetRGBA MinimapDots(5).dotColor(0), 0, 0, 255        ' ally 5  : blue
    SetRGBA MinimapDots(5).dotColor(1), 0, 0, 255
    SetRGBA MinimapDots(5).dotColor(2), 0, 0, 255
    SetRGBA MinimapDots(5).dotColor(3), 255, 255, 0
End Sub
