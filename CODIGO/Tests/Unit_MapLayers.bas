Attribute VB_Name = "Unit_MapLayers"
' Argentum 20 - Game Client Program
' Copyright (C) 2026 Noland Studios
' Licensed under the GNU Affero General Public License, version 3 or later,
' as described in the repository licence.txt.
Option Explicit

#If UNIT_TEST = 1 Then
Public CaptureWalkableDraws As Boolean
Private overlapPixels(0 To 95, 0 To 95) As Long
Private drawCount As Long
Private coastSubmission As Long
Private bridgeSubmission As Long

Public Function test_suite_map_layers() As Boolean
    Call UnitTesting.RunTest("maplayers_signed_five_layers", Recursos.TestCsmLayers(False, 0))
    Call UnitTesting.RunTest("maplayers_legacy_four_layers", Recursos.TestCsmLayers(True, 0))
    Call UnitTesting.RunTest("maplayers_layer2_walkable_water_explicit_block", Recursos.TestCsmLayers(False, 2))
    Call UnitTesting.RunTest("maplayers_layer3_walkable_water_explicit_block", Recursos.TestCsmLayers(False, 3))
    Call UnitTesting.RunTest("maplayers_overlay_steps_height_priority", TestOverlayBehavior())
    Call UnitTesting.RunTest("maplayers_contiguous_tile_graphics", TestContiguousTiles())
    Call UnitTesting.RunTest("maplayers_installed_resources", Recursos.TestInstalledCsmMaps(App.Path & "\..\Recursos\Mapas"))
    Call UnitTesting.RunTest("maplayers_map78_coast_behind_bridge", TestOverlappingLowerLayers(12, False))
    Call UnitTesting.RunTest("maplayers_map211_coast_behind_bridge", TestOverlappingLowerLayers(86, False))
    Call UnitTesting.RunTest("maplayers_later_column_coast_behind_bridge", TestOverlappingLowerLayers(12, True))
    test_suite_map_layers = True
End Function

Private Function TestOverlayBehavior() As Boolean
    On Error GoTo Fail
    ReDim MapData(1 To 100, 1 To 100)
    TestOverlayBehavior = GetTerrenoDePaso(20, GetWalkableOverlayGraphic(50, 50)) = CONST_AGUA
    MapData(50, 50).Graphic(2).GrhIndex = 12682
    TestOverlayBehavior = TestOverlayBehavior And GetTerrainHeight(50, 50) = 5
    TestOverlayBehavior = TestOverlayBehavior And GetTerrenoDePaso(20, GetWalkableOverlayGraphic(50, 50)) = CONST_PISO
    MapData(50, 50).Graphic(3).GrhIndex = 12683
    TestOverlayBehavior = TestOverlayBehavior And GetWalkableOverlayGraphic(50, 50) = 12683
    TestOverlayBehavior = TestOverlayBehavior And GetTerrainHeight(50, 50) = 10
    TestOverlayBehavior = TestOverlayBehavior And GetTerrenoDePaso(20, GetWalkableOverlayGraphic(50, 50)) = CONST_PISO
    MapData(50, 50).Graphic(2).GrhIndex = 0
    TestOverlayBehavior = TestOverlayBehavior And GetTerrainHeight(50, 50) = 10
    Exit Function
Fail:
    TestOverlayBehavior = False
End Function

Private Function TestContiguousTiles() As Boolean
    On Error GoTo Fail
    ReDim MapData(1 To 100, 1 To 100)
    Dim stride As Long, layer As Long
    stride = VarPtr(MapData(1, 1).Graphic(2)) - VarPtr(MapData(1, 1).Graphic(1))
    TestContiguousTiles = stride > 0
    For layer = 2 To MAP_LAYER_COUNT
        TestContiguousTiles = TestContiguousTiles And _
            VarPtr(MapData(1, 1).Graphic(layer)) - VarPtr(MapData(1, 1).Graphic(layer - 1)) = stride
    Next layer
    stride = VarPtr(MapData(2, 1)) - VarPtr(MapData(1, 1))
    TestContiguousTiles = TestContiguousTiles And stride > 0 And _
        VarPtr(MapData(1, 2)) - VarPtr(MapData(1, 1)) = stride * 100
    Exit Function
Fail:
    TestContiguousTiles = False
End Function
Public Sub RecordWalkableDraw(ByVal grhIndex As Long, ByVal x As Integer, ByVal y As Integer, _
                              ByVal width As Integer, ByVal height As Integer)
    Dim pixelX As Long, pixelY As Long
    drawCount = drawCount + 1
    If grhIndex = 85194 Then coastSubmission = drawCount
    If grhIndex = 12682 Then bridgeSubmission = drawCount
    ' Opaque fixture pixels isolate paint order; positions/dimensions come
    ' from the real Draw_Grh animation-frame, centering and culling path.
    For pixelY = y To y + height - 1
        For pixelX = x To x + width - 1
            If pixelX >= 0 And pixelX <= 95 And pixelY >= 0 And pixelY <= 95 Then
                overlapPixels(pixelX, pixelY) = grhIndex
            End If
        Next pixelX
    Next pixelY
End Sub

Private Function TestOverlappingLowerLayers(ByVal rootX As Integer, ByVal laterColumn As Boolean) As Boolean
    On Error GoTo Fail
    Dim savedMaxGrh As Long, savedCulling As Rect
    Dim coastX As Integer, coastY As Integer, maxX As Integer
    Dim pixelX As Long, pixelY As Long
    savedMaxGrh = MaxGrh
    savedCulling = RenderCullingRect
    ReDim MapData(1 To 100, 1 To 100)
    ReDim GrhData(1 To 85305)
    MaxGrh = 85305
    ' Maps 78 and 211: bridge 12682 at row 31, coast 85305 at row 32.
    ' Coast first frame 85194 is 96x64 and extends north/across columns.
    With GrhData(85305)
        .NumFrames = 1
        ReDim .Frames(1 To 1)
        .Frames(1) = 85194
    End With
    With GrhData(85194)
        .pixelWidth = 96: .pixelHeight = 64
        .TileWidth = 3: .TileHeight = 2
    End With
    With GrhData(12682)
        .NumFrames = 1
        ReDim .Frames(1 To 1)
        .Frames(1) = 12682
        .pixelWidth = 32: .pixelHeight = 32
        .TileWidth = 1: .TileHeight = 1
    End With
    coastX = rootX: coastY = 32: maxX = rootX
    If laterColumn Then
        coastX = rootX + 1: coastY = 31: maxX = coastX
    End If
    MapData(rootX, 31).Graphic(3).GrhIndex = 12682
    MapData(coastX, coastY).Graphic(2).GrhIndex = 85305
    RenderCullingRect.Left = 0: RenderCullingRect.Top = 0
    RenderCullingRect.Right = 96: RenderCullingRect.Bottom = 96
    Erase overlapPixels
    drawCount = 0: coastSubmission = 0: bridgeSubmission = 0
    CaptureWalkableDraws = True
    Call TileEngine_RenderScreen.RenderWalkableMapLayers(rootX, maxX, 31, 32, 32, 32)
    CaptureWalkableDraws = False
    TestOverlappingLowerLayers = drawCount = 2 And coastSubmission = 1 And bridgeSubmission = 2
    ' Every pixel of the bridge must stay visible despite the coast overlap.
    For pixelY = 32 To 63
        For pixelX = 32 To 63
            TestOverlappingLowerLayers = TestOverlappingLowerLayers And overlapPixels(pixelX, pixelY) = 12682
        Next pixelX
    Next pixelY
    MaxGrh = savedMaxGrh
    RenderCullingRect = savedCulling
    Erase GrhData
    Exit Function
Fail:
    CaptureWalkableDraws = False
    MaxGrh = savedMaxGrh
    RenderCullingRect = savedCulling
    Erase GrhData
    TestOverlappingLowerLayers = False
End Function
#End If
