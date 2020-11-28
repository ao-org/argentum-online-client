Attribute VB_Name = "ModVerticesPrecalculados"
Public Sub PrecalcularVertices(ByVal TilesAncho As Byte, ByVal TilesAlto As Byte)

    ReDim m_Data(Capacity * 4 - 1) As TYPE_VERTEX
   
    '
    '  Create the vertice buffer
    '
    Set m_VBuffer = DirectDevice.CreateVertexBuffer(24 * 4 * Capacity, D3DUSAGE_DYNAMIC, D3DFVF_XYZ Or D3DFVF_DIFFUSE Or D3DFVF_TEX1, D3DPOOL_DEFAULT)
 
    '
    '  Create the indice buffer, fill it with pre-baked indices
    '
    Set m_IBuffer = DirectDevice.CreateIndexBuffer(12 * 4 * Capacity, D3DUSAGE_WRITEONLY, D3DFMT_INDEX16, D3DPOOL_DEFAULT)
   
    Dim lpIndices() As Integer
   
    ReDim lpIndices(Capacity * 6 - 1) As Integer
   
    Dim i As Long, j As Integer
   
    For i = 0 To UBound(lpIndices) Step 6
        lpIndices(i) = j
        lpIndices(i + 1) = j + 1
        lpIndices(i + 2) = j + 2
        lpIndices(i + 3) = j + 2
        lpIndices(i + 4) = j + 3
        lpIndices(i + 5) = j
       
        j = j + 4
    Next
   
    Call D3DIndexBuffer8SetData(m_IBuffer, 0, UBound(lpIndices), 0, lpIndices(0))
       
End Sub
