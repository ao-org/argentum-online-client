VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsGraphicalList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Por Javier Podavini (El Yind)
'09/07/2009

Private ListPic As PictureBox
Private lista() As String
Private p_ListIndex As Integer
Private p_ListCount As Integer
Private p_Visible As Boolean
Private p_ForeColor As Long
Private p_Pixel_Alto As Integer
Private p_Scroll As Integer
Private p_OldY As Integer
Private p_DownBarrita As Integer
Private p_ShowBarrita As Boolean
Public VisibleCount As Integer
Public BarraHeight As Integer
Public BarraWidth As Integer


Private Sub Class_Initialize()
p_Visible = True
Clear
End Sub

Public Sub Initialize(ByRef LPic As PictureBox, ByVal ForeColor As Long)
    Set ListPic = LPic
    
    p_ForeColor = ForeColor
    Pixel_Alto = 15
    BarraHeight = 60
    BarraWidth = 10
    
    ListPic.ForeColor = ForeColor
    ListPic.Cls
    
End Sub

Public Property Get List(ByVal Index As Integer) As String
    List = lista(Index)
End Property
Public Property Let List(ByVal Index As Integer, ByVal Valor As String)
    lista(Index) = Valor
    Dibujar
End Property

Public Property Get Visible() As Integer
    Visible = p_Visible
End Property
Public Property Let Visible(ByVal Valor As Integer)
    p_Visible = Valor
    ListPic.Visible = p_Visible
End Property

Public Property Get ListIndex() As Integer
    ListIndex = p_ListIndex
End Property
Public Property Let ListIndex(ByVal Valor As Integer)
    p_ListIndex = Valor
    If p_ListIndex >= p_ListCount Then p_ListIndex = p_ListCount - 1
    'If p_ListIndex + Scroll + 1 > VisibleCount Then Scroll = p_ListIndex + Scroll + 1 - VisibleCount
    Dibujar
End Property

Public Property Get OldY() As Integer
    OldY = p_OldY
End Property
Public Property Let OldY(ByVal Valor As Integer)
    p_OldY = Valor
End Property

Public Property Get DownBarrita() As Integer
    DownBarrita = p_DownBarrita
End Property
Public Property Let DownBarrita(ByVal Valor As Integer)
    p_DownBarrita = Valor
    Dibujar
End Property

Public Property Get Pixel_Alto() As Integer
    Pixel_Alto = p_Pixel_Alto
End Property
Public Property Let Pixel_Alto(ByVal Valor As Integer)
    p_Pixel_Alto = Valor
    VisibleCount = ListPic.ScaleHeight / Valor
End Property

Public Property Get Scroll() As Integer
    Scroll = p_Scroll
End Property

Public Property Get ShowBarrita() As Boolean
    ShowBarrita = p_ShowBarrita
End Property
Public Property Let ShowBarrita(ByVal Valor As Boolean)
    If p_ShowBarrita <> Valor Then
        p_ShowBarrita = Valor
        Dibujar
    End If
End Property
Public Property Let Scroll(ByVal Valor As Integer)
    If Valor > p_ListCount - ListPic.ScaleHeight / p_Pixel_Alto Then Valor = p_ListCount - VisibleCount
    If Valor < 0 Then Valor = 0
    p_Scroll = Valor
    Dibujar
End Property

Public Property Get ListCount() As Integer
    ListCount = p_ListCount
End Property

Public Sub AddItem(Item As String)
ReDim Preserve lista(p_ListCount)
lista(p_ListCount) = Item
p_ListCount = p_ListCount + 1
If p_ListCount = 1 Then p_ListIndex = 0
Dibujar
End Sub
Public Sub Clear()
ReDim lista(0)
p_ListCount = 0
p_ListIndex = -1
p_Scroll = 0
If Not ListPic Is Nothing Then Dibujar
End Sub
Private Sub Dibujar()
Dim i As Integer
Dim PosY As Integer
ListPic.Cls
For i = p_Scroll To p_Scroll + VisibleCount
    If i < p_ListCount Then
        PosY = (i - p_Scroll) * p_Pixel_Alto

        If i = p_ListIndex Then
            ListPic.Line (1, posY)-(ListPic.ScaleWidth - 1, posY + p_Pixel_Alto + 1), &H202020, BF
            ListPic.Line (1, posY)-(ListPic.ScaleWidth - 1, posY + p_Pixel_Alto + 1), &H404040, B
            ListPic.ForeColor = &H828282
        Else
            ListPic.ForeColor = p_ForeColor
        End If
        ListPic.CurrentX = 4
        ListPic.CurrentY = PosY
       
        ListPic.Print lista(i)
    End If
Next i
ListPic.DrawStyle = 0
PosY = 0
If ShowBarrita Then
    If VisibleCount < ListCount Then posY = p_Scroll * (ListPic.ScaleHeight - BarraHeight) / (p_ListCount - VisibleCount)
    Dim barritaColor As Long
    If DownBarrita Then
        barritaColor = &H505050
    Else
        barritaColor = &H404040
    End If
    ListPic.Line (ListPic.ScaleWidth - BarraWidth, posY + 1)-(ListPic.ScaleWidth - 1, posY + BarraHeight - 2), barritaColor, BF
    ListPic.Line (ListPic.ScaleWidth - BarraWidth, posY + 1)-(ListPic.ScaleWidth - 1, posY + BarraHeight - 2), &H505050, B
End If
End Sub

