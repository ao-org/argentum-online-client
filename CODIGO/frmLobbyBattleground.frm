VERSION 5.00
Begin VB.Form frmLobbyBattleground 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Lobby Battleground"
   ClientHeight    =   5985
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   13800
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmLobbyBattleground.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   399
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   920
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox pEvents 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Height          =   4200
      Left            =   120
      ScaleHeight     =   280
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   905
      TabIndex        =   2
      Top             =   1080
      Width           =   13575
   End
   Begin VB.CommandButton btnCrear 
      Caption         =   "Crear Evento"
      Height          =   495
      Left            =   12000
      TabIndex        =   0
      Top             =   5400
      Width           =   1695
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Lobby Battleground"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   0
      TabIndex        =   1
      Top             =   120
      Width           =   13695
   End
End
Attribute VB_Name = "frmLobbyBattleground"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim MouseIndex  As Integer
Dim Scroll      As Integer
Const MAX_LIST  As Integer = 6
Dim LobbyList() As t_LobbyData
Dim LobbyLoaded As Boolean

Private Sub btnCrear_Click()
    Unload Me
    frmCreateBattleground.Show
End Sub

Private Sub Form_Load()
    LobbyLoaded = False
    ListRefresh
End Sub

Friend Sub SetLobbyList(ByRef List() As t_LobbyData)
    LobbyList = List
    LobbyLoaded = True
    Scroll = 0
    ListRefresh
End Sub

Private Sub pEvents_Click()
    On Error GoTo errhandler:
    If Not LobbyLoaded Then Exit Sub
    
    Dim Password As String
    If MouseIndex > 0 And MouseIndex <= UBound(LobbyList) Then
        If LobbyList(MouseIndex).IsPrivate Then
            Password = InputBox(JsonLanguage.Item("MENSAJE_EVENTO_CONTRASEÑA"))
            If Len(Password) = 0 Then Exit Sub
        End If
        Call WriteParticipar(LobbyList(MouseIndex).id, Password)
        Unload Me
    ElseIf MouseIndex = -1 Then
        Scroll = Scroll - 1
        If Scroll < 0 Then Scroll = 0
    ElseIf MouseIndex = -2 Then
        Scroll = Scroll + 1
        Dim maxScroll As Integer
        maxScroll = UBound(LobbyList) + 1 - MAX_LIST
        If maxScroll < 0 Then maxScroll = 0
        If Scroll > maxScroll Then Scroll = maxScroll
    End If
    ListRefresh
    Exit Sub
errhandler:
    Call RegistrarError(Err.Number, Err.Description, "frmLobbyBattleground.Click", Erl)
    Resume Next
End Sub

Private Sub pEvents_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Not LobbyLoaded Then Exit Sub
    
    Dim i As Integer
    Dim prevIndex As Integer
    prevIndex = MouseIndex
    MouseIndex = 0
    
    For i = 1 To MAX_LIST
        If x >= 0 And x <= pEvents.Width And y >= i * 45 And y <= 35 + i * 45 Then
            If i + Scroll <= UBound(LobbyList) Then
                MouseIndex = i + Scroll
            End If
            Exit For
        End If
    Next i
    
    If x >= pEvents.Width - 15 And x <= pEvents.Width - 5 And y >= 45 And y <= 50 Then
        MouseIndex = -1
    End If
    If x >= pEvents.Width - 15 And x <= pEvents.Width - 5 And y >= pEvents.Height - 15 And y <= pEvents.Height - 10 Then
        MouseIndex = -2
    End If
    
    If MouseIndex <> prevIndex Then ListRefresh
End Sub

Private Sub ListRefresh()
    On Error GoTo ErrHandler
    Dim OffX As Integer
    Dim Offy As Integer

    With pEvents
        pEvents.Cls
        .ForeColor = vbWhite
        .FontSize = 10
        OffX = 10
        Offy = 10

        'encabezados
        .currentX = OffX: .currentY = Offy: pEvents.Print "Descripción"
        OffX = OffX + 140
        .currentX = OffX: .currentY = Offy: pEvents.Print "Tipo"
        OffX = OffX + 100
        .currentX = OffX: .currentY = Offy: pEvents.Print "Tamaño de grupos"
        OffX = OffX + 130
        .currentX = OffX: .currentY = Offy: pEvents.Print "Formato de equipos"
        OffX = OffX + 130
        .currentX = OffX: .currentY = Offy: pEvents.Print "Límite de nivel"
        OffX = OffX + 130
        .currentX = OffX: .currentY = Offy: pEvents.Print "Jugadores"
        OffX = OffX + 100
        .currentX = OffX: .currentY = Offy: pEvents.Print "Valor"

        pEvents.Line (0, 40)-(.Width, 40)

        If Not LobbyLoaded Then
            .currentX = 10: .currentY = 60
            pEvents.Print "Cargando eventos..."
            GoTo DrawArrows
        End If
        
        Dim i As Integer
        For i = 1 To MAX_LIST
            If i + Scroll > UBound(LobbyList) Then Exit For
            Offy = 10 + i * 45
            OffX = 10
            .currentX = OffX: .currentY = Offy
            pEvents.Print LobbyList(i + Scroll).Description
            OffX = OffX + 140
            .currentX = OffX: .currentY = Offy
            pEvents.Print LobbyList(i + Scroll).ScenarioType
            OffX = OffX + 100
            .currentX = OffX: .currentY = Offy
            pEvents.Print LobbyList(i + Scroll).TeamSize
            OffX = OffX + 130
            .currentX = OffX: .currentY = Offy
            Select Case LobbyList(i + Scroll).TeamType
                Case e_TeamTypes.eRandom: pEvents.Print "Aleatorio"
                Case e_TeamTypes.ePremade: pEvents.Print "Grupos"
            End Select
            OffX = OffX + 130
            .currentX = OffX: .currentY = Offy
            pEvents.Print LobbyList(i + Scroll).MinLevel & "/" & LobbyList(i + Scroll).MaxLevel
            OffX = OffX + 130
            .currentX = OffX: .currentY = Offy
            pEvents.Print LobbyList(i + Scroll).RegisteredPlayers & "/" & LobbyList(i + Scroll).MaxPlayers
            OffX = OffX + 100
            .currentX = OffX: .currentY = Offy
            pEvents.Print LobbyList(i + Scroll).InscriptionPrice
            OffX = OffX + 80
            
            If MouseIndex = i + Scroll Then
                pEvents.ForeColor = RGB(45, 45, 45)
                pEvents.Line (OffX - 10, Offy - 10)-(OffX + 60, Offy + 25), , BF
                pEvents.ForeColor = vbWhite
                pEvents.Line (OffX - 10, Offy - 10)-(OffX + 60, Offy + 25), , B
            Else
                pEvents.Line (OffX - 10, Offy - 10)-(OffX + 60, Offy + 25), , B
            End If
            .currentX = OffX: .currentY = Offy
            pEvents.Print IIf(LobbyList(i + Scroll).IsPrivate, "* Ingresar", "Ingresar")
        Next i
        
DrawArrows:
        
        .ForeColor = IIf(MouseIndex = -2, RGB(200, 0, 0), vbWhite)
        pEvents.Line (pEvents.Width - 15, pEvents.Height - 15)-(pEvents.Width - 10, pEvents.Height - 10)
        pEvents.Line (pEvents.Width - 10, pEvents.Height - 10)-(pEvents.Width - 4, pEvents.Height - 15)

        .ForeColor = IIf(MouseIndex = -1, RGB(200, 0, 0), vbWhite)
        pEvents.Line (pEvents.Width - 15, 50)-(pEvents.Width - 10, 45)
        pEvents.Line (pEvents.Width - 10, 45)-(pEvents.Width - 4, 50)

        .ForeColor = vbWhite
    End With
    Exit Sub
errhandler:
    Call RegistrarError(Err.Number, Err.Description, "frmLobbyBattleground.Paint", Erl)
    Resume Next
End Sub
