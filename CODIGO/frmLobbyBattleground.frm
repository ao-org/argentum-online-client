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
Dim MouseIndex As Integer
Dim Scroll As Integer
Const MAX_LIST As Integer = 6

Private Sub btnCrear_Click()
    frmCreateBattleground.Show
    Unload Me
End Sub

Private Sub Form_Load()
    ListRefresh
End Sub
Private Sub pEvents_Click()
On Error GoTo ErrHandler:
    Dim Password As String
    If MouseIndex > 0 And MouseIndex <= UBound(LobbyList) Then
        If LobbyList(MouseIndex).IsPrivate Then
            Password = InputBox("Este evento tiene contraseña, por favor ingresala:")
            If Len(Password) = 0 Then Exit Sub
        End If
        Call WriteParticipar(LobbyList(MouseIndex).id, Password)
        Unload Me
    ElseIf MouseIndex = -1 Then
        Scroll = Scroll - 1
        If Scroll < 0 Then Scroll = 0
    ElseIf MouseIndex = -2 Then
        Scroll = Scroll + 1
        If Scroll + MAX_LIST > UBound(LobbyList) Then Scroll = UBound(LobbyList) - MAX_LIST
        If Scroll < 0 Then Scroll = 0
    End If
    Exit Sub
ErrHandler:
    Call RegistrarError(Err.Number, Err.Description, "frmLobbyBattleground.Click", Erl)
    Resume Next
End Sub

Private Sub pEvents_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim i As Integer
    MouseIndex = 0
    For i = 1 To MAX_LIST
        If X >= 810 And X <= 810 + 70 And Y >= i * 45 And Y <= 35 + i * 45 Then
            MouseIndex = i + Scroll
            Exit For
        End If
    Next i
    If X >= pEvents.Width - 15 And X <= pEvents.Width - 5 And Y >= 45 And Y <= 50 Then
        MouseIndex = -1
    End If
    If X >= pEvents.Width - 15 And X <= pEvents.Width - 5 And Y >= pEvents.Height - 15 And Y <= pEvents.Height - 10 Then
        MouseIndex = -2
    End If
    ListRefresh
End Sub

Private Sub ListRefresh()
On Error GoTo ErrHandler:
    Dim OffX As Integer
    Dim Offy As Integer
    With pEvents
        pEvents.Cls
        .ForeColor = vbWhite
        .FontSize = 10
        OffX = 10
        Offy = 10
        
        .CurrentX = OffX
        .CurrentY = Offy
        pEvents.Print "Descripción"
        
        OffX = OffX + 140
        .CurrentX = OffX
        .CurrentY = Offy
        pEvents.Print "Tipo"
        
        OffX = OffX + 100
        .CurrentX = OffX
        .CurrentY = Offy
        pEvents.Print "Tamaño de grupos"
        
        OffX = OffX + 130
        .CurrentX = OffX
        .CurrentY = Offy
        pEvents.Print "Formato de equipos"
        
        OffX = OffX + 130
        .CurrentX = OffX
        .CurrentY = Offy
        pEvents.Print "Límite de nivel"
        
        OffX = OffX + 130
        .CurrentX = OffX
        .CurrentY = Offy
        pEvents.Print "Jugadores"
        
        OffX = OffX + 100
        .CurrentX = OffX
        .CurrentY = Offy
        pEvents.Print "Valor"
        
        pEvents.Line (0, 40)-(.Width, 40)
        
        Dim i As Integer
        For i = 1 To MAX_LIST
            If i + Scroll > UBound(LobbyList) Then Exit For
            
            Offy = Offy + 45
            OffX = 10
            .CurrentX = OffX
            .CurrentY = Offy
            pEvents.Print LobbyList(i + Scroll).Description
            
            OffX = OffX + 140
            .CurrentX = OffX
            .CurrentY = Offy
            pEvents.Print LobbyList(i + Scroll).ScenarioType
            
            OffX = OffX + 100
            .CurrentX = OffX
            .CurrentY = Offy
            pEvents.Print LobbyList(i + Scroll).TeamSize
            
            OffX = OffX + 130
            .CurrentX = OffX
            .CurrentY = Offy
            Select Case LobbyList(i + Scroll).TeamType
                Case 1
                    pEvents.Print "Aleatorio"
                Case 2
                    pEvents.Print "Grupos"
            End Select
            
            
            OffX = OffX + 130
            .CurrentX = OffX
            .CurrentY = Offy
            pEvents.Print LobbyList(i + Scroll).MinLevel & "/" & LobbyList(i + Scroll).MaxLevel
            
            OffX = OffX + 130
            .CurrentX = OffX
            .CurrentY = Offy
            pEvents.Print LobbyList(i + Scroll).RegisteredPlayers & "/" & LobbyList(i + Scroll).MaxPlayers
            
            OffX = OffX + 100
            .CurrentX = OffX
            .CurrentY = Offy
            pEvents.Print LobbyList(i + Scroll).InscriptionPrice
            
            OffX = OffX + 80
            If MouseIndex = i Then
                pEvents.ForeColor = RGB(45, 45, 45)
                pEvents.Line (OffX - 10, Offy - 10)-(OffX + 60, Offy + 25), , BF
                pEvents.ForeColor = vbWhite
                pEvents.Line (OffX - 10, Offy - 10)-(OffX + 60, Offy + 25), , B
                
            Else
                pEvents.Line (OffX - 10, Offy - 10)-(OffX + 60, Offy + 25), , B
            End If
            .CurrentX = OffX
            .CurrentY = Offy
            pEvents.Print IIf(LobbyList(i + Scroll).IsPrivate, "* Ingresar", "Ingresar")
        Next i
        
        .ForeColor = vbWhite
        If MouseIndex = -2 Then .ForeColor = RGB(200, 0, 0)
        pEvents.Line (pEvents.Width - 15, pEvents.Height - 15)-(pEvents.Width - 15 + 5, pEvents.Height - 15 + 5)
        pEvents.Line (pEvents.Width - 15 + 5, pEvents.Height - 15 + 5)-(pEvents.Width - 15 + 11, pEvents.Height - 16)
        
        .ForeColor = vbWhite
        If MouseIndex = -1 Then .ForeColor = RGB(200, 0, 0)
        pEvents.Line (pEvents.Width - 15, 50)-(pEvents.Width - 15 + 5, 50 - 5)
        pEvents.Line (pEvents.Width - 15 + 5, 50 - 5)-(pEvents.Width - 15 + 11, 51)
        .ForeColor = vbWhite
    End With

    Exit Sub
ErrHandler:
    Call RegistrarError(Err.Number, Err.Description, "frmLobbyBattleground.Paint", Erl)
    Resume Next
End Sub
