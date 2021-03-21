VERSION 5.00
Begin VB.Form frmGuildAdm 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Lista de clanes registrados"
   ClientHeight    =   5865
   ClientLeft      =   0
   ClientTop       =   -180
   ClientWidth     =   6225
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   391
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   415
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox Filtro 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   545
      TabIndex        =   3
      Top             =   1615
      Width           =   1575
   End
   Begin VB.ListBox GuildsList 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   2565
      ItemData        =   "frmGuildAdm.frx":0000
      Left            =   495
      List            =   "frmGuildAdm.frx":0002
      TabIndex        =   1
      Top             =   2160
      Width           =   4080
   End
   Begin VB.ComboBox Combo1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      ItemData        =   "frmGuildAdm.frx":0004
      Left            =   2280
      List            =   "frmGuildAdm.frx":0011
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   1600
      Width           =   1655
   End
   Begin VB.Label lblClose 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   435
      Left            =   5760
      TabIndex        =   2
      Top             =   0
      Width           =   405
   End
   Begin VB.Image Image3 
      Height          =   420
      Left            =   4645
      Tag             =   "0"
      Top             =   4230
      Width           =   390
   End
   Begin VB.Image Image2 
      Height          =   420
      Left            =   480
      Tag             =   "0"
      Top             =   5040
      Width           =   1950
   End
   Begin VB.Image Image1 
      Height          =   425
      Left            =   4005
      Tag             =   "0"
      Top             =   1560
      Width           =   450
   End
End
Attribute VB_Name = "frmGuildAdm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    
    On Error GoTo Form_KeyDown_Err
    

    If KeyCode = vbKeyEscape Then Unload Me

    
    Exit Sub

Form_KeyDown_Err:
    Call RegistrarError(Err.number, Err.Description, "frmGuildAdm.Form_KeyDown", Erl)
    Resume Next
    
End Sub


Private Sub Form_Load()
    
    On Error GoTo Form_Load_Err
    

    Call FormParser.Parse_Form(Me)
    
    Me.Picture = LoadInterface(Language + "\ventanaclanes.bmp")
    
    
    Combo1.ListIndex = 2

    
    Exit Sub

Form_Load_Err:
    Call RegistrarError(Err.number, Err.Description, "frmGuildAdm.Form_Load", Erl)
    Resume Next
    
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    On Error GoTo Form_MouseMove_Err
    
    If Image1.Tag = "1" Then
        Image1.Picture = Nothing
        Image1.Tag = "0"
    End If
    
    If Image2.Tag = "1" Then
        Image2.Picture = Nothing
        Image2.Tag = "0"
    End If

    If Image3.Tag = "1" Then
        Image3.Picture = Nothing
        Image3.Tag = "0"
    End If
    
    Exit Sub

Form_MouseMove_Err:
    Call RegistrarError(Err.number, Err.Description, "frmGuildAdm.Form_MouseMove", Erl)
    Resume Next
    
End Sub

Private Sub GuildsList_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    On Error GoTo GuildsList_MouseMove_Err
    
    Call Form_MouseMove(Button, Shift, x, y)
    
    Exit Sub

GuildsList_MouseMove_Err:
    Call RegistrarError(Err.number, Err.Description, "frmGuildAdm.GuildsList_MouseMove", Erl)
    Resume Next
    
End Sub

Private Sub Image1_Click()
    
    On Error GoTo Image1_Click_Err
    
    Dim i As Long

    frmGuildAdm.guildslist.Clear
    
    If Not ListaClanes Then Exit Sub
    
    If Len(Filtro.Text) <> 0 Then
        For i = 0 To UBound(ClanesList)

            If Combo1.ListIndex < 2 Then
                If ClanesList(i).Alineacion = Combo1.ListIndex Then
                    If InStr(1, UCase$(ClanesList(i).nombre), UCase$(Filtro.Text)) <> 0 Then
                        Call frmGuildAdm.guildslist.AddItem(ClanesList(i).nombre)
                    End If
                End If
            ElseIf InStr(1, UCase$(ClanesList(i).nombre), UCase$(Filtro.Text)) <> 0 Then
                Call frmGuildAdm.guildslist.AddItem(ClanesList(i).nombre)
            End If
    
        Next i
        
    Else
        For i = 0 To UBound(ClanesList)

            If Combo1.ListIndex < 2 Then
                If ClanesList(i).Alineacion = Combo1.ListIndex Then
                    Call frmGuildAdm.guildslist.AddItem(ClanesList(i).nombre)
    
                End If
    
            Else
                
                Call frmGuildAdm.guildslist.AddItem(ClanesList(i).nombre)
    
            End If
    
        Next i
    End If
    
    Exit Sub

Image1_Click_Err:
    Call RegistrarError(Err.number, Err.Description, "frmGuildAdm.Image1_Click", Erl)
    Resume Next
    
End Sub

Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    On Error GoTo Image1_MouseDown_Err
    
    Image1.Picture = LoadInterface(Language + "\boton-buscar-off.bmp")
    Image1.Tag = "1"
    
    Exit Sub

Image1_MouseDown_Err:
    Call RegistrarError(Err.number, Err.Description, "frmGuildAdm.Image1_MouseDown", Erl)
    Resume Next
    
End Sub

Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    On Error GoTo Image1_MouseMove_Err
    

    If Image1.Tag = "0" Then
        Image1.Picture = LoadInterface(Language + "\boton-buscar-over.bmp")
        Image1.Tag = "1"
    
    End If

    
    Exit Sub

Image1_MouseMove_Err:
    Call RegistrarError(Err.number, Err.Description, "frmGuildAdm.Image1_MouseMove", Erl)
    Resume Next
    
End Sub

Private Sub Image2_Click()
    
    On Error GoTo Image2_Click_Err
    

    If UserEstado = 1 Then 'Muerto

        With FontTypes(FontTypeNames.FONTTYPE_INFO)
            Call ShowConsoleMsg("¡¡Estás muerto!!", .red, .green, .blue, .bold, .italic)

        End With

        Exit Sub

    End If
                   
    Call WriteQuieroFundarClan

    
    Exit Sub

Image2_Click_Err:
    Call RegistrarError(Err.number, Err.Description, "frmGuildAdm.Image2_Click", Erl)
    Resume Next
    
End Sub

Private Sub Image2_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    On Error GoTo Image2_MouseMove_Err
    

    If Image2.Tag = "0" Then
        Image2.Picture = LoadInterface(Language + "\boton-fundar-clan-es-over.bmp")
        Image2.Tag = "1"

    End If

    
    Exit Sub

Image2_MouseMove_Err:
    Call RegistrarError(Err.number, Err.Description, "frmGuildAdm.Image2_MouseMove", Erl)
    Resume Next
    
End Sub

Private Sub Image3_Click()
    
    On Error GoTo Image3_Click_Err
    
    
    'Si nos encontramos con un guild con nombre vacío algo sospechoso está pasando, x las dudas no hacemos nada.
    If Len(guildslist.List(guildslist.ListIndex)) = 0 Then Exit Sub
    
    frmGuildBrief.EsLeader = False
    
    Call WriteGuildRequestDetails(guildslist.List(guildslist.ListIndex))

    
    Exit Sub

Image3_Click_Err:
    Call RegistrarError(Err.number, Err.Description, "frmGuildAdm.Image3_Click", Erl)
    Resume Next
    
End Sub


Private Sub lblClose_Click()
    
    On Error GoTo lblClose_Click_Err
    
    Unload Me
    
    Exit Sub

lblClose_Click_Err:
    Call RegistrarError(Err.number, Err.Description, "frmGuildAdm.lblClose_Click", Erl)
    Resume Next
    
End Sub
