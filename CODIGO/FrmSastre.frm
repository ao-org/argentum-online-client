VERSION 5.00
Begin VB.Form FrmSastre 
   BackColor       =   &H80000007&
   BorderStyle     =   0  'None
   Caption         =   "Trabajar de sastre"
   ClientHeight    =   5670
   ClientLeft      =   0
   ClientTop       =   -75
   ClientWidth     =   6525
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5670
   ScaleWidth      =   6525
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox cantidad 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   250
      Left            =   3750
      MaxLength       =   3
      TabIndex        =   4
      Text            =   "1"
      Top             =   4500
      Width           =   660
   End
   Begin VB.ListBox List2 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   810
      ItemData        =   "FrmSastre.frx":0000
      Left            =   5535
      List            =   "FrmSastre.frx":0007
      TabIndex        =   3
      Top             =   2955
      Width           =   525
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   810
      Left            =   3870
      TabIndex        =   2
      Top             =   2955
      Width           =   1605
   End
   Begin VB.ListBox lstArmas 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   2955
      Left            =   720
      TabIndex        =   1
      Top             =   2160
      Width           =   2400
   End
   Begin VB.PictureBox picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   465
      Left            =   4750
      ScaleHeight     =   32
      ScaleMode       =   0  'User
      ScaleWidth      =   32
      TabIndex        =   0
      Top             =   1750
      Width           =   480
   End
   Begin VB.Label desc 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3960
      TabIndex        =   5
      Top             =   3840
      Width           =   2055
   End
   Begin VB.Image Command3 
      Height          =   450
      Left            =   4590
      Tag             =   "0"
      Top             =   4420
      Width           =   1740
   End
   Begin VB.Image Command4 
      Height          =   465
      Left            =   3960
      Tag             =   "0"
      Top             =   4990
      Width           =   2130
   End
   Begin VB.Image Command2 
      Height          =   585
      Left            =   2100
      Tag             =   "0"
      Top             =   1460
      Width           =   660
   End
   Begin VB.Image Command1 
      Height          =   585
      Left            =   1050
      Tag             =   "0"
      Top             =   1450
      Width           =   660
   End
End
Attribute VB_Name = "FrmSastre"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private indice As Byte

Private Sub Command1_Click()

    Dim i As Byte

    indice = 1
    lstArmas.Clear

    For i = 1 To UBound(SastreRopas())

        If SastreRopas(i).Index = 0 Then Exit For
        lstArmas.AddItem (ObjData(SastreRopas(i).Index).Name)
    Next i

    Command1.Picture = LoadInterface("sastreria_vestimentahover.bmp")
    Command2.Picture = Nothing

End Sub

Private Sub Command2_Click()

    Dim i As Byte

    indice = 2
    lstArmas.Clear

    For i = 1 To UBound(SastreGorros())

        If SastreGorros(i).Index = 0 Then Exit For
        lstArmas.AddItem (ObjData(SastreGorros(i).Index).Name)
    Next i
    
    Command2.Picture = LoadInterface("sastreria_gorroshover.bmp")
    Command1.Picture = Nothing

End Sub

Private Sub Command3_Click()

    On Error Resume Next

    If indice = 1 Then
        If cantidad > 1 Then
            UserMacro.Intervalo = IntervaloTrabajo
            UserMacro.cantidad = cantidad
            UserMacro.TIPO = 3
            UserMacro.Index = SastreRopas(lstArmas.ListIndex + 1).Index
            AddtoRichTextBox frmmain.RecTxt, "Comienzas a trabajar.", 2, 51, 223, 1, 1
            UserMacro.Activado = True
            frmmain.MacroLadder.Interval = IntervaloTrabajo
            frmmain.MacroLadder.Enabled = True
        Else
            Call WriteCraftSastre(SastreRopas(lstArmas.ListIndex + 1).Index)

            If frmmain.macrotrabajo.Enabled Then MacroBltIndex = SastreRopas(lstArmas.ListIndex + 1).Index

        End If

        Unload Me
    ElseIf indice = 2 Then

        If cantidad > 1 Then
            UserMacro.cantidad = cantidad
            UserMacro.TIPO = 3
            UserMacro.Index = SastreGorros(lstArmas.ListIndex + 1).Index
            AddtoRichTextBox frmmain.RecTxt, "Comienzas a trabajar.", 2, 51, 223, 1, 1
            UserMacro.Intervalo = IntervaloTrabajo
            UserMacro.Activado = True
            frmmain.MacroLadder.Interval = IntervaloTrabajo
            frmmain.MacroLadder.Enabled = True
        Else
            Call WriteCraftSastre(SastreGorros(lstArmas.ListIndex + 1).Index)

            If frmmain.macrotrabajo.Enabled Then MacroBltIndex = SastreGorros(lstArmas.ListIndex + 1).Index

        End If

    End If

    Unload Me

End Sub

Private Sub Command4_Click()
    Unload Me

End Sub

Private Sub Command4_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

    'Command4.Picture = LoadInterface("trabajar_salirpress.bmp")
    'Command4.Tag = "1"
End Sub

Private Sub Command4_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

    If Command4.Tag = "0" Then
        Command4.Picture = LoadInterface("trabajar_salirhover.bmp")
        Command4.Tag = "1"

    End If

    Command3.Picture = Nothing
    Command3.Tag = "0"

End Sub

Private Sub Form_Load()
    Call FormParser.Parse_Form(Me)
    indice = 1

End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)

    If (KeyAscii = 27) Then
        Unload Me

    End If

End Sub

Private Sub Command2_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

    'Command2.Picture = LoadInterface("sastreria_gorrospress.bmp")
    ' Command2.Tag = "1"
End Sub

Private Sub Command2_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    ' If Command2.Tag = "0" Then
    '   Command2.Picture = LoadInterface("sastreria_gorroshover.bmp")
    '   Command2.Tag = "1"
    ' End If
    
End Sub

Private Sub Command1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

    'Command1.Picture = LoadInterface("sastreria_vestimentapress.bmp")
    ' Command1.Tag = "1"
End Sub

Private Sub Command3_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

    ' Command3.Picture = LoadInterface("trabajar_construirpress.bmp")
    ' Command3.Tag = "1"
End Sub

Private Sub Command3_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

    If Command3.Tag = "0" Then
        Command3.Picture = LoadInterface("trabajar_construirhover.bmp")
        Command3.Tag = "1"

    End If

    Command4.Picture = Nothing
    Command4.Tag = "0"

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

    If indice <> 1 Then
        Command1.Picture = Nothing
        Command1.Tag = "0"

    End If

    If indice <> 2 Then
        Command2.Picture = Nothing
        Command2.Tag = "0"

    End If

    Command3.Picture = Nothing
    Command3.Tag = "0"
    Command4.Picture = Nothing
    Command4.Tag = "0"

End Sub

Private Sub List1_Click()

    On Error Resume Next

    Dim grh As Long

    If List1.ListIndex = 0 Then
        grh = 697
    ElseIf List1.ListIndex = 1 Then
        grh = 699
    ElseIf List1.ListIndex = 2 Then
        grh = 698

    End If

    Call Grh_Render_To_Hdc(picture1, grh, 0, 0, False)

End Sub

Private Sub lstArmas_Click()

    On Error Resume Next

    List1.Clear
    List2.Clear
    List1.AddItem ("Piel de lobo")
    List1.AddItem ("Piel de oso pardo")
    List1.AddItem ("Piel de oso polar")

    If indice = 1 Then
        Call Grh_Render_To_Hdc(picture1, ObjData(SastreRopas(lstArmas.ListIndex + 1).Index).GrhIndex, 0, 0)
        List2.AddItem (ObjData(SastreRopas(lstArmas.ListIndex + 1).Index).PielLobo)
        List2.AddItem (ObjData(SastreRopas(lstArmas.ListIndex + 1).Index).PielOsoPardo)
        List2.AddItem (ObjData(SastreRopas(lstArmas.ListIndex + 1).Index).PielOsoPolar)
        desc.Caption = "Defensa: " & ObjData(SastreRopas(lstArmas.ListIndex + 1).Index).MinDef & "/" & ObjData(SastreRopas(lstArmas.ListIndex + 1).Index).MaxDef
    ElseIf indice = 2 Then
        Call Grh_Render_To_Hdc(picture1, ObjData(SastreGorros(lstArmas.ListIndex + 1).Index).GrhIndex, 0, 0)
        List2.AddItem (ObjData(SastreGorros(lstArmas.ListIndex + 1).Index).PielLobo)
        List2.AddItem (ObjData(SastreGorros(lstArmas.ListIndex + 1).Index).PielOsoPardo)
        List2.AddItem (ObjData(SastreGorros(lstArmas.ListIndex + 1).Index).PielOsoPolar)
        desc.Caption = "Defensa: " & ObjData(SastreGorros(lstArmas.ListIndex + 1).Index).MinDef & "/" & ObjData(SastreGorros(lstArmas.ListIndex + 1).Index).MaxDef

    End If

End Sub
