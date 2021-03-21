VERSION 5.00
Begin VB.Form frmCarp 
   Appearance      =   0  'Flat
   BackColor       =   &H80000001&
   BorderStyle     =   0  'None
   Caption         =   "Trabajar de carpintero"
   ClientHeight    =   5595
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6525
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5595
   ScaleWidth      =   6525
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
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
      Height          =   3540
      Left            =   700
      TabIndex        =   4
      Top             =   1560
      Width           =   2400
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
      Height          =   1200
      Left            =   3855
      TabIndex        =   3
      Top             =   2950
      Width           =   1605
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
      Height          =   1200
      ItemData        =   "frmCarp.frx":0000
      Left            =   5520
      List            =   "frmCarp.frx":0007
      TabIndex        =   2
      Top             =   2950
      Width           =   525
   End
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
      Left            =   3735
      MaxLength       =   3
      TabIndex        =   1
      Text            =   "1"
      Top             =   4450
      Width           =   660
   End
   Begin VB.PictureBox picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   465
      Left            =   4740
      ScaleHeight     =   32
      ScaleMode       =   0  'User
      ScaleWidth      =   32
      TabIndex        =   0
      Top             =   1770
      Width           =   480
   End
   Begin VB.Label desc 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
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
      Height          =   255
      Left            =   3600
      TabIndex        =   5
      Top             =   2440
      Width           =   2775
   End
   Begin VB.Image Command4 
      Height          =   465
      Left            =   3960
      Tag             =   "0"
      Top             =   5000
      Width           =   2130
   End
   Begin VB.Image Command3 
      Height          =   450
      Left            =   4590
      Tag             =   "0"
      Top             =   4430
      Width           =   1740
   End
End
Attribute VB_Name = "frmCarp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command3_Click()
    
    On Error GoTo Command3_Click_Err
    
    
    'Si el indice seleccionado es -1 es xq no seleccionamos un item de la lista.
    If lstArmas.ListIndex = -1 Then Exit Sub

    If cantidad > 1 Then
    
        UserMacro.cantidad = cantidad
        UserMacro.TIPO = 2
        UserMacro.Index = ObjCarpintero(lstArmas.ListIndex + 1)
        
        Call AddtoRichTextBox(frmMain.RecTxt, "Comienzas a trabajar.", 2, 51, 223, 1, 1)
        
        UserMacro.Activado = True
        frmMain.MacroLadder.Interval = IntervaloTrabajoConstruir
        frmMain.MacroLadder.Enabled = True
    
    Else
    
        Call WriteCraftCarpenter(ObjCarpintero(lstArmas.ListIndex + 1))

        If frmMain.macrotrabajo.Enabled Then
            MacroBltIndex = ObjCarpintero(lstArmas.ListIndex + 1)
        End If
    
    End If

    Unload Me
    
    
    Exit Sub

Command3_Click_Err:
    Call RegistrarError(Err.number, Err.Description, "frmCarp.Command3_Click", Erl)
    Resume Next
    
End Sub

Private Sub Command4_Click()
    
    On Error GoTo Command4_Click_Err
    
    Unload Me
    
    Exit Sub

Command4_Click_Err:
    Call RegistrarError(Err.number, Err.Description, "frmCarp.Command4_Click", Erl)
    Resume Next
    
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    
    On Error GoTo Form_KeyPress_Err
    

    If (KeyAscii = vbKeyEscape) Then Unload Me

    
    Exit Sub

Form_KeyPress_Err:
    Call RegistrarError(Err.number, Err.Description, "frmCarp.Form_KeyPress", Erl)
    Resume Next
    
End Sub

Private Sub Form_Load()
    
    On Error GoTo Form_Load_Err
    
    Call FormParser.Parse_Form(Me)
    
    Exit Sub

Form_Load_Err:
    Call RegistrarError(Err.number, Err.Description, "frmCarp.Form_Load", Erl)
    Resume Next
    
End Sub

Private Sub Command3_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    'Command3.Picture = LoadInterface(Language + "\trabajar_construirpress.bmp")
    'Command3.Tag = "1"
End Sub

Private Sub Command3_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    On Error GoTo Command3_MouseMove_Err
    

    If Command3.Tag = "0" Then
        Command3.Picture = LoadInterface(Language + "\trabajar_construirhover.bmp")
        Command3.Tag = "1"

    End If
    
    Command4.Picture = Nothing
    Command4.Tag = "0"

    
    Exit Sub

Command3_MouseMove_Err:
    Call RegistrarError(Err.number, Err.Description, "frmCarp.Command3_MouseMove", Erl)
    Resume Next
    
End Sub

Private Sub Command4_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    'Command4.Picture = LoadInterface(Language + "\trabajar_salirpress.bmp")
    'Command4.Tag = "1"
End Sub

Private Sub Command4_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    On Error GoTo Command4_MouseMove_Err
    

    If Command4.Tag = "0" Then
        Command4.Picture = LoadInterface(Language + "\trabajar_salirhover.bmp")
        Command4.Tag = "1"

    End If

    Command3.Picture = Nothing
    Command3.Tag = "0"

    
    Exit Sub

Command4_MouseMove_Err:
    Call RegistrarError(Err.number, Err.Description, "frmCarp.Command4_MouseMove", Erl)
    Resume Next
    
End Sub

Private Sub List1_Click()
    
    On Error GoTo List1_Click_Err
    

    

    Dim SR As RECT, DR As RECT

    SR.Left = 0
    SR.Top = 0
    SR.Right = 32
    SR.Bottom = 32

    DR.Left = 0
    DR.Top = 0
    DR.Right = 32
    DR.Bottom = 32
    
    Call Grh_Render_To_Hdc(picture1, 550, 0, 0, False)

    
    Exit Sub

List1_Click_Err:
    Call RegistrarError(Err.number, Err.Description, "frmCarp.List1_Click", Erl)
    Resume Next
    
End Sub

Private Sub lstArmas_Click()
    
    On Error GoTo lstArmas_Click_Err
    

    

    Dim SR As RECT, DR As RECT

    SR.Left = 0
    SR.Top = 0
    SR.Right = 32
    SR.Bottom = 32

    DR.Left = 0
    DR.Top = 0
    DR.Right = 32
    DR.Bottom = 32
    
    Call frmCarp.List1.Clear
    Call frmCarp.List2.Clear
    
    Call frmCarp.List1.AddItem("Leña")
    Call frmCarp.List1.AddItem("Leña elfica")
    Call frmCarp.List2.AddItem(ObjData(ObjCarpintero(lstArmas.ListIndex + 1)).Madera)
    Call frmCarp.List2.AddItem(ObjData(ObjCarpintero(lstArmas.ListIndex + 1)).MaderaElfica)

    desc.Caption = ObjData(ObjCarpintero(lstArmas.ListIndex + 1)).Texto

    Call Grh_Render_To_Hdc(Me.picture1, ObjData(ObjCarpintero(lstArmas.ListIndex + 1)).GrhIndex, 0, 0)
    
    picture1.Visible = True
    
    
    Exit Sub

lstArmas_Click_Err:
    Call RegistrarError(Err.number, Err.Description, "frmCarp.lstArmas_Click", Erl)
    Resume Next
    
End Sub
