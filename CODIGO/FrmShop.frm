VERSION 5.00
Begin VB.Form FrmShop 
   BorderStyle     =   0  'None
   Caption         =   "Shop Donador"
   ClientHeight    =   5580
   ClientLeft      =   0
   ClientTop       =   -75
   ClientWidth     =   6495
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5580
   ScaleWidth      =   6495
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   465
      Left            =   4735
      ScaleHeight     =   32
      ScaleMode       =   0  'User
      ScaleWidth      =   32
      TabIndex        =   3
      Top             =   1760
      Width           =   480
   End
   Begin VB.ListBox lstArmas 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   2730
      Left            =   480
      TabIndex        =   0
      Top             =   1350
      Width           =   2840
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Ningún item de esta tienda se cae"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   255
      Left            =   3720
      TabIndex        =   7
      Top             =   3855
      Width           =   2625
   End
   Begin VB.Image Image3 
      Height          =   450
      Left            =   490
      Tag             =   "0"
      Top             =   4790
      Width           =   1020
   End
   Begin VB.Image Image2 
      Height          =   450
      Left            =   1530
      Tag             =   "0"
      Top             =   4770
      Width           =   1785
   End
   Begin VB.Image Image1 
      Height          =   570
      Left            =   630
      Tag             =   "0"
      Top             =   4190
      Width           =   2670
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0 "
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   255
      Left            =   4560
      TabIndex        =   6
      Top             =   2460
      Width           =   975
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Seleccione un item."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   855
      Left            =   3800
      TabIndex        =   5
      Top             =   3120
      Width           =   2500
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Ninguno"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   495
      Left            =   3720
      TabIndex        =   4
      Top             =   2880
      Width           =   2535
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0 días"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Left            =   4200
      TabIndex        =   2
      Top             =   4520
      Width           =   1605
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0 creditos."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Left            =   4080
      TabIndex        =   1
      Top             =   5240
      Width           =   1845
   End
End
Attribute VB_Name = "FrmShop"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' Constantes para SendMessage
Const WM_SYSCOMMAND As Long = &H112&

Const MOUSE_MOVE    As Long = &HF012&

Private Declare Function ReleaseCapture Lib "user32" () As Long

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Long) As Long

Private Sub moverForm()
    
    On Error GoTo moverForm_Err
    

    Dim res As Long

    ReleaseCapture
    res = SendMessage(Me.hwnd, WM_SYSCOMMAND, MOUSE_MOVE, 0)

    
    Exit Sub

moverForm_Err:
    Call RegistrarError(Err.number, Err.Description, "FrmShop.moverForm", Erl)
    Resume Next
    
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    
    On Error GoTo Form_KeyUp_Err
    

    If KeyCode = 27 Then
        Unload Me

    End If

    
    Exit Sub

Form_KeyUp_Err:
    Call RegistrarError(Err.number, Err.Description, "FrmShop.Form_KeyUp", Erl)
    Resume Next
    
End Sub

Private Sub Form_Load()
    
    On Error GoTo Form_Load_Err
    
    Call FormParser.Parse_Form(Me)

    
    Exit Sub

Form_Load_Err:
    Call RegistrarError(Err.number, Err.Description, "FrmShop.Form_Load", Erl)
    Resume Next
    
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    On Error GoTo Form_MouseMove_Err
    
    moverForm

    Image1.Picture = Nothing
    Image1.Tag = "0"
    Image2.Picture = Nothing
    Image2.Tag = "0"

    Image3.Picture = Nothing
    Image3.Tag = "0"

    
    Exit Sub

Form_MouseMove_Err:
    Call RegistrarError(Err.number, Err.Description, "FrmShop.Form_MouseMove", Erl)
    Resume Next
    
End Sub

Private Sub Image1_Click()
    
    On Error GoTo Image1_Click_Err
    

    If lstArmas.ListIndex >= 0 Then
        Call WriteComprarItem(lstArmas.ListIndex + 1)
    Else

        With FontTypes(FontTypeNames.FONTTYPE_INFOIAO)
            Call ShowConsoleMsg("No seleccionaste ningun item.", .red, .green, .blue, .bold, .italic)

        End With

    End If

    
    Exit Sub

Image1_Click_Err:
    Call RegistrarError(Err.number, Err.Description, "FrmShop.Image1_Click", Erl)
    Resume Next
    
End Sub

Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

    'Image1.Picture = LoadInterface("comprarpress.bmp")
    'Image1.Tag = "0"
End Sub

Private Sub Image2_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

    'Image2.Picture = LoadInterface("cargarcodepress.bmp")
    'Image2.Tag = "0"
End Sub

Private Sub Image2_Click()
    
    On Error GoTo Image2_Click_Err
    

    Dim Codigo As String

    Codigo = InputBox("Por favor, inserte el codigo que desea canjear.", "Canje de codigos")

    If Codigo <> "" Then
        Call WriteCodigo(Codigo)

    End If

    
    Exit Sub

Image2_Click_Err:
    Call RegistrarError(Err.number, Err.Description, "FrmShop.Image2_Click", Erl)
    Resume Next
    
End Sub

Private Sub Image2_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    On Error GoTo Image2_MouseMove_Err
    

    If Image2.Tag = "0" Then
        Image2.Picture = LoadInterface("cargarcodehover.bmp")
        Image2.Tag = "1"

    End If
    
    Image1.Picture = Nothing
    Image1.Tag = "0"
    
    Image3.Picture = Nothing
    Image3.Tag = "0"

    
    Exit Sub

Image2_MouseMove_Err:
    Call RegistrarError(Err.number, Err.Description, "FrmShop.Image2_MouseMove", Erl)
    Resume Next
    
End Sub

Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    On Error GoTo Image1_MouseMove_Err
    

    If Image1.Tag = "0" Then
        Image1.Picture = LoadInterface("comprarhover.bmp")
        Image1.Tag = "1"

    End If
    
    Image2.Picture = Nothing
    Image2.Tag = "0"
    
    Image3.Picture = Nothing
    Image3.Tag = "0"
    
    
    Exit Sub

Image1_MouseMove_Err:
    Call RegistrarError(Err.number, Err.Description, "FrmShop.Image1_MouseMove", Erl)
    Resume Next
    
End Sub

Private Sub Image3_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    On Error GoTo Image3_MouseMove_Err
    

    If Image3.Tag = "0" Then
        Image3.Picture = LoadInterface("donar.bmp")
        Image3.Tag = "1"

    End If

    Image1.Picture = Nothing
    Image1.Tag = "0"

    Image2.Picture = Nothing
    Image2.Tag = "0"

    
    Exit Sub

Image3_MouseMove_Err:
    Call RegistrarError(Err.number, Err.Description, "FrmShop.Image3_MouseMove", Erl)
    Resume Next
    
End Sub

Private Sub Image3_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    On Error GoTo Image3_MouseUp_Err
    
    ShellExecute Me.hwnd, "open", "https://www.argentum20.com/creditos/", "", "", 0

    
    Exit Sub

Image3_MouseUp_Err:
    Call RegistrarError(Err.number, Err.Description, "FrmShop.Image3_MouseUp", Erl)
    Resume Next
    
End Sub

Private Sub lstArmas_Click()
    
    On Error GoTo lstArmas_Click_Err
    
    Label1.Caption = ObjData(ObjDonador(lstArmas.ListIndex + 1).Index).Name
    Label2.Caption = ObjData(ObjDonador(lstArmas.ListIndex + 1).Index).Texto
    Label4.Caption = ObjDonador(lstArmas.ListIndex + 1).precio

    Call Grh_Render_To_Hdc(picture1, ObjData(ObjDonador(lstArmas.ListIndex + 1).Index).GrhIndex, 0, 0)
    picture1.Visible = True

    
    Exit Sub

lstArmas_Click_Err:
    Call RegistrarError(Err.number, Err.Description, "FrmShop.lstArmas_Click", Erl)
    Resume Next
    
End Sub
