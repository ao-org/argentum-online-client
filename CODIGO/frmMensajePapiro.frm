VERSION 5.00
Begin VB.Form frmMensajePapiro 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   9270
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11580
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9270
   ScaleWidth      =   11580
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7335
      Left            =   1560
      TabIndex        =   0
      Top             =   1320
      Width           =   8655
   End
   Begin VB.Image Image1 
      Height          =   495
      Left            =   9840
      Top             =   720
      Width           =   495
   End
End
Attribute VB_Name = "frmMensajePapiro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function GetWindowLong Lib "user32" _
    Alias "GetWindowLongW" ( _
    ByVal hwnd As Long, _
    ByVal nIndex As Long) As Long
    
Private Declare Function SetWindowLong Lib "user32" _
    Alias "SetWindowLongW" ( _
    ByVal hwnd As Long, _
    ByVal nIndex As Long, _
    ByVal dwNewLong As Long) As Long
    
Private Declare Function SetLayeredWindowAttributes Lib "user32" ( _
    ByVal hwnd As Long, _
    ByVal crKey As Long, _
    ByVal bAlpha As Byte, _
    ByVal dwFlags As Long) As Long
    
Private Const GWL_EXSTYLE = -20
Private Const WS_EX_LAYERED = &H80000
Private Const LWA_COLORKEY = &H1&

Private Sub Form_Load()

    Me.Picture = LoadInterface("ventana-papiro.bmp")
    MakeFormTransparent Me, vbBlack
    'Set the Form "transparent by color."

    
    Me.Label1.Caption = "El malvado Gúl Belthor ha vuelto. Sí, el terrible hermanastro menor del rey orco " & _
    "ha escapado de su prisión en las profundidades de la montaña Penthar. Aún no sabemos cómo lo ha hecho o " & _
    "quién ha ayudado al terrible hechicero a escapar, pero eso no es lo importante ahora. Gúl una vez libre, no " & _
    "perdió el tiempo en conjurar su oscura magia. Usando sus poderes logró arrebatar el poder de casi todos los seres " & _
    "de las tierras de Argentum, sólo unos pocos héroes quedan en pie, el resto han visto sus fuerzas ser arrebatadas por " & _
    "una misteriosa magia a lo largo de estos días. Pero eso no es lo peor, lo terrible es lo que el desalmado ha hecho con ese " & _
    " nuevo poder. El nigromante logró escabullirse invisible entre los cementerios de todas las ciudades, reviviendo allí a los " & _
    " anteriores reyes de todas las razas. " & vbNewLine & vbNewLine & " Con este comunicado convoco a los reyes de todas las razas y los representantes de ambas " & _
    " facciones a una reunión de emergencia en el Mesón Hostigado. Es momento de dejar de lado las " & _
    " diferencias para defender nuestro mundo de la oscuridad. Los espero allí." & vbNewLine & "                              Rey Luther "


End Sub

Private Sub Form_LostFocus()
    Unload Me
End Sub

Private Sub Image1_Click()
    Unload Me
End Sub

