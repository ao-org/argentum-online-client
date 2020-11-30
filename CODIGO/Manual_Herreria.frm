VERSION 5.00
Begin VB.Form Manual_Herreria 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Manual"
   ClientHeight    =   6810
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6255
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6810
   ScaleWidth      =   6255
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Caption         =   "Herreria"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6735
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   6015
      Begin VB.CommandButton Command2 
         Caption         =   "Indice"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2400
         TabIndex        =   24
         Top             =   6120
         Width           =   1215
      End
      Begin VB.ListBox List1 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3930
         Left            =   240
         TabIndex        =   23
         Top             =   1200
         Width           =   2655
      End
      Begin VB.Frame Frame1 
         Caption         =   "Informacion:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4095
         Left            =   3000
         TabIndex        =   1
         Top             =   1080
         Width           =   2895
         Begin VB.Frame Frame3 
            Caption         =   "Item:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   855
            Left            =   2040
            TabIndex        =   2
            Top             =   2760
            Width           =   735
            Begin VB.PictureBox picture1 
               Appearance      =   0  'Flat
               BackColor       =   &H00000000&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   465
               Left            =   120
               ScaleHeight     =   32
               ScaleMode       =   0  'User
               ScaleWidth      =   32
               TabIndex        =   3
               Top             =   240
               Width           =   480
            End
         End
         Begin VB.Label Label3 
            Caption         =   "-"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   855
            Left            =   120
            TabIndex        =   22
            Top             =   2880
            Width           =   1935
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nombre: "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   120
            TabIndex        =   21
            Top             =   240
            Width           =   750
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Precio:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   120
            TabIndex        =   20
            Top             =   720
            Width           =   570
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Lingotes:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   120
            TabIndex        =   19
            Top             =   1440
            Width           =   765
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Skill manualidades necesario:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   120
            TabIndex        =   18
            Top             =   2160
            Width           =   2490
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nombre:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   120
            TabIndex        =   17
            Top             =   480
            Width           =   615
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "precio"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   800
            TabIndex        =   16
            Top             =   720
            Width           =   435
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "raices"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   1080
            TabIndex        =   15
            Top             =   1920
            Width           =   420
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "raices"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   240
            TabIndex        =   14
            Top             =   2400
            Width           =   540
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Clases permitidas:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   120
            TabIndex        =   13
            Top             =   2640
            Width           =   1545
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "raices"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   360
            TabIndex        =   12
            Top             =   1920
            Width           =   420
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Hierro:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   240
            TabIndex        =   11
            Top             =   1680
            Width           =   570
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Defensa:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   120
            TabIndex        =   10
            Top             =   960
            Width           =   735
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "40"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   920
            TabIndex        =   9
            Top             =   960
            Width           =   180
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Min Nivel:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   120
            TabIndex        =   8
            Top             =   1200
            Width           =   795
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "40"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   960
            TabIndex        =   7
            Top             =   1200
            Width           =   180
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Plata:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   1080
            TabIndex        =   6
            Top             =   1680
            Width           =   480
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "raices"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   1920
            TabIndex        =   5
            Top             =   1920
            Width           =   420
         End
         Begin VB.Label Label21 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Oro:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   1920
            TabIndex        =   4
            Top             =   1680
            Width           =   345
         End
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         Caption         =   $"Manual_Herreria.frx":0000
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   495
         Left            =   360
         TabIndex        =   26
         Top             =   360
         Width           =   5295
      End
      Begin VB.Label Label11 
         Caption         =   "Nota: Los precios sugeridos son con 0 (cero) Skills en Comerciar."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   25
         Top             =   5400
         Width           =   5535
      End
   End
End
Attribute VB_Name = "Manual_Herreria"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private ListaHerreria(1 To 49) As Integer

Private Sub Command2_Click()
    
    On Error GoTo Command2_Click_Err
    
    Unload Me
    Manual.Show

    
    Exit Sub

Command2_Click_Err:
    Call RegistrarError(Err.number, Err.Description, "Manual_Herreria.Command2_Click", Erl)
    Resume Next
    
End Sub

Private Sub Form_Load()
    
    On Error GoTo Form_Load_Err
    

    Dim i As Byte

    ListaHerreria(1) = 1828
    ListaHerreria(2) = 1821
    ListaHerreria(3) = 1790
    ListaHerreria(4) = 1830
    ListaHerreria(5) = 1789
    ListaHerreria(6) = 1849
    ListaHerreria(7) = 1825
    ListaHerreria(8) = 1875
    ListaHerreria(9) = 1819
    ListaHerreria(10) = 1818
    ListaHerreria(11) = 1822
    ListaHerreria(12) = 1823
    ListaHerreria(13) = 1829
    ListaHerreria(14) = 1831
    ListaHerreria(15) = 1832
    ListaHerreria(16) = 1834
    ListaHerreria(17) = 1858
    ListaHerreria(18) = 1868
    ListaHerreria(19) = 1876

    ListaHerreria(20) = 1903
    ListaHerreria(21) = 1904
    ListaHerreria(22) = 1907
    ListaHerreria(23) = 1906
    ListaHerreria(24) = 1932
    ListaHerreria(25) = 1934
    ListaHerreria(26) = 1936
    ListaHerreria(27) = 1938
    ListaHerreria(28) = 1940
    ListaHerreria(29) = 1922

    ListaHerreria(30) = 1762
    ListaHerreria(31) = 1763
    ListaHerreria(32) = 1768
    ListaHerreria(33) = 1767
    ListaHerreria(34) = 1772
    ListaHerreria(35) = 1709
    ListaHerreria(36) = 1696
    ListaHerreria(37) = 1695
    ListaHerreria(38) = 1711
    ListaHerreria(39) = 1720

    ListaHerreria(40) = 1727
    ListaHerreria(41) = 1728
    ListaHerreria(42) = 1725
    ListaHerreria(43) = 1726
    ListaHerreria(44) = 1699
    ListaHerreria(45) = 1703
    ListaHerreria(46) = 1704
    ListaHerreria(47) = 1705
    ListaHerreria(48) = 1701
    ListaHerreria(49) = 1722

    For i = 1 To 49
        List1.AddItem ObjData(ListaHerreria(i)).Name
    Next i

    List1.ListIndex = 0

    
    Exit Sub

Form_Load_Err:
    Call RegistrarError(Err.number, Err.Description, "Manual_Herreria.Form_Load", Erl)
    Resume Next
    
End Sub

Private Sub List1_Click()
    
    On Error GoTo List1_Click_Err
    

    If ObjData(ListaHerreria(List1.ListIndex + 1)).ObjType = 2 Or ObjData(ListaHerreria(List1.ListIndex + 1)).ObjType = 46 Then
        Label14.Caption = "Golpe:"
        Label17.Caption = ObjData(ListaHerreria(List1.ListIndex + 1)).MinHit & "/" & ObjData(ListaHerreria(List1.ListIndex + 1)).MaxHit
    Else
        Label14.Caption = "Defensa:"
        Label17.Caption = ObjData(ListaHerreria(List1.ListIndex + 1)).MinDef & "/" & ObjData(ListaHerreria(List1.ListIndex + 1)).MaxDef

    End If

    Label6.Caption = ObjData(ListaHerreria(List1.ListIndex + 1)).Name
    Label7.Caption = ObjData(ListaHerreria(List1.ListIndex + 1)).Valor

    'Label17.Caption = traer min level
    Label13.Caption = ObjData(ListaHerreria(List1.ListIndex + 1)).LingH
    Label8.Caption = ObjData(ListaHerreria(List1.ListIndex + 1)).LingP
    Label20.Caption = ObjData(ListaHerreria(List1.ListIndex + 1)).LingO
    Label9.Caption = ObjData(ListaHerreria(List1.ListIndex + 1)).SkHerreria

    Label3.Caption = ObjData(ListaHerreria(List1.ListIndex + 1)).Texto

    Call Grh_Render_To_Hdc(picture1, ObjData(ListaHerreria(List1.ListIndex + 1)).GrhIndex, 0, 0, False)

    
    Exit Sub

List1_Click_Err:
    Call RegistrarError(Err.number, Err.Description, "Manual_Herreria.List1_Click", Erl)
    Resume Next
    
End Sub
