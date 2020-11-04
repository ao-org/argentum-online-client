VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmObjetos 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Crear objeto"
   ClientHeight    =   5775
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4515
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5775
   ScaleWidth      =   4515
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   465
      Left            =   3840
      ScaleHeight     =   32
      ScaleMode       =   0  'User
      ScaleWidth      =   32
      TabIndex        =   9
      Top             =   360
      Width           =   480
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Nombre Completo"
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
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   600
      Value           =   2  'Grayed
      Width           =   1575
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1800
      TabIndex        =   5
      Text            =   "1"
      Top             =   4440
      Width           =   615
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Buscar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2760
      TabIndex        =   4
      Top             =   240
      Width           =   855
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Left            =   240
      TabIndex        =   3
      Top             =   240
      Width           =   2415
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1710
      Left            =   5760
      TabIndex        =   2
      Top             =   1080
      Width           =   3090
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Salir"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   3000
      MouseIcon       =   "FrmObjetos.frx":0000
      MousePointer    =   99  'Custom
      TabIndex        =   1
      Top             =   4800
      Width           =   1290
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Crear"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      MouseIcon       =   "FrmObjetos.frx":0152
      MousePointer    =   99  'Custom
      TabIndex        =   0
      Top             =   4800
      Width           =   2730
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   3300
      Left            =   120
      TabIndex        =   8
      Top             =   960
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   5821
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      HideColumnHeaders=   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   0
      BackColor       =   16777215
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Objeto"
         Object.Width           =   5294
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Indice"
         Object.Width           =   1412
      EndProperty
   End
   Begin VB.Label Label1 
      Caption         =   "Cantidad:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   840
      TabIndex        =   6
      Top             =   4440
      Width           =   735
   End
End
Attribute VB_Name = "FrmObjetos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
  
  
'UDT necesarias para usar con SendMessage
Private Type POINTAPI
    x As Long
    y As Long
End Type
Private Type LVFINDINFO
    flags As Long
    psz As String
    lParam As Long
    pt As POINTAPI
    vkDirection As Long
End Type
'Función Api SendMessage
Private Declare Function SendMessage _
    Lib "user32" _
    Alias "SendMessageA" ( _
        ByVal hwnd As Long, _
        ByVal wMsg As Long, _
        ByVal wParam As Long, _
        lParam As Any) As Long
  
'Constantes para SendMessage
Private Const LVM_FIRST = &H1000
Private Const LVM_FINDITEM = (LVM_FIRST + 13)
Private Const LVFI_PARAM = &H1
Private Const LVFI_STRING = &H2
Private Const LVFI_PARTIAL = &H8
Private Const LVFI_WRAP = &H20
Private Const LVFI_NEARESTXY = &H40
  
'Variable de retorno y para la estructura
Dim lRet As Long, LFI As LVFINDINFO
  
  
  
'Procedimiento que busca: Se le envía el control ListView y el texto a buscar
Private Sub Buscar_ListView(ListView As ListView, _
                            Cadena As String)
  
    'Esto define si la cadena debe estar completa o si encuentra una parte _
    seleccionará el primer Item del ListView que mas se le paresca
    'If Option1 Then
        'cadena completa
      '  LFI.flags = LVFI_WRAP
    'Else
       ' 'Cadena parcial
        LFI.flags = LVFI_PARTIAL Or LVFI_WRAP Or LVFI_WRAP
    'End If
      
    If Cadena = "" Then
       Exit Sub
    End If
  
    'Se le asigna a esta variable la cadena que luego se le envía a SendMessage
    LFI.psz = Cadena
      
    'Le enviamos el mensaje LVM_FINDITEM, la estructura y rel ListView
    lRet = SendMessage(ListView.hwnd, LVM_FINDITEM, -1, LFI)
      
    If lRet >= 0 Then
        'Seleccionamos el item del Listview
        ListView.SelectedItem = ListView.ListItems(lRet + 1)
        'Propiedad opcional
        ListView.HideSelection = False
        'Si el item se encuentra fuera del área visible desplazamos la lista _
        para poder visualizarlo con el método EnsureVisible
        ListView.SelectedItem.EnsureVisible
        ListView.SetFocus
    Else
        'No se encontró
        MsgBox (" Elemento no encontrado "), vbInformation
    End If
  
  
End Sub
  
    
Private Sub Command3_Click()
    'WyroX: Cambio la lógica del buscar
    'Call Buscar_ListView(ListView1, Text1)
    
    FrmObjetos.ListView1.ListItems.Clear
    
    Dim i As Long
    For i = 1 To NumOBJs
        If InStr(1, Tilde(UCase$(ObjData(i).name)), Tilde(UCase$(Text1)), vbTextCompare) Then
            Dim subelemento As ListItem
            Set subelemento = FrmObjetos.ListView1.ListItems.Add(, , ObjData(i).name)
            
            subelemento.SubItems(1) = i
        End If
    Next i
End Sub
  
Private Sub Command1_Click()
If ListView1.SelectedItem.SubItems(1) <> "" Then

Call WriteCreateItem(ListView1.SelectedItem.SubItems(1), Text2.Text)
End If
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)
    If ListView1.SelectedItem.SubItems(1) <> "" Then
        Call Grh_Render_To_Hdc(picture1, ObjData(ListView1.SelectedItem.SubItems(1)).GrhIndex, 0, 0, False)
    End If
End Sub
