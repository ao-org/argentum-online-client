VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmObjetos 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Crear objeto"
   ClientHeight    =   5772
   ClientLeft      =   48
   ClientTop       =   396
   ClientWidth     =   4512
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5772
   ScaleWidth      =   4512
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   465
      Left            =   3960
      ScaleHeight     =   32
      ScaleMode       =   0  'User
      ScaleWidth      =   32
      TabIndex        =   7
      Top             =   120
      Width           =   480
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1800
      TabIndex        =   3
      Text            =   "1"
      Top             =   4440
      Width           =   615
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Left            =   1200
      TabIndex        =   0
      Top             =   240
      Width           =   2415
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1608
      Left            =   5760
      TabIndex        =   5
      Top             =   1080
      Width           =   3090
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Salir"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.4
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
      TabIndex        =   4
      Top             =   4800
      Width           =   1290
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Crear"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.4
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
      TabIndex        =   2
      Top             =   4800
      Width           =   2730
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   3540
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   4215
      _ExtentX        =   7430
      _ExtentY        =   6244
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
         Size            =   9.6
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
   Begin VB.Label Label2 
      Caption         =   "Filtrar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   480
      TabIndex        =   8
      Top             =   250
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "Cantidad:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.4
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
'    Argentum 20 - Game Client Program
'    Copyright (C) 2022 - Noland Studios
'
'    This program is free software: you can redistribute it and/or modify
'    it under the terms of the GNU Affero General Public License as published by
'    the Free Software Foundation, either version 3 of the License, or
'    (at your option) any later version.
'
'    This program is distributed in the hope that it will be useful,
'    but WITHOUT ANY WARRANTY; without even the implied warranty of
'    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'    GNU Affero General Public License for more details.
'    You should have received a copy of the GNU Affero General Public License
'    along with this program.  If not, see <https://www.gnu.org/licenses/>.
'
'
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
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
  
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
Private Sub Buscar_ListView(ListView As ListView, Cadena As String)
    
    On Error GoTo Buscar_ListView_Err
    
  
    'Esto define si la cadena debe estar completa o si encuentra una parte _
     seleccionará el primer Item del ListView que mas se le paresca
    'If Option1 Then
    'cadena completa
    '  LFI.flags = LVFI_WRAP
    'Else
    ' 'Cadena parcial
    LFI.flags = LVFI_PARTIAL Or LVFI_WRAP Or LVFI_WRAP
    'End If
      
    If Len(Cadena) = 0 Then Exit Sub

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
  
    
    Exit Sub

Buscar_ListView_Err:
    Call RegistrarError(Err.number, Err.Description, "FrmObjetos.Buscar_ListView", Erl)
    Resume Next
    
End Sub
  
Private Sub Command1_Click()
    
    On Error GoTo Command1_Click_Err
    

    If Len(ListView1.SelectedItem.SubItems(1)) <> 0 Then
        
        If Text2.Text > MAX_INVENTORY_OBJS Then Exit Sub
        
        Call WriteCreateItem(ListView1.SelectedItem.SubItems(1), Text2.Text)

    End If

    
    Exit Sub

Command1_Click_Err:
    Call RegistrarError(Err.number, Err.Description, "FrmObjetos.Command1_Click", Erl)
    Resume Next
    
End Sub

Private Sub Command2_Click()
    
    On Error GoTo Command2_Click_Err
    
    Unload Me
    
    Exit Sub

Command2_Click_Err:
    Call RegistrarError(Err.number, Err.Description, "FrmObjetos.Command2_Click", Erl)
    Resume Next
    
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    
    On Error GoTo Form_KeyPress_Err
    

    If (KeyAscii = vbKeyEscape) Then Unload Me

    
    Exit Sub

Form_KeyPress_Err:
    Call RegistrarError(Err.number, Err.Description, "FrmObjetos.Form_KeyPress", Erl)
    Resume Next
    
End Sub

Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)
    
    On Error GoTo ListView1_ItemClick_Err
    

    If Len(ListView1.SelectedItem.SubItems(1)) <> 0 Then
    
        Call Grh_Render_To_Hdc(picture1, ObjData(ListView1.SelectedItem.SubItems(1)).GrhIndex, 0, 0, False)

    End If

    
    Exit Sub

ListView1_ItemClick_Err:
    Call RegistrarError(Err.number, Err.Description, "FrmObjetos.ListView1_ItemClick", Erl)
    Resume Next
    
End Sub

Private Sub ListView1_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyEscape Then
    
        Unload Me
    
    ElseIf KeyCode = vbKeyReturn Then
        If Len(ListView1.SelectedItem.SubItems(1)) <> 0 Then
        
            If Text2.Text > MAX_INVENTORY_OBJS Then Exit Sub
            
            Call WriteCreateItem(ListView1.SelectedItem.SubItems(1), Text2.Text)
    
        End If
    End If

End Sub

Private Sub Text1_Change()

On Error GoTo Handle

    FrmObjetos.ListView1.ListItems.Clear
    
    Dim i As Long

    For i = 1 To NumOBJs

        If InStr(1, Tilde(ObjData(i).Name), Tilde(Text1), vbTextCompare) Then

            Dim subelemento As ListItem

            Set subelemento = FrmObjetos.ListView1.ListItems.Add(, , ObjData(i).Name)
            
            subelemento.SubItems(1) = i

        End If

    Next i
    
    Exit Sub

Handle:
    Call RegistrarError(Err.number, Err.Description, "FrmObjetos.Text1_Change(Buscar)")
    Resume Next
End Sub

Private Sub Text1_KeyUp(KeyCode As Integer, Shift As Integer)
    
    On Error GoTo Text1_KeyDown_Err

    If KeyCode = vbKeyEscape Then
        Unload Me
    
    ElseIf KeyCode = vbKeyReturn Then
        ListView1.SetFocus
    End If
    
    Exit Sub

Text1_KeyDown_Err:
    Call RegistrarError(Err.number, Err.Description, "FrmObjetos.Text1_KeyDown", Erl)
    Resume Next
    
End Sub
