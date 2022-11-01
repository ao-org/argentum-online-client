VERSION 5.00
Begin VB.Form frmShopPjsAO20 
   BackColor       =   &H00404040&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   2328
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4776
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2328
   ScaleWidth      =   4776
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtValor 
      Alignment       =   2  'Center
      BackColor       =   &H80000007&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   300
      Left            =   720
      TabIndex        =   0
      Text            =   "0"
      Top             =   960
      Width           =   3015
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Costo por publicar: 100.000 monedas de oro"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   240
      Left            =   80
      TabIndex        =   4
      Top             =   2040
      Width           =   4605
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   4440
      TabIndex        =   3
      Top             =   120
      Width           =   225
   End
   Begin VB.Label lblPublicar 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Publicar personaje"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   495
      Left            =   840
      TabIndex        =   2
      Top             =   1440
      Width           =   2775
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Ingrese el valor a publicar su personaje en ARS(Pesos Argentinos)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   1215
      Left            =   0
      TabIndex        =   1
      Top             =   240
      Width           =   4695
   End
End
Attribute VB_Name = "frmShopPjsAO20"
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
Private Sub Form_Load()

End Sub

Private Sub Label2_Click()
    Call cerrarFormulario
End Sub

Private Sub lblPublicar_Click()
    If Val(txtValor.Text <= 0) Then
        Call MsgBox("El valor ingresado del personaje es inválido.")
        Exit Sub
    End If
    
    If MsgBox("Estás publicando a " & username & " a un valor de " & txtValor.Text & ", se descontarán las 100.000 monedas de oro. En caso de querer cancelar la misma deberás hacerlo desde la página web.", vbYesNo + vbQuestion, "Publicar personaje") = vbYes Then
        Call writePublicarPersonajeMAO(Val(txtValor.Text))
        Call cerrarFormulario
    End If
    
End Sub
Private Sub cerrarFormulario()
    txtValor.Text = ""
    Unload Me
End Sub

Private Sub txtValor_Change()
    textval = txtValor.Text
    If IsNumeric(textval) Then
      numval = textval
    Else
      txtValor.Text = CStr(numval)
    End If
End Sub

