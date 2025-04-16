VERSION 5.00
Begin VB.Form frmShopPjsAO20 
   BackColor       =   &H00404040&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3225
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6570
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3225
   ScaleWidth      =   6570
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.ComboBox cmbPaymentMethod 
      BackColor       =   &H00FF80FF&
      Height          =   315
      Left            =   3360
      TabIndex        =   5
      Text            =   "Metodo de Pago"
      ToolTipText     =   "Selecciona el metodo de pago para publicar el personaje"
      Top             =   2160
      Width           =   2415
   End
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
      Left            =   600
      TabIndex        =   0
      Text            =   "50000"
      Top             =   2160
      Width           =   2415
   End
   Begin VB.Label lblCostGold 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Costo por publicar: 20.000 monedas de oro"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   360
      Left            =   360
      TabIndex        =   4
      Top             =   1320
      Width           =   5805
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
      Left            =   6120
      TabIndex        =   3
      Top             =   120
      Width           =   225
   End
   Begin VB.Label lblPublicar 
      Alignment       =   2  'Center
      BackColor       =   &H80000017&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Publicar personaje"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   495
      Left            =   1800
      TabIndex        =   2
      Top             =   2640
      Width           =   2655
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Ingrese el valor de venta en ARS(Pesos Argentinos) y seleccione metodo de pago publicar su personaje en MAO"
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
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   5535
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

Private Sub Form_Load()
    lblCostGold.Caption = "Costo por publicar: 50.000 monedas de oro o 500 Creditos Patreon."
    cmbPaymentMethod.AddItem "Oro"
    cmbPaymentMethod.AddItem "Creditos Patreon"
    cmbPaymentMethod.ListIndex = 0  ' Set default to Gold
End Sub

Private Sub Label2_Click()
    Call cerrarFormulario
End Sub

Private Sub lblPublicar_Click()
    If Val(txtValor.Text <= 0) Then
        Call MsgBox(JsonLanguage.Item("MENSAJE_VALOR_PERSONAJE_INVALIDO"), vbCritical, JsonLanguage.Item("MENSAJE_TITULO_ERROR"))
        Exit Sub
    End If
    
    Dim paymentMethod As String
    Dim costMessage As String
    paymentMethod = cmbPaymentMethod.Text

    Select Case paymentMethod
        Case "Oro"
            costMessage = "50.000 monedas de oro"
        Case "Creditos Patreon"
            costMessage = "500 Créditos Patreon"
        Case Else
            Call MsgBox("Por favor, seleccione un método de pago.")
            Exit Sub
    End Select

    If MsgBox(JsonLanguage.Item("MENSAJE_PUBLICAR_PERSONAJE") & userName & JsonLanguage.Item("MENSAJE_PUBLICAR_PERSONAJE_VALOR") & txtValor.Text & JsonLanguage.Item("MENSAJE_PUBLICAR_PERSONAJE_COSTO"), vbYesNo + vbQuestion, JsonLanguage.Item("MENSAJE_TITULO_PUBLICAR_PERSONAJE")) = vbYes Then
        Call writePublicarPersonajeMAO(Val(txtValor.Text), paymentMethod)
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

