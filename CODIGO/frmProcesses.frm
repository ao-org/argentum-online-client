VERSION 5.00
Begin VB.Form frmProcesses 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Procesos del usuario X"
   ClientHeight    =   5550
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5940
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5550
   ScaleWidth      =   5940
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox ProcesosTxt 
      BackColor       =   &H00E0E0E0&
      Height          =   4935
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   120
      Width           =   5655
   End
   Begin VB.CommandButton Close 
      Caption         =   "Salir"
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   5160
      Width           =   5895
   End
End
Attribute VB_Name = "frmProcesses"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub ShowProcesses(DATA As String)

    On Error GoTo Handler

    Dim f() As String
    
    f = Split(DATA, "*:*")
    
    If UBound(f) <= 0 Then Exit Sub
    
    If LenB(f(1)) = 0 Or f(1) = "ERROR" Then
        Me.Caption = "Error al leer los procesos."
        
        ProcesosTxt.Text = "Error al leer los procesos."

    Else

        Me.Caption = "Procesos del usuario " & f(0)

        ProcesosTxt.Text = f(1)

    End If
    
    Me.Show
    
    Exit Sub
    
Handler:
    Call RegistrarError(Err.Number, Err.Description, "frmProcesses.ShowProcesses")

End Sub

Private Sub Close_Click()

    Unload Me

End Sub

