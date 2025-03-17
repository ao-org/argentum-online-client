VERSION 5.00
Begin VB.Form frmDebug 
   Caption         =   "DebugTools"
   ClientHeight    =   7200
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10365
   LinkTopic       =   "Form1"
   ScaleHeight     =   7200
   ScaleWidth      =   10365
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox TraceBox 
      Height          =   4335
      Left            =   240
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Text            =   "frmDebug.frx":0000
      Top             =   240
      Width           =   9735
   End
End
Attribute VB_Name = "frmDebug"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Me.TraceBox.Text = vbNullString

End Sub

Public Sub add_text_tracebox(ByVal s As String)
    Debug.Print s
    Me.TraceBox.Text = Me.TraceBox.Text & s & vbCrLf
End Sub
